# encoding: utf-8

require 'fileutils'
require 'nokogiri'
require 'csv'
require 'cgi'
require 'xlsx_writer'
require 'conv/headers'

module Conv
  class FromXml
    include Conv::Headers

    class Output
      def initialize argv
        FileUtils.mkdir argv[:o] if ! argv[:o].nil? && ! File.exists?(argv[:o])
        @output_file = "#{argv[:o]+'/' if ! argv[:o].nil? && Dir.exists?(argv[:o])}#{File.basename(argv[:f], '.xml')}.#{argv[:x]?'xlsx':'csv'}"

        case argv[:x]
        when true
          @doc = XlsxWriter.new
          @sheet = @doc.add_sheet("Data").tap{|s| s.add_row FromXml::HEADERS }
          @out = self
        else
          @out = CSV.open(@output_file, "wb", :headers => FromXml::HEADERS, :write_headers => true, :encoding => 'Shift_JIS')
        end
      end
    
      def get_output
        yield @out
        @out.close
      end
  
      def << row
        @sheet.add_row row
      end
  
      def close
        FileUtils::mv @doc.path, "#{@output_file}"
        @doc.cleanup
      end
    end

    def initialize argv
      @argv = argv
    end

    def integer_string?(str)
      begin
        Integer(str)
        true
      rescue ArgumentError
        false
      end
    end

    def getCheckItem item, root, row, know_how_detail
      row[CHECK_ITEM_NAME] = item.xpath('xmlns:CheckItemName').text.encode(Encoding::Shift_JIS, undef: :replace)
      row[SEARCH_PROCESS] = item.xpath('xmlns:SearchProcess').text.encode(Encoding::Shift_JIS, undef: :replace)
      row[SEARCH_EXIST] = item.attr('searchExistance')
      row[FACTOR] = item.xpath('xmlns:PortabilityFactor').text.encode(Encoding::Shift_JIS, undef: :replace)
      row[DEGREE] = item.xpath('xmlns:PortabilityDegree').text.encode(Encoding::Shift_JIS, undef: :replace)
      row[DEGREE_DETAIL] = item.xpath('xmlns:DegreeDetail').text.encode(Encoding::Shift_JIS, undef: :replace)
      row[VISUAL_CONFIRM] = item.xpath('xmlns:VisualConfirm').text.encode(Encoding::Shift_JIS, undef: :replace)
      row[HEARING_CONFIRM] = item.xpath('xmlns:HearingConfirm').text.encode(Encoding::Shift_JIS, undef: :replace)

      row[KNOWLEDGE_DETAIL] = item.xpath('xmlns:CheckItemNo').text.encode(Encoding::Shift_JIS, undef: :replace) == '1' ? know_how_detail : ''

      search_info_ary = Array.new
      root.xpath("//xmlns:SearchInfomation[@searchInfoId='#{item.attr('searchRefKey')}']").each do |e|
        row[FILE_TYPE] = e.xpath('xmlns:FileType').text.encode(Encoding::Shift_JIS, undef: :replace)
        row[KEYWORD_1] = e.xpath('xmlns:SearchKey1').text.encode(Encoding::Shift_JIS, undef: :replace)
        row[KEYWORD_2] = e.xpath('xmlns:SearchKey2').text.encode(Encoding::Shift_JIS, undef: :replace)
        row[MODULE] = e.xpath('xmlns:PythonModule').text.encode(Encoding::Shift_JIS, undef: :replace)
        row[LINE_NUM_APPROPRIATE] = e.xpath("xmlns:Appropriate").attr('lineNumberAppropriate').text.encode(Encoding::Shift_JIS, undef: :replace)
        row[APPROPRIATE_CONTENTS] = e.xpath('xmlns:Appropriate/xmlns:AppropriateContents').text.encode(Encoding::Shift_JIS, undef: :replace)
        line_no = e.xpath('xmlns:LineNumberInfomation/xmlns:LineNumber').text.encode(Encoding::Shift_JIS, undef: :replace)
        row[TODO] = line_no if ! integer_string? line_no
        row[LINE_NUM] = line_no if integer_string? line_no
        row[LINE_NUM_CONTENTS] = e.xpath('xmlns:LineNumberInfomation/xmlns:LineNumberContents').text.encode(Encoding::Shift_JIS, undef: :replace)
        row[INVEST] = e.xpath('xmlns:LineNumberInfomation/xmlns:Investigation').text.encode(Encoding::Shift_JIS, undef: :replace)
      end
      row
    end

    def getEntryCategory out, root, node
      cat_id = node.text.encode(Encoding::Shift_JIS, undef: :replace)
      chap_name = root.xpath("//xmlns:ChapterList/xmlns:Chapter/xmlns:ChapterName[following-sibling::xmlns:ChildChapter/xmlns:ChapterCategoryRefKey/text()='#{cat_id}']").text.encode(Encoding::Shift_JIS, undef: :replace)
      chap_no = root.xpath("//xmlns:ChildChapter/xmlns:ChildCapterNo[following-sibling::xmlns:ChapterCategoryRefKey/text()='#{cat_id}']").text.encode(Encoding::Shift_JIS, undef: :replace)

      cat = root.xpath("//xmlns:CategoryList/xmlns:Category[@categoryId='#{cat_id}']")
      cat_name = cat.xpath('xmlns:CategoryName/text()').text.encode(Encoding::Shift_JIS, undef: :replace)

      par_cat_id = root.xpath("//xmlns:EntryCategoryRefKey[following-sibling::xmlns:ChildEntry/xmlns:EntryCategoryRefKey/text()='#{cat_id}']").text.encode(Encoding::Shift_JIS, undef: :replace)
      par_name = root.xpath("//xmlns:CategoryList/xmlns:Category[@categoryId='#{par_cat_id}']/xmlns:CategoryName/text()").text.encode(Encoding::Shift_JIS, undef: :replace)

      cat.xpath('xmlns:KnowhowRefKey').tap{|s| out << [chap_no, chap_name, cat_name, par_name] if s.empty?}.each do |know_how_ref|
        know_how = root.xpath("//xmlns:KnowhowList/xmlns:KnowhowInfomation[@knowhowId='#{know_how_ref.text.encode(Encoding::Shift_JIS, undef: :replace)}']")
        know_how_name = know_how.xpath('xmlns:KnowhowName').text.encode(Encoding::Shift_JIS, undef: :replace)
        know_how_detail_ref = know_how.attribute('knowhowDetailRefKey') if ! know_how.empty?
        know_how_detail = root.xpath("//xmlns:DocBook[@articleId='#{know_how_detail_ref}']/ns3:article/ns3:section/node()").to_ary.map{|e| e.to_xml.strip}.join

        row = CSV::Row.new(HEADERS, [chap_no, chap_name, cat_name, par_name, know_how_name])
        know_how.xpath('xmlns:CheckItem').tap{|s| out << [chap_no, chap_name, cat_name, par_name, know_how_name, know_how_detail] if s.empty?}.each do |item|
          out << getCheckItem(item, root, row, know_how_detail).to_hash.values
        end

      end
    end

    def process
      Output.new(@argv).get_output do |out|
        file = Nokogiri::XML(open(@argv[:f]))
        file.tap{|s| s.xpath("//xmlns:EntryCategoryRefKey").each {|node| getEntryCategory out, s, node}}
        out << ["@@" + file.xpath("//xmlns:PortabilityKnowhowTitle").text.encode(Encoding::Shift_JIS, undef: :replace)]
      end
    end
  end
end
