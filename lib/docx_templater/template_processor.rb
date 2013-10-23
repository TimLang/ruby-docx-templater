# -*- encoding : utf-8 -*-
require 'debugger'

module DocxTemplater
  class TemplateProcessor
    attr_reader :data, :escape_html

    LOOP_PLACE_HOLDER = :items
       # data is expected to be a hash of symbols => string or arrays of hashes.
    def initialize(data, cached_images, escape_html=true)
      @data = data
      @cached_images = cached_images
      @escape_html = escape_html
      @items_cache = {}
    end

    def render(document)
      # in order to sloved encoding bugs
      document = Nokogiri::XML(document, nil, 'utf-8').to_s
      #document.encoding = 'utf-8'
      #document = document.to_s
      data.each do |key, value|
        if value.class == Array
          document = enter_multiple_values(document, key)
        else
          generate_paragraph(document, key, value)
        end
      end
      document
    end

    def process_media document, images
      document = Nokogiri::XML(document)

      relationships = document.search('Relationships').first 
      max_index = relationships.search('Relationship').map{|r| r['Id'].gsub!(/rId(\d+)/, '\1')}.max

      images.each_with_index do |(k, img), i|
        node = Nokogiri::XML::Node.new('Relationship', document)
        node['Id'] = "rId#{max_index.to_i + i + 1}"
        node['Type'] = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
        node['Target'] = "media/#{img.name}"
      
        relationships.add_child(node)
        @cached_images[k].embed_id = node['Id']
      end
      document.to_s
    end

    private

    def generate_paragraph document, key, value
      document.gsub!("$#{key.to_s.upcase}$", excute_newline(value).join) 
    end

    def generate_each_paragraph document, key, value
      document.gsub("$EACH:#{key.to_s.upcase}$", excute_newline(value).join)
    end


    def excute_newline value
      if value =~ /\n/
        value.split(/\n/).inject([]) do |result, str|
          result << PARAGRAPH_ROW.gsub(/\$text\$/, (excute_nested_image_with_text(str)).join)
        end
      else
        [excute_nested_image_with_text(value)]
      end
    end

    def excute_nested_image_with_text value
      value.split(/(\${image\d+})/).inject([]) do |result, str|
        unless str =~ /\${image\d+}/
          result << TEXT_ROW.gsub('$text$', safe(str))
        else
          img = @cached_images[str.gsub(/\${(\w+)}/, '\1').to_sym]
          result << DRAWING_ROW.gsub(/\$embed_id\$/, img.embed_id)
          .gsub(/\$image_width\$/, get_word_image_dimension(img.width))
          .gsub(/\$image_height\$/, get_word_image_dimension(img.height))
          .gsub(/\$image_name\$/, 'tim') if img
        end
      end
    end

    def get_word_image_dimension unit
      (unit * (1/28.to_f) * 270000).round.to_s
    end

    def safe(text)
      if escape_html
        text.to_s.gsub('&', '&amp;').gsub('>', '&gt;').gsub('<', '&lt;')
      else
        text.to_s
      end
    end

    def enter_multiple_values(document, key)
      DocxTemplater::log("enter_multiple_values for: #{key}")
      # TODO ideally we would not re-parse xml doc every time
      xml = Nokogiri::XML(document)
      begin_row = "#BEGIN_ROW:#{key.to_s.upcase}#"
      end_row = "#END_ROW:#{key.to_s.upcase}#"
      begin_row_template = xml.xpath("//w:tr[contains(., '#{begin_row}')]", xml.root.namespaces).first
      end_row_template = xml.xpath("//w:tr[contains(., '#{end_row}')]", xml.root.namespaces).first
      #DocxTemplater::log("begin_row_template: #{begin_row_template.to_s}")
      #DocxTemplater::log("end_row_template: #{end_row_template.to_s}")
      raise "unmatched template markers: #{begin_row} nil: #{begin_row_template.nil?}, #{end_row} nil: #{end_row_template.nil?}. This could be because word broke up tags with it's own xml entries. See README." unless begin_row_template && end_row_template

      row_templates = []
      row = begin_row_template.next_sibling
      while (row != end_row_template)
        row_templates.unshift(row)
        row = row.next_sibling
      end
      #DocxTemplater::log("row_templates: (#{row_templates.count}) #{row_templates.map(&:to_s).inspect}")

      # for each data, reversed so they come out in the right order
      data[key].reverse.each do |each_data|
        #DocxTemplater::log("each_data: #{each_data.inspect}")

        # dup so we have new nodes to append
        row_templates.map(&:dup).each do |new_row|
          #DocxTemplater::log("   new_row: #{new_row}")
          innards = new_row.inner_html
          if !(matches = innards.scan(/\$EACH:([^\$]+)\$/)).empty?
            #DocxTemplater::log("   matches: #{matches.inspect}")
            matches.map(&:first).each do |each_key|
              #DocxTemplater::log("      each_key: #{each_key}")
              if each_key =~ /items_(.+)/i 
                cache_key = each_key.to_sym
                @items_cache[cache_key] = TR_WRAPPER_ROW.gsub(/\$text\$/, innards)
                each_data[LOOP_PLACE_HOLDER].reverse.each_with_index do |e, i|
                  innards = [] if i == 0
                  obj_key = (each_key =~ /items_(.+)/i) ? each_key.gsub(/items_(.+)/i, $1).downcase : ''
                  innards << BLANK_ROW
                  if e[:choice]
                    if e[:choice].class == Array
                      e[:choice].reverse.each do |c|
                        innards << generate_each_paragraph(@items_cache[cache_key], each_key, safe(c))
                      end
                    else
                      innards << generate_each_paragraph(@items_cache[cache_key], each_key, safe(e[:choice]))
                    end
                  end
                  innards << generate_each_paragraph(@items_cache[cache_key], each_key, safe(e[obj_key.to_sym]))
                end if each_data[LOOP_PLACE_HOLDER]
              else
                innards = generate_each_paragraph(innards, each_key, safe(each_data[each_key.downcase.to_sym]))
              end
            end
          end
          # change all the internals of the new node, even if we did not template
          if innards.class == Array
            innards.each do |inn|
              begin_row_template.add_next_sibling(inn)
            end
          else
            begin_row_template.add_next_sibling(innards)
          end
          #DocxTemplater::log("new_row new innards: #{new_row.inner_html}")
        end
      end
      (row_templates + [begin_row_template, end_row_template]).map(&:unlink)
      xml.to_s
    end

    # hard coding, my god!

    SPACE_TEXT = '<w:t xml:space="preserve">  </w:t>'

    BLANK_ROW = "<w:tr w:rsidR=\"00D779AB\" w:rsidTr=\"00B812D2\">\n        <w:tc>\n          <w:tcPr>\n            <w:tcW w:w=\"8522\" w:type=\"dxa\"/>\n          </w:tcPr>\n          <w:p w:rsidR=\"00D779AB\" w:rsidRDefault=\"00D779AB\" w:rsidP=\"00C44DF6\">\n            <w:pPr>\n              <w:rPr>\n                <w:rFonts w:hint=\"eastAsia\"/>\n              </w:rPr>\n            </w:pPr>\n            <w:r>\n              <w:rPr>\n                <w:rFonts w:hint=\"eastAsia\"/>\n              </w:rPr>\n              <w:t></w:t>\n            </w:r>\n          </w:p>\n        </w:tc>\n      </w:tr>"

    TR_WRAPPER_ROW = '\n <w:tr w:rsidR="004D5284" w:rsidTr="004D5284">\n $text$ </w:tr>\n'

    TEXT_ROW = '<w:r>
              <w:rPr>
                <w:rFonts w:hint="eastAsia"/>
              </w:rPr>
              <w:t>$text$</w:t>
            </w:r>'

    PARAGRAPH_ROW = '<w:p w:rsidR="0019258A" w:rsidRDefault="00214805">
	<w:pPr>
		<w:rPr>
			<w:rFonts w:hint="eastAsia"/>
		</w:rPr>
	</w:pPr>$text$</w:p>'

    DRAWING_ROW = '<w:r>
              <w:rPr>
                <w:rFonts w:hint="eastAsia"/>
                <w:noProof/>
              </w:rPr>
              <w:drawing>
                <wp:inline distT="0" distB="0" distL="0" distR="0">
                  <wp:extent cx="$image_width$" cy="$image_height$"/>
                  <wp:effectExtent l="19050" t="0" r="7990" b="0"/>
                  <wp:docPr id="2" name="图片" descr="$image_name$"/>
                  <wp:cNvGraphicFramePr>
                    <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
                  </wp:cNvGraphicFramePr>
                  <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                      <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                        <pic:nvPicPr>
                          <pic:cNvPr id="0" name="$image_name$"/>
                          <pic:cNvPicPr/>
                        </pic:nvPicPr>
                        <pic:blipFill>
                          <a:blip r:embed="$embed_id$" cstate="print"/>
                          <a:stretch>
                            <a:fillRect/>
                          </a:stretch>
                        </pic:blipFill>
                        <pic:spPr>
                          <a:xfrm>
                            <a:off x="0" y="0"/>
                            <a:ext cx="$image_width$" cy="$image_height$"/>
                          </a:xfrm>
                          <a:prstGeom prst="rect">
                            <a:avLst/>
                          </a:prstGeom>
                        </pic:spPr>
                      </pic:pic>
                    </a:graphicData>
                  </a:graphic>
                </wp:inline>
              </w:drawing>
            </w:r>'

  end
end
