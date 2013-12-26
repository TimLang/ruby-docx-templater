# -*- encoding : utf-8 -*-
require 'debugger'
module DocxTemplater
  class TemplateProcessor
    attr_reader :data, :escape_html

    include DocxTemplater::Element

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
          if key =~ /^c_(\d+)/
            document = excute_cols(document, key)
          else
            document = enter_multiple_values(document, key)
          end
        else
          generate_paragraph(document, key, value)
        end
      end
      clean_unreplaced_tags(document)
    end

    def process_media document, images
      document = Nokogiri::XML(document)

      relationships = document.search('Relationships').first 
      max_index = relationships.search('Relationship').map{|r| r['Id'].gsub(/rId(\d+)/, '\1').to_i}.max

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

    def clean_unreplaced_tags document
      xml = Nokogiri::XML(document)
      last_tbl = xml.search('//w:tbl').find{|a| !a.xpath(".//w:tc[contains(., '#B_COL')]", xml.root.namespaces).empty?}
      col_templates = last_tbl.xpath(".//w:tc[contains(., '#B_COL')]", xml.root.namespaces).each do |c|
        c.parent.unlink
      end if last_tbl
      xml.to_s
    end

    def generate_paragraph document, key, value
      document.gsub!("$#{key.to_s.upcase}$", excute_newline(value).join) 
    end

    def generate_each_paragraph document, key, value
      if value =~ /\n/
        document.gsub(/<w:p.*?>[\s\S]*?<\/w:p>/, excute_newline(value).join)
      else
        document.gsub("$EACH:#{key.to_s.upcase}$", excute_newline(value).join)
      end
    end


    def excute_newline value
      if value =~ /(_``table:.+?``_)|\n/
        value.split(/(_``table:.+?``_)|\n/).inject([]) do |result, str|
          if str =~ /(_``table:.+?``_)/
            result << create_paragraph(excute_nested_table(str).join)
          else
            result << create_paragraph(excute_nested_image_with_text(str).join)
          end
        end
      else
        [excute_nested_image_with_text(value)]
      end
    end

    def excute_nested_image_with_text value
      value.split(/(\${image\d+})/).inject([]) do |result, str|
        unless str =~ /\${image\d+}/
          result << create_text(safe(str))
        else
          img = @cached_images[str.gsub(/\${(\w+)}/, '\1').to_sym]
          if img
            result << create_image(img.embed_id, 'Tim', get_word_image_dimension(img.width), get_word_image_dimension(img.height))
          else
            result << create_text('')
          end
        end
      end
    end
    
    def excute_nested_table value
      table_regex = /(_``table:.+``_)/
      value.split(table_regex).inject([]) do |result ,str|
        unless str =~ table_regex
          result << create_text(safe(str))
        else
          begin
            result << create_table(
              eval(value.sub(/_``table:(.+)``_/, '\1')).deep_map! do |a| 
                excute_nested_image_with_text(safe(a)).join
              end
            )
          rescue SyntaxError => e
            result << create_text('')
          end
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
                @items_cache[cache_key] = create_tr_wrapper(innards)
                #@items_cache[cache_key] = innards
                each_data[LOOP_PLACE_HOLDER].reverse.each_with_index do |e, i|
                  innards = [] if i == 0
                  obj_key = (each_key =~ /items_(.+)/i) ? each_key.gsub(/items_(.+)/i, $1).downcase : ''
                  innards << create_blank_tr
                  if e[:choice]
                    #indentation for choices
                    tpl = @items_cache[cache_key].sub('<w:pPr>', " <w:pPr>\n        <w:ind w:leftChars=\"100\" w:left=\"180\"/>")
                    if e[:choice].class == Array
                      e[:choice].reverse.each do |c|
                        innards << generate_each_paragraph(tpl, each_key, safe(c))
                      end
                    else
                      innards << generate_each_paragraph(tpl, each_key, safe(e[:choice]))
                    end
                  end
                  innards << generate_each_paragraph(@items_cache[cache_key], each_key, safe(e[obj_key.to_sym]))
                end if each_data[LOOP_PLACE_HOLDER]
              else
                value = safe(each_data[each_key.downcase.to_sym])
                innards = value=='' ? '' : (create_blank_tr + generate_each_paragraph(new_row.to_html, each_key, value))
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
  
    def excute_cols document, key
      xml = Nokogiri::XML(document)
      begin_col = "#B_COL:#{key.to_s.upcase}#"
      end_col = "#E_COL:#{key.to_s.upcase}#"
      last_tbl = xml.search('//w:tbl').find{|a| !a.xpath(".//w:tc[contains(., '#{begin_col}')]", xml.root.namespaces).empty?}
      begin_col_template = last_tbl.xpath(".//w:tc[contains(., '#{begin_col}')]", xml.root.namespaces).first
      end_col_template = last_tbl.xpath(".//w:tc[contains(., '#{end_col}')]", xml.root.namespaces).first
      #DocxTemplater::log("begin_col_template: #{begin_col_template.to_s}")
      #DocxTemplater::log("end_col_template: #{end_col_template.to_s}")
      raise "unmatched template markers: #{begin_col} nil: #{begin_col_template.nil?}, #{end_col} nil: #{end_col_template.nil?}. This could be because word broke up tags with it's own xml entries. See README." unless begin_col_template && end_col_template

      col_templates = []
      col = begin_col_template.next_sibling
      while (col != end_col_template)
        col_templates.unshift(col)
        col = col.next_sibling
      end

      data[key].reverse.each do |each_data|
        #DocxTemplater::log("each_data: #{each_data.inspect}")

        # dup so we have new nodes to append
        col_templates.map(&:dup).each do |new_col|
          #DocxTemplater::log("   new_col: #{new_col}")
          innards = new_col.inner_html
          if !(matches = innards.scan(/\$EACH:([^\$]+)\$/)).empty?
            #DocxTemplater::log("   matches: #{matches.inspect}")
            matches.map(&:first).each do |each_key|
              #DocxTemplater::log("      each_key: #{each_key}")
              innards = generate_each_paragraph(innards, each_key, safe(each_data[each_key.downcase.to_sym]))
              #innards.gsub!("$EACH:#{each_key}$", safe(each_data[each_key.downcase.to_sym]))
            end
          end
          # change all the internals of the new node, even if we did not template
          new_col.inner_html = innards
          #DocxTemplater::log("new_col new innards: #{new_col.inner_html}")

          begin_col_template.add_next_sibling(new_col)
        end
      end
      (col_templates + [begin_col_template, end_col_template]).map(&:unlink)
      xml.to_s
    end

  end
end
