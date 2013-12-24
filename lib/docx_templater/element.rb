
require 'builder'

module DocxTemplater

  module Element

    GRID_SPAN_REGEX = /(__{(\d+)}__)/

    def create_paragraph element
      do_builder do |builder|
        xml = builder.tag!('w:p', {'w:rsidR' => '0019258A', 'w:rsidRDefault' => '00214805'}) do
          builder.tag!('w:pPr'){builder.tag!('w:rPr'){builder.tag!('w:rFonts', {'w:hint' => 'eastAsia'})}}
          builder << element
        end
      end
    end

    def create_text text
      do_builder do |builder|
        builder.tag!('w:r') do
          builder.tag!('w:rPr'){builder.tag!('w:rPr'){builder.tag!('w:rFonts', {'w:hint' => 'eastAsia'})}}
          builder.tag!('w:t') {builder << text}
        end
      end
    end

    def create_tr_wrapper item
      do_builder do |builder|
        builder.tag!('w:tr', {'w:rsidR' => '004D5284', 'w:rsidTr' => '004D5284'}) do
          builder << item
        end
      end
    end

    def create_whilespace
      do_builder do |builder|
        builder.tag!('w:t', {'xml:space' => 'preserve'})
      end
    end

    #=~ 20.3*136=2763
    def create_table rows, width="2763"
      return nil if !rows.is_a?(Array) || !rows[0].is_a?(Array)
      do_builder do |builder|
        builder.tag!('w:r') do
          builder.tag!('w:rPr') do
            builder.tag!('w:rFonts', {'w:hint' => 'eastAsia'})
            builder.tag!('w:noProof')
          end
          builder.tag!('w:tbl') do
            builder.tag!('w:tblPr') do
              builder.tag!('w:tblStyle', {'w:val' => 'a5'})
              builder.tag!('w:tblW', {'w:w' => '0', 'w:type' => 'auto'})
              builder.tag!('w:tblLook', {'w:val' => '04A0'})
            end
            builder.tag!('w:tblGrid') do
              rows.size.times do
                builder.tag!('w:gridCol', {'w:w' => width})
              end
            end

            rows.each do |row|
              #TR
              builder.tag!('w:tr', {'w:rsidR' => '006217C8', 'w:rsidTr' => '006217C8'}) do

                row.each do |col|
                  #TD
                  builder.tag!('w:tc') do
                    builder.tag!('w:tcPr') do
                      builder.tag!('w:tcW', {'w:w' => width, 'w:type' => 'dxa'})
                      builder.tag!('w:gridSpan', {'w:val'=>$2}) if col =~ GRID_SPAN_REGEX
                    end
                    builder.tag!('w:p', {'w:rsidR' => '006217C8', 'w:rsidRDefault' => '006217C8'}) do
                      builder.tag!('w:r') do
                        builder.tag!('w:rPr'){builder.tag!('w:rFonts', {'w:hint' => 'eastAsia'})}
                        builder.tag!('w:t') { builder << col.sub(GRID_SPAN_REGEX, '').to_s}
                      end
                    end
                  end
                end

              end
            end
            #
          end
        end
      end
    end

    def create_image embed_id, name, width, height
      do_builder do |builder|
        builder.tag!('w:r') do
          builder.tag!('w:rPr') do
            builder.tag!('w:rFonts', {'w:hint' => 'eastAsia'})
            builder.tag!('w:noProof')
          end
          builder.tag!('w:drawing') do
            builder.tag!('wp:inline', {'distT'=>'0', 'distB'=>'0','distL'=>'0','distR'=>'0'}) do
              builder.tag!('wp:extent',{'cx'=>width,'cy'=>height})
              builder.tag!('wp:effectExtent',{'l'=>'19050','t'=>'0','r'=>'7990','b'=>'0'})
              builder.tag!('wp:docPr',{'id'=>'2','name'=>name,'descr'=>name})
              builder.tag!('wp:cNvGraphicFramePr') do
                builder.tag!('a:graphicFrameLocks', {'xmlns:a'=>'http://schemas.openxmlformats.org/drawingml/2006/main','noChangeAspect'=>'1'})
              end
              builder.tag!('a:graphic',{'xmlns:a'=>'http://schemas.openxmlformats.org/drawingml/2006/main'}) do
                builder.tag!('a:graphicData',{'uri' => 'http://schemas.openxmlformats.org/drawingml/2006/picture'}) do
                  builder.tag!('pic:pic',{'xmlns:pic'=>'http://schemas.openxmlformats.org/drawingml/2006/picture'}) do
                    builder.tag!('pic:nvPicPr') do
                      builder.tag!('pic:cNvPr', {'id'=>'0','name'=>name})
                      builder.tag!('pic:cNvPicPr')
                    end
                    builder.tag!('pic:blipFill') do
                      builder.tag!('a:blip',{'r:embed'=>embed_id,'cstate'=>'print'})
                      builder.tag!('a:stretch'){builder.tag!('a:fillRect')}
                    end
                    builder.tag!('pic:spPr') do
                      builder.tag!('a:xfrm') do
                        builder.tag!('a:off',{'x'=>'0','y'=>'0'})
                        builder.tag!('a:ext',{'cx'=>width,'cy'=>height})
                      end
                      builder.tag!('a:prstGeom',{'prst'=>'rect'}){builder.tag!('a:avLst')}
                    end
                  end
                end

              end
            end
          end
        end
      end
    end

    def do_builder
      yield Builder::XmlMarkup.new
    end
    private :do_builder

  end
end
