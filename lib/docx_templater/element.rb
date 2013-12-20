
require 'builder'

module DocxTemplater

  module Element

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
                    builder.tag!('w:tcPr'){builder.tag!('w:tcW', {'w:w' => width, 'w:type' => 'dxa'})}
                    builder.tag!('w:p', {'w:rsidR' => '006217C8', 'w:rsidRDefault' => '006217C8'}) do
                      builder.tag!('w:r') do
                        builder.tag!('w:rPr'){builder.tag!('w:rFonts', {'w:hint' => 'eastAsia'})}
                        builder.tag!('w:t') { builder << col.to_s}
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

    def do_builder
      yield Builder::XmlMarkup.new
    end
    private :do_builder

  end
end
