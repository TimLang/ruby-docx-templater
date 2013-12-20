
require 'builder'

module DocxTemplater

  module Element

    def create_paragraph element
      do_builder do |builder|
        xml = builder.tag!('w:p', {'w:rsidR' => '0019258A', 'w:rsidRDefault' => '00214805'}) do
          builder.tag!('w:rPr'){builder.tag!('w:rPr'){builder.tag!('w:rFonts', {'w:hint' => 'eastAsia'})}}
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

    def create_whilespace
      do_builder do |builder|
        builder.tag!('w:t', {'xml:space' => 'preserve'})
      end
    end

    def do_builder
      yield Builder::XmlMarkup.new
    end
    private :do_builder

  end
end
