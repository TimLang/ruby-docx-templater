
#TODO using Table::TD model to refactor td attrs,instead of using regex
module DocxTemplater

  module Table

    class TD
     
      attr_reader :text, :width, :colspan, :rowspan

      def initialize text, width, colspan, rowspan
        @text = text
        @width = width
        @colspan = colspan
        @rowspan = rowspan
      end

    end

  end
end
