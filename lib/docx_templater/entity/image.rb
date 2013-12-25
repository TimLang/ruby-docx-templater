
module DocxTemplater

  class Image
    
    attr_reader :name, :base64_str
    attr_accessor :embed_id, :width, :height

    def initialize name, base64_str, width=nil, height=nil 
      @name = name
      @base64_str = base64_str
      @width = width
      @height = height
    end

  end

end
