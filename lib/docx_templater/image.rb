
module DocxTemplater

  class Image
    
    attr_reader :name, :base64_str
    attr_accessor :embed_id, :width, :height

    def initialize name, base64_str 
      @name = name
      @base64_str = base64_str
    end

  end

end
