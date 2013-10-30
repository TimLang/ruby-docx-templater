# -*- encoding : utf-8 -*-
module DocxTemplater
  class DocxCreator
    attr_reader :template_path, :template_processor, :cached_images

    def initialize(template_path, data, cached_images, escape_html=true)
      @template_path = template_path
      @template_processor = TemplateProcessor.new(data, cached_images, escape_html)
      @cached_images = cached_images
    end

    def generate_docx_file(file_name = "output_#{Time.now.strftime("%Y-%m-%d_%H%M")}.docx")
      buffer = generate_docx_bytes
      File.open(file_name, 'w') { |f| f.write(buffer) }
    end

    def generate_docx_bytes
      buffer = ''
      read_existing_template_docx do |template|
        create_new_zip_in_memory(buffer, template)
      end
      buffer
    end

    private

    def copy_or_template(entry_name, f)
      # Inside the word document archive is one file with contents of the actual document. Modify it.
      return template_processor.render(f.read) if entry_name == 'word/document.xml'
      return template_processor.process_media(f.read, @cached_images) if entry_name == 'word/_rels/document.xml.rels'
      f.read
    end

    def read_existing_template_docx
      ZipRuby::Archive.open(template_path) do |template|
        yield template
      end
    end

    def create_new_zip_in_memory(buffer, template)
      n_entries = template.num_files
      ZipRuby::Archive.open_buffer(buffer, ZipRuby::CREATE) do |archive|

        generate_images_on_word(archive)

        n_entries.times do |i|
          entry_name = template.get_name(i)
          template.fopen(entry_name) do |f|
            archive.add_buffer(entry_name, copy_or_template(entry_name, f))
          end
        end
      end
    end

    def generate_images_on_word archive
      @cached_images.each do |k, img|
        archive.add_buffer("word/media/#{img.name}", convert_img_format_to_jpeg(img.base64_str, img))
      end
    end

    def convert_img_format_to_jpeg(base64_str, img)
      image = MiniMagick::Image.from_blob(Base64.decode64(base64_str))
    
      image.resize("#{img.width} x #{img.height}") if img.width && img.height

      img.width = image['width']
      img.height = image['height']

      #image.format 'jpeg'
      image.to_blob
    end

  end
end
