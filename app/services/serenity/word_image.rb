module Serenity
  class WordImage
    IMAGE_DIR_NAME = "word/media"

    # attr_accessor :file_path, :context, :file_path_output

    def initialize(context, content, zipfile)
      @path_images = eval('@path_images', context) #teste
      @content = content
      @context = context
      @zipfile = zipfile
    end

    def generate_replacements_docx(name_file)
      require 'nokogiri'
      replacements = []

      if @path_images && @path_images.kind_of?(Array)
        
        xml_data = Nokogiri::XML(@content)
        @path_images.each do |image_path|
          img_substituta = "#{IMAGE_DIR_NAME}/#{::File.basename(image_path)}"
          unless @zipfile.find_entry(img_substituta)
            @zipfile.add(img_substituta,  image_path)
          end
        end



        if name_file == '[Content_Types].xml'
          xml_data.xpath("//xmlns:Override[@ContentType='image/png']").each_with_index do |node, index|
            placeholder_path = node.attribute('PartName').value
            odt_image_path = ::File.join("/#{IMAGE_DIR_NAME}/", ::File.basename(placeholder_path))
            img_substituta = "#{IMAGE_DIR_NAME}/#{::File.basename(@path_images[index])}"
            replacements << [odt_image_path, "/#{img_substituta}"]
          end
        end
        if name_file == 'word/_rels/document.xml.rels'
          xml_data.xpath("//xmlns:Relationship[contains(@Target, '.jpg') or contains(@Target, '.jpeg') or contains(@Target, '.png') or contains(@Target, '.gif')]").each_with_index do |node, index|
            placeholder_path = node.attribute('Target').value
            odt_image_path = ::File.join('media/', ::File.basename(placeholder_path))
            img_substituta = "#{::File.basename(@path_images[index])}"
            replacements << [odt_image_path, "media/#{img_substituta}"]
          end
        end

        replacements.each do |r|
          @content.sub!(r.first, r.last)
        end
      end
    end

  end
end
