module Serenity
  class WordImage
    IMAGE_DIR_NAME = "Pictures"

    attr_accessor :file_path, :context, :file_path_output

    def initialize(file_path, context, content_img, file_path_output)
      @replacements = []
      @file_path = file_path
      @context = context
      @images = eval('@images', context)
      @file_path_output = file_path_output
      @xml_content = content_img
    end

    # WordImage.adicionar_imagem_no_word(1,1)
    def adicionar_imagem_no_word(x, y, xml_file, zipfile)
      require 'nokogiri'

      if @images && @images.kind_of?(Hash)
        xml_data = Nokogiri::XML(@xml_content)
        @images.each do |image_name, replacement_path|
          unless zipfile.find_entry('word/media/logo.png')
            zipfile.add('word/media/logo.png',  ::File.join(Rails.public_path, 'logo.png'))
          end
        end

        if xml_file == '[Content_Types].xml'
          # xml_data.xpath("//Override[@ContentType='image/png']") - caso fosse apenas o overrides..
          xml_data.xpath("//xmlns:Override[@ContentType='image/png']").each_with_index do |node, index|
            placeholder_path = node.attribute('PartName').value
            odt_image_path = ::File.join('/word/media/', ::File.basename(placeholder_path))
            @replacements << [odt_image_path, "/word/media/#{@images.to_a[index][1]}"]
          end
        end
        if xml_file == 'word/_rels/document.xml.rels'
          xml_data.xpath("//xmlns:Relationship[contains(@Target, '.jpg') or contains(@Target, '.jpeg') or contains(@Target, '.png') or contains(@Target, '.gif')]").each_with_index do |node, index|
            placeholder_path = node.attribute('Target').value
            odt_image_path = ::File.join('media/', ::File.basename(placeholder_path))
            @replacements << [odt_image_path, "media/#{@images.to_a[index][1]}"]
          end
        end

        @replacements.each do |r|
          @xml_content = @xml_content.gsub(r.first, r.last)
        end

      end

      return @xml_content
    end

  end
end
