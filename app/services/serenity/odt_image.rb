module Serenity
  class OdtImage
    include Debug
    IMAGE_DIR_NAME = "Pictures"

    # attr_accessor :content

    def initialize(content, context, zipfile)
      @path_images = eval('@path_images', context) #teste
      @content = content
      @context = context
      @zipfile = zipfile
    end

    def generate_replacements
      require 'nokogiri'
      replacements = []
      if @path_images && @path_images.kind_of?(Array)
        xml_data = Nokogiri::XML(@content)
          xml_data.xpath('//draw:frame[starts-with(@draw:name, "Figura")]').each_with_index do |node, index|
            img_substituta = "#{IMAGE_DIR_NAME}/#{::File.basename(@path_images[index])}"
            unless @zipfile.find_entry(img_substituta)
              @zipfile.add(img_substituta, @path_images[index])
            end
            image_elem = node.at_xpath("//draw:image")
            placeholder_path = image_elem.attribute('href').value
            odt_image_path = ::File.join(IMAGE_DIR_NAME, ::File.basename(placeholder_path))
            replacements << [odt_image_path, img_substituta]
          end
      end
      replacements.each do |r|
        @content.sub!(r.first, r.last)
      end
      true
    end

  end
end
