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
            regex = /<draw:image[^>]*xlink:href="([^"]*)"/
            match_data = regex.match(node.to_s)
            href_value = match_data[1]
            replacements << [href_value, img_substituta]
          end
      end
      replacements.each do |r|
        @content.sub!(r.first, r.last)
      end
      true
    end

  end
end
