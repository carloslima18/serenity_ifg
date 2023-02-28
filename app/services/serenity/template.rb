require 'zip/zip'
require 'fileutils'
require 'docx'

module Serenity
  class Template
    attr_accessor :template

    def initialize(template, output)
      FileUtils.cp(template, output)
      @template = template
      @template_output = output
    end

    def process context
      tmpfiles = []
      Zip::ZipFile.open(@template_output) do |zipfile|

        arquivos_imgs = %w(word/_rels/document.xml.rels [Content_Types].xml)


        if is_xlsx_file?(@template)
          Serenity::XlsxCompiler.new(@template, context, @template_output).evaluate_excel_file
        else
          if @template.include?('docx')
            arquivos = %w(word/document.xml word/_rels/document.xml.rels [Content_Types].xml)
            # <Override PartName="/word/media/image1.png" ContentType="image/png"/>
            # /word/_rels/document.xml.rels
            #
          else
            arquivos = %w(content.xml styles.xml)
          end
          arquivos.each do |xml_file|
            content = zipfile.read(xml_file)


          if arquivos_imgs.include?(xml_file)
            content = Serenity::WordImage.new(@template, context, content, @template_output).adicionar_imagem_no_word(1, 2, xml_file, zipfile)
          end
          content


            #add img tentativa 1
            unless @template.include?('docx')
              unless zipfile.find_entry('Pictures/logo.png')
                zipfile.add('Pictures/logo.png',  ::File.join(Rails.public_path, 'logo.png'))
              end
              images_replacements = ImagesProcessor.new(content, context).generate_replacements
              images_replacements.each do |r|
                content = content.gsub(r.first, "Pictures/#{r.last}")
              end
            end


            eruby = Erubis::Eruby.new(content.force_encoding('ASCII-8BIT').force_encoding('UTF-8')
                                             .gsub('&lt;%=', '<%=').gsub('&lt;%', '<%').gsub('%&gt;', '%>'), :bufvar=>'@_out')
            out = eval(eruby.src, context)
            tmpfiles << (file = Tempfile.new("serenity"))
            file << out
            file.close
            zipfile.replace(xml_file, file.path)
          end
        end
      end
    end

    def is_xlsx_file?(path)
      extension = File.extname(path)
      extension == ".xlsx"
    end

  end
end
