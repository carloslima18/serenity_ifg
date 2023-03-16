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
          else
            arquivos = %w(content.xml styles.xml)
          end
          arquivos.each do |name_file|
            content = zipfile.read(name_file)

            if arquivos_imgs.include?(name_file) and @template.include?('docx')
              WordImage.new(context, content, zipfile).generate_replacements_docx(name_file)
            end

            unless @template.include?('docx')
              OdtImage.new(content, context, zipfile).generate_replacements
            end


            eruby = Erubis::Eruby.new(HTMLEntities.new.decode(content.force_encoding('ASCII-8BIT')
                                                             .force_encoding('UTF-8')), :bufvar => '@_out')


            file_output = eval(eruby.src, context)
            tmpfiles << (file = Tempfile.new("serenity"))
            file << file_output
            file.close
            zipfile.replace(name_file, file.path)
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
