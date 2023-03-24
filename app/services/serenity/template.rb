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

        if docx_file?(@template)
          arquivos_imgs = %w(word/_rels/document.xml.rels [Content_Types].xml)
          arquivos = %w(word/document.xml word/_rels/document.xml.rels [Content_Types].xml)
        elsif xlsx_file?(@template)
          arquivos = %w(xl/sharedStrings.xml)
        else
          arquivos_imgs = %w(content.xml)
          arquivos = %w(content.xml styles.xml)
        end

        arquivos.each do |name_file|
          content = zipfile.read(name_file)

          if docx_file?(@template) and arquivos_imgs.include?(name_file)
            DocxImage.new(context, content, zipfile).generate_replacements_docx(name_file)
          end

          if odt_file?(@template) and arquivos_imgs.include?(name_file)
            OdtImage.new(content, context, zipfile).generate_replacements
          end

          eruby = Erubis::Eruby.new(HTMLEntities.new.decode(content.force_encoding('ASCII-8BIT')
                                                                   .force_encoding('UTF-8')), :bufvar => '@_out', :pattern => '{% %}')

          file_output = eval(eruby.src, context)
          tmpfiles << (file = Tempfile.new("serenity"))
          file << file_output
          file.close
          zipfile.replace(name_file, file.path)
        end
      end
    end

    def xlsx_file?(path)
      File.extname(path) == ".xlsx"
    end

    def docx_file?(path)
      File.extname(path) == ".docx"
    end

    def odt_file?(path)
      File.extname(path) == ".odt"
    end

    def process_odt_eruby context
      tmpfiles = []
      Zip::ZipFile.open(@template_output) do |zipfile|
        if @template_output.include?('docx')
          arquivos = %w(word/document.xml)
        else
          arquivos = %w(content.xml styles.xml)
        end

        arquivos.each do |xml_file|
          content = zipfile.read(xml_file)
          odteruby = OdtEruby.new(XmlReader.new(content.force_encoding('ASCII-8BIT').force_encoding('UTF-8')))
          out = odteruby.evaluate(context)

          tmpfiles << (file = Tempfile.new("serenity"))
          file << out
          file.close

          zipfile.replace(xml_file, file.path)
        end
      end
    end

  end
end
