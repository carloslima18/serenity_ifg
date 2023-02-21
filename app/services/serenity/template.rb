require 'zip/zip'
require 'fileutils'

module Serenity
  class Template
    attr_accessor :template

    def initialize(template, output)
      FileUtils.cp(template, output)
      @template = output
    end

    def process context
      tmpfiles = []
      Zip::ZipFile.open(@template) do |zipfile|
        if template.include?('docx')
          arquivos = %w(word/document.xml)
        else
          arquivos = %w(content.xml styles.xml)
        end
        arquivos.each do |xml_file|
          content = zipfile.read(xml_file)
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
end
