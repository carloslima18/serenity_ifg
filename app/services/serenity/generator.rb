module Serenity
  module Generator
    def render_odt template_path, output_path = output_name(template_path)
      template = Template.new template_path, output_path
      template.process binding
    end

    private

    def output_name input
      extension = input.match(/(\.)[a-z]+/).to_s
      input_output = input.gsub(extension, '')
      "#{input_output}_output#{extension}"
    end
  end
end
