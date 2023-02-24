module Serenity
  class XlsxCompiler
    attr_accessor :file_path, :context, :file_path_output

    def initialize(file_path, context, file_path_output)
      @file_path = file_path
      @context = context
      @file_path_output = file_path_output
    end

    def evaluate_excel_file
      path = @file_path
      context = @context
      file_path_output = @file_path_output
      p = Axlsx::Package.new
      wb = p.workbook
      sheet = wb.add_worksheet(name: 'Sheet1')

      xlsx = Roo::Spreadsheet.open(path)
      xlsx.default_sheet = xlsx.sheets.first

      (1..xlsx.last_row).each do |row|
        linha = []
        (1..xlsx.last_column).each do |col|
          cell_value = xlsx.cell(row, col)

          if cell_value.is_a?(String) && cell_value =~ /<%.*%>/
            cell_value = cell_value.sub(/<%=/, '').sub(/<%/, '').sub(/%>/, '')
            begin
              cell_value = eval(cell_value, context)
            rescue Exception => e
              sheet.add_row(["Error: #{e.message}"])
            end
          end
          linha << cell_value
        end
        sheet.add_row(linha)
      end
      p.serialize(file_path_output)
    end
  end
end
