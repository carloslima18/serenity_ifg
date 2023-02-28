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

    def substituir_variaveis_em_arquivo_xlsx
      file_path = @file_path
      context = @context
      # Abre o arquivo XLSX usando a gem "roo-xls"
      workbook = Roo::Spreadsheet.open(file_path)

      # Inicializa um objeto "Axlsx" para escrever o arquivo XLSX modificado
      output_workbook = Roo::Spreadsheet.open(@file_path_output)
      if output_workbook.sheets.include?('Planilha1')
        output_worksheet = output_workbook.sheet('Planilha1')
      end
      # Itera sobre todas as planilhas do arquivo

      coluna = 0
      workbook.each_with_index do |sheet_name, row|
        sheet_name.each_with_index do |cell, col|
          if cell =~ /<%.*%>/
            begin
              # Avalia o código Ruby contido na célula
              codigo_ruby = cell.sub(/<%=/, '').sub(/<%/, '').sub(/%>/, '')
              resultado = eval(codigo_ruby, context)
              # Substitui a célula com o resultado da avaliação
              output_worksheet.set(row, col, resultado) unless resultado.nil?
              # output_worksheet.cell(0,0)
            rescue Exception => e
              puts "Erro ao substituir as variáveis na célula #{sheet_name}!#{Axlsx::cell_r(col, row)}: #{e.message}"
            end
          else
            # Copia a célula para o arquivo de saída sem modificação
            output_worksheet.set(row, col, resultado) unless resultado.nil?
          end
        end
      end
      # Salva o arquivo de saída na pasta "public" da aplicação
      # output_workbook.write(@file_path_output)
    end

    def compile_excel
      workbook = Roo::Spreadsheet.open(@file_path)
      sheet = workbook.sheet(0)

      # Lê cada linha da planilha e concatena em uma única string
      codigo_rails = ""
      sheet.each do |linha|
        codigo_rails << linha.join("\n") + "\n"
      end

      # Compila o código Ruby
      begin
        # ERB.new(codigo_rails).result(@context)
        resultado = ERB.new(codigo_rails).result(@context)
        # resultado = eval(codigo_rails, @context)
      rescue Exception => e
        puts "Erro ao compilar o código: #{e.message}"
        resultado = "Erro ao compilar o código: #{e.message}"
      end

      # Salva o resultado no arquivo
      output_path = File.join(Rails.public_path, 'aaaaaaaaaaaaaa_output.xlsx')
      File.open(output_path, 'w') do |file|
        file.write(resultado)
      end
    end

    def compilar_codigo_em_arquivo_xlsx
      arquivo = @file_path
      contexto = @context
      # Abre o arquivo Excel com a gem roo
      planilha = Roo::Excelx.new(arquivo)

      # Cria um novo arquivo Excel com a gem write_xlsx
      novo_arquivo = WriteXLSX.new("public/#{File.basename(arquivo)}")

      # Para cada planilha da planilha original, cria uma nova planilha no novo arquivo
      planilha.sheets.each do |sheet_name|
        # Cria um objeto para manipular a planilha original
        sheet = planilha.sheet(sheet_name)

        # Cria um objeto para manipular a nova planilha
        nova_sheet = novo_arquivo.add_worksheet(sheet_name)

        # Copia as configurações de formatação da planilha original para a nova planilha
        (sheet.first_row..sheet.last_row).each do |row_index|
          (sheet.first_column..sheet.last_column).each do |col_index|
            cell = sheet.cell(row_index, col_index)
            xf_index = cell.style
            xf = planilha.cell_xfs[xf_index]
            nova_sheet.write(row_index - 1, col_index - 1, cell, xf)
          end
        end

        # Para cada célula da planilha
        (sheet.first_row..sheet.last_row).each do |row_index|
          (sheet.first_column..sheet.last_column).each do |col_index|
            value = sheet.cell(row_index, col_index).to_s

            # Verifica se a célula contém um código Ruby e substitui as variáveis
            if value.start_with?("<%")
              # Remove as tags <% e %>
              codigo_ruby = value.gsub("<%", "").gsub("%>", "").strip

              # Executa o código Ruby no contexto da aplicação
              resultado = eval(codigo_ruby, contexto)

              # Escreve o resultado na nova planilha
              nova_sheet.write(row_index - 1, col_index - 1, resultado)
            elsif value != ""
              # Copia o valor da célula original para a nova planilha
              nova_sheet.write(row_index - 1, col_index - 1, value)
            end
          end
        end
      end

      # Salva e fecha o novo arquivo Excel
      novo_arquivo.close
    end

  end
end
