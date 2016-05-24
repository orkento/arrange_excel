require 'win32ole'

class WIN32OLE
  include Enumerable
end

module ArrangeExcel

  # シートのカーソルをA1に合わせます
  # 左上にスクロールします
  # 拡大率を100%にします
  def arrange_worksheet!(sheet, log = false)
    visible = sheet.visible
    sheet.visible = -1
    sheet.activate
    sheet.cells(1, 1).activate
    window = sheet.parent.windows(1)
    window.zoom = 100
    window.scrollRow = 1
    window.scrollColumn = 1
    sheet.visible = visible
    puts("success:#{sheet.parent.name}/#{sheet.name}") if log
  end

  def arrange_worksheet(sheet, log = false)
    success = false
    arrange_worksheet! sheet, log
    success = true
  rescue WIN32OLERuntimeError
    puts("failure:#{sheet.parent.name}/#{sheet.name}") if log
  ensure
    return success
  end

  # ブック内のすべてのシートを整え、最初のシートをアクティブにします
  def arrange_workbook!(book, log = false)
    book.worksheets.reverse_each do |sheet|
      arrange_worksheet! sheet, log
    end
  end

  def arrange_workbook(book, log = false)
    book.worksheets.reverse_each do |sheet|
      arrange_worksheet sheet, log
    end
  end

  # 指定したファイルを整えます
  def arrange_excel_files(file_names, log = false)
    excel = WIN32OLE.new 'Excel.Application'
    excel.DisplayAlerts = false
    file_names.each do |file_name|
      fso = WIN32OLE.new 'Scripting.FileSystemObject'
      absolute_path = fso.GetAbsolutePathName file_name
      puts("file:#{absolute_path}") if log
      
      book = excel.Workbooks.Open absolute_path
      
      arrange_workbook book, log
      
      book.save
      book.close
    end
  rescue WIN32OLERuntimeError
    puts 'Please close this workbook.'
  rescue => e
    puts e
  ensure
    excel.Workbooks.Close
    excel.quit
  end

  def arrnge_excel_file(file_name, log = false)
    arrange_excel_files([file_name], log)
  end
  
end