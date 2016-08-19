require 'win32ole'

class WIN32OLE
  include Enumerable
end

module ArrangeExcel

  # シートのカーソルをA1に合わせます
  def select_a1(sheet)
    arrange_active_sheet sheet {|s| s.Range("A1".Select)}
  end
  # 左上にスクロールします
  def scroll_to_a1(sheet)
    visible = sheet.visible
    visible = -1
    sheet.Activate
    window = sheet.Parent.Windows(1)
    window.ScrollRow = 1
    window.ScrollColumn = 1
    sheet.Visible = visible
  end
  # 拡大率を100%にします
  def resize_to_100_percent(sheet)
    visible = sheet.Visible
    visible = -1
    sheet.Activate
    window = sheet.Parent.Windows(1)
    window.Zoom = 100
    sheet.Visible = visible
  end
  
  def delete_auto_filter(sheet)
    sheet.AutoFilterMode = 0
  end
  
  def arrange_worksheet!(sheet, log = false)
    sheet.AutoFilterMode = 0
    visible = sheet.visible
    sheet.visible = -1
    sheet.activate
    sheet.Range("A1").Select
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
  
  private
    def arrange_active_sheet(sheet, &block)
      visible = sheet.visible
      visible = -1
      sheet.Activate
      yield(sheet)
      sheet.Visible = visible
    end
    
    def active_window(sheet)
      sheet.Parent.Windows(1)
    end
end