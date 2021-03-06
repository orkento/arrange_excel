﻿require 'win32ole'

class WIN32OLE
  include Enumerable
end

module ArrangeExcel

  # シートのカーソルをA1に合わせます
  def select_a1(sheet)
    arrange_active_sheet(sheet){ |s| s.Range("A1").Select }
  end
  # 左上にスクロールします
  def scroll_to_a1(sheet)
    arrange_active_sheet(sheet) do |s|
      window = active_window s
      window.ScrollRow = 1
      window.ScrollColumn = 1
    end
  end
  # 拡大率を100%にします
  def resize_to_100_percent(sheet)
    arrange_active_sheet(sheet) do |s|
      window = active_window s
      window.Zoom = 100
    end
  end
  # オートフィルターを削除します
  def delete_auto_filter(sheet)
    sheet.AutoFilterMode = 0
  end
  
  def arrange_worksheet!(sheet, log = false)
    select_a1 sheet
    scroll_to_a1 sheet
    resize_to_100_percent sheet
    delete_auto_filter sheet
    puts("success:#{sheet.parent.name}/#{sheet.name}") if log
  end

  def arrange_worksheet(sheet, log = false)
    success = false
    arrange_worksheet! sheet, log
    success = true
  rescue WIN32OLERuntimeError
    puts("failure:#{sheet.parent.name}/#{sheet.name}") if log
  rescue => e
    puts e
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