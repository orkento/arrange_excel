require_relative 'arrange_excel.rb'

=begin

=== 説明 ===

#必須
- Windows OS
- Excelアプリケーション
- arrange_excel.rb(同じディレクトリに配置する)
- Rubyが実行できる環境(Rubyコマンドプロンプト、Rubyを入れたCygwin等)

#コマンド例
ruby arrange_excel.exe.rb
->実行フォルダ内のファイルをすべて処理する

ruby arrange_excel.exe.rb file1.xlsx file2.xls dir/
->指定ファイル、または指定ディレクトリ内のファイルすべてを処理する

ruby arrange_excel.exe.rb -r dir/
->指定したディレクトリをサブディレクトリも含め、再帰的に処理する

ruby arrange_excel.exe.rb -v
->実行フォルダ内のファイルをすべて処理し、処理経過を表示する

=end

include ArrangeExcel

argv = ARGV.map{|e| e.downcase}

sub_dir = !argv.delete('-r').nil?
log     = !argv.delete('-v').nil?

extensions = %w(xls xlsx xlsm)

commands = argv.empty? ? ['./'] : ARGV

file_names = commands.inject([]) do |files, command|
               if File.file?(command) && extensions.map{|e| '.' + e}.include?(File.extname(command))
                 files + Dir.glob(command)
               elsif File.directory? command
                 path = if command[-1] != '/'
                          command + '/'
                        else
                          command
                        end
                 files + Dir.glob(path + (sub_dir ? "**/" : "") + "*.{#{extensions.join(',')}}")
               else
                 files
               end
             end

arrange_excel_files file_names, log