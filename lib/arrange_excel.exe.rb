require_relative 'arrange_excel.rb'

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