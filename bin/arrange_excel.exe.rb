require_relative '../lib/arrange_excel.rb'

include ArrangeExcel

argv = ARGV.map{|e| e.downcase}

sub_dir = !argv.delete('-r').nil?
log     = !argv.delete('-v').nil?

extensions = %w(xls xlsx xlsm)

commands = argv.empty? ? ['./'] : ARGV

file_names = commands.inject([]) do |files, command|

               if File.file?(command) && extensions.include?(command.split('.').last)
                 files + Dir.glob(command.encode('utf-8'))
               elsif File.directory? command
                 path = if command[-1] != '/'
                          command + '/'
                        else
                          command
                        end
                 files + Dir.glob((path + (sub_dir ? "**/" : "") + "*.{#{extensions.join(',')}}").encode('utf-8'))
               else
                 puts "#{command} is nothing." unless command.chr == '-'
                 files
               end
             end

arrange_excel_files file_names, log