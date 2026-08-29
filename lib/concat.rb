# frozen_string_literal: true

vba_path = Dir.glob('vba/*.bas')

all_crlf = ''
vba_path.each do |vba_script|
  puts vba_script
  crlf = File.read(vba_script)
  name = File.basename(vba_script, '.bas')
  all_crlf += "'######### #{name}\n#{crlf}"
end

File.write('module.bas', all_crlf.gsub(/\r\n|\r|\n/, "\r\n"))
