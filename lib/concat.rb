# frozen_string_literal: true

vba_path = Dir.glob('vba/*')

all_crlf = ''
vba_path.each do |vba_script|
  puts vba_script
  crlf = File.read(vba_script)
  all_crlf += "\n#{crlf}"
end

File.write('module.bas', all_crlf.gsub(/\r\n|\r|\n/, "\r\n"))
