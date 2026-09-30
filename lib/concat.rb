# frozen_string_literal: true

# vba/*.bas のファイルは Shift_JIS ではなく UTF-8 で保存されている。
# ロケールによっては既定の外部エンコーディングが US-ASCII になり
# gsub が "invalid byte sequence in US-ASCII" で失敗するため、明示的に UTF-8 を指定する。
ENCODING = 'UTF-8'

vba_path = Dir.glob('vba/*.bas').sort

all_crlf = +''
vba_path.each do |vba_script|
  puts vba_script
  crlf = File.read(vba_script, encoding: ENCODING)
  name = File.basename(vba_script, '.bas')
  all_crlf << "'######### #{name}\n" << crlf
end

File.write('module.bas', all_crlf.gsub(/\r\n|\r|\n/, "\r\n"), encoding: ENCODING)
