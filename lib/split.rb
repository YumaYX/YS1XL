# frozen_string_literal: true

# lib/concat.rb と同じく UTF-8 を明示的に指定する。
ENCODING = 'UTF-8'
# VBE と同じ CRLF で書き出す。.gitattributes の *.bas eol=crlf とも整合させる。
EOL = "\r\n"

current_file = nil
buffer = []

File.foreach('module.bas', encoding: ENCODING) do |line|
  line = line.chomp

  if line.start_with?("'#########")
    # 以前のファイルを閉じて書き出し
    if current_file
      buffer.each { |l| current_file.write(l, EOL) }
      current_file.close
      buffer.clear
    end

    name = line.split[1] || 'unknown'
    current_file = File.open("vba/#{name}.bas", 'wb')
    next
  end

  buffer << line if current_file
end

# 最後のファイルを書き出し
if current_file
  buffer.each { |l| current_file.write(l, EOL) }
  current_file.close
end
