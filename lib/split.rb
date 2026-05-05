# frozen_string_literal: true

current_file = nil
buffer = []

File.foreach("module.bas") do |line|
  line = line.chomp

  if line.start_with?("'#########")
    # 以前のファイルを閉じて書き出し
    if current_file
      buffer.each { |l| current_file.puts(l) }
      current_file.close
      buffer.clear
    end

    name = line.split[1] || "unknown"
    current_file = File.open("vba/#{name}.bas", "w")
  end

  buffer << line if current_file
end

# 最後のファイルを書き出し
if current_file
  buffer.each { |l| current_file.puts(l) }
  current_file.close
end