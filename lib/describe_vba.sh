#!/usr/bin/env bash

rm -f output.txt

ls -1 vba/*.bas | while read line
do
  echo "${line}"
  if [ -e "${line}.md" ]; then
    echo "Skip"
  else
    ys-ollama2file "${line}"
    mv output.txt "${line}".md
  fi
done

rm -f output.txt
