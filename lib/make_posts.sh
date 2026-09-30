#!/usr/bin/env bash

set -eu

rm -rf _posts
mkdir -p _posts

for f in vba/*.md
do
  echo "${f}"

  base="${f##*/}"      # IPaddress.bas.md
  name="${base%.md}"   # IPaddress.bas

  # 出力ファイル名に .md を二重に付けないよう "${f##*/}" をそのまま使う
  cat <<EOF > "_posts/1999-12-31-${base}"
---
layout: post
category: 
title: "${name}"
---

$(cat "${f}")

[${name}](https://github.com/YumaYX/YS1XL/blob/main/vba/${name})

<small>※本内容はAIによる自動生成のため、正確性が保証されていません。参考情報としてご利用ください。</small>

EOF
done
