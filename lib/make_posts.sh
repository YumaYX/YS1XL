#!/usr/bin/env bash

rm -rf _posts
mkdir -p _posts

for f in $(ls -1 vba/*.md)
do
  echo "${f}"
  cat <<EOF > _posts/1999-12-31-$(basename "${f}").md
---
layout: post
category: 
title: "$(basename "${f%%.md}")"
---

$(cat "${f}")


<https://github.com/YumaYX/YS1XL/>

<small>本内容はAIによる自動生成のため、正確性が保証されていません。参考情報としてご利用ください。</small>

EOF
done

find _site -type f -exec sed -i -e 's/<table>/<table class="u-full-width">/g' {} +
