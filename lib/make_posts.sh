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
title: "$(basename "${f}")"
---
$(cat "${f}")
EOF
done
