#!/usr/bin/env bash

# 生成済み HTML の <table> に.full-width クラスを付与する。
#
# jekyll build は既定で _site を作り直すため、この処理は
# 必ずビルドの「後」に実行すること。lib/make_posts.sh の中から
# 呼ぶと次のビルドで消える。

set -eu

if [ ! -d _site ]; then
  echo "_site が存在しないため何もしません（先に jekyll build を実行してください）" >&2
  exit 0
fi

count_before=$(grep -rl '<table>' _site 2>/dev/null | wc -l)

find _site -type f -name '*.html' -exec \
  sed -i -e 's/<table>/<table class="u-full-width">/g' {} +

echo "fix_tables: ${count_before} ファイルに適用"
