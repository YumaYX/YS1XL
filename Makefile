.PHONY: default site posts concat split fix_tables describe cmt

default: site

# Jekyll でプレビューを起動する
site:
	bundle install
	mkdir -p _site
	sh lib/make_posts.sh
	bundle exec jekyll server -H 0.0.0.0 -P 3000

# _posts/ を生成する（vba/*.bas.md から）
posts:
	sh lib/make_posts.sh

split:
	ruby lib/split.rb

concat:
	ruby lib/concat.rb

# ビルド後に <table> へ full-width クラスを付与する
# jekyll build が _site を作り直すため、必ず build の後に実行すること
fix_tables:
	sh lib/fix_tables.sh

# 説明文（.bas.md）が無いファイルだけ AI で生成する
# 外部ツール ys-ollama2file に依存する
describe:
	sh lib/describe_vba.sh

# 変更をステージしてコミットする（メッセージは -m で渡す）
#   make cmt m="修正内容"
cmt:
	git add -A
	git commit -m "$(m)"
