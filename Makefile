default:
	bundle exec jekyll server -H 0.0.0.0 -P 3000
cat:
	cat Makefile

split:
	ruby lib/split.rb

concat:
	ruby lib/concat.rb

describe:
	time sh lib/describe_vba.sh

posts:
	sh lib/make_posts.sh

cmt:
	git add .
	git commit --allow-empty-message -am ""

commit: cmt concat
	make cmt
