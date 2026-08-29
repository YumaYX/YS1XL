---
layout: default
---

<section>

<p>再利用可能なExcel VBAの関数集ドキュメント</p>

</section>

<section>
<ul>
{% for post in site.posts reversed %}
<li>
<a href="{{ post.url | relative_url }}">{{ post.title }}</a>: {{ post.content | strip_html | split: '。' | first }}。
</li>
{% endfor %}
</ul>
</section>

