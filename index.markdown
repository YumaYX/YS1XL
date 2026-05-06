---
layout: default
---



<ul>
{% for post in site.posts reversed %}
<li>
<a href="{{ post.url | relative_url }}">{{ post.title }}</a>
<p>{{ post.content | strip_html | truncate: 120 }}</p>
</li>
{% endfor %}
</ul>
