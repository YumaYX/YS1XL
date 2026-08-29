---
layout: default
---

<section>

<p>再利用可能なExcel VBAの関数集ドキュメント</p>

</section>


<section>

<table>
  <thead>
    <tr>
      <th>タイトル</th>
      <th>概要</th>
    </tr>
  </thead>
  <tbody>
    {% for post in site.posts reversed %}
    <tr>
      <td>
        <a href="{{ post.url | relative_url }}">{{ post.title }}</a>
      </td>
      <td>
        {{ post.content | strip_html | split: '。' | first }}。
      </td>
    </tr>
    {% endfor %}
  </tbody>
</table>

</section>

