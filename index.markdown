---
layout: default
---

<table>
  <thead>
    <tr>
      <th>Function</th>
      <th>Description</th>
    </tr>
  </thead>

  <tbody>
    {% for post in site.posts %}
    <tr>

      <td>
        <a href="{{ post.url | relative_url }}">
          <strong>{{ post.title }}</strong>
        </a>
      </td>

      <td>
{{ post.content | strip_html | truncate: 120 }}
      </td>

    </tr>
    {% endfor %}
  </tbody>
</table>