# Markdown to Excel

This is a little program to convert a [Markdown][] file to an Open Office XML spreadsheet. It uses [openpyxl][] and [mistune][].

Note: the program requires support for rich text cell content, and actually serves as a "demo" of the feature. As such, you will need a special version of openpyxl (see `requirements.txt`) or the [specific repository][rich_text]

  [Markdown]: https://daringfireball.net/projects/markdown/
  [openpyxl]: https://openpyxl.readthedocs.io/en/default/index.html
  [mistune]: https://github.com/lepture/mistune
  [rich_text]: https://bitbucket.org/Cilyan/openpyxl/branch/2.6

## Known limitations

Some features of Markdown are not supported:
- Tables
- Images
- HTML tags
- Footnotes (for the moment)
- List item continuation after a nested list (is considered as a new list item)
