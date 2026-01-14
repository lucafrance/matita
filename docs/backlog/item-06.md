---
title: item-06 Add support for `Worksheet.Rows(index)` ans `Worksheet.Columns(index)`
---

```
# Works now
wks.rows.item(2).style = "Heading 1"

# Not supported
wks.rows(2).style = "Heading 1"
```
