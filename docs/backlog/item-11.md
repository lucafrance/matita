---
title: item-11 Fix invalid collection return class types
---

E.g. There is no class `Excel.Area`. Is `Range.Areas` a collection of Ranges?

Could be fixed by updating the documentation.
If the return class of the `Item` method of a collection is defined, it should be prioritised over the class inferred by the name.

```
class Areas:

    def __init__(self, areas=None):
        self.com_object= areas

    def __call__(self, item):
        return Area(self.com_object(item))
```
