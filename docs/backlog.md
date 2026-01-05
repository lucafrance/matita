## item-1 Add snake case aliases for all methods and properties

VBA properties and methods are written in [CamelCase](https://developer.mozilla.org/en-US/docs/Glossary/Camel_case).
In Python, the [PEP8](https://peps.python.org/pep-0008/#function-and-variable-names) naming convention is [snake_case](https://developer.mozilla.org/en-US/docs/Glossary/Snake_case).

> Function names should be lowercase, with words separated by underscores as necessary to improve readability.

Add snake_case aliases for all methods and properties to improve intuitiveness for Python developers.

```python
from matita.office import excel
wks = excel.application().new().workbooks.open("file.xlsx").worksheets(1)

# Supported
wks.AutoFilter
wks.autofilter()

# New
wks.auto_filter
```

## item-2 Parse *data type* of parameters

The column `data type` of the parameters table is not yet parsed.
E.g. from [Sequence.AddEffect method (PowerPoint)](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.sequence.addeffect).

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Shape_|Required|**[Shape](PowerPoint.Shape.md)**|The shape to which the animation effect is added.|
| _effectId_|Required|**[MsoAnimEffect](PowerPoint.MsoAnimEffect.md)**|The animation effect to be applied.|
| _Level_|Optional|**[MsoAnimateByLevel](PowerPoint.MsoAnimateByLevel.md)**|For charts, diagrams, or text, the level to which the animation effect will be applied. The default value is **msoAnimationLevelNone**.|
| _trigger_|Optional|**[MsoAnimTriggerType](PowerPoint.MsoAnimTriggerType.md)**|The action that triggers the animation effect. The default value is **msoAnimTriggerOnPageClick**.|
| _Index_|Optional|**Long**|The position at which the effect will be placed in the collection of animation effects. The default value is -1 (added to the end). |

The parsed information can be used to improve the type hints in the generated code.
E.g. the method signature of `AddEffect` can be improved from

```python
   def AddEffect(self, Shape=None, effectId=None, Level=None, trigger=None, Index=None):
```
to

```python
   def AddEffect(self, Shape: Shape = None, effectId: int = None, Level: int = None, trigger: int = None, Index: int = None)
```

Enums are defined as `int`.

## item-3 Parse *default value* of parameters

Follow-up to item-2.
More complex, because the default value is embedded in the description text.
