import string

import pythoncom

def com_arguments(arguments: list) -> list:
    """Ensures that optional arguments are correctly represented for COM calls.
    
    Optional arguments should be `None` until there are defined arguments.
    After the last define argument, all subsequent optional arguments should be
    represented as `pythoncom.Missing`.
    """
    com_arguments = []
    replacement_value = pythoncom.Missing
    for p in reversed(arguments):
        if p is None:
            com_arguments.insert(0, replacement_value)
        else:
            com_arguments.insert(0, p)
            replacement_value = None
    return com_arguments

def unwrap(obj):
    """Return the underlying COM object if available"""
    
    if hasattr(obj, "com_object"):
        return obj.com_object
    return obj

def camel_to_snake_case(s):
    snake_s =  "".join(
        [f"_{c.lower()}" if c in string.ascii_uppercase else c for c in s]
    )
    snake_s = snake_s.strip("_")
    return snake_s
