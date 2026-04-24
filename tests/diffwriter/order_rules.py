"""Layer-2 child-ordering rules.

OOXML permits certain siblings to appear in any order (e.g. `<row>` inside
`<sheetData>` can appear in any order as long as the `r` attribute is set),
while others are strictly ordered by schema. This module tells the XML-tree
diff which lists to sort before comparing.

Only sort where the spec ALLOWS any order. Sorting a schema-required
order would mask real bugs.
"""

# element_tag -> (attribute_key_to_sort_by, is_numeric)
#
# Keys are raw element local-names (no namespace prefix). The Layer 2 diff
# normalizes the document, strips namespaces, and applies these rules.
SORT_BY_ATTRIBUTE = {
    "row": ("r", True),
    "c": ("r", False),  # "r" is like "A1", not plain int — sort as str
    "mergeCell": ("ref", False),
    "definedName": ("name", False),
    "sheet": ("sheetId", True),
    "Relationship": ("Id", False),
    "Default": ("Extension", False),
    "Override": ("PartName", False),
    "numFmt": ("numFmtId", True),
    "si": (None, False),  # sharedStrings: order is load-bearing, don't sort
    "xf": (None, False),  # cellXfs index is the identity, don't sort
}

# Elements whose children are never sorted (order is semantic).
PRESERVE_ORDER = {
    "sst",
    "cellXfs",
    "fonts",
    "fills",
    "borders",
    "numFmts",
    "conditionalFormatting",
    "dataValidations",
    "cfRule",
}
