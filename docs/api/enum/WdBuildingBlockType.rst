.. _WdBuildingBlockType:

``WD_BUILDING_BLOCK_TYPE``
==========================

Specifies the type of a building block. A building block can have more than one type.

Example::

    from docx.enum.building_block import WD_BUILDING_BLOCK_TYPE

    block = document.building_blocks.get("Sign-off")
    block.types = (WD_BUILDING_BLOCK_TYPE.NORMAL,)

----

NONE
    No type.

NORMAL
    Normal building block.

AUTO_EXPAND
    Automatically replace the name with the content.

TOOLBAR
    AutoText user interface entry.

SPELLER
    AutoCorrect entry.

FORM_FIELD
    Form field help text.

SDT_PLACEHOLDER
    Structured document tag placeholder text.
