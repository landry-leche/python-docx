.. _WdBuildingBlockBehavior:

``WD_BUILDING_BLOCK_BEHAVIOR``
==============================

Specifies how Word inserts a building block into a document. A building block can have
more than one behavior.

Example::

    from docx.enum.building_block import WD_BUILDING_BLOCK_BEHAVIOR

    block = document.building_blocks.get("Sign-off")
    block.behaviors = (WD_BUILDING_BLOCK_BEHAVIOR.PARAGRAPH,)

----

CONTENT
    Insert the content at the specified location.

PARAGRAPH
    Ensure the entry is in a new paragraph.

PAGE
    Ensure the entry is on a new page.
