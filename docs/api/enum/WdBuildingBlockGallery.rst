.. _WdBuildingBlockGallery:

``WD_BUILDING_BLOCK_GALLERY``
=============================

Specifies the gallery a building block is classified in.

Example::

    from docx.enum.building_block import WD_BUILDING_BLOCK_GALLERY

    document.building_blocks.add(
        "Disclaimer",
        "This message is confidential.",
        gallery=WD_BUILDING_BLOCK_GALLERY.AUTO_TEXT,
    )

The integer values of the members are internal to |docx|; they are not the values of the
``WdBuildingBlockTypes`` enumeration in the Word VBA object model.

----

PLACEHOLDER
    Structured document tag placeholder text gallery.

ANY
    All galleries.

DEFAULT
    No gallery classification.

DOCUMENT_PARTS
    Document parts gallery (*Quick Parts* in the Word UI).

COVER_PAGE
    Cover page gallery.

EQUATIONS
    Equations gallery.

FOOTERS
    Footers gallery.

HEADERS
    Headers gallery.

PAGE_NUMBERS
    Page numbers gallery.

TABLES
    Tables gallery.

WATERMARKS
    Watermarks gallery.

AUTO_TEXT
    AutoText gallery.

TEXT_BOX
    Text box gallery.

PAGE_NUMBERS_TOP
    Page numbers at top gallery.

PAGE_NUMBERS_BOTTOM
    Page numbers at bottom gallery.

PAGE_NUMBERS_MARGINS
    Page numbers at margins gallery.

TABLE_OF_CONTENTS
    Table of contents gallery.

BIBLIOGRAPHY
    Bibliography gallery.

CUSTOM_QUICK_PARTS
    Custom Quick Parts gallery.

CUSTOM_COVER_PAGE
    Custom cover page gallery.

CUSTOM_EQUATIONS
    Custom equations gallery.

CUSTOM_FOOTERS
    Custom footers gallery.

CUSTOM_HEADERS
    Custom headers gallery.

CUSTOM_PAGE_NUMBERS
    Custom page numbers gallery.

CUSTOM_TABLES
    Custom tables gallery.

CUSTOM_WATERMARKS
    Custom watermarks gallery.

CUSTOM_AUTO_TEXT
    Custom AutoText gallery.

CUSTOM_TEXT_BOX
    Custom text box gallery.

CUSTOM_PAGE_NUMBERS_TOP
    Custom page numbers at top gallery.

CUSTOM_PAGE_NUMBERS_BOTTOM
    Custom page numbers at bottom gallery.

CUSTOM_PAGE_NUMBERS_MARGINS
    Custom page numbers at margins gallery.

CUSTOM_TABLE_OF_CONTENTS
    Custom table of contents gallery.

CUSTOM_BIBLIOGRAPHY
    Custom bibliography gallery.

CUSTOM_1, CUSTOM_2, CUSTOM_3, CUSTOM_4, CUSTOM_5
    Custom galleries 1 through 5.
