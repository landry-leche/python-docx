"""Enumerations related to Word building blocks (the "glossary document" part)."""

from __future__ import annotations

from .base import BaseXmlEnum


class WD_BUILDING_BLOCK_GALLERY(BaseXmlEnum):
    """Specifies the gallery a building block is classified in.

    Corresponds to the `ST_DocPartGallery` simple type of WordprocessingML, serialized on
    the `w:gallery` element of a building block.

    The integer values of the members are internal to python-docx; they match the
    ordering of the schema type but are not the values of the Word VBA
    `WdBuildingBlockTypes` enumeration.
    """

    PLACEHOLDER = (0, "placeholder", "Structured document tag placeholder text gallery.")
    ANY = (1, "any", "All galleries.")
    DEFAULT = (2, "default", "No gallery classification.")
    DOCUMENT_PARTS = (3, "docParts", "Document parts gallery (\"Quick Parts\" in the Word UI).")
    COVER_PAGE = (4, "coverPg", "Cover page gallery.")
    EQUATIONS = (5, "eq", "Equations gallery.")
    FOOTERS = (6, "ftrs", "Footers gallery.")
    HEADERS = (7, "hdrs", "Headers gallery.")
    PAGE_NUMBERS = (8, "pgNum", "Page numbers gallery.")
    TABLES = (9, "tbls", "Tables gallery.")
    WATERMARKS = (10, "watermarks", "Watermarks gallery.")
    AUTO_TEXT = (11, "autoTxt", "AutoText gallery.")
    TEXT_BOX = (12, "txtBox", "Text box gallery.")
    PAGE_NUMBERS_TOP = (13, "pgNumT", "Page numbers at top gallery.")
    PAGE_NUMBERS_BOTTOM = (14, "pgNumB", "Page numbers at bottom gallery.")
    PAGE_NUMBERS_MARGINS = (15, "pgNumMargins", "Page numbers at margins gallery.")
    TABLE_OF_CONTENTS = (16, "tblOfContents", "Table of contents gallery.")
    BIBLIOGRAPHY = (17, "bib", "Bibliography gallery.")
    CUSTOM_QUICK_PARTS = (18, "custQuickParts", "Custom Quick Parts gallery.")
    CUSTOM_COVER_PAGE = (19, "custCoverPg", "Custom cover page gallery.")
    CUSTOM_EQUATIONS = (20, "custEq", "Custom equations gallery.")
    CUSTOM_FOOTERS = (21, "custFtrs", "Custom footers gallery.")
    CUSTOM_HEADERS = (22, "custHdrs", "Custom headers gallery.")
    CUSTOM_PAGE_NUMBERS = (23, "custPgNum", "Custom page numbers gallery.")
    CUSTOM_TABLES = (24, "custTbls", "Custom tables gallery.")
    CUSTOM_WATERMARKS = (25, "custWatermarks", "Custom watermarks gallery.")
    CUSTOM_AUTO_TEXT = (26, "custAutoTxt", "Custom AutoText gallery.")
    CUSTOM_TEXT_BOX = (27, "custTxtBox", "Custom text box gallery.")
    CUSTOM_PAGE_NUMBERS_TOP = (28, "custPgNumT", "Custom page numbers at top gallery.")
    CUSTOM_PAGE_NUMBERS_BOTTOM = (29, "custPgNumB", "Custom page numbers at bottom gallery.")
    CUSTOM_PAGE_NUMBERS_MARGINS = (30, "custPgNumMargins", "Custom page numbers at margins.")
    CUSTOM_TABLE_OF_CONTENTS = (31, "custTblOfContents", "Custom table of contents gallery.")
    CUSTOM_BIBLIOGRAPHY = (32, "custBib", "Custom bibliography gallery.")
    CUSTOM_1 = (33, "custom1", "Custom gallery 1.")
    CUSTOM_2 = (34, "custom2", "Custom gallery 2.")
    CUSTOM_3 = (35, "custom3", "Custom gallery 3.")
    CUSTOM_4 = (36, "custom4", "Custom gallery 4.")
    CUSTOM_5 = (37, "custom5", "Custom gallery 5.")


class WD_BUILDING_BLOCK_TYPE(BaseXmlEnum):
    """Specifies the type of a building block.

    Corresponds to the `ST_DocPartType` simple type, serialized on the `w:type` children
    of the `w:types` element of a building block.
    """

    NONE = (0, "none", "No type.")
    NORMAL = (1, "normal", "Normal building block.")
    AUTO_EXPAND = (2, "autoExp", "Automatically replace the name with the content.")
    TOOLBAR = (3, "toolbar", "AutoText user interface entry.")
    SPELLER = (4, "speller", "AutoCorrect entry.")
    FORM_FIELD = (5, "formFld", "Form field help text.")
    SDT_PLACEHOLDER = (6, "bbPlcHdr", "Structured document tag placeholder text.")


class WD_BUILDING_BLOCK_BEHAVIOR(BaseXmlEnum):
    """Specifies how a building block is inserted into a document.

    Corresponds to the `ST_DocPartBehavior` simple type, serialized on the `w:behavior`
    children of the `w:behaviors` element of a building block.
    """

    CONTENT = (0, "content", "Insert the content at the specified location.")
    PARAGRAPH = (1, "p", "Ensure the entry is in a new paragraph.")
    PAGE = (2, "pg", "Ensure the entry is on a new page.")
