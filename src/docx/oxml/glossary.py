"""Custom element classes related to the glossary document, aka. Word "building blocks".

A glossary document is a package part (`/word/glossary/document.xml`) related to the main
document part. Its root is `w:glossaryDocument`, containing a `w:docParts` collection of
`w:docPart` elements, each of which is one building block: a `w:docPartPr` (properties)
followed by a `w:docPartBody` (the content, much like a `w:body`).

Note that a few tag names used inside `w:docPartPr` (`w:style` and `w:type`) are also used
elsewhere in WordprocessingML for unrelated things and are mapped to different custom
element classes by tag name alone. The elements here therefore read and write the `w:val`
attribute of those children directly, rather than relying on their registered classes.
"""

from __future__ import annotations

import uuid
from typing import TYPE_CHECKING, Callable, Iterable, List, cast

from docx.enum.building_block import (
    WD_BUILDING_BLOCK_BEHAVIOR,
    WD_BUILDING_BLOCK_GALLERY,
    WD_BUILDING_BLOCK_TYPE,
)
from docx.oxml.ns import qn
from docx.oxml.parser import OxmlElement
from docx.oxml.shared import CT_String
from docx.oxml.xmlchemy import BaseOxmlElement, RequiredAttribute, ZeroOrMore, ZeroOrOne

if TYPE_CHECKING:
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P

# -- Word's default category name for user-defined building blocks --
DEFAULT_CATEGORY_NAME = "General"


def new_guid() -> str:
    """A new GUID string in the braced, upper-case form Word writes, e.g. `{6F9619FF-...}`."""
    return "{%s}" % str(uuid.uuid4()).upper()


class CT_GlossaryDocument(BaseOxmlElement):
    """`w:glossaryDocument`, root element of the glossary document part."""

    get_or_add_docParts: Callable[[], CT_DocParts]

    docParts: CT_DocParts | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:docParts", successors=()
    )


class CT_DocParts(BaseOxmlElement):
    """`w:docParts`, the collection of building blocks in a glossary document."""

    docPart_lst: List[CT_DocPart]

    docPart = ZeroOrMore("w:docPart")

    def add_doc_part(self, name: str) -> CT_DocPart:
        """Newly added `w:docPart` having `name`, a fresh GUID and one empty paragraph."""
        doc_part = CT_DocPart.new(name)
        self.append(doc_part)
        return doc_part


class CT_DocPart(BaseOxmlElement):
    """`w:docPart`, a single building block: properties plus content."""

    get_or_add_docPartPr: Callable[[], CT_DocPartPr]
    get_or_add_docPartBody: Callable[[], CT_DocPartBody]

    docPartPr: CT_DocPartPr | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:docPartPr", successors=("w:docPartBody",)
    )
    docPartBody: CT_DocPartBody | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:docPartBody", successors=()
    )

    @classmethod
    def new(cls, name: str) -> CT_DocPart:
        """A new `w:docPart` element with the minimum content Word expects.

        That is a name, a GUID and a body containing a single empty paragraph (Word does
        not allow an empty building block).
        """
        doc_part = cast(CT_DocPart, OxmlElement("w:docPart"))
        doc_part_pr = doc_part.get_or_add_docPartPr()
        doc_part_pr.name_val = name
        doc_part_pr.guid_val = new_guid()
        doc_part.get_or_add_docPartBody().add_p()
        return doc_part


class CT_DocPartPr(BaseOxmlElement):
    """`w:docPartPr`, the properties of a building block.

    Children are written in the order Word writes them: name, style, category, types,
    behaviors, description, guid.
    """

    get_or_add_name: Callable[[], CT_String]
    get_or_add_category: Callable[[], CT_DocPartCategory]
    get_or_add_types: Callable[[], CT_DocPartTypes]
    get_or_add_behaviors: Callable[[], CT_DocPartBehaviors]
    get_or_add_description: Callable[[], CT_String]
    get_or_add_guid: Callable[[], CT_String]
    _remove_category: Callable[[], None]
    _remove_types: Callable[[], None]
    _remove_behaviors: Callable[[], None]
    _remove_description: Callable[[], None]
    _remove_guid: Callable[[], None]

    name: CT_String | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:name",
        successors=(
            "w:style",
            "w:category",
            "w:types",
            "w:behaviors",
            "w:description",
            "w:guid",
        ),
    )
    category: CT_DocPartCategory | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:category", successors=("w:types", "w:behaviors", "w:description", "w:guid")
    )
    types: CT_DocPartTypes | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:types", successors=("w:behaviors", "w:description", "w:guid")
    )
    behaviors: CT_DocPartBehaviors | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:behaviors", successors=("w:description", "w:guid")
    )
    description: CT_String | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:description", successors=("w:guid",)
    )
    guid: CT_String | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:guid", successors=()
    )

    # -- name ------------------------------------------------------------------------------

    @property
    def name_val(self) -> str:
        """Value of the `w:name/@w:val` attribute, empty string when there is no name."""
        name = self.name
        return name.val if name is not None else ""

    @name_val.setter
    def name_val(self, value: str):
        self.get_or_add_name().val = value

    # -- category and gallery --------------------------------------------------------------
    # -- (`w:category` requires both `w:name` and `w:gallery`, so one is never added alone) --

    @property
    def category_name(self) -> str | None:
        """Name of the category this building block is filed under, |None| if unclassified."""
        category = self.category
        if category is None or category.name is None:
            return None
        return category.name.val

    @category_name.setter
    def category_name(self, value: str):
        category = self.get_or_add_category()
        category.get_or_add_name().val = value
        if category.gallery is None:
            category.get_or_add_gallery().val = WD_BUILDING_BLOCK_GALLERY.DEFAULT

    @property
    def gallery_val(self) -> WD_BUILDING_BLOCK_GALLERY | None:
        """Gallery this building block is classified in, |None| if unclassified."""
        category = self.category
        if category is None or category.gallery is None:
            return None
        return category.gallery.val

    @gallery_val.setter
    def gallery_val(self, value: WD_BUILDING_BLOCK_GALLERY):
        category = self.get_or_add_category()
        category.get_or_add_gallery().val = value
        if category.name is None:
            category.get_or_add_name().val = DEFAULT_CATEGORY_NAME

    # -- description and guid --------------------------------------------------------------

    @property
    def description_val(self) -> str | None:
        """Value of `w:description/@w:val`, |None| when there is no description."""
        description = self.description
        return description.val if description is not None else None

    @description_val.setter
    def description_val(self, value: str | None):
        if value is None:
            self._remove_description()
            return
        self.get_or_add_description().val = value

    @property
    def guid_val(self) -> str | None:
        """Value of `w:guid/@w:val`, |None| when there is no GUID."""
        guid = self.guid
        return guid.val if guid is not None else None

    @guid_val.setter
    def guid_val(self, value: str | None):
        if value is None:
            self._remove_guid()
            return
        self.get_or_add_guid().val = value

    # -- types and behaviors ---------------------------------------------------------------

    @property
    def type_vals(self) -> tuple[WD_BUILDING_BLOCK_TYPE, ...]:
        """The types of this building block, empty when none are specified."""
        types = self.types
        return types.type_vals if types is not None else ()

    @type_vals.setter
    def type_vals(self, values: Iterable[WD_BUILDING_BLOCK_TYPE]):
        values = tuple(values)
        if not values:
            self._remove_types()
            return
        self.get_or_add_types().type_vals = values

    @property
    def behavior_vals(self) -> tuple[WD_BUILDING_BLOCK_BEHAVIOR, ...]:
        """The insertion behaviors of this building block, empty when none are specified."""
        behaviors = self.behaviors
        return behaviors.behavior_vals if behaviors is not None else ()

    @behavior_vals.setter
    def behavior_vals(self, values: Iterable[WD_BUILDING_BLOCK_BEHAVIOR]):
        values = tuple(values)
        if not values:
            self._remove_behaviors()
            return
        self.get_or_add_behaviors().behavior_vals = values


class CT_DocPartCategory(BaseOxmlElement):
    """`w:category`, the category and gallery a building block is filed under."""

    get_or_add_name: Callable[[], CT_String]
    get_or_add_gallery: Callable[[], CT_DocPartGallery]

    name: CT_String | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:name", successors=("w:gallery",)
    )
    gallery: CT_DocPartGallery | None = ZeroOrOne(  # pyright: ignore[reportAssignmentType]
        "w:gallery", successors=()
    )


class CT_DocPartGallery(BaseOxmlElement):
    """`w:gallery`, having a `w:val` attribute from the `ST_DocPartGallery` enumeration."""

    val: WD_BUILDING_BLOCK_GALLERY = RequiredAttribute(  # pyright: ignore[reportAssignmentType]
        "w:val", WD_BUILDING_BLOCK_GALLERY
    )


class CT_DocPartTypes(BaseOxmlElement):
    """`w:types`, containing one `w:type` child for each type of the building block.

    The `w:type` tag is shared with the section-type element, which is why the children are
    accessed through their raw `w:val` attribute.
    """

    @property
    def type_vals(self) -> tuple[WD_BUILDING_BLOCK_TYPE, ...]:
        return tuple(
            WD_BUILDING_BLOCK_TYPE.from_xml(val)
            for val in (t.get(qn("w:val")) for t in self.findall(qn("w:type")))
            if val is not None
        )

    @type_vals.setter
    def type_vals(self, values: Iterable[WD_BUILDING_BLOCK_TYPE]):
        for child in self.findall(qn("w:type")):
            self.remove(child)
        for value in values:
            child = OxmlElement("w:type")
            child.set(qn("w:val"), cast(str, WD_BUILDING_BLOCK_TYPE.to_xml(value)))
            self.append(child)


class CT_DocPartBehaviors(BaseOxmlElement):
    """`w:behaviors`, containing one `w:behavior` child for each insertion behavior."""

    @property
    def behavior_vals(self) -> tuple[WD_BUILDING_BLOCK_BEHAVIOR, ...]:
        return tuple(
            WD_BUILDING_BLOCK_BEHAVIOR.from_xml(val)
            for val in (b.get(qn("w:val")) for b in self.findall(qn("w:behavior")))
            if val is not None
        )

    @behavior_vals.setter
    def behavior_vals(self, values: Iterable[WD_BUILDING_BLOCK_BEHAVIOR]):
        for child in self.findall(qn("w:behavior")):
            self.remove(child)
        for value in values:
            child = OxmlElement("w:behavior")
            child.set(qn("w:val"), cast(str, WD_BUILDING_BLOCK_BEHAVIOR.to_xml(value)))
            self.append(child)


class CT_DocPartBody(BaseOxmlElement):
    """`w:docPartBody`, the content of a building block.

    A block-item container like `w:body`: it holds paragraphs and tables (and, rarely, a
    final `w:sectPr`).
    """

    add_p: Callable[[], CT_P]
    p_lst: List[CT_P]
    tbl_lst: List[CT_Tbl]

    _insert_tbl: Callable[[CT_Tbl], CT_Tbl]

    p = ZeroOrMore("w:p", successors=("w:sectPr",))
    tbl = ZeroOrMore("w:tbl", successors=("w:sectPr",))

    @property
    def inner_content_elements(self) -> List[CT_P | CT_Tbl]:
        """All `w:p` and `w:tbl` child elements, in document order."""
        return self.xpath("./w:p | ./w:tbl")

    @property
    def content_elements(self) -> List[BaseOxmlElement]:
        """Every child element that is content, in order: all but the final `w:sectPr`.

        Unlike `.inner_content_elements` this includes block-level content controls
        (`w:sdt`), bookmarks and any other block-level element, which a faithful copy of
        the block needs.
        """
        return [
            child
            for child in self
            if isinstance(child.tag, str) and child.tag != qn("w:sectPr")
        ]
