"""Collection providing access to the Word building blocks of a document.

Word calls these "building blocks" (Quick Parts, AutoText, cover pages, headers, footers,
equations, and so on); in the file format they are `w:docPart` elements of the *glossary
document*, a part related to the main document part.
"""

from __future__ import annotations

import copy
import io
from typing import TYPE_CHECKING, Iterable, Iterator, Sequence, cast

from docx.blkcntnr import BlockItemContainer
from docx.enum.building_block import (
    WD_BUILDING_BLOCK_BEHAVIOR,
    WD_BUILDING_BLOCK_GALLERY,
    WD_BUILDING_BLOCK_TYPE,
)
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.glossary import DEFAULT_CATEGORY_NAME
from docx.oxml.ns import nsmap, qn
from docx.oxml.parser import OxmlElement
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph

if TYPE_CHECKING:
    from docx.document import Document
    from docx.oxml.glossary import CT_DocPart
    from docx.oxml.xmlchemy import BaseOxmlElement
    from docx.parts.document import DocumentPart
    from docx.parts.glossary import GlossaryPart
    from docx.parts.story import StoryPart

_R_NS_PREFIX = "{%s}" % nsmap["r"]

# -- children of a story that hold properties rather than content --
_KEPT_WHEN_REPLACING = frozenset((qn("w:sectPr"), qn("w:tcPr")))

# -- the elements `append_to()` knows to be block-item containers --
_BLOCK_CONTAINER_TAGS = frozenset(
    qn(tag) for tag in ("w:body", "w:tc", "w:hdr", "w:ftr", "w:comment", "w:docPartBody")
)


class BuildingBlocks:
    """Collection of the building blocks (Quick Parts, AutoText, etc.) of a document.

    Obtained from `document.building_blocks`. Reading never modifies the document: the
    glossary part that stores building blocks is only created when the first one is added.
    """

    def __init__(self, document_part: DocumentPart):
        self._document_part = document_part

    def __iter__(self) -> Iterator[BuildingBlock]:
        """Iterate over the building blocks in this collection, in file order."""
        glossary_part = self._document_part.glossary_part
        if glossary_part is None:
            return
        doc_parts = glossary_part.element.docParts
        if doc_parts is None:
            return
        for doc_part in doc_parts.docPart_lst:
            yield BuildingBlock(doc_part, glossary_part)

    def __len__(self) -> int:
        """The number of building blocks in this collection."""
        return sum(1 for _ in self)

    def add(
        self,
        name: str,
        text: str = "",
        *,
        gallery: WD_BUILDING_BLOCK_GALLERY = WD_BUILDING_BLOCK_GALLERY.DOCUMENT_PARTS,
        category: str = DEFAULT_CATEGORY_NAME,
        description: str | None = None,
        types: Iterable[WD_BUILDING_BLOCK_TYPE] = (WD_BUILDING_BLOCK_TYPE.NORMAL,),
        behaviors: Iterable[WD_BUILDING_BLOCK_BEHAVIOR] = (WD_BUILDING_BLOCK_BEHAVIOR.CONTENT,),
    ) -> BuildingBlock:
        """Add a new building block and return it.

        `name` must be unique within its `gallery` and `category`; |ValueError| is raised
        otherwise. The defaults produce what Word calls a custom Quick Part: gallery "Quick
        Parts" (`DOCUMENT_PARTS`), category "General", "Insert content" behavior.

        The block always starts with one paragraph (Word does not allow an empty building
        block). If `text` is provided it is placed in that paragraph, and each additional
        line (separated by `"\\n"`) becomes a further paragraph. Richer content can be added
        afterward with the returned block's `add_paragraph()`, `add_table()` and the
        paragraphs it exposes.
        """
        block = self._add_block(name, gallery, category, description, types, behaviors)
        if text:
            first_line, *more_lines = text.split("\n")
            block.paragraphs[0].add_run(first_line)
            for line in more_lines:
                block.add_paragraph(line)
        return block

    def add_from(
        self,
        name: str,
        items: Sequence[Paragraph | Table],
        *,
        gallery: WD_BUILDING_BLOCK_GALLERY = WD_BUILDING_BLOCK_GALLERY.DOCUMENT_PARTS,
        category: str = DEFAULT_CATEGORY_NAME,
        description: str | None = None,
        types: Iterable[WD_BUILDING_BLOCK_TYPE] = (WD_BUILDING_BLOCK_TYPE.NORMAL,),
        behaviors: Iterable[WD_BUILDING_BLOCK_BEHAVIOR] = (WD_BUILDING_BLOCK_BEHAVIOR.CONTENT,),
    ) -> BuildingBlock:
        """Add a new building block holding a copy of `items` and return it.

        This is the equivalent of selecting content in Word and choosing "Save Selection to
        Quick Part Gallery". `items` is a non-empty sequence of |Paragraph| and |Table|
        objects, typically taken from `document.paragraphs`, `document.tables` or
        `document.iter_inner_content()`. The originals are left untouched. Images and
        hyperlinks in the content are carried over along with their relationships.
        """
        if not items:
            raise ValueError("`items` must contain at least one paragraph or table")
        block = self._add_block(name, gallery, category, description, types, behaviors)
        body = block._element  # pyright: ignore[reportPrivateUsage]
        for p in list(body.p_lst):
            body.remove(p)
        for item in items:
            clone = copy.deepcopy(item._element)  # pyright: ignore[reportPrivateUsage]
            _remap_relationships([clone], item.part, block.part)
            body.append(clone)
        # -- Word requires a story to end with a paragraph, never a table --
        if body.inner_content_elements and not isinstance(body.inner_content_elements[-1], CT_P):
            body.add_p()
        return block

    def find(
        self,
        name: str | None = None,
        *,
        gallery: WD_BUILDING_BLOCK_GALLERY | None = None,
        category: str | None = None,
    ) -> list[BuildingBlock]:
        """List of building blocks matching all the criteria given, in file order.

        A criterion left as |None| matches anything. `name` and `category` are compared
        exactly (case-sensitive).
        """
        return [
            block
            for block in self
            if (name is None or block.name == name)
            and (gallery is None or block.gallery == gallery)
            and (category is None or block.category == category)
        ]

    def get(
        self,
        name: str,
        *,
        gallery: WD_BUILDING_BLOCK_GALLERY | None = None,
        category: str | None = None,
    ) -> BuildingBlock | None:
        """The first building block named `name`, or |None| if there isn't one.

        Names are unique only within a gallery and category, so pass `gallery` and/or
        `category` to disambiguate when needed.
        """
        matches = self.find(name, gallery=gallery, category=category)
        return matches[0] if matches else None

    def remove(self, block: BuildingBlock) -> None:
        """Remove `block` from the document.

        Relationships (images, hyperlinks) that only that building block used are removed
        along with it. The `block` object must not be used afterward.
        """
        glossary_part = self._document_part.glossary_part
        doc_part = block._doc_part  # pyright: ignore[reportPrivateUsage]
        parent = doc_part.getparent()
        if glossary_part is None or parent is None or parent is not (
            glossary_part.element.docParts
        ):
            raise ValueError("`block` is not a building block of this document")

        released = _referenced_rIds(doc_part)
        parent.remove(doc_part)
        still_used = _referenced_rIds(glossary_part.element)
        for rId in released - still_used:
            if rId in glossary_part.rels:
                del glossary_part.rels[rId]

    def _add_block(
        self,
        name: str,
        gallery: WD_BUILDING_BLOCK_GALLERY,
        category: str,
        description: str | None,
        types: Iterable[WD_BUILDING_BLOCK_TYPE],
        behaviors: Iterable[WD_BUILDING_BLOCK_BEHAVIOR],
    ) -> BuildingBlock:
        """A newly-added, minimally-populated building block (one empty paragraph)."""
        if not name:
            raise ValueError("a building block must have a non-empty name")
        if self.get(name, gallery=gallery, category=category) is not None:
            raise ValueError(
                f"a building block named {name!r} already exists in gallery {gallery.name}"
                f" and category {category!r}"
            )
        glossary_part = self._document_part.get_or_add_glossary_part()
        doc_parts = glossary_part.element.get_or_add_docParts()
        doc_part = doc_parts.add_doc_part(name)
        block = BuildingBlock(doc_part, glossary_part)
        block.gallery = gallery
        block.category = category
        block.description = description
        block.types = tuple(types)
        block.behaviors = tuple(behaviors)
        return block


class BuildingBlock(BlockItemContainer):
    """A single building block: a named, categorized piece of reusable content.

    Its content is accessed like that of a table cell or header, with `.paragraphs`,
    `.tables`, `.add_paragraph()`, `.add_table()` and `.iter_inner_content()`.
    """

    def __init__(self, doc_part: CT_DocPart, glossary_part: GlossaryPart):
        body = doc_part.docPartBody
        if body is None:
            body = doc_part.get_or_add_docPartBody()
            body.add_p()
        super().__init__(cast("BaseOxmlElement", body), glossary_part)  # pyright: ignore
        self._doc_part = doc_part

    def __eq__(self, other: object) -> bool:
        return isinstance(other, BuildingBlock) and other._doc_part is self._doc_part

    def __hash__(self) -> int:
        return id(self._doc_part)

    def add_table(self, rows: int, cols: int, width=None) -> Table:  # pyright: ignore
        """Table of `rows` x `cols` added to the end of this building block.

        Like a table cell, a building block must end with a paragraph, so an empty one
        is added after the table. `width` is distributed evenly across the columns and
        defaults to 6 inches.
        """
        from docx.shared import Inches

        table = super().add_table(rows, cols, width if width is not None else Inches(6))
        self.add_paragraph()
        return table

    def append_to(
        self, container: BlockItemContainer | Document, *, replace: bool = False
    ) -> list[Paragraph | Table]:
        """Copy this building block's content to the end of `container`, return the copies.

        `container` is the |Document| or anything else that holds block items: a table cell,
        a header or footer, a comment, another building block. In a document body the
        content goes before the final section properties, so it lands at the end of the
        last section.

        With `replace=True` the existing content of `container` is removed first, so that
        the block *becomes* the content. That is what Word does when you pick a block from
        its Header or Footer gallery. It is not allowed for the document body itself
        (|ValueError|). Whatever the container, a story always ends with a paragraph: if
        the block's content would leave it ending with a table, an empty paragraph is added.

        Beware of headers and footers that are *linked to the previous section*: reading
        or writing one reaches the header of the section it inherits from, so appending to
        every section's header would add the block to the same header several times. Check
        `is_linked_to_previous` first, as shown in the user guide.

        This is a block-level insertion: every paragraph, table and content control of the
        block is copied as-is, the block's `behaviors` are not interpreted. Images and
        hyperlinks are carried over with their relationships. Styles the content uses that the
        destination does not have are imported into it, as Word does; a style the destination
        already has (same name) is left as it is there. Bookmarks and drawing ids are not
        renumbered.
        """
        from docx.document import Document

        if isinstance(container, Document):
            container = container._body  # pyright: ignore[reportPrivateUsage]
        target = container._element  # pyright: ignore[reportPrivateUsage]
        if target.tag not in _BLOCK_CONTAINER_TAGS:
            raise TypeError(
                f"cannot append paragraphs and tables to a {target.tag.split('}')[-1]!r} element"
            )
        if replace and target.tag == qn("w:body"):
            raise ValueError("`replace=True` is not supported for the document body")
        clones = self._clone_content(container.part)
        if replace:
            for child in list(target):
                if child.tag not in _KEPT_WHEN_REPLACING:
                    target.remove(child)
        sectPr = target.find(qn("w:sectPr"))
        for clone in clones:
            if sectPr is not None:
                sectPr.addprevious(clone)
            else:
                target.append(clone)
        if target.tag != qn("w:body"):
            _ensure_ends_with_paragraph(target)
        proxies = (_proxy_for(clone, container) for clone in clones)
        return [proxy for proxy in proxies if proxy is not None]

    def insert_before(self, item: Paragraph | Table) -> list[Paragraph | Table]:
        """Copy this building block's content immediately before `item`, return the copies.

        `item` is a |Paragraph| or |Table| in any story (body, cell, header...). See
        `append_to()` for what is and isn't carried over.
        """
        target = item._element  # pyright: ignore[reportPrivateUsage]
        clones = self._clone_content(item.part)
        for clone in clones:
            target.addprevious(clone)
        parent = item._parent  # pyright: ignore[reportPrivateUsage]
        proxies = (_proxy_for(clone, parent) for clone in clones)
        return [proxy for proxy in proxies if proxy is not None]

    @property
    def behaviors(self) -> tuple[WD_BUILDING_BLOCK_BEHAVIOR, ...]:
        """How Word inserts this block: a tuple of `WD_BUILDING_BLOCK_BEHAVIOR` members.

        Assigning an empty sequence removes the behaviors element.
        """
        pr = self._doc_part.docPartPr
        return pr.behavior_vals if pr is not None else ()

    @behaviors.setter
    def behaviors(self, values: Iterable[WD_BUILDING_BLOCK_BEHAVIOR]):
        self._doc_part.get_or_add_docPartPr().behavior_vals = values

    @property
    def category(self) -> str | None:
        """Name of the category this block is listed under, |None| if it has none.

        Assigning a name to a block with no classification also sets its gallery to
        `WD_BUILDING_BLOCK_GALLERY.DEFAULT`, because the file format requires both.
        """
        pr = self._doc_part.docPartPr
        return pr.category_name if pr is not None else None

    @category.setter
    def category(self, value: str):
        self._doc_part.get_or_add_docPartPr().category_name = value

    @property
    def description(self) -> str | None:
        """Description shown by Word for this block, |None| if it has none."""
        pr = self._doc_part.docPartPr
        return pr.description_val if pr is not None else None

    @description.setter
    def description(self, value: str | None):
        self._doc_part.get_or_add_docPartPr().description_val = value

    @property
    def gallery(self) -> WD_BUILDING_BLOCK_GALLERY | None:
        """The gallery this block is classified in, |None| if it has none.

        Assigning a gallery to a block with no classification also sets its category to
        "General", because the file format requires both.
        """
        pr = self._doc_part.docPartPr
        return pr.gallery_val if pr is not None else None

    @gallery.setter
    def gallery(self, value: WD_BUILDING_BLOCK_GALLERY):
        self._doc_part.get_or_add_docPartPr().gallery_val = value

    @property
    def guid(self) -> str | None:
        """The GUID Word uses to identify this block, like `{6F9619FF-...}`, or |None|."""
        pr = self._doc_part.docPartPr
        return pr.guid_val if pr is not None else None

    @property
    def name(self) -> str:
        """The name of this building block, which is what Word shows in its galleries.

        Not checked for uniqueness when assigned; use `BuildingBlocks.get()` first if a
        clash matters.
        """
        pr = self._doc_part.docPartPr
        return pr.name_val if pr is not None else ""

    @name.setter
    def name(self, value: str):
        if not value:
            raise ValueError("a building block must have a non-empty name")
        self._doc_part.get_or_add_docPartPr().name_val = value

    @property
    def text(self) -> str:
        """The text of this block's paragraphs, one per line. Tables are not included."""
        return "\n".join(p.text for p in self.paragraphs)

    @property
    def types(self) -> tuple[WD_BUILDING_BLOCK_TYPE, ...]:
        """The block's types: a tuple of `WD_BUILDING_BLOCK_TYPE` members.

        Assigning an empty sequence removes the types element.
        """
        pr = self._doc_part.docPartPr
        return pr.type_vals if pr is not None else ()

    @types.setter
    def types(self, values: Iterable[WD_BUILDING_BLOCK_TYPE]):
        self._doc_part.get_or_add_docPartPr().type_vals = values

    def _clone_content(self, dest_part: StoryPart) -> list[BaseOxmlElement]:
        """Deep copies of everything this block contains, relationships re-pointed.

        That is every block-level child of the body, so paragraphs and tables but also
        content controls (`w:sdt`), which the built-in header and footer blocks use.
        Relationships the content depends on (images, hyperlinks...) belong to the glossary
        part; the copies are given equivalent relationships on `dest_part`.
        """
        clones = [
            copy.deepcopy(e)
            for e in self._element.content_elements  # pyright: ignore
        ]
        _remap_relationships(clones, self.part, dest_part)
        _import_styles(clones, self.part, dest_part)
        return clones


def _ensure_ends_with_paragraph(story: BaseOxmlElement) -> None:
    """Make sure the block-level content of `story` ends with a paragraph.

    Word requires that of table cells, headers, footers, comments and the like, and
    repairs the file (or refuses to open it) when it isn't so.
    """
    content = [child for child in story if child.tag not in _KEPT_WHEN_REPLACING]
    if not content or content[-1].tag != qn("w:p"):
        story.append(OxmlElement("w:p"))


def _proxy_for(element: BaseOxmlElement, parent) -> Paragraph | Table | None:
    """A |Paragraph| or |Table| proxy for `element`, |None| for any other kind of element."""
    if isinstance(element, CT_P):
        return Paragraph(element, parent)
    if isinstance(element, CT_Tbl):
        return Table(element, parent)
    return None


def _referenced_rIds(root: BaseOxmlElement) -> set[str]:
    """Every relationship id referenced by an `r:`-namespace attribute at or below `root`."""
    return {
        value
        for element in root.iter()
        for name, value in element.attrib.items()
        if name.startswith(_R_NS_PREFIX)
    }


def _remap_relationships(
    elements: Iterable[BaseOxmlElement], src_part: StoryPart, dest_part: StoryPart
) -> None:
    """Re-point relationship references in `elements` from `src_part` to `dest_part`.

    `elements` are copies that were taken from `src_part`, so their `r:id`, `r:embed` and
    other `r:`-namespace attributes hold relationship ids that are only meaningful there.
    For each, an equivalent relationship (same type, same target, internal or external) is
    obtained on `dest_part` and the attribute updated to its id.

    When the two parts belong to different packages (a block taken from a template, say),
    a part cannot simply be shared: its part-name, like `/word/media/image1.png`, may
    already be taken by a different part in the destination package. Images are therefore
    added to the destination package afresh, which names them uniquely and reuses an
    identical image the destination already has. Other kinds of internal target cannot be
    copied between packages yet and raise |ValueError|.
    """
    if src_part is dest_part:
        return
    same_package = src_part.package is dest_part.package
    for root in elements:
        for element in root.iter():
            for attr_name, rId in list(element.attrib.items()):
                if not attr_name.startswith(_R_NS_PREFIX):
                    continue
                rel = src_part.rels.get(rId)
                if rel is None:
                    continue
                if rel.is_external:
                    new_rId = dest_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
                elif same_package:
                    new_rId = dest_part.relate_to(rel.target_part, rel.reltype)
                elif rel.reltype == RT.IMAGE:
                    new_rId, _ = dest_part.get_or_add_image(io.BytesIO(rel.target_part.blob))
                else:
                    kind = rel.reltype.rsplit("/", 1)[-1]
                    raise ValueError(
                        f"cannot copy content with a {kind!r} relationship between documents"
                    )
                element.set(attr_name, new_rId)


_STYLE_REFERENCE_TAGS = ("w:pStyle", "w:rStyle", "w:tblStyle")


def _styles_element_of(part: StoryPart):
    """The `w:styles` element related to `part`, or |None| when it has none.

    A building block's own styles part is related to the glossary part.
    """
    try:
        return part.part_related_by(RT.STYLES).element  # pyright: ignore
    except KeyError:
        return None


def _import_styles(
    clones: Sequence[BaseOxmlElement], src_part: StoryPart, dest_part: StoryPart
) -> None:
    """Make the styles `clones` refer to exist in `dest_part`'s document.

    Word keeps the formatting of a building block when it is inserted: styles the target
    lacks are brought along, together with the styles they are based on, are followed by
    (`w:next`) or linked to. A style the destination already has, matched by name, wins and
    the references in `clones` are pointed at its id. Definitions are looked for in the
    glossary's own styles first, then in the styles of the block's document.
    """
    tags = [qn(tag) for tag in _STYLE_REFERENCE_TAGS]
    references = [ref for clone in clones for ref in clone.iter(*tags)]
    if not references:
        return
    sources = [
        e
        for e in (
            _styles_element_of(src_part),
            _styles_element_of(src_part._document_part),  # pyright: ignore
        )
        if e is not None
    ]
    dest_document_part = dest_part._document_part  # pyright: ignore[reportPrivateUsage]
    dest_styles = _styles_element_of(dest_document_part)
    sources = [e for e in sources if e is not dest_styles]
    if not sources:
        return
    dest = dest_document_part._styles_part.element  # pyright: ignore[reportPrivateUsage]
    mapping: dict[str, str] = {}

    def find_source(style_id: str):
        for styles in sources:
            style = styles.get_by_id(style_id)
            if style is not None:
                return style
        return None

    def bring(style_id: str) -> str:
        """Id in the destination of the style that `style_id` is in the source."""
        if style_id in mapping:
            return mapping[style_id]
        source = find_source(style_id)
        if source is None:
            mapping[style_id] = style_id
            return style_id
        name = (source.name_val or "").lower()
        for existing in dest.xpath("w:style"):
            if (existing.name_val or "").lower() == name and existing.type == source.type:
                mapping[style_id] = existing.styleId or style_id
                return mapping[style_id]
        new_id = style_id
        taken = {s.styleId for s in dest.xpath("w:style")}
        counter = 2
        while new_id in taken:
            new_id = f"{style_id}{counter}"
            counter += 1
        mapping[style_id] = new_id
        style = copy.deepcopy(source)
        style.styleId = new_id
        style.default = None  # -- there is one default style per type; it is the target's --
        dest.append(style)
        for tag in ("w:basedOn", "w:next", "w:link"):
            related = style.find(qn(tag))
            val = related.get(qn("w:val")) if related is not None else None
            if related is not None and val:
                related.set(qn("w:val"), bring(val))
        return new_id

    for ref in references:
        style_id = ref.get(qn("w:val"))
        if style_id:
            ref.set(qn("w:val"), bring(style_id))
