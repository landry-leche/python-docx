# pyright: reportPrivateUsage=false

"""Behavior tests for `docx.building_blocks`, exercised against real `Document` objects."""

from __future__ import annotations

import io
import zipfile

import pytest

from docx import Document
from docx.building_blocks import BuildingBlock, BuildingBlocks
from docx.enum.building_block import (
    WD_BUILDING_BLOCK_BEHAVIOR,
    WD_BUILDING_BLOCK_GALLERY,
    WD_BUILDING_BLOCK_TYPE,
)
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn
from docx.oxml.parser import parse_xml
from docx.parts.glossary import GlossaryPart
from docx.shared import Inches
from docx.text.paragraph import Paragraph

from .unitutil.file import test_file


def _round_trip(document) -> "Document":  # pyright: ignore
    stream = io.BytesIO()
    document.save(stream)
    stream.seek(0)
    return Document(stream)


def _child_tags(element) -> list[str]:
    return [child.tag.split("}")[1] for child in element]


class DescribeBuildingBlocks:
    """Collection-level behavior of `Document.building_blocks`."""

    def it_is_empty_and_creates_no_part_for_a_document_without_building_blocks(self):
        document = Document()

        building_blocks = document.building_blocks

        assert isinstance(building_blocks, BuildingBlocks)
        assert len(building_blocks) == 0
        assert list(building_blocks) == []
        assert building_blocks.get("anything") is None
        assert building_blocks.find(category="General") == []
        assert document.part.glossary_part is None
        assert RT.GLOSSARY_DOCUMENT not in {r.reltype for r in document.part.rels.values()}

    def it_creates_the_glossary_part_when_the_first_block_is_added(self):
        document = Document()

        document.building_blocks.add("Sign-off")

        glossary_part = document.part.glossary_part
        assert isinstance(glossary_part, GlossaryPart)
        assert glossary_part.partname == "/word/glossary/document.xml"
        assert glossary_part.content_type == CT.WML_DOCUMENT_GLOSSARY
        assert len(document.building_blocks) == 1

    def it_adds_a_custom_quick_part_by_default(self):
        document = Document()

        block = document.building_blocks.add("Sign-off", "Best regards,\nJane", description="End")

        assert isinstance(block, BuildingBlock)
        assert block.name == "Sign-off"
        assert block.gallery == WD_BUILDING_BLOCK_GALLERY.DOCUMENT_PARTS
        assert block.category == "General"
        assert block.description == "End"
        assert block.types == (WD_BUILDING_BLOCK_TYPE.NORMAL,)
        assert block.behaviors == (WD_BUILDING_BLOCK_BEHAVIOR.CONTENT,)
        assert block.guid is not None and block.guid.startswith("{") and block.guid.endswith("}")
        assert [p.text for p in block.paragraphs] == ["Best regards,", "Jane"]
        assert block.text == "Best regards,\nJane"

    def it_starts_every_block_with_one_paragraph(self):
        block = Document().building_blocks.add("Empty")

        assert len(block.paragraphs) == 1
        assert block.paragraphs[0].text == ""

    def it_writes_docPartPr_children_in_the_order_Word_does(self):
        block = Document().building_blocks.add("Order", description="d")

        pr = block._doc_part.docPartPr
        assert pr is not None
        assert _child_tags(pr) == ["name", "category", "types", "behaviors", "description", "guid"]
        assert _child_tags(pr.category) == ["name", "gallery"]

    def it_rejects_empty_and_duplicate_names(self):
        building_blocks = Document().building_blocks
        building_blocks.add("Dup")

        with pytest.raises(ValueError, match="non-empty"):
            building_blocks.add("")
        with pytest.raises(ValueError, match="already exists"):
            building_blocks.add("Dup")

    def but_allows_the_same_name_in_a_different_gallery_or_category(self):
        building_blocks = Document().building_blocks
        building_blocks.add("Same")

        building_blocks.add("Same", gallery=WD_BUILDING_BLOCK_GALLERY.AUTO_TEXT)
        building_blocks.add("Same", category="Other")

        assert len(building_blocks) == 3

    def it_finds_blocks_by_name_gallery_and_category(self):
        building_blocks = Document().building_blocks
        a = building_blocks.add("A", category="Legal")
        building_blocks.add("B", category="Legal", gallery=WD_BUILDING_BLOCK_GALLERY.AUTO_TEXT)
        building_blocks.add("C")

        assert [b.name for b in building_blocks.find(category="Legal")] == ["A", "B"]
        assert [b.name for b in building_blocks.find(gallery=WD_BUILDING_BLOCK_GALLERY.AUTO_TEXT)] == [
            "B"
        ]
        assert building_blocks.get("A") == a
        assert building_blocks.get("A", category="Nope") is None
        assert [b.name for b in building_blocks] == ["A", "B", "C"]

    def it_round_trips_through_save_and_load(self):
        document = Document()
        document.building_blocks.add(
            "Sign-off",
            "Best regards,\nJane",
            gallery=WD_BUILDING_BLOCK_GALLERY.AUTO_TEXT,
            category="Mail",
            description="End",
            types=(WD_BUILDING_BLOCK_TYPE.AUTO_EXPAND, WD_BUILDING_BLOCK_TYPE.NORMAL),
            behaviors=(WD_BUILDING_BLOCK_BEHAVIOR.PARAGRAPH,),
        )

        reloaded = _round_trip(document)

        assert isinstance(reloaded.part.glossary_part, GlossaryPart)
        (block,) = list(reloaded.building_blocks)
        assert block.name == "Sign-off"
        assert block.gallery == WD_BUILDING_BLOCK_GALLERY.AUTO_TEXT
        assert block.category == "Mail"
        assert block.description == "End"
        assert block.types == (WD_BUILDING_BLOCK_TYPE.AUTO_EXPAND, WD_BUILDING_BLOCK_TYPE.NORMAL)
        assert block.behaviors == (WD_BUILDING_BLOCK_BEHAVIOR.PARAGRAPH,)
        assert block.text == "Best regards,\nJane"

    def it_declares_the_glossary_content_type_and_relationship_in_the_saved_package(self):
        document = Document()
        document.building_blocks.add("X")
        stream = io.BytesIO()
        document.save(stream)

        with zipfile.ZipFile(stream) as zf:
            content_types = zf.read("[Content_Types].xml").decode()
            doc_rels = zf.read("word/_rels/document.xml.rels").decode()
            assert "word/glossary/document.xml" in zf.namelist()

        assert "/word/glossary/document.xml" in content_types
        assert CT.WML_DOCUMENT_GLOSSARY in content_types
        assert RT.GLOSSARY_DOCUMENT in doc_rels
        assert 'Target="glossary/document.xml"' in doc_rels

    def it_removes_a_block(self):
        building_blocks = Document().building_blocks
        keep = building_blocks.add("Keep")
        drop = building_blocks.add("Drop")

        building_blocks.remove(drop)

        assert list(building_blocks) == [keep]
        assert building_blocks.get("Drop") is None

    def it_refuses_to_remove_a_block_from_another_document(self):
        other = Document().building_blocks.add("Foreign")
        building_blocks = Document().building_blocks
        building_blocks.add("Mine")

        with pytest.raises(ValueError, match="not a building block of this document"):
            building_blocks.remove(other)


class DescribeBuildingBlock:
    """Behavior of a single `BuildingBlock`."""

    def it_lets_its_properties_be_changed(self):
        block = Document().building_blocks.add("Old")

        block.name = "New"
        block.gallery = WD_BUILDING_BLOCK_GALLERY.HEADERS
        block.category = "Built-in"
        block.description = None
        block.types = ()
        block.behaviors = (WD_BUILDING_BLOCK_BEHAVIOR.PAGE,)

        assert block.name == "New"
        assert block.gallery == WD_BUILDING_BLOCK_GALLERY.HEADERS
        assert block.category == "Built-in"
        assert block.description is None
        assert block.types == ()
        assert block.behaviors == (WD_BUILDING_BLOCK_BEHAVIOR.PAGE,)
        pr = block._doc_part.docPartPr
        assert _child_tags(pr) == ["name", "category", "behaviors", "guid"]

    def it_keeps_child_order_when_optional_children_are_added_later(self):
        block = Document().building_blocks.add("X", description=None)
        block.types = ()
        block.behaviors = ()
        pr = block._doc_part.docPartPr
        assert _child_tags(pr) == ["name", "category", "guid"]

        block.description = "later"
        block.behaviors = (WD_BUILDING_BLOCK_BEHAVIOR.CONTENT,)
        block.types = (WD_BUILDING_BLOCK_TYPE.NORMAL,)

        assert _child_tags(pr) == ["name", "category", "types", "behaviors", "description", "guid"]

    def it_rejects_an_empty_name(self):
        block = Document().building_blocks.add("X")

        with pytest.raises(ValueError, match="non-empty"):
            block.name = ""

    @pytest.mark.parametrize("which", ["gallery", "category"])
    def it_always_writes_both_halves_of_a_category(self, which: str):
        # -- `w:category` requires both `w:name` and `w:gallery` --
        glossary = Document().building_blocks.add("Seed")
        doc_part = parse_xml(
            '<w:docPart xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:docPartPr><w:name w:val="Bare"/></w:docPartPr>'
            "<w:docPartBody><w:p/></w:docPartBody></w:docPart>"
        )
        bare = BuildingBlock(doc_part, glossary.part)  # pyright: ignore
        assert bare.gallery is None and bare.category is None

        if which == "gallery":
            bare.gallery = WD_BUILDING_BLOCK_GALLERY.EQUATIONS
            assert bare.category == "General"
        else:
            bare.category = "Mine"
            assert bare.gallery == WD_BUILDING_BLOCK_GALLERY.DEFAULT

        category = doc_part.docPartPr.category
        assert _child_tags(category) == ["name", "gallery"]

    def it_exposes_block_item_content_like_a_cell(self):
        block = Document().building_blocks.add("Mixed", "Intro")

        block.add_paragraph("Second")
        table = block.add_table(2, 2)
        table.cell(0, 0).text = "x"

        kinds = [type(item).__name__ for item in block.iter_inner_content()]
        # -- a trailing paragraph follows the table, as a story must not end in one --
        assert kinds == ["Paragraph", "Paragraph", "Table", "Paragraph"]
        assert len(block.tables) == 1
        assert block.tables[0].cell(0, 0).text == "x"
        assert block.text == "Intro\nSecond\n"

    def it_inserts_a_copy_of_its_content_before_a_paragraph(self):
        document = Document()
        document.add_paragraph("first")
        anchor = document.add_paragraph("anchor")
        block = document.building_blocks.add("Sign-off", "Regards,\nJane")

        copies = block.insert_before(anchor)

        assert [p.text for p in document.paragraphs] == ["first", "Regards,", "Jane", "anchor"]
        assert [type(c) for c in copies] == [Paragraph, Paragraph]
        assert [c.text for c in copies] == ["Regards,", "Jane"]
        # -- the block itself is unchanged and independent of the copies --
        copies[0].add_run(" (edited)")
        assert block.text == "Regards,\nJane"

    def it_appends_a_copy_of_its_content_before_the_final_sectPr(self):
        document = Document()
        document.add_paragraph("body")
        block = document.building_blocks.add("Tail", "tail")

        block.append_to(document)

        body = document.element.body
        assert [p.text for p in document.paragraphs][-2:] == ["body", "tail"]
        assert body[-1].tag == qn("w:sectPr")

    def it_can_append_to_a_table_cell(self):
        document = Document()
        table = document.add_table(1, 1)
        block = document.building_blocks.add("Note", "noted")

        block.append_to(table.cell(0, 0))

        assert [p.text for p in table.cell(0, 0).paragraphs] == ["", "noted"]

    def it_inserts_a_block_taken_from_another_document(self):
        source = Document()
        block = source.building_blocks.add("Imported", "from elsewhere")
        target = Document()
        anchor = target.add_paragraph("anchor")

        block.insert_before(anchor)

        assert [p.text for p in target.paragraphs] == ["from elsewhere", "anchor"]


class DescribeBuildingBlocksAndRelationships:
    """Images and hyperlinks must keep working as content moves between parts."""

    def it_carries_an_image_into_a_block_and_back_out(self):
        document = Document()
        picture_para = document.add_paragraph()
        picture_para.add_run().add_picture(test_file("python-icon.png"), width=Inches(1))

        block = document.building_blocks.add_from("Logo", [picture_para])

        glossary_part = document.part.glossary_part
        (embed,) = block._element.xpath(".//a:blip/@r:embed")
        assert glossary_part.rels[embed].reltype == RT.IMAGE
        image_part = glossary_part.rels[embed].target_part
        with open(test_file("python-icon.png"), "rb") as f:
            assert image_part.blob == f.read()

        (copy,) = block.append_to(document)
        (copy_embed,) = copy._element.xpath(".//a:blip/@r:embed")
        assert document.part.rels[copy_embed].target_part is image_part

    def it_survives_save_and_load_with_an_image(self):
        document = Document()
        picture_para = document.add_paragraph()
        picture_para.add_run().add_picture(test_file("python-icon.png"))
        document.building_blocks.add_from("Logo", [picture_para])

        reloaded = _round_trip(document)

        block = reloaded.building_blocks.get("Logo")
        (embed,) = block._element.xpath(".//a:blip/@r:embed")
        rel = reloaded.part.glossary_part.rels[embed]
        assert rel.reltype == RT.IMAGE
        assert len(rel.target_part.blob) > 0

    def it_remaps_an_external_hyperlink(self):
        document = Document()
        paragraph = document.add_paragraph()
        rId = document.part.relate_to("https://example.com/", RT.HYPERLINK, is_external=True)
        paragraph._p.append(
            parse_xml(
                '<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
                ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
                f' r:id="{rId}"><w:r><w:t>link</w:t></w:r></w:hyperlink>'
            )
        )

        block = document.building_blocks.add_from("Link", [paragraph])

        glossary_part = document.part.glossary_part
        (block_rId,) = block._element.xpath(".//w:hyperlink/@r:id")
        rel = glossary_part.rels[block_rId]
        assert rel.is_external and rel.target_ref == "https://example.com/"

    def it_drops_relationships_only_the_removed_block_used(self):
        document = Document()
        p1 = document.add_paragraph()
        p1.add_run().add_picture(test_file("python-icon.png"))
        p2 = document.add_paragraph()
        p2.add_run().add_picture(test_file("monty-truth.png"))
        building_blocks = document.building_blocks
        one = building_blocks.add_from("One", [p1])
        two = building_blocks.add_from("Two", [p2])
        glossary_part = document.part.glossary_part
        image_rels = lambda: [r for r in glossary_part.rels.values() if r.reltype == RT.IMAGE]
        assert len(image_rels()) == 2

        building_blocks.remove(one)

        assert len(image_rels()) == 1
        (embed,) = two._element.xpath(".//a:blip/@r:embed")
        assert embed in glossary_part.rels

    def it_keeps_a_shared_relationship_while_another_block_still_uses_it(self):
        document = Document()
        p = document.add_paragraph()
        p.add_run().add_picture(test_file("python-icon.png"))
        building_blocks = document.building_blocks
        one = building_blocks.add_from("One", [p])
        two = building_blocks.add_from("Two", [p])
        glossary_part = document.part.glossary_part

        building_blocks.remove(one)

        (embed,) = two._element.xpath(".//a:blip/@r:embed")
        assert embed in glossary_part.rels

    def it_ends_a_block_made_from_a_table_with_a_paragraph(self):
        document = Document()
        table = document.add_table(1, 2)

        block = document.building_blocks.add_from("Grid", [table])

        kinds = [type(item).__name__ for item in block.iter_inner_content()]
        assert kinds == ["Table", "Paragraph"]

    def it_rejects_an_empty_selection(self):
        with pytest.raises(ValueError, match="at least one"):
            Document().building_blocks.add_from("Nothing", [])


class DescribeExistingGlossaryParts:
    """Packages that already carry a glossary, like those saved by Word."""

    def it_preserves_the_other_parts_related_to_an_existing_glossary(self):
        # -- build a package whose glossary relates to its own styles part, as Word's do --
        document = Document()
        document.building_blocks.add("Existing", "body text")
        stream = io.BytesIO()
        document.save(stream)

        patched = io.BytesIO()
        with zipfile.ZipFile(stream) as zin, zipfile.ZipFile(patched, "w") as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "[Content_Types].xml":
                    data = data.replace(
                        b"</Types>",
                        b'<Override PartName="/word/glossary/styles.xml" ContentType="'
                        + CT.WML_STYLES.encode()
                        + b'"/></Types>',
                    )
                zout.writestr(item, data)
            zout.writestr(
                "word/glossary/_rels/document.xml.rels",
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                f'<Relationship Id="rId1" Type="{RT.STYLES}" Target="styles.xml"/>'
                "</Relationships>",
            )
            zout.writestr(
                "word/glossary/styles.xml",
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
            )
        patched.seek(0)

        loaded = Document(patched)
        block = loaded.building_blocks.get("Existing")
        assert block is not None and block.text == "body text"
        loaded.building_blocks.add("Another")
        resaved = _round_trip(loaded)

        glossary_part = resaved.part.glossary_part
        assert [b.name for b in resaved.building_blocks] == ["Existing", "Another"]
        assert glossary_part.rels["rId1"].reltype == RT.STYLES
        assert glossary_part.rels["rId1"].target_part.partname == "/word/glossary/styles.xml"
