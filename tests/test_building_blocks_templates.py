# pyright: reportPrivateUsage=false

"""Tests for the "building block from a template into headers and footers" use case.

The scenario: Word keeps reusable blocks in a template (`.dotx` or `.dotm`); an existing
`.docx` with several sections should get one of those blocks into its headers or footers.
"""

from __future__ import annotations

import io
import zipfile

import pytest

from docx import Document
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn
from docx.oxml.parser import parse_xml
from docx.shared import Inches

from .unitutil.file import test_file

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _retype(data: bytes, old: str, new: str) -> bytes:
    """Copy of the package `data` whose main part is declared with content type `new`."""
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as zin, zipfile.ZipFile(out, "w") as zout:
        for item in zin.infolist():
            blob = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                assert old.encode() in blob
                blob = blob.replace(old.encode(), new.encode())
            zout.writestr(item, blob)
    return out.getvalue()


def _template_bytes(template_type: str, *blocks: tuple[str, str]) -> bytes:
    """A template package of `template_type` holding the (name, text) building `blocks`."""
    document = Document()
    for name, text in blocks:
        document.building_blocks.add(name, text)
    stream = io.BytesIO()
    document.save(stream)
    return _retype(stream.getvalue(), CT.WML_DOCUMENT_MAIN, template_type)


def _three_section_document():
    document = Document()
    document.add_paragraph("one")
    document.add_section()
    document.add_paragraph("two")
    document.add_section()
    document.add_paragraph("three")
    return document


def _texts(story) -> list[str]:
    return [p.text for p in story.paragraphs]


class DescribeTemplateFiles:
    """`Document()` opens Word templates, which is where Word keeps its building blocks."""

    @pytest.mark.parametrize(
        "template_type", [CT.WML_TEMPLATE_MAIN, CT.WML_TEMPLATE_MAIN_MACRO_ENABLED]
    )
    def it_opens_a_template_and_saves_it_as_a_template(self, template_type: str):
        template = Document(io.BytesIO(_template_bytes(template_type, ("Footer text", "Secret"))))

        assert template.part.content_type == template_type
        assert [b.name for b in template.building_blocks] == ["Footer text"]
        template.building_blocks.add("Another")
        stream = io.BytesIO()
        template.save(stream)

        with zipfile.ZipFile(stream) as zf:
            content_types = zf.read("[Content_Types].xml").decode()
        assert template_type in content_types
        assert CT.WML_DOCUMENT_MAIN not in content_types
        reopened = Document(io.BytesIO(stream.getvalue()))
        assert [b.name for b in reopened.building_blocks] == ["Footer text", "Another"]

    def it_still_rejects_a_package_that_is_not_a_word_file(self):
        stream = io.BytesIO()
        Document().save(stream)
        not_word = _retype(stream.getvalue(), CT.WML_DOCUMENT_MAIN, CT.SML_SHEET_MAIN)

        with pytest.raises(ValueError, match="is not a Word file"):
            Document(io.BytesIO(not_word))


class DescribeInsertingIntoHeadersAndFooters:
    """Inserting a block into the headers and footers of the sections of a document."""

    def it_inserts_into_each_section_that_has_its_own_header_and_footer(self):
        document = _three_section_document()
        one, two, three = document.sections
        one.header.is_linked_to_previous = False
        three.header.is_linked_to_previous = False
        one.footer.is_linked_to_previous = False
        block = document.building_blocks.add("Banner", "ACME Corp")

        for section in document.sections:
            for header_footer in (section.header, section.footer):
                if not header_footer.is_linked_to_previous:
                    block.append_to(header_footer, replace=True)

        assert _texts(one.header) == ["ACME Corp"]
        assert _texts(three.header) == ["ACME Corp"]
        assert _texts(one.footer) == ["ACME Corp"]
        # -- a linked header shows the one it inherits; it was neither touched nor doubled --
        assert two.header.is_linked_to_previous
        assert _texts(two.header) == ["ACME Corp"]
        assert three.footer.is_linked_to_previous

    def it_reaches_the_inherited_header_when_a_header_is_linked(self):
        document = _three_section_document()
        document.sections[0].header.is_linked_to_previous = False
        block = document.building_blocks.add("Banner", "ACME")

        for section in document.sections:  # -- naive loop: no `is_linked_to_previous` check --
            block.append_to(section.header)

        assert _texts(document.sections[0].header).count("ACME") == 3

    def it_can_replace_the_content_of_a_header(self):
        document = Document()
        header = document.sections[0].header
        header.is_linked_to_previous = False
        header.paragraphs[0].text = "old"
        block = document.building_blocks.add("New", "new text")

        copies = block.append_to(header, replace=True)

        assert _texts(header) == ["new text"]
        assert [c.text for c in copies] == ["new text"]

    def it_adds_after_the_existing_content_by_default(self):
        document = Document()
        header = document.sections[0].header
        header.is_linked_to_previous = False
        header.paragraphs[0].text = "kept"

        document.building_blocks.add("New", "added").append_to(header)

        assert _texts(header) == ["kept", "added"]

    def it_refuses_to_replace_the_content_of_the_document_body(self):
        document = Document()
        document.add_paragraph("precious")
        block = document.building_blocks.add("Block", "text")

        with pytest.raises(ValueError, match="document body"):
            block.append_to(document, replace=True)

        assert _texts(document) == ["precious"]

    def it_keeps_the_properties_of_a_cell_when_replacing_its_content(self):
        document = Document()
        cell = document.add_table(1, 1).cell(0, 0)
        cell.width = Inches(2)
        block = document.building_blocks.add("Block", "text")

        block.append_to(cell, replace=True)

        assert _texts(cell) == ["text"]
        assert cell._tc.tcPr is not None and cell.width == Inches(2)

    def it_copies_content_controls_as_well_as_paragraphs_and_tables(self):
        document = Document()
        block = document.building_blocks.add("Controlled", "tail")
        block._element.insert(
            0,
            parse_xml(
                f'<w:sdt xmlns:w="{W_NS}"><w:sdtPr/><w:sdtContent>'
                "<w:p><w:r><w:t>In control</w:t></w:r></w:p></w:sdtContent></w:sdt>"
            ),
        )
        header = document.sections[0].header
        header.is_linked_to_previous = False

        copies = block.append_to(header, replace=True)

        tags = [child.tag.split("}")[1] for child in header._element]
        assert tags == ["sdt", "p"]
        assert header._element.xpath(".//w:sdt//w:t/text()") == ["In control"]
        # -- only paragraphs and tables get proxies, the content control is just carried --
        assert [c.text for c in copies] == ["tail"]
        assert len(list(block.iter_inner_content())) == 1

    def it_never_leaves_a_header_ending_with_a_table(self):
        document = Document()
        block = document.building_blocks.add("Grid")
        block.add_table(1, 1)
        block._element.remove(block._element.p_lst[-1])
        assert block._element[-1].tag == qn("w:tbl")
        header = document.sections[0].header
        header.is_linked_to_previous = False

        block.append_to(header, replace=True)

        assert header._element[-1].tag == qn("w:p")
        assert len(header.tables) == 1

    def it_rejects_an_element_that_is_not_a_story(self):
        document = Document()
        block = document.building_blocks.add("Block", "text")
        paragraph = document.add_paragraph("p")

        class FakeContainer:
            _element = paragraph._p
            part = document.part

        with pytest.raises(TypeError, match="cannot append"):
            block.append_to(FakeContainer())  # pyright: ignore


class DescribeImagesAcrossDocuments:
    """A block's pictures must follow it into a footer of another package without clashes."""

    def _picture_block(self, image: str, name: str = "Logo"):
        source = Document()
        paragraph = source.add_paragraph()
        paragraph.add_run().add_picture(test_file(image))
        return source, source.building_blocks.add_from(name, [paragraph])

    def it_maps_an_image_onto_the_footer_part(self):
        _, block = self._picture_block("python-icon.png")
        target = Document()
        footer = target.sections[0].footer
        footer.is_linked_to_previous = False

        block.append_to(footer, replace=True)

        (embed,) = footer._element.xpath(".//a:blip/@r:embed")
        rel = footer.part.rels[embed]
        assert rel.reltype == RT.IMAGE
        with open(test_file("python-icon.png"), "rb") as f:
            assert rel.target_part.blob == f.read()

    def it_keeps_two_different_images_apart_when_their_part_names_would_clash(self):
        # -- both documents name their only image `/word/media/image1.png` --
        _, block = self._picture_block("python-icon.png")
        target = Document()
        target.add_paragraph().add_run().add_picture(test_file("monty-truth.png"))
        footer = target.sections[0].footer
        footer.is_linked_to_previous = False

        block.append_to(footer, replace=True)
        stream = io.BytesIO()
        target.save(stream)

        with zipfile.ZipFile(stream) as zf:
            names = zf.namelist()
            media = sorted(n for n in names if n.startswith("word/media/"))
            assert len(names) == len(set(names))
            assert len(media) == 2
            blobs = {zf.read(n) for n in media}
        with open(test_file("python-icon.png"), "rb") as a, open(
            test_file("monty-truth.png"), "rb"
        ) as b:
            assert blobs == {a.read(), b.read()}

    def it_reuses_an_identical_image_the_target_already_has(self):
        _, block = self._picture_block("python-icon.png")
        target = Document()
        target.add_paragraph().add_run().add_picture(test_file("python-icon.png"))
        footer = target.sections[0].footer
        footer.is_linked_to_previous = False

        block.append_to(footer, replace=True)
        stream = io.BytesIO()
        target.save(stream)

        with zipfile.ZipFile(stream) as zf:
            assert len([n for n in zf.namelist() if n.startswith("word/media/")]) == 1

    def it_refuses_other_internal_relationships_between_documents(self):
        source = Document()
        block = source.building_blocks.add("Odd", "text")
        part = source.part
        odd_part = part.package.get_or_add_image_part(test_file("python-icon.png"))  # a stand-in
        rId = block.part.relate_to(odd_part, RT.OLE_OBJECT)
        block._element.p_lst[0].append(
            parse_xml(f'<w:r xmlns:w="{W_NS}" xmlns:r="{RT_NS}"><w:object r:id="{rId}"/></w:r>')
        )
        target = Document()
        header = target.sections[0].header
        header.is_linked_to_previous = False

        with pytest.raises(ValueError, match="between documents"):
            block.append_to(header)


RT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


class DescribeTheWholeScenario:
    """A block stored in a template ends up in the footers of an existing document."""

    def it_puts_a_template_block_into_the_footers_of_an_existing_document(self):
        template = Document(
            io.BytesIO(
                _template_bytes(
                    CT.WML_TEMPLATE_MAIN, ("Company footer", "ACME Corp - Confidential")
                )
            )
        )
        block = template.building_blocks.get("Company footer")
        assert block is not None

        existing = _three_section_document()
        existing.sections[0].footer.is_linked_to_previous = False
        existing.sections[2].footer.is_linked_to_previous = False
        for section in existing.sections:
            if not section.footer.is_linked_to_previous:
                block.append_to(section.footer, replace=True)
        stream = io.BytesIO()
        existing.save(stream)

        saved = Document(io.BytesIO(stream.getvalue()))
        footers = [s.footer for s in saved.sections]
        assert [f.is_linked_to_previous for f in footers] == [False, True, False]
        assert _texts(footers[0]) == ["ACME Corp - Confidential"]
        assert _texts(footers[1]) == ["ACME Corp - Confidential"]  # inherited from section 1
        assert _texts(footers[2]) == ["ACME Corp - Confidential"]
        assert saved.part.content_type == CT.WML_DOCUMENT_MAIN
        assert saved.part.glossary_part is None  # -- the template's glossary stays behind --
