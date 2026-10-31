"""Contains the building blocks (glossary document) of the document."""

from __future__ import annotations

import os
from typing import TYPE_CHECKING, cast

from typing_extensions import Self

from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.packuri import PackURI
from docx.oxml.parser import parse_xml
from docx.parts.story import StoryPart

if TYPE_CHECKING:
    from docx.oxml.glossary import CT_GlossaryDocument
    from docx.package import Package


class GlossaryPart(StoryPart):
    """Container part for the building blocks ("glossary document") of a document.

    In a Word package this is `/word/glossary/document.xml`, related to the main document
    part by a relationship of type `glossaryDocument`. Its content is a `w:glossaryDocument`
    element holding one `w:docPart` element per building block.

    A building block's content is a story just like a comment's is, so this part is a
    |StoryPart|; styles and images in building-block content are resolved the same way
    they are for the main document.
    """

    def __init__(
        self,
        partname: PackURI,
        content_type: str,
        element: CT_GlossaryDocument,
        package: Package,
    ):
        super().__init__(partname, content_type, element, package)
        self._glossary_element = element

    @classmethod
    def default(cls, package: Package) -> Self:
        """A newly created glossary part, containing an empty `w:docParts` collection."""
        partname = PackURI("/word/glossary/document.xml")
        content_type = CT.WML_DOCUMENT_GLOSSARY
        element = cast("CT_GlossaryDocument", parse_xml(cls._default_glossary_xml()))
        return cls(partname, content_type, element, package)

    @classmethod
    def _default_glossary_xml(cls) -> bytes:
        """A byte-string containing XML for a default glossary part."""
        path = os.path.join(os.path.split(__file__)[0], "..", "templates", "default-glossary.xml")
        with open(path, "rb") as f:
            xml_bytes = f.read()
        return xml_bytes
