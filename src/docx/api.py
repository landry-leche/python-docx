"""Directly exposed API functions and classes, :func:`Document` for now.

Provides a syntactically more convenient API for interacting with the OpcPackage graph.
"""

from __future__ import annotations

import os
from typing import IO, TYPE_CHECKING, cast

from docx.opc.constants import CONTENT_TYPE as CT
from docx.package import Package

if TYPE_CHECKING:
    from docx.document import Document as DocumentObject
    from docx.parts.document import DocumentPart


# -- content types a main document part may have: document, template, macro-enabled template --
_MAIN_DOCUMENT_CONTENT_TYPES = frozenset(
    (CT.WML_DOCUMENT_MAIN, CT.WML_TEMPLATE_MAIN, CT.WML_TEMPLATE_MAIN_MACRO_ENABLED)
)


def Document(docx: str | IO[bytes] | None = None) -> DocumentObject:
    """Return a |Document| object loaded from `docx`, where `docx` can be either a path
    to a ``.docx`` file (a string) or a file-like object.

    A Word template (``.dotx``, or the macro-enabled ``.dotm``) is also accepted. It is
    loaded like any other document, and saving it writes a template again, not a
    ``.docx``.

    If `docx` is missing or ``None``, the built-in default document "template" is
    loaded.
    """
    docx = _default_docx_path() if docx is None else docx
    document_part = cast("DocumentPart", Package.open(docx).main_document_part)
    if document_part.content_type not in _MAIN_DOCUMENT_CONTENT_TYPES:
        tmpl = "file '%s' is not a Word file, content type is '%s'"
        raise ValueError(tmpl % (docx, document_part.content_type))
    return document_part.document


def _default_docx_path():
    """Return the path to the built-in default .docx package."""
    _thisdir = os.path.split(__file__)[0]
    return os.path.join(_thisdir, "templates", "default.docx")
