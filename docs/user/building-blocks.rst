.. _building_blocks:

Working with Building Blocks
============================

Word lets you save a piece of content once and reuse it anywhere, as a *building block*.
The Word UI calls them by different names depending on where you meet them: *Quick Parts*,
*AutoText*, *cover pages*, *headers* and *footers*, *page numbers*, *equations*,
*watermarks*, and so on. They are all the same thing underneath, a named piece of content
filed in a *gallery* and a *category*.

*python-docx* can list, create, edit, insert and remove building blocks, so you can keep a
library of boilerplate in a document and assemble new documents from it.

.. note::

   Building blocks are accessed through a :class:`.Document` object. Word usually keeps its
   built-in galleries in *template* files (``.dotx`` and ``.dotm``), which
   :func:`docx.Document` does not open; see :ref:`building_blocks_limitations`. A
   ``.docx`` file can hold building blocks too, and that is what this page works with.


Building block anatomy
----------------------

A building block has these properties, all of which are available as attributes of a
:class:`.BuildingBlock` object:

``name``
    What Word displays for the block in a gallery. It is unique within a gallery and
    category.

``gallery``
    One of the galleries Word offers, a member of
    :class:`~docx.enum.building_block.WD_BUILDING_BLOCK_GALLERY`. Blocks you create
    yourself usually go in ``DOCUMENT_PARTS`` (the *Quick Parts* gallery) or ``AUTO_TEXT``.

``category``
    A free-form name Word uses to group blocks within a gallery. ``"General"`` is Word's
    default.

``description``
    Optional text Word shows as a tooltip.

``types`` and ``behaviors``
    Technical flags. ``types`` says what kind of entry the block is (a normal block, an
    AutoCorrect entry, structured-document-tag placeholder text, ...) and ``behaviors`` says
    whether Word inserts it as inline content, as its own paragraph, or on its own page.
    The defaults are the ones Word uses for a custom Quick Part.

``guid``
    The identifier Word uses to recognize the block. Assigned automatically.

The content of the block is a sequence of paragraphs and tables, handled exactly like the
content of a table cell or a header: a block has ``paragraphs``, ``tables``,
``add_paragraph()``, ``add_table()`` and ``iter_inner_content()``.

In the file, building blocks are stored in a separate *glossary document* part
(``word/glossary/document.xml``) related to the main document. Each one looks like this:

.. code-block:: xml

    <w:docPart>
      <w:docPartPr>
        <w:name w:val="Sign-off"/>
        <w:category>
          <w:name w:val="General"/>
          <w:gallery w:val="docParts"/>
        </w:category>
        <w:types><w:type w:val="normal"/></w:types>
        <w:behaviors><w:behavior w:val="content"/></w:behaviors>
        <w:guid w:val="{F846BD45-7528-439E-A91B-00BC3A8C332B}"/>
      </w:docPartPr>
      <w:docPartBody>
        <w:p><w:r><w:t>Best regards,</w:t></w:r></w:p>
        <w:p><w:r><w:t>Jane Doe</w:t></w:r></w:p>
      </w:docPartBody>
    </w:docPart>

You never need to touch this XML; it is shown only so you can recognize it if you open a
``.docx`` file in an archive tool.


Accessing the building blocks of a document
-------------------------------------------

The collection is available from the :attr:`.Document.building_blocks` property::

    >>> from docx import Document
    >>> document = Document()
    >>> building_blocks = document.building_blocks
    >>> len(building_blocks)
    0

The collection can be iterated and has a length. Reading it never changes the document: the
glossary part that stores the blocks is only created when the first block is added, so a
document you merely inspect is saved exactly as it was.


Adding a building block
-----------------------

:meth:`.BuildingBlocks.add` creates a block. With no further arguments it creates what Word
calls a custom Quick Part: gallery *Quick Parts*, category *General*::

    >>> signoff = building_blocks.add("Sign-off", "Best regards,\nJane Doe")
    >>> signoff.name
    'Sign-off'
    >>> print(signoff.gallery)
    DOCUMENT_PARTS (3)
    >>> signoff.category
    'General'
    >>> signoff.text
    'Best regards,\nJane Doe'

The optional second argument is the block's text. A newline starts a new paragraph. A block
always has at least one paragraph, because Word does not allow an empty one, so omitting the
text gives you a block with a single empty paragraph.

Use the keyword arguments to choose the gallery and category and to add a description::

    >>> from docx.enum.building_block import WD_BUILDING_BLOCK_GALLERY
    >>> disclaimer = building_blocks.add(
    ...     "Disclaimer",
    ...     "This message is confidential.",
    ...     gallery=WD_BUILDING_BLOCK_GALLERY.AUTO_TEXT,
    ...     category="Legal",
    ...     description="Standard confidentiality notice",
    ... )
    >>> [block.name for block in building_blocks]
    ['Sign-off', 'Disclaimer']

A name must be non-empty and unique within its gallery and category. :exc:`ValueError` is
raised otherwise::

    >>> building_blocks.add("Sign-off")
    Traceback (most recent call last):
      ...
    ValueError: a building block named 'Sign-off' already exists in gallery DOCUMENT_PARTS and category 'General'

The same name can be used again in a different gallery or category.


Finding building blocks
-----------------------

:meth:`.BuildingBlocks.get` returns the first block with a given name, or |None|::

    >>> building_blocks.get("Disclaimer").category
    'Legal'
    >>> building_blocks.get("No such block") is None
    True

Because names are only unique within a gallery and category, ``get()`` accepts ``gallery``
and ``category`` keyword arguments to disambiguate. :meth:`.BuildingBlocks.find` returns a
list of every block that matches whichever of ``name``, ``gallery`` and ``category`` you
give it::

    >>> [block.name for block in building_blocks.find(category="Legal")]
    ['Disclaimer']
    >>> autotext = WD_BUILDING_BLOCK_GALLERY.AUTO_TEXT
    >>> [block.name for block in building_blocks.find(gallery=autotext)]
    ['Disclaimer']

Names and categories are compared exactly, including case.


Adding rich content to a building block
---------------------------------------

The content of a building block is a *block-item container*, like the document body or a
table cell. It can hold any number of paragraphs and tables, and the paragraphs can have
formatted runs and pictures. The methods are the ones you already know from the document::

    >>> terms = building_blocks.add("Terms")
    >>> run = terms.paragraphs[0].add_run("Payment terms")
    >>> run.bold = True
    >>> table = terms.add_table(rows=1, cols=2)
    >>> table.cell(0, 0).text = "Net 30"
    >>> [type(item).__name__ for item in terms.iter_inner_content()]
    ['Paragraph', 'Table', 'Paragraph']

Notice the paragraph after the table. A story in Word cannot end with a table, so
``add_table()`` adds an empty paragraph after it, as it does for a table cell. If you do not
give a width, the table is 6 inches wide, spread evenly over its columns.


Saving existing content as a building block
-------------------------------------------

The equivalent of selecting some content in Word and choosing *Save Selection to Quick Part
Gallery* is :meth:`.BuildingBlocks.add_from`. Pass it the paragraphs and tables you want, in
order::

    >>> from docx.shared import Inches
    >>> paragraph = document.add_paragraph()
    >>> picture = paragraph.add_run().add_picture("logo.png", width=Inches(1))
    >>> logo = building_blocks.add_from("Logo", [paragraph], description="Company logo")
    >>> len(logo.paragraphs)
    1

The block holds a *copy*; the paragraph in the document is left alone. Pictures and
hyperlinks in the content are carried over together with the relationships they need, and
the picture is stored only once in the file. ``add_from()`` accepts the same ``gallery``,
``category`` and ``description`` arguments as ``add()``. If the last item you pass is a
table, an empty paragraph is added after it, for the reason given above.

Any ``Paragraph`` or ``Table`` can be used, including ones from another document.


Inserting a building block into a document
------------------------------------------

There are two ways to use a block, depending on where you want the content.
:meth:`.BuildingBlock.insert_before` puts a copy of the block immediately before a paragraph
or table::

    >>> closing = document.add_paragraph("Thank you for your business.")
    >>> copies = signoff.insert_before(closing)
    >>> [paragraph.text for paragraph in document.paragraphs][-3:]
    ['Best regards,', 'Jane Doe', 'Thank you for your business.']

:meth:`.BuildingBlock.append_to` puts a copy at the end of a container. The container can be
the document itself, a table cell, a header or a footer::

    >>> copies = disclaimer.append_to(document)
    >>> document.paragraphs[-1].text
    'This message is confidential.'

    >>> table = document.add_table(rows=1, cols=1)
    >>> copies = terms.append_to(table.cell(0, 0))

Both methods return the new paragraphs and tables, so you can adjust them. When appending to
the document, the content is placed at the end of the last section, ahead of the final
section properties.

A block can be inserted into a different document from the one it lives in, which is how you
share a library of blocks between documents::

    >>> other = Document()
    >>> copies = signoff.append_to(other)
    >>> [paragraph.text for paragraph in other.paragraphs]
    ['Best regards,', 'Jane Doe']

Insertion is *block-level*: every paragraph and table of the block is copied exactly as it is
stored. In particular, the block's ``behaviors`` (which tell Word to insert inline, as a
paragraph, or on a new page) are not applied. If you want a page break, add one yourself.


Changing a building block
-------------------------

Name, gallery, category and description can all be assigned. ``types`` and ``behaviors`` are
sequences of enumeration members::

    >>> from docx.enum.building_block import WD_BUILDING_BLOCK_BEHAVIOR
    >>> signoff.name = "Closing"
    >>> signoff.description = "Closing for letters"
    >>> signoff.behaviors = (WD_BUILDING_BLOCK_BEHAVIOR.PARAGRAPH,)
    >>> [behavior.name for behavior in signoff.behaviors]
    ['PARAGRAPH']
    >>> [type_.name for type_ in signoff.types]
    ['NORMAL']
    >>> signoff.guid
    '{...}'

Assigning ``None`` to ``description`` removes it, and assigning an empty sequence to ``types``
or ``behaviors`` removes that property. A name cannot be set to an empty string. The name is
*not* checked for uniqueness when you assign it; call ``get()`` first if a clash matters.

A category in the file always has both a name and a gallery. A block that has neither, which
can happen in files made by other software, gets Word's default category ``"General"`` when
you assign only a gallery, and the ``DEFAULT`` gallery when you assign only a category.


Removing a building block
-------------------------

Pass the block to :meth:`.BuildingBlocks.remove`::

    >>> building_blocks.remove(building_blocks.get("Logo"))
    >>> [block.name for block in building_blocks]
    ['Closing', 'Disclaimer', 'Terms']

Pictures and hyperlinks that no other block uses are removed from the glossary along with
the block. The ``block`` object must not be used afterward, and :exc:`ValueError` is raised
if it does not belong to this document.


Saving and reopening
--------------------

Building blocks are saved with the document and are there when it is opened again::

    >>> document.save("report.docx")
    >>> reopened = Document("report.docx")
    >>> [block.name for block in reopened.building_blocks]
    ['Closing', 'Disclaimer', 'Terms']

A document that already has a glossary part keeps everything in it that *python-docx* does
not manage, such as the styles and settings Word stores alongside the blocks.


.. _building_blocks_limitations:

Limitations
-----------

- **Templates.** ``docx.Document()`` opens only documents whose main part is a regular
  ``.docx`` document, so it raises :exc:`ValueError` for ``.dotx`` and ``.dotm`` templates.
  That is where Word keeps its built-in galleries and the ``Normal`` template's custom
  Quick Parts, so those cannot be reached directly yet.
- **Styles.** A glossary created by *python-docx* has no styles of its own. Paragraph and
  character styles named in a block are looked up in the main document's styles.
- **Insertion is a copy, not an expansion.** The block's ``behaviors`` are not interpreted
  (see above). Bookmark ids and drawing ids are not renumbered when content is copied, so
  inserting the same block several times repeats them.
- **Not modelled.** The style reference (``w:style``) of a building block and the
  ``decorate`` attribute of its name are preserved when a file is saved but are not
  exposed as properties.
