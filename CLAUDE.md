# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

`docx-comments` is a Python module for complete Word document comment manipulation. It fills gaps in python-docx by providing:
- Anchored comments with proper OOXML structure
- Threaded replies
- Comment resolution (done status)
- Full Word Online compatibility

## Development Commands

```bash
# Setup
uv venv && uv pip install -e ".[dev]"

# Run tests
pytest                          # Full suite
pytest tests/test_basic.py -v   # Single file with verbose
pytest -k "test_add_comment"    # Single test by name

# Type checking
mypy src/docx_comments

# Linting
ruff check src/
ruff format src/
```

## Architecture

### OOXML Comment System

Word comments require coordination across four XML parts:

1. **comments.xml** - Comment content (text, author, timestamp)
2. **document.xml** - Anchors linking comments to text ranges
3. **commentsExtended.xml** - Threading (parent-child relationships, done status)
4. **commentsIds.xml** - Durable IDs for persistence across edits

### Module Structure

- `manager.py` - `CommentManager` class: main public API
  - `add_comment()`, `reply_to_comment()`, `resolve_comment()`
  - `list_comments()`, `get_comment_threads()`
  - `get_authors()`, `get_document_author()` - Author introspection

- `xml_parts.py` - Handlers for XML parts
  - `CommentsPart` - Main comments.xml (handles XmlPart vs generic Part serialisation)
  - `CommentsExtendedPart` - Threading info (w15:paraIdParent, w15:done)
  - `CommentsIdsPart` - Durable IDs (w16cid:durableId)
  - `ensure_comment_parts()` - Creates missing parts with proper relationships

- `anchors.py` - `CommentAnchor` class: manages document.xml anchors
  - Inserts `commentRangeStart`, `commentRangeEnd`, `commentReference`
  - Handles empty paragraphs and reply co-location

- `models.py` - Data classes: `CommentInfo`, `CommentThread`

### Key Implementation Details

**ID Generation** (`manager.py:37-49`):
- `comment_id`: Large random integer (10 digits)
- `para_id`: 8 uppercase hex chars (links comments.xml to extended parts)
- `durable_id`: 8 uppercase hex chars (persistence across edits)

**Namespace Prefixes**:
- `w:` - Main WordprocessingML (2006)
- `w14:` - Word 2010 extensions (paraId on paragraphs)
- `w15:` - Word 2012 extensions (threading, done status)
- `w16cid:` - Word 2016 extensions (durable IDs)

**Part Creation** (`xml_parts.py`): Uses python-docx internals (`docx.opc.part.Part`, `docx.opc.packuri.PackURI`) to create new parts with correct content types and relationships.

**XmlPart Serialisation**: python-docx loads comments.xml as an `XmlPart` subclass which uses `_element` for serialisation, not `_blob`. The `CommentsPart` handler detects this and works with `_element` directly for existing documents, falling back to blob-based caching for newly created parts.

## Testing Notes

Tests use `tmp_path` fixture for save/reload verification. The test file `tests/test_basic.py` covers:
- Manager initialisation with new documents
- Comment add/reply/resolve operations
- Thread grouping logic
- Model property correctness
- Word Online compatibility (XML structure validation via zipfile inspection)

## References

### OOXML Specification

- [ECMA-376 Standard](https://ecma-international.org/publications-and-standards/standards/ecma-376/) - Office Open XML File Formats (free download)
- [ISO/IEC 29500](https://www.loc.gov/preservation/digital/formats/fdd/fdd000395.shtml) - OOXML Format Family overview
- [MS-DOCX Extensions](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-docx/) - Microsoft's DOCX extensions documentation

### Comment Elements

- [commentRangeStart](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_commentRangeStart_topic_ID0EFJMV.html) - Comment anchor range start element
- [commentRangeEnd](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_commentRangeEnd_topic_ID0ESCLV.html) - Comment anchor range end element
- [commentReference](https://c-rex.net/samples/ooxml/e1/part4/OOXML_P4_DOCX_commentReference_topic_ID0E4PNV.html) - Comment content reference mark
- [Comments Overview](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_Comments_topic_ID0EEHJV.html) - OOXML comments specification

### Threading & Extended Parts

- [CommentEx Class](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2013.word.commentex) - Office 2013 comment threading (paraId, paraIdParent, done)
- [commentsIds](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-docx/22977b5a-5bb5-4f27-b7a1-c6d216c2bb94) - Durable IDs specification (Office 2016+)
- [Open XML SDK Issue #484](https://github.com/OfficeDev/Open-XML-SDK/issues/484) - commentsIds part discussion

### Implementation Guides

- [MS Learn: Insert Comment](https://learn.microsoft.com/en-us/office/open-xml/word/how-to-insert-a-comment-into-a-word-processing-document) - C# implementation guide
- [WordprocessingML Structure](https://learn.microsoft.com/en-us/office/open-xml/word/structure-of-a-wordprocessingml-document) - Document structure overview
- [Office Open XML Anatomy](http://officeopenxml.com/anatomyofOOXML.php) - Package structure explained

### Related Libraries

- [python-docx](https://python-docx.readthedocs.io/) - Python library for Word documents
- [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK) - Microsoft's .NET SDK for OOXML
