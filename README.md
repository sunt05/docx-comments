# docx-comments

[![PyPI version](https://badge.fury.io/py/docx-comments.svg)](https://pypi.org/project/docx-comments/)
[![Python versions](https://img.shields.io/pypi/pyversions/docx-comments.svg)](https://pypi.org/project/docx-comments/)
[![CI](https://github.com/sunt05/docx-comments/actions/workflows/ci.yml/badge.svg)](https://github.com/sunt05/docx-comments/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Python module for complete Word document comment manipulation - adding, replying, and resolving comments with full Word Online compatibility.

## Problem

python-docx can read Word comments but cannot properly create or reply to them:
- Creates `comments.xml` but no anchors in `document.xml`
- Missing `commentsExtended.xml` (threading)
- Missing `commentsIds.xml` (durable IDs)

Microsoft Graph API does NOT support Word comments (only Excel).

## Solution

This module provides complete OOXML comment manipulation based on ECMA-376 / ISO/IEC 29500:
- Add anchored comments to specific text ranges
- Reply to existing comments (threaded)
- Mark comments as resolved
- Unresolve comments and toggle done status
- Delete comments or entire threads
- Move comment anchors to new locations
- Full Word Online compatibility
- Optional people.xml identity linkage (Word account presence)

## Installation

```bash
pip install docx-comments
```

## Usage

```python
from docx import Document
from docx_comments import CommentManager, PersonInfo

doc = Document("document.docx")
mgr = CommentManager(doc)

# Author must be a PersonInfo object, not a raw string.

# Add anchored comment to text range
comment_id = mgr.add_comment(
    paragraph=doc.paragraphs[0],
    start_run=0,
    end_run=2,
    text="Please review this section",
    author=PersonInfo(author="Reviewer Name"),
    initials="RN",
    person=True,  # ensure people.xml entry exists for identity linkage
)

# Reply to existing comment
reply_id = mgr.reply_to_comment(
    parent_id=comment_id,
    text="Addressed in this revision",
    author=PersonInfo(author="Author Name"),
    initials="AN"
)

# Mark comment as resolved
mgr.resolve_comment(comment_id)

# Mark comment as unresolved
mgr.unresolve_comment(comment_id)

# Move a comment to a new paragraph
mgr.move_comment(
    comment_id=comment_id,
    paragraph=doc.paragraphs[1],
)

# Delete a comment thread (root + replies)
mgr.delete_thread(comment_id)

# List all comment threads
for thread in mgr.get_comment_threads():
    print(f"Root: {thread.root.text} by {thread.root.author}")
    for reply in thread.replies:
        print(f"  Reply: {reply.text} by {reply.author}")

doc.save("document_reviewed.docx")
```

## Identity Linkage (people.xml)

Word maps `w:comment/@w:author` to account identity using `word/people.xml`. By default, this library does
not create or modify `people.xml` unless you opt in.

```python
# Create a minimal people.xml entry without presence metadata
person = mgr.ensure_person("Reviewer Name")

# Or fetch an existing person entry (raises if missing)
try:
    person = mgr.get_person("Reviewer Name")
except KeyError:
    person = mgr.ensure_person("Reviewer Name")

# Resolve a default author from the system or a DOCX source
person, initials = mgr.get_default_author_person()

# Merge people.xml entries from another document (adds missing authors only)
template_doc = Document("template.docx")
mgr.merge_people_from(template_doc, include_presence=False)

# Or request it when adding a comment
mgr.add_comment(
    paragraph=doc.paragraphs[0],
    text="Linked to people.xml",
    author=person,
    person=True,
)

# You can also pass a PersonInfo object from an existing people.xml
person = mgr.get_people()[0]
mgr.add_comment(
    paragraph=doc.paragraphs[0],
    text="Author from PersonInfo",
    author=person,
)

# Optional presence metadata (only if you explicitly supply it)
mgr.ensure_person(
    "Reviewer Name",
    presence={"provider_id": "provider", "user_id": "user"},
)
```

Note: Word comments are keyed by the author string (`w:comment/@w:author`). If two people share
the same name string, Word does not provide a separate comment author ID to disambiguate them.
Using `people.xml` presence metadata can improve account linkage, but cannot fully resolve
same-name conflicts.

You can also point the resolver at a known DOCX (kept private) using an environment variable:

```bash
export DOCX_COMMENTS_AUTHOR_DOCX="/path/to/author-source.docx"
```

Then call:

```python
person, initials = mgr.get_default_author_person(include_presence=True)
```

If the DOCX contains more than one `w15:person` entry, a warning is raised and the resolver
falls back to system user info.

## OOXML Parts Handled

This module manages five XML parts:

1. **comments.xml** - Comment content and metadata
2. **document.xml** - Anchors (`commentRangeStart/End`, `commentReference`)
3. **commentsExtended.xml** - Threading (`paraId`, `paraIdParent`, `done`)
4. **commentsIds.xml** - Durable IDs for persistence
5. **people.xml** - Optional identity linkage (`w15:person`)

## Requirements

- Python 3.9+
- python-docx >= 1.0.0
- lxml

## References

### OOXML Specification

- [ECMA-376](https://ecma-international.org/publications-and-standards/standards/ecma-376/) - Office Open XML File Formats (free download)
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

### Related Libraries

- [python-docx](https://python-docx.readthedocs.io/) - Python library for Word documents (foundation for this module)
- [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK) - Microsoft's .NET SDK for OOXML

## Acknowledgements

This project was conceptualised by [Ting Sun](https://github.com/sunt05) and implemented with the assistance of [Claude Code](https://claude.ai/code) (Anthropic's AI coding assistant) under his guidance. The collaboration involved iterative development of the OOXML comment handling logic, with Claude Code contributing to code implementation and Ting Sun providing architectural direction and domain expertise.

## License

MIT
