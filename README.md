# docx-comments

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
- Full Word Online compatibility

## Installation

```bash
pip install docx-comments
```

## Usage

```python
from docx import Document
from docx_comments import CommentManager

doc = Document("document.docx")
mgr = CommentManager(doc)

# Add anchored comment to text range
comment_id = mgr.add_comment(
    paragraph=doc.paragraphs[0],
    start_run=0,
    end_run=2,
    text="Please review this section",
    author="Reviewer Name",
    initials="RN"
)

# Reply to existing comment
reply_id = mgr.reply_to_comment(
    parent_id=comment_id,
    text="Addressed in this revision",
    author="Author Name",
    initials="AN"
)

# Mark comment as resolved
mgr.resolve_comment(comment_id)

# List all comment threads
for thread in mgr.get_comment_threads():
    print(f"Root: {thread.root.text} by {thread.root.author}")
    for reply in thread.replies:
        print(f"  Reply: {reply.text} by {reply.author}")

doc.save("document_reviewed.docx")
```

## OOXML Parts Handled

This module manages four XML parts:

1. **comments.xml** - Comment content and metadata
2. **document.xml** - Anchors (`commentRangeStart/End`, `commentReference`)
3. **commentsExtended.xml** - Threading (`paraId`, `paraIdParent`, `done`)
4. **commentsIds.xml** - Durable IDs for persistence

## Requirements

- Python 3.9+
- python-docx >= 1.0.0
- lxml

## Development

```bash
# Clone and setup
git clone https://github.com/sunt05/docx-comments.git
cd docx-comments
uv venv && uv pip install -e ".[dev]"

# Run tests
pytest

# Type checking
mypy src/docx_comments
```

## License

MIT

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