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
uv venv
source .venv/bin/activate
uv pip install -e ".[dev]"

# Run tests
pytest

# Linting
ruff check src/docx_comments
```

## License

MIT

## References

- [ECMA-376 Standard](https://ecma-international.org/publications-and-standards/standards/ecma-376/)
- [MS Learn: Insert Comment](https://learn.microsoft.com/en-us/office/open-xml/word/how-to-insert-a-comment-into-a-word-processing-document)
- [python-docx](https://python-docx.readthedocs.io/)

## Tracking

Linear issue: [PER-406](https://linear.app/sunt05/issue/PER-406)
