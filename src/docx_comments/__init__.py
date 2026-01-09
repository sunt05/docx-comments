"""
docx-comments: Complete Word document comment manipulation.

This module provides full OOXML comment support including:
- Adding anchored comments to specific text ranges
- Replying to existing comments (threaded)
- Marking comments as resolved
- Full Word Online compatibility
"""

from docx_comments.manager import CommentManager
from docx_comments.models import CommentInfo, CommentThread

__version__ = "0.1.0"
__all__ = ["CommentManager", "CommentThread", "CommentInfo"]
