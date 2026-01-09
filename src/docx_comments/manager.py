"""Main CommentManager class for Word document comment manipulation."""

from __future__ import annotations

import random
import uuid
from datetime import datetime, timezone
from typing import TYPE_CHECKING, Iterator, Optional

from lxml import etree

from docx_comments.anchors import CommentAnchor
from docx_comments.models import CommentInfo, CommentThread
from docx_comments.xml_parts import (
    CommentsExtendedPart,
    CommentsIdsPart,
    CommentsPart,
    ensure_comment_parts,
)

if TYPE_CHECKING:
    from docx import Document
    from docx.text.paragraph import Paragraph


# OOXML Namespaces
NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
NS_W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
NS_W16CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid"


def _qn(ns: str, name: str) -> str:
    """Create qualified name with namespace."""
    return f"{{{ns}}}{name}"


def _generate_id() -> str:
    """Generate a random comment ID (large positive integer as string)."""
    return str(random.randint(1_000_000_000, 9_999_999_999))


def _generate_para_id() -> str:
    """Generate a paragraph ID (8 uppercase hex characters)."""
    return uuid.uuid4().hex[:8].upper()


def _generate_durable_id() -> str:
    """Generate a durable ID (8 uppercase hex characters)."""
    return uuid.uuid4().hex[:8].upper()


class CommentManager:
    """
    Manager for Word document comments.

    Provides complete comment manipulation including:
    - Adding anchored comments to specific text ranges
    - Replying to existing comments (threaded)
    - Marking comments as resolved
    - Full Word Online compatibility

    Example:
        >>> from docx import Document
        >>> from docx_comments import CommentManager
        >>>
        >>> doc = Document("document.docx")
        >>> mgr = CommentManager(doc)
        >>>
        >>> # Add comment
        >>> comment_id = mgr.add_comment(
        ...     paragraph=doc.paragraphs[0],
        ...     text="Review this",
        ...     author="Reviewer"
        ... )
        >>>
        >>> # Reply to comment
        >>> reply_id = mgr.reply_to_comment(comment_id, "Fixed", "Author")
        >>>
        >>> doc.save("reviewed.docx")
    """

    def __init__(self, document: Document) -> None:
        """
        Initialize CommentManager with a python-docx Document.

        Args:
            document: A python-docx Document instance.
        """
        self._document = document
        self._comments_handler: Optional[CommentsPart] = None
        self._ensure_parts()

    def _ensure_parts(self) -> None:
        """Ensure all required comment parts exist in the document."""
        ensure_comment_parts(self._document)
        # Cache the comments part handler
        self._comments_handler = CommentsPart(self._document)

    @property
    def _comments_xml(self) -> etree._Element:
        """Get the comments.xml root element."""
        if self._comments_handler is None:
            self._comments_handler = CommentsPart(self._document)
        return self._comments_handler.xml

    def _save_comments(self) -> None:
        """Save changes to comments.xml."""
        if self._comments_handler is not None:
            self._comments_handler._save()

    def list_comments(self) -> Iterator[CommentInfo]:
        """
        List all comments in the document.

        Yields:
            CommentInfo objects for each comment.
        """
        # Build para_id to comment mapping from comments.xml
        para_id_map: dict[str, dict] = {}

        for comment_elem in self._comments_xml.findall(_qn(NS_W, "comment")):
            comment_id = comment_elem.get(_qn(NS_W, "id"))
            author = comment_elem.get(_qn(NS_W, "author"), "")
            initials = comment_elem.get(_qn(NS_W, "initials"))
            date_str = comment_elem.get(_qn(NS_W, "date"))

            # Get text content
            text_parts = []
            for t_elem in comment_elem.findall(f".//{_qn(NS_W, 't')}"):
                if t_elem.text:
                    text_parts.append(t_elem.text)
            text = "".join(text_parts)

            # Get para_id from first paragraph
            para = comment_elem.find(_qn(NS_W, "p"))
            para_id = None
            if para is not None:
                para_id = para.get(_qn(NS_W14, "paraId"))

            # Parse timestamp (OOXML uses UTC, normalize all to tz-aware)
            timestamp = None
            if date_str:
                try:
                    if date_str.endswith("Z"):
                        timestamp = datetime.fromisoformat(date_str.replace("Z", "+00:00"))
                    else:
                        # Assume UTC if no timezone specified
                        timestamp = datetime.fromisoformat(date_str).replace(tzinfo=timezone.utc)
                except ValueError:
                    pass

            para_id_map[para_id] = {
                "comment_id": comment_id,
                "para_id": para_id,
                "text": text,
                "author": author,
                "initials": initials,
                "timestamp": timestamp,
            }

        # Get threading info from commentsExtended.xml
        ext_part = CommentsExtendedPart(self._document)
        threading = ext_part.get_threading_info()

        # Get durable IDs from commentsIds.xml
        ids_part = CommentsIdsPart(self._document)
        durable_ids = ids_part.get_durable_ids()

        # Build CommentInfo objects
        for para_id, info in para_id_map.items():
            thread_info = threading.get(para_id, {})
            yield CommentInfo(
                comment_id=info["comment_id"],
                para_id=para_id or "",
                text=info["text"],
                author=info["author"],
                initials=info["initials"],
                timestamp=info["timestamp"],
                parent_para_id=thread_info.get("parent_para_id"),
                is_resolved=thread_info.get("done", False),
                durable_id=durable_ids.get(para_id),
            )

    def get_comment_threads(self) -> list[CommentThread]:
        """
        Get all comment threads (grouped by root comment).

        Returns:
            List of CommentThread objects.
        """
        comments = list(self.list_comments())

        # Separate roots and replies
        roots: dict[str, CommentInfo] = {}
        replies_by_parent: dict[str, list[CommentInfo]] = {}

        for comment in comments:
            if comment.is_reply and comment.parent_para_id:
                if comment.parent_para_id not in replies_by_parent:
                    replies_by_parent[comment.parent_para_id] = []
                replies_by_parent[comment.parent_para_id].append(comment)
            else:
                roots[comment.para_id] = comment

        # Build threads
        threads = []
        for para_id, root in roots.items():
            replies = replies_by_parent.get(para_id, [])
            # Sort replies by timestamp (use tz-aware min for comparison)
            min_dt = datetime.min.replace(tzinfo=timezone.utc)
            replies.sort(key=lambda c: c.timestamp or min_dt)
            threads.append(CommentThread(root=root, replies=replies))

        return threads

    def add_comment(
        self,
        paragraph: Paragraph,
        text: str,
        author: str,
        initials: Optional[str] = None,
        start_run: int = 0,
        end_run: Optional[int] = None,
    ) -> str:
        """
        Add a new anchored comment to a paragraph.

        Args:
            paragraph: The paragraph to comment on.
            text: Comment text.
            author: Author name.
            initials: Author initials (optional).
            start_run: Index of first run to anchor (default: 0).
            end_run: Index of last run to anchor (default: all runs).

        Returns:
            The comment ID of the new comment.
        """
        comment_id = _generate_id()
        para_id = _generate_para_id()
        text_id = _generate_para_id()
        durable_id = _generate_durable_id()

        # 1. Add to comments.xml
        self._add_comment_xml(
            comment_id=comment_id,
            para_id=para_id,
            text_id=text_id,
            text=text,
            author=author,
            initials=initials,
        )

        # 2. Add anchors to document.xml
        anchor = CommentAnchor(self._document)
        anchor.add_anchors(
            paragraph=paragraph,
            comment_id=comment_id,
            start_run=start_run,
            end_run=end_run,
        )

        # 3. Add to commentsExtended.xml (root comment, no parent)
        ext_part = CommentsExtendedPart(self._document)
        ext_part.add_comment_ex(para_id=para_id, parent_para_id=None, done=False)

        # 4. Add to commentsIds.xml
        ids_part = CommentsIdsPart(self._document)
        ids_part.add_comment_id(para_id=para_id, durable_id=durable_id)

        return comment_id

    def reply_to_comment(
        self,
        parent_id: str,
        text: str,
        author: str,
        initials: Optional[str] = None,
    ) -> str:
        """
        Reply to an existing comment.

        Args:
            parent_id: Comment ID of the parent comment.
            text: Reply text.
            author: Author name.
            initials: Author initials (optional).

        Returns:
            The comment ID of the reply.

        Raises:
            ValueError: If parent comment not found.
        """
        # Find parent comment's para_id
        parent_para_id = None
        parent_paragraph = None

        for comment in self.list_comments():
            if comment.comment_id == parent_id:
                parent_para_id = comment.para_id
                break

        if not parent_para_id:
            raise ValueError(f"Parent comment {parent_id} not found")

        # Find the paragraph that the parent comment is anchored to
        anchor = CommentAnchor(self._document)
        parent_paragraph = anchor.find_paragraph_with_comment(parent_id)

        if not parent_paragraph:
            raise ValueError(f"Could not find anchor for parent comment {parent_id}")

        comment_id = _generate_id()
        para_id = _generate_para_id()
        text_id = _generate_para_id()
        durable_id = _generate_durable_id()

        # 1. Add to comments.xml
        self._add_comment_xml(
            comment_id=comment_id,
            para_id=para_id,
            text_id=text_id,
            text=text,
            author=author,
            initials=initials,
        )

        # 2. Add anchors to same location as parent
        anchor.add_anchors_at_comment(
            parent_comment_id=parent_id,
            new_comment_id=comment_id,
        )

        # 3. Add to commentsExtended.xml with parent link
        ext_part = CommentsExtendedPart(self._document)
        ext_part.add_comment_ex(para_id=para_id, parent_para_id=parent_para_id, done=False)

        # 4. Add to commentsIds.xml
        ids_part = CommentsIdsPart(self._document)
        ids_part.add_comment_id(para_id=para_id, durable_id=durable_id)

        return comment_id

    def resolve_comment(self, comment_id: str) -> None:
        """
        Mark a comment as resolved.

        Args:
            comment_id: The comment ID to resolve.

        Raises:
            ValueError: If comment not found.
        """
        # Find comment's para_id
        para_id = None
        for comment in self.list_comments():
            if comment.comment_id == comment_id:
                para_id = comment.para_id
                break

        if not para_id:
            raise ValueError(f"Comment {comment_id} not found")

        ext_part = CommentsExtendedPart(self._document)
        ext_part.set_done(para_id, done=True)

    def _add_comment_xml(
        self,
        comment_id: str,
        para_id: str,
        text_id: str,
        text: str,
        author: str,
        initials: Optional[str],
    ) -> None:
        """Add a comment element to comments.xml."""
        rsid_r = uuid.uuid4().hex[:8].upper()
        rsid_default = uuid.uuid4().hex[:8].upper()
        rsid_rpr = uuid.uuid4().hex[:8].upper()

        # Build comment element
        comment = etree.SubElement(self._comments_xml, _qn(NS_W, "comment"))
        comment.set(_qn(NS_W, "id"), comment_id)
        comment.set(_qn(NS_W, "author"), author)
        if initials:
            comment.set(_qn(NS_W, "initials"), initials)
        comment.set(_qn(NS_W, "date"), datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"))

        # Add paragraph
        para = etree.SubElement(comment, _qn(NS_W, "p"))
        para.set(_qn(NS_W, "rsidR"), rsid_r)
        para.set(_qn(NS_W, "rsidRDefault"), rsid_default)
        para.set(_qn(NS_W14, "paraId"), para_id)
        para.set(_qn(NS_W14, "textId"), text_id)

        # Add paragraph properties with CommentText style
        pPr = etree.SubElement(para, _qn(NS_W, "pPr"))
        pStyle = etree.SubElement(pPr, _qn(NS_W, "pStyle"))
        pStyle.set(_qn(NS_W, "val"), "CommentText")

        # Add run with annotationRef
        run1 = etree.SubElement(para, _qn(NS_W, "r"))
        rPr = etree.SubElement(run1, _qn(NS_W, "rPr"))
        rStyle = etree.SubElement(rPr, _qn(NS_W, "rStyle"))
        rStyle.set(_qn(NS_W, "val"), "CommentReference")
        etree.SubElement(run1, _qn(NS_W, "annotationRef"))

        # Add run with text
        run2 = etree.SubElement(para, _qn(NS_W, "r"))
        run2.set(_qn(NS_W, "rsidRPr"), rsid_rpr)
        t = etree.SubElement(run2, _qn(NS_W, "t"))
        t.text = text

        # Save changes to the part
        self._save_comments()
