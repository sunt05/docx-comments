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
    CommentsExtensiblePart,
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


def _generate_long_hex_id() -> str:
    """Generate an 8-hex-digit ST_LongHexNumber within the valid range."""
    return f"{random.randint(1, 0x7FFFFFFE):08X}"


def _generate_para_id() -> str:
    """Generate a paragraph ID (8 uppercase hex characters)."""
    return _generate_long_hex_id()


def _generate_durable_id() -> str:
    """Generate a durable ID (8 uppercase hex characters)."""
    return _generate_long_hex_id()


def _format_utc(dt: datetime) -> str:
    """Format a timezone-aware datetime in UTC."""
    return dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _parse_comment_date(date_str: Optional[str]) -> Optional[datetime]:
    """Parse a comment date string into a tz-aware datetime."""
    if not date_str:
        return None
    try:
        if date_str.endswith("Z"):
            return datetime.fromisoformat(date_str.replace("Z", "+00:00"))
        parsed = datetime.fromisoformat(date_str)
        if parsed.tzinfo is None:
            return parsed.replace(tzinfo=timezone.utc)
        return parsed.astimezone(timezone.utc)
    except ValueError:
        return None


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

    def __init__(self, document: Document, auto_migrate: bool = False) -> None:
        """
        Initialize CommentManager with a python-docx Document.

        Args:
            document: A python-docx Document instance.
            auto_migrate: Whether to backfill missing comment metadata on init.
        """
        self._document = document
        self._comments_handler: Optional[CommentsPart] = None
        self._ensure_parts()
        if auto_migrate:
            self.migrate_comment_metadata()

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

    def migrate_comment_metadata(self) -> None:
        """
        Backfill missing comment metadata in existing documents.

        Ensures:
        - w14:paraId and w14:textId on comment paragraphs
        - commentsExtended.xml entries (commentEx)
        - commentsIds.xml entries (durableId)
        - commentsExtensible.xml entries (commentExtensible)
        """
        ensure_comment_parts(self._document)

        ext_part = CommentsExtendedPart(self._document)
        ids_part = CommentsIdsPart(self._document)
        extensible_part = CommentsExtensiblePart(self._document)
        threading = ext_part.get_threading_info()
        durable_ids = ids_part.get_durable_ids()
        extensible_info = extensible_part.get_extensible_info()

        updated_comments = False

        for comment_elem in self._comments_xml.findall(_qn(NS_W, "comment")):
            para_ids = []
            for para in comment_elem.findall(_qn(NS_W, "p")):
                para_id = para.get(_qn(NS_W14, "paraId"))
                if not para_id:
                    para_id = _generate_para_id()
                    para.set(_qn(NS_W14, "paraId"), para_id)
                    updated_comments = True
                para_ids.append(para_id)

                text_id = para.get(_qn(NS_W14, "textId"))
                if not text_id:
                    text_id = _generate_para_id()
                    para.set(_qn(NS_W14, "textId"), text_id)
                    updated_comments = True

            if not para_ids:
                continue

            primary_para_id = None
            for pid in reversed(para_ids):
                if pid in threading:
                    primary_para_id = pid
                    break
            if primary_para_id is None:
                for pid in reversed(para_ids):
                    if pid in durable_ids:
                        primary_para_id = pid
                        break
            if primary_para_id is None:
                primary_para_id = para_ids[-1]

            if primary_para_id not in threading:
                ext_part.add_comment_ex(
                    para_id=primary_para_id, parent_para_id=None, done=False
                )
                threading[primary_para_id] = {
                    "parent_para_id": None,
                    "done": False,
                }

            if primary_para_id not in durable_ids:
                durable_ids[primary_para_id] = _generate_durable_id()
                ids_part.add_comment_id(
                    para_id=primary_para_id,
                    durable_id=durable_ids[primary_para_id],
                )

            durable_id = durable_ids.get(primary_para_id)
            if durable_id and durable_id not in extensible_info:
                date_str = comment_elem.get(_qn(NS_W, "date"))
                timestamp = _parse_comment_date(date_str)
                date_utc = _format_utc(timestamp) if timestamp else None
                extensible_part.add_comment_extensible(
                    durable_id=durable_id,
                    date_utc=date_utc,
                )

        if updated_comments:
            self._save_comments()

    def list_comments(self) -> Iterator[CommentInfo]:
        """
        List all comments in the document.

        Yields:
            CommentInfo objects for each comment.
        """
        # Collect comments from comments.xml
        comments_data: list[dict] = []

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

            # Collect paraIds from all comment paragraphs (some comments span multiple paragraphs)
            para_ids = []
            for para in comment_elem.findall(_qn(NS_W, "p")):
                para_id = para.get(_qn(NS_W14, "paraId"))
                if para_id:
                    para_ids.append(para_id)

            # Parse timestamp (OOXML uses UTC, normalize all to tz-aware)
            timestamp = _parse_comment_date(date_str)

            comments_data.append(
                {
                    "comment_id": comment_id,
                    "para_ids": para_ids,
                    "text": text,
                    "author": author,
                    "initials": initials,
                    "timestamp": timestamp,
                }
            )

        # Get threading info from commentsExtended.xml
        ext_part = CommentsExtendedPart(self._document)
        threading = ext_part.get_threading_info()

        # Get durable IDs from commentsIds.xml
        ids_part = CommentsIdsPart(self._document)
        durable_ids = ids_part.get_durable_ids()

        # Build CommentInfo objects
        for info in comments_data:
            para_ids = info["para_ids"]
            para_id = None
            for pid in reversed(para_ids):
                if pid in threading:
                    para_id = pid
                    break
            if para_id is None:
                for pid in reversed(para_ids):
                    if pid in durable_ids:
                        para_id = pid
                        break
            if para_id is None and para_ids:
                para_id = para_ids[-1]

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

        # Index comments by para_id for parent traversal
        by_para_id = {c.para_id: c for c in comments if c.para_id}

        def thread_key(comment: CommentInfo) -> str:
            return comment.para_id or comment.comment_id

        def root_for(comment: CommentInfo) -> CommentInfo:
            current = comment
            seen: set[str] = set()
            while current.parent_para_id and current.parent_para_id in by_para_id:
                if current.parent_para_id in seen:
                    break
                seen.add(current.parent_para_id)
                current = by_para_id[current.parent_para_id]
            return current

        # Build threads by walking parent chains (supports reply-to-reply)
        threads_by_root: dict[str, CommentThread] = {}
        for comment in comments:
            root = root_for(comment)
            root_key = thread_key(root)
            thread = threads_by_root.get(root_key)
            if thread is None:
                thread = CommentThread(root=root, replies=[])
                threads_by_root[root_key] = thread

            if comment is not root:
                thread.replies.append(comment)

        # Sort replies by timestamp (use tz-aware min for comparison)
        min_dt = datetime.min.replace(tzinfo=timezone.utc)
        for thread in threads_by_root.values():
            thread.replies.sort(key=lambda c: c.timestamp or min_dt)

        return list(threads_by_root.values())

    def get_authors(self) -> dict[str, str]:
        """
        Get all unique authors who have commented on this document.

        Returns:
            Dict mapping author name to initials, e.g. {"Sun, Ting": "ST"}
        """
        authors: dict[str, str] = {}
        for comment in self.list_comments():
            if not comment.author:
                continue
            if comment.author not in authors:
                authors[comment.author] = comment.initials or ""
            elif not authors[comment.author] and comment.initials:
                # Prefer first non-empty initials when available
                authors[comment.author] = comment.initials
        return authors

    def get_document_author(self) -> tuple[str, Optional[str]]:
        """
        Get the document owner's name and initials.

        Uses document core properties for the author name, then looks up
        initials from existing comments by that author.

        Returns:
            Tuple of (author_name, initials). author_name is always a string
            but may be empty ("") if no author is set in document properties.
            Initials may be None if the document owner hasn't made any comments.
        """
        author = self._document.core_properties.author or ""
        if not author:
            # Fallback to last_modified_by
            author = self._document.core_properties.last_modified_by or ""

        # Look for initials in existing comments
        initials = None
        for comment in self.list_comments():
            if comment.author == author and comment.initials:
                initials = comment.initials
                break

        return author, initials

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
        timestamp = self._add_comment_xml(
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

        # 5. Add to commentsExtensible.xml (modern comments metadata)
        extensible_part = CommentsExtensiblePart(self._document)
        extensible_part.add_comment_extensible(
            durable_id=durable_id,
            date_utc=_format_utc(timestamp),
        )

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
        # Find parent comment's para_id and resolve root for compatibility.
        comments = list(self.list_comments())
        parent_comment = next((c for c in comments if c.comment_id == parent_id), None)

        if parent_comment is None or not parent_comment.para_id:
            self.migrate_comment_metadata()
            comments = list(self.list_comments())
            parent_comment = next((c for c in comments if c.comment_id == parent_id), None)
            if parent_comment is None or not parent_comment.para_id:
                raise ValueError(f"Parent comment {parent_id} not found")

        parent_para_id = parent_comment.para_id
        parent_parent_para_id = parent_comment.parent_para_id

        by_para_id = {c.para_id: c for c in comments if c.para_id}
        root_comment = parent_comment
        seen: set[str] = set()
        while (
            root_comment.parent_para_id
            and root_comment.parent_para_id in by_para_id
            and root_comment.parent_para_id not in seen
        ):
            seen.add(root_comment.parent_para_id)
            root_comment = by_para_id[root_comment.parent_para_id]

        # Word UI doesn't support nested replies; attach to the root comment.
        effective_parent_para_id = root_comment.para_id or parent_para_id
        effective_parent_parent_para_id = root_comment.parent_para_id

        anchor = CommentAnchor(self._document)

        comment_id = _generate_id()
        para_id = _generate_para_id()
        text_id = _generate_para_id()
        durable_id = _generate_durable_id()

        # 1. Add to comments.xml
        timestamp = self._add_comment_xml(
            comment_id=comment_id,
            para_id=para_id,
            text_id=text_id,
            text=text,
            author=author,
            initials=initials,
        )

        # 2. Add anchors at the root comment location for Word threading compatibility.
        anchor_parent_id = root_comment.comment_id or parent_id
        anchor.add_anchors_at_comment(
            parent_comment_id=anchor_parent_id,
            new_comment_id=comment_id,
        )

        # 3. Ensure parent exists in commentsExtended.xml, then add reply link
        ext_part = CommentsExtendedPart(self._document)
        threading = ext_part.get_threading_info()
        if parent_para_id not in threading:
            ext_part.add_comment_ex(
                para_id=parent_para_id,
                parent_para_id=parent_parent_para_id,
                done=False,
            )
        if effective_parent_para_id not in threading and effective_parent_para_id != parent_para_id:
            ext_part.add_comment_ex(
                para_id=effective_parent_para_id,
                parent_para_id=effective_parent_parent_para_id,
                done=False,
            )
        ext_part.add_comment_ex(
            para_id=para_id,
            parent_para_id=effective_parent_para_id,
            done=False,
        )

        # 4. Add to commentsIds.xml
        ids_part = CommentsIdsPart(self._document)
        ids_part.add_comment_id(para_id=para_id, durable_id=durable_id)

        # 5. Add to commentsExtensible.xml (modern comments metadata)
        extensible_part = CommentsExtensiblePart(self._document)
        extensible_part.add_comment_extensible(
            durable_id=durable_id,
            date_utc=_format_utc(timestamp),
        )

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
        timestamp: Optional[datetime] = None,
    ) -> datetime:
        """Add a comment element to comments.xml and return its timestamp."""
        rsid_r = uuid.uuid4().hex[:8].upper()
        rsid_default = uuid.uuid4().hex[:8].upper()
        rsid_rpr = uuid.uuid4().hex[:8].upper()

        # Build comment element
        comment = etree.SubElement(self._comments_xml, _qn(NS_W, "comment"))
        comment.set(_qn(NS_W, "id"), comment_id)
        comment.set(_qn(NS_W, "author"), author)
        if initials:
            comment.set(_qn(NS_W, "initials"), initials)
        # Use local time with offset so Word displays the expected timestamp.
        if timestamp is None:
            timestamp = datetime.now().astimezone()
        comment.set(
            _qn(NS_W, "date"),
            timestamp.isoformat(timespec="seconds"),
        )

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
        return timestamp
