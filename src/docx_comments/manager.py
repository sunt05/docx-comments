"""Main CommentManager class for Word document comment manipulation."""

from __future__ import annotations

import os
import random
import uuid
from datetime import datetime, timezone
from typing import TYPE_CHECKING, Any, Iterator, Optional, Union

from lxml import etree

from docx_comments.anchors import CommentAnchor
from docx_comments.models import CommentInfo, CommentThread, PersonInfo
from docx_comments.system_author import _default_person_from_system
from docx_comments.xml_parts import (
    CommentsExtendedPart,
    CommentsExtensiblePart,
    CommentsIdsPart,
    CommentsPart,
    PeoplePart,
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

PersonSpec = Union[PersonInfo, str, dict[str, Any], bool]


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
        >>> from docx_comments import CommentManager, PersonInfo
        >>>
        >>> doc = Document("document.docx")
        >>> mgr = CommentManager(doc)
        >>>
        >>> # Add comment
        >>> comment_id = mgr.add_comment(
        ...     paragraph=doc.paragraphs[0],
        ...     text="Review this",
        ...     author=PersonInfo(author="Reviewer")
        ... )
        >>>
        >>> # Reply to comment
        >>> reply_id = mgr.reply_to_comment(
        ...     comment_id,
        ...     "Fixed",
        ...     PersonInfo(author="Author")
        ... )
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

    def _comment_index(
        self,
    ) -> tuple[list[CommentInfo], dict[str, CommentInfo], dict[str, CommentInfo]]:
        comments = list(self.list_comments())
        by_id = {c.comment_id: c for c in comments}
        by_para_id = {c.para_id: c for c in comments if c.para_id}
        return comments, by_id, by_para_id

    def _root_for(
        self, comment: CommentInfo, by_para_id: dict[str, CommentInfo]
    ) -> CommentInfo:
        current = comment
        seen: set[str] = set()
        while current.parent_para_id and current.parent_para_id in by_para_id:
            if current.parent_para_id in seen:
                break
            seen.add(current.parent_para_id)
            current = by_para_id[current.parent_para_id]
        return current

    @staticmethod
    def _thread_key(comment: CommentInfo) -> str:
        return comment.para_id or comment.comment_id

    def _thread_comments_for(self, comment_id: str) -> list[CommentInfo]:
        comments, by_id, by_para_id = self._comment_index()
        target = by_id.get(comment_id)
        if target is None:
            raise ValueError(f"Comment {comment_id} not found")
        root = self._root_for(target, by_para_id)
        root_key = self._thread_key(root)
        return [
            comment
            for comment in comments
            if self._thread_key(self._root_for(comment, by_para_id)) == root_key
        ]

    def _collect_comment_para_ids(self) -> set[str]:
        para_ids: set[str] = set()
        for comment_elem in self._comments_xml.findall(_qn(NS_W, "comment")):
            for para in comment_elem.findall(_qn(NS_W, "p")):
                para_id = para.get(_qn(NS_W14, "paraId"))
                if para_id:
                    para_ids.add(para_id)
        return para_ids

    def _cleanup_orphan_metadata(self, valid_para_ids: set[str]) -> None:
        ext_part = CommentsExtendedPart(self._document)
        ids_part = CommentsIdsPart(self._document)
        extensible_part = CommentsExtensiblePart(self._document)

        orphan_para_ids: set[str] = set()
        for elem in list(ext_part.xml):
            if etree.QName(elem).localname != "commentEx":
                continue
            para_id = elem.get(_qn(NS_W15, "paraId"))
            if para_id and para_id not in valid_para_ids:
                orphan_para_ids.add(para_id)

        for elem in list(ids_part.xml):
            if etree.QName(elem).localname != "commentId":
                continue
            para_id = elem.get(_qn(NS_W16CID, "paraId"))
            if para_id and para_id not in valid_para_ids:
                orphan_para_ids.add(para_id)

        removed_durable_ids: set[str] = set()
        for para_id in orphan_para_ids:
            ext_part.remove_comment_ex(para_id)
            durable_id = ids_part.remove_comment_id(para_id)
            if durable_id:
                removed_durable_ids.add(durable_id)

        for durable_id in removed_durable_ids:
            extensible_part.remove_comment_extensible(durable_id)

    def _detach_orphan_replies(self, valid_para_ids: set[str]) -> None:
        ext_part = CommentsExtendedPart(self._document)
        for comment in self.list_comments():
            if not comment.para_id:
                continue
            parent = comment.parent_para_id
            if parent and parent not in valid_para_ids:
                ext_part.set_parent(comment.para_id, None)

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
            ext_entry = extensible_info.get(durable_id) if durable_id else None
            if durable_id and (
                durable_id not in extensible_info
                or not (ext_entry or {}).get("date_utc")
            ):
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

    def get_people(self) -> list[PersonInfo]:
        """
        List people entries from word/people.xml.

        Returns:
            List of PersonInfo entries. Empty if people.xml is absent.
        """
        people_part = PeoplePart(self._document)
        return people_part.get_people()

    def get_person(self, author: str) -> PersonInfo:
        """
        Get a single person entry by author name.

        Args:
            author: Author name to look up in people.xml.

        Returns:
            PersonInfo if found.

        Raises:
            KeyError: If no matching person is found.
        """
        people_part = PeoplePart(self._document)
        return people_part.get_person(author)

    def ensure_person(
        self, author: str, presence: Optional[dict[str, str]] = None
    ) -> PersonInfo:
        """
        Ensure a people.xml entry exists for an author.

        Args:
            author: Author name to match w:comment/@w:author.
            presence: Optional presence metadata with provider_id/user_id.

        Returns:
            PersonInfo for the ensured entry.
        """
        people_part = PeoplePart(self._document)
        return people_part.ensure_person(author, presence)

    def _parse_author_spec(self, author: PersonInfo) -> tuple[str, Optional[dict[str, str]]]:
        if not isinstance(author, PersonInfo):
            raise TypeError("author must be a PersonInfo")

        author_name = author.author
        if not author_name:
            raise ValueError("author must be non-empty")

        presence = None
        if author.provider_id and author.user_id:
            presence = {
                "provider_id": author.provider_id,
                "user_id": author.user_id,
            }
        elif author.provider_id or author.user_id:
            raise ValueError("author presence must include provider_id and user_id")

        return author_name, presence

    def _get_default_author_person(
        self,
        docx_path: Optional[str] = None,
        include_presence: bool = False,
        strict_docx: bool = False,
    ) -> tuple[PersonInfo, Optional[str]]:
        """
        Internal helper to resolve a default author PersonInfo.

        Preference order:
        1) DOCX file from `docx_path` or env var DOCX_COMMENTS_AUTHOR_DOCX
        2) System Office user info (macOS plist / Windows registry)
        3) Current document core properties

        Returns:
            (PersonInfo, initials)

        Raises:
            ValueError: If no author can be resolved.
        """
        person, initials = _default_person_from_system(
            docx_path=docx_path,
            include_presence=include_presence,
            strict_docx=strict_docx,
        )
        if person:
            return person, initials

        if strict_docx and (docx_path or os.environ.get("DOCX_COMMENTS_AUTHOR_DOCX")):
            raise ValueError("default author DOCX did not yield an author")

        author_name = self._document.core_properties.author or ""
        if not author_name:
            author_name = self._document.core_properties.last_modified_by or ""
        if author_name:
            return PersonInfo(author=author_name), None

        raise ValueError("no default author could be resolved")

    def get_default_author_person(
        self,
        docx_path: Optional[str] = None,
        include_presence: bool = False,
        strict_docx: bool = False,
    ) -> tuple[PersonInfo, Optional[str]]:
        """
        Resolve a default author PersonInfo.

        Args:
            docx_path: Optional path to a DOCX file used as the author source.
            include_presence: Whether to include presence metadata from people.xml.
            strict_docx: If True and a DOCX source is provided (or env var set),
                raise when the DOCX cannot provide an author, without falling back.
                A DOCX with multiple people entries triggers a warning and falls back.

        Returns:
            (PersonInfo, initials)
        """
        return self._get_default_author_person(
            docx_path=docx_path,
            include_presence=include_presence,
            strict_docx=strict_docx,
        )

    def merge_people_from(
        self, source: Document, include_presence: bool = False
    ) -> list[PersonInfo]:
        """
        Merge people entries from another document.

        Args:
            source: Document to import people.xml entries from.
            include_presence: Whether to copy presence metadata.

        Returns:
            List of PersonInfo entries added to this document.
        """
        source_part = PeoplePart(source)
        target_part = PeoplePart(self._document)
        return target_part.merge_from(source_part, include_presence)

    def _ensure_person_for_comment(
        self,
        author: str,
        person: Optional[PersonSpec],
    ) -> None:
        if person is None or person is False:
            return

        if isinstance(person, bool):
            if person:
                self.ensure_person(author)
            return

        presence: Optional[dict[str, str]] = None
        person_author = author

        if isinstance(person, PersonInfo):
            person_author = person.author
            if person.provider_id and person.user_id:
                presence = {
                    "provider_id": person.provider_id,
                    "user_id": person.user_id,
                }
        elif isinstance(person, str):
            person_author = person
        elif isinstance(person, dict):
            if "author" in person and isinstance(person["author"], str):
                person_author = person["author"]
            raw_presence = person.get("presence")
            if isinstance(raw_presence, dict):
                presence = raw_presence  # type: ignore[assignment]
            else:
                provider_id = person.get("provider_id") or person.get("providerId")
                user_id = person.get("user_id") or person.get("userId")
                if provider_id and user_id:
                    presence = {
                        "provider_id": str(provider_id),
                        "user_id": str(user_id),
                    }
                elif provider_id or user_id:
                    raise ValueError("presence must include provider_id and user_id")
        else:
            raise TypeError("person must be a bool, str, dict, or PersonInfo")

        if person_author != author:
            raise ValueError("person author must match comment author to link identity")

        self.ensure_person(person_author, presence)

    def add_comment(
        self,
        paragraph: Paragraph,
        text: str,
        author: PersonInfo,
        initials: Optional[str] = None,
        start_run: int = 0,
        end_run: Optional[int] = None,
        person: Optional[PersonSpec] = None,
    ) -> str:
        """
        Add a new anchored comment to a paragraph.

        Args:
            paragraph: The paragraph to comment on.
            text: Comment text.
            author: PersonInfo instance.
            initials: Author initials (optional).
            start_run: Index of first run to anchor (default: 0).
            end_run: Index of last run to anchor (default: all runs).
            person: Optional people.xml entry to link author identity.

        Returns:
            The comment ID of the new comment.
        """
        author_name, author_presence = self._parse_author_spec(author)
        person_spec = person
        if person_spec is None and author_presence:
            person_spec = {"presence": author_presence}
        elif person_spec is True and author_presence:
            person_spec = {"presence": author_presence}

        comment_id = _generate_id()
        para_id = _generate_para_id()
        text_id = _generate_para_id()
        durable_id = _generate_durable_id()

        self._ensure_person_for_comment(author_name, person_spec)

        # 1. Add to comments.xml
        timestamp = self._add_comment_xml(
            comment_id=comment_id,
            para_id=para_id,
            text_id=text_id,
            text=text,
            author=author_name,
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
        author: PersonInfo,
        initials: Optional[str] = None,
        person: Optional[PersonSpec] = None,
    ) -> str:
        """
        Reply to an existing comment.

        Args:
            parent_id: Comment ID of the parent comment.
            text: Reply text.
            author: PersonInfo instance.
            initials: Author initials (optional).
            person: Optional people.xml entry to link author identity.

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

        author_name, author_presence = self._parse_author_spec(author)
        person_spec = person
        if person_spec is None and author_presence:
            person_spec = {"presence": author_presence}
        elif person_spec is True and author_presence:
            person_spec = {"presence": author_presence}

        comment_id = _generate_id()
        para_id = _generate_para_id()
        text_id = _generate_para_id()
        durable_id = _generate_durable_id()

        self._ensure_person_for_comment(author_name, person_spec)

        # 1. Add to comments.xml
        timestamp = self._add_comment_xml(
            comment_id=comment_id,
            para_id=para_id,
            text_id=text_id,
            text=text,
            author=author_name,
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
        self.set_comment_resolved(comment_id, True)

    def unresolve_comment(self, comment_id: str) -> None:
        """
        Mark a comment as unresolved.

        Args:
            comment_id: The comment ID to unresolve.

        Raises:
            ValueError: If comment not found.
        """
        self.set_comment_resolved(comment_id, False)

    def set_comment_resolved(self, comment_id: str, resolved: bool) -> None:
        """
        Set the resolved status for a comment.

        Args:
            comment_id: The comment ID to update.
            resolved: True to resolve, False to unresolve.

        Raises:
            ValueError: If comment not found.
        """
        para_id = None
        for comment in self.list_comments():
            if comment.comment_id == comment_id:
                para_id = comment.para_id
                break

        if not para_id:
            raise ValueError(f"Comment {comment_id} not found")

        ext_part = CommentsExtendedPart(self._document)
        ext_part.set_done(para_id, done=resolved)

    def delete_comment(self, comment_id: str) -> None:
        """
        Delete a single comment.

        Replies remain in the document but are detached from the deleted parent.

        Args:
            comment_id: The comment ID to delete.

        Raises:
            ValueError: If comment not found.
        """
        self.migrate_comment_metadata()

        if self._comments_handler is None:
            self._comments_handler = CommentsPart(self._document)

        removed_para_ids = self._comments_handler.remove_comment(comment_id)
        if removed_para_ids is None:
            raise ValueError(f"Comment {comment_id} not found")

        # Remove anchors for this comment.
        anchor = CommentAnchor(self._document)
        anchor.remove_anchors(comment_id)

        # Remove comment metadata entries.
        deleted_para_ids = {pid for pid in removed_para_ids if pid}
        self._cleanup_comment_metadata(deleted_para_ids)

        remaining_para_ids = self._collect_comment_para_ids()
        self._cleanup_orphan_metadata(remaining_para_ids)
        self._detach_orphan_replies(remaining_para_ids)

    def delete_thread(self, comment_id: str) -> None:
        """
        Delete an entire comment thread (root + replies).

        Args:
            comment_id: Any comment ID within the thread.

        Raises:
            ValueError: If comment not found.
        """
        self.migrate_comment_metadata()
        thread_comments = self._thread_comments_for(comment_id)

        if self._comments_handler is None:
            self._comments_handler = CommentsPart(self._document)

        anchor = CommentAnchor(self._document)
        deleted_para_ids: set[str] = set()

        for comment in thread_comments:
            removed_para_ids = self._comments_handler.remove_comment(comment.comment_id)
            if removed_para_ids is None:
                raise ValueError(f"Comment {comment.comment_id} not found")
            deleted_para_ids.update(pid for pid in removed_para_ids if pid)
            anchor.remove_anchors(comment.comment_id)

        self._cleanup_comment_metadata(deleted_para_ids)
        remaining_para_ids = self._collect_comment_para_ids()
        self._cleanup_orphan_metadata(remaining_para_ids)
        self._detach_orphan_replies(remaining_para_ids)

    def move_comment(
        self,
        comment_id: str,
        paragraph: Paragraph,
        start_run: int = 0,
        end_run: Optional[int] = None,
    ) -> None:
        """
        Move a single comment anchor to a new paragraph.

        Args:
            comment_id: The comment ID to move.
            paragraph: Paragraph to anchor the comment to.
            start_run: Index of first run to anchor.
            end_run: Index of last run to anchor (default: last run).

        Raises:
            ValueError: If comment not found.
        """
        _, by_id, _ = self._comment_index()
        if comment_id not in by_id:
            raise ValueError(f"Comment {comment_id} not found")
        anchor = CommentAnchor(self._document)
        anchor.remove_anchors(comment_id)
        anchor.add_anchors(paragraph, comment_id, start_run=start_run, end_run=end_run)

    def move_thread(
        self,
        comment_id: str,
        paragraph: Paragraph,
        start_run: int = 0,
        end_run: Optional[int] = None,
    ) -> None:
        """
        Move an entire comment thread (root + replies) to a new paragraph.

        Args:
            comment_id: Any comment ID within the thread.
            paragraph: Paragraph to anchor the thread to.
            start_run: Index of first run to anchor (root comment).
            end_run: Index of last run to anchor (root comment).

        Raises:
            ValueError: If comment not found.
        """
        thread_comments = self._thread_comments_for(comment_id)
        by_para_id = {c.para_id: c for c in thread_comments if c.para_id}
        target = next(
            (comment for comment in thread_comments if comment.comment_id == comment_id),
            None,
        )
        if target is None:
            raise ValueError(f"Comment {comment_id} not found")
        root = self._root_for(target, by_para_id)

        anchor = CommentAnchor(self._document)
        for comment in thread_comments:
            anchor.remove_anchors(comment.comment_id)

        anchor.add_anchors(
            paragraph,
            root.comment_id,
            start_run=start_run,
            end_run=end_run,
        )

        # Re-anchor replies at the root comment location.
        for comment in thread_comments:
            if comment.comment_id == root.comment_id:
                continue
            anchor.add_anchors_at_comment(
                parent_comment_id=root.comment_id,
                new_comment_id=comment.comment_id,
            )

    def _cleanup_comment_metadata(self, para_ids: set[str]) -> None:
        if not para_ids:
            return

        ext_part = CommentsExtendedPart(self._document)
        ids_part = CommentsIdsPart(self._document)
        extensible_part = CommentsExtensiblePart(self._document)
        durable_ids = ids_part.get_durable_ids()

        for para_id in para_ids:
            ext_part.remove_comment_ex(para_id)
            ids_part.remove_comment_id(para_id)
            durable_id = durable_ids.get(para_id)
            if durable_id:
                extensible_part.remove_comment_extensible(durable_id)

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
