"""Handler for comment anchors in document.xml."""

from __future__ import annotations

from typing import TYPE_CHECKING, Optional

from lxml import etree

if TYPE_CHECKING:
    from docx import Document
    from docx.text.paragraph import Paragraph


# OOXML Namespace
NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _qn(ns: str, name: str) -> str:
    """Create qualified name with namespace."""
    return f"{{{ns}}}{name}"


class CommentAnchor:
    """Handler for comment anchors in document.xml."""

    def __init__(self, document: Document) -> None:
        self._document = document

    @property
    def _body(self) -> etree._Element:
        """Get the document body element."""
        return self._document.element.body

    def add_anchors(
        self,
        paragraph: Paragraph,
        comment_id: str,
        start_run: int = 0,
        end_run: Optional[int] = None,
    ) -> None:
        """
        Add comment anchors to a paragraph.

        Creates commentRangeStart, commentRangeEnd, and commentReference
        elements around the specified runs.

        Args:
            paragraph: The paragraph to anchor the comment to.
            comment_id: The comment ID.
            start_run: Index of first run to anchor (default: 0).
            end_run: Index of last run to anchor (default: all runs).
        """
        para_elem = paragraph._element
        runs = para_elem.findall(_qn(NS_W, "r"))

        if not runs:
            # If no runs, anchor at paragraph level
            self._add_anchors_to_empty_paragraph(para_elem, comment_id)
            return

        # Validate run indices
        if end_run is None:
            end_run = len(runs) - 1
        if start_run < 0 or start_run >= len(runs):
            start_run = 0
        if end_run < start_run or end_run >= len(runs):
            end_run = len(runs) - 1

        # Insert commentRangeStart before start_run
        range_start = etree.Element(_qn(NS_W, "commentRangeStart"))
        range_start.set(_qn(NS_W, "id"), comment_id)
        runs[start_run].addprevious(range_start)

        # Insert commentRangeEnd after end_run
        range_end = etree.Element(_qn(NS_W, "commentRangeEnd"))
        range_end.set(_qn(NS_W, "id"), comment_id)
        runs[end_run].addnext(range_end)

        # Insert commentReference run after commentRangeEnd
        ref_run = etree.Element(_qn(NS_W, "r"))
        ref = etree.SubElement(ref_run, _qn(NS_W, "commentReference"))
        ref.set(_qn(NS_W, "id"), comment_id)
        range_end.addnext(ref_run)

    def _add_anchors_to_empty_paragraph(
        self,
        para_elem: etree._Element,
        comment_id: str,
    ) -> None:
        """Add anchors to a paragraph with no runs."""
        # Create commentRangeStart
        range_start = etree.Element(_qn(NS_W, "commentRangeStart"))
        range_start.set(_qn(NS_W, "id"), comment_id)

        # Create commentRangeEnd
        range_end = etree.Element(_qn(NS_W, "commentRangeEnd"))
        range_end.set(_qn(NS_W, "id"), comment_id)

        # Create commentReference run
        ref_run = etree.Element(_qn(NS_W, "r"))
        ref = etree.SubElement(ref_run, _qn(NS_W, "commentReference"))
        ref.set(_qn(NS_W, "id"), comment_id)

        # Insert after pPr if present, else at start
        pPr = para_elem.find(_qn(NS_W, "pPr"))
        if pPr is not None:
            pPr.addnext(range_start)
        else:
            para_elem.insert(0, range_start)

        range_start.addnext(range_end)
        range_end.addnext(ref_run)

    def add_anchors_at_comment(
        self,
        parent_comment_id: str,
        new_comment_id: str,
    ) -> None:
        """
        Add anchors for a new comment at the same location as an existing comment.

        Used for reply comments that should anchor to the same text.

        Args:
            parent_comment_id: ID of the existing comment.
            new_comment_id: ID of the new comment.
        """
        # Find the parent comment's anchors
        parent_start = self._body.find(
            f".//{_qn(NS_W, 'commentRangeStart')}[@{_qn(NS_W, 'id')}='{parent_comment_id}']"
        )
        parent_end = self._body.find(
            f".//{_qn(NS_W, 'commentRangeEnd')}[@{_qn(NS_W, 'id')}='{parent_comment_id}']"
        )
        parent_ref = self._body.find(
            f".//{_qn(NS_W, 'commentReference')}[@{_qn(NS_W, 'id')}='{parent_comment_id}']"
        )

        if parent_start is None or parent_end is None:
            raise ValueError(f"Could not find anchors for comment {parent_comment_id}")

        # Add new anchors right after parent anchors
        new_start = etree.Element(_qn(NS_W, "commentRangeStart"))
        new_start.set(_qn(NS_W, "id"), new_comment_id)
        parent_start.addnext(new_start)

        new_end = etree.Element(_qn(NS_W, "commentRangeEnd"))
        new_end.set(_qn(NS_W, "id"), new_comment_id)
        parent_end.addnext(new_end)

        # Add reference run
        ref_run = etree.Element(_qn(NS_W, "r"))
        ref = etree.SubElement(ref_run, _qn(NS_W, "commentReference"))
        ref.set(_qn(NS_W, "id"), new_comment_id)

        if parent_ref is not None:
            # Insert after parent's reference run
            parent_ref_run = parent_ref.getparent()
            parent_ref_run.addnext(ref_run)
        else:
            # Insert after commentRangeEnd
            new_end.addnext(ref_run)

    def find_paragraph_with_comment(self, comment_id: str) -> Optional[Paragraph]:
        """
        Find the paragraph that contains a comment's anchor.

        Args:
            comment_id: The comment ID to find.

        Returns:
            The Paragraph object, or None if not found.
        """
        # Find commentRangeStart for this comment
        range_start = self._body.find(
            f".//{_qn(NS_W, 'commentRangeStart')}[@{_qn(NS_W, 'id')}='{comment_id}']"
        )

        if range_start is None:
            return None

        # Walk up to find parent paragraph
        parent = range_start.getparent()
        while parent is not None:
            if etree.QName(parent).localname == "p":
                # Find matching python-docx Paragraph
                for para in self._document.paragraphs:
                    if para._element is parent:
                        return para
                break
            parent = parent.getparent()

        return None

    def remove_anchors(self, comment_id: str) -> None:
        """
        Remove all anchors for a comment.

        Args:
            comment_id: The comment ID whose anchors to remove.
        """
        # Find and remove all anchor elements
        for tag in ["commentRangeStart", "commentRangeEnd"]:
            elem = self._body.find(
                f".//{_qn(NS_W, tag)}[@{_qn(NS_W, 'id')}='{comment_id}']"
            )
            if elem is not None:
                elem.getparent().remove(elem)

        # Find and remove commentReference (and its parent run)
        ref = self._body.find(
            f".//{_qn(NS_W, 'commentReference')}[@{_qn(NS_W, 'id')}='{comment_id}']"
        )
        if ref is not None:
            ref_run = ref.getparent()
            if ref_run is not None and etree.QName(ref_run).localname == "r":
                # Check if run only contains the reference
                if len(ref_run) == 1:
                    ref_run.getparent().remove(ref_run)
                else:
                    ref.getparent().remove(ref)
