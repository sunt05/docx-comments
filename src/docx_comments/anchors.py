"""Handler for comment anchors in document.xml."""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator, Optional

from lxml import etree

if TYPE_CHECKING:
    from docx import Document
    from docx.text.paragraph import Paragraph


# OOXML Namespace
NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _qn(ns: str, name: str) -> str:
    """Create qualified name with namespace."""
    return f"{{{ns}}}{name}"


class CommentAnchor:
    """Handler for comment anchors in document.xml."""

    def __init__(self, document: Document) -> None:
        self._document = document

    def _part_element(self, part) -> Optional[etree._Element]:
        """Get an XML element for a part, ensuring it is writable when possible."""
        if part is None:
            return None

        if hasattr(part, "element"):
            try:
                elem = part.element
                if elem is not None:
                    return elem
            except (AttributeError, TypeError, ValueError, etree.XMLSyntaxError):
                pass

        if hasattr(part, "_element"):
            if getattr(part, "_element", None) is None:
                try:
                    part._element = etree.fromstring(part.blob)
                except (AttributeError, TypeError, etree.XMLSyntaxError):
                    return None
            return part._element

        return None

    def _iter_anchor_roots(self) -> Iterator[etree._Element]:
        """Yield XML roots that can contain comment anchors."""
        seen: set[int] = set()

        def add_root(elem: Optional[etree._Element]) -> None:
            if elem is None:
                return
            elem_id = id(elem)
            if elem_id in seen:
                return
            seen.add(elem_id)
            roots.append(elem)

        roots: list[etree._Element] = []
        add_root(self._document.element)

        # Headers/footers across sections, without forcing part creation.
        doc_part = getattr(self._document, "part", None)
        related_parts = getattr(doc_part, "related_parts", {}) if doc_part else {}
        for section in getattr(self._document, "sections", []):
            sect_pr = getattr(section, "_sectPr", None)
            if sect_pr is None:
                continue
            for ref_tag in ("headerReference", "footerReference"):
                for ref in sect_pr.findall(_qn(NS_W, ref_tag)):
                    r_id = ref.get(_qn(NS_R, "id"))
                    if not r_id:
                        continue
                    part = related_parts.get(r_id)
                    add_root(self._part_element(part))

        # Footnotes/endnotes parts (if available).
        if doc_part is not None:
            for attr in ("footnotes_part", "endnotes_part"):
                part = getattr(doc_part, attr, None)
                add_root(self._part_element(part))

        for root in roots:
            yield root

    def _find_anchor_elements(
        self, comment_id: str
    ) -> tuple[Optional[etree._Element], Optional[etree._Element], Optional[etree._Element]]:
        start_xpath = (
            f".//{_qn(NS_W, 'commentRangeStart')}[@{_qn(NS_W, 'id')}='{comment_id}']"
        )
        end_xpath = (
            f".//{_qn(NS_W, 'commentRangeEnd')}[@{_qn(NS_W, 'id')}='{comment_id}']"
        )
        ref_xpath = (
            f".//{_qn(NS_W, 'commentReference')}[@{_qn(NS_W, 'id')}='{comment_id}']"
        )

        for root in self._iter_anchor_roots():
            start = root.find(start_xpath)
            end = root.find(end_xpath)
            if start is None or end is None:
                continue
            ref = root.find(ref_xpath)
            return start, end, ref

        return None, None, None

    def _iter_paragraphs(self) -> Iterator[Paragraph]:
        for para in self._document.paragraphs:
            yield para

        for section in getattr(self._document, "sections", []):
            for attr, ref_tag, ref_type in (
                ("header", "headerReference", None),
                ("footer", "footerReference", None),
                ("first_page_header", "headerReference", "first"),
                ("first_page_footer", "footerReference", "first"),
                ("even_page_header", "headerReference", "even"),
                ("even_page_footer", "footerReference", "even"),
            ):
                if not self._section_has_ref(section, ref_tag, ref_type):
                    continue
                part = getattr(section, attr, None)
                if part is None:
                    continue
                for para in part.paragraphs:
                    yield para

    def _section_has_ref(self, section, ref_tag: str, ref_type: Optional[str]) -> bool:
        sect_pr = getattr(section, "_sectPr", None)
        if sect_pr is None:
            return False
        for ref in sect_pr.findall(_qn(NS_W, ref_tag)):
            ref_type_attr = ref.get(_qn(NS_W, "type"))
            if ref_type is None:
                if ref_type_attr in (None, "default"):
                    return True
            elif ref_type_attr == ref_type:
                return True
        return False

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
        parent_start, parent_end, parent_ref = self._find_anchor_elements(parent_comment_id)

        if parent_start is None or parent_end is None:
            raise ValueError(f"Could not find anchors for comment {parent_comment_id}")

        # Add new anchors after any existing anchor group for this location.
        def is_comment_ref_run(elem: etree._Element) -> bool:
            if etree.QName(elem).localname != "r":
                return False
            return elem.find(_qn(NS_W, "commentReference")) is not None

        # Insert new start after the last commentRangeStart in the group.
        insert_start_after = parent_start
        sibling = parent_start.getnext()
        while sibling is not None and etree.QName(sibling).localname == "commentRangeStart":
            insert_start_after = sibling
            sibling = sibling.getnext()

        new_start = etree.Element(_qn(NS_W, "commentRangeStart"))
        new_start.set(_qn(NS_W, "id"), new_comment_id)
        insert_start_after.addnext(new_start)

        # Insert new end after the last commentRangeEnd in the group.
        insert_end_after = parent_end
        sibling = parent_end.getnext()
        while sibling is not None and etree.QName(sibling).localname == "commentRangeEnd":
            insert_end_after = sibling
            sibling = sibling.getnext()

        new_end = etree.Element(_qn(NS_W, "commentRangeEnd"))
        new_end.set(_qn(NS_W, "id"), new_comment_id)
        insert_end_after.addnext(new_end)

        # Add reference run after existing commentReference runs (if any).
        ref_run = etree.Element(_qn(NS_W, "r"))
        ref = etree.SubElement(ref_run, _qn(NS_W, "commentReference"))
        ref.set(_qn(NS_W, "id"), new_comment_id)
        insert_ref_after = new_end
        sibling = new_end.getnext()
        while sibling is not None and is_comment_ref_run(sibling):
            insert_ref_after = sibling
            sibling = sibling.getnext()
        insert_ref_after.addnext(ref_run)

    def find_paragraph_with_comment(self, comment_id: str) -> Optional[Paragraph]:
        """
        Find the paragraph that contains a comment's anchor.

        Args:
            comment_id: The comment ID to find.

        Returns:
            The Paragraph object, or None if not found.
        """
        # Find commentRangeStart for this comment
        range_start, _, _ = self._find_anchor_elements(comment_id)

        if range_start is None:
            return None

        # Walk up to find parent paragraph
        parent = range_start.getparent()
        while parent is not None:
            if etree.QName(parent).localname == "p":
                # Find matching python-docx Paragraph
                for para in self._iter_paragraphs():
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
        for root in self._iter_anchor_roots():
            for tag in ["commentRangeStart", "commentRangeEnd"]:
                for elem in root.findall(
                    f".//{_qn(NS_W, tag)}[@{_qn(NS_W, 'id')}='{comment_id}']"
                ):
                    elem.getparent().remove(elem)

            # Find and remove commentReference (and its parent run)
            for ref in root.findall(
                f".//{_qn(NS_W, 'commentReference')}[@{_qn(NS_W, 'id')}='{comment_id}']"
            ):
                ref_run = ref.getparent()
                if ref_run is not None and etree.QName(ref_run).localname == "r":
                    # Check if run only contains the reference
                    if len(ref_run) == 1:
                        ref_run.getparent().remove(ref_run)
                    else:
                        ref.getparent().remove(ref)
