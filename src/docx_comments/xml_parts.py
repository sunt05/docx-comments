"""Handlers for XML parts: comments.xml, commentsExtended.xml, and commentsIds.xml."""

from __future__ import annotations

from typing import TYPE_CHECKING, Optional

from lxml import etree

if TYPE_CHECKING:
    from docx import Document


# OOXML Namespaces
NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
NS_W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
NS_W16CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"

# Relationship types
REL_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
REL_COMMENTS_EXT = (
    "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"
)
REL_COMMENTS_IDS = (
    "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds"
)

# Content types
CT_COMMENTS = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
CT_COMMENTS_EXT = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"
)
CT_COMMENTS_IDS = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"
)


def _qn(ns: str, name: str) -> str:
    """Create qualified name with namespace."""
    return f"{{{ns}}}{name}"


class CommentsPart:
    """Handler for word/comments.xml part.

    Note: python-docx loads comments.xml as an XmlPart subclass (CommentsPart)
    which uses _element for serialization, not _blob. We must work with
    part._element directly to persist changes.
    """

    def __init__(self, document: Document) -> None:
        self._document = document

    def _get_part(self):
        """Get the comments part from document relationships."""
        for rel in self._document.part.rels.values():
            if REL_COMMENTS in rel.reltype:
                return rel.target_part
        return None

    def ensure_exists(self) -> None:
        """Ensure the comments part exists, creating if needed."""
        if self._get_part() is None:
            self._create_part()

    def _create_part(self) -> None:
        """Create a new comments.xml part."""
        from docx.opc.packuri import PackURI
        from docx.opc.part import Part

        # Create XML content with required namespaces
        nsmap = {
            "w": NS_W,
            "w14": NS_W14,
            "w15": NS_W15,
            "mc": NS_MC,
        }
        root = etree.Element(
            _qn(NS_W, "comments"),
            nsmap=nsmap,
        )
        root.set(_qn(NS_MC, "Ignorable"), "w14 w15")

        xml_content = etree.tostring(
            root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone="yes",
        )

        # Create generic part (python-docx will load it as XmlPart on next open)
        part = Part(
            PackURI("/word/comments.xml"),
            CT_COMMENTS,
            xml_content,
            self._document.part.package,
        )
        self._document.part.relate_to(part, REL_COMMENTS)

    @property
    def xml(self) -> etree._Element:
        """Get the XML root element.

        Handles two cases:
        - XmlPart (from existing document): use part._element directly
        - Generic Part (newly created): parse and cache from blob
        """
        part = self._get_part()
        if part is None:
            # Shouldn't happen after ensure_exists
            return etree.Element(_qn(NS_W, "comments"))

        # Check if it's an XmlPart (has _element attribute)
        if hasattr(part, "_element"):
            return part._element

        # Generic Part - need to parse blob and maintain cache
        if not hasattr(self, "_xml_cache") or self._xml_cache is None:
            self._xml_cache = etree.fromstring(part.blob)
        return self._xml_cache

    def _save(self) -> None:
        """Save changes back to the part.

        - XmlPart: changes to _element persist automatically
        - Generic Part: need to update _blob
        """
        part = self._get_part()
        if part is None:
            return

        # XmlPart doesn't need explicit save
        if hasattr(part, "_element"):
            return

        # Generic Part - update _blob from cached xml
        if hasattr(self, "_xml_cache") and self._xml_cache is not None:
            part._blob = etree.tostring(
                self._xml_cache,
                xml_declaration=True,
                encoding="UTF-8",
                standalone="yes",
            )


def ensure_comment_parts(document: Document) -> None:
    """
    Ensure all required comment parts exist in the document.

    Creates:
    - comments.xml if missing
    - commentsExtended.xml if missing
    - commentsIds.xml if missing
    """
    # Ensure comments.xml (main comments part)
    comments_part = CommentsPart(document)
    comments_part.ensure_exists()

    # Ensure commentsExtended.xml
    ext_part = CommentsExtendedPart(document)
    ext_part.ensure_exists()

    # Ensure commentsIds.xml
    ids_part = CommentsIdsPart(document)
    ids_part.ensure_exists()


class CommentsExtendedPart:
    """Handler for word/commentsExtended.xml part."""

    def __init__(self, document: Document) -> None:
        self._document = document
        self._xml: Optional[etree._Element] = None

    def _get_part(self):
        """Get the commentsExtended part from document relationships."""
        for rel in self._document.part.rels.values():
            if REL_COMMENTS_EXT in rel.reltype:
                return rel.target_part
        return None

    def ensure_exists(self) -> None:
        """Ensure the commentsExtended part exists, creating if needed."""
        if self._get_part() is None:
            self._create_part()

    def _create_part(self) -> None:
        """Create a new commentsExtended.xml part."""
        # Create XML content
        nsmap = {
            "mc": NS_MC,
            "w15": NS_W15,
        }
        root = etree.Element(
            _qn(NS_W15, "commentsEx"),
            nsmap=nsmap,
        )
        root.set(_qn(NS_MC, "Ignorable"), "w15")

        xml_content = etree.tostring(
            root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone="yes",
        )

        # Add part to document
        # Note: This requires accessing python-docx internals
        from docx.opc.packuri import PackURI
        from docx.opc.part import Part

        part = Part(
            PackURI("/word/commentsExtended.xml"),
            CT_COMMENTS_EXT,
            xml_content,
            self._document.part.package,
        )
        self._document.part.relate_to(part, REL_COMMENTS_EXT)

    @property
    def xml(self) -> etree._Element:
        """Get the XML root element."""
        if self._xml is None:
            part = self._get_part()
            if part:
                self._xml = etree.fromstring(part.blob)
            else:
                # Return empty element if part doesn't exist
                self._xml = etree.Element(_qn(NS_W15, "commentsEx"))
        return self._xml

    def _save(self) -> None:
        """Save changes back to the part."""
        part = self._get_part()
        if part:
            part._blob = etree.tostring(
                self.xml,
                xml_declaration=True,
                encoding="UTF-8",
                standalone="yes",
            )

    def get_threading_info(self) -> dict[str, dict]:
        """
        Get threading information for all comments.

        Returns:
            Dict mapping para_id to {"parent_para_id": str|None, "done": bool}
        """
        result = {}
        for elem in self.xml:
            if etree.QName(elem).localname == "commentEx":
                para_id = elem.get(_qn(NS_W15, "paraId"))
                parent = elem.get(_qn(NS_W15, "paraIdParent"))
                done = elem.get(_qn(NS_W15, "done"), "0") == "1"
                if para_id:
                    result[para_id] = {
                        "parent_para_id": parent,
                        "done": done,
                    }
        return result

    def add_comment_ex(
        self,
        para_id: str,
        parent_para_id: Optional[str] = None,
        done: bool = False,
    ) -> None:
        """
        Add a commentEx entry.

        Args:
            para_id: Paragraph ID of the comment.
            parent_para_id: Paragraph ID of parent (for replies).
            done: Whether comment is resolved.
        """
        elem = etree.SubElement(self.xml, _qn(NS_W15, "commentEx"))
        elem.set(_qn(NS_W15, "paraId"), para_id)
        elem.set(_qn(NS_W15, "done"), "1" if done else "0")
        if parent_para_id:
            elem.set(_qn(NS_W15, "paraIdParent"), parent_para_id)
        self._save()

    def set_done(self, para_id: str, done: bool) -> None:
        """
        Set the done status for a comment.

        Args:
            para_id: Paragraph ID of the comment.
            done: Whether comment is resolved.
        """
        for elem in self.xml:
            if etree.QName(elem).localname == "commentEx":
                if elem.get(_qn(NS_W15, "paraId")) == para_id:
                    elem.set(_qn(NS_W15, "done"), "1" if done else "0")
                    self._save()
                    return
        raise ValueError(f"Comment with para_id {para_id} not found in commentsExtended")


class CommentsIdsPart:
    """Handler for word/commentsIds.xml part."""

    def __init__(self, document: Document) -> None:
        self._document = document
        self._xml: Optional[etree._Element] = None

    def _get_part(self):
        """Get the commentsIds part from document relationships."""
        for rel in self._document.part.rels.values():
            if REL_COMMENTS_IDS in rel.reltype:
                return rel.target_part
        return None

    def ensure_exists(self) -> None:
        """Ensure the commentsIds part exists, creating if needed."""
        if self._get_part() is None:
            self._create_part()

    def _create_part(self) -> None:
        """Create a new commentsIds.xml part."""
        # Create XML content
        nsmap = {
            "mc": NS_MC,
            "w16cid": NS_W16CID,
        }
        root = etree.Element(
            _qn(NS_W16CID, "commentsIds"),
            nsmap=nsmap,
        )
        root.set(_qn(NS_MC, "Ignorable"), "w16cid")

        xml_content = etree.tostring(
            root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone="yes",
        )

        # Add part to document
        from docx.opc.packuri import PackURI
        from docx.opc.part import Part

        part = Part(
            PackURI("/word/commentsIds.xml"),
            CT_COMMENTS_IDS,
            xml_content,
            self._document.part.package,
        )
        self._document.part.relate_to(part, REL_COMMENTS_IDS)

    @property
    def xml(self) -> etree._Element:
        """Get the XML root element."""
        if self._xml is None:
            part = self._get_part()
            if part:
                self._xml = etree.fromstring(part.blob)
            else:
                # Return empty element if part doesn't exist
                self._xml = etree.Element(_qn(NS_W16CID, "commentsIds"))
        return self._xml

    def _save(self) -> None:
        """Save changes back to the part."""
        part = self._get_part()
        if part:
            part._blob = etree.tostring(
                self.xml,
                xml_declaration=True,
                encoding="UTF-8",
                standalone="yes",
            )

    def get_durable_ids(self) -> dict[str, str]:
        """
        Get durable IDs for all comments.

        Returns:
            Dict mapping para_id to durable_id.
        """
        result = {}
        for elem in self.xml:
            if etree.QName(elem).localname == "commentId":
                para_id = elem.get(_qn(NS_W16CID, "paraId"))
                durable_id = elem.get(_qn(NS_W16CID, "durableId"))
                if para_id and durable_id:
                    result[para_id] = durable_id
        return result

    def add_comment_id(self, para_id: str, durable_id: str) -> None:
        """
        Add a commentId entry.

        Args:
            para_id: Paragraph ID of the comment.
            durable_id: Durable ID for persistence.
        """
        elem = etree.SubElement(self.xml, _qn(NS_W16CID, "commentId"))
        elem.set(_qn(NS_W16CID, "paraId"), para_id)
        elem.set(_qn(NS_W16CID, "durableId"), durable_id)
        self._save()
