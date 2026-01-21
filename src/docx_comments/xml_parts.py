"""Handlers for XML parts: comments.xml, commentsExtended.xml, and commentsIds.xml."""

from __future__ import annotations

from typing import TYPE_CHECKING, Optional

from docx.opc.packuri import PackURI
from docx.opc.part import Part
from lxml import etree

if TYPE_CHECKING:
    from docx import Document


# OOXML Namespaces
NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
NS_W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
NS_W16CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
NS_W16CEX = "http://schemas.microsoft.com/office/word/2018/wordml/cex"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"

# Relationship types
REL_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
REL_COMMENTS_EXT = (
    "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"
)
REL_COMMENTS_IDS = (
    "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds"
)
REL_COMMENTS_EXTENSIBLE = (
    "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible"
)

# Content types
CT_COMMENTS = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
CT_COMMENTS_EXT = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"
)
CT_COMMENTS_IDS = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"
)
CT_COMMENTS_EXTENSIBLE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml"
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
        self._xml: Optional[etree._Element] = None

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

        # Prefer public accessor when available (ensures _element initialized)
        if hasattr(part, "element"):
            try:
                elem = part.element
                if elem is not None:
                    return elem
            except (AttributeError, TypeError, ValueError, etree.XMLSyntaxError):
                # Best-effort fallback for python-docx element access.
                pass

        # Fallback for XmlPart with private _element (ensure initialized)
        if hasattr(part, "_element"):
            if getattr(part, "_element", None) is None:
                try:
                    part._element = etree.fromstring(part.blob)
                except (etree.XMLSyntaxError, AttributeError, TypeError):
                    # XMLSyntaxError: malformed XML in blob
                    # AttributeError: part lacks blob attribute
                    # TypeError: blob is None or wrong type
                    return etree.Element(_qn(NS_W, "comments"))
            return part._element

        # Generic Part - need to parse blob and maintain cache
        if self._xml is None:
            self._xml = etree.fromstring(part.blob)
        return self._xml

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
        if self._xml is not None:
            part._blob = etree.tostring(
                self._xml,
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
    - commentsExtensible.xml if missing
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

    # Ensure commentsExtensible.xml (modern comments metadata)
    extensible_part = CommentsExtensiblePart(document)
    extensible_part.ensure_exists()


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
        elem = etree.Element(_qn(NS_W15, "commentEx"))
        elem.set(_qn(NS_W15, "paraId"), para_id)
        elem.set(_qn(NS_W15, "done"), "1" if done else "0")
        if parent_para_id:
            elem.set(_qn(NS_W15, "paraIdParent"), parent_para_id)
        inserted = False
        if parent_para_id:
            for existing in self.xml:
                if (
                    etree.QName(existing).localname == "commentEx"
                    and existing.get(_qn(NS_W15, "paraId")) == parent_para_id
                ):
                    existing.addnext(elem)
                    inserted = True
                    break
        if not inserted:
            self.xml.append(elem)
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


class CommentsExtensiblePart:
    """Handler for word/commentsExtensible.xml part."""

    def __init__(self, document: Document) -> None:
        self._document = document
        self._xml: Optional[etree._Element] = None

    def _get_part(self):
        """Get the commentsExtensible part from document relationships."""
        doc_part = self._document.part
        for rel in doc_part.rels.values():
            if "commentsExtensible" in rel.reltype:
                return rel.target_part
        package = getattr(doc_part, "package", None)
        if package is not None:
            for part in getattr(package, "parts", []):
                if str(part.partname) == "/word/commentsExtensible.xml":
                    return part
        return None

    def ensure_exists(self) -> None:
        """Ensure the commentsExtensible part exists, creating if needed."""
        if self._get_part() is None:
            self._create_part()

    def _create_part(self) -> None:
        """Create a new commentsExtensible.xml part."""
        nsmap = {
            "mc": NS_MC,
            "w16cex": NS_W16CEX,
        }
        root = etree.Element(
            _qn(NS_W16CEX, "commentsExtensible"),
            nsmap=nsmap,
        )
        root.set(_qn(NS_MC, "Ignorable"), "w16cex")

        xml_content = etree.tostring(
            root,
            xml_declaration=True,
            encoding="UTF-8",
            standalone="yes",
        )

        part = Part(
            PackURI("/word/commentsExtensible.xml"),
            CT_COMMENTS_EXTENSIBLE,
            xml_content,
            self._document.part.package,
        )
        self._document.part.relate_to(part, REL_COMMENTS_EXTENSIBLE)

    @property
    def xml(self) -> etree._Element:
        """Get the XML root element."""
        if self._xml is None:
            part = self._get_part()
            if part:
                self._xml = etree.fromstring(part.blob)
            else:
                self._xml = etree.Element(_qn(NS_W16CEX, "commentsExtensible"))
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

    def get_extensible_info(self) -> dict[str, dict]:
        """
        Get metadata entries from commentsExtensible.xml.

        Returns:
            Dict mapping durable_id to {"date_utc": str|None}.
        """
        result = {}
        for elem in self.xml:
            if etree.QName(elem).localname == "commentExtensible":
                durable_id = elem.get(_qn(NS_W16CEX, "durableId"))
                date_utc = elem.get(_qn(NS_W16CEX, "dateUtc"))
                if durable_id:
                    result[durable_id] = {"date_utc": date_utc}
        return result

    def add_comment_extensible(self, durable_id: str, date_utc: Optional[str] = None) -> None:
        """
        Add or update a commentExtensible entry.

        Args:
            durable_id: Durable ID for the comment.
            date_utc: Optional UTC timestamp (ISO8601, Z-terminated).
        """
        for elem in self.xml:
            if (
                etree.QName(elem).localname == "commentExtensible"
                and elem.get(_qn(NS_W16CEX, "durableId")) == durable_id
            ):
                if date_utc and not elem.get(_qn(NS_W16CEX, "dateUtc")):
                    elem.set(_qn(NS_W16CEX, "dateUtc"), date_utc)
                    self._save()
                return

        elem = etree.SubElement(self.xml, _qn(NS_W16CEX, "commentExtensible"))
        elem.set(_qn(NS_W16CEX, "durableId"), durable_id)
        if date_utc:
            elem.set(_qn(NS_W16CEX, "dateUtc"), date_utc)
        self._save()


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
