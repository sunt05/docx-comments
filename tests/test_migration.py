"""Tests for metadata migration."""

from docx import Document

from docx_comments import CommentManager


class TestCommentMigration:
    """Tests for comment metadata migration."""

    def test_migrate_comment_metadata(self):
        """Test backfilling missing comment metadata."""
        from docx_comments.xml_parts import CommentsExtendedPart, CommentsIdsPart

        doc = Document()
        para = doc.add_paragraph("Text with comment")
        mgr = CommentManager(doc)
        mgr.add_comment(para, "Comment", "Author")

        # Remove paraId/textId from comments.xml
        ns_w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        ns_w14 = "http://schemas.microsoft.com/office/word/2010/wordml"
        comment_elem = mgr._comments_xml.find(f"{{{ns_w}}}comment")
        para_elem = comment_elem.find(f"{{{ns_w}}}p")
        para_elem.attrib.pop(f"{{{ns_w14}}}paraId", None)
        para_elem.attrib.pop(f"{{{ns_w14}}}textId", None)
        mgr._save_comments()

        # Clear threading and durable IDs
        ext_part = CommentsExtendedPart(doc)
        ids_part = CommentsIdsPart(doc)
        for elem in list(ext_part.xml):
            ext_part.xml.remove(elem)
        ext_part._save()
        for elem in list(ids_part.xml):
            ids_part.xml.remove(elem)
        ids_part._save()

        # Run migration
        mgr.migrate_comment_metadata()

        # Verify metadata restored
        para_id = para_elem.get(f"{{{ns_w14}}}paraId")
        text_id = para_elem.get(f"{{{ns_w14}}}textId")
        assert para_id is not None
        assert text_id is not None
        assert para_id in CommentsExtendedPart(doc).get_threading_info()
        assert para_id in CommentsIdsPart(doc).get_durable_ids()
