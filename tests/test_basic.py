"""Basic tests for docx-comments module."""

import pytest
from docx import Document

from docx_comments import CommentManager


class TestCommentManager:
    """Tests for CommentManager class."""

    def test_init_new_document(self):
        """Test initializing with a new document."""
        doc = Document()
        doc.add_paragraph("Test paragraph")
        mgr = CommentManager(doc)
        assert mgr is not None

    def test_list_comments_empty(self):
        """Test listing comments on document with no comments."""
        doc = Document()
        doc.add_paragraph("Test paragraph")
        mgr = CommentManager(doc)
        comments = list(mgr.list_comments())
        assert len(comments) == 0

    def test_add_comment(self, tmp_path):
        """Test adding a comment to a paragraph."""
        doc = Document()
        para = doc.add_paragraph("This is test text to comment on.")
        mgr = CommentManager(doc)

        comment_id = mgr.add_comment(
            paragraph=para,
            text="This is a test comment",
            author="Test Author",
            initials="TA",
        )

        assert comment_id is not None
        assert len(comment_id) > 0

        # Verify comment was added
        comments = list(mgr.list_comments())
        assert len(comments) == 1
        assert comments[0].text == "This is a test comment"
        assert comments[0].author == "Test Author"
        assert comments[0].initials == "TA"

        # Save and reload
        output_path = tmp_path / "test_output.docx"
        doc.save(str(output_path))

        # Reload and verify
        doc2 = Document(str(output_path))
        mgr2 = CommentManager(doc2)
        comments2 = list(mgr2.list_comments())
        assert len(comments2) == 1

    def test_reply_to_comment(self, tmp_path):
        """Test replying to a comment."""
        doc = Document()
        para = doc.add_paragraph("Text to comment on.")
        mgr = CommentManager(doc)

        # Add root comment
        root_id = mgr.add_comment(
            paragraph=para,
            text="Root comment",
            author="Author1",
        )

        # Add reply
        reply_id = mgr.reply_to_comment(
            parent_id=root_id,
            text="Reply comment",
            author="Author2",
        )

        assert reply_id != root_id

        # Check threading
        threads = mgr.get_comment_threads()
        assert len(threads) == 1
        assert threads[0].root.text == "Root comment"
        assert len(threads[0].replies) == 1
        assert threads[0].replies[0].text == "Reply comment"

    def test_resolve_comment(self):
        """Test marking a comment as resolved."""
        doc = Document()
        para = doc.add_paragraph("Test text")
        mgr = CommentManager(doc)

        comment_id = mgr.add_comment(
            paragraph=para,
            text="Comment to resolve",
            author="Author",
        )

        # Initially not resolved
        comments = list(mgr.list_comments())
        assert not comments[0].is_resolved

        # Resolve
        mgr.resolve_comment(comment_id)

        # Verify resolved
        comments = list(mgr.list_comments())
        assert comments[0].is_resolved

    def test_get_comment_threads(self):
        """Test getting comment threads."""
        doc = Document()
        para1 = doc.add_paragraph("First paragraph")
        para2 = doc.add_paragraph("Second paragraph")
        mgr = CommentManager(doc)

        # Add two independent comments
        id1 = mgr.add_comment(para1, "Comment 1", "Author1")
        id2 = mgr.add_comment(para2, "Comment 2", "Author2")

        # Add replies to first comment
        mgr.reply_to_comment(id1, "Reply 1a", "Author3")
        mgr.reply_to_comment(id1, "Reply 1b", "Author4")

        threads = mgr.get_comment_threads()
        assert len(threads) == 2

        # Find thread with replies
        thread_with_replies = next(t for t in threads if t.reply_count > 0)
        assert thread_with_replies.root.text == "Comment 1"
        assert thread_with_replies.reply_count == 2


class TestCommentInfo:
    """Tests for CommentInfo model."""

    def test_is_reply(self):
        """Test is_reply property."""
        from docx_comments.models import CommentInfo

        root = CommentInfo(
            comment_id="1",
            para_id="ABC",
            text="Root",
            author="Author",
        )
        assert not root.is_reply

        reply = CommentInfo(
            comment_id="2",
            para_id="DEF",
            text="Reply",
            author="Author",
            parent_para_id="ABC",
        )
        assert reply.is_reply


class TestCommentThread:
    """Tests for CommentThread model."""

    def test_all_comments(self):
        """Test all_comments property."""
        from docx_comments.models import CommentInfo, CommentThread

        root = CommentInfo("1", "ABC", "Root", "Author")
        reply1 = CommentInfo("2", "DEF", "Reply1", "Author", parent_para_id="ABC")
        reply2 = CommentInfo("3", "GHI", "Reply2", "Author", parent_para_id="ABC")

        thread = CommentThread(root=root, replies=[reply1, reply2])

        assert len(thread.all_comments) == 3
        assert thread.all_comments[0] is root


class TestWordOnlineCompatibility:
    """Tests for Word Online compatibility by validating XML structure."""

    def test_xml_parts_created(self, tmp_path):
        """Test that all required XML parts are created."""
        from zipfile import ZipFile

        doc = Document()
        para = doc.add_paragraph("Test text")
        mgr = CommentManager(doc)

        # Add a comment
        mgr.add_comment(para, "Test comment", "Author")

        # Save document
        output_path = tmp_path / "test_parts.docx"
        doc.save(str(output_path))

        # Check ZIP contents
        with ZipFile(str(output_path), "r") as zf:
            names = zf.namelist()
            assert "word/comments.xml" in names
            assert "word/commentsExtended.xml" in names
            assert "word/commentsIds.xml" in names

    def test_comments_xml_structure(self, tmp_path):
        """Test comments.xml has correct structure."""
        from lxml import etree
        from zipfile import ZipFile

        doc = Document()
        para = doc.add_paragraph("Test text")
        mgr = CommentManager(doc)

        comment_id = mgr.add_comment(
            para, "Review this section", "Reviewer", initials="R"
        )

        output_path = tmp_path / "test_comments_xml.docx"
        doc.save(str(output_path))

        with ZipFile(str(output_path), "r") as zf:
            xml = etree.fromstring(zf.read("word/comments.xml"))

        # Check root element
        assert xml.tag.endswith("}comments")

        # Find comment element
        ns_w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        comments = xml.findall(f"{{{ns_w}}}comment")
        assert len(comments) == 1

        comment = comments[0]
        assert comment.get(f"{{{ns_w}}}author") == "Reviewer"
        assert comment.get(f"{{{ns_w}}}initials") == "R"
        assert comment.get(f"{{{ns_w}}}id") == comment_id

    def test_threading_xml_structure(self, tmp_path):
        """Test commentsExtended.xml has correct threading structure."""
        from lxml import etree
        from zipfile import ZipFile

        doc = Document()
        para = doc.add_paragraph("Test text")
        mgr = CommentManager(doc)

        # Add root comment and reply
        root_id = mgr.add_comment(para, "Root comment", "Author1")
        reply_id = mgr.reply_to_comment(root_id, "Reply comment", "Author2")

        output_path = tmp_path / "test_threading.docx"
        doc.save(str(output_path))

        with ZipFile(str(output_path), "r") as zf:
            xml = etree.fromstring(zf.read("word/commentsExtended.xml"))

        # Check root element
        assert xml.tag.endswith("}commentsEx")

        # Find commentEx elements
        ns_w15 = "http://schemas.microsoft.com/office/word/2012/wordml"
        comment_exs = xml.findall(f"{{{ns_w15}}}commentEx")
        assert len(comment_exs) == 2

        # Check that reply has parent link
        para_ids = set()
        parent_links = {}
        for ce in comment_exs:
            para_id = ce.get(f"{{{ns_w15}}}paraId")
            parent = ce.get(f"{{{ns_w15}}}paraIdParent")
            para_ids.add(para_id)
            if parent:
                parent_links[para_id] = parent

        # One comment should have a parent link
        assert len(parent_links) == 1
        # The parent should exist
        assert list(parent_links.values())[0] in para_ids

    def test_resolved_status_in_xml(self, tmp_path):
        """Test that resolved status is correctly saved in XML."""
        from lxml import etree
        from zipfile import ZipFile

        doc = Document()
        para = doc.add_paragraph("Test text")
        mgr = CommentManager(doc)

        comment_id = mgr.add_comment(para, "Comment to resolve", "Author")
        mgr.resolve_comment(comment_id)

        output_path = tmp_path / "test_resolved.docx"
        doc.save(str(output_path))

        with ZipFile(str(output_path), "r") as zf:
            xml = etree.fromstring(zf.read("word/commentsExtended.xml"))

        ns_w15 = "http://schemas.microsoft.com/office/word/2012/wordml"
        comment_exs = xml.findall(f"{{{ns_w15}}}commentEx")
        assert len(comment_exs) == 1
        assert comment_exs[0].get(f"{{{ns_w15}}}done") == "1"

    def test_document_xml_anchors(self, tmp_path):
        """Test that document.xml has proper comment anchors."""
        from lxml import etree
        from zipfile import ZipFile

        doc = Document()
        para = doc.add_paragraph("This is test text to comment on.")
        mgr = CommentManager(doc)

        comment_id = mgr.add_comment(para, "Test comment", "Author")

        output_path = tmp_path / "test_anchors.docx"
        doc.save(str(output_path))

        with ZipFile(str(output_path), "r") as zf:
            xml = etree.fromstring(zf.read("word/document.xml"))

        ns_w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        # Find comment anchors
        range_start = xml.find(f".//{{{ns_w}}}commentRangeStart")
        range_end = xml.find(f".//{{{ns_w}}}commentRangeEnd")
        comment_ref = xml.find(f".//{{{ns_w}}}commentReference")

        assert range_start is not None, "commentRangeStart not found"
        assert range_end is not None, "commentRangeEnd not found"
        assert comment_ref is not None, "commentReference not found"

        # Verify IDs match
        assert range_start.get(f"{{{ns_w}}}id") == comment_id
        assert range_end.get(f"{{{ns_w}}}id") == comment_id
        assert comment_ref.get(f"{{{ns_w}}}id") == comment_id

    def test_full_roundtrip(self, tmp_path):
        """Test full save/reload roundtrip with all features."""
        doc = Document()
        para1 = doc.add_paragraph("First paragraph for comments")
        para2 = doc.add_paragraph("Second paragraph for comments")
        mgr = CommentManager(doc)

        # Add various comments
        id1 = mgr.add_comment(para1, "Comment on first para", "Alice", "A")
        id2 = mgr.add_comment(para2, "Comment on second para", "Bob", "B")
        id3 = mgr.reply_to_comment(id1, "Reply to Alice", "Charlie", "C")
        mgr.resolve_comment(id2)

        # Save
        output_path = tmp_path / "test_roundtrip.docx"
        doc.save(str(output_path))

        # Reload
        doc2 = Document(str(output_path))
        mgr2 = CommentManager(doc2)

        # Verify comments
        comments = list(mgr2.list_comments())
        assert len(comments) == 3

        # Verify threads
        threads = mgr2.get_comment_threads()
        assert len(threads) == 2

        # Find thread with reply
        thread_with_reply = next(t for t in threads if t.reply_count > 0)
        assert thread_with_reply.root.text == "Comment on first para"
        assert thread_with_reply.replies[0].text == "Reply to Alice"

        # Find resolved thread
        resolved_thread = next(t for t in threads if t.root.text == "Comment on second para")
        assert resolved_thread.is_resolved
