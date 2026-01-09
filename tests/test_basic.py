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
