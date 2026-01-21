"""Tests for comment data models."""

from docx_comments.models import CommentInfo, CommentThread


class TestCommentInfo:
    """Tests for CommentInfo model."""

    def test_is_reply(self):
        """Test is_reply property."""
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
        root = CommentInfo("1", "ABC", "Root", "Author")
        reply1 = CommentInfo("2", "DEF", "Reply1", "Author", parent_para_id="ABC")
        reply2 = CommentInfo("3", "GHI", "Reply2", "Author", parent_para_id="ABC")

        thread = CommentThread(root=root, replies=[reply1, reply2])

        assert len(thread.all_comments) == 3
        assert thread.all_comments[0] is root
