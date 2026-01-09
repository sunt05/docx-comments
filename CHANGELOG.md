# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.1.0] - 2025-01-09

### Added

- Initial release
- `CommentManager` class for managing Word document comments
- `add_comment()` - Add anchored comments to specific text ranges
- `reply_to_comment()` - Create threaded replies to existing comments
- `resolve_comment()` - Mark comments as resolved (done status)
- `list_comments()` - List all comments in the document
- `get_comment_threads()` - Get comments grouped by thread
- `get_authors()` - Get all comment authors
- `get_document_author()` - Get document core properties author
- Full Word Online compatibility with proper OOXML structure
- Support for all four comment-related XML parts:
  - `comments.xml` - Comment content
  - `document.xml` - Anchors (commentRangeStart/End, commentReference)
  - `commentsExtended.xml` - Threading (paraId, paraIdParent, done)
  - `commentsIds.xml` - Durable IDs

[Unreleased]: https://github.com/sunt05/docx-comments/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/sunt05/docx-comments/releases/tag/v0.1.0
