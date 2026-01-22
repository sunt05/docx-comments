# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- `unresolve_comment()` and `set_comment_resolved()` for toggling done status
- `delete_comment()` and `delete_thread()` for removing comments and threads
- `move_comment()` and `move_thread()` for re-anchoring comments

## [0.2.0] - 2026-01-21

### Added

- Optional `people.xml` identity linkage for comment authors
- `PersonInfo` data model for people.xml authors and presence metadata
- `CommentManager` people APIs: `get_people()`, `get_person()`, `ensure_person()`,
  `merge_people_from()`, `get_default_author_person()`
- System author resolution from Office profiles or a DOCX source via
  `DOCX_COMMENTS_AUTHOR_DOCX`

### Changed

- **Breaking**: `author` parameters now require `PersonInfo` instead of raw strings
- `add_comment()`/`reply_to_comment()` can optionally ensure people.xml entries

## [0.1.1] - 2026-01-21

### Changed

- Switch to git tag-based versioning via hatch-vcs

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

[Unreleased]: https://github.com/sunt05/docx-comments/compare/v0.2.0...HEAD
[0.2.0]: https://github.com/sunt05/docx-comments/compare/v0.1.1...v0.2.0
[0.1.1]: https://github.com/sunt05/docx-comments/compare/v0.1.0...v0.1.1
[0.1.0]: https://github.com/sunt05/docx-comments/releases/tag/v0.1.0
