# Features in Detail

## RTF Generation
- Proper paragraph formatting with double line breaks (`\par\par`)
- Line breaks within paragraphs using `\line`
- Bold formatting for headings
- UTF-8 encoding (codepage 65001) for full Unicode support
- Clean RTF structure with proper control words and grouping
- Comprehensive text formatting support:
  - Bold text (`<b>` and `<strong>` tags)
  - Italic text (`<i>` and `<em>` tags)
  - Underlined text (`<u>` tags)
  - Superscript text (`<sup>` tags)
  - Subscript text (`<sub>` tags)
  - Nested formatting (e.g., bold text within italic text)
- Proper handling of formatting hierarchy and nesting
- Consistent formatting state management across paragraphs
- Automatic formatting command closure to prevent formatting leaks

## Character Support
- Automatic detection of document codepage
- Fallback to UTF-8 if no codepage is specified
- Proper handling of all Latin-based scripts:
  - Western European (French, German, Spanish, etc.)
  - Eastern European (Polish, Czech, Hungarian, etc.)
  - Nordic (Swedish, Danish, Norwegian)
  - Baltic (Latvian, Lithuanian)
  - Turkish
  - Romanian
- Proper escaping of RTF special characters

## File Naming
- Automatic sanitization of filenames
- Support for diacritics in filenames
- Sequential numbering of output files
- Safe character replacement for special characters

## Browser Compatibility
- Smart delay between downloads to prevent browser limits
- Proper MIME type handling for RTF files
- Clean file download mechanism
- Progress indication during processing

## Dependencies

- mammoth: For DOCX file conversion
- file-saver: For saving files
- http: For serving the application