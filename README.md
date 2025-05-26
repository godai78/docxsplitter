# DOCX Splitter

A web application that splits Microsoft Word (DOCX) files by headings and saves them as separate RTF files with proper formatting and character encoding.

## Features

- Split DOCX files by headings (H1-H6)
- Clean and modern web interface
- Automatic file naming based on headings
- Works in any modern web browser
- Generates valid RTF files with proper styling and formatting
- Or switch to the `htmlout` branch and output simple HTML files)
- Full Unicode support for all Latin-based scripts (Western, Eastern European, Nordic, Baltic, etc.)
- Automatic codepage detection from source documents
- Proper handling of diacritics and special characters
- Built-in Node.js server for easy deployment
- Smart delay between downloads to avoid browser limits

## Usage

### Local Development
1. Clone this repository
2. Install dependencies:
```bash
npm install
```
3. Start the server:
```bash
node server.js
```
4. Open your browser and navigate to `http://localhost:3003`
5. Upload a DOCX file and click "Split Document"

### Web Server Deployment
1. Upload the following files to your web server:
	- `index.html`
	- `app.js`
	- `styles.css`
	- `server.js`
	- `package.json`
	- `package-lock.json`
	- `LICENSE`
	- `README.md`
	- `pepeagent.gif`

2. Ensure your web server:
	- Serves static files
	- Has Node.js installed
	- Allows file downloads

3. Access the application through your web server's URL

## Features in Detail

### RTF Generation
- Proper paragraph formatting with double line breaks (`\par\par`)
- Line breaks within paragraphs using `\line`
- Bold formatting for headings
- UTF-8 encoding (codepage 65001) for full Unicode support
- Clean RTF structure with proper control words and grouping

### Character Support
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

### File Naming
- Automatic sanitization of filenames
- Support for diacritics in filenames
- Sequential numbering of output files
- Safe character replacement for special characters

### Browser Compatibility
- Smart delay between downloads to prevent browser limits
- Proper MIME type handling for RTF files
- Clean file download mechanism
- Progress indication during processing

## Dependencies

- mammoth: For DOCX file conversion
- file-saver: For saving files
- http: For serving the application

## License

This work is dedicated to the public domain under the CC0 1.0 Universal Public Domain Dedication. See LICENSE file for details.
