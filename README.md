# DOCX Splitter v 0.6

A web application that splits Microsoft Word `.docx` files and RTF files by selected heading level and saves them as separate `.rtf` files.

![Pepe Agent](pepeagent.gif)

## Features

- Split `.docx` and `.rtf` files by headings (H1-H6)
- Select which heading level to split at (splits at selected level and all levels above it)
- Clean and modern web interface
- Automatic file naming based on headings
- Works in any modern web browser
- Generates valid `.rtf` files with proper styling and formatting
- Automatic code page detection from source documents and proper handling of diacritics and special characters
- Full Unicode support for all Latin-based scripts (Western, Eastern European, Nordic, Baltic, etc.)
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

### GitHub Pages

A fully functional test version of the app based off the `master` branch runs at https://godai78.github.io/docxsplitter/

## License

Ths app was vibe-coded by godai78 using Windsurf and Cursor IDEs.

This work is dedicated to the public domain under the CC0 1.0 Universal Public Domain Dedication. See LICENSE file for details.
