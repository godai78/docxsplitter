# DOCX Splitter

A web application that splits Microsoft Word (DOCX) files by headings and saves them as separate HTML files.

## Features

- Split DOCX files by headings (H1-H6)
- Clean and modern web interface
- Automatic file naming based on headings
- Works in any modern web browser
- Generates valid HTML files with proper styling
- Supports multiple language character sets in filenames
- Built-in Node.js server for easy deployment

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

## Dependencies

- mammoth: For DOCX file conversion
- file-saver: For saving files
- http: For serving the application

## License

This work is dedicated to the public domain under the CC0 1.0 Universal Public Domain Dedication. See LICENSE file for details.
