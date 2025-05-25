# DOCX Splitter

![PepeAgent](https://media.tenor.com/WF8QZgZzv0YAAAAC/pepe-agent.gif)



An AI slop web application that splits Microsoft Word (DOCX) files by headings and saves them as separate files.

## Features

- Split DOCX files by headings (H1-H6)
- Clean and modern web interface
- Automatic file naming based on headings
- Works in any modern web browser

## Usage

### Local Development
1. Clone this repository
2. Start the development server:
```bash
python3 -m http.server 3000
```
3. Open your browser and navigate to `http://localhost:3000`
4. Upload a DOCX file and click "Split Document"

### Web Server Deployment
1. Upload the following files to your web server:
   - `index.html`
   - `app.js`
   - `styles.css`
   - `package.json`
   - `package-lock.json`
   - `LICENSE`
   - `README.md`

2. Ensure your web server:
   - Serves static files
   - Has proper MIME type configuration for `.docx` files
   - Allows file downloads

3. Access the application through your web server's URL

## Dependencies

- mammoth: For DOCX file conversion
- file-saver: For saving files
- http-server: For serving the application

## License

This work is dedicated to the public domain under the CC0 1.0 Universal Public Domain Dedication. See LICENSE file for details.
