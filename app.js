document.addEventListener('DOMContentLoaded', () => {
	const docxFile = document.getElementById('docxFile');
	const splitButton = document.getElementById('splitButton');
	const status = document.getElementById('status');
	const output = document.getElementById('output');
	const headingLevel = document.getElementById('headingLevel');

	splitButton.addEventListener('click', async () => {
		const file = docxFile.files[0];
		if (!file) {
			setStatus('Please select a file first', 'error');
			return;
		}

		// Check if the file is a supported format
		const fileExtension = file.name.toLowerCase().split('.').pop();
		if (!['docx', 'rtf'].includes(fileExtension)) {
			setStatus('Unsupported file format. Please use DOCX or RTF files.', 'error');
			return;
		}

		try {
			setStatus('Processing document...', 'info');
			let content;
			
			if (fileExtension === 'docx') {
				content = await processDocx(file);
			} else {
				content = await processRtf(file);
			}
			
			// Check if there are headers of the selected level
			if (!hasHeadersOfLevel(content, parseInt(headingLevel.value))) {
				setStatus('No headers of the selected size present in the document', 'error');
				return;
			}
			
			const sections = splitByHeadings(content, parseInt(headingLevel.value));
			 
			setStatus('Saving files...', 'info');
			await saveSectionsAsDocx(sections);
			 
			setStatus('Files have been saved successfully!', 'success');
			 
			// Display output information
			output.innerHTML = sections.map((section, index) => `
				<div class="output-item">
					<h3>File ${index + 1}</h3>
					<p>Title: ${section.title}</p>
				</div>
			`).join('');

		} catch (error) {
			setStatus(`Error: ${error.message}`, 'error');
			console.error('Error:', error);
		}
	});

	function setStatus(message, type = 'info') {
		status.className = `status ${type}`;
		status.textContent = message;
	}

	async function processDocx(file) {
		const reader = new FileReader();
		return new Promise((resolve, reject) => {
			reader.onload = async (event) => {
				try {
					const arrayBuffer = event.target.result;
					const result = await mammoth.convertToHtml({
						arrayBuffer: arrayBuffer,
						options: {
							styleMap: [],
							convertImage: mammoth.images.imgElement(() => {}),
							ignoreEmptyParagraphs: false,
							idPrefix: "",
							transformDocument: mammoth.transforms.paragraph(element => {
								// Extract codepage from the document if available
								const codepage = element.styleId ? element.styleId.match(/cp(\d+)/) : null;
								if (codepage) {
									globalCodepage = codepage[1];
								} else {
									globalCodepage = "65001"; // Fallback to UTF-8 if no codepage found
								}
								return element;
							})
						}
					});
					resolve(result.value);
				} catch (error) {
					reject(error);
				}
			};
			reader.onerror = (error) => reject(error);
			reader.readAsArrayBuffer(file);
		});
	}

	async function processRtf(file) {
		const reader = new FileReader();
		return new Promise((resolve, reject) => {
			reader.onload = async (event) => {
				try {
					// Read as ArrayBuffer to preserve encoding
					const arrayBuffer = event.target.result;
					
					// First read the header to detect codepage
					const headerView = new Uint8Array(arrayBuffer.slice(0, 1000));
					const headerText = new TextDecoder('latin1').decode(headerView);
					
					// Look for codepage declaration in RTF header
					const codepageMatch = headerText.match(/\\ansicpg(\d+)/);
					if (!codepageMatch) {
						throw new Error('Could not detect codepage in RTF file');
					}
					
					const codepage = codepageMatch[1];
					console.log('Detected codepage from RTF:', codepage);
					
					// Set the global codepage for consistent encoding handling
					globalCodepage = codepage;
					
					// Read the entire file using the detected codepage
					const decoder = new TextDecoder(`windows-${codepage}`);
					const rtfContent = decoder.decode(arrayBuffer);
					
					console.log('=== RTF Processing Debug ===');
					
					// Skip the RTF header and font table to get to the actual content
					const contentStart = rtfContent.indexOf('\\viewkind');
					if (contentStart === -1) {
						throw new Error('Could not find document content in RTF file');
					}
					
					const documentContent = rtfContent.substring(contentStart);
					console.log('Document content (first 2000 chars):', documentContent.substring(0, 2000));
					
					// First, let's find all paragraphs
					const paragraphs = documentContent.split('\\par');
					console.log('Found paragraphs:', paragraphs.length);
					
					let processedContent = '';
					let headerCount = 0;
					
					// Process each paragraph
					paragraphs.forEach(paragraph => {
						// Skip empty paragraphs
						if (!paragraph.trim()) return;
						
						// Check if this paragraph has header-like formatting
						const hasBold = paragraph.includes('\\b');
						const hasFontSize = paragraph.match(/\\fs(\d+)/);
						const fontSize = hasFontSize ? parseInt(hasFontSize[1]) : 0;
						
						// Extract the actual text, properly handling Unicode
						let text = paragraph
							.replace(/\\[a-z0-9]+\s?/g, '')
							.replace(/{|}/g, '')
							.replace(/^arsid\s*/i, '')  // Remove arsid heading if present
							.trim();
						
						// Convert hex escape sequences to their proper characters
						text = text.replace(/\\'([0-9a-f]{2})/gi, (match, hex) => {
							const code = parseInt(hex, 16);
							// Convert from Windows codepage to Unicode
							const buffer = new Uint8Array([code]);
							return new TextDecoder(`windows-${codepage}`).decode(buffer);
						});
						
						if (!text) return;
						
						console.log('Processing paragraph:', {
							text: text.substring(0, 50),
							hasBold,
							fontSize,
							original: paragraph.substring(0, 100)
						});
						
						// Determine if this is a header based on formatting
						if (hasBold || fontSize >= 20) {
							// Determine heading level based on font size
							let level = 1;
							if (fontSize >= 40) level = 1;
							else if (fontSize >= 32) level = 2;
							else if (fontSize >= 28) level = 3;
							else if (fontSize >= 24) level = 4;
							else if (fontSize >= 20) level = 5;
							else level = 6;
							
							console.log('Found header:', {
								text,
								level,
								fontSize
							});
							
							processedContent += `<h${level}>${text}</h${level}>\n`;
							headerCount++;
						} else {
							processedContent += `<p>${text}</p>\n`;
						}
					});
					
					console.log('Total headers found:', headerCount);
					console.log('Final HTML content:', processedContent);
					console.log('=== End RTF Processing Debug ===');
					
					resolve(processedContent);
				} catch (error) {
					console.error('Error processing RTF:', error);
					reject(error);
				}
			};
			reader.onerror = (error) => reject(error);
			reader.readAsArrayBuffer(file);
		});
	}

	// Global variable to store the codepage
	let globalCodepage = "65001"; // Default to UTF-8

	function splitByHeadings(html, targetLevel) {
		const parser = new DOMParser();
		const doc = parser.parseFromString(html, 'text/html');
		const sections = [];
		let currentSection = null;

		// Get all elements in order
		const elements = Array.from(doc.body.childNodes);
		 
		elements.forEach(element => {
			if (element.nodeType === Node.ELEMENT_NODE) {
				const headingLevel = getHeadingLevel(element);
				// Split at the target level and all levels above it (lower numbers)
				if (headingLevel && headingLevel <= targetLevel) {
					// Start new section
					if (currentSection) {
						sections.push(currentSection);
					}
					
					// Get the text content and remove any RTF artifacts
					const title = element.textContent
						.replace(/^[^a-zA-Z]*/, '')  // Remove everything before first letter
						.trim();
					
					currentSection = {
						title: title,
						content: ''
					};
				} else if (currentSection) {
					// Add content to current section
					currentSection.content += element.outerHTML;
				}
			}
		});

		if (currentSection) {
			sections.push(currentSection);
		}

		return sections;
	}

	function getHeadingLevel(element) {
		const tagName = element.tagName.toLowerCase();
		if (tagName.startsWith('h') && tagName.length === 2) {
			const level = parseInt(tagName[1]);
			return (level >= 1 && level <= 6) ? level : null;
		}
		return null;
	}

	function hasHeadersOfLevel(html, targetLevel) {
		console.log('=== Header Level Check Debug ===');
		console.log('Target level:', targetLevel);
		console.log('HTML content:', html);
		
		const parser = new DOMParser();
		const doc = parser.parseFromString(html, 'text/html');
		const elements = Array.from(doc.body.childNodes);
		
		console.log('Document elements:', elements.map(el => ({
			tagName: el.tagName,
			textContent: el.textContent?.substring(0, 50)
		})));
		
		const hasHeaders = elements.some(element => {
			if (element.nodeType === Node.ELEMENT_NODE) {
				const headingLevel = getHeadingLevel(element);
				console.log('Element:', element.tagName, 'Level:', headingLevel, 'Text:', element.textContent?.substring(0, 50));
				return headingLevel && headingLevel <= targetLevel;
			}
			return false;
		});
		
		console.log('Has headers:', hasHeaders);
		console.log('=== End Header Level Check Debug ===');
		return hasHeaders;
	}

	async function saveSectionsAsDocx(sections) {
		for (let i = 0; i < sections.length; i++) {
			const section = sections[i];
			const filename = `${String(i + 1).padStart(2, '0')}_${sanitizeFilename(section.title)}.rtf`;
			
			// Create RTF header with proper Unicode support
			const rtfHeader = `{\\rtf1\\ansi\\ansicpg65001\\deff0\\deflang1033
{\\fonttbl{\\f0\\fnil\\fcharset0 Arial;}}
{\\colortbl ;\\red0\\green0\\blue0;}
\\viewkind4\\uc1\\pard\\cf1\\f0\\fs24`;

			// Parse the HTML content
			const parser = new DOMParser();
			const doc = parser.parseFromString(section.content, 'text/html');

			// RTF formatting commands mapping
			const rtfCommands = {
				'b': { open: '\\b ', close: '\\b0 ' },
				'strong': { open: '\\b ', close: '\\b0 ' },
				'i': { open: '\\i ', close: '\\i0 ' },
				'em': { open: '\\i ', close: '\\i0 ' },
				'u': { open: '\\ul ', close: '\\ulnone ' },
				'sup': { open: '\\super ', close: '\\nosupersub ' },
				'sub': { open: '\\sub ', close: '\\nosupersub ' }
			};

			// Function to convert text to RTF format with proper Unicode handling
			function convertToRtfText(text) {
				return text.split('').map(char => {
					const code = char.charCodeAt(0);
					if (code < 128) {
						// ASCII characters
						if (char === '\\' || char === '{' || char === '}') {
							return '\\' + char;
						}
						return char;
					} else {
						// Unicode characters - use \uN? format
						return '\\u' + code + '?';
					}
				}).join('');
			}

			// Function to process a node and its children
			function processNode(node, formattingStack = []) {
				let rtfContent = '';

				if (node.nodeType === Node.TEXT_NODE) {
					// Apply current formatting stack and convert text
					const text = node.textContent;
					if (text) {
						// Preserve leading and trailing spaces
						const leadingSpace = text.match(/^\s*/)[0];
						const trailingSpace = text.match(/\s*$/)[0];
						const trimmedText = text.trim();
						
						if (trimmedText) {
							// Add leading space if exists
							if (leadingSpace) {
								rtfContent += convertToRtfText(leadingSpace);
							}
							
							// Wrap text in formatting groups if there are active formats
							if (formattingStack.length > 0) {
								rtfContent += '{' + formattingStack.join('') + convertToRtfText(trimmedText) + '}';
							} else {
								rtfContent += convertToRtfText(trimmedText);
							}
							
							// Add trailing space if exists
							if (trailingSpace) {
								rtfContent += convertToRtfText(trailingSpace);
							}
						} else {
							// If text is only whitespace, preserve it
							rtfContent += convertToRtfText(text);
						}
					}
				} else if (node.nodeType === Node.ELEMENT_NODE) {
					const tagName = node.tagName.toLowerCase();
					const rtfCommand = rtfCommands[tagName];

					if (rtfCommand) {
						// Handle formatting tags with proper grouping
						const newFormattingStack = [...formattingStack, rtfCommand.open];
						let childContent = '';
						
						// Process children with updated formatting stack
						Array.from(node.childNodes).forEach(child => {
							childContent += processNode(child, newFormattingStack);
						});
						
						// Wrap the formatted content in a group with proper closure
						rtfContent += '{' + childContent + rtfCommand.close + '}';
					} else if (tagName === 'p') {
						// Handle paragraphs
						rtfContent += '\\par\n';
						Array.from(node.childNodes).forEach(child => {
							rtfContent += processNode(child, [...formattingStack]);
						});
						rtfContent += '\\par\n';
					} else if (tagName === 'br') {
						// Handle line breaks
						rtfContent += '\\line\n';
					} else {
						// Process other elements
						Array.from(node.childNodes).forEach(child => {
							rtfContent += processNode(child, [...formattingStack]);
						});
					}
				}

				return rtfContent;
			}

			// Process the title with proper formatting
			const titleRtf = '{' + '\\b ' + convertToRtfText(section.title) + '\\b0' + '}';

			// Process the content
			const contentRtf = processNode(doc.body);

			// Create the final RTF content
			const rtfContent = rtfHeader + 
				'\\par\n\\par\n' + 
				titleRtf + 
				'\\par\n\\par\n' + 
				contentRtf + 
				'\\par\n\\par\n}';
			
			// Create a blob and save it
			const blob = new Blob([rtfContent], { type: 'application/rtf' });
			
			// Add a small delay between downloads to avoid browser limits
			await new Promise(resolve => setTimeout(resolve, 1000)); // 1 second delay
			
			// Create a temporary link to trigger download
			const link = document.createElement('a');
			link.href = URL.createObjectURL(blob);
			link.download = filename;
			document.body.appendChild(link);
			link.click();
			
			// Clean up
			URL.revokeObjectURL(link.href);
			document.body.removeChild(link);
		}
	}

	function sanitizeFilename(filename) {
		// Replace common diacritics with their base letters
		const diacritics = {
			// Polish
			'ą': 'a', 'ć': 'c', 'ę': 'e', 'ł': 'l', 'ń': 'n', 'ó': 'o', 'ś': 's', 'ź': 'z', 'ż': 'z',
			// German
			'ä': 'a', 'ö': 'o', 'ü': 'u', 'ß': 'ss',
			// French
			'à': 'a', 'â': 'a', 'ã': 'a', 'á': 'a', 'ä': 'a', 'å': 'a',
			'ç': 'c',
			'è': 'e', 'é': 'e', 'ê': 'e', 'ë': 'e',
			'ì': 'i', 'í': 'i', 'î': 'i', 'ï': 'i',
			'ñ': 'n',
			'ò': 'o', 'ó': 'o', 'ô': 'o', 'ö': 'o', 'õ': 'o',
			'ù': 'u', 'ú': 'u', 'û': 'u', 'ü': 'u',
			'ý': 'y',
			// Spanish
			'á': 'a', 'é': 'e', 'í': 'i', 'ó': 'o', 'ú': 'u',
			// Czech
			'č': 'c', 'ď': 'd', 'ě': 'e', 'ř': 'r', 'š': 's', 'ť': 't', 'ů': 'u', 'ž': 'z',
			// Hungarian
			'á': 'a', 'é': 'e', 'í': 'i', 'ó': 'o', 'ő': 'o', 'ú': 'u', 'ű': 'u',
			// Turkish
			'ç': 'c', 'ğ': 'g', 'ı': 'i', 'ö': 'o', 'ş': 's', 'ü': 'u',
			// Greek
			'α': 'a', 'β': 'b', 'γ': 'g', 'δ': 'd', 'ε': 'e', 'ζ': 'z', 'η': 'e',
			'θ': 'th', 'ι': 'i', 'κ': 'k', 'λ': 'l', 'μ': 'm', 'ν': 'n', 'ξ': 'x',
			'ο': 'o', 'π': 'p', 'ρ': 'r', 'σ': 's', 'τ': 't', 'υ': 'y', 'φ': 'f',
			'χ': 'ch', 'ψ': 'ps', 'ω': 'o',
			// Russian
			'а': 'a', 'б': 'b', 'в': 'v', 'г': 'g', 'д': 'd', 'е': 'e', 'ё': 'yo',
			'ж': 'zh', 'з': 'z', 'и': 'i', 'й': 'y', 'к': 'k', 'л': 'l', 'м': 'm',
			'н': 'n', 'о': 'o', 'п': 'p', 'р': 'r', 'с': 's', 'т': 't', 'у': 'u',
			'ф': 'f', 'х': 'h', 'ц': 'c', 'ч': 'ch', 'ш': 'sh', 'щ': 'shch',
			'ъ': '', 'ы': 'y', 'ь': '', 'э': 'e', 'ю': 'yu', 'я': 'ya',
			// Hebrew
			'א': 'a', 'ב': 'b', 'ג': 'g', 'ד': 'd', 'ה': 'h', 'ו': 'v', 'ז': 'z',
			'ח': 'kh', 'ט': 't', 'י': 'y', 'כ': 'k', 'ל': 'l', 'מ': 'm', 'נ': 'n',
			'ס': 's', 'ע': 'a', 'פ': 'p', 'צ': 'ts', 'ק': 'k', 'ר': 'r', 'ש': 'sh',
			'ת': 't',
			// Arabic
			'ا': 'a', 'ب': 'b', 'ت': 't', 'ث': 'th', 'ج': 'j', 'ح': 'h', 'خ': 'kh',
			'د': 'd', 'ذ': 'dh', 'ر': 'r', 'ز': 'z', 'س': 's', 'ش': 'sh', 'ص': 's',
			'ض': 'd', 'ط': 't', 'ظ': 'z', 'ع': 'a', 'غ': 'gh', 'ف': 'f', 'ق': 'q',
			'ك': 'k', 'ل': 'l', 'م': 'm', 'ن': 'n', 'ه': 'h', 'و': 'w', 'ي': 'y'
		};

		// Replace diacritics and special characters
		let sanitized = filename.toLowerCase();
		for (const [char, replacement] of Object.entries(diacritics)) {
			sanitized = sanitized.split(char).join(replacement);
		}

		// Replace remaining special characters with underscores
		sanitized = sanitized
			.replace(/[^a-z0-9]/g, '_')
			.replace(/_+/g, '_')
			.replace(/^_+|_+$/g, '');

		return sanitized;
	}
});
