import { LightningElement, track, api } from 'lwc';
import { loadScript } from 'lightning/platformResourceLoader';

import JSZip from '@salesforce/resourceUrl/JSZip';
import FastXML from '@salesforce/resourceUrl/FastXML';

import getRecordFields from '@salesforce/apex/WordMergeController.getRecordFields';
import generatePdfAndAttachToContact from '@salesforce/apex/PDFGeneratorService.generatePdfAndAttachToContact';

export default class WordMergeHtmlViewer extends LightningElement {
    @api recordId;
    @track isDownloadDisabled = true;
    @track fieldsFetched = false;
    @track fields = {};
    htmlOutput;

    FastXMLLoaded = false;
    JSZipLoaded = false;

    connectedCallback() {
        this.fetchFields();
    }

    renderedCallback() {
        if (!this.FastXMLLoaded) {
            loadScript(this, FastXML + '/fxparser.min.js')
                .then(() => {
                    if (window.fxparser?.XMLParser) {
                        window.XMLParser = window.fxparser.XMLParser;
                    } else {
                        console.error('XMLParser not found');
                    }
                    this.FastXMLLoaded = true;
                    return this.loadJSZip();
                })
                .catch(error => {
                    console.error('Error loading FastXML:', error);
                });
        }

        if (this.FastXMLLoaded && this.JSZipLoaded && !this.fieldsFetched) {
            this.fieldsFetched = true;
            this.fetchFields();
        }
    }

    loadJSZip() {
        if (!this.JSZipLoaded) {
            return loadScript(this, JSZip + '/jszip.min.js')
                .then(() => {
                    this.JSZipLoaded = true;
                })
                .catch(error => {
                    console.error('Error loading JSZip:', error);
                });
        }
        return Promise.resolve();
    }

    fetchFields() {
        getRecordFields({ recordId: this.recordId })
            .then(result => {
                this.fields = JSON.parse(JSON.stringify(result));
                this.isDownloadDisabled = false;
            })
            .catch(error => {
                console.error('Error fetching record fields:', error);
            });
    }

    async handleFileChange(event) {
        const file = event.target.files[0];
        if (!file) {
            console.warn('No file selected');
            return;
        }
        if (!this.JSZipLoaded || !this.FastXMLLoaded) {
            console.error('Libraries not loaded');
            return;
        }
        if (!this.fields || Object.keys(this.fields).length === 0) {
            console.error('Merge fields not fetched');
            return;
        }

        const reader = new FileReader();
        reader.onload = async () => {
            try {
                const arrayBuffer = reader.result;
                if (!window.JSZip || !window.XMLParser) {
                    console.error('JSZip or XMLParser missing');
                    return;
                }
                const zip = await window.JSZip.loadAsync(arrayBuffer);
                const documentXml = await zip.file('word/document.xml')?.async('text');
                const stylesXml = await zip.file('word/styles.xml')?.async('text');

                if (!documentXml || !stylesXml) {
                    console.error('Missing document.xml or styles.xml');
                    return;
                }

                const ParserClass = window.XMLParser.default;
                const parser = new ParserClass({ ignoreAttributes: false });

                const docJson = parser.parse(documentXml);
                const stylesJson = parser.parse(stylesXml);

                const css = this.extractCssFromStyles(stylesJson);
                const html = this.buildHtmlWithStyles(docJson, stylesJson);
                const finalHtml = this.mergeFields(html, this.fields);
                const inlinedHtml = this.inlineCssFromClasses(finalHtml, css);

                this.htmlOutput = inlinedHtml;
                this.renderHtmlToDom(this.htmlOutput);

            } catch (err) {
                console.error('Error processing file:', err);
            }
        };
        reader.onerror = e => {
            console.error('FileReader error:', e);
        };
        reader.readAsArrayBuffer(file);
    }

    extractCssFromStyles(stylesJson) {
        let cssOutput = '';
        const styles = stylesJson['w:styles']['w:style'];
        const styleArray = Array.isArray(styles) ? styles : [styles];

        styleArray.forEach(style => {
            const styleId = style['@_w:styleId'];
            const rPr = style['w:rPr'];
            if (!rPr || !styleId) return;

            let css = '';
            const colorVal = rPr['w:color']?.['@_w:val'];
            if (colorVal && colorVal.toLowerCase() !== 'auto') {
                css += `color: #${colorVal}; `;
            }
            if (rPr['w:b']) {
                css += `font-weight: bold; `;
            }
            if (rPr['w:i']) {
                css += `font-style: italic; `;
            }
            if (rPr['w:sz']) {
                const fontSize = parseInt(rPr['w:sz']['@_w:val'], 10) / 2;
                css += `font-size: ${fontSize}pt; `;
            }

            cssOutput += `.${styleId} { ${css} }\n`;
        });

        return cssOutput;
    }

    buildHtmlWithStyles(docJson, stylesJson) {
        const body = docJson['w:document']['w:body'];
        let html = '';

        // Handle paragraphs
        const paras = body['w:p'];
        if (paras) {
            const paraArray = Array.isArray(paras) ? paras : [paras];
            paraArray.forEach(p => {
                html += this.htmlFromParagraph(p);
            });
        }

        // Handle tables
        const tables = body['w:tbl'];
        if (tables) {
            const tblArray = Array.isArray(tables) ? tables : [tables];
            tblArray.forEach(tbl => {
                html += this.htmlFromTable(tbl, stylesJson);
            });
        }

        return html;
    }

    htmlFromParagraph(p) {
        const pStyle = p['w:pPr']?.['w:pStyle']?.['@_w:val'] || 'Normal';
        const runs = p['w:r'];
        const inner = this.htmlFromRuns(runs);
        if (inner.trim()) {
            return `<p class="${pStyle}">${inner}</p>`;
        }
        return '';
    }

    htmlFromRuns(runs) {
        const runArray = Array.isArray(runs) ? runs : [runs];
        return runArray.map(run => {
            let text = '';
            const t = run['w:t'];
            if (t) {
                text = (typeof t === 'string') ? t : (t['#text'] || t['_text'] || '');
            }

            const rPr = run['w:rPr'];
            if (rPr) {
                let style = '';
                if (rPr['w:b']) style += 'font-weight:bold;';
                if (rPr['w:i']) style += 'font-style:italic;';
                if (rPr['w:u']) style += 'text-decoration:underline;';
                if (rPr['w:color'] && rPr['w:color']['@_w:val'] && rPr['w:color']['@_w:val'].toLowerCase() !== 'auto') {
                    style += `color:#${rPr['w:color']['@_w:val']};`;
                }
                if (rPr['w:sz']) {
                    const fontSize = parseInt(rPr['w:sz']['@_w:val'], 10) / 2;
                    style += `font-size:${fontSize}pt;`;
                }
                return `<span style="${style}">${text}</span>`;
            }
            return text;
        }).join('');
    }

    lookupTableStyle(styleId, stylesJson) {
        const styles = stylesJson['w:styles']['w:style'];
        const styleArray = Array.isArray(styles) ? styles : [styles];
        return styleArray.find(s => s['@_w:styleId'] === styleId);
    }

    htmlFromTable(tbl, stylesJson) {
        // 1. Look for a style reference in the table
        const tblPr = tbl['w:tblPr'];
        const styleRef = tblPr?.['w:tblStyle']?.['@_w:val'];
        let styleDef = null;
        if (styleRef) {
            styleDef = this.lookupTableStyle(styleRef, stylesJson);
        }

        // 2. Extract style-level border definitions (if styleDef has them)
        let styleBorders = {};
        if (styleDef && styleDef['w:tblPr'] && styleDef['w:tblPr']['w:tblBorders']) {
            styleBorders = styleDef['w:tblPr']['w:tblBorders'];
        }

        // 3. Extract table-level overrides (explicit tblBorders in tblPr)
        let overrideBorders = {};
        if (tblPr && tblPr['w:tblBorders']) {
            overrideBorders = tblPr['w:tblBorders'];
        }

        // Helper to compute css string for a border element
        const borderToCss = (borderObj, side) => {
            if (!borderObj) return '';
            const b = borderObj[`w:${side}`];
            if (!b || !b['@_w:val']) return '';
            // val might be “single”, “double”, etc.  sz is size, color is color val
            const val = b['@_w:val'];
            const sz = b['@_w:sz'] ? parseInt(b['@_w:sz'], 10) : null;
            const color = b['@_w:color'];
            // Convert sz (Word uses e.g. twips or eighths of a point) to px approx
            let widthPx = '';
            if (sz !== null) {
                // A rough heuristic: Word “sz” is in ½‑points (or in eighths) — you may calibrate
                widthPx = `${sz / 8}px`;
            }
            let css = `border-${side}: ${widthPx || '1px'} ${val} ${color ? `#${color}` : 'black'};`;
            return css;
        };

        // Build a combined border CSS for the table
        let tableStyle = 'border-collapse: collapse; ';
        const sides = ['top','left','bottom','right','insideH','insideV'];
        sides.forEach(side => {
            // Table-level override takes precedence, else style-level
            const borderCss = borderToCss(overrideBorders, side) || borderToCss(styleBorders, side);
            tableStyle += borderCss;
        });

        // Build the table HTML
        let html = `<table style="${tableStyle}">`;

        const rows = tbl['w:tr'];
        if (rows) {
            const rowArray = Array.isArray(rows) ? rows : [rows];
            rowArray.forEach(tr => {
                html += '<tr>';
                const cells = tr['w:tc'];
                if (cells) {
                    const cellArray = Array.isArray(cells) ? cells : [cells];
                    cellArray.forEach(tc => {
                        // Within each cell, you might have tcPr (for cell-level borders) and paragraphs
                        let cellStyle = '';
                        const tcPr = tc['w:tcPr'];
                        if (tcPr && tcPr['w:tcBorders']) {
                            const b = tcPr['w:tcBorders'];
                            // check top, left, bottom, right
                            ['top','left','bottom','right'].forEach(side => {
                                const sideCss = borderToCss(b, side);
                                if (sideCss) cellStyle += sideCss;
                            });
                        }
                        // fallback: if no cell-level border, you might inherit from table
                        if (!cellStyle) {
                            // Optionally inherit table border for each cell
                            // e.g. `border: 1px solid black;`
                            // But we might skip because table border covers edges
                        }

                        // Extract content inside the cell (paragraphs)
                        let cellHtml = '';
                        const cellParas = tc['w:p'];
                        if (cellParas) {
                            const cpArray = Array.isArray(cellParas) ? cellParas : [cellParas];
                            cpArray.forEach(cp => {
                                cellHtml += this.htmlFromParagraph(cp);
                            });
                        }

                        html += `<td style="${cellStyle}">${cellHtml}</td>`;
                    });
                }
                html += '</tr>';
            });
        }

        html += '</table>';
        return html;
    }

    mergeFields(html, fields) {
        let merged = html;
        Object.keys(fields).forEach(key => {
            const regex = new RegExp(`{{${key}}}`, 'gi');
            merged = merged.replace(regex, fields[key] || '');
        });
        merged = merged.replace(/{{(.*?)}}/g, '<span class="unmerged">$1</span>');
        return merged;
    }

    inlineCssFromClasses(html, cssText) {
        const cssMap = {};
        const regex = /\.([^{\s]+)\s*{([^}]*)}/g;
        let match;
        while ((match = regex.exec(cssText)) !== null) {
            const cls = match[1];
            let styles = match[2].trim().replace(/\s+/g, ' ');
            cssMap[cls] = styles;
        }

        const parser = new DOMParser();
        const doc = parser.parseFromString(html, 'text/html');

        Object.keys(cssMap).forEach(cls => {
            const els = doc.querySelectorAll('.' + cls);
            els.forEach(el => {
                const existing = el.getAttribute('style') || '';
                el.setAttribute('style', `${existing} ${cssMap[cls]}`.trim());
                el.classList.remove(cls);
            });
        });

        return doc.body.innerHTML;
    }

    renderHtmlToDom(html) {
        const container = this.template.querySelector('.doc-preview');
        if (container) container.innerHTML = html || '<p>No content</p>';
    }

    handlePreviewAsPdf() {
        if (!this.htmlOutput) {
            console.warn('No content to preview');
            return;
        }
        const encoded = btoa(encodeURIComponent(this.htmlOutput));
        const vfUrl = '/apex/DownloadWordPage?html=' + encoded;
        window.open(vfUrl, '_blank');
    }

    handleSavePdfToSalesforce() {
        if (!this.htmlOutput) {
            console.warn('No content to save');
            return;
        }
        generatePdfAndAttachToContact({
            htmlContent: this.htmlOutput,
            contactId: this.recordId
        })
        .then(() => {
            console.log('Saved PDF');
        })
        .catch(error => {
            console.error('Error saving PDF:', error);
        });
    }

    handleBothSaveAndDownload() {
        this.handleSavePdfToSalesforce();
        this.handlePreviewAsPdf();
    }
}
