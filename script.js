/**
 * CodeToTable Tool
 * A vanilla JS tool to convert code into styled HTML tables with line numbers.
 */

document.addEventListener('DOMContentLoaded', () => {
    const themeToggle = document.getElementById('theme-toggle');
    const convertBtn = document.getElementById('convert-btn');
    const clearBtn = document.getElementById('clear-btn');
    const copyCodeBtn = document.getElementById('copy-code-btn');
    const copyWordBtn = document.getElementById('copy-word-btn');
    const copyHtmlBtn = document.getElementById('copy-html-btn');
    const codeInput = document.getElementById('code-input');
    const languageSelect = document.getElementById('language-select');
    const fontSizeInput = document.getElementById('font-size');
    const fontFamilySelect = document.getElementById('font-family');
    const resultSection = document.getElementById('result-section');
    const tableOutputWrapper = document.getElementById('table-output-wrapper');
    const dropZone = document.getElementById('drop-zone');

    // --- Theme Management ---
    themeToggle.addEventListener('click', () => {
        const isDark = document.body.classList.contains('theme-dark');
        if (isDark) {
            document.body.classList.replace('theme-dark', 'theme-light');
            document.getElementById('prism-theme').href = 'https://cdnjs.cloudflare.com/ajax/libs/prism/1.29.0/themes/prism.min.css';
        } else {
            document.body.classList.replace('theme-light', 'theme-dark');
            document.getElementById('prism-theme').href = 'https://cdnjs.cloudflare.com/ajax/libs/prism/1.29.0/themes/prism-tomorrow.min.css';
        }
    });

    // --- File Drag & Drop ---
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('dragging');
    });

    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('dragging');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('dragging');
        const file = e.dataTransfer.files[0];
        if (file) {
            readFile(file);
        }
    });

    function readFile(file) {
        const reader = new FileReader();
        reader.onload = (e) => {
            codeInput.value = e.target.result;
            const ext = file.name.split('.').pop().toLowerCase();
            const langMap = {
                'js': 'javascript', 'py': 'python', 'java': 'java',
                'html': 'html', 'css': 'css', 'cpp': 'cpp',
                'go': 'go', 'rs': 'rust'
            };
            if (langMap[ext]) languageSelect.value = langMap[ext];
        };
        reader.readAsText(file);
    }

    // --- Core Logic: Table Rendering ---
    const TableRenderer = {
        render(code, language, fontSize, fontFamily) {
            const lines = code.split(/\r?\n/);
            const table = document.createElement('table');
            table.className = 'code-table';
            table.style.setProperty('--selected-font', fontFamily);
            table.style.setProperty('--selected-size', `${fontSize}px`);

            lines.forEach((lineText, index) => {
                const tr = document.createElement('tr');

                const tdGutter = document.createElement('td');
                tdGutter.className = 'gutter';
                tdGutter.textContent = index + 1;

                const tdCode = document.createElement('td');
                tdCode.className = 'code';

                if (language && language !== 'auto') {
                    const grammar = Prism.languages[language];
                    if (grammar) {
                        const highlighted = Prism.highlight(lineText || ' ', grammar, language);
                        tdCode.innerHTML = highlighted;
                    } else {
                        tdCode.textContent = lineText || ' ';
                    }
                } else {
                    tdCode.textContent = lineText || ' ';
                }

                tr.appendChild(tdGutter);
                tr.appendChild(tdCode);
                table.appendChild(tr);
            });

            return table;
        }
    };

    // --- Action Handlers ---
    convertBtn.addEventListener('click', () => {
        const code = codeInput.value.trim();
        if (!code) {
            alert('Please paste some code first!');
            return;
        }

        const lang = languageSelect.value;
        const fontSize = fontSizeInput.value;
        const fontFamily = fontFamilySelect.value;

        tableOutputWrapper.innerHTML = '';
        const table = TableRenderer.render(code, lang, fontSize, fontFamily);
        tableOutputWrapper.appendChild(table);

        resultSection.classList.remove('hidden');
        resultSection.scrollIntoView({ behavior: 'smooth' });
    });

    clearBtn.addEventListener('click', () => {
        codeInput.value = '';
        resultSection.classList.add('hidden');
        tableOutputWrapper.innerHTML = '';
    });

    copyCodeBtn.addEventListener('click', async () => {
        const table = tableOutputWrapper.querySelector('table');
        if (!table) return;

        const codeCells = table.querySelectorAll('.code');
        const codeText = Array.from(codeCells).map(cell => cell.textContent).join('\n');

        try {
            await navigator.clipboard.writeText(codeText);
            showCopyFeedback(copyCodeBtn, 'Copied Code!');
        } catch (err) {
            console.error('Failed to copy: ', err);
        }
    });

    // ─────────────────────────────────────────────────────────────────────────
    // FIXED: Copy for Word
    // Uses the modern Clipboard API with text/html MIME type so Word receives
    // a proper rich-text table, not plain text.
    // Inline styles are resolved BEFORE innerHTML is overwritten.
    // ─────────────────────────────────────────────────────────────────────────
    copyWordBtn.addEventListener('click', async () => {
        const table = tableOutputWrapper.querySelector('table');
        if (!table) return;

        const wordFontFamily = "Consolas, 'Courier New', monospace";
        const tableStyle = window.getComputedStyle(table);
        const tableBg = tableStyle.backgroundColor;

        // Build HTML rows by walking the ORIGINAL DOM so getComputedStyle works.
        // We collect everything as a string — no cloneNode needed.
        let rowsHtml = '';

        const rows = table.querySelectorAll('tr');
        rows.forEach((row) => {
            const gutter = row.querySelector('.gutter');
            const codeCell = row.querySelector('.code');

            const gStyle = window.getComputedStyle(gutter);
            const cStyle = window.getComputedStyle(codeCell);

            // Resolve span colors BEFORE building the code cell HTML string.
            // Walk every span in the live DOM and bake its computed color into a style attribute.
            const spans = codeCell.querySelectorAll('span');
            spans.forEach((span) => {
                const sColor = window.getComputedStyle(span).color;
                const sWeight = window.getComputedStyle(span).fontWeight;
                // Temporarily stamp the inline style so outerHTML captures it.
                span.setAttribute('data-word-style', `color:${sColor};font-weight:${sWeight};font-family:${wordFontFamily};`);
            });

            // Grab the innerHTML with baked span styles.
            // Replace data-word-style with style so Word sees them.
            let codeCellHtml = codeCell.innerHTML.replace(/data-word-style="/g, 'style="');

            // Clean up temp attributes from the live DOM.
            spans.forEach((span) => span.removeAttribute('data-word-style'));

            const gutterBg = gStyle.backgroundColor;
            const gutterColor = gStyle.color;
            const codeBg = cStyle.backgroundColor || tableBg;
            const codeColor = cStyle.color;

            rowsHtml += `
            <tr>
              <td style="
                width:40pt;
                text-align:right;
                padding:2pt 6pt;
                background-color:${gutterBg};
                color:${gutterColor};
                font-family:Arial,sans-serif;
                font-size:10pt;
                border-right:1px solid ${gutterColor};
                vertical-align:top;
                white-space:nowrap;
                user-select:none;
              ">${gutter.textContent}</td>
              <td style="
                padding:2pt 8pt;
                background-color:${codeBg};
                color:${codeColor};
                font-family:${wordFontFamily};
                font-size:10pt;
                white-space:pre;
                vertical-align:top;
              "><span style="font-family:${wordFontFamily};color:${codeColor};">${codeCellHtml}</span></td>
            </tr>`;
        });

        // Wrap in a full table with inline styles Word can parse.
        const wordHtml = `
        <html>
        <body>
        <table style="
          border-collapse:collapse;
          width:100%;
          background-color:${tableBg};
          font-size:10pt;
        " border="0" cellspacing="0" cellpadding="0">
          ${rowsHtml}
        </table>
        </body>
        </html>`;

        try {
            // Modern Clipboard API: write text/html so Word gets a real table.
            await navigator.clipboard.write([
                new ClipboardItem({
                    'text/html': new Blob([wordHtml], { type: 'text/html' }),
                    'text/plain': new Blob(
                        [Array.from(table.querySelectorAll('.code')).map(c => c.textContent).join('\n')],
                        { type: 'text/plain' }
                    )
                })
            ]);
            showCopyFeedback(copyWordBtn, 'Copied for Word!');
        } catch (err) {
            // Fallback: try execCommand with a hidden div as a last resort.
            console.warn('Clipboard API failed, trying execCommand fallback:', err);
            try {
                const container = document.createElement('div');
                container.style.cssText = 'position:fixed;left:-9999px;top:0;opacity:0;';
                container.innerHTML = wordHtml;
                document.body.appendChild(container);

                const range = document.createRange();
                range.selectNodeContents(container);
                const sel = window.getSelection();
                sel.removeAllRanges();
                sel.addRange(range);
                document.execCommand('copy');
                sel.removeAllRanges();
                document.body.removeChild(container);
                showCopyFeedback(copyWordBtn, 'Copied for Word!');
            } catch (fallbackErr) {
                console.error('Both copy methods failed:', fallbackErr);
                alert('Copy failed. Your browser may not support rich-text clipboard.\nTry using Chrome or Edge.');
            }
        }
    });

    copyHtmlBtn.addEventListener('click', async () => {
        const table = tableOutputWrapper.querySelector('table');
        if (!table) return;

        const clonedTable = table.cloneNode(true);

        try {
            await navigator.clipboard.writeText(clonedTable.outerHTML);
            showCopyFeedback(copyHtmlBtn, 'Copied HTML!');
        } catch (err) {
            console.error('Failed to copy: ', err);
        }
    });

    function showCopyFeedback(btn, text) {
        const originalText = btn.textContent;
        btn.textContent = text;
        btn.classList.add('btn-success');
        setTimeout(() => {
            btn.textContent = originalText;
            btn.classList.remove('btn-success');
        }, 2000);
    }
});
