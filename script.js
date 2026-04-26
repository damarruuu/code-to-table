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
        if (file) readFile(file);
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
                        tdCode.innerHTML = Prism.highlight(lineText || ' ', grammar, language);
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
        if (!code) { alert('Please paste some code first!'); return; }

        tableOutputWrapper.innerHTML = '';
        const table = TableRenderer.render(
            code,
            languageSelect.value,
            fontSizeInput.value,
            fontFamilySelect.value
        );
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
        const codeText = Array.from(table.querySelectorAll('.code'))
            .map(c => c.textContent).join('\n');
        try {
            await navigator.clipboard.writeText(codeText);
            showCopyFeedback(copyCodeBtn, 'Copied Code!');
        } catch (err) { console.error(err); }
    });

    // ─────────────────────────────────────────────────────────────────────────
    // Copy for Word — builds a REAL bordered table that Word renders as a grid
    // ─────────────────────────────────────────────────────────────────────────
    copyWordBtn.addEventListener('click', async () => {
        const table = tableOutputWrapper.querySelector('table');
        if (!table) return;

        const WORD_MONO = "Consolas, 'Courier New', monospace";
        const BORDER    = "1px solid #d0d0d0";
        const GUTTER_BG = "#f6f8fa";
        const GUTTER_FG = "#6e7781";
        const CODE_BG   = "#ffffff";
        const CODE_FG   = "#24292e";
        const FONT_SIZE = "10pt";

        // Step 1: bake computed span colors into a temp attribute on the LIVE dom
        // (getComputedStyle only works on elements that are in the document)
        const allSpans = table.querySelectorAll('.code span');
        allSpans.forEach(span => {
            const cs = window.getComputedStyle(span);
            span.setAttribute('data-ws',
                `color:${cs.color};font-weight:${cs.fontWeight};font-style:${cs.fontStyle};font-family:${WORD_MONO};`
            );
        });

        // Step 2: build rows HTML — swap temp attr name to "style" in the string
        let rowsHtml = '';
        table.querySelectorAll('tr').forEach(row => {
            const gutter   = row.querySelector('.gutter');
            const codeCell = row.querySelector('.code');

            const codeCellHtml = codeCell.innerHTML
                .replace(/data-ws="/g, 'style="');

            rowsHtml += `
<tr>
  <td style="width:36pt;min-width:36pt;text-align:right;vertical-align:top;padding:1pt 6pt 1pt 4pt;background-color:${GUTTER_BG};color:${GUTTER_FG};font-family:Arial,sans-serif;font-size:${FONT_SIZE};border:${BORDER};white-space:nowrap;">${gutter.textContent}</td>
  <td style="vertical-align:top;padding:1pt 8pt;background-color:${CODE_BG};color:${CODE_FG};font-family:${WORD_MONO};font-size:${FONT_SIZE};border:${BORDER};white-space:pre;">${codeCellHtml}</td>
</tr>`;
        });

        // Step 3: clean up temp attributes from live DOM
        allSpans.forEach(span => span.removeAttribute('data-ws'));

        // Step 4: wrap in a Word-compatible HTML document
        // The mso- namespace hints and border="1" ensure Word draws the grid
        const wordHtml = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
<head><meta charset="utf-8"><style>table{border-collapse:collapse;}td{mso-border-alt:solid #d0d0d0 .5pt;}</style></head>
<body>
<table border="1" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:${BORDER};">
${rowsHtml}
</table>
</body></html>`;

        // Step 5: write as text/html to clipboard — Word reads this as a real table
        try {
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
            // Fallback for file:// or older browsers
            console.warn('Clipboard API blocked, trying execCommand fallback:', err);
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
            } catch (fb) {
                console.error('Both copy methods failed:', fb);
                alert('Copy failed. Serve the page over HTTPS, or use Chrome/Edge.');
            }
        }
    });

    copyHtmlBtn.addEventListener('click', async () => {
        const table = tableOutputWrapper.querySelector('table');
        if (!table) return;
        try {
            await navigator.clipboard.writeText(table.cloneNode(true).outerHTML);
            showCopyFeedback(copyHtmlBtn, 'Copied HTML!');
        } catch (err) { console.error(err); }
    });

    function showCopyFeedback(btn, text) {
        const original = btn.textContent;
        btn.textContent = text;
        btn.classList.add('btn-success');
        setTimeout(() => {
            btn.textContent = original;
            btn.classList.remove('btn-success');
        }, 2000);
    }
});
