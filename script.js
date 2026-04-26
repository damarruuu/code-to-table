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
    dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('dragging'); });
    dropZone.addEventListener('dragleave', () => { dropZone.classList.remove('dragging'); });
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
            const langMap = { 'js':'javascript','py':'python','java':'java','html':'html','css':'css','cpp':'cpp','go':'go','rs':'rust' };
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
        const table = TableRenderer.render(code, languageSelect.value, fontSizeInput.value, fontFamilySelect.value);
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
        const codeText = Array.from(table.querySelectorAll('.code')).map(c => c.textContent).join('\n');
        try {
            await navigator.clipboard.writeText(codeText);
            showCopyFeedback(copyCodeBtn, 'Copied Code!');
        } catch (err) { console.error(err); }
    });

    // ─────────────────────────────────────────────────────────────────────────
    // Prism token → inline color maps
    // Dark theme = prism-tomorrow, Light theme = prism (default)
    // These match the actual CSS colors from the Prism CDN stylesheets exactly.
    // Word ignores CSS classes — every color MUST be inline style.
    // ─────────────────────────────────────────────────────────────────────────
    const PRISM_DARK_COLORS = {
        'comment':        { color: '#999999', style: 'italic' },
        'prolog':         { color: '#999999' },
        'doctype':        { color: '#999999' },
        'cdata':          { color: '#999999' },
        'punctuation':    { color: '#cccccc' },
        'property':       { color: '#f8c555' },
        'tag':            { color: '#f8c555' },
        'constant':       { color: '#f8c555' },
        'symbol':         { color: '#f8c555' },
        'deleted':        { color: '#f8c555' },
        'boolean':        { color: '#f08d49' },
        'number':         { color: '#f08d49' },
        'selector':       { color: '#7ec699' },
        'attr-name':      { color: '#7ec699' },
        'string':         { color: '#7ec699' },
        'char':           { color: '#7ec699' },
        'builtin':        { color: '#7ec699' },
        'inserted':       { color: '#7ec699' },
        'operator':       { color: '#67cdcc' },
        'entity':         { color: '#67cdcc' },
        'url':            { color: '#67cdcc' },
        'variable':       { color: '#e2777a' },
        'atrule':         { color: '#cc99cd' },
        'attr-value':     { color: '#cc99cd' },
        'function':       { color: '#cc99cd' },
        'keyword':        { color: '#cc99cd' },
        'regex':          { color: '#de9b30' },
        'important':      { color: '#de9b30', weight: 'bold' },
    };

    const PRISM_LIGHT_COLORS = {
        'comment':        { color: '#708090', style: 'italic' },
        'prolog':         { color: '#708090' },
        'doctype':        { color: '#708090' },
        'cdata':          { color: '#708090' },
        'punctuation':    { color: '#999999' },
        'property':       { color: '#905' },
        'tag':            { color: '#905' },
        'boolean':        { color: '#905' },
        'number':         { color: '#905' },
        'constant':       { color: '#905' },
        'symbol':         { color: '#905' },
        'deleted':        { color: '#905' },
        'selector':       { color: '#690' },
        'attr-name':      { color: '#690' },
        'string':         { color: '#690' },
        'char':           { color: '#690' },
        'builtin':        { color: '#690' },
        'inserted':       { color: '#690' },
        'operator':       { color: '#9a6e3a' },
        'entity':         { color: '#9a6e3a' },
        'url':            { color: '#9a6e3a' },
        'atrule':         { color: '#07a' },
        'attr-value':     { color: '#07a' },
        'keyword':        { color: '#07a' },
        'function':       { color: '#DD4A68' },
        'class-name':     { color: '#DD4A68' },
        'regex':          { color: '#e90' },
        'important':      { color: '#e90', weight: 'bold' },
        'variable':       { color: '#e90' },
    };

    /**
     * Takes a Prism-highlighted HTML string and converts all
     * <span class="token xyz"> into <span style="color:..."> so that
     * Word, Google Docs, and any rich-text editor can render the colors.
     */
    function inlineTokenColors(html, isDark) {
        const colorMap = isDark ? PRISM_DARK_COLORS : PRISM_LIGHT_COLORS;
        const defaultColor = isDark ? '#ccc' : '#000';

        // Parse the HTML string into a temporary element
        const tmp = document.createElement('div');
        tmp.innerHTML = html;

        tmp.querySelectorAll('span.token').forEach(span => {
            // A span can have multiple token classes e.g. "token keyword operator"
            const classes = Array.from(span.classList).filter(c => c !== 'token');
            let resolved = null;
            for (const cls of classes) {
                if (colorMap[cls]) { resolved = colorMap[cls]; break; }
            }

            const color  = resolved ? resolved.color  : defaultColor;
            const weight = resolved && resolved.weight ? resolved.weight : 'normal';
            const fstyle = resolved && resolved.style  ? resolved.style  : 'normal';

            span.removeAttribute('class');
            span.style.cssText = `color:${color};font-weight:${weight};font-style:${fstyle};font-family:Consolas,'Courier New',monospace;`;
        });

        return tmp.innerHTML;
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Copy for Word — real bordered table with fully inlined syntax colors
    // ─────────────────────────────────────────────────────────────────────────
    copyWordBtn.addEventListener('click', async () => {
        const table = tableOutputWrapper.querySelector('table');
        if (!table) return;

        const isDark    = document.body.classList.contains('theme-dark');
        const WORD_MONO = "Consolas, 'Courier New', monospace";
        const BORDER    = "1px solid #d0d0d0";
        const GUTTER_BG = isDark ? "#1d2021" : "#f6f8fa";
        const GUTTER_FG = isDark ? "#6e7781" : "#8a929a";
        const CODE_BG   = isDark ? "#2d2d2d" : "#ffffff";
        const CODE_FG   = isDark ? "#cccccc" : "#24292e";

        let rowsHtml = '';
        table.querySelectorAll('tr').forEach(row => {
            const gutter   = row.querySelector('.gutter');
            const codeCell = row.querySelector('.code');

            // Convert Prism class-based colors → inline styles
            const inlinedCode = inlineTokenColors(codeCell.innerHTML, isDark);

            rowsHtml += `
<tr>
  <td style="width:36pt;min-width:36pt;text-align:right;vertical-align:top;padding:1pt 6pt 1pt 4pt;background-color:${GUTTER_BG};color:${GUTTER_FG};font-family:Arial,sans-serif;font-size:10pt;border:${BORDER};white-space:nowrap;">${gutter.textContent}</td>
  <td style="vertical-align:top;padding:1pt 8pt;background-color:${CODE_BG};color:${CODE_FG};font-family:${WORD_MONO};font-size:10pt;border:${BORDER};white-space:pre;">${inlinedCode}</td>
</tr>`;
        });

        const wordHtml = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
<head><meta charset="utf-8"><style>table{border-collapse:collapse;}td{mso-border-alt:solid #d0d0d0 .5pt;}</style></head>
<body>
<table border="1" cellspacing="0" cellpadding="0" style="border-collapse:collapse;border:${BORDER};">
${rowsHtml}
</table>
</body></html>`;

        const plainText = Array.from(table.querySelectorAll('.code')).map(c => c.textContent).join('\n');

        try {
            await navigator.clipboard.write([
                new ClipboardItem({
                    'text/html':  new Blob([wordHtml],  { type: 'text/html' }),
                    'text/plain': new Blob([plainText], { type: 'text/plain' })
                })
            ]);
            showCopyFeedback(copyWordBtn, 'Copied for Word!');
        } catch (err) {
            console.warn('Clipboard API failed:', err);
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
                console.error('Both methods failed:', fb);
                alert('Copy failed. Try Chrome or Edge.');
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
        setTimeout(() => { btn.textContent = original; btn.classList.remove('btn-success'); }, 2000);
    }
});
