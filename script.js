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
            // Attempt to detect language from extension
            const ext = file.name.split('.').pop().toLowerCase();
            const langMap = {
                'js': 'javascript',
                'py': 'python',
                'java': 'java',
                'html': 'html',
                'css': 'css',
                'cpp': 'cpp',
                'go': 'go',
                'rs': 'rust'
            };
            if (langMap[ext]) {
                languageSelect.value = langMap[ext];
            }
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
                
                // Gutter (Line Number)
                const tdGutter = document.createElement('td');
                tdGutter.className = 'gutter';
                tdGutter.textContent = index + 1;
                
                // Code Cell
                const tdCode = document.createElement('td');
                tdCode.className = 'code';
                
                // Highlight code if language is provided
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

        // Clear previous output
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

        // Extract code text only (ignoring gutter)
        const codeCells = table.querySelectorAll('.code');
        const codeText = Array.from(codeCells).map(cell => cell.textContent).join('\n');

        try {
            await navigator.clipboard.writeText(codeText);
            showCopyFeedback(copyCodeBtn, 'Copied Code!');
        } catch (err) {
            console.error('Failed to copy: ', err);
        }
    });

    copyWordBtn.addEventListener('click', () => {
        const table = tableOutputWrapper.querySelector('table');
        if (!table) return;

        // Get computed styles for consistent coloring
        const tableStyle = window.getComputedStyle(table);
        const bgColor = tableStyle.backgroundColor;
        const textColor = tableStyle.color;
        const fontSize = tableStyle.fontSize;
        const fontFamily = "Consolas, 'Courier New', monospace"; // Word-friendly fonts

        // Clone the table
        const clonedTable = table.cloneNode(true);
        clonedTable.setAttribute('width', '100%');
        clonedTable.setAttribute('border', '0');
        clonedTable.setAttribute('cellspacing', '0');
        clonedTable.setAttribute('cellpadding', '5'); // Add some padding for Word
        clonedTable.style.borderCollapse = 'collapse';
        clonedTable.style.backgroundColor = bgColor;

        const originalRows = table.querySelectorAll('tr');
        const clonedRows = clonedTable.querySelectorAll('tr');

        originalRows.forEach((row, rowIndex) => {
            const clonedRow = clonedRows[rowIndex];
            clonedRow.style.backgroundColor = bgColor;

            // Gutter
            const gutter = row.querySelector('.gutter');
            const clonedGutter = clonedRow.querySelector('.gutter');
            const gStyle = window.getComputedStyle(gutter);
            
            clonedGutter.style.backgroundColor = gStyle.backgroundColor;
            clonedGutter.style.color = gStyle.color;
            clonedGutter.style.width = '40pt'; // Word likes pt
            clonedGutter.style.textAlign = 'right';
            clonedGutter.style.borderRight = `1px solid ${gStyle.color}`;
            clonedGutter.style.fontFamily = "Arial, sans-serif";
            clonedGutter.style.fontSize = '10pt';

            // Code
            const codeCell = row.querySelector('.code');
            const clonedCodeCell = clonedRow.querySelector('.code');
            const cStyle = window.getComputedStyle(codeCell);
            
            clonedCodeCell.style.backgroundColor = bgColor;
            clonedCodeCell.style.color = cStyle.color;
            clonedCodeCell.style.fontFamily = fontFamily;
            clonedCodeCell.style.fontSize = '10pt';
            clonedCodeCell.style.whiteSpace = 'pre';

            // Convert leading spaces to &nbsp; for Word compatibility
            let html = clonedCodeCell.innerHTML;
            // This is a bit tricky since we have spans. Let's do a simple replace on the text parts.
            // For simplicity, we'll wrap the whole cell content in a <pre> if it's not already
            clonedCodeCell.innerHTML = `<pre style="margin:0; font-family:${fontFamily}; color:${cStyle.color};">${html}</pre>`;

            // Fix spans
            const originalSpans = codeCell.querySelectorAll('span');
            const clonedSpans = clonedCodeCell.querySelectorAll('span');
            originalSpans.forEach((span, spanIndex) => {
                const sStyle = window.getComputedStyle(span);
                clonedSpans[spanIndex].style.color = sStyle.color;
                clonedSpans[spanIndex].style.fontWeight = sStyle.fontWeight;
            });
        });

        // Use a temporary container for selection
        const container = document.createElement('div');
        container.style.position = 'fixed';
        container.style.left = '-9999px';
        container.style.top = '0';
        // Force the background color on a wrapper too
        container.style.backgroundColor = bgColor;
        container.appendChild(clonedTable);
        document.body.appendChild(container);

        // Select and Copy
        const range = document.createRange();
        range.selectNodeContents(container);
        const selection = window.getSelection();
        selection.removeAllRanges();
        selection.addRange(range);

        try {
            const successful = document.execCommand('copy');
            if (successful) {
                showCopyFeedback(copyWordBtn, 'Copied for Word!');
            } else {
                throw new Error('execCommand copy failed');
            }
        } catch (err) {
            console.error('Legacy copy failed: ', err);
            alert('Failed to copy. Please select the table manually and copy.');
        }

        // Cleanup
        selection.removeAllRanges();
        document.body.removeChild(container);
    });

    copyHtmlBtn.addEventListener('click', async () => {
        const table = tableOutputWrapper.querySelector('table');
        if (!table) return;

        // Clone the table to inject inline styles for better portability
        const clonedTable = table.cloneNode(true);
        // Add a class or inline styles if needed for the target environment
        
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
