class ProDocEditor {
    constructor() {
        // DOM Elements
        this.editor = document.getElementById('editor');
        this.docTitle = document.getElementById('docTitle');
        this.sidebar = document.getElementById('sidebar');
        this.outlineList = document.getElementById('outlineList');
        this.wordCount = document.getElementById('wordCount');
        this.charCount = document.getElementById('charCount');
        this.paraCount = document.getElementById('paraCount');
        this.readTime = document.getElementById('readTime');
        this.cursorPos = document.getElementById('cursorPos');
        this.zoomLevel = document.getElementById('zoomLevel');
        this.saveStatus = document.getElementById('saveStatus');
        this.page = document.querySelector('.page');
        this.editorArea = document.getElementById('editorArea');
        
        // State
        this.currentZoom = 100;
        this.autoSaveTimer = null;
        this.history = [];
        this.historyIndex = -1;
        this.isDirty = false;
        this.savedRange = null; // Crucial for modal insertions
        this.selectedTableSize = { rows: 1, cols: 1 };

        this.document = {
            title: '无标题文档',
            content: this.editor.innerHTML,
            created: Date.now(),
            modified: Date.now(),
            version: '1.0.0',
            metadata: {}
        };
        
        this.init();
    }
    
    init() {
        this.loadDocument();
        this.setupEventListeners();
        this.setupToolbar();
        this.setupModals();
        this.setupScrolling();
        this.startAutoSave();
        this.updateStats();
        this.generateOutline();
        this.saveToHistory(this.editor.innerHTML); // Save initial state
    }
    
    // --- Selection Management ---
    saveSelection() {
        const selection = window.getSelection();
        if (selection.rangeCount > 0) {
            this.savedRange = selection.getRangeAt(0).cloneRange();
        } else {
            this.savedRange = null;
        }
    }

    restoreSelection() {
        if (this.savedRange) {
            this.editor.focus();
            const selection = window.getSelection();
            selection.removeAllRanges();
            selection.addRange(this.savedRange);
        } else {
            this.editor.focus();
            const range = document.createRange();
            range.selectNodeContents(this.editor);
            range.collapse(false);
            const selection = window.getSelection();
            selection.removeAllRanges();
            selection.addRange(range);
        }
    }
    
    // --- Document Lifecycle ---
    loadDocument() {
        const saved = localStorage.getItem('prodoc_document');
        if (saved) {
            try {
                const docData = JSON.parse(saved);
                this.document = docData;
                this.docTitle.textContent = docData.title;
                this.editor.innerHTML = docData.content;
                this.showToast('文档已从本地恢复');
            } catch (e) {
                console.error('Failed to load document from localStorage:', e);
                this.showToast('恢复文档失败', 'error');
            }
        }
    }
    
    saveDocument(showToast = false) {
        this.document.title = this.docTitle.textContent.trim() || '无标题文档';
        this.document.content = this.editor.innerHTML;
        this.document.modified = Date.now();
        this.updateStats();
        this.document.metadata = {
            wordCount: parseInt(this.wordCount.textContent, 10),
            charCount: parseInt(this.charCount.textContent, 10),
            paraCount: parseInt(this.paraCount.textContent, 10)
        };
        
        localStorage.setItem('prodoc_document', JSON.stringify(this.document));
        this.updateSaveStatus('已保存');
        if (showToast) {
            this.showToast('文档已手动保存');
        }
        this.isDirty = false;
    }
    
    scheduleAutoSave() {
        clearTimeout(this.autoSaveTimer);
        this.updateSaveStatus('正在输入...');
        this.autoSaveTimer = setTimeout(() => {
            if (this.isDirty) {
                this.saveDocument();
            }
        }, 2500);
    }

    startAutoSave() {
        this.autoSaveTimer = setInterval(() => {
            if (this.isDirty) {
                this.saveDocument();
            }
        }, 15000);
    }
    
    // --- Event Setup ---
    setupEventListeners() {
        document.getElementById('menuToggle').addEventListener('click', () => this.sidebar.classList.toggle('active'));
        document.getElementById('closeSidebar').addEventListener('click', () => this.sidebar.classList.remove('active'));
        document.getElementById('saveBtn').addEventListener('click', () => this.saveDocument(true));
        document.getElementById('importBtn').addEventListener('click', () => this.importDocument());
        document.getElementById('exportBtn').addEventListener('click', () => this.showExportModal());
        document.getElementById('zoomIn').addEventListener('click', () => this.setZoom(Math.min(this.currentZoom + 10, 200)));
        document.getElementById('zoomOut').addEventListener('click', () => this.setZoom(Math.max(this.currentZoom - 10, 50)));

        this.editor.addEventListener('input', () => {
            this.isDirty = true;
            this.scheduleAutoSave();
        });
        
        let updateTimeout;
        this.editor.addEventListener('input', () => {
            clearTimeout(updateTimeout);
            updateTimeout = setTimeout(() => {
                this.updateStats();
                this.generateOutline();
            }, 300);
        });

        this.editor.addEventListener('keyup', (e) => {
            if (!e.ctrlKey && !e.metaKey) {
                 this.saveToHistory(this.editor.innerHTML);
            }
            this.updateToolbarState();
        });

        this.editor.addEventListener('mouseup', () => {
            this.saveToHistory(this.editor.innerHTML);
            this.updateToolbarState();
        });

        this.editor.addEventListener('click', (e) => {
             if (e.target.tagName === 'A' && (e.ctrlKey || e.metaKey)) {
                e.preventDefault();
                window.open(e.target.href, '_blank');
            }
        });

        this.docTitle.addEventListener('input', () => {
            this.isDirty = true;
            this.scheduleAutoSave();
        });

        this.editor.addEventListener('keydown', (e) => {
            if (this.handleBlockquoteExit(e)) {
                return;
            }

            if (e.ctrlKey || e.metaKey) {
                switch(e.key.toLowerCase()) {
                    case 's': e.preventDefault(); this.saveDocument(true); break;
                    case 'z': e.preventDefault(); e.shiftKey ? this.redo() : this.undo(); break;
                    case 'y': e.preventDefault(); this.redo(); break;
                    case 'b': e.preventDefault(); this.formatText('bold'); break;
                    case 'i': e.preventDefault(); this.formatText('italic'); break;
                    case 'u': e.preventDefault(); this.formatText('underline'); break;
                }
            }
        });
    }

    // --- FINAL, CORRECTED BLOCKQUOTE HANDLER ---
    handleBlockquoteExit(e) {
        if (e.key !== 'Enter' && e.key !== 'Backspace') {
            return false;
        }

        const selection = window.getSelection();
        if (!selection || !selection.isCollapsed || selection.rangeCount === 0) {
            return false;
        }

        const range = selection.getRangeAt(0);
        const currentElement = range.startContainer;
        
        const blockquote = (currentElement.nodeType === Node.ELEMENT_NODE ? currentElement : currentElement.parentElement).closest('blockquote');
        if (!blockquote) {
            return false;
        }

        const currentBlock = (currentElement.nodeType === Node.ELEMENT_NODE ? currentElement : currentElement.parentElement).closest('p, h1, h2, h3, li');

        // Handle "Enter" on an empty line to exit the blockquote
        if (e.key === 'Enter' && currentBlock && (currentBlock.textContent.trim() === '' || currentBlock.innerHTML === '<br>')) {
            e.preventDefault();
            
            // This time, we use the reliable 'outdent' command which correctly turns the blockquote line into a normal paragraph.
            document.execCommand('outdent', false, null);
            this.updateToolbarState();
            return true;
        }
        
        // Handle "Backspace" at the very beginning of any line inside a blockquote
        if (e.key === 'Backspace' && range.startOffset === 0) {
             const preCaretRange = range.cloneRange();
             preCaretRange.selectNodeContents(currentBlock || currentElement);
             preCaretRange.setEnd(range.startContainer, range.startOffset);
             if (preCaretRange.toString() === '') {
                 e.preventDefault();
                 document.execCommand('outdent', false, null);
                 this.updateToolbarState();
                 return true;
             }
        }
        
        return false;
    }

    setupToolbar() {
        document.querySelectorAll('.tool-btn[data-action]').forEach(btn => {
            btn.addEventListener('click', () => {
                this.formatText(btn.dataset.action);
            });
        });

        document.getElementById('fontFamily').addEventListener('change', (e) => this.formatText('fontName', e.target.value));
        document.getElementById('fontSize').addEventListener('change', (e) => this.formatText('fontSize', '7', e.target.value));
        document.getElementById('textColor').addEventListener('input', (e) => this.formatText('foreColor', e.target.value));
        document.getElementById('bgColor').addEventListener('input', (e) => this.formatText('backColor', e.target.value));
        
        document.getElementById('insertLink').addEventListener('click', () => this.showLinkModal());
        document.getElementById('insertImage').addEventListener('click', () => this.showImageModal());
        document.getElementById('insertVideo').addEventListener('click', () => this.showVideoModal());
        document.getElementById('insertTable').addEventListener('click', () => this.showTableModal());
        document.getElementById('insertCode').addEventListener('click', () => this.insertCodeBlock());
        document.getElementById('insertQuote').addEventListener('click', () => this.insertQuote());
    }

    setupModals() {
        document.querySelectorAll('.modal').forEach(modal => {
            modal.addEventListener('click', (e) => {
                if (e.target.classList.contains('modal-backdrop') || e.target.closest('.modal-close')) {
                    modal.classList.remove('active');
                }
            });
        });
        
        document.getElementById('confirmLink').addEventListener('click', () => this.confirmLink());
        document.getElementById('confirmImage').addEventListener('click', () => this.confirmImage());
        document.getElementById('confirmVideo').addEventListener('click', () => this.confirmVideo());
        document.getElementById('confirmTable').addEventListener('click', () => this.confirmTable());
        
        document.querySelectorAll('.export-option').forEach(option => {
            option.addEventListener('click', () => {
                this.exportDocument(option.dataset.format);
                document.getElementById('exportModal').classList.remove('active');
            });
        });
    }

    setupScrolling() {
        this.editorArea.addEventListener('wheel', (e) => {
            if (e.ctrlKey) {
                e.preventDefault();
                const newZoom = this.currentZoom + (e.deltaY < 0 ? 10 : -10);
                this.setZoom(Math.max(50, Math.min(200, newZoom)));
            }
        }, { passive: false });
    }
    
    // --- UI & State Updates ---
    updateStats() {
        const text = this.editor.innerText;
        const words = text.match(/\S+/g)?.length || 0;
        const chars = text.length;
        const paragraphs = this.editor.querySelectorAll('p, h1, h2, h3, h4, h5, h6, li, blockquote').length || 1;
        const readTime = Math.max(1, Math.ceil(words / 200));
        
        this.wordCount.textContent = words;
        this.charCount.textContent = chars;
        this.paraCount.textContent = paragraphs;
        this.readTime.textContent = `${readTime}分钟`;
    }

    generateOutline() {
        const headings = this.editor.querySelectorAll('h1, h2, h3');
        this.outlineList.innerHTML = '';
        
        if (headings.length === 0) {
            this.outlineList.innerHTML = '<div class="outline-item">使用标题格式生成大纲</div>';
            return;
        }
        
        headings.forEach(heading => {
            if (!heading.textContent.trim()) return;
            const item = document.createElement('div');
            item.className = 'outline-item';
            item.style.paddingLeft = `${(parseInt(heading.tagName[1]) - 1) * 16}px`;
            item.textContent = heading.textContent;
            item.addEventListener('click', () => heading.scrollIntoView({ behavior: 'smooth', block: 'center' }));
            this.outlineList.appendChild(item);
        });
    }

    updateToolbarState() {
        const commands = ['bold', 'italic', 'underline', 'strikethrough', 'insertOrderedList', 'insertUnorderedList', 'justifyLeft', 'justifyCenter', 'justifyRight', 'justifyFull'];
        commands.forEach(command => {
            const btn = document.querySelector(`.tool-btn[data-action="${command}"]`);
            if (btn) {
                if (document.queryCommandState(command)) {
                    btn.classList.add('active');
                } else {
                    btn.classList.remove('active');
                }
            }
        });
        
        const quoteBtn = document.getElementById('insertQuote');
        const selection = window.getSelection();
        if (selection.rangeCount > 0) {
            const node = selection.getRangeAt(0).startContainer;
            const parentBlockquote = (node.nodeType === 3 ? node.parentNode : node).closest('blockquote');
            quoteBtn.classList.toggle('active', !!parentBlockquote);
        }
    }
    
    updateSaveStatus(status) {
        this.saveStatus.textContent = status;
        this.saveStatus.style.color = status === '已保存' ? 'var(--success)' : 'var(--text-muted)';
    }

    // --- Rich Text Formatting ---
    formatText(command, value = null, arg = null) {
        this.editor.focus();
        document.execCommand(command, false, value);
        
        if (command === 'fontSize' && arg) {
             const fontElements = document.getElementsByTagName("font");
             for (let i = 0, len = fontElements.length; i < len; ++i) {
                if (fontElements[i].size == "7") {
                    fontElements[i].removeAttribute("size");
                    fontElements[i].style.fontSize = arg + "px";
                }
            }
        }
        
        this.updateToolbarState();
        this.isDirty = true;
        this.saveToHistory(this.editor.innerHTML);
    }

    // --- Content Insertion ---
    showLinkModal() {
        this.saveSelection();
        const selection = window.getSelection();
        const parentAnchor = selection.focusNode?.parentElement.closest('a');

        if(parentAnchor) {
            document.getElementById('linkText').value = parentAnchor.innerText;
            document.getElementById('linkUrl').value = parentAnchor.href;
            document.getElementById('linkNewTab').checked = parentAnchor.target === '_blank';
        } else {
            document.getElementById('linkText').value = selection.toString();
            document.getElementById('linkUrl').value = 'https://';
            document.getElementById('linkNewTab').checked = true;
        }
        document.getElementById('linkModal').classList.add('active');
    }

    confirmLink() {
        this.restoreSelection();
        const url = document.getElementById('linkUrl').value;
        if (!url || !url.match(/^(https?|ftp):\/\//i)) {
            this.showToast('请输入有效的链接地址', 'warning');
            return;
        }
        
        const text = document.getElementById('linkText').value || url;

        const selection = window.getSelection();
        const parentAnchor = selection.focusNode?.parentElement.closest('a');
        if(parentAnchor) {
            parentAnchor.href = url;
            parentAnchor.innerText = text;
            if (document.getElementById('linkNewTab').checked) {
                parentAnchor.target = '_blank';
            } else {
                parentAnchor.removeAttribute('target');
            }
        } else {
            const newTab = document.getElementById('linkNewTab').checked;
            const html = `<a href="${url}" ${newTab ? 'target="_blank"' : ''}>${text}</a>`;
            document.execCommand('insertHTML', false, html);
        }

        document.getElementById('linkModal').classList.remove('active');
        this.showToast('链接已更新');
        this.saveToHistory(this.editor.innerHTML);
    }
    
    showImageModal() {
        this.saveSelection();
        document.getElementById('imageModal').classList.add('active');
    }
    
    confirmImage() {
        const fileInput = document.getElementById('imageFile');
        const urlInput = document.getElementById('imageUrl').value;
        
        const processAndInsert = (src) => {
            const width = document.getElementById('imageWidth').value;
            const height = document.getElementById('imageHeight').value;
            this.insertImageElement(src, width, height);
            document.getElementById('imageModal').classList.remove('active');
            fileInput.value = '';
            document.getElementById('imageUrl').value = '';
        };

        if(urlInput && urlInput.startsWith('http')) {
            processAndInsert(urlInput);
        } else if (fileInput.files[0]) {
             const reader = new FileReader();
             reader.onload = (e) => processAndInsert(e.target.result);
             reader.readAsDataURL(fileInput.files[0]);
        } else {
            this.showToast('请选择图片文件或输入有效的图片URL', 'warning');
            return;
        }
    }
    
    insertImageElement(src, width, height) {
        this.restoreSelection();
        let style = `max-width: 100%; width: ${width}px; height: auto;`;
        if (height) style = `max-width: 100%; width: ${width}px; height: ${height}px;`;
        if (document.getElementById('imageRounded').checked) style += 'border-radius: var(--radius-lg);';
        
        let className = '';
        if (document.getElementById('imageShadow').checked) className += ' image-shadow';
        if (document.getElementById('imageBorder').checked) className += ' image-border';

        const html = `<p style="text-align: center;"><img src="${src}" style="${style}" class="${className.trim()}"></p>`;
        document.execCommand('insertHTML', false, html);

        this.showToast('图片已插入');
        this.saveToHistory(this.editor.innerHTML);
    }

    showVideoModal() {
        this.saveSelection();
        document.getElementById('videoModal').classList.add('active');
    }

    confirmVideo() {
        const url = document.getElementById('videoUrl').value;
        const width = document.getElementById('videoWidth').value;
        const height = document.getElementById('videoHeight').value;
        
        if (!url || !url.startsWith('http')) {
            this.showToast('请输入有效的视频链接', 'warning');
            return;
        }

        let html;
        if (url.includes("youtube.com/watch?v=")) {
            const videoId = url.split('v=')[1].split('&')[0];
            html = `<iframe src="https://www.youtube.com/embed/${videoId}" width="${width}" height="${height}" style="border:none; border-radius: var(--radius);" allowfullscreen></iframe>`;
        } else if (url.includes("bilibili.com/video/")) {
            const bvid = url.match(/BV[a-zA-Z0-9]+/);
            if(bvid) html = `<iframe src="//player.bilibili.com/player.html?bvid=${bvid[0]}&page=1" width="${width}" height="${height}" style="border:none; border-radius: var(--radius);" allowfullscreen></iframe>`;
        } else if (url.match(/\.(mp4|webm|ogg)$/i)) {
            html = `<video src="${url}" width="${width}" height="${height}" controls style="border-radius: var(--radius);"></video>`;
        } else {
            this.showToast('不支持的视频链接格式', 'warning');
            return;
        }
        
        this.restoreSelection();
        document.execCommand('insertHTML', false, `<p style="text-align: center;">${html}</p>`);
        
        document.getElementById('videoModal').classList.remove('active');
        this.showToast('视频已插入');
        this.saveToHistory(this.editor.innerHTML);
    }
    
    showTableModal() {
        this.saveSelection();
        const grid = document.getElementById('sizeGrid');
        grid.innerHTML = '';
        const modal = document.getElementById('tableModal');
        for (let i = 0; i < 100; i++) {
            const cell = document.createElement('div');
            cell.className = 'grid-cell';
            const row = Math.floor(i / 10) + 1;
            const col = (i % 10) + 1;
            cell.dataset.row = row;
            cell.dataset.col = col;
            cell.addEventListener('mouseenter', () => this.updateTablePreview(row, col));
            cell.addEventListener('click', () => {
                this.insertTable(row, col);
                modal.classList.remove('active');
            });
            grid.appendChild(cell);
        }
        this.updateTablePreview(1,1);
        modal.classList.add('active');
    }

    updateTablePreview(rows, cols) {
        this.selectedTableSize = { rows, cols };
        document.getElementById('sizePreview').textContent = `${rows} × ${cols} 表格`;
        document.querySelectorAll('#sizeGrid .grid-cell').forEach(cell => {
            if (cell.dataset.row <= rows && cell.dataset.col <= cols) {
                cell.classList.add('selected');
            } else {
                cell.classList.remove('selected');
            }
        });
    }

    confirmTable() {
        const { rows, cols } = this.selectedTableSize;
        if (rows > 0 && cols > 0) {
            this.insertTable(rows, cols);
        }
        document.getElementById('tableModal').classList.remove('active');
    }

    insertTable(rows, cols) {
        this.restoreSelection();
        let tableHTML = '<table style="width: 100%; border-collapse: collapse;"><tbody>';
        for (let i = 0; i < rows; i++) {
            tableHTML += '<tr>';
            for (let j = 0; j < cols; j++) {
                const cellType = (i === 0) ? 'th' : 'td';
                let style = 'border: 1px solid #dee2e6; padding: 8px; min-width: 50px;';
                if (i === 0) style += 'background-color: #f8f9fa; font-weight: 600; text-align: center;';
                tableHTML += `<${cellType} style="${style}"><p><br></p></${cellType}>`;
            }
            tableHTML += '</tr>';
        }
        tableHTML += '</tbody></table><p><br></p>';
        document.execCommand('insertHTML', false, tableHTML);
        this.showToast('表格已插入');
        this.saveToHistory(this.editor.innerHTML);
    }

    insertCodeBlock() {
        const html = `<pre style="background: var(--bg-tertiary); padding: 16px; border-radius: var(--radius); overflow: auto; margin: 16px 0;"><code>在此输入代码...</code></pre><p><br></p>`;
        document.execCommand('insertHTML', false, html);
        this.showToast('代码块已插入');
    }
    
    insertQuote() {
        this.formatText('formatBlock', 'blockquote');
        this.showToast('引用格式已应用');
    }
    
    // --- History (Undo/Redo) ---
    saveToHistory(content) {
        if (this.history[this.historyIndex] === content) return;
        if (this.historyIndex < this.history.length - 1) {
            this.history = this.history.slice(0, this.historyIndex + 1);
        }
        this.history.push(content);
        this.historyIndex++;
        if (this.history.length > 50) {
            this.history.shift();
            this.historyIndex--;
        }
    }
    
    undo() {
        if (this.historyIndex > 0) {
            this.historyIndex--;
            this.editor.innerHTML = this.history[this.historyIndex];
            this.updateAfterHistory();
            this.showToast('已撤销');
        }
    }
    
    redo() {
        if (this.historyIndex < this.history.length - 1) {
            this.historyIndex++;
            this.editor.innerHTML = this.history[this.historyIndex];
            this.updateAfterHistory();
            this.showToast('已重做');
        }
    }

    updateAfterHistory() {
        this.updateStats();
        this.generateOutline();
        this.updateToolbarState();
        this.isDirty = true;
    }

    // --- Import/Export ---
    showExportModal() {
        document.getElementById('exportModal').classList.add('active');
    }

    exportDocument(format) {
        const title = this.docTitle.textContent.trim() || '文档';
        switch(format) {
            case 'docx': this.exportWord(title); break;
            case 'pdf': this.exportPDF(title); break;
            case 'html': this.exportHTML(title); break;
            case 'txt': this.exportText(title); break;
            case 'md': this.exportMarkdown(title, this.editor.innerHTML); break;
            case 'json': this.exportJSON(title); break;
            default: this.showToast('该导出格式暂未实现', 'warning');
        }
    }

    async exportWord(title) {
        this.showToast('正在创建 Word 文档...', 'info');
        const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType } = docx;

        const parseNodeToDocx = (node) => {
            if (node.nodeType === Node.TEXT_NODE) {
                return [new TextRun(node.textContent)];
            }
            if (node.nodeType !== Node.ELEMENT_NODE) {
                return [];
            }

            let runs = [];
            node.childNodes.forEach(childNode => {
                let childRuns = [];
                if (childNode.nodeType === Node.TEXT_NODE) {
                    childRuns.push(new TextRun(childNode.textContent));
                } else if (childNode.nodeType === Node.ELEMENT_NODE) {
                    childRuns = parseNodeToDocx(childNode);
                }
                
                childRuns.forEach(run => {
                    let currentRun = run;
                    if (childNode.parentElement.nodeName === 'STRONG' || childNode.parentElement.nodeName === 'B') currentRun = currentRun.bold();
                    if (childNode.parentElement.nodeName === 'EM' || childNode.parentElement.nodeName === 'I') currentRun = currentRun.italic();
                    if (childNode.parentElement.nodeName === 'U') currentRun = currentRun.underline();
                    if (childNode.parentElement.nodeName === 'S') currentRun = currentRun.strike();
                    runs.push(currentRun);
                });
            });

            const paragraphOptions = {};
            switch (node.nodeName) {
                case 'H1': paragraphOptions.heading = HeadingLevel.HEADING_1; break;
                case 'H2': paragraphOptions.heading = HeadingLevel.HEADING_2; break;
                case 'H3': paragraphOptions.heading = HeadingLevel.HEADING_3; break;
                case 'P': break;
                case 'BLOCKQUOTE': paragraphOptions.style = "IntenseQuote"; break;
                default: return runs;
            }

            if (node.style.textAlign) {
                 switch (node.style.textAlign) {
                    case 'center': paragraphOptions.alignment = AlignmentType.CENTER; break;
                    case 'right': paragraphOptions.alignment = AlignmentType.RIGHT; break;
                    case 'justify': paragraphOptions.alignment = AlignmentType.JUSTIFY; break;
                 }
            }
            
            if (runs.length > 0 || node.innerHTML === '<br>') {
                 return [new Paragraph({ ...paragraphOptions, children: runs })];
            }
            return [];
        };
        
        const docxChildren = Array.from(this.editor.childNodes).flatMap(parseNodeToDocx);

        const doc = new Document({
            sections: [{
                children: [
                    new Paragraph({ text: title, heading: HeadingLevel.TITLE }),
                    new Paragraph(""),
                    ...docxChildren
                ],
            }],
        });

        try {
            const blob = await Packer.toBlob(doc);
            this.downloadFile(blob, `${title}.docx`, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
            this.showToast('Word 文档已导出');
        } catch (err) {
            console.error(err);
            this.showToast('Word 导出失败', 'error');
        }
    }

    async exportPDF(title) {
        this.showToast('正在创建 PDF 文档...', 'info');
        const { jsPDF } = window.jspdf;
        const editorPage = this.page;
        const originalTransform = editorPage.style.transform;
        
        editorPage.style.transform = 'scale(1)';
        
        try {
            const canvas = await html2canvas(editorPage, { scale: 2, useCORS: true });
            
            editorPage.style.transform = originalTransform;

            const imgData = canvas.toDataURL('image/png');
            const pdf = new jsPDF('p', 'mm', 'a4');
            const pdfWidth = pdf.internal.pageSize.getWidth();
            const pdfHeight = pdf.internal.pageSize.getHeight();
            const imgProps = pdf.getImageProperties(imgData);
            const imgRatio = imgProps.width / imgProps.height;
            const canvasPdfHeight = pdfWidth / imgRatio;
            
            let heightLeft = canvasPdfHeight;
            let position = 0;

            pdf.addImage(imgData, 'PNG', 0, position, pdfWidth, canvasPdfHeight);
            heightLeft -= pdfHeight;

            while (heightLeft > 0) {
                position = -heightLeft;
                pdf.addPage();
                pdf.addImage(imgData, 'PNG', 0, position, pdfWidth, canvasPdfHeight);
                heightLeft -= pdfHeight;
            }

            pdf.save(`${title}.pdf`);
            this.showToast('PDF 文档已导出');

        } catch(err) {
            console.error('PDF export failed:', err);
            this.showToast('PDF 导出失败', 'error');
            editorPage.style.transform = originalTransform;
        }
    }

    exportHTML(title) {
        const cssContent = Array.from(document.styleSheets).map(sheet => {
            try { return Array.from(sheet.cssRules).map(rule => rule.cssText).join(''); }
            catch (e) { console.warn("Cannot read stylesheet " + sheet.href); }
        }).filter(Boolean).join('\n');

        const content = this.editor.innerHTML;
        const html = `<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"><title>${title}</title><style>${cssContent}</style></head><body style="background:#fff;padding:40px;"><div class="page">${content}</div></body></html>`;
        this.downloadFile(html, `${title}.html`, 'text/html');
        this.showToast('HTML 文档已导出');
    }
    
    exportText(title) {
        const text = this.editor.innerText;
        this.downloadFile(text, `${title}.txt`, 'text/plain;charset=utf-8');
        this.showToast('文本文档已导出');
    }

    exportMarkdown(title, content) {
        let markdown = content;
        
        markdown = markdown.replace(/<h1[^>]*>(.*?)<\/h1>/gi, '# $1\n\n');
        markdown = markdown.replace(/<h2[^>]*>(.*?)<\/h2>/gi, '## $1\n\n');
        markdown = markdown.replace(/<h3[^>]*>(.*?)<\/h3>/gi, '### $1\n\n');
        markdown = markdown.replace(/<strong>(.*?)<\/strong>/gi, '**$1**').replace(/<b>(.*?)<\/b>/gi, '**$1**');
        markdown = markdown.replace(/<em>(.*?)<\/em>/gi, '*$1*').replace(/<i>(.*?)<\/i>/gi, '*$1*');
        markdown = markdown.replace(/<ul[^>]*>([\s\S]*?)<\/ul>/gi, (match, m1) => m1.replace(/<li[^>]*>(.*?)<\/li>/gi, '- $1\n') + '\n');
        markdown = markdown.replace(/<ol[^>]*>([\s\S]*?)<\/ol>/gi, (match, m1) => {
            let i = 1;
            return m1.replace(/<li[^>]*>(.*?)<\/li>/gi, (m, m1_li) => `${i++}. ${m1_li}\n`) + '\n';
        });
        markdown = markdown.replace(/<blockquote[^>]*>(.*?)<\/blockquote>/gi, (m, m1) => `> ${m1.replace(/<p[^>]*>(.*?)<\/p>/gi, '$1\n')}\n`);
        markdown = markdown.replace(/<a href="([^"]*)"[^>]*>(.*?)<\/a>/gi, '[$2]($1)');
        markdown = markdown.replace(/<img src="([^"]*)"[^>]*>/gi, '![Image]($1)');
        markdown = markdown.replace(/<pre[^>]*><code[^>]*>(.*?)<\/code><\/pre>/gi, '```\n$1\n```\n\n');
        markdown = markdown.replace(/<p[^>]*>(.*?)<\/p>/gi, '$1\n\n');
        markdown = markdown.replace(/<br\s*\/?>/gi, '\n');
        
        const tempDiv = document.createElement('div');
        tempDiv.innerHTML = markdown;
        markdown = tempDiv.textContent || tempDiv.innerText || '';

        this.downloadFile(markdown.trim(), `${title}.md`, 'text/markdown;charset=utf-8');
        this.showToast('Markdown 文档已导出');
    }

    exportJSON(title) {
        const data = {
            ...this.document,
            title: title,
            content: this.editor.innerHTML,
            plainText: this.editor.innerText,
            exportedAt: Date.now(),
            exportFormat: 'json'
        };
        const jsonString = JSON.stringify(data, null, 2);
        this.downloadFile(jsonString, `${title}.json`, 'application/json;charset=utf-8');
        this.showToast('JSON 文档已导出');
    }

    importDocument() {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.json,.txt,.md,.html';
        input.style.display = 'none';

        input.onchange = (e) => {
            const file = e.target.files[0];
            if (!file) return;

            const reader = new FileReader();

            reader.onload = (event) => {
                const content = event.target.result;
                const fileName = file.name;

                try {
                    if (fileName.endsWith('.json')) {
                        const data = JSON.parse(content);
                        if (data.title && data.content) {
                            this.docTitle.textContent = data.title;
                            this.editor.innerHTML = data.content;
                            this.document = { ...this.document, ...data };
                            this.showToast('JSON 文档已成功导入');
                        } else {
                            this.showToast('导入失败：无效的 ProDoc JSON 格式', 'error');
                            return;
                        }
                    } else { 
                        this.docTitle.textContent = fileName.split('.').slice(0, -1).join('.') || '导入的文档';
                        this.editor.innerHTML = `<p>${content.replace(/\n\n/g, '</p><p>').replace(/\n/g, '<br>')}</p>`;
                        this.showToast(`已导入 ${fileName}`);
                    }

                    this.updateAfterHistory();
                    this.saveToHistory(this.editor.innerHTML);
                    this.saveDocument();
                } catch (err) {
                    console.error("Import failed:", err);
                    this.showToast('文件导入或解析失败', 'error');
                }
            };

            reader.onerror = () => {
                this.showToast('读取文件时出错', 'error');
            };

            reader.readAsText(file);
        };

        document.body.appendChild(input);
        input.click();
        document.body.removeChild(input);
    }
    
    downloadFile(content, filename, mimeType) {
        const blob = new Blob([content], { type: mimeType });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    // --- Utilities ---
    setZoom(level) {
        this.currentZoom = level;
        this.page.style.transform = `scale(${level / 100})`;
        this.zoomLevel.textContent = `${level}%`;
    }

    showToast(message, type = 'success') {
        const toast = document.getElementById('toast');
        toast.textContent = message;
        toast.className = `toast show`;
        const colors = { success: 'var(--success)', warning: 'var(--warning)', error: 'var(--danger)', info: 'var(--info)' };
        toast.style.background = colors[type] || 'var(--text-primary)';
        setTimeout(() => toast.classList.remove('show'), 3000);
    }
}

document.addEventListener('DOMContentLoaded', () => {
    window.proDocEditor = new ProDocEditor();
    window.addEventListener('beforeunload', (e) => {
        if (window.proDocEditor.isDirty) {
            e.preventDefault();
            e.returnValue = '您有未保存的更改，确定要离开吗？';
        }
    });
});