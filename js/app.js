/**
 * Main Application Logic
 * Handles UI interactions and orchestrates the conversion process
 */

// Messages for the parse failures DocxParser tells apart with error.code
const PARSE_FAILURE_KEYS = {
    corrupt: 'error.corrupt',
    encrypted: 'error.encrypted',
    legacyDoc: 'error.legacyDoc',
    strict: 'error.strict'
};

// Messages for the conversion findings DocxParser reports, by finding code
const FINDING_KEYS = {
    headersFooters: 'notes.headersFooters',
    comments: 'notes.comments',
    nestedTables: 'notes.nestedTables',
    missingImage: 'notes.missingImage',
    unsupportedImage: 'notes.unsupportedImage'
};

// Findings whose examples are document text, shown in quotes; the others are format names
const QUOTED_EXAMPLE_CODES = ['nestedTables', 'missingImage'];

class WordToUdfApp {
    constructor() {
        // DOM Elements
        this.uploadZone = document.getElementById('uploadZone');
        this.fileInput = document.getElementById('fileInput');
        this.fileInfo = document.getElementById('fileInfo');
        this.fileName = document.getElementById('fileName');
        this.fileSize = document.getElementById('fileSize');
        this.removeFile = document.getElementById('removeFile');
        this.convertBtn = document.getElementById('convertBtn');
        this.progressContainer = document.getElementById('progressContainer');
        this.progressFill = document.getElementById('progressFill');
        this.progressText = document.getElementById('progressText');
        this.downloadSection = document.getElementById('downloadSection');
        this.downloadBtn = document.getElementById('downloadBtn');
        this.convertAnother = document.getElementById('convertAnother');
        this.conversionNotes = document.getElementById('conversionNotes');
        this.conversionNotesList = document.getElementById('conversionNotesList');
        this.errorMessage = document.getElementById('errorMessage');
        this.errorText = document.getElementById('errorText');

        // State
        this.selectedFile = null;
        this.convertedBlob = null;
        this.outputFilename = 'document.udf';

        // Modules
        this.parser = new DocxParser();
        this.generator = new UdfGenerator();

        // Initialize
        this.init();
    }

    init() {
        this.bindEvents();
    }

    bindEvents() {
        // File input change
        this.fileInput.addEventListener('change', (e) => this.handleFileSelect(e));

        // Drag and drop
        this.uploadZone.addEventListener('dragover', (e) => this.handleDragOver(e));
        this.uploadZone.addEventListener('dragleave', (e) => this.handleDragLeave(e));
        this.uploadZone.addEventListener('drop', (e) => this.handleDrop(e));

        // Remove file
        this.removeFile.addEventListener('click', () => this.handleRemoveFile());

        // Convert button
        this.convertBtn.addEventListener('click', () => this.handleConvert());

        // Download button
        this.downloadBtn.addEventListener('click', () => this.handleDownload());

        // Convert another
        this.convertAnother.addEventListener('click', () => this.handleConvertAnother());
    }

    handleDragOver(e) {
        e.preventDefault();
        e.stopPropagation();
        this.uploadZone.classList.add('drag-over');
    }

    handleDragLeave(e) {
        e.preventDefault();
        e.stopPropagation();
        this.uploadZone.classList.remove('drag-over');
    }

    handleDrop(e) {
        e.preventDefault();
        e.stopPropagation();
        this.uploadZone.classList.remove('drag-over');

        const files = e.dataTransfer.files;
        if (files.length > 0) {
            this.processFile(files[0]);
        }
    }

    handleFileSelect(e) {
        const files = e.target.files;
        if (files.length > 0) {
            this.processFile(files[0]);
        }
    }

    processFile(file) {
        // Validate file type
        const extension = getFileExtension(file.name);
        if (extension !== 'docx') {
            this.showError('error.invalidType');
            return;
        }

        this.selectedFile = file;
        this.outputFilename = toUdfFilename(file.name);

        // Update UI
        this.fileName.textContent = file.name;
        this.fileSize.textContent = formatFileSize(file.size);
        this.uploadZone.classList.add('hidden');
        this.fileInfo.classList.remove('hidden');
        this.convertBtn.disabled = false;
        this.hideError();
    }

    handleRemoveFile() {
        this.selectedFile = null;
        this.fileInput.value = '';
        this.fileInfo.classList.add('hidden');
        this.uploadZone.classList.remove('hidden');
        this.convertBtn.disabled = true;
        this.hideError();
    }

    async handleConvert() {
        if (!this.selectedFile) return;

        // The parser reports any failure with an English message (and an error.code for the
        // failures it can name); show the translated message for the step that failed and
        // keep the raw error for the console
        let failureKey = 'error.parse';

        try {
            // Show progress
            this.convertBtn.classList.add('hidden');
            this.progressContainer.classList.remove('hidden');
            this.hideError();
            this.clearNotes();

            // Update progress: Parsing
            this.updateProgress(20, 'progress.parsing');
            await sleep(300);

            // Parse DOCX
            const document = await this.parser.parse(this.selectedFile);

            // Update progress: Converting
            failureKey = 'error.generic';
            this.updateProgress(60, 'progress.converting');
            await sleep(300);

            // Generate UDF
            this.convertedBlob = await this.generator.generate(document);

            // Update progress: Complete
            this.updateProgress(100, 'progress.complete');
            await sleep(500);

            // Show download section, with notes on what the conversion dropped or simplified
            this.progressContainer.classList.add('hidden');
            this.fileInfo.classList.add('hidden');
            this.showNotes(document.findings || []);
            this.downloadSection.classList.remove('hidden');

        } catch (error) {
            console.error('Conversion error:', error);
            // Damaged, password-protected, .doc and Strict Open XML files get a message saying what to do
            if (failureKey === 'error.parse' && error &&
                Object.prototype.hasOwnProperty.call(PARSE_FAILURE_KEYS, error.code)) {
                failureKey = PARSE_FAILURE_KEYS[error.code];
            }
            this.progressContainer.classList.add('hidden');
            this.convertBtn.classList.remove('hidden');
            this.showError(failureKey);
        }
    }

    updateProgress(percent, messageKey) {
        this.progressFill.style.width = `${percent}%`;
        setTranslatedText(this.progressText, messageKey);
    }

    handleDownload() {
        if (this.convertedBlob) {
            downloadBlob(this.convertedBlob, this.outputFilename);
        }
    }

    handleConvertAnother() {
        // Reset state
        this.selectedFile = null;
        this.convertedBlob = null;
        this.fileInput.value = '';

        // Reset UI
        this.downloadSection.classList.add('hidden');
        this.uploadZone.classList.remove('hidden');
        this.convertBtn.classList.remove('hidden');
        this.convertBtn.disabled = true;
        this.progressFill.style.width = '0%';
        this.clearNotes();
        this.hideError();
    }

    /**
     * List the parser's findings under the download button; hidden when there are none.
     * Examples are document text, so they only ever go in through textContent.
     * @param {Array<{code: string, count: number, examples: Array<string>}>} findings
     */
    showNotes(findings) {
        this.clearNotes();
        const known = findings.filter(finding => Object.prototype.hasOwnProperty.call(FINDING_KEYS, finding.code));
        if (known.length === 0) return;

        for (const { code, count, examples } of known) {
            const item = document.createElement('li');
            item.className = 'conversion-note';
            item.dataset.code = code;

            const message = document.createElement('p');
            message.className = 'conversion-note-text';
            setTranslatedText(message, FINDING_KEYS[code]);
            item.appendChild(message);

            if (count > 1) {
                const countLine = document.createElement('p');
                countLine.className = 'conversion-note-count';
                setTranslatedText(countLine, 'notes.count', { count });
                item.appendChild(countLine);
            }

            if (examples.length > 0) {
                const label = document.createElement('p');
                label.className = 'conversion-note-examples-label';
                setTranslatedText(label, 'notes.examples');
                item.appendChild(label);

                const list = document.createElement('ul');
                list.className = 'conversion-note-examples';
                for (const example of examples) {
                    const entry = document.createElement('li');
                    entry.textContent = QUOTED_EXAMPLE_CODES.includes(code) ? `“${example}”` : example;
                    list.appendChild(entry);
                }
                item.appendChild(list);
            }

            this.conversionNotesList.appendChild(item);
        }
        this.conversionNotes.classList.remove('hidden');
    }

    clearNotes() {
        this.conversionNotesList.replaceChildren();
        this.conversionNotes.classList.add('hidden');
    }

    showError(messageKey) {
        setTranslatedText(this.errorText, messageKey);
        this.errorMessage.classList.remove('hidden');
    }

    hideError() {
        this.errorMessage.classList.add('hidden');
    }
}

// Initialize app when DOM is ready
document.addEventListener('DOMContentLoaded', () => {
    window.app = new WordToUdfApp();
});
