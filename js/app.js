/**
 * Main Application Logic
 * Handles UI interactions and orchestrates the conversion process
 */

// Messages for the parse failures DocxParser tells apart with error.code
const PARSE_FAILURE_KEYS = {
    corrupt: 'error.corrupt',
    encrypted: 'error.encrypted',
    legacyDoc: 'error.legacyDoc',
    strict: 'error.strict',
    tooLarge: 'error.tooLarge'
};

// Messages for the conversion findings DocxParser reports, by finding code
const FINDING_KEYS = {
    headersFooters: 'notes.headersFooters',
    comments: 'notes.comments',
    nestedTables: 'notes.nestedTables',
    missingImage: 'notes.missingImage',
    unsupportedImage: 'notes.unsupportedImage',
    notesMoved: 'notes.notesMoved',
    endnotesMoved: 'notes.endnotesMoved',
    lineSpacing: 'notes.lineSpacing',
    tabStops: 'notes.tabStops'
};

// The conversion's real steps, in order; the bar fills one share per step
const PROGRESS_STAGES = ['progress.parsing', 'progress.converting', 'progress.complete'];

// Findings whose examples are document text, shown in quotes; the others are format names
const QUOTED_EXAMPLE_CODES = ['nestedTables', 'missingImage', 'notesMoved', 'endnotesMoved', 'lineSpacing', 'tabStops'];

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
        this.progressBar = document.getElementById('progressBar');
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
        // Bumped by every new conversion and every change of file; a conversion whose
        // number is no longer current belongs to a file that was removed or replaced
        this.conversionRun = 0;

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

        this.cancelConversion();
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

    /**
     * Remove the selected file. During a conversion this cancels it: the parser can't stop
     * mid-way, so its result is discarded when it arrives.
     */
    handleRemoveFile() {
        this.cancelConversion();
        this.selectedFile = null;
        this.fileInput.value = '';
        this.fileInfo.classList.add('hidden');
        this.progressContainer.classList.add('hidden');
        this.resetProgress();
        this.uploadZone.classList.remove('hidden');
        this.convertBtn.classList.remove('hidden');
        this.convertBtn.disabled = true;
        this.hideError();
    }

    /** Make any conversion still running stale, so nothing it produces is shown */
    cancelConversion() {
        this.conversionRun++;
    }

    async handleConvert() {
        const file = this.selectedFile;
        if (!file) return;

        // Tie this run to its file: the name is taken now, and once the file is removed or
        // replaced (or "convert another" is pressed) the run drops whatever it produces
        const run = ++this.conversionRun;
        const isStale = () => run !== this.conversionRun;
        const outputFilename = toUdfFilename(file.name);

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

            // The parser and generator keep state across their awaits, and a cancelled run
            // can't be stopped, so each run gets its own: a discarded run never touches the next
            const parser = this.createParser();
            const generator = this.createGenerator();

            // Read the Word document, once the browser has painted that it is doing so
            await this.showStage('progress.parsing');
            if (isStale()) return;
            const document = await parser.parse(file);
            if (isStale()) return;

            // Build the UDF
            failureKey = 'error.generic';
            await this.showStage('progress.converting');
            if (isStale()) return;
            const blob = await generator.generate(document);
            if (isStale()) return;

            this.convertedBlob = blob;
            this.outputFilename = outputFilename;
            this.updateProgress('progress.complete');

            // Show download section, with notes on what the conversion dropped or simplified,
            // and take keyboard and screen-reader users to the result
            this.progressContainer.classList.add('hidden');
            this.fileInfo.classList.add('hidden');
            this.showNotes(document.findings || []);
            this.downloadSection.classList.remove('hidden');
            this.downloadBtn.focus();

        } catch (error) {
            // A conversion of a file that is gone has nothing to report
            if (isStale()) return;
            console.error('Conversion error:', error);
            // Damaged, password-protected, .doc, Strict Open XML and too-large files get a message saying what to do
            if (failureKey === 'error.parse' && error &&
                Object.prototype.hasOwnProperty.call(PARSE_FAILURE_KEYS, error.code)) {
                failureKey = PARSE_FAILURE_KEYS[error.code];
            }
            this.progressContainer.classList.add('hidden');
            this.convertBtn.classList.remove('hidden');
            this.showError(failureKey);
            // The role="alert" message announces itself; put focus back on the convert button,
            // which lost it when it was hidden, unless the user has moved it elsewhere
            const focused = window.document.activeElement;
            if (!focused || focused === window.document.body) this.convertBtn.focus();
        }
    }

    createParser() {
        return new DocxParser();
    }

    createGenerator() {
        return new UdfGenerator();
    }

    /**
     * Show a stage and give the browser a frame to paint it before the work starts
     * @param {string} messageKey - One of PROGRESS_STAGES
     */
    async showStage(messageKey) {
        this.updateProgress(messageKey);
        await nextFrame();
    }

    /**
     * Fill the bar to a stage's share and show its text
     * @param {string} messageKey - One of PROGRESS_STAGES
     */
    updateProgress(messageKey) {
        const stage = PROGRESS_STAGES.indexOf(messageKey) + 1;
        this.setProgressValue(Math.round(stage / PROGRESS_STAGES.length * 100));
        setTranslatedText(this.progressText, messageKey);
    }

    resetProgress() {
        this.setProgressValue(0);
    }

    setProgressValue(percent) {
        this.progressFill.style.width = `${percent}%`;
        this.progressBar.setAttribute('aria-valuenow', String(percent));
    }

    handleDownload() {
        if (this.convertedBlob) {
            downloadBlob(this.convertedBlob, this.outputFilename);
        }
    }

    handleConvertAnother() {
        // Drop the result, then reset the file, progress and buttons as removing the file does
        this.convertedBlob = null;
        this.downloadSection.classList.add('hidden');
        this.clearNotes();
        this.handleRemoveFile();
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
