/**
 * Main Application Logic
 * Handles UI interactions and orchestrates the conversion process
 */

// Messages for the parse failures DocxParser tells apart with error.code
const PARSE_FAILURE_KEYS = {
    corrupt: 'error.corrupt',
    embeddedOnly: 'error.embeddedOnly',
    encrypted: 'error.encrypted',
    legacyDoc: 'error.legacyDoc',
    strict: 'error.strict',
    tooLarge: 'error.tooLarge'
};

// Messages for the conversion findings DocxParser reports, by finding code
const FINDING_KEYS = {
    headersFooters: 'notes.headersFooters',
    comments: 'notes.comments',
    hiddenText: 'notes.hiddenText',
    nestedTables: 'notes.nestedTables',
    missingImage: 'notes.missingImage',
    unsupportedImage: 'notes.unsupportedImage',
    unconvertedDrawing: 'notes.unconvertedDrawing',
    notesMoved: 'notes.notesMoved',
    endnotesMoved: 'notes.endnotesMoved',
    lineSpacing: 'notes.lineSpacing',
    tabStops: 'notes.tabStops',
    symbols: 'notes.symbols',
    tableBorders: 'notes.tableBorders',
    equations: 'notes.equations',
    embeddedContent: 'notes.embeddedContent',
    trackedChanges: 'notes.trackedChanges',
    sectionLayout: 'notes.sectionLayout',
    subscript: 'notes.subscript',
    doubleStrikethrough: 'notes.doubleStrikethrough',
    underlineStyles: 'notes.underlineStyles',
    listMarkers: 'notes.listMarkers',
    paragraphBorders: 'notes.paragraphBorders',
    paragraphShading: 'notes.paragraphShading',
    hyperlinks: 'notes.hyperlinks'
};

// The conversion's real steps, in order; the bar fills one share per step
const PROGRESS_STAGES = ['progress.parsing', 'progress.converting', 'progress.complete'];

// Findings whose examples are document text, shown in quotes; the others are format names
const QUOTED_EXAMPLE_CODES = ['hiddenText', 'nestedTables', 'missingImage', 'unconvertedDrawing', 'notesMoved', 'endnotesMoved', 'lineSpacing', 'tabStops',
    'equations', 'embeddedContent', 'trackedChanges', 'sectionLayout', 'subscript', 'doubleStrikethrough',
    'underlineStyles', 'listMarkers', 'paragraphBorders', 'paragraphShading', 'hyperlinks'];

// What separates the parts of a location line
const LOCATION_SEPARATOR = ' · ';

/**
 * The translated parts of a finding location's line, before its example: the note it is in,
 * where it is, and the heading it is under. A document-wide location has none.
 * @param {Object} location - See DocxParser.addFinding
 * @returns {Array<[string, Object]>} [dictionary key, params] pairs, in order
 */
function locationParts(location) {
    const parts = [];
    if (location.note && (location.note.kind === 'footnote' || location.note.kind === 'endnote')) {
        parts.push([`notes.location.${location.note.kind}`, {}]);
    }
    if (location.kind === 'paragraph') {
        parts.push(['notes.location.paragraph', { paragraph: location.paragraph }]);
    } else if (location.kind === 'cell') {
        parts.push(['notes.location.cell', { table: location.table, row: location.row, cell: location.cell }]);
    } else if (location.kind === 'header' || location.kind === 'footer') {
        parts.push([`notes.location.${location.kind}`, {}]);
    } else {
        return [];
    }
    if (typeof location.heading === 'string' && location.heading !== '') {
        parts.push(['notes.location.heading', { heading: location.heading }]);
    }
    return parts;
}

class WordToUdfApp {
    constructor() {
        // DOM Elements
        this.uploadCard = document.getElementById('uploadCard');
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
        this.resultStatus = document.getElementById('resultStatus');
        this.convertAnother = document.getElementById('convertAnother');
        this.replaceSourceBtn = document.getElementById('replaceSource');
        this.replaceInput = document.getElementById('replaceInput');
        this.resultSourceFile = document.getElementById('resultSourceFile');
        this.resultConverted = document.getElementById('resultConverted');
        this.resultOutput = document.getElementById('resultOutput');
        this.conversionNotes = document.getElementById('conversionNotes');
        this.conversionNotesList = document.getElementById('conversionNotesList');
        this.conversionNotesHelp = document.getElementById('conversionNotesHelp');
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

        // Replace source: pick a new revision and convert it in place of the current result.
        // The input is emptied first, so choosing a file of the same name still counts as a change
        this.replaceSourceBtn.addEventListener('click', () => {
            this.replaceInput.value = '';
            this.replaceInput.click();
        });
        this.replaceInput.addEventListener('change', (e) => {
            const files = e.target.files;
            if (files.length > 0) this.replaceSource(files[0]);
        });
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
        // The focused file input went with the upload zone; the next step is converting
        this.focusIfLost(this.convertBtn, this.uploadZone);
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
        // Choosing a file is the next step, whichever of the now hidden or disabled controls had focus
        this.focusIfLost(this.fileInput, this.fileInfo, this.progressContainer, this.downloadSection, this.convertBtn);
    }

    /**
     * Move focus to `target` when the focused control is gone: focus is on <body>, or inside one
     * of the parts just hidden. A place the user moved focus to on purpose is left alone.
     * @param {HTMLElement} target - The control for the next step
     * @param {...HTMLElement} hiddenParts - What the step hid or disabled
     */
    focusIfLost(target, ...hiddenParts) {
        const focused = window.document.activeElement;
        if (!focused || focused === window.document.body || hiddenParts.some(part => part.contains(focused))) {
            target.focus();
        }
    }

    /** Make any conversion still running stale, so nothing it produces is shown */
    cancelConversion() {
        this.conversionRun++;
    }

    /**
     * Replace the source of the result on screen with a new file (usually a new revision of
     * the same document) and convert it straight away. The old download, source line and
     * notes go at once, before any work; a file that isn't a .docx leaves them in place.
     * @param {File} file
     * @returns {Promise<void>|undefined} The conversion, when one starts
     */
    replaceSource(file) {
        if (getFileExtension(file.name) !== 'docx') {
            this.showError('error.invalidType');
            return;
        }
        this.clearResult();
        this.processFile(file);
        return this.handleConvert();
    }

    async handleConvert() {
        const file = this.selectedFile;
        if (!file) return;

        // Tie this run to its file: the name is taken now, and once the file is removed or
        // replaced (or "convert another" is pressed) the run drops whatever it produces
        const run = ++this.conversionRun;
        const isStale = () => run !== this.conversionRun;
        const outputFilename = toUdfFilename(file.name);
        // What the result will say about its source, read from the file this run converts
        const source = {
            name: file.name,
            size: formatFileSize(file.size),
            lastModified: Number.isFinite(file.lastModified) ? file.lastModified : null,
            output: outputFilename
        };

        // The parser reports any failure with an English message (and an error.code for the
        // failures it can name); show the translated message for the step that failed and
        // keep the raw error for the console
        let failureKey = 'error.parse';

        try {
            // Show progress. The status region is revealed empty and filled a frame later:
            // screen readers announce changes to a live region, not what a region shows
            // when it first appears
            this.convertBtn.classList.add('hidden');
            delete this.progressText.dataset.i18n;
            this.progressText.textContent = '';
            this.progressContainer.classList.remove('hidden');
            this.hideError();
            this.clearResult();
            // KALDIR / REMOVE stays visible and cancels the run, so it takes the focus the
            // hidden convert button (or replace button, for a replacement) had
            this.focusIfLost(this.removeFile, this.convertBtn, this.downloadSection);
            await nextFrame();
            if (isStale()) return;

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
            // and take keyboard and screen-reader users to the result, unless they moved focus
            // out of the converter while it ran
            this.progressContainer.classList.add('hidden');
            this.fileInfo.classList.add('hidden');
            this.showSource(source, Date.now());
            const noteCount = this.showNotes(document.findings || []);
            this.downloadSection.classList.remove('hidden');
            this.focusIfLost(this.downloadBtn, this.uploadCard);

            // Say once that it finished and how many notes follow. Like the progress status,
            // the region is revealed empty and filled a frame later, so it is announced
            await nextFrame();
            if (isStale()) return;
            this.announceResult(noteCount);

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
            // The role="alert" message announces itself; put focus back on the convert button
            // (from KALDIR / REMOVE, or from a hidden control), unless the user has moved it
            // out of the converter
            this.focusIfLost(this.convertBtn, this.uploadCard);
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
        this.clearResult();
        this.handleRemoveFile();
    }

    /** Drop the result: its download, announcement, source line and notes go together */
    clearResult() {
        this.convertedBlob = null;
        this.downloadSection.classList.add('hidden');
        this.clearAnnouncement();
        this.clearSource();
        this.clearNotes();
    }

    /**
     * Fill the result's status region with one sentence: the conversion finished, and how many
     * notes are listed with it. A language change rewrites it in place.
     * @param {number} noteCount - The notes listed under the download button
     */
    announceResult(noteCount) {
        if (noteCount === 0) {
            setTranslatedText(this.resultStatus, 'result.announce');
        } else if (noteCount === 1) {
            setTranslatedText(this.resultStatus, 'result.announceOneNote');
        } else {
            setTranslatedText(this.resultStatus, 'result.announceNotes', { count: noteCount });
        }
    }

    clearAnnouncement() {
        delete this.resultStatus.dataset.i18n;
        delete this.resultStatus.dataset.i18nParams;
        this.resultStatus.textContent = '';
    }

    /**
     * Say which file and revision the result came from. The name is the user's, so it only
     * ever goes in through textContent; the times are written in the page language, and again
     * on every language change.
     * @param {{name: string, size: string, lastModified: ?number, output: string}} source
     * @param {number} convertedAt - When the conversion finished (ms since the epoch)
     */
    showSource({ name, size, lastModified, output }, convertedAt) {
        if (lastModified === null) {
            setTranslatedText(this.resultSourceFile, 'result.sourceUndated', { name, size });
        } else {
            setTranslatedText(this.resultSourceFile, 'result.source', { name, size, modified: { dateTime: lastModified } });
        }
        setTranslatedText(this.resultConverted, 'result.converted', { time: { time: convertedAt } });
        setTranslatedText(this.resultOutput, 'result.output', { output });
    }

    clearSource() {
        for (const line of [this.resultSourceFile, this.resultConverted, this.resultOutput]) {
            delete line.dataset.i18n;
            delete line.dataset.i18nParams;
            line.textContent = '';
        }
    }

    /**
     * List the parser's findings under the download button; hidden when there are none. A finding
     * located somewhere in particular lists where, one line per place with its example; otherwise
     * its examples are listed. Examples and headings are document text, so they only ever go in
     * through textContent.
     * @param {Array<{code: string, count: number, examples: Array<string>,
     *   occurrences?: Array<{example: string, location: Object}>}>} findings
     * @returns {number} How many notes are listed
     */
    showNotes(findings) {
        this.clearNotes();
        const known = findings.filter(finding => Object.prototype.hasOwnProperty.call(FINDING_KEYS, finding.code));
        if (known.length === 0) return 0;

        let located = false;
        for (const { code, count, examples, occurrences = [] } of known) {
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

            const shown = text => (QUOTED_EXAMPLE_CODES.includes(code) ? `“${text}”` : text);
            const places = occurrences
                .map(({ example, location }) => ({ example, parts: locationParts(location || {}) }))
                .filter(place => place.parts.length > 0);
            if (places.length > 0 || examples.length > 0) {
                const label = document.createElement('p');
                label.className = 'conversion-note-examples-label';
                setTranslatedText(label, places.length > 0 ? 'notes.where' : 'notes.examples');
                item.appendChild(label);

                const list = document.createElement('ul');
                list.className = 'conversion-note-examples';
                if (places.length > 0) {
                    located = true;
                    for (const { example, parts } of places) {
                        const entry = document.createElement('li');
                        entry.className = 'conversion-note-location';
                        parts.forEach(([key, params], i) => {
                            if (i > 0) entry.append(LOCATION_SEPARATOR);
                            const part = document.createElement('span');
                            setTranslatedText(part, key, params);
                            entry.appendChild(part);
                        });
                        if (example) entry.append(LOCATION_SEPARATOR + shown(example));
                        list.appendChild(entry);
                    }
                } else {
                    for (const example of examples) {
                        const entry = document.createElement('li');
                        entry.textContent = shown(example);
                        list.appendChild(entry);
                    }
                }
                item.appendChild(list);
            }

            this.conversionNotesList.appendChild(item);
        }
        this.conversionNotesHelp.classList.toggle('hidden', !located);
        this.conversionNotes.classList.remove('hidden');
        return known.length;
    }

    clearNotes() {
        this.conversionNotesList.replaceChildren();
        this.conversionNotesHelp.classList.add('hidden');
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
