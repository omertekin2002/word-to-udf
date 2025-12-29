/**
 * Main Application Logic
 * Handles UI interactions and orchestrates the conversion process
 */

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
        this.errorMessage = document.getElementById('errorMessage');
        this.errorText = document.getElementById('errorText');
        this.themeToggle = document.getElementById('themeToggle');

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
        this.initTheme();
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

        // Theme toggle
        this.themeToggle.addEventListener('click', () => this.toggleTheme());
    }

    initTheme() {
        // Check for saved theme preference - light mode by default
        const savedTheme = localStorage.getItem('theme');

        if (savedTheme === 'dark') {
            document.documentElement.setAttribute('data-theme', 'dark');
            this.themeToggle.querySelector('.icon').textContent = '🌙';
        } else {
            // Light mode is default (no data-theme attribute needed)
            document.documentElement.removeAttribute('data-theme');
            this.themeToggle.querySelector('.icon').textContent = '☀️';
        }
    }

    toggleTheme() {
        const currentTheme = document.documentElement.getAttribute('data-theme');

        if (currentTheme === 'dark') {
            // Switch to light
            document.documentElement.removeAttribute('data-theme');
            this.themeToggle.querySelector('.icon').textContent = '☀️';
            localStorage.setItem('theme', 'light');
        } else {
            // Switch to dark
            document.documentElement.setAttribute('data-theme', 'dark');
            this.themeToggle.querySelector('.icon').textContent = '🌙';
            localStorage.setItem('theme', 'dark');
        }
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
            this.showError('Please select a valid Word document (.docx file)');
            return;
        }

        this.selectedFile = file;
        this.outputFilename = file.name.replace('.docx', '.udf');

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

        try {
            // Show progress
            this.convertBtn.classList.add('hidden');
            this.progressContainer.classList.remove('hidden');
            this.hideError();

            // Update progress: Parsing
            this.updateProgress(20, 'Parsing Word document...');
            await sleep(300);

            // Parse DOCX
            const document = await this.parser.parse(this.selectedFile);

            // Update progress: Converting
            this.updateProgress(60, 'Converting to UDF format...');
            await sleep(300);

            // Generate UDF
            this.convertedBlob = await this.generator.generate(document);

            // Update progress: Complete
            this.updateProgress(100, 'Conversion complete!');
            await sleep(500);

            // Show download section
            this.progressContainer.classList.add('hidden');
            this.fileInfo.classList.add('hidden');
            this.downloadSection.classList.remove('hidden');

        } catch (error) {
            console.error('Conversion error:', error);
            this.progressContainer.classList.add('hidden');
            this.convertBtn.classList.remove('hidden');
            this.showError(error.message || 'An error occurred during conversion. Please try again.');
        }
    }

    updateProgress(percent, text) {
        this.progressFill.style.width = `${percent}%`;
        this.progressText.textContent = text;
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
        this.hideError();
    }

    showError(message) {
        this.errorText.textContent = message;
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
