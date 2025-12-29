/**
 * Utility functions for Word to UDF Converter
 */

/**
 * Format file size in human-readable format
 * @param {number} bytes - File size in bytes
 * @returns {string} Formatted file size
 */
function formatFileSize(bytes) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

/**
 * Trigger file download from a Blob
 * @param {Blob} blob - The file blob
 * @param {string} filename - The filename for download
 */
function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/**
 * Escape special XML characters
 * @param {string} text - Text to escape
 * @returns {string} Escaped text
 */
function escapeXml(text) {
  if (!text) return '';
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Convert ArrayBuffer to Base64 string
 * @param {ArrayBuffer} buffer - The array buffer
 * @returns {string} Base64 encoded string
 */
function arrayBufferToBase64(buffer) {
  let binary = '';
  const bytes = new Uint8Array(buffer);
  const len = bytes.byteLength;
  for (let i = 0; i < len; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary);
}

/**
 * Get file extension from filename
 * @param {string} filename - The filename
 * @returns {string} The extension (without dot)
 */
function getFileExtension(filename) {
  return filename.split('.').pop().toLowerCase();
}

/**
 * Generate a unique ID
 * @returns {string} Unique ID
 */
function generateId() {
  return Date.now().toString(36) + Math.random().toString(36).substr(2);
}

/**
 * Sleep for specified milliseconds
 * @param {number} ms - Milliseconds to sleep
 * @returns {Promise} Promise that resolves after delay
 */
function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * Parse XML string to Document
 * @param {string} xmlString - XML content
 * @returns {Document} Parsed XML document
 */
function parseXml(xmlString) {
  const parser = new DOMParser();
  return parser.parseFromString(xmlString, 'application/xml');
}

/**
 * Get text content from XML element, handling namespaces
 * @param {Element} element - The XML element
 * @param {string} tagName - Tag name to find
 * @param {string} namespace - XML namespace
 * @returns {string|null} Text content or null
 */
function getXmlText(element, tagName, namespace) {
  const el = element.getElementsByTagNameNS(namespace, tagName)[0];
  return el ? el.textContent : null;
}

// Word ML namespace
const WORD_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const DRAWING_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const RELATIONSHIPS_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PICTURE_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const WORDML_DRAWING_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
