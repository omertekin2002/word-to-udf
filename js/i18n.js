/**
 * Interface language: Turkish by default, with English available
 * Strings live in the dictionaries below; markup marks them with data-i18n attributes:
 *   data-i18n="key"            - sets the element's text
 *   data-i18n-html="key"       - sets the element's markup (dictionary strings only, never user input)
 *   data-i18n-attr="attr:key"  - sets an attribute; separate several pairs with ";"
 */

const I18N_STORAGE_KEY = 'word-to-udf.lang';
const I18N_DEFAULT_LANG = 'tr';

const I18N_STRINGS = {
    tr: {
        'meta.title': 'Word’den UDF’ye Dönüştürücü | UYAP Belge Dönüştürücü',
        'meta.description': 'Word belgelerini (.docx) çevrimiçi olarak UYAP UDF biçimine dönüştürün. Ücretsiz, hızlı ve güvenli: dosyalarınız tarayıcınızdan hiç çıkmaz.',
        'banner.tagline': '.DOCX’TEN UYAP UDF’YE. ÇEVRİMİÇİ.',
        'lang.label': 'Dil',
        'nav.quickLinks': 'Hızlı bağlantılar',
        'nav.source': 'KAYNAK',
        'privacy.lead': 'Dosyalarınız tarayıcınızdan hiç çıkmaz.',
        'privacy.body': 'Dönüştürme tamamen kendi bilgisayarınızda yapılır; hiçbir şey bir sunucuya yüklenmez.',
        'converts.title': 'DÖNÜŞTÜRÜLENLER',
        'converts.formatting': 'Kalın, italik, yazı tipleri',
        'converts.alignment': 'Hizalama',
        'converts.tables': 'Tablolar ve kenarlıklar',
        'converts.images': 'Görseller',
        'notYet.title': 'HENÜZ YOK',
        'notYet.headers': 'Üst ve alt bilgiler',
        'notYet.trackChanges': 'Değişiklik izleme, açıklamalar',
        'notYet.ole': 'Gömülü OLE nesneleri',
        'convert.heading': 'DÖNÜŞTÜR',
        'upload.title': 'BİR WORD BELGESİ SEÇİN',
        'upload.inputLabel': 'Word belgesi yükleyin',
        'upload.text': '<strong>.docx</strong> dosyanızı buraya sürükleyip bırakın',
        'upload.subtext': 'ya da <u>göz atmak için tıklayın</u>',
        'file.remove': 'KALDIR',
        'file.removeLabel': 'Dosyayı kaldır',
        'convert.button': 'UDF’YE DÖNÜŞTÜR →',
        'progress.parsing': 'Word belgesi okunuyor...',
        'progress.converting': 'UDF biçimine dönüştürülüyor...',
        'progress.complete': 'Dönüştürme tamamlandı!',
        'download.burst': 'TAMAM!',
        'download.success': 'Dönüştürme tamamlandı.',
        'download.button': 'UDF DOSYASINI İNDİR',
        'download.another': 'Başka bir dosya dönüştür',
        'error.icon': 'HATA',
        'error.default': 'Bir hata oluştu',
        'error.invalidType': 'Lütfen geçerli bir Word belgesi (.docx dosyası) seçin.',
        'error.parse': 'DOCX dosyası okunamadı. Geçerli bir Word belgesi olduğundan emin olun.',
        'error.generic': 'Dönüştürme sırasında bir hata oluştu. Lütfen tekrar deneyin.',
        'seal.local.label': '%100 tarayıcıda',
        'seal.local.top': '%100',
        'seal.local.mid': 'YEREL',
        'seal.local.bot': 'SUNUCUSUZ',
        'seal.open.label': 'Ücretsiz ve açık kaynak',
        'seal.open.top': 'ÜCRETSİZ',
        'seal.open.mid': 'AÇIK',
        'seal.open.bot': 'KAYNAK',
        'footer.navLabel': 'Alt bilgi',
        'footer.convert': 'DÖNÜŞTÜR',
        'footer.source': 'KAYNAK',
        'footer.copyright': 'Telif hakkı',
        'footer.license': 'MIT Lisansı.',
        'footer.smallPrint': 'Bu site en iyi 3.0 ve üzeri tarayıcı sürümleriyle görüntülenir.'
    },
    en: {
        'meta.title': 'Word to UDF Converter | UYAP Document Converter',
        'meta.description': 'Convert Word documents (.docx) to UYAP UDF format online. Free, fast, and secure - your files never leave your browser.',
        'banner.tagline': 'CONVERT .DOCX TO UYAP UDF. ONLINE.',
        'lang.label': 'Language',
        'nav.quickLinks': 'Quick links',
        'nav.source': 'SOURCE',
        'privacy.lead': 'Your files never leave your browser.',
        'privacy.body': 'The whole conversion runs on your own computer — nothing is uploaded to a server.',
        'converts.title': 'CONVERTS',
        'converts.formatting': 'Bold, italic, fonts',
        'converts.alignment': 'Alignment',
        'converts.tables': 'Tables & borders',
        'converts.images': 'Images',
        'notYet.title': 'NOT YET',
        'notYet.headers': 'Headers & footers',
        'notYet.trackChanges': 'Track changes, comments',
        'notYet.ole': 'Embedded OLE objects',
        'convert.heading': 'CONVERT',
        'upload.title': 'SELECT A WORD DOCUMENT',
        'upload.inputLabel': 'Upload Word document',
        'upload.text': 'Drag and drop your <strong>.docx</strong> file here',
        'upload.subtext': 'or <u>click to browse</u>',
        'file.remove': 'REMOVE',
        'file.removeLabel': 'Remove file',
        'convert.button': 'CONVERT TO UDF →',
        'progress.parsing': 'Parsing Word document...',
        'progress.converting': 'Converting to UDF format...',
        'progress.complete': 'Conversion complete!',
        'download.burst': 'DONE!',
        'download.success': 'Conversion complete.',
        'download.button': 'DOWNLOAD UDF FILE',
        'download.another': 'Convert another file',
        'error.icon': 'ERROR',
        'error.default': 'An error occurred',
        'error.invalidType': 'Please select a valid Word document (.docx file)',
        'error.parse': 'Failed to parse DOCX file. Make sure it is a valid Word document.',
        'error.generic': 'An error occurred during conversion. Please try again.',
        'seal.local.label': '100% client-side',
        'seal.local.top': '100%',
        'seal.local.mid': 'LOCAL',
        'seal.local.bot': 'NO UPLOAD',
        'seal.open.label': 'Free and open source',
        'seal.open.top': 'OPEN',
        'seal.open.mid': 'FREE',
        'seal.open.bot': 'SOURCE',
        'footer.navLabel': 'Footer',
        'footer.convert': 'CONVERT',
        'footer.source': 'SOURCE',
        'footer.copyright': 'Copyright',
        'footer.license': 'MIT License.',
        'footer.smallPrint': 'This site is best viewed with browser versions 3.0 and higher.'
    }
};

let currentLang = I18N_DEFAULT_LANG;

function isSupportedLang(lang) {
    return Object.prototype.hasOwnProperty.call(I18N_STRINGS, lang);
}

/**
 * Look up a string in the active language
 * @param {string} key - Dictionary key
 * @param {Object} [params] - Values for {name} placeholders
 * @returns {string} The translated string, or the key itself if it is missing
 */
function t(key, params = {}) {
    const dictionary = I18N_STRINGS[currentLang];
    const text = Object.prototype.hasOwnProperty.call(dictionary, key) ? dictionary[key] : key;
    return text.replace(/\{(\w+)\}/g, (match, name) => (name in params ? String(params[name]) : match));
}

/**
 * Pick the language: ?lang= in the URL, then a saved choice, then Turkish.
 * The browser's own language is deliberately ignored.
 * @returns {string}
 */
function detectLanguage() {
    const fromUrl = new URLSearchParams(window.location.search).get('lang');
    if (isSupportedLang(fromUrl)) return fromUrl;

    let saved = null;
    try {
        saved = window.localStorage.getItem(I18N_STORAGE_KEY);
    } catch (e) {
        // Storage can be unavailable (privacy modes, file:// in some browsers)
    }
    return isSupportedLang(saved) ? saved : I18N_DEFAULT_LANG;
}

/**
 * Show the whole page in a language
 * @param {string} lang - 'tr' or 'en'
 */
function applyLanguage(lang) {
    currentLang = isSupportedLang(lang) ? lang : I18N_DEFAULT_LANG;

    document.documentElement.lang = currentLang;
    document.title = t('meta.title');

    for (const el of document.querySelectorAll('[data-i18n]')) {
        el.textContent = t(el.dataset.i18n);
    }
    for (const el of document.querySelectorAll('[data-i18n-html]')) {
        el.innerHTML = t(el.dataset.i18nHtml);
    }
    for (const el of document.querySelectorAll('[data-i18n-attr]')) {
        for (const pair of el.dataset.i18nAttr.split(';')) {
            const [attr, key] = pair.split(':').map(part => part.trim());
            if (attr && key) el.setAttribute(attr, t(key));
        }
    }
    for (const button of document.querySelectorAll('[data-set-lang]')) {
        button.setAttribute('aria-pressed', String(button.dataset.setLang === currentLang));
    }
}

/**
 * Switch language from the toggle: apply it, remember it, and keep a ?lang= in the URL in step
 * @param {string} lang - 'tr' or 'en'
 */
function setLanguage(lang) {
    if (!isSupportedLang(lang)) return;
    applyLanguage(lang);

    try {
        window.localStorage.setItem(I18N_STORAGE_KEY, lang);
    } catch (e) {
        // Not saved; the choice still holds for this visit
    }

    const url = new URL(window.location.href);
    if (url.searchParams.has('lang')) {
        url.searchParams.set('lang', lang);
        window.history.replaceState(null, '', url);
    }
}

/**
 * Show a dictionary string in an element and keep it translated on later language changes
 * @param {HTMLElement} el
 * @param {string} key - Dictionary key
 */
function setTranslatedText(el, key) {
    el.dataset.i18n = key;
    el.textContent = t(key);
}

function currentLanguage() {
    return currentLang;
}

function initI18n() {
    for (const button of document.querySelectorAll('[data-set-lang]')) {
        button.addEventListener('click', () => setLanguage(button.dataset.setLang));
    }
    applyLanguage(detectLanguage());
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initI18n);
} else {
    initI18n();
}
