/**
 * Interface language: Turkish by default, with English available
 * Strings live in the dictionaries below; markup marks them with data-i18n attributes:
 *   data-i18n="key"            - sets the element's text
 *   data-i18n-params='{"n":1}' - values for the {placeholders} of its data-i18n string (JSON);
 *                                a value {"dateTime": ms} or {"time": ms} is a moment, written
 *                                in the active language each time the string is shown
 *   data-i18n-html="key"       - sets the element's markup (dictionary strings only, never user input)
 *   data-i18n-attr="attr:key"  - sets an attribute; separate several pairs with ";"
 */

const I18N_STORAGE_KEY = 'word-to-udf.lang';
const I18N_DEFAULT_LANG = 'tr';

const I18N_STRINGS = {
    tr: {
        'meta.title': 'Word’den UDF’ye Dönüştürücü | UYAP Belge Dönüştürücü',
        'meta.description': 'Word belgelerini (.docx) çevrimiçi olarak UYAP UDF biçimine dönüştürün. Ücretsiz, hızlı ve güvenli: dosyalarınız tarayıcınızdan hiç çıkmaz.',
        'banner.tagline': '.DOCX’TEN UYAP UDF’YE.',
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
        'notYet.headers': 'Sayfa numaraları, ilk sayfa üst bilgileri',
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
        'progress.label': 'Dönüştürme ilerlemesi',
        'progress.parsing': 'Word belgesi okunuyor...',
        'progress.converting': 'UDF biçimine dönüştürülüyor...',
        'progress.complete': 'Dönüştürme tamamlandı!',
        'download.burst': 'TAMAM!',
        'download.success': 'Dönüştürme tamamlandı.',
        'download.button': 'UDF DOSYASINI İNDİR',
        'download.another': 'Başka bir dosya dönüştür',
        'download.replace': 'KAYNAĞI DEĞİŞTİR',
        'download.replaceInputLabel': 'Yeni kaynak Word belgesini seçin',
        'result.source': 'Kaynak: {name} ({size}, değiştirilme {modified})',
        'result.sourceUndated': 'Kaynak: {name} ({size})',
        'result.converted': 'Dönüştürüldü: {time}',
        'result.output': 'İndirilecek: {output}',
        'result.announce': 'Dönüştürme tamamlandı.',
        'result.announceOneNote': 'Dönüştürme tamamlandı. Bu dönüşümle ilgili 1 not var.',
        'result.announceNotes': 'Dönüştürme tamamlandı. Bu dönüşümle ilgili {count} not var.',
        'notes.title': 'Bu dönüşümle ilgili notlar',
        'notes.count': 'Sayı: {count}',
        'notes.examples': 'Örnekler:',
        'notes.where': 'Nerede:',
        'notes.locationHelp': 'Paragraf numaraları, belgenin gövdesindeki paragrafları baştan sayar: boş satırlar dahil, tablolardaki ve metin kutularındaki paragraflar hariç. Word’de Giriş sekmesindeki ¶ düğmesi her paragrafın sonunu gösterir. Tablolar da baştan sayılır; satırlar yukarıdan aşağıya, hücreler soldan sağa.',
        'notes.location.paragraph': 'Paragraf {paragraph}',
        'notes.location.cell': 'Tablo {table}, satır {row}, hücre {cell}',
        'notes.location.header': 'Üst bilgi',
        'notes.location.footer': 'Alt bilgi',
        'notes.location.footnote': 'Dipnot',
        'notes.location.endnote': 'Son not',
        'notes.location.heading': '“{heading}” başlığı altında',
        'notes.headersFooters': 'Üst ve alt bilgilerin bir kısmı dönüştürülmedi: ilk sayfa ve çift sayfa üst/alt bilgileri, önceki bölümlerin farklı üst/alt bilgileri, ayrıca son bölümün üst/alt bilgisindeki sayfa numaraları ve diğer alanlar, metin kutuları, filigranlar, tablolar ve gruplanmış çizimlerde ilkinden sonraki resimler. Bu içerik gerekiyorsa Word’de son bölümün ana üst ya da alt bilgisine metin veya resim olarak ekleyin ya da UYAP editöründe ekleyin.',
        'notes.comments': 'Açıklamalar (yorumlar) dönüştürülmedi. Gerekli bir açıklama varsa metnini Word’de belgenin gövdesine yazın.',
        'notes.hiddenText': 'Gizli metin dönüştürülmedi: Word’de gizli olarak biçimlendirilmiş metin, Word’ün göstermediği ve yazdırmadığı gibi UDF’ye de yazılmadı. Bu metin dilekçede yer almalıysa Word’de seçip Yazı Tipi ayarlarında “Gizli” seçeneğini kaldırın.',
        'notes.nestedTables': 'Üç ya da daha fazla tablonun içine yerleştirilmiş bir tablo düzleştirildi: metni, içinde bulunduğu hücreye sırayla yazıldı. Düzen önemliyse Word’de bu tabloyu daha az iç içe olacak şekilde taşıyın ya da düz metne dönüştürün.',
        'notes.missingImage': 'Verisi belgede bulunmayan görseller atlandı. Bu görselleri Word’de silip Ekle › Resimler ile yeniden ekleyin.',
        'notes.unsupportedImage': 'Bazı görseller PNG ya da JPEG değil; UDF yalnızca bu iki biçimi tanımlar, bu yüzden UYAP’ta görünmeyebilirler. Görselleri PNG ya da JPEG olarak kaydedip Word’de yeniden ekleyin.',
        'notes.unconvertedDrawing': 'Bazı grafikler, SmartArt diyagramları, şekiller ya da gömülü nesneler dönüştürülemediği için atlandı. Gerekiyorsa Word’de her birini kopyalayıp Özel Yapıştır › Resim (PNG) ile resim olarak yapıştırın ve belgeyi yeniden dönüştürün.',
        'notes.notesMoved': 'Dipnotlar sayfalarının altında değil: metinleri biçimlendirme olmadan, bir sonraki sayfa sonundan önceye ya da belgenin sonuna yazıldı. Dönüştürülen dosyada dipnot numaralarını ve yerlerini kontrol edin.',
        'notes.endnotesMoved': 'Belge son notları her bölümün sonuna koyuyor ya da numaralarını yeniden başlatıyor, ancak hepsi belgenin sonuna taşındı ve baştan numaralandı. Dönüştürülen dosyada son not numaralarını ve yerlerini kontrol edin ya da Word’de son notları sürekli numarayla belgenin sonuna alın.',
        'notes.lineSpacing': 'Bazı paragraflarda satır aralığı “Tam” ya da “En az” olarak ayarlı; UDF bunu tanımlamadığı için bu paragraflara tek satır aralığı verildi. Aralık önemliyse Word’ün Paragraf ayarlarında satır aralığı olarak “Tek”, “1,5 satır”, “Çift” ya da “Birden çok” seçin.',
        'notes.tabStops': 'Bazı sekme durakları atlandı ya da sadeleşti: çubuk sekmeleri yazılmadı, alt çizgi ve kalın çizgi öncü karakterleri kaldırıldı. Önemliyse Word’ün Sekmeler ayarlarında nokta ya da tire öncü karakteri seçin.',
        'notes.symbols': 'Sembol yazı tiplerinden eklenen bazı semboller (onay kutuları, işaretler, madde imleri gibi) Unicode karşılıklarıyla yazıldı; örneklerde okun ardından gösterilir. UYAP editörünün yazı tipleri bunları farklı çizebilir, dönüştürülen dosyada kontrol edin. Örneklerde oku olmayan sembollerin karşılığı yok, bu yüzden atlandı: gerekiyorsa Word’de düz metinle değiştirin.',
        'notes.tableBorders': 'Bazı tablo hücrelerinin kenarlıkları sadeleşti: UDF her hücreye tek bir çizgi stili, kalınlığı ve rengi verir. Kenarları farklı çizgilerle çizilmiş hücrelerin tüm kenarları en kalın çizgiyle yazıldı; düz, noktalı, kesik ve çift çizgi dışındaki stiller (üçlü, dalgalı, kabartma gibi) düz çizgi olarak yazıldı. Bu hücrelerin kenarlıklarını UYAP editöründe kontrol edin ya da Word’ün Kenarlıklar ayarlarında hücrenin tüm kenarlarına aynı düz, noktalı, kesik ya da çift çizgiyi verin.',
        'notes.tablePlacement': 'Bazı tablolar sol kenar boşluğunda, metnin akışı içinde yazıldı: UDF’de tablonun hizalaması, girintisi ya da konumu tanımlı değil, bu yüzden ortalanmış, sağa yaslanmış, girintili ya da metnin kaydırıldığı (kayan) tablolar sola yaslı ve metin içinde yazıldı. Tablonun yeri önemliyse Word’de tabloyu sola hizalayıp kaydırmasını kaldırın (Tablo Özellikleri → Tablo → Hizalama: Sol, Metni kaydırma: Hiçbiri) ya da tabloyu UYAP editöründe kontrol edin.',
        'notes.equations': 'Denklemler düz, tek satırlık metin olarak yazıldı (kesirler a/b, üsler a^b biçiminde). Bunları UYAP editöründe kontrol edin ya da Word’de düz metin veya resimle değiştirin.',
        'notes.embeddedContent': 'Bu dosyaya başka bir programın gömdüğü içerik (HTML, RTF ya da Word parçaları) dönüştürülmedi. Dosyayı Word’de açıp yeniden kaydedin ve kaydettiğiniz dosyayı dönüştürün: Word gömülü içeriği normal metne çevirir.',
        'notes.trackedChanges': 'Belgede kabul edilmemiş izlenen değişiklikler var (eklenen, silinen ya da taşınan metin, satırlar, biçim değişiklikleri). Belge, tüm değişiklikler kabul edilmiş gibi dönüştürüldü: eklenen metin yazıldı, silinen metin, paragraflar ve satırlar atlandı. Emin olmak için değişiklikleri Word’de kabul edin ya da reddedin (Gözden Geçir › Kabul Et) ve belgeyi yeniden dönüştürün.',
        'notes.sectionLayout': 'Bölümlerin sayfa düzeninin bir kısmı korunamadı: UDF’de tüm belge için tek bir sayfa yapısı var. Önceki bölümlerin farklı sayfa yönü (yatay ya da dikey), kâğıt boyutu ya da kenar boşlukları yerine son bölümünkiler kullanıldı; birden çok sütun tek sütun olarak yazıldı; tek ya da çift numaralı sayfada başlayan bölüm sonları düz sayfa sonu olarak yazıldı. Bu sayfaları UYAP editöründe kontrol edin ya da Word’de içeriği son bölümün sayfa düzenine taşıyın.',
        'notes.subscript': 'Alt simge korunamadı: UDF’de alt simge tanımlı değil, bu yüzden alt simge olarak biçimlendirilmiş metin (H₂O’daki 2 gibi) normal satır üzerinde yazıldı. Bu metni UYAP editöründe kontrol edin ya da Word’de alt simge kullanmadan yeniden yazın (örneğin “H2O”).',
        'notes.doubleStrikethrough': 'Çift üstü çizili metin tek çizgiyle yazıldı: UDF’de çift üstü çizili biçim tanımlı değil, bu yüzden Word’de çift çizgiyle üstü çizilmiş metin tek çizgiyle üstü çizili olarak yazıldı. Bu metni UYAP editöründe kontrol edin.',
        'notes.underlineStyles': 'Bazı altı çizili metinler düz tek çizgiyle yazıldı: UDF’de yalnızca tek çizgili altı çizili biçim var, bu yüzden çift, kalın, noktalı, kesik ya da dalgalı çizgiyle altı çizilmiş metin tek çizgiyle yazıldı; yalnızca sözcüklerin altı çizilmiş metinde sözcükler arasındaki boşlukların da altı çizildi. Word’de tek çizgili altı çizili biçim kullanın ya da bu metni UYAP editöründe kontrol edin.',
        'notes.listMarkers': 'Bazı liste işaretleri UDF’nin sunduğu en yakın biçimle yazıldı: UDF’de yalnızca birkaç madde işareti ve “1.”, “1)”, “a.”, “I.” gibi numara biçimleri var. Bu yüzden başka madde işaretleri (tire, onay işareti, resimli madde işareti gibi), “MADDE 1:” ya da “(1)” gibi numara metinleri, “1.2.” gibi çok düzeyli numaralar, 1’den başlamayan numaralar ve diğer numara biçimleri korunamadı. Listenin numaralarını ya da madde işaretlerini UYAP editöründe kontrol edin.',
        'notes.paragraphBorders': 'Paragraf kenarlıkları yazılmadı: UDF yalnızca tablo hücrelerine kenarlık verebilir, bu yüzden kutu içine alınmış ya da üstü veya altı çizgili paragrafların çizgileri dönüştürülmedi. Kenarlık önemliyse Word’de paragrafı tek hücreli bir tabloya taşıyın (tablo kenarlıklarıyla birlikte dönüştürülür) ya da paragrafı UYAP editöründe kontrol edin.',
        'notes.paragraphShading': 'Paragraf gölgelendirmesi yazılmadı: UDF’de paragrafa arka plan rengi verilemiyor, bu yüzden gri “AÇIKLAMALAR” başlığı gibi gölgeli paragraflar arka plansız yazıldı. Renk önemliyse Word’de paragrafı tek hücreli bir tabloya taşıyıp hücreye gölgelendirme verin (hücre gölgelendirmesi dönüştürülür) ya da paragrafı UYAP editöründe kontrol edin.',
        'notes.hyperlinks': 'Bazı bağlantıların adresi yazılmadı: UDF’de bağlantı tanımlı değil, bu yüzden yalnızca bağlantının metni yazıldı, gösterdiği web ya da e-posta adresi atlandı. Örneklerde adres okun ardından gösterilir. Okuyucunun adrese ihtiyacı varsa adresi Word’de metnin içine yazın (örneğin bağlantıdan sonra parantez içinde) ve belgeyi yeniden dönüştürün.',
        'error.icon': 'HATA',
        'error.default': 'Bir hata oluştu',
        'error.invalidType': 'Lütfen geçerli bir Word belgesi (.docx dosyası) seçin.',
        'error.parse': 'DOCX dosyası okunamadı. Geçerli bir Word belgesi olduğundan emin olun.',
        'error.corrupt': 'Word belgesi bozuk: metni okunamıyor. Belgeyi Word’de açıp yeniden kaydedin, sonra tekrar deneyin.',
        'error.embeddedOnly': 'Bu Word belgesinin içeriğinin tamamı başka bir programın gömdüğü içerikten (HTML, RTF ya da Word parçaları) oluşuyor ve doğrudan dönüştürülemiyor. Belgeyi Word’de açıp yeniden kaydedin ve kaydettiğiniz dosyayı dönüştürün: Word gömülü içeriği normal metne çevirir.',
        'error.encrypted': 'Bu Word belgesi parolayla korunuyor. Parolayı Word’de kaldırın (Dosya › Bilgi › Belgeyi Koru), belgeyi kaydedin ve tekrar deneyin.',
        'error.legacyDoc': 'Bu dosya eski bir Word 97–2003 belgesi (.doc). Word’de açıp “Farklı Kaydet” ile .docx olarak kaydedin ve tekrar deneyin.',
        'error.strict': 'Bu Word belgesi desteklenmeyen Strict Open XML biçiminde kaydedilmiş. Word’de açıp “Farklı Kaydet” ile “Word Belgesi (.docx)” olarak kaydedin ve tekrar deneyin.',
        'error.tooLarge': 'Bu Word belgesi tarayıcıda dönüştürülemeyecek kadar büyük ya da bozuk. Belgeyi Word’de açıp yeniden kaydedin; olmazsa büyük resimleri sıkıştırın (Resim Biçimi › Resimleri Sıkıştır) veya kaldırın, sonra tekrar deneyin.',
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
        'banner.tagline': 'CONVERT .DOCX TO UYAP UDF.',
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
        'notYet.headers': 'Page numbers, first-page headers',
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
        'progress.label': 'Conversion progress',
        'progress.parsing': 'Parsing Word document...',
        'progress.converting': 'Converting to UDF format...',
        'progress.complete': 'Conversion complete!',
        'download.burst': 'DONE!',
        'download.success': 'Conversion complete.',
        'download.button': 'DOWNLOAD UDF FILE',
        'download.another': 'Convert another file',
        'download.replace': 'REPLACE SOURCE',
        'download.replaceInputLabel': 'Choose the new source Word document',
        'result.source': 'Source: {name} ({size}, modified {modified})',
        'result.sourceUndated': 'Source: {name} ({size})',
        'result.converted': 'Converted: {time}',
        'result.output': 'Downloads as: {output}',
        'result.announce': 'Conversion complete.',
        'result.announceOneNote': 'Conversion complete. 1 note on this conversion.',
        'result.announceNotes': 'Conversion complete. {count} notes on this conversion.',
        'notes.title': 'Notes on this conversion',
        'notes.count': 'Count: {count}',
        'notes.examples': 'Examples:',
        'notes.where': 'Where:',
        'notes.locationHelp': 'Paragraph numbers count the paragraphs in the body of the document from the start: empty lines included, paragraphs in tables and text boxes not. In Word, the ¶ button on the Home tab marks the end of each paragraph. Tables are counted from the start too; rows top to bottom, cells left to right.',
        'notes.location.paragraph': 'Paragraph {paragraph}',
        'notes.location.cell': 'Table {table}, row {row}, cell {cell}',
        'notes.location.header': 'Header',
        'notes.location.footer': 'Footer',
        'notes.location.footnote': 'Footnote',
        'notes.location.endnote': 'Endnote',
        'notes.location.heading': 'under “{heading}”',
        'notes.headersFooters': 'Some header and footer content was not converted: first-page and even-page headers and footers, earlier sections’ different headers and footers, and page numbers and other fields, text boxes, watermarks, tables and extra pictures in grouped drawings in the last section’s header or footer. If it is needed, add it as text or a picture in the last section’s main header or footer in Word, or add it in the UYAP editor.',
        'notes.comments': 'Comments were not converted. If a comment matters, write its text into the body of the document in Word.',
        'notes.hiddenText': 'Hidden text was left out: text formatted as hidden in Word, which Word neither shows nor prints, was not written into the UDF. If it belongs in the filing, select it in Word and clear “Hidden” in the Font settings.',
        'notes.nestedTables': 'A table nested inside three or more other tables was flattened: its text was written into the cell holding it, in reading order. If the layout matters, nest that table less deeply or turn it into plain text in Word.',
        'notes.missingImage': 'Pictures whose image data is missing from the document were skipped. Delete them in Word and insert them again with Insert › Pictures.',
        'notes.unsupportedImage': 'Some pictures are not PNG or JPEG, the only formats UDF defines, so they may not show in UYAP. Save them as PNG or JPEG and insert them again in Word.',
        'notes.unconvertedDrawing': 'Some charts, SmartArt diagrams, shapes or embedded objects could not be converted and were left out. If they matter, copy each one in Word, paste it back as a picture with Paste Special › Picture (PNG), and convert again.',
        'notes.notesMoved': 'Footnotes are no longer at the bottom of their page: their text was written without formatting before the next page break, or at the end of the document. Check the note numbers and placement in the converted file.',
        'notes.endnotesMoved': 'The document places endnotes at the end of each section or restarts their numbering, but they were all moved to the end of the document and numbered from the start. Check the endnote numbers and placement in the converted file, or set endnotes to the end of the document with continuous numbering in Word.',
        'notes.lineSpacing': 'Some paragraphs have an “Exactly” or “At least” line spacing, which UDF can’t express, so they got single line spacing. If the spacing matters, choose “Single”, “1.5 lines”, “Double” or “Multiple” as the line spacing in Word’s Paragraph settings.',
        'notes.tabStops': 'Some tab stops were dropped or simplified: bar tab stops were left out, and underline and heavy leaders were removed. If they matter, choose a dot or hyphen leader in Word’s Tabs settings.',
        'notes.symbols': 'Some symbols inserted from symbol fonts (such as check boxes, marks and bullets) were written as their Unicode equivalents, shown after the arrow in the examples. The UYAP editor’s fonts may draw them differently, so check them in the converted file. Symbols with no arrow have no equivalent and were left out: if they matter, replace them with plain text in Word.',
        'notes.tableBorders': 'Some table cell borders were simplified: UDF gives each cell a single line style, width and color. Cells whose edges are drawn with different lines had all their edges written with the heaviest one, and line styles other than single, dotted, dashed and double (such as triple, wavy or 3-D) were written as solid lines. Check these cells’ borders in the UYAP editor, or give all of a cell’s edges the same single, dotted, dashed or double line in Word’s Borders settings.',
        'notes.tablePlacement': 'Some tables were written at the left margin, in the text flow: UDF tables have no alignment, indent or position, so centred, right-aligned, indented and floating (text-wrapped) tables were written left-aligned and inline. If a table’s position matters, align it left and make it non-floating in Word (Table Properties → Table → Alignment: Left, Text wrapping: None), or check it in the UYAP editor.',
        'notes.equations': 'Equations were written as plain linear text (fractions as a/b, powers as a^b). Check them in the UYAP editor, or replace them in Word with ordinary text or a picture.',
        'notes.embeddedContent': 'Content embedded in this file by another program (HTML, RTF or Word chunks) was left out. Open the file in Word, save it again, and convert the saved file: Word turns embedded content into ordinary text.',
        'notes.trackedChanges': 'The document has tracked changes that haven’t been accepted (inserted, deleted or moved text, rows, formatting changes). It was converted as if every change were accepted: inserted text was written, and deleted text, paragraphs and rows were left out. To be sure, accept or reject the changes in Word (Review › Accept) and convert again.',
        'notes.sectionLayout': 'Some section layout couldn’t be kept: UDF has one page setup for the whole document. Earlier sections’ different orientation (landscape or portrait), paper size or margins were replaced by the last section’s; multiple columns were written as a single column; and section breaks that start on an odd or even page were written as plain page breaks. Check these pages in the UYAP editor, or move their content into the last section’s layout in Word.',
        'notes.subscript': 'Subscript couldn’t be kept: UDF defines no subscript, so text formatted as subscript (such as the 2 in H₂O) was written on the baseline. Check it in the UYAP editor, or rewrite it in Word without subscript (e.g. “H2O”).',
        'notes.doubleStrikethrough': 'Double strikethrough was written as a single strikethrough: UDF defines no double strikethrough, so text struck through with two lines in Word was written struck through with one. Check it in the UYAP editor.',
        'notes.underlineStyles': 'Some underlines were written as a plain single underline: UDF has only a single underline, so text underlined with a double, thick, dotted, dashed or wavy line was written with a single one, and for words-only underlining the spaces between words were underlined too. Use a single underline in Word, or check the text in the UYAP editor.',
        'notes.listMarkers': 'Some list markers were written in the closest form UDF has: UDF has only a few bullet shapes and number styles such as “1.”, “1)”, “a.” and “I.”. Other bullets (such as dashes, check marks or picture bullets), number text such as “MADDE 1:” or “(1)”, multi-level numbers such as “1.2.”, numbering that doesn’t start at 1 and other number formats couldn’t be kept. Check the list’s numbering or bullets in the UYAP editor.',
        'notes.paragraphBorders': 'Paragraph borders were left out: UDF can only give table cells borders, so the lines of boxed paragraphs, or of paragraphs with a line above or below, were not converted. If a border matters, move the paragraph into a one-cell table in Word (its borders are converted), or check the paragraph in the UYAP editor.',
        'notes.paragraphShading': 'Paragraph shading was left out: UDF can’t give a paragraph a background color, so shaded paragraphs, such as a grey “AÇIKLAMALAR” banner, were written without one. If the color matters, move the paragraph into a one-cell table in Word and shade the cell (cell shading is converted), or check the paragraph in the UYAP editor.',
        'notes.hyperlinks': 'Some links’ addresses were left out: UDF has no links, so only a link’s text was written and the web or e-mail address it points to was left out. Examples show the address after the arrow. If the reader needs the address, type it into the text in Word (for example in brackets after the link) and convert again.',
        'error.icon': 'ERROR',
        'error.default': 'An error occurred',
        'error.invalidType': 'Please select a valid Word document (.docx file)',
        'error.parse': 'Failed to parse DOCX file. Make sure it is a valid Word document.',
        'error.corrupt': 'The Word document is damaged: its text can’t be read. Open it in Word, save it again, then try again.',
        'error.embeddedOnly': 'All of this Word document’s content is embedded by another program (HTML, RTF or Word chunks), which can’t be converted directly. Open it in Word, save it again, and convert the saved file: Word turns embedded content into ordinary text.',
        'error.encrypted': 'This Word document is password-protected. Remove the password in Word (File › Info › Protect Document), save it, and try again.',
        'error.legacyDoc': 'This file is an old Word 97–2003 document (.doc). Open it in Word, use “Save As” to save it as .docx, and try again.',
        'error.strict': 'This Word document is saved in the Strict Open XML format, which isn’t supported. Open it in Word, use “Save As” to save it as “Word Document (.docx)”, and try again.',
        'error.tooLarge': 'This Word document is too large, or too damaged, to convert in the browser. Open it in Word and save it again; if that doesn’t help, compress (Picture Format › Compress Pictures) or remove large pictures, then try again.',
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

// How moment params ({"dateTime": ms} or {"time": ms}) are written, in the active language
const I18N_MOMENT_FORMATS = {
    dateTime: { year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit' },
    time: { hour: '2-digit', minute: '2-digit', second: '2-digit' }
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
    return text.replace(/\{(\w+)\}/g, (match, name) => (name in params ? formatParam(params[name]) : match));
}

/**
 * Write a placeholder value: a moment ({dateTime: ms} or {time: ms}) with Intl.DateTimeFormat
 * in the active language, anything else as a string
 * @param {*} value
 * @returns {string}
 */
function formatParam(value) {
    if (value !== null && typeof value === 'object') {
        for (const [style, options] of Object.entries(I18N_MOMENT_FORMATS)) {
            if (typeof value[style] === 'number') return new Intl.DateTimeFormat(currentLang, options).format(value[style]);
        }
    }
    return String(value);
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
        el.textContent = t(el.dataset.i18n, el.dataset.i18nParams ? JSON.parse(el.dataset.i18nParams) : {});
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
 * @param {Object} [params] - Values for {name} placeholders, kept for later language changes
 */
function setTranslatedText(el, key, params) {
    el.dataset.i18n = key;
    if (params) el.dataset.i18nParams = JSON.stringify(params);
    else delete el.dataset.i18nParams;
    el.textContent = t(key, params);
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
