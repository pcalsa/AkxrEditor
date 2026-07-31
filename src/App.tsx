import { useCallback, useEffect, useMemo, useRef, useState, type PointerEvent as ReactPointerEvent } from 'react'
import { EditorContent, useEditor, type Editor } from '@tiptap/react'
import { Mark, mergeAttributes } from '@tiptap/core'
import StarterKit from '@tiptap/starter-kit'
import Underline from '@tiptap/extension-underline'
import TextAlign from '@tiptap/extension-text-align'
import { Table, TableCell, TableHeader, TableRow } from '@tiptap/extension-table'
import { Color } from '@tiptap/extension-color'
import { FontFamily } from '@tiptap/extension-font-family'
import { TextStyle, FontSize } from '@tiptap/extension-text-style'
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  ArrowRight,
  Baseline,
  Bold,
  Brush,
  Circle as CircleIcon,
  Download,
  Eraser,
  FileArchive,
  FileText,
  Highlighter,
  ImagePlus,
  Italic,
  List,
  ListOrdered,
  MousePointer2,
  Move,
  PaintBucket,
  Paintbrush,
  PenLine,
  Plus,
  RectangleHorizontal,
  Redo2,
  RotateCw,
  Rows3,
  ScissorsLineDashed,
  Shapes,
  Table2,
  Trash2,
  Type,
  Underline as UnderlineIcon,
  Undo2,
  Upload,
  X,
} from 'lucide-react'
import {
  ActiveSelection,
  Canvas as FabricCanvas,
  Circle,
  Control,
  Group,
  IText,
  Line,
  PencilBrush,
  Point,
  Rect,
  Textbox,
  Triangle,
  util as fabricUtil,
  type FabricObject,
} from 'fabric'
import TurndownService from 'turndown'
import mammoth from 'mammoth'
import JSZip from 'jszip'
import * as pdfjsLib from 'pdfjs-dist'
import pdfWorker from 'pdfjs-dist/build/pdf.worker.mjs?url'
import { Document as DocxDocument, Packer, Paragraph, TextRun } from 'docx'
import { jsPDF } from 'jspdf'
import './App.css'

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker

type Mode = 'editor' | 'convert'
type ExportFormat = 'pdf' | 'docx' | 'txt' | 'md' | 'html' | 'png' | 'jpg'
type EditorTool = 'write' | 'select-text' | 'select-paint' | 'brush' | 'eraser' | 'bucket' | 'place-text'
type LayerKind = 'text' | 'paint' | 'images'
type UndoKind = 'text' | 'paint' | 'pages'
type PageSizeKey = 'a4' | 'a5' | 'a4Landscape' | 'a5Landscape'
type AppLanguage = 'en' | 'de' | 'ar'
type PageMargins = {
  top: number
  right: number
  bottom: number
  left: number
}
type SelectedTarget =
  | { kind: 'image'; pageId: string; imageId: string }
  | { kind: 'paint'; pageId: string }
  | { kind: 'table-cell'; pageId: string; cellIndex: number }
  | { kind: 'text'; pageId: string }
type PageContextTarget = { kind: 'page'; pageId: string }
type ContextMenuTarget = SelectedTarget | PageContextTarget
type CanvasPoint = { x: number; y: number }
type ContextMenuState = {
  pastePoint?: CanvasPoint
  target: ContextMenuTarget
  x: number
  y: number
}
type FloatingImage = {
  id: string
  src: string
  x: number
  y: number
  width: number
  height: number
  rotation: number
}
type PageData = {
  id: string
  html: string
  images: FloatingImage[]
  margins?: PageMargins
  fontFamily?: string
  pageImage?: string
  size?: PageSizeKey
}
type MovableTextMode = 'plain' | 'bullet' | 'numbered' | 'table'
type ManagedIText = IText & {
  akxrTextMode?: MovableTextMode
  akxrFormattingText?: boolean
}
type HorizontalTableGroup = Group & {
  akxrObjectType?: 'horizontal-table'
  akxrTableCellCount?: number
  akxrTableCellWidth?: number
}
type PendingImageExport = {
  baseName: string
  format: 'png' | 'jpg'
  pageCount: number
}

const exportFormats: ExportFormat[] = ['pdf', 'docx', 'txt', 'md', 'html', 'png', 'jpg']
const canvasObjectCustomProperties = ['akxrObjectType', 'akxrTableCellCount', 'akxrTableCellWidth', 'akxrTextMode']
const defaultFontSize = '12px'
const fontSizes = ['8px', '10px', '12px', '14px', '16px', '18px', '20px', '24px', '32px', '48px', '72px']
const fontFamilies = ['Calibri', 'Cambria', 'Arial', 'Georgia', 'Courier New', 'Times New Roman']
const languages: { code: AppLanguage; label: string }[] = [
  { code: 'en', label: 'English' },
  { code: 'de', label: 'Deutsch' },
  { code: 'ar', label: 'العربية' },
]
const translations = {
  en: {
    addA4Page: 'Add A4 page',
    addA5Page: 'Add A5 page',
    addImage: 'Add image',
    addImagePage: 'Add image page',
    addPdfPages: 'Add PDF pages',
    alignCenter: 'Align center',
    alignLeft: 'Align left',
    alignRight: 'Align right',
    arrangeAndExport: 'Arrange and export',
    arrow: 'Arrow',
    bold: 'Bold',
    brush: 'Brush',
    bulletList: 'Bullet list',
    bucket: 'Paint bucket',
    cancel: 'Cancel',
    canvasConversionUnavailable: 'Canvas conversion is not available in this browser.',
    canvasExportUnavailable: 'Canvas export is not available in this browser.',
    canvasRenderUnavailable: 'Canvas rendering is not available in this browser.',
    chooseFile: 'Choose file',
    circle: 'Circle',
    clearDrawing: 'Clear drawing',
    commonFontSizes: 'Common font sizes',
    convert: 'Convert',
    continue: 'Continue',
    convertAFIle: 'Convert a File',
    convertDropZone: 'Drop into the ribbon, export from the browser',
    convertFinePrint: 'PDF to DOCX uses text extraction only; browser conversion cannot preserve full PDF layout.',
    convertPrompt: 'Upload a DOCX, PDF, TXT, MD, PNG, or JPG file to convert it locally.',
    converted: 'Converted {file} to {format}.',
    converting: 'Converting in your browser...',
    couldNotAddImagePage: 'Could not add that image page.',
    couldNotAddPdfPages: 'Could not add that PDF.',
    couldNotExportCanvas: 'Could not export the canvas.',
    couldNotExport: 'Could not export this document.',
    couldNotLoadImage: 'Could not load that image.',
    couldNotOpen: 'Could not open that file.',
    couldNotReadFile: 'Could not read that file.',
    copy: 'Copy',
    delete: 'Delete',
    deleteCell: 'Delete cell',
    deletePage: 'Delete page',
    deletePageNumber: 'Delete page {number}',
    deleteTable: 'Delete table',
    drawing: 'Drawing',
    editor: 'Editor',
    editorModes: 'Editor modes',
    eraser: 'Eraser',
    export: 'Export',
    exportFileName: 'Export file name',
    exportFormat: 'Export format',
    fileName: 'File name',
    imageExportMultiplePagesMessage:
      'This document has {count} pages. Each page will be downloaded separately as a {format} file.',
    imageExportMultiplePagesTitle: 'Multiple pages',
    imageOcrUnavailable: 'Image to text, Markdown, or DOCX needs OCR, which is not included in this client-side build.',
    files: 'Files',
    fillColor: 'Fill color',
    fontFamily: 'Font family',
    fontSizeInPixels: 'Font size in pixels',
    fontSizeSlider: 'Font size slider',
    formatting: 'Formatting',
    heading: 'Heading',
    heading1: 'Heading 1',
    heading2: 'Heading 2',
    heading3: 'Heading 3',
    history: 'History',
    insertTable: 'Insert table',
    italic: 'Italic',
    language: 'Language',
    line: 'Line',
    move: 'Move',
    numberedList: 'Numbered list',
    openFile: 'Open file',
    page: 'Page {number}',
    pageSizeA4: 'A4',
    pageSizeA5: 'A5',
    pageSizeA4Landscape: 'A4 landscape',
    pageSizeA5Landscape: 'A5 landscape',
    paintFront: 'Paint front',
    paragraph: 'Paragraph',
    pdfDocxExtractedOnly: 'PDF to DOCX exports extracted text only.',
    paste: 'Paste',
    placeText: 'Place text',
    rectangle: 'Rectangle',
    redo: 'Redo',
    resizeHeight: 'Resize height',
    resizeProportionally: 'Resize proportionally',
    resizeWidth: 'Resize width',
    rotate: 'Rotate',
    selectedToBack: 'Selected to back',
    selectedToFront: 'Selected to front',
    selectPaint: 'Select paint',
    selectText: 'Select text',
    setFontSize: 'Set font size to {size}px',
    strokeColor: 'Stroke color',
    targetFormat: 'Target format',
    textColor: 'Text color',
    textFront: 'Text front',
    textGap: 'Text gap',
    textSize: 'Text size',
    textTools: 'Text tools',
    underline: 'Underline',
    undo: 'Undo',
    unavailableConversion: 'This conversion is not available in the browser.',
    unsupportedSourceFile: 'Unsupported source file. Use DOCX, PDF, TXT, MD, HTML, PNG, or JPG.',
    unsavedWarning: 'Are you sure you want to reload? You will lose all unsaved changes.',
    write: 'Write',
    xShape: 'X',
  },
  de: {
    addA4Page: 'A4-Seite hinzufugen',
    addA5Page: 'A5-Seite hinzufugen',
    addImage: 'Bild hinzufugen',
    addImagePage: 'Bildseite hinzufugen',
    addPdfPages: 'PDF-Seiten hinzufugen',
    alignCenter: 'Zentriert ausrichten',
    alignLeft: 'Links ausrichten',
    alignRight: 'Rechts ausrichten',
    arrangeAndExport: 'Anordnen und exportieren',
    arrow: 'Pfeil',
    bold: 'Fett',
    brush: 'Pinsel',
    bulletList: 'Aufzahlung',
    bucket: 'Fulleimer',
    cancel: 'Abbrechen',
    canvasConversionUnavailable: 'Canvas-Konvertierung ist in diesem Browser nicht verfugbar.',
    canvasExportUnavailable: 'Canvas-Export ist in diesem Browser nicht verfugbar.',
    canvasRenderUnavailable: 'Canvas-Darstellung ist in diesem Browser nicht verfugbar.',
    chooseFile: 'Datei auswahlen',
    circle: 'Kreis',
    clearDrawing: 'Zeichnung loschen',
    commonFontSizes: 'Ubliche Schriftgrossen',
    convert: 'Konvertieren',
    continue: 'Fortfahren',
    convertAFIle: 'Datei konvertieren',
    convertDropZone: 'In die Leiste ziehen, im Browser exportieren',
    convertFinePrint: 'PDF zu DOCX nutzt nur Textextraktion; die Browser-Konvertierung kann das vollstandige PDF-Layout nicht erhalten.',
    convertPrompt: 'Lade eine DOCX-, PDF-, TXT-, MD-, PNG- oder JPG-Datei hoch, um sie lokal zu konvertieren.',
    converted: '{file} wurde in {format} konvertiert.',
    converting: 'Konvertierung im Browser...',
    couldNotAddImagePage: 'Diese Bildseite konnte nicht hinzugefugt werden.',
    couldNotAddPdfPages: 'Diese PDF konnte nicht hinzugefugt werden.',
    couldNotExportCanvas: 'Die Zeichenflache konnte nicht exportiert werden.',
    couldNotExport: 'Dieses Dokument konnte nicht exportiert werden.',
    couldNotLoadImage: 'Dieses Bild konnte nicht geladen werden.',
    couldNotOpen: 'Diese Datei konnte nicht geoffnet werden.',
    couldNotReadFile: 'Diese Datei konnte nicht gelesen werden.',
    copy: 'Kopieren',
    delete: 'Loschen',
    deleteCell: 'Zelle loschen',
    deletePage: 'Seite loschen',
    deletePageNumber: 'Seite {number} loschen',
    deleteTable: 'Tabelle loschen',
    drawing: 'Zeichnen',
    editor: 'Editor',
    editorModes: 'Editor-Modi',
    eraser: 'Radierer',
    export: 'Exportieren',
    exportFileName: 'Export-Dateiname',
    exportFormat: 'Exportformat',
    fileName: 'Dateiname',
    imageExportMultiplePagesMessage:
      'Dieses Dokument hat {count} Seiten. Jede Seite wird separat als {format}-Datei heruntergeladen.',
    imageExportMultiplePagesTitle: 'Mehrere Seiten',
    imageOcrUnavailable: 'Bild zu Text, Markdown oder DOCX benotigt OCR, das in diesem clientseitigen Build nicht enthalten ist.',
    files: 'Dateien',
    fillColor: 'Fullfarbe',
    fontFamily: 'Schriftart',
    fontSizeInPixels: 'Schriftgrosse in Pixeln',
    fontSizeSlider: 'Schriftgrossen-Schieberegler',
    formatting: 'Formatierung',
    heading: 'Uberschrift',
    heading1: 'Uberschrift 1',
    heading2: 'Uberschrift 2',
    heading3: 'Uberschrift 3',
    history: 'Verlauf',
    insertTable: 'Tabelle einfugen',
    italic: 'Kursiv',
    language: 'Sprache',
    line: 'Linie',
    move: 'Verschieben',
    numberedList: 'Nummerierte Liste',
    openFile: 'Datei offnen',
    page: 'Seite {number}',
    pageSizeA4: 'A4',
    pageSizeA5: 'A5',
    pageSizeA4Landscape: 'A4 quer',
    pageSizeA5Landscape: 'A5 quer',
    paintFront: 'Zeichnung nach vorne',
    paragraph: 'Absatz',
    pdfDocxExtractedOnly: 'PDF zu DOCX exportiert nur extrahierten Text.',
    paste: 'Einfugen',
    placeText: 'Text platzieren',
    rectangle: 'Rechteck',
    redo: 'Wiederholen',
    resizeHeight: 'Hohe andern',
    resizeProportionally: 'Proportional skalieren',
    resizeWidth: 'Breite andern',
    rotate: 'Drehen',
    selectedToBack: 'Auswahl nach hinten',
    selectedToFront: 'Auswahl nach vorne',
    selectPaint: 'Zeichnung auswahlen',
    selectText: 'Text auswahlen',
    setFontSize: 'Schriftgrosse auf {size}px setzen',
    strokeColor: 'Linienfarbe',
    targetFormat: 'Zielformat',
    textColor: 'Textfarbe',
    textFront: 'Text nach vorne',
    textGap: 'Textabstand',
    textSize: 'Textgrosse',
    textTools: 'Textwerkzeuge',
    underline: 'Unterstrichen',
    undo: 'Ruckgangig',
    unavailableConversion: 'Diese Konvertierung ist im Browser nicht verfugbar.',
    unsupportedSourceFile: 'Nicht unterstutzte Quelldatei. Verwende DOCX, PDF, TXT, MD, HTML, PNG oder JPG.',
    unsavedWarning: 'Mochtest du wirklich neu laden? Alle nicht gespeicherten Anderungen gehen verloren.',
    write: 'Schreiben',
    xShape: 'X',
  },
  ar: {
    cancel: 'الغاء',
    continue: 'متابعة',
    bucket: 'دلو التعبئة',
    canvasConversionUnavailable: 'تحويل اللوحة غير متاح في هذا المتصفح.',
    canvasExportUnavailable: 'تصدير اللوحة غير متاح في هذا المتصفح.',
    canvasRenderUnavailable: 'عرض اللوحة غير متاح في هذا المتصفح.',
    couldNotExportCanvas: 'تعذر تصدير اللوحة.',
    couldNotLoadImage: 'تعذر تحميل هذه الصورة.',
    couldNotReadFile: 'تعذر قراءة هذا الملف.',
    imageExportMultiplePagesMessage:
      'يحتوي هذا المستند على {count} صفحات. سيتم تنزيل كل صفحة بشكل منفصل كملف {format}.',
    imageExportMultiplePagesTitle: 'صفحات متعددة',
    imageOcrUnavailable: 'تحويل الصورة الى نص او Markdown او DOCX يحتاج الى OCR، وهو غير مضمن في هذا الاصدار المحلي.',
    pageSizeA4: 'A4',
    pageSizeA5: 'A5',
    pageSizeA4Landscape: 'A4 افقي',
    pageSizeA5Landscape: 'A5 افقي',
    unsupportedSourceFile: 'نوع الملف غير مدعوم. استخدم DOCX او PDF او TXT او MD او HTML او PNG او JPG.',
    addA4Page: 'اضافة صفحة A4',
    addA5Page: 'اضافة صفحة A5',
    addImage: 'اضافة صورة',
    addImagePage: 'اضافة صفحة صورة',
    addPdfPages: 'اضافة صفحات PDF',
    alignCenter: 'محاذاة للوسط',
    alignLeft: 'محاذاة لليسار',
    alignRight: 'محاذاة لليمين',
    arrangeAndExport: 'ترتيب وتصدير',
    arrow: 'سهم',
    bold: 'عريض',
    brush: 'فرشاة',
    bulletList: 'قائمة نقطية',
    chooseFile: 'اختر ملفا',
    circle: 'دائرة',
    clearDrawing: 'مسح الرسم',
    commonFontSizes: 'احجام الخط الشائعة',
    convert: 'تحويل',
    convertAFIle: 'تحويل ملف',
    convertDropZone: 'اسحب الى الشريط ثم صدر من المتصفح',
    convertFinePrint: 'تحويل PDF الى DOCX يستخدم استخراج النص فقط؛ تحويل المتصفح لا يحافظ على تخطيط PDF الكامل.',
    convertPrompt: 'ارفع ملف DOCX او PDF او TXT او MD او PNG او JPG لتحويله محليا.',
    converted: 'تم تحويل {file} الى {format}.',
    converting: 'جار التحويل في المتصفح...',
    couldNotAddImagePage: 'تعذر اضافة صفحة الصورة.',
    couldNotAddPdfPages: 'تعذر اضافة ملف PDF.',
    couldNotExport: 'تعذر تصدير هذا المستند.',
    couldNotOpen: 'تعذر فتح هذا الملف.',
    copy: 'نسخ',
    delete: 'حذف',
    deleteCell: 'حذف الخلية',
    deletePage: 'حذف الصفحة',
    deletePageNumber: 'حذف الصفحة {number}',
    deleteTable: 'حذف الجدول',
    drawing: 'الرسم',
    editor: 'المحرر',
    editorModes: 'اوضاع المحرر',
    eraser: 'ممحاة',
    export: 'تصدير',
    exportFileName: 'اسم ملف التصدير',
    exportFormat: 'صيغة التصدير',
    fileName: 'اسم الملف',
    files: 'الملفات',
    fillColor: 'لون التعبئة',
    fontFamily: 'نوع الخط',
    fontSizeInPixels: 'حجم الخط بالبكسل',
    fontSizeSlider: 'شريط حجم الخط',
    formatting: 'التنسيق',
    heading: 'عنوان',
    heading1: 'عنوان 1',
    heading2: 'عنوان 2',
    heading3: 'عنوان 3',
    history: 'السجل',
    insertTable: 'ادراج جدول',
    italic: 'مائل',
    language: 'اللغة',
    line: 'خط',
    move: 'نقل',
    numberedList: 'قائمة مرقمة',
    openFile: 'فتح ملف',
    page: 'صفحة {number}',
    paintFront: 'الرسم للامام',
    paragraph: 'فقرة',
    pdfDocxExtractedOnly: 'تصدير PDF الى DOCX يستخرج النص فقط.',
    paste: 'لصق',
    placeText: 'وضع نص',
    rectangle: 'مستطيل',
    redo: 'اعادة',
    resizeHeight: 'تغيير الارتفاع',
    resizeProportionally: 'تغيير الحجم بالتناسب',
    resizeWidth: 'تغيير العرض',
    rotate: 'تدوير',
    selectedToBack: 'المحدد للخلف',
    selectedToFront: 'المحدد للامام',
    selectPaint: 'تحديد الرسم',
    selectText: 'تحديد النص',
    setFontSize: 'تعيين حجم الخط الى {size}px',
    strokeColor: 'لون الخط',
    targetFormat: 'الصيغة المطلوبة',
    textColor: 'لون النص',
    textFront: 'النص للامام',
    textGap: 'تباعد النص',
    textSize: 'حجم النص',
    textTools: 'ادوات النص',
    underline: 'تحته خط',
    undo: 'تراجع',
    unavailableConversion: 'هذا التحويل غير متاح في المتصفح.',
    unsavedWarning: 'هل انت متاكد من اعادة التحميل؟ ستفقد كل التغييرات غير المحفوظة.',
    write: 'كتابة',
    xShape: 'X',
  },
} satisfies Record<AppLanguage, Record<string, string>>
type TranslationKey = keyof typeof translations.en
type Translate = (key: TranslationKey, values?: Record<string, string | number>) => string

function translate(language: AppLanguage, key: TranslationKey, values: Record<string, string | number> = {}) {
  const text = translations[language][key] ?? translations.en[key]
  return Object.entries(values).reduce((currentText, [name, value]) => currentText.replaceAll(`{${name}}`, String(value)), text)
}

function localizedErrorMessage(error: unknown, t: Translate, fallbackKey: TranslationKey) {
  const message = error instanceof Error ? error.message : typeof error === 'string' ? error : ''
  const key = errorMessageTranslationKey(message)
  return key ? t(key) : message || t(fallbackKey)
}

function errorMessageTranslationKey(message: string): TranslationKey | null {
  if (message === 'Canvas export is not available in this browser.') return 'canvasExportUnavailable'
  if (message === 'Canvas conversion is not available in this browser.') return 'canvasConversionUnavailable'
  if (message === 'Canvas rendering is not available in this browser.') return 'canvasRenderUnavailable'
  if (message === 'Could not export the canvas.') return 'couldNotExportCanvas'
  if (message === 'Could not load that image.') return 'couldNotLoadImage'
  if (message === 'Could not read that file.') return 'couldNotReadFile'
  if (message === 'Image to text, Markdown, or DOCX needs OCR, which is not included in this client-side build.') return 'imageOcrUnavailable'
  if (message === 'Unsupported source file. Use DOCX, PDF, TXT, MD, HTML, PNG, or JPG.') return 'unsupportedSourceFile'
  return null
}
const pageSizes: Record<
  PageSizeKey,
  { label: string; width: number; height: number; docxWidth: number; docxHeight: number; orientation: 'portrait' | 'landscape' }
> = {
  a4: { label: 'A4', width: 794, height: 1123, docxWidth: 11906, docxHeight: 16838, orientation: 'portrait' },
  a5: { label: 'A5', width: 559, height: 794, docxWidth: 8391, docxHeight: 11906, orientation: 'portrait' },
  a4Landscape: { label: 'A4 landscape', width: 1123, height: 794, docxWidth: 16838, docxHeight: 11906, orientation: 'landscape' },
  a5Landscape: { label: 'A5 landscape', width: 794, height: 559, docxWidth: 11906, docxHeight: 8391, orientation: 'landscape' },
}
const defaultPageSizeKey: PageSizeKey = 'a4'
const defaultPageSize = pageSizes[defaultPageSizeKey]
const defaultPageMargins: PageMargins = { top: 64, right: 72, bottom: 72, left: 72 }

const StackOrder = Mark.create({
  name: 'stackOrder',
  inclusive: false,

  addAttributes() {
    return {
      order: {
        default: 'front',
        parseHTML: (element) => (element.getAttribute('data-stack-order') === 'back' ? 'back' : 'front'),
        renderHTML: (attributes) => ({
          'data-stack-order': attributes.order === 'back' ? 'back' : 'front',
        }),
      },
    }
  },

  parseHTML() {
    return [{ tag: 'span[data-stack-order]' }]
  },

  renderHTML({ HTMLAttributes }) {
    return ['span', mergeAttributes(HTMLAttributes), 0]
  },
})

const starterContent = '<p></p>'

function App() {
  const [language, setLanguage] = useState<AppLanguage>('en')
  const [mode, setMode] = useState<Mode>('editor')
  const [pages, setPages] = useState<PageData[]>([{ id: createId(), html: starterContent, images: [], size: defaultPageSizeKey }])
  const [activePageId, setActivePageId] = useState(() => pages[0].id)
  const activePageIdRef = useRef(activePageId)
  const loadingPage = useRef(false)
  const pagesRef = useRef<HTMLDivElement>(null)
  const pageDataRef = useRef(pages)
  const pageHistory = useRef<PageData[][]>([])
  const pageFuture = useRef<PageData[][]>([])
  const hasUnsavedChanges = useRef(false)
  const t = useCallback<Translate>((key, values) => translate(language, key, values), [language])
  const unsavedWarning = t('unsavedWarning')

  const updateActivePageHtml = useCallback((html: string) => {
    hasUnsavedChanges.current = true
    const pageId = activePageIdRef.current
    setPages((currentPages) => currentPages.map((page) => (page.id === pageId ? { ...page, html } : page)))
  }, [])

  const editor = useEditor({
    extensions: [
      StarterKit,
      Underline,
      StackOrder,
      TextStyle,
      FontSize,
      Color,
      FontFamily,
      TextAlign.configure({ types: ['heading', 'paragraph'] }),
      Table.configure({ resizable: true }),
      TableRow,
      TableHeader,
      TableCell,
    ],
    content: starterContent,
    onUpdate: ({ editor: currentEditor }) => {
      if (!loadingPage.current) updateActivePageHtml(currentEditor.getHTML())
    },
  })

  useEffect(() => {
    activePageIdRef.current = activePageId
  }, [activePageId])

  useEffect(() => {
    pageDataRef.current = pages
  }, [pages])

  useEffect(() => {
    const confirmBeforeUnload = (event: BeforeUnloadEvent) => {
      if (!hasUnsavedChanges.current) return
      event.preventDefault()
      event.returnValue = unsavedWarning
      return unsavedWarning
    }

    window.addEventListener('beforeunload', confirmBeforeUnload)
    return () => window.removeEventListener('beforeunload', confirmBeforeUnload)
  }, [unsavedWarning])

  const markUnsaved = useCallback(() => {
    hasUnsavedChanges.current = true
  }, [])

  const restorePages = useCallback(
    (nextPages: PageData[]) => {
      const restoredPages = clonePages(nextPages)
      const nextActivePageId = restoredPages.some((page) => page.id === activePageIdRef.current) ? activePageIdRef.current : restoredPages[0]?.id
      if (!nextActivePageId) return false
      loadingPage.current = true
      setPages(restoredPages)
      setActivePageId(nextActivePageId)
      const nextPage = restoredPages.find((page) => page.id === nextActivePageId)
      if (nextPage) editor?.commands.setContent(nextPage.html)
      window.setTimeout(() => {
        loadingPage.current = false
        scrollToPage(nextActivePageId)
      }, 0)
      return true
    },
    [editor],
  )

  const recordPagesHistory = useCallback(() => {
    pageHistory.current.push(clonePages(pageDataRef.current))
    pageHistory.current = pageHistory.current.slice(-50)
    pageFuture.current = []
  }, [])

  const undoPagesChange = useCallback(() => {
    const previousPages = pageHistory.current.pop()
    if (!previousPages) return false
    pageFuture.current.push(clonePages(pageDataRef.current))
    return restorePages(previousPages)
  }, [restorePages])

  const redoPagesChange = useCallback(() => {
    const nextPages = pageFuture.current.pop()
    if (!nextPages) return false
    pageHistory.current.push(clonePages(pageDataRef.current))
    return restorePages(nextPages)
  }, [restorePages])

  useEffect(() => {
    const nextPage = pageDataRef.current.find((page) => page.id === activePageId)
    if (!editor || !nextPage) return
    loadingPage.current = true
    editor.commands.setContent(nextPage.html)
    window.setTimeout(() => {
      loadingPage.current = false
    }, 0)
  }, [activePageId, editor])

  const addPage = (size: PageSizeKey = defaultPageSizeKey) => {
    markUnsaved()
    const nextPage = { id: createId(), html: '<p></p>', images: [], size }
    setPages((currentPages) => [...currentPages, nextPage])
    setActivePageId(nextPage.id)
    window.setTimeout(() => scrollToPage(nextPage.id), 0)
  }

  const addImagePage = async (file: File) => {
    markUnsaved()
    const nextPage = { id: createId(), html: '<p></p>', images: [], pageImage: await fileToDataUrl(file), size: defaultPageSizeKey }
    setPages((currentPages) => [...currentPages, nextPage])
    setActivePageId(nextPage.id)
    editor?.commands.setContent(nextPage.html)
    window.setTimeout(() => scrollToPage(nextPage.id), 0)
  }

  const addPdfPages = async (file: File) => {
    const nextPages = await pagesFromPdf(file)
    const firstPageId = nextPages[0]?.id
    if (!firstPageId) return
    recordPagesHistory()
    markUnsaved()
    setPages((currentPages) => [...currentPages, ...nextPages])
    setActivePageId(firstPageId)
    editor?.commands.setContent(nextPages[0].html)
    window.setTimeout(() => scrollToPage(firstPageId), 0)
  }

  const fillPageImageWithBucket = async (pageId: string, x: number, y: number, fillColor: string) => {
    const page = pageDataRef.current.find((item) => item.id === pageId)
    if (!page?.pageImage) return false
    const size = pageSizeFor(page.size)
    const nextPageImage = await floodFillImageDataUrl(page.pageImage, x, y, size.width, size.height, fillColor)
    if (!nextPageImage) return false
    if (nextPageImage === page.pageImage) return true
    recordPagesHistory()
    markUnsaved()
    setPages((currentPages) => currentPages.map((item) => (item.id === pageId ? { ...item, pageImage: nextPageImage } : item)))
    return true
  }

  const deletePage = (pageId: string) => {
    if (pages.length <= 1) return
    markUnsaved()
    const pageIndex = pages.findIndex((page) => page.id === pageId)
    const nextPages = pages.filter((page) => page.id !== pageId)
    const nextActivePageId =
      activePageId === pageId ? nextPages[Math.max(0, Math.min(pageIndex, nextPages.length - 1))].id : activePageId
    setPages(nextPages)
    setActivePageId(nextActivePageId)
    window.setTimeout(() => scrollToPage(nextActivePageId), 0)
  }

  const reorderPage = (draggedPageId: string, targetPageId: string) => {
    if (draggedPageId === targetPageId) return
    markUnsaved()
    setPages((currentPages) => {
      const fromIndex = currentPages.findIndex((page) => page.id === draggedPageId)
      const toIndex = currentPages.findIndex((page) => page.id === targetPageId)
      if (fromIndex < 0 || toIndex < 0 || fromIndex === toIndex) return currentPages
      const nextPages = [...currentPages]
      const [draggedPage] = nextPages.splice(fromIndex, 1)
      nextPages.splice(toIndex, 0, draggedPage)
      return nextPages
    })
  }

  const updateImage = (pageId: string, imageId: string, patch: Partial<FloatingImage>) => {
    markUnsaved()
    setPages((currentPages) =>
      currentPages.map((page) =>
        page.id === pageId
          ? {
              ...page,
              images: page.images.map((image) => (image.id === imageId ? { ...image, ...patch } : image)),
            }
          : page,
      ),
    )
  }

  const deleteImage = (pageId: string, imageId: string) => {
    markUnsaved()
    setPages((currentPages) =>
      currentPages.map((page) => (page.id === pageId ? { ...page, images: page.images.filter((image) => image.id !== imageId) } : page)),
    )
  }

  const moveImageLayer = (pageId: string, imageId: string, direction: 'front' | 'back') => {
    markUnsaved()
    setPages((currentPages) =>
      currentPages.map((page) => {
        if (page.id !== pageId) return page
        const image = page.images.find((item) => item.id === imageId)
        if (!image) return page
        const remaining = page.images.filter((item) => item.id !== imageId)
        return { ...page, images: direction === 'front' ? [...remaining, image] : [image, ...remaining] }
      }),
    )
  }

  const addImageToActivePage = (src: string, width = 260, height = 180) => {
    markUnsaved()
    const pageSize = pageSizeFor(pageDataRef.current.find((page) => page.id === activePageId)?.size)
    setPages((currentPages) =>
      currentPages.map((page) =>
        page.id === activePageId
          ? {
              ...page,
              images: [
                ...page.images,
                {
                  id: createId(),
                  src,
                  x: 120,
                  y: 130,
                  width: Math.min(width, pageSize.width - 240),
                  height: Math.min(height, pageSize.height - 260),
                  rotation: 0,
                },
              ],
            }
          : page,
      ),
    )
  }

  const openDocument = async (file: File) => {
    markUnsaved()
    const nextPages = await pagesFromFile(file)
    const firstPageId = nextPages[0].id
    loadingPage.current = true
    setPages(nextPages)
    setActivePageId(firstPageId)
    editor?.commands.setContent(nextPages[0].html)
    window.setTimeout(() => {
      loadingPage.current = false
      scrollToPage(firstPageId)
    }, 0)
    return nextPages
  }

  return (
    <main className="app-shell" lang={language}>
      <Sidebar mode={mode} setMode={setMode} t={t} />
      <section className="workspace">
        <header className="topbar">
          <h1>AKXREDITOR</h1>
          {mode !== 'editor' && (
            <label className="language-picker">
              <span>{t('language')}</span>
              <select aria-label={t('language')} onChange={(event) => setLanguage(event.target.value as AppLanguage)} value={language}>
                {languages.map((item) => (
                  <option key={item.code} value={item.code}>
                    {item.label}
                  </option>
                ))}
              </select>
            </label>
          )}
        </header>

        {mode === 'editor' && (
          <EditorMode
            activePageId={activePageId}
            addImageToActivePage={addImageToActivePage}
            addImagePage={addImagePage}
            addPage={addPage}
            addPdfPages={addPdfPages}
            deletePage={deletePage}
            editor={editor}
            pages={pages}
            pagesRef={pagesRef}
            openDocument={openDocument}
            recordPagesHistory={recordPagesHistory}
            markUnsaved={markUnsaved}
            redoPagesChange={redoPagesChange}
            reorderPage={reorderPage}
            setActivePageId={setActivePageId}
            undoPagesChange={undoPagesChange}
            updateImage={updateImage}
            deleteImage={deleteImage}
            fillPageImageWithBucket={fillPageImageWithBucket}
            moveImageLayer={moveImageLayer}
            language={language}
            setLanguage={setLanguage}
            t={t}
          />
        )}
        {mode === 'convert' && <ConvertMode t={t} />}
      </section>
    </main>
  )
}

function Sidebar({ mode, setMode, t }: { mode: Mode; setMode: (mode: Mode) => void; t: Translate }) {
  const items = [
    { id: 'editor' as const, label: t('editor'), icon: FileText },
    { id: 'convert' as const, label: t('convert'), icon: FileArchive },
  ]

  return (
    <aside className="sidebar">
      <div className="brand-mark">AX</div>
      <nav className="mode-tabs" aria-label={t('editorModes')}>
        {items.map((item) => {
          const Icon = item.icon
          return (
            <button
              className={mode === item.id ? 'active' : ''}
              key={item.id}
              onClick={() => setMode(item.id)}
              type="button"
            >
              <Icon size={18} />
              <span>{item.label}</span>
            </button>
          )
        })}
      </nav>
    </aside>
  )
}

function EditorMode({
  activePageId,
  addImageToActivePage,
  addImagePage,
  addPage,
  addPdfPages,
  deletePage,
  editor,
  pages,
  pagesRef,
  openDocument,
  recordPagesHistory,
  markUnsaved,
  redoPagesChange,
  reorderPage,
  setActivePageId,
  undoPagesChange,
  updateImage,
  deleteImage,
  fillPageImageWithBucket,
  moveImageLayer,
  language,
  setLanguage,
  t,
}: {
  activePageId: string
  addImageToActivePage: (src: string, width?: number, height?: number) => void
  addImagePage: (file: File) => Promise<void>
  addPage: (size?: PageSizeKey) => void
  addPdfPages: (file: File) => Promise<void>
  deletePage: (pageId: string) => void
  editor: Editor | null
  pages: PageData[]
  pagesRef: React.RefObject<HTMLDivElement | null>
  openDocument: (file: File) => Promise<PageData[]>
  recordPagesHistory: () => void
  markUnsaved: () => void
  redoPagesChange: () => boolean
  reorderPage: (draggedPageId: string, targetPageId: string) => void
  setActivePageId: (pageId: string) => void
  undoPagesChange: () => boolean
  updateImage: (pageId: string, imageId: string, patch: Partial<FloatingImage>) => void
  deleteImage: (pageId: string, imageId: string) => void
  fillPageImageWithBucket: (pageId: string, x: number, y: number, fillColor: string) => Promise<boolean>
  moveImageLayer: (pageId: string, imageId: string, direction: 'front' | 'back') => void
  language: AppLanguage
  setLanguage: (language: AppLanguage) => void
  t: Translate
}) {
  const openInputRef = useRef<HTMLInputElement>(null)
  const imageInputRef = useRef<HTMLInputElement>(null)
  const imagePageInputRef = useRef<HTMLInputElement>(null)
  const pdfPagesInputRef = useRef<HTMLInputElement>(null)
  const canvasMap = useRef(new Map<string, FabricCanvas>())
  const histories = useRef(new Map<string, string[]>())
  const futures = useRef(new Map<string, string[]>())
  const loadingCanvas = useRef(false)
  const actionHistory = useRef<UndoKind[]>([])
  const redoHistory = useRef<UndoKind[]>([])
  const fabricClipboard = useRef<FabricObject | null>(null)
  const [tool, setTool] = useState<EditorTool>('write')
  const [layerOrder, setLayerOrder] = useState<LayerKind[]>(['text', 'images', 'paint'])
  const [textColor, setTextColor] = useState('#171224')
  const [textFontFamily, setTextFontFamily] = useState(fontFamilies[0])
  const [textFontSize, setTextFontSize] = useState(defaultFontSize)
  const [textLineGap, setTextLineGap] = useState(1.72)
  const [exportFileName, setExportFileName] = useState('akxreditor-document')
  const [textBold, setTextBold] = useState(false)
  const [textItalic, setTextItalic] = useState(false)
  const [textUnderline, setTextUnderline] = useState(false)
  const [stroke, setStroke] = useState('#6d5dfc')
  const [fill, setFill] = useState('#d93f75')
  const [brushSize, setBrushSize] = useState(8)
  const [selectedImageId, setSelectedImageId] = useState<string | null>(null)
  const [selectedTarget, setSelectedTarget] = useState<SelectedTarget | null>(null)
  const [pendingImageExport, setPendingImageExport] = useState<PendingImageExport | null>(null)
  const [contextMenu, setContextMenu] = useState<ContextMenuState | null>(null)
  const [hasFabricClipboard, setHasFabricClipboard] = useState(false)
  const [draggedPageId, setDraggedPageId] = useState<string | null>(null)
  const [dragOverPageId, setDragOverPageId] = useState<string | null>(null)
  const layerZ = {
    text: layerOrder.indexOf('text') + 1,
    images: layerOrder.indexOf('images') + 1,
    paint: layerOrder.indexOf('paint') + 1,
  }
  const relativeLayerZ = {
    images: layerOrder.indexOf('images') - layerOrder.indexOf('text'),
    paint: layerOrder.indexOf('paint') - layerOrder.indexOf('text'),
  }
  const canSelectPaint = tool === 'select-paint'
  const canSelectText = tool === 'select-text'

  const moveLayer = (layer: LayerKind, direction: 'front' | 'back') => {
    setLayerOrder((currentOrder) => {
      const remaining = currentOrder.filter((item) => item !== layer)
      return direction === 'front' ? [...remaining, layer] : [layer, ...remaining]
    })
  }

  const moveLayerToFront = (layer: LayerKind) => moveLayer(layer, 'front')

  const activateTool = (nextTool: EditorTool) => {
    setTool(nextTool)
    if (nextTool === 'brush' || nextTool === 'eraser' || nextTool === 'bucket' || nextTool === 'place-text') moveLayerToFront('paint')
  }

  const toggleTextPlacement = () => {
    activateTool(tool === 'place-text' ? 'select-paint' : 'place-text')
  }

  const createPaintText = (text: string, left: number, top: number) =>
    new IText(text, {
      fill: textColor,
      fontFamily: textFontFamily,
      fontSize: Number.parseFloat(textFontSize),
      fontStyle: textItalic ? 'italic' : 'normal',
      fontWeight: textBold ? 'bold' : 'normal',
      left: Math.round(left),
      lineHeight: textLineGap,
      noScaleCache: true,
      objectCaching: false,
      top: Math.round(top),
      underline: textUnderline,
    })

  const textStyleVars = {
    '--editor-line-height': textLineGap,
    '--editor-text-color': textColor,
  } as React.CSSProperties

  const activatePage = (pageId: string) => {
    setSelectedImageId(null)
    setSelectedTarget(null)
    setContextMenu(null)
    setActivePageId(pageId)
    window.setTimeout(() => scrollToPage(pageId), 0)
  }

  const rememberAction = useCallback((kind: UndoKind) => {
    markUnsaved()
    actionHistory.current.push(kind)
    redoHistory.current = []
  }, [markUnsaved])

  const rememberTextAction = () => rememberAction('text')

  const rememberPagesAction = () => {
    recordPagesHistory()
    rememberAction('pages')
  }

  const rememberPageImageBucketFill = () => {
    actionHistory.current.push('pages')
    redoHistory.current = []
  }

  const runTextCommand = (command: () => void) => {
    rememberTextAction()
    command()
  }

  const handleSelectDrawing = useCallback((pageId: string, tableCellIndex?: number) => {
    setSelectedImageId(null)
    setSelectedTarget(tableCellIndex == null ? { kind: 'paint', pageId } : { kind: 'table-cell', pageId, cellIndex: tableCellIndex })
  }, [])

  const registerCanvas = useCallback((pageId: string, canvas: FabricCanvas | null) => {
    if (canvas) {
      canvasMap.current.set(pageId, canvas)
      if (!histories.current.has(pageId)) {
        histories.current.set(pageId, [canvasSnapshot(canvas)])
        futures.current.set(pageId, [])
      }
    } else {
      canvasMap.current.delete(pageId)
    }
  }, [])

  const recordCanvasHistory = useCallback((pageId: string, canvas: FabricCanvas) => {
    if (loadingCanvas.current) return
    const history = histories.current.get(pageId) ?? []
    history.push(canvasSnapshot(canvas))
    histories.current.set(pageId, history.slice(-50))
    futures.current.set(pageId, [])
    actionHistory.current.push('paint')
    redoHistory.current = []
  }, [])

  const loadCanvasSnapshot = (pageId: string, snapshot: string) => {
    const canvas = canvasMap.current.get(pageId)
    if (!canvas) return
    loadingCanvas.current = true
    void canvas.loadFromJSON(JSON.parse(snapshot)).then(() => {
      restoreManagedCanvasObjects(canvas)
      canvas.requestRenderAll()
      loadingCanvas.current = false
    })
  }

  const activeCanvas = () => canvasMap.current.get(activePageId)

  const selectedPaintText = () => {
    const object = activeCanvas()?.getActiveObject()
    return object instanceof IText || isHorizontalTableGroup(object)
  }

  const clearCanvasSelection = (pageId = activePageId) => {
    const canvas = canvasMap.current.get(pageId)
    if (!canvas) return
    canvas.discardActiveObject()
    canvas.requestRenderAll()
  }

  const addFabricObject = (object: FabricObject) => {
    const canvas = activeCanvas()
    if (!canvas) return
    setTool('select-paint')
    moveLayerToFront('paint')
    setSelectedImageId(null)
    setSelectedTarget({ kind: 'paint', pageId: activePageId })
    canvas.add(object)
    canvas.setActiveObject(object)
    canvas.requestRenderAll()
  }

  const addPageWithHistory = (size: PageSizeKey = defaultPageSizeKey) => {
    rememberPagesAction()
    addPage(size)
  }

  const addImagePageWithHistory = (file: File) => {
    rememberPagesAction()
    return addImagePage(file)
  }

  const deletePageWithHistory = (pageId: string) => {
    rememberPagesAction()
    deletePage(pageId)
  }

  const updateImageWithHistory = (pageId: string, imageId: string, patch: Partial<FloatingImage>) => {
    updateImage(pageId, imageId, patch)
  }

  const moveImageLayerWithHistory = (pageId: string, imageId: string, direction: 'front' | 'back') => {
    rememberPagesAction()
    moveImageLayer(pageId, imageId, direction)
  }

  const deleteImageWithHistory = (pageId: string, imageId: string) => {
    rememberPagesAction()
    deleteImage(pageId, imageId)
  }

  const updateSelectedPaintTextStyle = (
    patch: Partial<{
      fill: string
      fontFamily: string
      fontSize: number
      fontStyle: 'normal' | 'italic'
      fontWeight: 'normal' | 'bold'
      lineHeight: number
      underline: boolean
    }>,
  ) => {
    const canvas = activeCanvas()
    const object = canvas?.getActiveObject()
    if (!canvas || (!isHorizontalTableGroup(object) && !(object instanceof IText))) return
    if (isHorizontalTableGroup(object)) {
      object.getObjects().forEach((item) => {
        if (item instanceof IText) {
          item.set(patch)
          sharpenCanvasText(item)
          item.initDimensions()
        }
      })
    } else {
      object.set(patch)
      sharpenCanvasText(object)
      object.initDimensions()
    }
    object.setCoords()
    canvas.requestRenderAll()
    recordCanvasHistory(activePageId, canvas)
  }

  const updateTextColor = (nextColor: string) => {
    setTextColor(nextColor)
    updateSelectedPaintTextStyle({ fill: nextColor })
    if (!selectedPaintText() && editor) runTextCommand(() => editor.chain().focus().selectAll().setColor(nextColor).run())
  }

  const updateTextFontFamily = (nextFamily: string) => {
    setTextFontFamily(nextFamily)
    updateSelectedPaintTextStyle({ fontFamily: nextFamily })
    if (!selectedPaintText()) runTextCommand(() => editor?.chain().focus().setFontFamily(nextFamily).run())
  }

  const updateTextFontSize = (nextSize: string) => {
    const normalizedSize = normalizeFontSize(nextSize)
    if (!normalizedSize) return
    setTextFontSize(normalizedSize)
    updateSelectedPaintTextStyle({ fontSize: Number.parseFloat(normalizedSize) })
    if (!selectedPaintText()) runTextCommand(() => editor?.chain().focus().setFontSize(normalizedSize).run())
  }

  const updateTextLineGap = (nextGap: number) => {
    setTextLineGap(nextGap)
    updateSelectedPaintTextStyle({ lineHeight: nextGap })
  }

  const toggleTextBold = () => {
    const nextBold = !textBold
    setTextBold(nextBold)
    updateSelectedPaintTextStyle({ fontWeight: nextBold ? 'bold' : 'normal' })
    if (!selectedPaintText()) runTextCommand(() => editor?.chain().focus().toggleBold().run())
  }

  const toggleTextItalic = () => {
    const nextItalic = !textItalic
    setTextItalic(nextItalic)
    updateSelectedPaintTextStyle({ fontStyle: nextItalic ? 'italic' : 'normal' })
    if (!selectedPaintText()) runTextCommand(() => editor?.chain().focus().toggleItalic().run())
  }

  const toggleTextUnderline = () => {
    const nextUnderline = !textUnderline
    setTextUnderline(nextUnderline)
    updateSelectedPaintTextStyle({ underline: nextUnderline })
    if (!selectedPaintText()) runTextCommand(() => editor?.chain().focus().toggleUnderline().run())
  }

  const addMovableText = (text: string, mode: MovableTextMode = 'plain') => {
    const canvas = activeCanvas()
    if (!canvas) return
    const position = nextMovableTextPosition(canvas)
    const object = createPaintText(text, position.left, position.top)
    configureManagedText(object, mode, canvas, { removeIfEmpty: mode === 'plain' })
    addFabricObject(object)
    object.enterEditing()
    if (mode === 'plain') {
      object.selectAll()
    } else {
      moveTextCursorToEnd(object)
    }
    canvas.requestRenderAll()
  }

  const addHorizontalTable = () => {
    const canvas = activeCanvas()
    if (!canvas) return
    const position = nextMovableTextPosition(canvas)
    const pageSize = pageSizeFor(pages.find((page) => page.id === activePageId)?.size)
    const table = createHorizontalTableGroup({
      cellWidth: Math.max(84, Math.min(150, (pageSize.width - position.left - 34) / 3)),
      cellCount: 3,
      fill: textColor,
      fontFamily: textFontFamily,
      fontSize: Number.parseFloat(textFontSize),
      fontStyle: textItalic ? 'italic' : 'normal',
      fontWeight: textBold ? 'bold' : 'normal',
      left: position.left,
      top: position.top,
      underline: textUnderline,
    })
    configureHorizontalTableGroup(table, canvas)
    addFabricObject(table)
  }

  const moveSelectedObjectLayer = (direction: 'front' | 'back') => {
    setContextMenu(null)
    if (selectedTarget?.kind === 'image') {
      moveImageLayerWithHistory(selectedTarget.pageId, selectedTarget.imageId, direction)
      return
    }
    if (selectedImageId) {
      moveImageLayerWithHistory(activePageId, selectedImageId, direction)
      return
    }
    const targetPageId = selectedTarget?.kind === 'paint' ? selectedTarget.pageId : activePageId
    const canvas = canvasMap.current.get(targetPageId)
    const object = canvas?.getActiveObject()
    if (canvas && object) {
      if (direction === 'front') {
        canvas.bringObjectToFront(object)
      } else {
        canvas.sendObjectToBack(object)
      }
      canvas.requestRenderAll()
      recordCanvasHistory(targetPageId, canvas)
      return
    }
    if (selectedTarget?.kind === 'text' || (editor && !editor.state.selection.empty)) {
      if (editor && !editor.state.selection.empty) {
        rememberTextAction()
        editor.chain().focus().setMark('stackOrder', { order: direction }).run()
      }
    }
  }

  const deleteSelectedTableCell = () => {
    setContextMenu(null)
    if (selectedTarget?.kind !== 'table-cell') return
    const canvas = canvasMap.current.get(selectedTarget.pageId)
    const object = canvas?.getActiveObject()
    if (!canvas || !isHorizontalTableGroup(object)) return
    deleteHorizontalTableCell(object, selectedTarget.cellIndex, canvas)
    canvas.requestRenderAll()
    recordCanvasHistory(selectedTarget.pageId, canvas)
    setSelectedTarget(null)
  }

  const deleteSelectedObject = () => {
    setContextMenu(null)
    if (selectedTarget?.kind === 'table-cell') {
      deleteSelectedTableCell()
      return
    }
    if (selectedTarget?.kind === 'image') {
      deleteImageWithHistory(selectedTarget.pageId, selectedTarget.imageId)
      setSelectedImageId(null)
      setSelectedTarget(null)
      return
    }
    if (selectedImageId) {
      deleteImageWithHistory(activePageId, selectedImageId)
      setSelectedImageId(null)
      setSelectedTarget(null)
      return
    }
    const targetPageId = selectedTarget?.kind === 'paint' ? selectedTarget.pageId : activePageId
    const canvas = canvasMap.current.get(targetPageId)
    const object = canvas?.getActiveObject()
    if (canvas && object) {
      removeCanvasObject(canvas, object)
      canvas.discardActiveObject()
      canvas.requestRenderAll()
      recordCanvasHistory(targetPageId, canvas)
      setSelectedTarget(null)
      return
    }
    if (selectedTarget?.kind === 'text' && editor && !editor.state.selection.empty) {
      rememberTextAction()
      editor.chain().focus().deleteSelection().run()
      setSelectedTarget(null)
    }
  }

  const copySelectedCanvasObject = async () => {
    const targetPageId = selectedTarget?.kind === 'paint' || selectedTarget?.kind === 'table-cell' ? selectedTarget.pageId : activePageId
    const canvas = canvasMap.current.get(targetPageId)
    const object = canvas?.getActiveObject()
    if (!canvas || !object) return false
    fabricClipboard.current = await object.clone(canvasObjectCustomProperties)
    setHasFabricClipboard(true)
    setContextMenu(null)
    return true
  }

  const pasteCanvasObject = async (pageId = activePageId, pastePoint?: CanvasPoint) => {
    const canvas = canvasMap.current.get(pageId)
    const source = fabricClipboard.current
    if (!canvas || !source) return false
    const clone = await source.clone(canvasObjectCustomProperties)
    const pastedObjects = positionClonedCanvasObject(clone, pastePoint)
    if (!pastedObjects.length) return false

    loadingCanvas.current = true
    canvas.discardActiveObject()
    pastedObjects.forEach((object) => {
      restorePastedCanvasObject(object, canvas)
      keepObjectInsideCanvas(object, canvas)
      canvas.add(object)
    })
    loadingCanvas.current = false

    if (pastedObjects.length === 1) {
      canvas.setActiveObject(pastedObjects[0])
    } else {
      canvas.setActiveObject(new ActiveSelection(pastedObjects, { canvas } as Partial<ConstructorParameters<typeof ActiveSelection>[1]>))
    }
    setTool('select-paint')
    moveLayerToFront('paint')
    setActivePageId(pageId)
    setSelectedImageId(null)
    setSelectedTarget({ kind: 'paint', pageId })
    canvas.requestRenderAll()
    recordCanvasHistory(pageId, canvas)
    setContextMenu(null)
    return true
  }

  const copyFromContextMenu = () => {
    void copySelectedCanvasObject()
  }

  const pasteFromContextMenu = () => {
    if (!contextMenu || !hasFabricClipboard) return
    void pasteCanvasObject(contextMenu.target.pageId, contextMenu.pastePoint)
  }

  const openContextMenu = (event: MouseEvent | React.MouseEvent, target: ContextMenuTarget, pastePoint?: CanvasPoint) => {
    event.preventDefault()
    event.stopPropagation()
    if (target.kind === 'page' && !fabricClipboard.current) {
      setContextMenu(null)
      return
    }
    setActivePageId(target.pageId)
    if (target.kind === 'page') {
      setSelectedTarget(null)
      setSelectedImageId(null)
      clearCanvasSelection(target.pageId)
    } else {
      setSelectedTarget(target)
    }
    setContextMenu({
      pastePoint,
      target,
      x: event.clientX,
      y: event.clientY,
    })
  }

  const savePagesAsImageFiles = async (format: 'png' | 'jpg', baseName: string) => {
    if (pages.length <= 1) {
      saveBlob(await exportPagesImageBlob(pages, format, (pageId) => canvasMap.current.get(pageId)?.toDataURL()), exportFileNameFor(baseName, format))
      return
    }

    const imageBlobs = await exportPageImageBlobs(pages, format, (pageId) => canvasMap.current.get(pageId)?.toDataURL())
    imageBlobs.forEach((blob, index) => {
      saveBlob(blob, exportImagePageFileNameFor(baseName, index + 1, format))
    })
  }

  const continueImageExport = () => {
    if (!pendingImageExport) return
    const { baseName, format } = pendingImageExport
    setPendingImageExport(null)
    void savePagesAsImageFiles(format, baseName).catch((error) => {
      window.alert(localizedErrorMessage(error, t, 'couldNotExport'))
    })
  }

  const exportDocument = async (format: ExportFormat, requestedBaseName = exportFileName) => {
    if (!pagesRef.current) return
    const baseName = safeFileBaseName(requestedBaseName)
    const html = pages.map((page) => page.html).join('<hr>')
    const text = pages.map((page, index) => `${t('page', { number: index + 1 })}\n${textFromHtml(page.html)}`).join('\n\n')

    if (format === 'txt') {
      saveBlob(new Blob([text], { type: 'text/plain;charset=utf-8' }), exportFileNameFor(baseName, format))
      return
    }
    if (format === 'md') {
      saveBlob(new Blob([markdownFromHtml(html)], { type: 'text/markdown;charset=utf-8' }), exportFileNameFor(baseName, format))
      return
    }
    if (format === 'html') {
      saveBlob(
        new Blob([wrapHtmlDocument(renderPagesHtml(pages, (pageId) => canvasMap.current.get(pageId)?.toDataURL()), baseName)], {
          type: 'text/html;charset=utf-8',
        }),
        exportFileNameFor(baseName, format),
      )
      return
    }
    if (format === 'docx') {
      saveBlob(await createDocxBlob(text, pages[0]?.size), exportFileNameFor(baseName, format))
      return
    }
    if (format === 'pdf') {
      saveBlob(await exportPagesPdfBlob(pages, (pageId) => canvasMap.current.get(pageId)?.toDataURL()), exportFileNameFor(baseName, format))
      return
    }
    if (pages.length > 1) {
      setPendingImageExport({ baseName, format, pageCount: pages.length })
      return
    }
    await savePagesAsImageFiles(format, baseName)
  }

  const insertImage = async (file: File) => {
    const src = await fileToDataUrl(file)
    const image = await loadImage(src)
    const maxWidth = 360
    const ratio = image.naturalHeight / image.naturalWidth
    rememberPagesAction()
    addImageToActivePage(src, maxWidth, maxWidth * ratio)
    moveLayerToFront('images')
    setTool('select-paint')
  }

  const undoCanvas = () => {
    const history = histories.current.get(activePageId) ?? []
    if (history.length < 2) return false
    const current = history.pop()
    if (current) futures.current.set(activePageId, [...(futures.current.get(activePageId) ?? []), current])
    const previous = history.at(-1)
    histories.current.set(activePageId, history)
    if (previous) loadCanvasSnapshot(activePageId, previous)
    return true
  }

  const redoCanvas = () => {
    const future = futures.current.get(activePageId) ?? []
    const next = future.pop()
    if (!next) return false
    histories.current.set(activePageId, [...(histories.current.get(activePageId) ?? []), next])
    futures.current.set(activePageId, future)
    loadCanvasSnapshot(activePageId, next)
    return true
  }

  useEffect(() => {
    const closeMenu = () => setContextMenu(null)
    const closeOnEscape = (event: KeyboardEvent) => {
      if (event.key === 'Escape') closeMenu()
    }
    window.addEventListener('click', closeMenu)
    window.addEventListener('scroll', closeMenu, true)
    window.addEventListener('keydown', closeOnEscape)
    return () => {
      window.removeEventListener('click', closeMenu)
      window.removeEventListener('scroll', closeMenu, true)
      window.removeEventListener('keydown', closeOnEscape)
    }
  }, [])

  const undoLast = () => {
    const kind = actionHistory.current.pop()
    if (!kind) return
    if (kind === 'paint') {
      if (undoCanvas()) redoHistory.current.push(kind)
      return
    }
    if (kind === 'pages') {
      if (undoPagesChange()) redoHistory.current.push(kind)
      return
    }
    editor?.chain().focus().undo().run()
    redoHistory.current.push(kind)
  }

  const redoLast = () => {
    const kind = redoHistory.current.pop()
    if (!kind) return
    if (kind === 'paint') {
      if (redoCanvas()) actionHistory.current.push(kind)
      return
    }
    if (kind === 'pages') {
      if (redoPagesChange()) actionHistory.current.push(kind)
      return
    }
    editor?.chain().focus().redo().run()
    actionHistory.current.push(kind)
  }

  useEffect(() => {
    const handleShortcut = (event: KeyboardEvent) => {
      const target = event.target
      const typingTarget = isTypingTarget(target)
      const isUndo = (event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'z' && !event.shiftKey
      const isRedo = (event.ctrlKey || event.metaKey) && (event.key.toLowerCase() === 'y' || (event.shiftKey && event.key.toLowerCase() === 'z'))
      const isCopy = (event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'c'
      const isPaste = (event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'v'

      if (isUndo) {
        event.preventDefault()
        undoLast()
        return
      }
      if (isRedo) {
        event.preventDefault()
        redoLast()
        return
      }
      if (isCopy && !typingTarget && activeCanvas()?.getActiveObject()) {
        event.preventDefault()
        void copySelectedCanvasObject()
        return
      }
      if (isPaste && !typingTarget && hasFabricClipboard) {
        event.preventDefault()
        void pasteCanvasObject(activePageId)
        return
      }
      if ((event.key === 'Delete' || event.key === 'Backspace') && !typingTarget && (selectedTarget || selectedImageId || activeCanvas()?.getActiveObject())) {
        event.preventDefault()
        deleteSelectedObject()
        return
      }
      if (event.key === 'Escape') {
        setContextMenu(null)
        setSelectedImageId(null)
        setSelectedTarget(null)
        clearCanvasSelection()
        if (tool === 'place-text' || tool === 'brush' || tool === 'eraser' || tool === 'bucket') setTool('select-paint')
      }
    }

    window.addEventListener('keydown', handleShortcut)
    return () => window.removeEventListener('keydown', handleShortcut)
  }, [activePageId, copySelectedCanvasObject, deleteSelectedObject, hasFabricClipboard, pasteCanvasObject, redoLast, selectedImageId, selectedTarget, tool, undoLast])

  const textFontSizeNumber = Number.parseFloat(textFontSize)
  const defaultFontSizeNumber = Number.parseFloat(defaultFontSize)
  const visibleTextFontSize = Number.isFinite(textFontSizeNumber) ? Math.round(textFontSizeNumber) : defaultFontSizeNumber
  const selectedFontSizePreset = fontSizes.includes(`${visibleTextFontSize}px`) ? `${visibleTextFontSize}px` : ''

  return (
    <section className="mode-panel" style={textStyleVars}>
      <div className="editor-command-strip">
        <ToolButton icon={Upload} label={t('openFile')} onClick={() => openInputRef.current?.click()} />
        <input
          accept=".docx,.pdf,.txt,.md,.markdown,.html,.htm,.png,.jpg,.jpeg"
          className="hidden-input"
          onChange={(event) => {
            const file = event.target.files?.[0]
            if (file) {
              void openDocument(file)
                .then((nextPages) => {
                  const source = extensionOf(file.name)
                  const openedOnlyImages = nextPages.every((page) => page.images.length > 0 && textFromHtml(page.html).trim().length === 0)
                  setExportFileName(baseNameFromFileName(file.name))
                  setTool(source === 'pdf' && !openedOnlyImages ? 'write' : openedOnlyImages ? 'select-paint' : 'write')
                })
                .catch((error) => {
                  window.alert(localizedErrorMessage(error, t, 'couldNotOpen'))
                })
            }
            event.currentTarget.value = ''
          }}
          ref={openInputRef}
          type="file"
        />
        <ToolButton icon={ImagePlus} label={t('addImage')} onClick={() => imageInputRef.current?.click()} />
        <input
          accept="image/*"
          className="hidden-input"
          onChange={(event) => {
            const file = event.target.files?.[0]
            if (file) void insertImage(file)
            event.currentTarget.value = ''
          }}
          ref={imageInputRef}
          type="file"
        />
        <label className="language-picker editor-language-picker">
          <span>{t('language')}</span>
          <select aria-label={t('language')} onChange={(event) => setLanguage(event.target.value as AppLanguage)} value={language}>
            {languages.map((item) => (
              <option key={item.code} value={item.code}>
                {item.label}
              </option>
            ))}
          </select>
        </label>
        <span className="command-spacer" />
        <IconButton icon={Undo2} label={t('undo')} onClick={undoLast} />
        <IconButton icon={Redo2} label={t('redo')} onClick={redoLast} />
      </div>

      <div className="editor-workbench">
        <aside className="editor-rail editor-rail-left" aria-label={t('textTools')}>
          <div className="rail-section">
            <ToolButton active={tool === 'write'} icon={Type} label={t('write')} onClick={() => activateTool('write')} />
            <ToolButton active={tool === 'place-text'} icon={Type} label={t('placeText')} onClick={toggleTextPlacement} />
            <ToolButton active={tool === 'select-text'} icon={Type} label={t('selectText')} onClick={() => activateTool('select-text')} />
          </div>
          <div className="rail-section">
            <select aria-label={t('fontFamily')} onChange={(event) => updateTextFontFamily(event.target.value)} value={textFontFamily}>
              {fontFamilies.map((family) => (
                <option key={family} value={family}>
                  {family}
                </option>
              ))}
            </select>
            <div className="font-size-control" title={t('textSize')}>
              <label className="font-size-field">
                <Type size={16} />
                <input aria-label={t('fontSizeInPixels')} max="240" min="1" onChange={(event) => updateTextFontSize(event.target.value)} step="1" type="number" value={visibleTextFontSize} />
                <span>px</span>
              </label>
              <select aria-label={t('commonFontSizes')} className="font-size-select" onChange={(event) => updateTextFontSize(event.target.value)} value={selectedFontSizePreset}>
                {!selectedFontSizePreset && <option value="">{visibleTextFontSize}px</option>}
                {fontSizes.map((size) => (
                  <option key={size} value={size}>
                    {Number.parseFloat(size)}px
                  </option>
                ))}
              </select>
            </div>
            <label className="color-chip" title={t('textColor')}>
              <Baseline size={16} />
              <input type="color" onChange={(event) => updateTextColor(event.target.value)} value={textColor} />
            </label>
            <label className="range-control" title={t('textGap')}>
              <Rows3 size={16} />
              <input min="1" max="2.6" onChange={(event) => updateTextLineGap(Number(event.target.value))} step="0.02" type="range" value={textLineGap} />
              <span>{textLineGap.toFixed(2)}</span>
            </label>
          </div>
          <div className="rail-section compact-icons" aria-label={t('formatting')}>
            <IconButton active={textBold || editor?.isActive('bold')} icon={Bold} label={t('bold')} onClick={toggleTextBold} />
            <IconButton active={textItalic || editor?.isActive('italic')} icon={Italic} label={t('italic')} onClick={toggleTextItalic} />
            <IconButton active={textUnderline || editor?.isActive('underline')} icon={UnderlineIcon} label={t('underline')} onClick={toggleTextUnderline} />
            <select
              aria-label={t('heading')}
              onChange={(event) => {
                const level = Number(event.target.value)
                runTextCommand(() => {
                  if (level === 0) editor?.chain().focus().setParagraph().run()
                  if (level > 0) editor?.chain().focus().toggleHeading({ level: level as 1 | 2 | 3 }).run()
                })
              }}
            >
              <option value="0">{t('paragraph')}</option>
              <option value="1">{t('heading1')}</option>
              <option value="2">{t('heading2')}</option>
              <option value="3">{t('heading3')}</option>
            </select>
            <IconButton icon={AlignLeft} label={t('alignLeft')} onClick={() => runTextCommand(() => editor?.chain().focus().setTextAlign('left').run())} />
            <IconButton icon={AlignCenter} label={t('alignCenter')} onClick={() => runTextCommand(() => editor?.chain().focus().setTextAlign('center').run())} />
            <IconButton icon={AlignRight} label={t('alignRight')} onClick={() => runTextCommand(() => editor?.chain().focus().setTextAlign('right').run())} />
            <IconButton icon={List} label={t('bulletList')} onClick={() => addMovableText('- ', 'bullet')} />
            <IconButton icon={ListOrdered} label={t('numberedList')} onClick={() => addMovableText('1. ', 'numbered')} />
            <IconButton icon={Table2} label={t('insertTable')} onClick={addHorizontalTable} />
          </div>
        </aside>

        <div className="editor-center">

      <div className="page-strip">
        {pages.map((page, index) => (
          <span
            className={[
              page.id === activePageId ? 'page-tab-group active' : 'page-tab-group',
              page.id === draggedPageId ? 'dragging' : '',
              page.id === dragOverPageId && page.id !== draggedPageId ? 'drag-over' : '',
            ]
              .filter(Boolean)
              .join(' ')}
            draggable
            key={page.id}
            onDragEnd={() => {
              setDraggedPageId(null)
              setDragOverPageId(null)
            }}
            onDragOver={(event) => {
              event.preventDefault()
              if (draggedPageId && draggedPageId !== page.id) setDragOverPageId(page.id)
            }}
            onDragStart={(event) => {
              setDraggedPageId(page.id)
              setActivePageId(page.id)
              event.dataTransfer.effectAllowed = 'move'
              event.dataTransfer.setData('text/plain', page.id)
            }}
            onDrop={(event) => {
              event.preventDefault()
              const sourcePageId = draggedPageId ?? event.dataTransfer.getData('text/plain')
              setDraggedPageId(null)
              setDragOverPageId(null)
              if (!sourcePageId || sourcePageId === page.id) return
              rememberPagesAction()
              reorderPage(sourcePageId, page.id)
            }}
          >
            <button
              className="page-tab"
              onClick={() => {
                activatePage(page.id)
              }}
              type="button"
            >
              <span>{t('page', { number: index + 1 })}</span>
              <span className="page-size-label">{t(pageSizeTranslationKey(page.size))}</span>
            </button>
            <button
              aria-label={t('deletePageNumber', { number: index + 1 })}
              className="delete-page-tab"
              disabled={pages.length <= 1}
              onClick={() => deletePageWithHistory(page.id)}
              title={t('deletePage')}
              type="button"
            >
              <Trash2 size={13} />
            </button>
          </span>
        ))}
        <button className="page-tab add-page-tab" onClick={() => addPageWithHistory()} type="button">
          <Plus size={15} />
          {t('addA4Page')}
        </button>
        <button className="page-tab add-page-tab" onClick={() => addPageWithHistory('a5')} type="button">
          <Plus size={15} />
          {t('addA5Page')}
        </button>
        <button className="page-tab add-page-tab" onClick={() => imagePageInputRef.current?.click()} type="button">
          <ImagePlus size={15} />
          {t('addImagePage')}
        </button>
        <button className="page-tab add-page-tab" onClick={() => pdfPagesInputRef.current?.click()} type="button">
          <FileText size={15} />
          {t('addPdfPages')}
        </button>
        <input
          accept="image/*"
          className="hidden-input"
          onChange={(event) => {
            const file = event.target.files?.[0]
            if (file) {
              void addImagePageWithHistory(file).catch((error) => {
                window.alert(localizedErrorMessage(error, t, 'couldNotAddImagePage'))
              })
            }
            event.currentTarget.value = ''
          }}
          ref={imagePageInputRef}
          type="file"
        />
        <input
          accept=".pdf,application/pdf"
          className="hidden-input"
          onChange={(event) => {
            const file = event.target.files?.[0]
            if (file) {
              void addPdfPages(file).catch((error) => {
                window.alert(localizedErrorMessage(error, t, 'couldNotAddPdfPages'))
              })
            }
            event.currentTarget.value = ''
          }}
          ref={pdfPagesInputRef}
          type="file"
        />
      </div>

      <div className="page-stage" ref={pagesRef}>
        {pages.map((page) => {
          const pageSize = pageSizeFor(page.size)
          return (
          <article
            className={page.id === activePageId ? 'paper active-paper' : 'paper'}
            id={`page-${page.id}`}
            key={page.id}
            style={pageStyle(page)}
            onClick={() => {
              activatePage(page.id)
              setSelectedImageId(null)
              setSelectedTarget(null)
              clearCanvasSelection(page.id)
            }}
          >
            {page.pageImage && <img alt="" className="page-image" draggable={false} src={page.pageImage} />}
            <PageCanvas
              active={page.id === activePageId}
              brushSize={brushSize}
              canSelectPaint={canSelectPaint}
              pageId={page.id}
              recordCanvasHistory={recordCanvasHistory}
              registerCanvas={registerCanvas}
              onSelectDrawing={handleSelectDrawing}
              onOpenDrawingContextMenu={(event, tableCellIndex, pastePoint, hasTarget) =>
                openContextMenu(
                  event,
                  hasTarget
                    ? tableCellIndex == null
                      ? { kind: 'paint', pageId: page.id }
                      : { kind: 'table-cell', pageId: page.id, cellIndex: tableCellIndex }
                    : { kind: 'page', pageId: page.id },
                  pastePoint,
                )
              }
              onPageImageBucketFill={rememberPageImageBucketFill}
              pageHeight={pageSize.height}
              pageWidth={pageSize.width}
              fillPageImageWithBucket={fillPageImageWithBucket}
              textColor={textColor}
              textBold={textBold}
              textFontFamily={textFontFamily}
              textFontSize={Number.parseFloat(textFontSize)}
              textItalic={textItalic}
              textLineGap={textLineGap}
              textUnderline={textUnderline}
              fill={fill}
              stroke={stroke}
              zIndex={relativeLayerZ.paint}
              tool={tool}
            />
            {page.id === activePageId ? (
              <div
                className={textLayerClass(tool, canSelectText)}
                onContextMenu={(event) => {
                  if (editor?.state.selection.empty) return
                  setSelectedImageId(null)
                  clearCanvasSelection(page.id)
                  openContextMenu(event, { kind: 'text', pageId: page.id })
                }}
                onInput={rememberTextAction}
                onClick={(event) => event.stopPropagation()}
                onMouseUp={() => {
                  if (editor && !editor.state.selection.empty) {
                    setSelectedImageId(null)
                    clearCanvasSelection(page.id)
                    setSelectedTarget({ kind: 'text', pageId: page.id })
                  }
                }}
                onPointerDown={(event) => {
                  event.stopPropagation()
                  if (tool !== 'select-text') setTool('write')
                  setSelectedImageId(null)
                  setSelectedTarget({ kind: 'text', pageId: page.id })
                  clearCanvasSelection(page.id)
                }}
              >
                <EditorContent editor={editor} />
              </div>
            ) : (
              <div className="static-page-content" dangerouslySetInnerHTML={{ __html: page.html }} />
            )}
            <div className={tool === 'select-paint' ? 'floating-layer interactive' : 'floating-layer'} style={{ zIndex: relativeLayerZ.images }}>
              {page.images.map((image) => (
                <FloatingImageObject
                  active={page.id === activePageId && image.id === selectedImageId}
                  image={image}
                  key={image.id}
                  pageHeight={pageSize.height}
                  pageWidth={pageSize.width}
                  onSelect={() => {
                    setActivePageId(page.id)
                    setSelectedImageId(image.id)
                    setSelectedTarget({ kind: 'image', pageId: page.id, imageId: image.id })
                    clearCanvasSelection(page.id)
                    setTool('select-paint')
                  }}
                  onOpenContextMenu={(event) => openContextMenu(event, { kind: 'image', pageId: page.id, imageId: image.id })}
                  onChangeStart={() => rememberPagesAction()}
                  t={t}
                  updateImage={(patch) => updateImageWithHistory(page.id, image.id, patch)}
                />
              ))}
            </div>
          </article>
          )
        })}
      </div>
        </div>

        <aside className="editor-rail editor-rail-right" aria-label={t('drawing')}>
          <div className="rail-section">
            <ToolButton active={tool === 'select-paint'} icon={MousePointer2} label={t('selectPaint')} onClick={() => activateTool('select-paint')} />
            <IconButton active={tool === 'brush'} icon={Brush} label={t('brush')} onClick={() => activateTool('brush')} />
            <IconButton active={tool === 'eraser'} icon={Eraser} label={t('eraser')} onClick={() => activateTool('eraser')} />
            <IconButton active={tool === 'bucket'} icon={PaintBucket} label={t('bucket')} onClick={() => activateTool('bucket')} />
            <label className="range-control">
              <PenLine size={16} />
              <input min="1" max="42" onChange={(event) => setBrushSize(Number(event.target.value))} type="range" value={brushSize} />
              <span>{brushSize}</span>
            </label>
            <label className="color-chip" title={t('strokeColor')}>
              <Highlighter size={16} />
              <input onChange={(event) => setStroke(event.target.value)} type="color" value={stroke} />
            </label>
            <label className="color-chip" title={t('fillColor')}>
              <Shapes size={16} />
              <input onChange={(event) => setFill(event.target.value)} type="color" value={fill} />
            </label>
          </div>
          <div className="rail-section compact-icons">
            <IconButton icon={RectangleHorizontal} label={t('rectangle')} onClick={() => addFabricObject(new Rect({ left: 160, top: 160, width: 190, height: 110, fill, stroke, strokeWidth: 3, rx: 8, ry: 8 }))} />
            <IconButton icon={CircleIcon} label={t('circle')} onClick={() => addFabricObject(new Circle({ left: 220, top: 150, radius: 64, fill, stroke, strokeWidth: 3 }))} />
            <IconButton
              icon={X}
              label={t('xShape')}
              onClick={() => {
                const size = 140
                const width = Math.max(3, brushSize)
                const firstStroke = new Line([0, 0, size, size], { stroke, strokeWidth: width, strokeLineCap: 'round' })
                const secondStroke = new Line([size, 0, 0, size], { stroke, strokeWidth: width, strokeLineCap: 'round' })
                addFabricObject(new Group([firstStroke, secondStroke], { left: 220, top: 160 }))
              }}
            />
            <IconButton icon={Rows3} label={t('line')} onClick={() => addFabricObject(new Line([150, 230, 470, 300], { stroke, strokeWidth: brushSize, strokeLineCap: 'round' }))} />
            <IconButton
              icon={ArrowRight}
              label={t('arrow')}
              onClick={() => {
                const line = new Line([0, 0, 240, 80], { stroke, strokeWidth: brushSize, strokeLineCap: 'round' })
                const head = new Triangle({ left: 240, top: 80, width: 24, height: 28, fill: stroke, angle: 110, originX: 'center', originY: 'center' })
                addFabricObject(new Group([line, head], { left: 170, top: 220 }))
              }}
            />
          </div>
          <div className="rail-section">
            <IconButton active={layerZ.text > layerZ.paint} icon={Type} label={t('textFront')} onClick={() => moveLayerToFront('text')} />
            <IconButton active={layerZ.paint > layerZ.text} icon={Paintbrush} label={t('paintFront')} onClick={() => moveLayerToFront('paint')} />
            <IconButton icon={ArrowRight} label={t('selectedToFront')} onClick={() => moveSelectedObjectLayer('front')} />
            <IconButton icon={Rows3} label={t('selectedToBack')} onClick={() => moveSelectedObjectLayer('back')} />
          </div>
          <span className="rail-spacer" />
          <div className="rail-section export-rail-section">
            <ExportMenu fileName={exportFileName} formats={exportFormats} onExport={exportDocument} onFileNameChange={setExportFileName} t={t} />
          </div>
        </aside>
      </div>
      {pendingImageExport && (
        <ImageExportConfirmation
          format={pendingImageExport.format}
          pageCount={pendingImageExport.pageCount}
          onCancel={() => setPendingImageExport(null)}
          onContinue={continueImageExport}
          t={t}
        />
      )}
      {contextMenu && (
        <SelectionContextMenu
          canDeleteCell={contextMenu.target.kind === 'table-cell'}
          canPaste={hasFabricClipboard}
          isPageMenu={contextMenu.target.kind === 'page'}
          x={contextMenu.x}
          y={contextMenu.y}
          onBack={() => moveSelectedObjectLayer('back')}
          onCopy={copyFromContextMenu}
          onDelete={deleteSelectedObject}
          onDeleteCell={deleteSelectedTableCell}
          onFront={() => moveSelectedObjectLayer('front')}
          onPaste={pasteFromContextMenu}
          t={t}
        />
      )}
    </section>
  )
}

function PageCanvas({
  active,
  brushSize,
  canSelectPaint,
  pageHeight,
  pageId,
  pageWidth,
  fillPageImageWithBucket,
  onPageImageBucketFill,
  recordCanvasHistory,
  registerCanvas,
  onSelectDrawing,
  onOpenDrawingContextMenu,
  textColor,
  textBold,
  textFontFamily,
  textFontSize,
  textItalic,
  textLineGap,
  textUnderline,
  fill,
  stroke,
  zIndex,
  tool,
}: {
  active: boolean
  brushSize: number
  canSelectPaint: boolean
  pageHeight: number
  pageId: string
  pageWidth: number
  fillPageImageWithBucket: (pageId: string, x: number, y: number, fillColor: string) => Promise<boolean>
  onPageImageBucketFill: () => void
  recordCanvasHistory: (pageId: string, canvas: FabricCanvas) => void
  registerCanvas: (pageId: string, canvas: FabricCanvas | null) => void
  onSelectDrawing: (pageId: string, tableCellIndex?: number) => void
  onOpenDrawingContextMenu: (event: MouseEvent, tableCellIndex: number | undefined, pastePoint: CanvasPoint, hasTarget: boolean) => void
  textColor: string
  textBold: boolean
  textFontFamily: string
  textFontSize: number
  textItalic: boolean
  textLineGap: number
  textUnderline: boolean
  fill: string
  stroke: string
  zIndex: number
  tool: EditorTool
}) {
  const canvasEl = useRef<HTMLCanvasElement>(null)
  const fabricRef = useRef<FabricCanvas | null>(null)
  const activeRef = useRef(active)
  const toolRef = useRef(tool)
  const textColorRef = useRef(textColor)
  const textBoldRef = useRef(textBold)
  const textFontFamilyRef = useRef(textFontFamily)
  const textFontSizeRef = useRef(textFontSize)
  const textItalicRef = useRef(textItalic)
  const textLineGapRef = useRef(textLineGap)
  const textUnderlineRef = useRef(textUnderline)
  const fillRef = useRef(fill)
  const strokeRef = useRef(stroke)
  const canSelectPaintRef = useRef(canSelectPaint)
  const callbacksRef = useRef({ fillPageImageWithBucket, onOpenDrawingContextMenu, onPageImageBucketFill, onSelectDrawing, recordCanvasHistory })

  useEffect(() => {
    activeRef.current = active
    toolRef.current = tool
    textColorRef.current = textColor
    textBoldRef.current = textBold
    textFontFamilyRef.current = textFontFamily
    textFontSizeRef.current = textFontSize
    textItalicRef.current = textItalic
    textLineGapRef.current = textLineGap
    textUnderlineRef.current = textUnderline
    fillRef.current = fill
    strokeRef.current = stroke
    canSelectPaintRef.current = canSelectPaint
    callbacksRef.current = { fillPageImageWithBucket, onOpenDrawingContextMenu, onPageImageBucketFill, onSelectDrawing, recordCanvasHistory }
  }, [
    active,
    canSelectPaint,
    fill,
    fillPageImageWithBucket,
    onOpenDrawingContextMenu,
    onPageImageBucketFill,
    onSelectDrawing,
    recordCanvasHistory,
    stroke,
    textBold,
    textColor,
    textFontFamily,
    textFontSize,
    textItalic,
    textLineGap,
    textUnderline,
    tool,
  ])

  useEffect(() => {
    if (!canvasEl.current) return
    const canvas = new FabricCanvas(canvasEl.current, {
      backgroundColor: 'rgba(255,255,255,0)',
      enableRetinaScaling: true,
      height: pageHeight,
      fireRightClick: true,
      preserveObjectStacking: true,
      selection: true,
      stopContextMenu: true,
      width: pageWidth,
    })
    fabricRef.current = canvas
    const record = () => callbacksRef.current.recordCanvasHistory(pageId, canvas)
    const constrain = (event: { target?: FabricObject }) => {
      if (event.target) keepObjectInsideCanvas(event.target, canvas)
    }
    canvas.on('object:added', record)
    canvas.on('object:modified', record)
    canvas.on('object:removed', record)
    canvas.on('object:moving', constrain)
    canvas.on('object:scaling', constrain)
    canvas.on('object:rotating', constrain)
    const tableCellIndexFromEvent = (event: { e?: Event; target?: FabricObject; subTargets?: FabricObject[] }) => {
      if (!(event.e instanceof MouseEvent)) return null
      const target = event.target ?? canvas.getActiveObject()
      const table = horizontalTableFromTarget(target) ?? horizontalTableFromTarget(canvas.getActiveObject())
      return table ? horizontalTableCellIndexFromPointer(table, canvas.getScenePoint(event.e), event.subTargets ?? (target ? [target] : [])) : null
    }
    const selectDrawing = (event: { e?: Event; target?: FabricObject; subTargets?: FabricObject[] }) => {
      callbacksRef.current.onSelectDrawing(pageId, tableCellIndexFromEvent(event) ?? undefined)
    }
    const openDrawingMenu = (event: { e: Event; target?: FabricObject; subTargets?: FabricObject[] }) => {
      const canOpenPaintMenu = canSelectPaintRef.current || toolRef.current === 'place-text'
      if (!activeRef.current || !canOpenPaintMenu || !(event.e instanceof MouseEvent)) return
      const target = event.target ?? canvas.getActiveObject()
      const pointer = canvas.getScenePoint(event.e)
      if (!target) {
        callbacksRef.current.onOpenDrawingContextMenu(event.e, undefined, pointer, false)
        return
      }
      const activeObject = canvas.getActiveObject()
      const table = horizontalTableFromTarget(target) ?? horizontalTableFromTarget(activeObject)
      const tableCellIndex = table ? horizontalTableCellIndexFromPointer(table, pointer, event.subTargets ?? [target]) : null
      canvas.setActiveObject(table ?? target)
      canvas.requestRenderAll()
      callbacksRef.current.onSelectDrawing(pageId, tableCellIndex ?? undefined)
      callbacksRef.current.onOpenDrawingContextMenu(event.e, tableCellIndex ?? undefined, pointer, true)
    }
    const handleMouseDown = (event: { e: Event; target?: FabricObject; subTargets?: FabricObject[] }) => {
      if (!activeRef.current || !(event.e instanceof MouseEvent)) return
      if (toolRef.current === 'bucket') {
        const table = horizontalTableFromTarget(event.target)
        const tableCellIndex = table ? horizontalTableCellIndexFromPointer(table, canvas.getScenePoint(event.e), event.subTargets ?? []) : null
        if (table && tableCellIndex != null) {
          fillHorizontalTableCell(table, tableCellIndex, fillRef.current)
          canvas.setActiveObject(table)
          callbacksRef.current.onSelectDrawing(pageId, tableCellIndex)
        } else if (event.target) {
          applyPaintBucketFill(event.target, fillRef.current)
          canvas.setActiveObject(horizontalTableFromTarget(event.target) ?? event.target)
          callbacksRef.current.onSelectDrawing(pageId)
        } else {
          const pointer = canvas.getScenePoint(event.e)
          void callbacksRef.current
            .fillPageImageWithBucket(pageId, pointer.x, pointer.y, fillRef.current)
            .catch(() => true)
            .then((filledPageImage) => {
              if (filledPageImage) {
                callbacksRef.current.onPageImageBucketFill()
                callbacksRef.current.onSelectDrawing(pageId)
                return
              }
              canvas.backgroundColor = fillRef.current
              callbacksRef.current.onSelectDrawing(pageId)
              canvas.requestRenderAll()
              callbacksRef.current.recordCanvasHistory(pageId, canvas)
            })
          return
        }
        canvas.requestRenderAll()
        callbacksRef.current.recordCanvasHistory(pageId, canvas)
        return
      }
      if (toolRef.current !== 'place-text') {
        const canSelectCell = canSelectPaintRef.current || toolRef.current === 'select-paint'
        if (canSelectCell && event.target) callbacksRef.current.onSelectDrawing(pageId, tableCellIndexFromEvent(event) ?? undefined)
        return
      }
      if (event.target) {
        callbacksRef.current.onSelectDrawing(pageId, tableCellIndexFromEvent(event) ?? undefined)
        return
      }
      const pointer = canvas.getScenePoint(event.e)
      const text = new IText('Type here', {
        fill: textColorRef.current,
        fontFamily: textFontFamilyRef.current,
        fontSize: textFontSizeRef.current,
        fontStyle: textItalicRef.current ? 'italic' : 'normal',
        fontWeight: textBoldRef.current ? 'bold' : 'normal',
        left: Math.round(clamp(pointer.x, 0, pageWidth - 120)),
        lineHeight: textLineGapRef.current,
        noScaleCache: true,
        objectCaching: false,
        top: Math.round(clamp(pointer.y, 0, pageHeight - 42)),
        underline: textUnderlineRef.current,
      })
      configureManagedText(text, 'plain', canvas)
      canvas.add(text)
      canvas.setActiveObject(text)
      text.enterEditing()
      text.selectAll()
      canvas.requestRenderAll()
      callbacksRef.current.onSelectDrawing(pageId)
    }
    canvas.on('selection:created', selectDrawing)
    canvas.on('selection:updated', selectDrawing)
    canvas.on('contextmenu', openDrawingMenu)
    canvas.on('mouse:down', handleMouseDown)
    registerCanvas(pageId, canvas)
    return () => {
      registerCanvas(pageId, null)
      canvas.dispose()
      fabricRef.current = null
    }
  }, [pageHeight, pageId, pageWidth, registerCanvas])

  useEffect(() => {
    const canvas = fabricRef.current
    if (!canvas) return
    const paintSelectionActive = active && (canSelectPaint || tool === 'bucket' || tool === 'place-text')
    canvas.isDrawingMode = active && (tool === 'brush' || tool === 'eraser')
    canvas.selection = paintSelectionActive
    canvas.skipTargetFind = !paintSelectionActive
    canvas.getObjects().forEach((object) => {
      object.selectable = paintSelectionActive
      object.evented = paintSelectionActive
    })
    if (canvas.isDrawingMode) {
      const brush = new PencilBrush(canvas)
      brush.width = brushSize
      brush.color = tool === 'eraser' ? '#ffffff' : stroke
      canvas.freeDrawingBrush = brush
    }
    canvas.requestRenderAll()
  }, [active, brushSize, canSelectPaint, stroke, tool])

  const interactive = active && (tool === 'brush' || tool === 'eraser' || tool === 'bucket' || canSelectPaint || tool === 'place-text')

  return (
    <div
      className={interactive ? 'drawing-layer interactive' : 'drawing-layer'}
      onClick={(event) => event.stopPropagation()}
      onPointerDown={(event) => event.stopPropagation()}
      style={{ zIndex }}
    >
      <canvas height={pageHeight} ref={canvasEl} width={pageWidth} />
    </div>
  )
}

function FloatingImageObject({
  active,
  image,
  onChangeStart,
  onOpenContextMenu,
  onSelect,
  pageHeight,
  pageWidth,
  t,
  updateImage,
}: {
  active: boolean
  image: FloatingImage
  onChangeStart: () => void
  onOpenContextMenu: (event: React.MouseEvent<HTMLDivElement>) => void
  onSelect: () => void
  pageHeight: number
  pageWidth: number
  t: Translate
  updateImage: (patch: Partial<FloatingImage>) => void
}) {
  const startDrag = (event: ReactPointerEvent<HTMLDivElement>) => {
    if (event.button !== 0) {
      event.stopPropagation()
      onSelect()
      return
    }
    event.preventDefault()
    event.stopPropagation()
    onSelect()
    onChangeStart()
    const startX = event.clientX
    const startY = event.clientY
    const origin = { x: image.x, y: image.y }
    const move = (moveEvent: PointerEvent) => {
      updateImage({
        x: clamp(origin.x + moveEvent.clientX - startX, 0, pageWidth - image.width),
        y: clamp(origin.y + moveEvent.clientY - startY, 0, pageHeight - image.height),
      })
    }
    bindPointerDrag(move)
  }

  const startCornerResize = (event: ReactPointerEvent<HTMLButtonElement>) => {
    event.preventDefault()
    event.stopPropagation()
    onSelect()
    onChangeStart()
    const startX = event.clientX
    const ratio = image.height / image.width
    const move = (moveEvent: PointerEvent) => {
      const nextWidth = clamp(image.width + moveEvent.clientX - startX, 64, pageWidth - image.x)
      updateImage({
        width: nextWidth,
        height: clamp(nextWidth * ratio, 48, pageHeight - image.y),
      })
    }
    bindPointerDrag(move)
  }

  const startHorizontalResize = (event: ReactPointerEvent<HTMLButtonElement>) => {
    event.preventDefault()
    event.stopPropagation()
    onSelect()
    onChangeStart()
    const startX = event.clientX
    const move = (moveEvent: PointerEvent) => {
      updateImage({
        width: clamp(image.width + moveEvent.clientX - startX, 64, pageWidth - image.x),
      })
    }
    bindPointerDrag(move)
  }

  const startVerticalResize = (event: ReactPointerEvent<HTMLButtonElement>) => {
    event.preventDefault()
    event.stopPropagation()
    onSelect()
    onChangeStart()
    const startY = event.clientY
    const move = (moveEvent: PointerEvent) => {
      updateImage({
        height: clamp(image.height + moveEvent.clientY - startY, 48, pageHeight - image.y),
      })
    }
    bindPointerDrag(move)
  }

  const startRotate = (event: ReactPointerEvent<HTMLButtonElement>) => {
    event.preventDefault()
    event.stopPropagation()
    onSelect()
    onChangeStart()
    const paperBounds = event.currentTarget.closest('.paper')?.getBoundingClientRect()
    if (!paperBounds) return
    const centerX = paperBounds.left + image.x + image.width / 2
    const centerY = paperBounds.top + image.y + image.height / 2
    const move = (moveEvent: PointerEvent) => {
      const radians = Math.atan2(moveEvent.clientY - centerY, moveEvent.clientX - centerX)
      updateImage({ rotation: Math.round((radians * 180) / Math.PI) })
    }
    bindPointerDrag(move)
  }

  return (
    <div
      className={active ? 'floating-image active' : 'floating-image'}
      onClick={(event) => {
        event.stopPropagation()
        onSelect()
      }}
      onContextMenu={(event) => {
        onSelect()
        onOpenContextMenu(event)
      }}
      onPointerDown={startDrag}
      style={{
        height: image.height,
        left: image.x,
        top: image.y,
        transform: `rotate(${image.rotation}deg)`,
        width: image.width,
      }}
    >
      <img alt="" draggable={false} src={image.src} />
      {active && (
        <>
          <button className="image-handle move-handle" title={t('move')} type="button">
            <Move size={12} />
          </button>
          <button className="image-handle resize-horizontal-handle" onPointerDown={startHorizontalResize} title={t('resizeWidth')} type="button" />
          <button className="image-handle resize-vertical-handle" onPointerDown={startVerticalResize} title={t('resizeHeight')} type="button" />
          <button className="image-handle resize-corner-handle" onPointerDown={startCornerResize} title={t('resizeProportionally')} type="button" />
          <button className="image-handle rotate-handle" onPointerDown={startRotate} title={t('rotate')} type="button">
            <RotateCw size={12} />
          </button>
        </>
      )}
    </div>
  )
}

function SelectionContextMenu({
  canDeleteCell,
  canPaste,
  isPageMenu,
  x,
  y,
  onBack,
  onCopy,
  onDelete,
  onDeleteCell,
  onFront,
  onPaste,
  t,
}: {
  canDeleteCell: boolean
  canPaste: boolean
  isPageMenu: boolean
  x: number
  y: number
  onBack: () => void
  onCopy: () => void
  onDelete: () => void
  onDeleteCell: () => void
  onFront: () => void
  onPaste: () => void
  t: Translate
}) {
  return (
    <div
      className="selection-context-menu"
      onClick={(event) => event.stopPropagation()}
      style={{
        left: Math.min(x, window.innerWidth - 188),
        top: Math.min(y, window.innerHeight - 196),
      }}
    >
      {!isPageMenu && (
        <>
          <button onClick={onCopy} type="button">
            {t('copy')}
          </button>
          {canPaste && (
            <button onClick={onPaste} type="button">
              {t('paste')}
            </button>
          )}
          <button onClick={onFront} type="button">
            {t('selectedToFront')}
          </button>
          <button onClick={onBack} type="button">
            {t('selectedToBack')}
          </button>
          {canDeleteCell && (
            <button onClick={onDeleteCell} type="button">
              {t('deleteCell')}
            </button>
          )}
          <button onClick={onDelete} type="button">
            {canDeleteCell ? t('deleteTable') : t('delete')}
          </button>
        </>
      )}
      {isPageMenu && canPaste && (
        <button onClick={onPaste} type="button">
          {t('paste')}
        </button>
      )}
    </div>
  )
}

type ConvertMessage =
  | { kind: 'prompt' }
  | { kind: 'converting' }
  | { fileName: string; kind: 'converted'; limitation: boolean; target: ExportFormat }
  | { kind: 'error'; message?: string }

function ConvertMode({ t }: { t: Translate }) {
  const [file, setFile] = useState<File | null>(null)
  const [target, setTarget] = useState<ExportFormat>('pdf')
  const [message, setMessage] = useState<ConvertMessage>({ kind: 'prompt' })
  const fileType = useMemo(() => (file ? extensionOf(file.name) : ''), [file])
  const displayMessage = useMemo(() => {
    if (message.kind === 'converting') return t('converting')
    if (message.kind === 'converted') {
      const limitation = message.limitation ? ` ${t('pdfDocxExtractedOnly')}` : ''
      return `${t('converted', { file: message.fileName, format: message.target.toUpperCase() })}${limitation}`
    }
    if (message.kind === 'error') return localizedErrorMessage(message.message, t, 'unavailableConversion')
    return t('convertPrompt')
  }, [message, t])

  const convert = async () => {
    if (!file) return
    setMessage({ kind: 'converting' })
    try {
      await convertFile(file, target)
      setMessage({ fileName: file.name, kind: 'converted', limitation: fileType === 'pdf' && target === 'docx', target })
    } catch (error) {
      setMessage({ kind: 'error', message: error instanceof Error ? error.message : undefined })
    }
  }

  return (
    <section className="mode-panel">
      <Ribbon title={t('convertAFIle')}>
        <label className="upload-button">
          <Upload size={17} />
          <span>{file ? file.name : t('chooseFile')}</span>
          <input
            accept=".docx,.pdf,.txt,.md,.markdown,.png,.jpg,.jpeg"
            className="hidden-input"
            onChange={(event) => {
              setFile(event.target.files?.[0] ?? null)
              setMessage({ kind: 'prompt' })
            }}
            type="file"
          />
        </label>
        <select aria-label={t('targetFormat')} onChange={(event) => setTarget(event.target.value as ExportFormat)} value={target}>
          {exportFormats.map((format) => (
            <option key={format} value={format}>
              {format.toUpperCase()}
            </option>
          ))}
        </select>
        <button className="primary-action" disabled={!file} onClick={() => void convert()} type="button">
          <ScissorsLineDashed size={17} />
          {t('convert')}
        </button>
      </Ribbon>
      <div className="converter-stage">
        <div className="drop-zone">
          <Upload size={42} />
          <h2>{file ? file.name : t('convertDropZone')}</h2>
          <p>{displayMessage}</p>
          <p className="fine-print">{t('convertFinePrint')}</p>
        </div>
      </div>
    </section>
  )
}

function Ribbon({ title, children }: { title: string; children: React.ReactNode }) {
  return (
    <div className="ribbon">
      <div className="ribbon-title">{title}</div>
      <div className="ribbon-controls">{children}</div>
    </div>
  )
}

function textLayerClass(tool: EditorTool, canSelectText: boolean) {
  const classes = ['text-editor-layer']
  if (tool === 'write' || canSelectText) classes.push('active')
  if (tool === 'write') classes.push('write-mode')
  if (canSelectText) classes.push('select-text-mode')
  return classes.join(' ')
}

function IconButton({
  active,
  icon: Icon,
  label,
  onClick,
}: {
  active?: boolean
  icon: React.ComponentType<{ size?: number }>
  label: string
  onClick: () => void
}) {
  return (
    <button aria-label={label} className={active ? 'icon-button active' : 'icon-button'} onClick={onClick} title={label} type="button">
      <Icon size={17} />
    </button>
  )
}

function ToolButton({
  active,
  icon: Icon,
  label,
  onClick,
}: {
  active?: boolean
  icon: React.ComponentType<{ size?: number }>
  label: string
  onClick: () => void
}) {
  return (
    <button aria-label={label} className={active ? 'tool-button active' : 'tool-button'} onClick={onClick} title={label} type="button">
      <Icon size={16} />
      <span>{label}</span>
    </button>
  )
}

function ExportMenu({
  fileName,
  formats,
  onExport,
  onFileNameChange,
  t,
}: {
  fileName: string
  formats: ExportFormat[]
  onExport: (format: ExportFormat, fileName: string) => void | Promise<void>
  onFileNameChange: (fileName: string) => void
  t: Translate
}) {
  return (
    <label className="export-menu export-menu-with-name">
      <Download size={17} />
      <input
        aria-label={t('exportFileName')}
        className="export-file-name"
        onChange={(event) => onFileNameChange(event.target.value)}
        placeholder={t('fileName')}
        type="text"
        value={fileName}
      />
      <select
        aria-label={t('exportFormat')}
        defaultValue=""
        onChange={(event) => {
          const value = event.target.value as ExportFormat
          if (value) {
            void Promise.resolve(onExport(value, fileName)).catch((error) => {
              window.alert(localizedErrorMessage(error, t, 'couldNotExport'))
            })
          }
          event.currentTarget.value = ''
        }}
      >
        <option value="" disabled>
          {t('export')}
        </option>
        {formats.map((format) => (
          <option key={format} value={format}>
            {format.toUpperCase()}
          </option>
        ))}
      </select>
    </label>
  )
}

function ImageExportConfirmation({
  format,
  pageCount,
  onCancel,
  onContinue,
  t,
}: {
  format: 'png' | 'jpg'
  pageCount: number
  onCancel: () => void
  onContinue: () => void
  t: Translate
}) {
  return (
    <div className="dialog-backdrop">
      <div aria-labelledby="image-export-confirmation-title" aria-modal="true" className="confirmation-dialog" role="dialog">
        <div className="dialog-icon">
          <Download size={22} />
        </div>
        <h2 id="image-export-confirmation-title">{t('imageExportMultiplePagesTitle')}</h2>
        <p>{t('imageExportMultiplePagesMessage', { count: pageCount, format: format.toUpperCase() })}</p>
        <div className="dialog-actions">
          <button className="dialog-secondary" onClick={onCancel} type="button">
            {t('cancel')}
          </button>
          <button className="primary-action" onClick={onContinue} type="button">
            {t('continue')}
          </button>
        </div>
      </div>
    </div>
  )
}

async function convertFile(file: File, target: ExportFormat) {
  const source = extensionOf(file.name)
  const baseName = file.name.replace(/\.[^.]+$/, '') || 'akxreditor-conversion'

  if (source === 'png' || source === 'jpg' || source === 'jpeg') {
    const dataUrl = await fileToDataUrl(file)
    if (target === 'png' || target === 'jpg') {
      saveBlob(await convertImageBlob(dataUrl, target), `${baseName}.${target}`)
      return
    }
    if (target === 'pdf') {
      saveBlob(
        await exportPagesPdfBlob([
          {
            id: createId(),
            html: starterContent,
            images: [],
            pageImage: dataUrl,
            size: defaultPageSizeKey,
          },
        ]),
        `${baseName}.pdf`,
      )
      return
    }
    if (target === 'html') {
      saveBlob(new Blob([wrapHtmlDocument(`<img src="${dataUrl}" alt="${file.name}" />`)], { type: 'text/html;charset=utf-8' }), `${baseName}.html`)
      return
    }
    throw new Error('Image to text, Markdown, or DOCX needs OCR, which is not included in this client-side build.')
  }

  const { text, html } = await readTextualFile(file, source)
  if (target === 'txt') {
    saveBlob(new Blob([text], { type: 'text/plain;charset=utf-8' }), `${baseName}.txt`)
    return
  }
  if (target === 'md') {
    saveBlob(new Blob([html ? markdownFromHtml(html) : text], { type: 'text/markdown;charset=utf-8' }), `${baseName}.md`)
    return
  }
  if (target === 'html') {
    saveBlob(new Blob([wrapHtmlDocument(html || `<pre>${escapeHtml(text)}</pre>`)], { type: 'text/html;charset=utf-8' }), `${baseName}.html`)
    return
  }
  if (target === 'docx') {
    saveBlob(await createDocxBlob(text), `${baseName}.docx`)
    return
  }
  if (target === 'pdf') {
    saveBlob(await exportPagesPdfBlob(pagesFromText(text)), `${baseName}.pdf`)
    return
  }
  saveBlob(await textToImageBlob(text, target), `${baseName}.${target}`)
}

async function pagesFromFile(file: File) {
  const source = extensionOf(file.name)
  if (source === 'pdf') return pagesFromPdf(file)
  if (source === 'png' || source === 'jpg' || source === 'jpeg') {
    const src = await fileToDataUrl(file)
    const image = await loadImage(src)
    const pageSize = defaultPageSize
    const scale = Math.min((pageSize.width - 240) / image.naturalWidth, (pageSize.height - 260) / image.naturalHeight, 1)
    return [
      {
        id: createId(),
        html: starterContent,
        size: defaultPageSizeKey,
        images: [
          {
            id: createId(),
            src,
            x: 96,
            y: 96,
            width: image.naturalWidth * scale,
            height: image.naturalHeight * scale,
            rotation: 0,
          },
        ],
      },
    ]
  }

  const { text, html, layout } = await readTextualFile(file, source)
  if (source === 'txt') return pagesFromText(text)
  return pagesFromHtml(html || textToEditableHtml(text), layout)
}

async function readDocxLayout(arrayBuffer: ArrayBuffer): Promise<Partial<Pick<PageData, 'fontFamily' | 'margins' | 'size'>>> {
  const zip = await JSZip.loadAsync(arrayBuffer)
  const documentXml = await zip.file('word/document.xml')?.async('text')
  const stylesXml = await zip.file('word/styles.xml')?.async('text')
  if (!documentXml) return { margins: defaultPageMargins, size: defaultPageSizeKey, fontFamily: 'Calibri' }

  const parser = new DOMParser()
  const documentDom = parser.parseFromString(documentXml, 'application/xml')
  const stylesDom = stylesXml ? parser.parseFromString(stylesXml, 'application/xml') : null
  const section = xmlFirst(documentDom, 'sectPr')
  const pageSize = section ? pageSizeFromSectPr(section) : defaultPageSizeKey
  const margins = section ? marginsFromSectPr(section) : defaultPageMargins
  const fontFamily = fontFromDocx(stylesDom) ?? fontFromDocx(documentDom) ?? 'Calibri'
  return { fontFamily, margins, size: pageSize }
}

function pageSizeFromSectPr(section: Element): PageSizeKey {
  const pageSize = xmlFirst(section, 'pgSz')
  const widthTwips = Number(xmlAttr(pageSize, 'w'))
  const heightTwips = Number(xmlAttr(pageSize, 'h'))
  if (!Number.isFinite(widthTwips) || !Number.isFinite(heightTwips)) return defaultPageSizeKey
  const isLandscape = xmlAttr(pageSize, 'orient') === 'landscape' || widthTwips > heightTwips
  const portraitWidth = Math.min(widthTwips, heightTwips)
  const portraitHeight = Math.max(widthTwips, heightTwips)
  const a5Distance = Math.abs(portraitWidth - pageSizes.a5.docxWidth) + Math.abs(portraitHeight - pageSizes.a5.docxHeight)
  const a4Distance = Math.abs(portraitWidth - pageSizes.a4.docxWidth) + Math.abs(portraitHeight - pageSizes.a4.docxHeight)
  if (a5Distance < a4Distance) return isLandscape ? 'a5Landscape' : 'a5'
  return isLandscape ? 'a4Landscape' : 'a4'
}

function marginsFromSectPr(section: Element): PageMargins {
  const pageMargins = xmlFirst(section, 'pgMar')
  return exactPageMargins({
    top: twipsToPx(Number(xmlAttr(pageMargins, 'top')) || defaultPageMargins.top * 15),
    right: twipsToPx(Number(xmlAttr(pageMargins, 'right')) || defaultPageMargins.right * 15),
    bottom: twipsToPx(Number(xmlAttr(pageMargins, 'bottom')) || defaultPageMargins.bottom * 15),
    left: twipsToPx(Number(xmlAttr(pageMargins, 'left')) || defaultPageMargins.left * 15),
  })
}

function fontFromDocx(document: Document | null) {
  const fonts = Array.from(document?.getElementsByTagName('*') ?? []).find((element) => element.localName === 'rFonts')
  return normalizeWordFont(xmlAttr(fonts, 'ascii') || xmlAttr(fonts, 'hAnsi') || xmlAttr(fonts, 'asciiTheme') || xmlAttr(fonts, 'hAnsiTheme'))
}

function xmlFirst(root: ParentNode, localName: string) {
  return Array.from(root.querySelectorAll('*')).find((element) => element.localName === localName)
}

function xmlAttr(element: Element | undefined, localName: string) {
  if (!element) return ''
  return (
    element.getAttribute(`w:${localName}`) ??
    Array.from(element.attributes).find((attribute) => attribute.localName === localName)?.value ??
    ''
  )
}

function twipsToPx(twips: number) {
  return Math.round(twips / 15)
}

function normalizeWordFont(font: string | undefined) {
  if (!font) return ''
  const normalizedFont = font.toLowerCase()
  if (normalizedFont.includes('major')) return 'Cambria'
  if (normalizedFont.includes('minor')) return 'Calibri'
  if (normalizedFont.includes('cambria')) return 'Cambria'
  if (normalizedFont.includes('calibri')) return 'Calibri'
  return font
}

async function pagesFromPdf(file: File) {
  const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise
  const pages: PageData[] = []
  const pageSize = defaultPageSize
  const pixelRatio = Math.min(Math.max(window.devicePixelRatio || 1, 2), 3)

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber)
    const baseViewport = page.getViewport({ scale: 1 })
    const outputWidth = Math.round(pageSize.width * pixelRatio)
    const outputHeight = Math.round(pageSize.height * pixelRatio)
    const renderScale = Math.min(outputWidth / baseViewport.width, outputHeight / baseViewport.height)
    const viewport = page.getViewport({ scale: renderScale })
    const renderCanvas = document.createElement('canvas')
    const context = renderCanvas.getContext('2d')
    if (!context) throw new Error('Canvas rendering is not available in this browser.')
    renderCanvas.width = Math.ceil(viewport.width)
    renderCanvas.height = Math.ceil(viewport.height)
    await page.render({ canvas: renderCanvas, canvasContext: context, viewport }).promise

    const outputCanvas = document.createElement('canvas')
    const outputContext = outputCanvas.getContext('2d')
    if (!outputContext) throw new Error('Canvas rendering is not available in this browser.')
    outputCanvas.width = outputWidth
    outputCanvas.height = outputHeight
    outputContext.fillStyle = '#ffffff'
    outputContext.fillRect(0, 0, outputCanvas.width, outputCanvas.height)
    outputContext.imageSmoothingEnabled = true
    outputContext.imageSmoothingQuality = 'high'
    outputContext.drawImage(renderCanvas, (outputWidth - renderCanvas.width) / 2, (outputHeight - renderCanvas.height) / 2)

    pages.push({
      id: createId(),
      html: starterContent,
      size: defaultPageSizeKey,
      images: [],
      pageImage: outputCanvas.toDataURL('image/png'),
    })
  }

  return pages.length ? pages : [{ id: createId(), html: starterContent, images: [], size: defaultPageSizeKey }]
}

async function readTextualFile(file: File, source: string) {
  if (source === 'docx') {
    const arrayBuffer = await file.arrayBuffer()
    const layout = await readDocxLayout(arrayBuffer)
    const text = (await mammoth.extractRawText({ arrayBuffer })).value
    const html = (await mammoth.convertToHtml({ arrayBuffer }, { ignoreEmptyParagraphs: false })).value
    return { text, html, layout }
  }
  if (source === 'pdf') {
    const text = await extractPdfText(file)
    return { text, html: `<pre>${escapeHtml(text)}</pre>` }
  }
  if (source === 'txt' || source === 'md' || source === 'markdown') {
    const text = await file.text()
    return { text, html: source === 'txt' ? textToEditableHtml(text) : markdownToHtml(text) }
  }
  if (source === 'html' || source === 'htm') {
    const rawHtml = await file.text()
    const parser = new DOMParser()
    const body = parser.parseFromString(rawHtml, 'text/html').body
    return { text: body.textContent ?? '', html: body.innerHTML }
  }
  throw new Error('Unsupported source file. Use DOCX, PDF, TXT, MD, HTML, PNG, or JPG.')
}

function pagesFromText(text: string) {
  return pagesFromHtml(textToEditableHtml(text))
}

function pagesFromHtml(html: string, layout: Partial<Pick<PageData, 'fontFamily' | 'margins' | 'size'>> = {}) {
  const normalizedLayout = {
    ...layout,
    margins: normalizePageMargins(layout.margins),
    size: layout.size ?? defaultPageSizeKey,
  }
  const parser = new DOMParser()
  const document = parser.parseFromString(`<main>${html || starterContent}</main>`, 'text/html')
  const root = document.body.firstElementChild
  const atoms = Array.from(root?.childNodes ?? []).flatMap(htmlAtomsFromNode)
  const pageHtmls = paginateHtmlAtoms(atoms, normalizedLayout)
  return (pageHtmls.length ? pageHtmls : [starterContent]).map((pageHtml) => ({
    id: createId(),
    fontFamily: normalizedLayout.fontFamily,
    html: pageHtml || starterContent,
    images: [],
    margins: normalizedLayout.margins,
    size: normalizedLayout.size,
  }))
}

function htmlAtomsFromNode(node: ChildNode): string[] {
  if (node.nodeType === Node.TEXT_NODE) return textToEditableHtmlAtoms(node.textContent ?? '')
  if (!(node instanceof HTMLElement)) return []
  if (!node.textContent?.trim() && node.tagName.toLowerCase() !== 'br') return []

  const tagName = node.tagName.toLowerCase()
  if (tagName === 'pre') return splitTextForPageAtoms(node.textContent ?? '', 850).map((chunk) => `<pre>${escapeHtml(chunk)}</pre>`)
  if (tagName === 'p' || tagName === 'blockquote') {
    const chunks = splitTextForPageAtoms(textWithLineBreaks(node), 1100)
    return chunks.map((chunk) => `<${tagName}>${escapeHtml(chunk).replace(/\n/g, '<br>') || '<br>'}</${tagName}>`)
  }
  if (tagName === 'h1' || tagName === 'h2' || tagName === 'h3') return [`<${tagName}>${escapeHtml(node.textContent ?? '')}</${tagName}>`]
  if (tagName === 'ul' || tagName === 'ol') return listAtomsFromNode(node, tagName)
  if (tagName === 'table') return [node.outerHTML]

  const childAtoms = Array.from(node.childNodes).flatMap(htmlAtomsFromNode)
  return childAtoms.length ? childAtoms : [node.outerHTML]
}

function textToEditableHtmlAtoms(text: string) {
  return text
    .replace(/\r\n/g, '\n')
    .split(/\n{2,}/)
    .flatMap((block) => splitTextForPageAtoms(block, 1100))
    .map((block) => `<p>${escapeHtml(block).replace(/\n/g, '<br>') || '<br>'}</p>`)
}

function listAtomsFromNode(node: HTMLElement, tagName: string) {
  const items = Array.from(node.children).filter((child) => child.tagName.toLowerCase() === 'li')
  const atoms: string[] = []
  let currentItems: string[] = []
  let currentLength = 0
  items.forEach((item) => {
    const itemHtml = item.outerHTML
    const itemLength = item.textContent?.length ?? itemHtml.length
    if (currentItems.length && currentLength + itemLength > 1000) {
      atoms.push(`<${tagName}>${currentItems.join('')}</${tagName}>`)
      currentItems = []
      currentLength = 0
    }
    currentItems.push(itemHtml)
    currentLength += itemLength
  })
  if (currentItems.length) atoms.push(`<${tagName}>${currentItems.join('')}</${tagName}>`)
  return atoms
}

function splitTextForPageAtoms(text: string, maxLength: number) {
  const normalizedText = text.replace(/\r\n/g, '\n').trim()
  if (!normalizedText) return ['']
  const chunks: string[] = []
  let current = ''
  normalizedText.split(/(\s+)/).forEach((part) => {
    if (current && current.length + part.length > maxLength) {
      chunks.push(current.trim())
      current = ''
    }
    current += part
  })
  if (current.trim()) chunks.push(current.trim())
  return chunks.length ? chunks : [normalizedText]
}

function textWithLineBreaks(element: HTMLElement) {
  return Array.from(element.childNodes)
    .map((child) => {
      if (child instanceof HTMLBRElement) return '\n'
      return child.textContent ?? ''
    })
    .join('')
}

function paginateHtmlAtoms(atoms: string[], layout: Partial<Pick<PageData, 'fontFamily' | 'margins' | 'size'>>) {
  if (!atoms.length) return [starterContent]
  const measurer = createPageMeasurer(layout)
  const pageHtmls: string[] = []
  let currentHtml = ''

  atoms.forEach((atom) => {
    const candidate = `${currentHtml}${atom}`
    if (!currentHtml || measurer.fits(candidate)) {
      currentHtml = candidate
      return
    }
    pageHtmls.push(currentHtml)
    currentHtml = atom
  })

  if (currentHtml) pageHtmls.push(currentHtml)
  measurer.remove()
  return pageHtmls
}

function createPageMeasurer(layout: Partial<Pick<PageData, 'fontFamily' | 'margins' | 'size'>>) {
  const size = pageSizeFor(layout.size)
  const margins = normalizePageMargins(layout.margins)
  const frame = document.createElement('div')
  const content = document.createElement('div')
  frame.style.position = 'fixed'
  frame.style.left = '-10000px'
  frame.style.top = '0'
  frame.style.visibility = 'hidden'
  frame.style.pointerEvents = 'none'
  frame.style.width = `${size.width}px`
  frame.style.height = `${size.height}px`
  content.className = 'static-page-content'
  content.style.width = `${size.width}px`
  content.style.height = `${size.height}px`
  content.style.boxSizing = 'border-box'
  content.style.padding = `${margins.top}px ${margins.right}px ${margins.bottom}px ${margins.left}px`
  content.style.fontFamily = fontStackFor(layout.fontFamily)
  content.style.fontSize = defaultFontSize
  content.style.lineHeight = '1.72'
  frame.appendChild(content)
  document.body.appendChild(frame)
  return {
    fits(pageHtml: string) {
      content.innerHTML = pageHtml || starterContent
      return content.scrollHeight <= content.clientHeight + 2 && content.scrollWidth <= content.clientWidth + 2
    },
    remove() {
      frame.remove()
    },
  }
}

function textToEditableHtml(text: string) {
  const blocks = text.replace(/\r\n/g, '\n').split(/\n{2,}/)
  return blocks
    .map((block) => {
      const lines = block.split('\n').map((line) => escapeHtml(line))
      return `<p>${lines.join('<br>') || '<br>'}</p>`
    })
    .join('')
}

function markdownToHtml(markdown: string) {
  const lines = markdown.replace(/\r\n/g, '\n').split('\n')
  const blocks: string[] = []
  let paragraph: string[] = []
  let listItems: string[] = []
  let orderedListItems: string[] = []

  const flushParagraph = () => {
    if (!paragraph.length) return
    blocks.push(`<p>${paragraph.map(formatInlineMarkdown).join('<br>')}</p>`)
    paragraph = []
  }
  const flushLists = () => {
    if (listItems.length) blocks.push(`<ul>${listItems.map((item) => `<li>${formatInlineMarkdown(item)}</li>`).join('')}</ul>`)
    if (orderedListItems.length) blocks.push(`<ol>${orderedListItems.map((item) => `<li>${formatInlineMarkdown(item)}</li>`).join('')}</ol>`)
    listItems = []
    orderedListItems = []
  }

  lines.forEach((line) => {
    const heading = /^(#{1,3})\s+(.+)$/.exec(line)
    const bullet = /^[-*]\s+(.+)$/.exec(line)
    const ordered = /^\d+\.\s+(.+)$/.exec(line)
    if (!line.trim()) {
      flushParagraph()
      flushLists()
      return
    }
    if (heading) {
      flushParagraph()
      flushLists()
      blocks.push(`<h${heading[1].length}>${formatInlineMarkdown(heading[2])}</h${heading[1].length}>`)
      return
    }
    if (bullet) {
      flushParagraph()
      if (orderedListItems.length) flushLists()
      listItems.push(bullet[1])
      return
    }
    if (ordered) {
      flushParagraph()
      if (listItems.length) flushLists()
      orderedListItems.push(ordered[1])
      return
    }
    flushLists()
    paragraph.push(line)
  })

  flushParagraph()
  flushLists()
  return blocks.join('') || starterContent
}

function formatInlineMarkdown(value: string) {
  return escapeHtml(value)
    .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
    .replace(/\*([^*]+)\*/g, '<em>$1</em>')
    .replace(/`([^`]+)`/g, '<code>$1</code>')
}

async function extractPdfText(file: File) {
  const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise
  const pages: string[] = []
  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber)
    const content = await page.getTextContent()
    const pageText = content.items.map((item) => ('str' in item ? item.str : '')).join(' ')
    pages.push(pageText)
  }
  return pages.join('\n\n')
}

function markdownFromHtml(html: string) {
  return new TurndownService({ headingStyle: 'atx', bulletListMarker: '-' }).turndown(html)
}

async function createDocxBlob(text: string, size: PageSizeKey = defaultPageSizeKey) {
  const pageSize = pageSizeFor(size)
  const paragraphs = text.split(/\n{2,}/).map(
    (paragraph) =>
      new Paragraph({
        children: [new TextRun(paragraph.trim() || ' ')],
        spacing: { after: 220 },
      }),
  )
  const docxDocument = new DocxDocument({
    sections: [
      {
        properties: {
          page: {
            size: {
              height: pageSize.docxHeight,
              width: pageSize.docxWidth,
            },
          },
        },
        children: paragraphs,
      },
    ],
  })
  return Packer.toBlob(docxDocument)
}

async function textToImageBlob(text: string, format: 'png' | 'jpg') {
  const canvas = document.createElement('canvas')
  const context = canvas.getContext('2d')
  if (!context) throw new Error('Canvas export is not available in this browser.')
  canvas.width = defaultPageSize.width
  canvas.height = defaultPageSize.height
  context.fillStyle = '#ffffff'
  context.fillRect(0, 0, canvas.width, canvas.height)
  context.fillStyle = '#12101d'
  context.font = `${defaultFontSize} Calibri, Carlito, Arial, sans-serif`
  wrapCanvasText(context, text || ' ', defaultPageMargins.left, defaultPageMargins.top, canvas.width - defaultPageMargins.left - defaultPageMargins.right, 28)
  return canvasToBlob(canvas, mimeForImageFormat(format), 0.95)
}

async function convertImageBlob(dataUrl: string, format: 'png' | 'jpg') {
  const image = await loadImage(dataUrl)
  const canvas = document.createElement('canvas')
  const context = canvas.getContext('2d')
  if (!context) throw new Error('Canvas conversion is not available in this browser.')
  canvas.width = image.naturalWidth
  canvas.height = image.naturalHeight
  if (format === 'jpg') {
    context.fillStyle = '#ffffff'
    context.fillRect(0, 0, canvas.width, canvas.height)
  }
  context.drawImage(image, 0, 0)
  return canvasToBlob(canvas, mimeForImageFormat(format), 0.95)
}

function wrapCanvasText(context: CanvasRenderingContext2D, text: string, x: number, y: number, maxWidth: number, lineHeight: number) {
  const words = text.replace(/\s+/g, ' ').split(' ')
  let line = ''
  words.forEach((word) => {
    const nextLine = `${line}${word} `
    if (context.measureText(nextLine).width > maxWidth && line) {
      context.fillText(line, x, y)
      line = `${word} `
      y += lineHeight
    } else {
      line = nextLine
    }
  })
  context.fillText(line, x, y)
}

function renderPagesHtml(pages: PageData[], canvasDataUrlForPage?: (pageId: string) => string | undefined) {
  return pages.map((page) => renderPageHtml(page, canvasDataUrlForPage?.(page.id))).join('')
}

function renderPageHtml(page: PageData, canvasDataUrl?: string) {
  const size = pageSizeFor(page.size)
  const margins = normalizePageMargins(page.margins)
  return `<section class="export-page" data-page-size="${page.size ?? defaultPageSizeKey}" style="--paper-width:${size.width}px;--paper-height:${size.height}px;--page-padding-top:${margins.top}px;--page-padding-right:${margins.right}px;--page-padding-bottom:${margins.bottom}px;--page-padding-left:${margins.left}px;--page-font-family:${escapeHtml(fontStackFor(page.fontFamily))};--page-font-size:${defaultFontSize};">${
    page.pageImage ? `<img class="export-page-image" src="${page.pageImage}" alt="">` : ''
  }<div class="export-page-content">${page.html}</div>${
    canvasDataUrl ? `<img class="export-canvas-layer" src="${canvasDataUrl}" alt="">` : ''
  }${page.images
    .map(
      (image) =>
        `<img class="export-floating-image" src="${image.src}" style="left:${image.x}px;top:${image.y}px;width:${image.width}px;height:${image.height}px;transform:rotate(${image.rotation}deg);" alt="">`,
    )
    .join('')}</section>`
}

function wrapHtmlDocument(body: string, title = 'AKXREDITOR Export') {
  return `<!doctype html><html><head><meta charset="utf-8"><title>${escapeHtml(title)}</title><style>${exportDocumentCss()}</style></head><body>${body}</body></html>`
}

async function exportPagesImageBlob(pages: PageData[], format: 'png' | 'jpg', canvasDataUrlForPage?: (pageId: string) => string | undefined) {
  const scale = 2
  const pageCanvases = await Promise.all(
    pages.map((page) => renderPageToCanvas(page, canvasDataUrlForPage?.(page.id), scale, format)),
  )
  return canvasToBlob(stitchPageCanvases(pageCanvases, format), mimeForImageFormat(format), 0.95)
}

async function exportPageImageBlobs(pages: PageData[], format: 'png' | 'jpg', canvasDataUrlForPage?: (pageId: string) => string | undefined) {
  const scale = 2
  const pageCanvases = await Promise.all(
    pages.map((page) => renderPageToCanvas(page, canvasDataUrlForPage?.(page.id), scale, format)),
  )
  return Promise.all(pageCanvases.map((canvas) => canvasToBlob(canvas, mimeForImageFormat(format), 0.95)))
}

async function exportPagesPdfBlob(pages: PageData[], canvasDataUrlForPage?: (pageId: string) => string | undefined) {
  const scale = 2
  const safePages = pages.length ? pages : [{ id: createId(), html: starterContent, images: [], size: defaultPageSizeKey }]
  const pageCanvases = await Promise.all(
    safePages.map((page) => renderPageToCanvas(page, canvasDataUrlForPage?.(page.id), scale, 'jpg')),
  )
  const firstSize = pageSizeFor(safePages[0].size)
  const pdf = new jsPDF({
    compress: true,
    format: [firstSize.width, firstSize.height],
    orientation: firstSize.orientation,
    unit: 'px',
  })

  pageCanvases.forEach((canvas, index) => {
    const page = safePages[index]
    const size = pageSizeFor(page.size)
    if (index > 0) pdf.addPage([size.width, size.height], size.orientation)
    pdf.addImage(canvas.toDataURL('image/jpeg', 0.95), 'JPEG', 0, 0, size.width, size.height)
  })

  return pdf.output('blob')
}

async function renderPageToCanvas(page: PageData, canvasDataUrl: string | undefined, scale: number, format: 'png' | 'jpg') {
  const size = pageSizeFor(page.size)
  const svgBody = xhtmlForSvg(`<style>${exportDocumentCss()}html,body{background:#fff;}.export-page{margin:0;box-shadow:none;}</style>${renderPageHtml(page, canvasDataUrl)}`)
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${size.width * scale}" height="${size.height * scale}" viewBox="0 0 ${size.width} ${size.height}"><foreignObject width="${size.width}" height="${size.height}">${svgBody}</foreignObject></svg>`
  const image = await loadImage(`data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`)
  const canvas = document.createElement('canvas')
  const context = canvas.getContext('2d')
  if (!context) throw new Error('Canvas export is not available in this browser.')
  canvas.width = size.width * scale
  canvas.height = size.height * scale
  if (format === 'jpg') {
    context.fillStyle = '#ffffff'
    context.fillRect(0, 0, canvas.width, canvas.height)
  }
  context.drawImage(image, 0, 0, canvas.width, canvas.height)
  return canvas
}

function stitchPageCanvases(canvases: HTMLCanvasElement[], format: 'png' | 'jpg') {
  const gap = 40
  const width = Math.max(...canvases.map((canvas) => canvas.width))
  const height = canvases.reduce((sum, canvas) => sum + canvas.height, 0) + gap * Math.max(canvases.length - 1, 0)
  const output = document.createElement('canvas')
  const context = output.getContext('2d')
  if (!context) throw new Error('Canvas export is not available in this browser.')
  output.width = width
  output.height = height
  context.fillStyle = format === 'jpg' ? '#ffffff' : 'rgba(255,255,255,0)'
  context.fillRect(0, 0, output.width, output.height)
  let y = 0
  canvases.forEach((canvas) => {
    context.drawImage(canvas, 0, y)
    y += canvas.height + gap
  })
  return output
}

function xhtmlForSvg(html: string) {
  const wrapper = document.createElement('div')
  wrapper.setAttribute('xmlns', 'http://www.w3.org/1999/xhtml')
  wrapper.innerHTML = html
  return new XMLSerializer().serializeToString(wrapper)
}

function exportDocumentCss() {
  return `
@font-face{font-family:Carlito;src:local("Carlito"),local("Calibri");}
@font-face{font-family:Caladea;src:local("Caladea"),local("Cambria");}
@page{size:A4;margin:0;}
@page akxr-a4{size:A4;margin:0;}
@page akxr-a5{size:A5;margin:0;}
@page akxr-a4-landscape{size:A4 landscape;margin:0;}
@page akxr-a5-landscape{size:A5 landscape;margin:0;}
*{box-sizing:border-box;}
html,body{margin:0;min-height:100%;background:#f3f4f6;color:#171224;}
body{font-family:Calibri,Carlito,Arial,sans-serif;line-height:1.72;}
img{max-width:100%;}
table{width:100%;border-collapse:collapse;}
td,th{min-width:60px;border:1px solid #c9c4d6;padding:8px;}
th{background:#f2efff;}
.export-page{isolation:isolate;position:relative;width:var(--paper-width);height:var(--paper-height);margin:32px auto;overflow:hidden;background:#fff;color:#171224;box-shadow:0 18px 50px rgba(15,23,42,.2);page-break-after:always;break-after:page;}
.export-page[data-page-size="a4"]{page:akxr-a4;}
.export-page[data-page-size="a5"]{page:akxr-a5;}
.export-page[data-page-size="a4Landscape"]{page:akxr-a4-landscape;}
.export-page[data-page-size="a5Landscape"]{page:akxr-a5-landscape;}
.export-page-image,.export-canvas-layer{position:absolute;inset:0;width:100%;height:100%;object-fit:fill;pointer-events:none;}
.export-page-image{z-index:0;}
.export-page-content{position:relative;z-index:1;height:100%;padding:var(--page-padding-top) var(--page-padding-right) var(--page-padding-bottom) var(--page-padding-left);overflow:hidden;font-family:var(--page-font-family);font-size:var(--page-font-size);line-height:1.72;text-align:left;}
.export-canvas-layer{z-index:2;}
.export-floating-image{position:absolute;z-index:3;object-fit:contain;transform-origin:center center;}
.export-page-content p{margin:0 0 .774em;}
.export-page-content p:last-child{margin-bottom:0;}
.export-page-content ul,.export-page-content ol{margin:0 0 .85em;padding-left:1.55em;list-style-position:outside;}
.export-page-content li{margin:.2em 0;padding-left:.15em;}
.export-page-content li p{margin:.15em 0;}
.export-page-content h1,.export-page-content h2,.export-page-content h3{margin:0 0 .65em;line-height:1.12;}
.export-page-content h1:not(:first-child),.export-page-content h2:not(:first-child),.export-page-content h3:not(:first-child){margin-top:1em;}
.export-page-content h1{font-size:42px;}
.export-page-content h2{font-size:28px;}
.export-page-content blockquote{margin:0 0 .85em;padding-left:1em;border-left:3px solid #d3cce3;}
.export-page-content pre{margin:0 0 .85em;white-space:pre-wrap;}
@media print{html,body{background:#fff;}.export-page{margin:0;box-shadow:none;}.export-page:last-child{break-after:auto;page-break-after:auto;}}
`
}

function saveBlob(blob: Blob, fileName: string) {
  const url = URL.createObjectURL(blob)
  saveDataUrl(url, fileName)
  window.setTimeout(() => URL.revokeObjectURL(url), 300)
}

function saveDataUrl(url: string, fileName: string) {
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = fileName
  anchor.click()
}

function canvasToBlob(canvas: HTMLCanvasElement, type: string, quality?: number) {
  return new Promise<Blob>((resolve, reject) => {
    canvas.toBlob((blob) => (blob ? resolve(blob) : reject(new Error('Could not export the canvas.'))), type, quality)
  })
}

function loadImage(src: string) {
  return new Promise<HTMLImageElement>((resolve, reject) => {
    const image = new window.Image()
    image.onload = () => resolve(image)
    image.onerror = () => reject(new Error('Could not load that image.'))
    image.src = src
  })
}

function fileToDataUrl(file: File) {
  return new Promise<string>((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(String(reader.result))
    reader.onerror = () => reject(new Error('Could not read that file.'))
    reader.readAsDataURL(file)
  })
}

function bindPointerDrag(move: (event: PointerEvent) => void) {
  const end = () => {
    window.removeEventListener('pointermove', move)
    window.removeEventListener('pointerup', end)
  }
  window.addEventListener('pointermove', move)
  window.addEventListener('pointerup', end)
}

function removeCanvasObject(canvas: FabricCanvas, object: FabricObject) {
  if (object.type === 'activeselection' && 'getObjects' in object && typeof object.getObjects === 'function') {
    object.getObjects().forEach((item: FabricObject) => canvas.remove(item))
    return
  }
  canvas.remove(object)
}

function isTypingTarget(target: EventTarget | null) {
  if (!(target instanceof HTMLElement)) return false
  if (target.isContentEditable) return true
  const tagName = target.tagName.toLowerCase()
  return tagName === 'input' || tagName === 'textarea' || tagName === 'select'
}

function scrollToPage(pageId: string) {
  document.getElementById(`page-${pageId}`)?.scrollIntoView({ behavior: 'smooth', block: 'center' })
}

function keepObjectInsideCanvas(object: FabricObject, canvas: FabricCanvas) {
  object.setCoords()
  const bounds = object.getBoundingRect()
  let nextLeft = object.left ?? 0
  let nextTop = object.top ?? 0
  if (bounds.left < 0) nextLeft -= bounds.left
  if (bounds.top < 0) nextTop -= bounds.top
  if (bounds.left + bounds.width > canvas.getWidth()) nextLeft -= bounds.left + bounds.width - canvas.getWidth()
  if (bounds.top + bounds.height > canvas.getHeight()) nextTop -= bounds.top + bounds.height - canvas.getHeight()
  object.set({ left: nextLeft, top: nextTop })
  object.setCoords()
}

function canvasSnapshot(canvas: FabricCanvas) {
  return JSON.stringify(canvas.toObject(canvasObjectCustomProperties))
}

function positionClonedCanvasObject(object: FabricObject, pastePoint?: CanvasPoint) {
  const objects = object instanceof ActiveSelection ? object.getObjects() : [object]
  if (!objects.length) return []
  const bounds = object.getBoundingRect()
  const offsetX = pastePoint ? pastePoint.x - bounds.left : 24
  const offsetY = pastePoint ? pastePoint.y - bounds.top : 24
  objects.forEach((item) => {
    item.set({
      left: Number(item.left ?? 0) + offsetX,
      top: Number(item.top ?? 0) + offsetY,
    })
    item.setCoords()
  })
  return objects
}

function restorePastedCanvasObject(object: FabricObject, canvas: FabricCanvas) {
  sharpenCanvasText(object)
  if (isHorizontalTableGroup(object)) {
    configureHorizontalTableGroup(object, canvas)
    return
  }
  if (object instanceof IText) {
    const mode = (object as ManagedIText).akxrTextMode
    if (mode) configureManagedText(object, mode, canvas, { removeIfEmpty: mode === 'plain' })
  }
  if (object instanceof Group) {
    object.getObjects().forEach((item) => restorePastedCanvasObject(item, canvas))
  }
}

function sharpenCanvasText(object: FabricObject, options: { snapPosition?: boolean } = {}) {
  if (object instanceof IText) {
    object.set({
      noScaleCache: true,
      objectCaching: false,
    })
    if (options.snapPosition) {
      object.set({
        left: Math.round(Number(object.left ?? 0)),
        top: Math.round(Number(object.top ?? 0)),
      })
    }
    object.initDimensions()
    object.setCoords()
    return
  }

  if (object instanceof Group) {
    object.getObjects().forEach((item) => sharpenCanvasText(item))
  }
}

function restoreManagedCanvasObjects(canvas: FabricCanvas) {
  canvas.getObjects().forEach((object) => {
    sharpenCanvasText(object)
    if (isHorizontalTableGroup(object)) configureHorizontalTableGroup(object, canvas)
  })
}

function configureManagedText(
  text: IText,
  mode: MovableTextMode,
  canvas: FabricCanvas,
  options: { removeIfEmpty?: boolean } = {},
) {
  const managed = text as ManagedIText
  managed.akxrTextMode = mode
  sharpenCanvasText(text, { snapPosition: true })

  text.on('changed', () => {
    if (managed.akxrFormattingText) return
    const currentText = text.text ?? ''
    const formattedText = formatManagedText(currentText, mode)
    if (formattedText === currentText) return
    const cursorAtEnd = (text.selectionStart ?? 0) >= currentText.length - 1
    const previousCursor = text.selectionStart ?? formattedText.length
    const cursorDelta = formattedText.length - currentText.length
    managed.akxrFormattingText = true
    text.set({ text: formattedText })
    text.initDimensions()
    text.setCoords()
    if (cursorAtEnd) {
      moveTextCursorToEnd(text)
    } else {
      const nextCursor = clamp(previousCursor + cursorDelta, 0, formattedText.length)
      text.selectionStart = nextCursor
      text.selectionEnd = nextCursor
      syncFabricTextarea(text)
    }
    canvas.requestRenderAll()
    managed.akxrFormattingText = false
  })

  text.on('editing:exited', () => {
    if (!options.removeIfEmpty || !isEmptyManagedText(text, mode)) return
    canvas.remove(text)
    canvas.discardActiveObject()
    canvas.requestRenderAll()
  })
}

function formatManagedText(value: string, mode: MovableTextMode) {
  if (mode === 'plain') return value
  const lines = value.split('\n')

  if (mode === 'bullet') {
    return lines.map((line) => `- ${listLineContent(line)}`).join('\n')
  }

  if (mode === 'numbered') {
    return lines.map((line, index) => `${index + 1}. ${listLineContent(line)}`).join('\n')
  }

  return lines.map((line) => normalizeTableLine(line)).join('\n')
}

function listLineContent(line: string) {
  return line.replace(/^\s*(?:[-*]\s+|\d+[.)]\s+|•\s*)/, '')
}

function normalizeTableLine(line: string) {
  const trimmedLine = line.trim()
  if (!trimmedLine) return '| Cell | Cell | Cell |'
  if (trimmedLine.includes('|')) return trimmedLine.startsWith('|') ? trimmedLine : `| ${trimmedLine}`
  return `| ${trimmedLine} | Cell | Cell |`
}

function isEmptyManagedText(text: IText, mode: MovableTextMode) {
  const value = (text.text ?? '').trim()
  if (!value || value === 'Type here') return true
  if (mode === 'bullet') return value.split('\n').every((line) => !listLineContent(line).trim())
  if (mode === 'numbered') return value.split('\n').every((line) => !listLineContent(line).trim())
  return false
}

function nextMovableTextPosition(canvas: FabricCanvas) {
  const textCount = canvas.getObjects().filter((object) => object instanceof IText || isHorizontalTableGroup(object)).length
  const step = textCount % 8
  return {
    left: 130 + step * 34,
    top: 155 + step * 30,
  }
}

function createHorizontalTableGroup({
  cellCount,
  cellWidth = 150,
  fill,
  fontFamily,
  fontSize,
  fontStyle,
  fontWeight,
  left,
  top,
  underline,
}: {
  cellCount: number
  cellWidth?: number
  fill: string
  fontFamily: string
  fontSize: number
  fontStyle: 'normal' | 'italic'
  fontWeight: 'normal' | 'bold'
  left: number
  top: number
  underline: boolean
}) {
  const cellHeight = 58
  const objects: FabricObject[] = []
  for (let index = 0; index < cellCount; index += 1) {
    objects.push(
      new Rect({
        fill: '#ffffff',
        height: cellHeight,
        left: index * cellWidth,
        rx: 3,
        ry: 3,
        stroke: '#111827',
        strokeWidth: 2,
        top: 0,
        width: cellWidth,
      }),
    )
    objects.push(
      new Textbox('', {
        editable: true,
        evented: true,
        fill,
        fontFamily,
        fontSize,
        fontStyle,
        fontWeight,
        left: index * cellWidth + cellWidth / 2,
        minWidth: cellWidth - 28,
        noScaleCache: true,
        objectCaching: false,
        originX: 'center',
        originY: 'center',
        selectable: true,
        splitByGrapheme: true,
        textAlign: 'center',
        top: cellHeight / 2,
        underline,
        width: cellWidth - 28,
      }),
    )
  }
  const table = new Group(objects, {
    borderColor: '#2563eb',
    cornerColor: '#2563eb',
    cornerStyle: 'circle',
    interactive: true,
    left,
    lockScalingFlip: true,
    subTargetCheck: true,
    top,
  }) as HorizontalTableGroup
  table.akxrObjectType = 'horizontal-table'
  table.akxrTableCellCount = cellCount
  table.akxrTableCellWidth = cellWidth
  return table
}

function configureHorizontalTableGroup(table: HorizontalTableGroup, canvas: FabricCanvas) {
  table.akxrObjectType = 'horizontal-table'
  table.akxrTableCellCount = Math.max(table.akxrTableCellCount ?? horizontalTableTextItems(table).length, 1)
  table.akxrTableCellWidth = table.akxrTableCellWidth ?? horizontalTableCellWidth(table)
  horizontalTableTextItems(table).forEach((text) => sharpenCanvasText(text))
  table.controls = {
    ...table.controls,
    addCell: new Control({
      actionName: 'add-table-cell',
      cursorStyle: 'pointer',
      mouseUpHandler: (_eventData, transform) => {
        if (!isHorizontalTableGroup(transform.target)) return false
        addHorizontalTableCell(transform.target, canvas)
        return true
      },
      offsetX: 30,
      sizeX: 26,
      sizeY: 26,
      touchSizeX: 34,
      touchSizeY: 34,
      x: 0.5,
      y: 0,
      render: renderAddCellControl,
    }),
  }
}

function addHorizontalTableCell(table: HorizontalTableGroup, canvas: FabricCanvas) {
  const texts = horizontalTableTextItems(table)
  const firstText = texts[0]
  const nextTable = createHorizontalTableGroup({
    cellWidth: table.akxrTableCellWidth ?? horizontalTableCellWidth(table),
    cellCount: Math.max(table.akxrTableCellCount ?? texts.length, 1) + 1,
    fill: String(firstText?.fill ?? '#111827'),
    fontFamily: String(firstText?.fontFamily ?? 'Inter'),
    fontSize: Number(firstText?.fontSize ?? 18),
    fontStyle: firstText?.fontStyle === 'italic' ? 'italic' : 'normal',
    fontWeight: firstText?.fontWeight === 'bold' ? 'bold' : 'normal',
    left: table.left ?? 130,
    top: table.top ?? 155,
    underline: Boolean(firstText?.underline),
  })
  horizontalTableTextItems(nextTable).forEach((text, index) => {
    text.set({ text: texts[index]?.text ?? '' })
  })
  nextTable.set({
    angle: table.angle,
    scaleX: table.scaleX,
    scaleY: table.scaleY,
  })
  configureHorizontalTableGroup(nextTable, canvas)
  canvas.remove(table)
  canvas.add(nextTable)
  canvas.setActiveObject(nextTable)
  nextTable.setCoords()
  keepObjectInsideCanvas(nextTable, canvas)
  canvas.requestRenderAll()
}

function deleteHorizontalTableCell(table: HorizontalTableGroup, cellIndex: number, canvas: FabricCanvas) {
  const texts = horizontalTableTextItems(table)
  if (cellIndex < 0 || cellIndex >= texts.length) return
  const remainingTexts = texts.filter((_, index) => index !== cellIndex)
  if (!remainingTexts.length) {
    canvas.remove(table)
    canvas.discardActiveObject()
    return
  }

  const firstText = remainingTexts[0]
  const nextTable = createHorizontalTableGroup({
    cellWidth: table.akxrTableCellWidth ?? horizontalTableCellWidth(table),
    cellCount: remainingTexts.length,
    fill: String(firstText.fill ?? '#111827'),
    fontFamily: String(firstText.fontFamily ?? 'Inter'),
    fontSize: Number(firstText.fontSize ?? 18),
    fontStyle: firstText.fontStyle === 'italic' ? 'italic' : 'normal',
    fontWeight: firstText.fontWeight === 'bold' ? 'bold' : 'normal',
    left: table.left ?? 130,
    top: table.top ?? 155,
    underline: Boolean(firstText.underline),
  })
  const nextTexts = horizontalTableTextItems(nextTable)
  remainingTexts.forEach((text, index) => {
    nextTexts[index]?.set({
      fill: text.fill,
      fontFamily: text.fontFamily,
      fontSize: text.fontSize,
      fontStyle: text.fontStyle,
      fontWeight: text.fontWeight,
      text: text.text ?? '',
      underline: text.underline,
    })
    nextTexts[index]?.initDimensions()
  })
  nextTable.set({
    angle: table.angle,
    scaleX: table.scaleX,
    scaleY: table.scaleY,
  })
  configureHorizontalTableGroup(nextTable, canvas)
  canvas.remove(table)
  canvas.add(nextTable)
  canvas.setActiveObject(nextTable)
  nextTable.setCoords()
  keepObjectInsideCanvas(nextTable, canvas)
}

function horizontalTableTextItems(table: HorizontalTableGroup) {
  return table.getObjects().filter((object): object is Textbox => object instanceof Textbox)
}

function fillHorizontalTableCell(table: HorizontalTableGroup, cellIndex: number, fill: string) {
  const cell = table.getObjects()[cellIndex * 2]
  if (cell instanceof Rect) cell.set({ fill })
  table.set({ dirty: true })
}

function applyPaintBucketFill(object: FabricObject, fill: string) {
  if (object instanceof Group) {
    object.getObjects().forEach((child) => applyPaintBucketFill(child, fill))
    object.set({ dirty: true })
    return
  }
  if (object instanceof Line) {
    object.set({ stroke: fill })
    return
  }
  object.set({ fill })
}

async function floodFillImageDataUrl(dataUrl: string, pageX: number, pageY: number, pageWidth: number, pageHeight: number, fill: string) {
  const image = await loadImage(dataUrl)
  const canvas = document.createElement('canvas')
  const context = canvas.getContext('2d')
  if (!context) throw new Error('Canvas rendering is not available in this browser.')
  canvas.width = image.naturalWidth
  canvas.height = image.naturalHeight
  context.drawImage(image, 0, 0)

  const x = clamp(Math.floor((pageX / pageWidth) * canvas.width), 0, canvas.width - 1)
  const y = clamp(Math.floor((pageY / pageHeight) * canvas.height), 0, canvas.height - 1)
  const fillColor = hexColorToRgba(fill)
  const imageData = context.getImageData(0, 0, canvas.width, canvas.height)
  const data = imageData.data
  const startIndex = (y * canvas.width + x) * 4
  const targetColor = [data[startIndex], data[startIndex + 1], data[startIndex + 2], data[startIndex + 3]]
  if (colorsClose(targetColor, fillColor, 0)) return dataUrl

  const tolerance = 28
  const pixelCount = canvas.width * canvas.height
  const visited = new Uint8Array(pixelCount)
  const stack = new Int32Array(pixelCount)
  let stackLength = 1
  let changed = false
  stack[0] = y * canvas.width + x
  visited[stack[0]] = 1

  while (stackLength > 0) {
    stackLength -= 1
    const pixel = stack[stackLength]
    const dataIndex = pixel * 4
    const currentColor = [data[dataIndex], data[dataIndex + 1], data[dataIndex + 2], data[dataIndex + 3]]
    if (!colorsClose(currentColor, targetColor, tolerance)) continue

    data[dataIndex] = fillColor[0]
    data[dataIndex + 1] = fillColor[1]
    data[dataIndex + 2] = fillColor[2]
    data[dataIndex + 3] = fillColor[3]
    changed = true

    const px = pixel % canvas.width
    const left = pixel - 1
    const right = pixel + 1
    const up = pixel - canvas.width
    const down = pixel + canvas.width
    if (px > 0 && !visited[left]) {
      visited[left] = 1
      stack[stackLength] = left
      stackLength += 1
    }
    if (px < canvas.width - 1 && !visited[right]) {
      visited[right] = 1
      stack[stackLength] = right
      stackLength += 1
    }
    if (pixel >= canvas.width && !visited[up]) {
      visited[up] = 1
      stack[stackLength] = up
      stackLength += 1
    }
    if (pixel < canvas.width * (canvas.height - 1) && !visited[down]) {
      visited[down] = 1
      stack[stackLength] = down
      stackLength += 1
    }
  }

  if (!changed) return dataUrl
  context.putImageData(imageData, 0, 0)
  return canvas.toDataURL('image/png')
}

function hexColorToRgba(color: string) {
  const hex = color.trim().replace(/^#/, '')
  if (/^[0-9a-f]{6}$/i.test(hex)) {
    return [
      Number.parseInt(hex.slice(0, 2), 16),
      Number.parseInt(hex.slice(2, 4), 16),
      Number.parseInt(hex.slice(4, 6), 16),
      255,
    ]
  }
  return [0, 0, 0, 255]
}

function colorsClose(a: number[], b: number[], tolerance: number) {
  return (
    Math.abs(a[0] - b[0]) <= tolerance &&
    Math.abs(a[1] - b[1]) <= tolerance &&
    Math.abs(a[2] - b[2]) <= tolerance &&
    Math.abs(a[3] - b[3]) <= tolerance
  )
}

function horizontalTableCellIndexFromTargets(table: HorizontalTableGroup, targets: FabricObject[]) {
  const objects = table.getObjects()
  const target = targets.find((item) => objects.includes(item))
  if (!target) return null
  const cellIndex = Math.floor(objects.indexOf(target) / 2)
  const cellCount = table.akxrTableCellCount ?? horizontalTableTextItems(table).length
  return cellIndex >= 0 && cellIndex < cellCount ? cellIndex : null
}

function horizontalTableCellIndexFromPointer(table: HorizontalTableGroup, pointer: Point, targets: FabricObject[]) {
  const targetCellIndex = horizontalTableCellIndexFromTargets(table, targets)
  if (targetCellIndex != null) return targetCellIndex

  const cellWidth = table.akxrTableCellWidth ?? horizontalTableCellWidth(table)
  const cellCount = table.akxrTableCellCount ?? horizontalTableTextItems(table).length
  const localPoint = fabricUtil.sendPointToPlane(pointer, undefined, table.calcTransformMatrix())
  const leftEdge = -((cellWidth * cellCount) / 2)
  const cellIndex = Math.floor((localPoint.x - leftEdge) / cellWidth)
  return localPoint.y >= -29 && localPoint.y <= 29 && cellIndex >= 0 && cellIndex < cellCount ? cellIndex : null
}

function horizontalTableCellWidth(table: HorizontalTableGroup) {
  const firstCell = table.getObjects().find((object) => object instanceof Rect) as Rect | undefined
  return Number(firstCell?.width ?? 150)
}

function isHorizontalTableGroup(object: FabricObject | undefined | null): object is HorizontalTableGroup {
  return object instanceof Group && (object as HorizontalTableGroup).akxrObjectType === 'horizontal-table'
}

function horizontalTableFromTarget(object: FabricObject | undefined | null): HorizontalTableGroup | null {
  if (isHorizontalTableGroup(object)) return object
  const group = (object as (FabricObject & { group?: FabricObject }) | undefined)?.group
  return isHorizontalTableGroup(group) ? group : null
}

function renderAddCellControl(ctx: CanvasRenderingContext2D, left: number, top: number) {
  ctx.save()
  ctx.translate(left, top)
  ctx.fillStyle = '#2563eb'
  ctx.strokeStyle = '#ffffff'
  ctx.lineWidth = 2
  ctx.beginPath()
  ctx.arc(0, 0, 13, 0, Math.PI * 2)
  ctx.fill()
  ctx.beginPath()
  ctx.moveTo(-6, 0)
  ctx.lineTo(6, 0)
  ctx.moveTo(0, -6)
  ctx.lineTo(0, 6)
  ctx.stroke()
  ctx.restore()
}

function moveTextCursorToEnd(text: IText) {
  const end = text.text?.length ?? 0
  text.selectionStart = end
  text.selectionEnd = end
  syncFabricTextarea(text)
}

function syncFabricTextarea(text: IText) {
  ;(text as unknown as { _updateTextarea?: () => void })._updateTextarea?.()
}

function textFromHtml(html: string) {
  const parser = new DOMParser()
  return parser.parseFromString(html, 'text/html').body.textContent ?? ''
}

function pageSizeFor(size: PageSizeKey = defaultPageSizeKey) {
  return pageSizes[size] ?? defaultPageSize
}

function pageSizeTranslationKey(size: PageSizeKey = defaultPageSizeKey): TranslationKey {
  if (size === 'a5') return 'pageSizeA5'
  if (size === 'a4Landscape') return 'pageSizeA4Landscape'
  if (size === 'a5Landscape') return 'pageSizeA5Landscape'
  return 'pageSizeA4'
}

function pageStyle(page: PageData) {
  const size = pageSizeFor(page.size)
  const margins = normalizePageMargins(page.margins)
  return {
    '--paper-width': `${size.width}px`,
    '--paper-height': `${size.height}px`,
    '--page-font-family': fontStackFor(page.fontFamily),
    '--page-font-size': defaultFontSize,
    '--page-padding-bottom': `${margins.bottom}px`,
    '--page-padding-left': `${margins.left}px`,
    '--page-padding-right': `${margins.right}px`,
    '--page-padding-top': `${margins.top}px`,
  } as React.CSSProperties
}

function normalizePageMargins(margins?: Partial<PageMargins>): PageMargins {
  return exactPageMargins(margins)
}

function exactPageMargins(margins?: Partial<PageMargins>): PageMargins {
  return {
    top: pageMarginValue(margins?.top, defaultPageMargins.top),
    right: pageMarginValue(margins?.right, defaultPageMargins.right),
    bottom: pageMarginValue(margins?.bottom, defaultPageMargins.bottom),
    left: pageMarginValue(margins?.left, defaultPageMargins.left),
  }
}

function pageMarginValue(value: number | undefined, fallback: number) {
  if (typeof value !== 'number' || !Number.isFinite(value)) return fallback
  return Math.max(0, Math.round(value))
}

function normalizeFontSize(size: string) {
  const numericSize = Number.parseFloat(size)
  if (!Number.isFinite(numericSize) || numericSize <= 0) return ''
  return `${clamp(Math.round(numericSize), 1, 400)}px`
}

function baseNameFromFileName(fileName: string) {
  return safeFileBaseName(fileName.replace(/\.[^.]+$/, ''))
}

function exportFileNameFor(baseName: string, format: ExportFormat) {
  return `${safeFileBaseName(baseName)}.${format}`
}

function exportImagePageFileNameFor(baseName: string, pageNumber: number, format: 'png' | 'jpg') {
  return `${safeFileBaseName(baseName)}-page-${pageNumber}.${format}`
}

function safeFileBaseName(fileName: string) {
  const illegalCharacters = new Set(['<', '>', ':', '"', '/', '\\', '|', '?', '*'])
  const sanitizedName = fileName
    .replace(/\.[a-z0-9]{1,6}$/i, '')
    .split('')
    .filter((character) => character.charCodeAt(0) >= 32 && !illegalCharacters.has(character))
    .join('')
    .trim()
  return sanitizedName || 'akxreditor-document'
}

function fontStackFor(fontFamily = 'Calibri') {
  const normalizedFamily = fontFamily.toLowerCase()
  if (normalizedFamily.includes('cambria')) return 'Cambria, Caladea, Georgia, serif'
  if (normalizedFamily.includes('calibri')) return 'Calibri, Carlito, Arial, sans-serif'
  if (normalizedFamily.includes('courier')) return '"Courier New", Courier, monospace'
  if (normalizedFamily.includes('times')) return '"Times New Roman", Times, serif'
  if (normalizedFamily.includes('georgia')) return 'Georgia, Cambria, Caladea, serif'
  return `${fontFamily}, Calibri, Carlito, Arial, sans-serif`
}

function clonePages(pages: PageData[]) {
  return pages.map((page) => ({
    ...page,
    size: page.size ?? defaultPageSizeKey,
    images: page.images.map((image) => ({ ...image })),
  }))
}

function extensionOf(fileName: string) {
  return fileName.split('.').pop()?.toLowerCase() ?? ''
}

function mimeForImageFormat(format: 'png' | 'jpg') {
  return format === 'jpg' ? 'image/jpeg' : 'image/png'
}

function clamp(value: number, min: number, max: number) {
  return Math.min(Math.max(value, min), max)
}

function createId() {
  return crypto.randomUUID()
}

function escapeHtml(value: string) {
  return value.replace(/[&<>"']/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;' })[char] ?? char)
}

export default App
