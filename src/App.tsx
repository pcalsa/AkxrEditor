import { HelpHint } from './HelpHint'
import { PageLibrary } from './PageLibrary'
import { AreaEraser } from './AreaEraser'
import { InkCutMark, eraseImage, eraseInlineImages, eraseTextTransaction, prepareDrawingErase, parseInkCut, restoreInkStyle, fullyErased, maskedGlyphBlob, type EraseStroke } from './ink'
import { translate, languages, type AppLanguage, type TranslationKey, type Translate } from './i18n'
import { useCallback, useEffect, useMemo, useRef, useState, type PointerEvent as ReactPointerEvent } from 'react'
import { EditorContent, useEditor, type Editor } from '@tiptap/react'
import { Mark, mergeAttributes } from '@tiptap/core'
import StarterKit from '@tiptap/starter-kit'
import TextAlign from '@tiptap/extension-text-align'
import { Table, TableCell, TableHeader, TableRow } from '@tiptap/extension-table'
import { Color } from '@tiptap/extension-color'
import { FontFamily } from '@tiptap/extension-font-family'
import { TextStyle, FontSize } from '@tiptap/extension-text-style'
import { PressureBrush } from './PressureBrush'
import ImageExtension from '@tiptap/extension-image'
import DOMPurify from 'dompurify'
import { marked } from 'marked'
import { readDraft, writeDraft, localVersions, StorageConflict } from './storage'
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
  ZoomIn,
  ZoomOut,
  Move,
  PaintBucket,
  Paintbrush,
  PenLine,
  Plus,
  RectangleHorizontal,
  Redo2,
  RotateCw,
  Save,
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
  StaticCanvas,
  Point,
  Rect,
  Textbox,
  Triangle,
  util as fabricUtil,
  type FabricObject,
} from 'fabric'
import TurndownService from 'turndown'
import './App.css'

async function pdfLibrary() {
  const pdf = await import('pdfjs-dist')
  pdf.GlobalWorkerOptions.workerSrc = (await import('pdfjs-dist/build/pdf.worker.mjs?url')).default
  return pdf
}

type Mode = 'editor' | 'convert'
type ExportFormat = 'pdf' | 'docx' | 'txt' | 'md' | 'html' | 'png' | 'jpg'
type EditorTool = 'write' | 'select-text' | 'select-paint' | 'brush' | 'eraser' | 'area-eraser' | 'bucket' | 'place-text'
type LayerKind = 'text' | 'paint' | 'images'
type PageSizeKey = 'a4' | 'a5' | 'a4Landscape' | 'a5Landscape'
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
export type FloatingImage = {
  alt?: string
  id: string
  src: string
  x: number
  y: number
  width: number
  height: number
  rotation: number
}
export type PageData = {
  direction?: 'rtl' | 'ltr'
  title?: string
  section?: string
  tags?: string
  createdAt?: string
  updatedAt?: string
  attachments?: { id: string; name: string; type: string; data: string }[]
  lineHeight?: number
  layerOrder?: LayerKind[]
  width?: number
  height?: number
  baseHeight?: number
  canvasJson?: string
  id: string
  html: string
  images: FloatingImage[]
  margins?: PageMargins
  fontFamily?: string
  pageImage?: string
  size?: PageSizeKey
}
export type SavedProject = {
  trash?: PageData[]
  activePageId?: string
  language?: AppLanguage
  pages: PageData[]
  version: 1
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
function localizedErrorMessage(error: unknown, t: Translate, fallbackKey: TranslationKey) {
  const message = error instanceof Error ? error.message : typeof error === 'string' ? error : ''
  const key = errorMessageTranslationKey(message)
  if(message.includes('PDF exceeds'))return t('pdfLimit')
  if(message.includes('exceeds') || message==='renderLimit')return t('renderLimit')
  if(message.includes('No extractable PDF text'))return t('noPdfText')
  if(message.includes('Import cancelled') || message==='eraseStale')return t('importStale')
  return key ? t(key) : t(fallbackKey)
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
const autosaveKey = 'akxr-editor-autosave-v1'
const pendingCanvasWrites = new Set<() => void>()

function App() {
  const [restoredDraft] = useState(() => readAutosavedProject())
  const [language, rawSetLanguage] = useState<AppLanguage>(restoredDraft?.language ?? 'de')
  const [offlineReady, setOfflineReady] = useState(false)
  useEffect(() => {
    if ('serviceWorker' in navigator) void navigator.serviceWorker.ready.then(() => setOfflineReady(true))
  }, [])
  const [mode, setMode] = useState<Mode>('editor')
  const [pages, rawSetPages] = useState<PageData[]>(() => restoredDraft?.pages ?? [{ id: createId(), html: starterContent, images: [], size: defaultPageSizeKey }])
  const [activePageId, setActivePageId] = useState(() => restoredDraft?.activePageId ?? restoredDraft?.pages[0]?.id ?? pages[0].id)
  const activePageIdRef = useRef(activePageId)
  const pagesRef = useRef<HTMLDivElement>(null)
  const pageDataRef = useRef(pages)
  const [trash, setTrash] = useState<PageData[]>(restoredDraft?.trash ?? [])
  const [storageStatus, setStorageStatus] = useState('loading')
  const [ready, setReady] = useState(false)
  const [query, setQuery] = useState('')
  const [remoteChanged, setRemoteChanged] = useState(false)
  const [importing, setImporting] = useState(false)
  const importController = useRef<AbortController | null>(null)
  const revision = useRef(0)
  const changeNumber = useRef(0)
  const lastCheckpoint = useRef(0)
  const saveQueue = useRef(Promise.resolve())
  const trashRef = useRef(trash)
  trashRef.current = trash
  const pageHistory = useRef<{ pages: PageData[]; trash: PageData[] }[]>([])
  const textGrouping = useRef({ pageId: '', time: 0 })
  const pageFuture = useRef<{ pages: PageData[]; trash: PageData[] }[]>([])
  const hasUnsavedChanges = useRef(false)
  const t = useCallback<Translate>((key, values) => translate(language, key, values), [language])
  const setLanguage = (next: AppLanguage) => {
    rawSetLanguage(next)
    setPages(current => current.map(page => ({ ...page, direction: next === 'ar' ? 'rtl' : 'ltr' })))
  }
  const unsavedWarning = t('unsavedWarning')

  const setPages = useCallback((update: PageData[] | ((current: PageData[]) => PageData[])) => {
    const current = pageDataRef.current
    let next = typeof update === 'function' ? update(current) : update
    if (next === current || next.length === current.length && next.every((page,index) => page === current[index])) return
    const layoutOnly = next.length === current.length && next.every((page,index) => page.id === current[index].id && Object.keys({ ...page, ...current[index] }).filter(key => key !== 'height').every(key => page[key as keyof PageData] === current[index][key as keyof PageData]))
    const changed = next.filter((page, index) => page !== current[index])
    const textOnly = next.length === current.length && changed.length === 1 &&
      changed[0].html !== current.find(page => page.id === changed[0].id)?.html
    const now = Date.now()
    if (!layoutOnly && (!textOnly || textGrouping.current.pageId !== changed[0].id || now - textGrouping.current.time > 500)) {
      pageHistory.current = [...pageHistory.current, { pages: current, trash: trashRef.current }].slice(-80)
    }
    if (!layoutOnly) next = next.map((page,index) => page !== current[index] ? { ...page, createdAt: page.createdAt ?? new Date().toISOString(), updatedAt: new Date().toISOString() } : page)
    textGrouping.current = { pageId: textOnly ? changed[0].id : '', time: now }
    if (!layoutOnly) pageFuture.current = []
    pageDataRef.current = next
    changeNumber.current++
    hasUnsavedChanges.current = true
    setStorageStatus('pending')
    rawSetPages(next)
  }, [])

  const updateActivePageHtml = useCallback((html: string) => {
    hasUnsavedChanges.current = true
    const pageId = activePageIdRef.current
    setPages((currentPages) => currentPages.map((page) => (page.id === pageId ? { ...page, html } : page)))
  }, [setPages])

  const editor = useEditor({
    extensions: [
      StarterKit.configure({ undoRedo: false }),
      ImageExtension.configure({ allowBase64: true }),
      StackOrder,
      InkCutMark,
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
    editorProps: { transformPastedHTML: sanitizeImportedHtml },
    onUpdate: ({ editor: currentEditor, transaction }) => {
      if (transaction.getMeta('uiEvent') === 'paste' || transaction.steps.some(step => step.toJSON().stepType !== 'replace')) textGrouping.current = { pageId: '', time: 0 }
      updateActivePageHtml(currentEditor.getHTML())
    },
  })

  useEffect(() => {
    activePageIdRef.current = activePageId
  }, [activePageId])

  useEffect(() => {
    pageDataRef.current = pages
  }, [pages])

  useEffect(() => {
    const channel = new BroadcastChannel('akxr-drafts')
    channel.onmessage = (event) => { if (event.data.revision > revision.current) setRemoteChanged(true) }
    return () => channel.close()
  }, [])

  const projectRef = useRef<SavedProject>({ pages, activePageId, language, trash, version: 1 })
  projectRef.current = { pages, activePageId, language, trash, version: 1 }
  const flushSave = useCallback((forceCheckpoint = false) => {
    if (!ready) return saveQueue.current
    pendingCanvasWrites.forEach(flush => flush())
    const project = { ...projectRef.current, pages: pageDataRef.current }
    const change = changeNumber.current
    saveQueue.current = saveQueue.current.catch(() => undefined).then(async () => {
      try {
        const checkpoint = forceCheckpoint || Date.now() - lastCheckpoint.current > 30_000
        revision.current = await writeDraft(project, revision.current, checkpoint)
        if (checkpoint) lastCheckpoint.current = Date.now()
        if (change === changeNumber.current) {
          hasUnsavedChanges.current = false
          setStorageStatus('saved')
        }
        const channel = new BroadcastChannel('akxr-drafts'); channel.postMessage({ revision: revision.current }); channel.close()
      } catch (error) {
        setStorageStatus(error instanceof StorageConflict ? 'conflict' : 'error')
        hasUnsavedChanges.current = true
      }
    })
    return saveQueue.current
  }, [ready])

  useEffect(() => {
    let cancelled = false
    void readDraft<SavedProject>().then((draft) => {
      if (cancelled) return
      if (draft) {
        const project = normalizeSavedProject(draft.project)
        revision.current = draft.revision
        pageDataRef.current = project.pages
        rawSetPages(project.pages)
        setActivePageId(project.activePageId!)
        rawSetLanguage(project.language ?? 'de')
        setTrash(project.trash ?? [])
      }
      setStorageStatus('pending')
      setReady(true)
    }).catch(() => { if (!cancelled) { setStorageStatus('error'); setReady(true) } })
    return () => { cancelled = true }
  }, [])

  useEffect(() => {
    if (!ready) return
    const timeout = window.setTimeout(() => void flushSave(), 500)
    return () => window.clearTimeout(timeout)
  }, [activePageId, language, pages, trash, flushSave, ready])

  useEffect(() => {
    const saveOnHide = () => { if (document.visibilityState === 'hidden') void flushSave() }
    const confirmBeforeUnload = (event: BeforeUnloadEvent) => {
      void flushSave()
      if (!hasUnsavedChanges.current) return
      event.preventDefault()
      event.returnValue = unsavedWarning
    }
    document.addEventListener('visibilitychange', saveOnHide)
    const saveOnPageHide = () => { void flushSave() }
    window.addEventListener('pagehide', saveOnPageHide)
    window.addEventListener('beforeunload', confirmBeforeUnload)
    return () => {
      document.removeEventListener('visibilitychange', saveOnHide)
      window.removeEventListener('pagehide', saveOnPageHide)
      window.removeEventListener('beforeunload', confirmBeforeUnload)
    }
  }, [flushSave, unsavedWarning])

  const markUnsaved = useCallback(() => {
    hasUnsavedChanges.current = true
  }, [])

  const startNewProject = useCallback(() => {
    if (!window.confirm(t('newProjectWarning'))) return
    void flushSave(true)
    const nextPage: PageData = { id: createId(), html: starterContent, images: [], size: defaultPageSizeKey, direction: language === 'ar' ? 'rtl' : 'ltr', createdAt: new Date().toISOString() }
    importController.current?.abort()
    setImporting(false)
    pageHistory.current = []
    pageFuture.current = []
    hasUnsavedChanges.current = false

    rawSetPages([nextPage])
    pageDataRef.current = [nextPage]
    setTrash([])
    changeNumber.current++
    hasUnsavedChanges.current = true
    setActivePageId(nextPage.id)
    editor?.commands.setContent(nextPage.html, { emitUpdate: false })
    window.setTimeout(() => {

      scrollToPage(nextPage.id)
    }, 0)
  }, [editor, t, flushSave, language])

  const updatePageCanvas = useCallback((pageId: string, canvasJson: string) => {
    markUnsaved()
    setPages((currentPages) => currentPages.map((page) => (page.id === pageId && page.canvasJson !== canvasJson ? { ...page, canvasJson } : page)))
  }, [markUnsaved, setPages])

  const restorePages = useCallback(
    (nextPages: PageData[], preferredPageId?: string) => {
      const restoredPages = nextPages
      const requested = preferredPageId ?? activePageIdRef.current
      const nextActivePageId = restoredPages.some((page) => page.id === requested) ? requested : restoredPages[0]?.id
      if (!nextActivePageId) return false

      rawSetPages(restoredPages)
      pageDataRef.current = restoredPages
      changeNumber.current++
      hasUnsavedChanges.current = true
      setStorageStatus('pending')
      textGrouping.current = { pageId: '', time: 0 }
      setActivePageId(nextActivePageId)
      const nextPage = restoredPages.find((page) => page.id === nextActivePageId)
      if (nextPage) editor?.commands.setContent(nextPage.html, { emitUpdate: false })
      window.setTimeout(() => {

        scrollToPage(nextActivePageId)
      }, 0)
      return true
    },
    [editor],
  )

  const undoPagesChange = useCallback(() => {
    pendingCanvasWrites.forEach(flush => flush())
    const previousPages = pageHistory.current.pop()
    if (!previousPages) return false
    const changedPage = previousPages.pages.find(page => pageDataRef.current.find(current => current.id === page.id) !== page)
    pageFuture.current.push({ pages: pageDataRef.current, trash: trashRef.current })
    setTrash(previousPages.trash)
    return restorePages(previousPages.pages, changedPage?.id)
  }, [restorePages])

  const redoPagesChange = useCallback(() => {
    pendingCanvasWrites.forEach(flush => flush())
    const nextPages = pageFuture.current.pop()
    if (!nextPages) return false
    const changedPage = nextPages.pages.find(page => pageDataRef.current.find(current => current.id === page.id) !== page)
    pageHistory.current.push({ pages: pageDataRef.current, trash: trashRef.current })
    setTrash(nextPages.trash)
    return restorePages(nextPages.pages, changedPage?.id)
  }, [restorePages])

  useEffect(() => {
    const nextPage = pageDataRef.current.find((page) => page.id === activePageId)
    if (!editor || !nextPage) return
    if (editor.getHTML() === nextPage.html) return

    editor.commands.setContent(nextPage.html, { emitUpdate: false })
    window.setTimeout(() => {

    }, 0)
  }, [activePageId, editor])

  const addPage = (size: PageSizeKey = defaultPageSizeKey) => {
    markUnsaved()
    const nextPage: PageData = { id: createId(), html: '<p></p>', images: [], size, direction: language === 'ar' ? 'rtl' : 'ltr', createdAt: new Date().toISOString() }
    setPages((currentPages) => [...currentPages, nextPage])
    setActivePageId(nextPage.id)
    window.setTimeout(() => scrollToPage(nextPage.id), 0)
  }

  const addImagePage = async (file: File) => {
    const change = changeNumber.current
    const nextPage = { id: createId(), html: '<p></p>', images: [], pageImage: await fileToDataUrl(file), size: defaultPageSizeKey }
    if (change !== changeNumber.current) throw new Error(t('importStale'))
    markUnsaved()
    setPages((currentPages) => [...currentPages, nextPage])
    setActivePageId(nextPage.id)
    editor?.commands.setContent(nextPage.html, { emitUpdate: false })
    window.setTimeout(() => scrollToPage(nextPage.id), 0)
  }

  const addPdfPages = async (file: File) => {
    importController.current?.abort()
    const controller = new AbortController()
    importController.current = controller
    const change = changeNumber.current
    setImporting(true)
    try {
      const nextPages = await pagesFromPdf(file, controller.signal)
      if (controller.signal.aborted || change !== changeNumber.current) throw new Error(t('importStale'))
      setPages(current => [...current, ...nextPages])
      setActivePageId(nextPages[0].id)
    } finally { if (controller === importController.current) setImporting(false) }
  }

  const fillPageImageWithBucket = async (pageId: string, x: number, y: number, fillColor: string) => {
    const change = changeNumber.current
    const page = pageDataRef.current.find((item) => item.id === pageId)
    if (!page?.pageImage) return false
    const size = pageSizeFor(page.size)
    const nextPageImage = await floodFillImageDataUrl(page.pageImage, x, y, size.width, size.height, fillColor)
    if (!nextPageImage || change !== changeNumber.current) return true
    if (pageDataRef.current.find(item => item.id === pageId)?.pageImage !== page.pageImage) return true
    if (nextPageImage === page.pageImage) return true
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
    setTrash(current => [...current, pages.find(page => page.id === pageId)!])
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

  const addImageToActivePage = (src: string, width = 260, height = 180, alt = '') => {
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
                  alt,
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
    if (!window.confirm(t('replaceWarning'))) return null
    importController.current?.abort()
    const controller = new AbortController()
    importController.current = controller
    const change = changeNumber.current
    setImporting(true)
    try {
      if (file.size > 50 * 1024 * 1024) throw new Error(t('fileTooLarge'))
      const project = extensionOf(file.name) === 'akxr' ? await readProjectFile(file) : null
      const nextPages = project?.pages ?? await pagesFromFile(file, controller.signal)
      if (controller.signal.aborted || change !== changeNumber.current) throw new Error(t('importStale'))
      await flushSave(true)
      if (controller.signal.aborted || change !== changeNumber.current) throw new Error(t('importStale'))
      const firstPageId = project?.activePageId ?? nextPages[0].id
      setPages(nextPages)
      if (project) { rawSetLanguage(project.language ?? 'de'); setTrash(project.trash ?? []) }
      setActivePageId(firstPageId)
      activePageIdRef.current = firstPageId
      editor?.commands.setContent(nextPages.find(page => page.id === firstPageId)!.html, { emitUpdate: false })
      return nextPages
    } finally { if (importController.current === controller) setImporting(false) }
  }

  const patchPage = useCallback((patch: Partial<PageData>, separateAction = false) => {
    if(separateAction)textGrouping.current={pageId:'',time:0}
    setPages(current => current.map(page => page.id === activePageId && Object.entries(patch).some(([key, value]) => page[key as keyof PageData] !== value) ? { ...page, ...patch } : page))
  }, [activePageId, setPages])
  const applyProject = (project: SavedProject) => {
    setPages(project.pages)
    setTrash(project.trash ?? [])
    rawSetLanguage(project.language ?? language)
    const pageId = project.activePageId ?? project.pages[0].id
    setActivePageId(pageId)
    activePageIdRef.current = pageId
    editor?.commands.setContent(project.pages.find(page => page.id === pageId)!.html, { emitUpdate: false })
  }
  const activePage = pages.find(page => page.id === activePageId) ?? pages[0]
  const searchResults = query.trim() ? pages.filter(page => `${page.title ?? ''} ${page.tags ?? ''} ${textFromHtml(page.html)} ${canvasText(page.canvasJson)}`.toLocaleLowerCase(language).includes(query.toLocaleLowerCase(language))) : []

  return (
    <main className="app-shell" dir={language === 'ar' ? 'rtl' : undefined} lang={language}>
      <Sidebar mode={mode} setMode={setMode} t={t} />
      <section className="workspace">
        <header className="topbar">
          <div className="topbar-title">
            <span>{t('creativeWorkspace')}</span>
            <h1>{mode === 'editor' ? t('editor') : t('convertAFIle')}</h1>
          </div>
          <div className="topbar-actions">
            {offlineReady && <span>{t('offlineReady')}</span>}
            {mode === 'editor' && <HelpHint text={t('recoveryLocal')}><span className={`save-indicator ${storageStatus}`} role="status"><span />{t(storageStatus === 'saved' ? 'savedLocally' : storageStatus === 'conflict' ? 'saveConflict' : storageStatus === 'error' ? 'saveError' : 'saving')}</span></HelpHint>}
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
          </div>
        </header>

        {remoteChanged && <p role="alert">{t('saveConflict')} <button onClick={async () => {
          if (!window.confirm(t('replaceWarning'))) return
          const draft = await readDraft<SavedProject>()
          if (draft) { revision.current = draft.revision; applyProject(normalizeSavedProject(draft.project)); setRemoteChanged(false) }
        }}>{t('loadOtherTab')}</button></p>}
        <div hidden={mode !== 'editor'} className="editor-persistent">
        {ready && <>
          <PageLibrary
            page={activePage} pages={pages} trash={trash} query={query} onQuery={setQuery} results={searchResults}
            onOpen={setActivePageId} onPatch={patchPage} language={language} t={t}
            onDuplicate={() => {
              const copy: PageData = { ...activePage, id:createId(), title:`${activePage.title || t('pageTitle')} (${t('copy')})`, images:activePage.images.map(image => ({...image,id:createId()})) }
              setPages(current=>[...current,copy]);setActivePageId(copy.id)
            }}
            onAttach={async files => {
              const pageId=activePageId
              const attachments:NonNullable<PageData['attachments']>=[]
              for(const file of files) {
                if(file.size>10*1024*1024)throw new Error(t('fileTooLarge'))
                attachments.push({id:createId(),name:file.name,type:file.type,data:await fileToDataUrl(file)})
              }
              setPages(current=>current.map(page=>page.id===pageId?{...page,attachments:[...(page.attachments??[]),...attachments]}:page))
            }}
            onDownload={async id=>{
              const attachment=activePage.attachments?.find(item=>item.id===id)
              if(attachment)saveBlob(await (await fetch(attachment.data)).blob(),attachment.name)
            }}
            onRestorePage={id=>{
              const page=trash.find(page=>page.id===id)
              if(!page)return
              const restored={...page,id:createId()}
              setPages(current=>[...current,restored]);setTrash(current=>current.filter(page=>page.id!==id))
            }}
            onLoadVersions={async ()=>{await flushSave(true);return localVersions<SavedProject>()}}
            onRestoreVersion={async version=>{
              const project=normalizeSavedProject(version.project)
              pendingCanvasWrites.forEach(flush=>flush())
              const change=changeNumber.current
              await flushSave(true)
              if(changeNumber.current!==change || hasUnsavedChanges.current)throw new Error('restoreFailed')
              applyProject(project)
            }}
          />
          {importing && <div role="status">{t('importing')}<button onClick={() => { importController.current?.abort(); setImporting(false) }}>{t('cancel')}</button></div>}
          <EditorMode
            activePageId={activePageId}
            patchPage={patchPage}
            trash={trash}
            onProjectSaved={() => { hasUnsavedChanges.current = false }}
            addImageToActivePage={addImageToActivePage}
            addImagePage={addImagePage}
            addPage={addPage}
            addPdfPages={addPdfPages}
            updatePageCanvas={updatePageCanvas}
            deletePage={deletePage}
            editor={editor}
            pages={pages}
            pagesRef={pagesRef}
            onNewProject={startNewProject}
            openDocument={openDocument}
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
        </>}
        </div>
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
      <div className="brand-lockup">
        <div className="brand-mark"><img alt="AKXR Editor" src="/akxr-logo.png" /></div>
        <div className="brand-copy">
          <strong>AKXR</strong>
          <span>Editor</span>
        </div>
      </div>
      <p className="sidebar-caption">{t('workspaceCaption')}</p>
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
  patchPage,
  trash,
  onProjectSaved,
  activePageId,
  addImageToActivePage,
  addImagePage,
  addPage,
  addPdfPages,
  updatePageCanvas,
  deletePage,
  editor,
  pages,
  pagesRef,
  onNewProject,
  openDocument,
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
  patchPage: (patch: Partial<PageData>, separateAction?:boolean) => void
  trash: PageData[]
  onProjectSaved: () => void
  activePageId: string
  addImageToActivePage: (src: string, width?: number, height?: number, alt?: string) => void
  addImagePage: (file: File) => Promise<void>
  addPage: (size?: PageSizeKey) => void
  addPdfPages: (file: File) => Promise<void>
  updatePageCanvas: (pageId: string, canvasJson: string) => void
  deletePage: (pageId: string) => void
  editor: Editor | null
  pages: PageData[]
  pagesRef: React.RefObject<HTMLDivElement | null>
  onNewProject: () => void
  openDocument: (file: File) => Promise<PageData[] | null>
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
  const loadingCanvas = useRef(false)
  const fabricClipboard = useRef<FabricObject | null>(null)
  const [tool, setTool] = useState<EditorTool>('write')
  const [zoom,setZoom]=useState(1)
  const [eraserSize,setEraserSize]=useState(24)
  const currentPages=useRef(pages)
  currentPages.current=pages
  const currentActiveId=useRef(activePageId)
  currentActiveId.current=activePageId
  const layerOrder = pages.find(page => page.id === activePageId)?.layerOrder ?? ['text', 'images', 'paint'] as LayerKind[]
  const setLayerOrder = (update: (current: LayerKind[]) => LayerKind[]) => {
    const next = update(layerOrder)
    if (next.some((layer,index) => layer !== layerOrder[index])) patchPage({ layerOrder: next })
  }
  const [textColor, setTextColor] = useState('#171224')
  const [textFontFamily, setTextFontFamily] = useState(fontFamilies[0])
  const [textFontSize, setTextFontSize] = useState(defaultFontSize)
  const textLineGap = pages.find(page => page.id === activePageId)?.lineHeight ?? 1.72
  const setTextLineGap = (lineHeight: number) => patchPage({ lineHeight })
  const [textTransparency, setTextTransparency] = useState(0)
  const [exportFileName, setExportFileName] = useState('akxreditor-document')
  const [textBold, setTextBold] = useState(false)
  const [textItalic, setTextItalic] = useState(false)
  const [textUnderline, setTextUnderline] = useState(false)
  const [stroke, setStroke] = useState('#6d5dfc')
  const [fill, setFill] = useState('#d93f75')
  const [strokeTransparency, setStrokeTransparency] = useState(0)
  const [fillTransparency, setFillTransparency] = useState(0)
  const [penOnly, setPenOnly] = useState(false)
  const [pressure, setPressure] = useState(true)
  const [brushSize, setBrushSize] = useState(8)
  const [selectedImageId, setSelectedImageId] = useState<string | null>(null)
  const [selectedTarget, setSelectedTarget] = useState<SelectedTarget | null>(null)
  const [pendingImageExport, setPendingImageExport] = useState<PendingImageExport | null>(null)
  const [contextMenu, setContextMenu] = useState<ContextMenuState | null>(null)
  const [hasFabricClipboard, setHasFabricClipboard] = useState(false)
  const [draggedPageId, setDraggedPageId] = useState<string | null>(null)
  const [dragOverPageId, setDragOverPageId] = useState<string | null>(null)
  const [statusMessage, setStatusMessage] = useState<TranslationKey | ''>('')
  const [, updateToolbar] = useState(0)
  const shortcutHandlerRef = useRef<(event: KeyboardEvent) => void>(() => undefined)
  const layerZ = {
    text: layerOrder.indexOf('text') + 1,
    images: layerOrder.indexOf('images') + 1,
    paint: layerOrder.indexOf('paint') + 1,
  }
  const canSelectPaint = tool === 'select-paint'
  const canSelectText = tool === 'select-text'
  const effectiveTextColor = colorWithTransparency(textColor, textTransparency)
  const effectiveStroke = colorWithTransparency(stroke, strokeTransparency)
  const effectiveFill = colorWithTransparency(fill, fillTransparency)

  const moveLayer = (layer: LayerKind, direction: 'front' | 'back') => {
    setLayerOrder((currentOrder) => {
      const remaining = currentOrder.filter((item) => item !== layer)
      return direction === 'front' ? [...remaining, layer] : [layer, ...remaining]
    })
  }

  const moveLayerToFront = (layer: LayerKind) => moveLayer(layer, 'front')

  const activateTool = (nextTool: EditorTool) => {
    setTool(nextTool)
    if(nextTool==='eraser' || nextTool==='area-eraser') { activeCanvas()?.discardActiveObject();activeCanvas()?.requestRenderAll();setSelectedTarget(null);setSelectedImageId(null) }
    if (nextTool === 'brush' || nextTool === 'eraser' || nextTool === 'bucket' || nextTool === 'place-text') moveLayerToFront('paint')
  }

  const toggleTextPlacement = () => {
    activateTool(tool === 'place-text' ? 'select-paint' : 'place-text')
  }

  const createPaintText = (text: string, left: number, top: number) =>
    new IText(text, {
      fill: effectiveTextColor,
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
  } as React.CSSProperties

  const activatePage = (pageId: string) => {
    setSelectedImageId(null)
    setSelectedTarget(null)
    setContextMenu(null)
    setActivePageId(pageId)
    window.setTimeout(() => scrollToPage(pageId), 0)
  }

  // All document mutations share App's bounded snapshot history.
  const rememberTextAction = () => {}
  const rememberPagesAction = () => {}
  const rememberPageImageBucketFill = () => {}
  const runTextCommand = (command: () => void) => {
    if (editor) syncEditorSelection(editor)
    command()
  }
  const handleSelectDrawing = useCallback((pageId: string, tableCellIndex?: number) => {
    setSelectedImageId(null)
    setSelectedTarget(tableCellIndex == null ? { kind: 'paint', pageId } : { kind: 'table-cell', pageId, cellIndex: tableCellIndex })
    const selected = canvasMap.current.get(pageId)?.getActiveObject()
    const text = isHorizontalTableGroup(selected) ? horizontalTableTextItems(selected)[tableCellIndex ?? 0] : selected
    if (text instanceof IText) {
      setTextBold(text.fontWeight === 'bold' || Number(text.fontWeight) >= 600)
      setTextItalic(text.fontStyle === 'italic')
      setTextUnderline(text.underline)
      setTextFontFamily(text.fontFamily)
      setTextFontSize(`${text.fontSize}px`)
    }
  }, [])
  const registerCanvas = useCallback((pageId: string, canvas: FabricCanvas | null) => {
    if (canvas) canvasMap.current.set(pageId, canvas)
    else canvasMap.current.delete(pageId)
  }, [])
  const recordCanvasHistory = useCallback((pageId: string, canvas: FabricCanvas) => {
    if (!loadingCanvas.current) updatePageCanvas(pageId, canvasSnapshot(canvas))
  }, [updatePageCanvas])

  const activeCanvas = useCallback(() => canvasMap.current.get(activePageId), [activePageId])

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
    applyCanvasSelectionStyle(object)
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
      textAlign: string
      underline: boolean
    }>,
  ) => {
    const canvas = activeCanvas()
    const object = canvas?.getActiveObject()
    if (!canvas || (!isHorizontalTableGroup(object) && !(object instanceof IText))) return
    if (isHorizontalTableGroup(object)) {
      object.getObjects().forEach((item, index) => {
        if (item instanceof IText && (selectedTarget?.kind !== 'table-cell' || Math.floor(index / 2) === selectedTarget.cellIndex)) {
          item.set(patch)
          sharpenCanvasText(item)
          item.initDimensions()
        }
      })
    } else {
      if (object.isEditing && object.selectionStart !== object.selectionEnd) object.setSelectionStyles(patch)
      else object.set(patch)
      sharpenCanvasText(object)
      object.initDimensions()
    }
    object.setCoords()
    canvas.requestRenderAll()
    recordCanvasHistory(activePageId, canvas)
  }

  const updateTextColor = (nextColor: string) => {
    setTextColor(nextColor)
    const nextEffectiveColor = colorWithTransparency(nextColor, textTransparency)
    updateSelectedPaintTextStyle({ fill: nextEffectiveColor })
    if (!selectedPaintText() && editor) runTextCommand(() => editor.chain().focus().setColor(nextEffectiveColor).run())
  }

  const updateTextTransparency = (nextTransparency: number) => {
    const normalizedTransparency = clamp(Math.round(nextTransparency), 0, 100)
    setTextTransparency(normalizedTransparency)
    const nextEffectiveColor = colorWithTransparency(textColor, normalizedTransparency)
    updateSelectedPaintTextStyle({ fill: nextEffectiveColor })
    if (!selectedPaintText() && editor) runTextCommand(() => editor.chain().focus().setColor(nextEffectiveColor).run())
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
    const selected = canvas.getActiveObject()
    if (selected instanceof IText && mode !== 'plain') {
      selected.off('changed'); selected.off('editing:exited')
      selected.set({ text: formatManagedText(selected.text, mode) })
      configureManagedText(selected, mode, canvas, { removeIfEmpty: true })
      canvas.requestRenderAll(); recordCanvasHistory(activePageId, canvas); return
    }
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
    if (selectedPaintText()) return
    editor?.chain().focus().insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run()
    setTool('write')
  }
  const alignText = (alignment: string) => {
    if (selectedPaintText()) updateSelectedPaintTextStyle({ textAlign: alignment })
    else editor?.chain().focus().setTextAlign(alignment).run()
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
    const targetPageId = selectedTarget?.kind === 'paint' || selectedTarget?.kind === 'table-cell' ? selectedTarget.pageId : activePageId
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
    const object = horizontalTableFromTarget(canvas?.getActiveObject())
    if (!canvas || !isHorizontalTableGroup(object)) return
    deleteHorizontalTableCell(object, selectedTarget.cellIndex, canvas)
    canvas.requestRenderAll()
    recordCanvasHistory(selectedTarget.pageId, canvas)
    setSelectedTarget(null)
  }

  const deleteSelectedObject = () => {
    setContextMenu(null)
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
    const targetPageId = selectedTarget?.kind === 'paint' || selectedTarget?.kind === 'table-cell' ? selectedTarget.pageId : activePageId
    const canvas = canvasMap.current.get(targetPageId)
    const object = selectedTarget?.kind === 'table-cell' ? horizontalTableFromTarget(canvas?.getActiveObject()) : canvas?.getActiveObject()
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
    const image = pages.find(page => page.id === selectedTarget?.pageId)?.images.find(image => selectedTarget?.kind === 'image' && image.id === selectedTarget.imageId)
    const operation = image ? fetch(image.src).then(response => response.blob()).then(async blob => {
      const png = blob.type === 'image/png' ? blob : await convertImageBlob(image.src, 'png')
      await navigator.clipboard.write([new ClipboardItem({ 'image/png': png })])
    }) : selectedTarget?.kind === 'text' && editor ? navigator.clipboard.writeText(editor.state.doc.textBetween(editor.state.selection.from, editor.state.selection.to, '\n')) : copySelectedCanvasObject()
    void Promise.resolve(operation).catch(error => window.alert(String(error)))
    setContextMenu(null)
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

  const saveProject = () => {
    const snapshotPages = pages.map((page) => ({
      ...page,
      canvasJson: canvasMap.current.get(page.id) ? canvasSnapshot(canvasMap.current.get(page.id)!) : page.canvasJson,
    }))
    const project: SavedProject = { activePageId, language, pages: snapshotPages, trash, version: 1 }
    saveBlob(new Blob([JSON.stringify(project)], { type: 'application/json;charset=utf-8' }), `${safeFileBaseName(exportFileName)}.akxr`)
    onProjectSaved()
    setStatusMessage('projectSaved')
    window.setTimeout(() => setStatusMessage(''), 3_200)
  }

  const continueImageExport = () => {
    if (!pendingImageExport) return
    const { baseName, format } = pendingImageExport
    setPendingImageExport(null)
    void savePagesAsImageFiles(format, baseName).catch((error) => {
      setStatusMessage('couldNotExport')
      window.alert(localizedErrorMessage(error, t, 'couldNotExport'))
      return false
    }).then((success) => {
      if (success === false) return
      setStatusMessage('exportReady')
      window.setTimeout(() => setStatusMessage(''), 3_200)
    })
  }

  const exportWithStatus = async (format: ExportFormat, fileName: string) => {
    setStatusMessage('exporting')
    try { await exportDocument(format, fileName) } catch (error) { setStatusMessage('couldNotExport'); window.alert(localizedErrorMessage(error, t, 'couldNotExport')); return }
    if ((format !== 'png' && format !== 'jpg') || pages.length <= 1) {
      setStatusMessage('exportReady')
      window.setTimeout(() => setStatusMessage(''), 3_200)
    }
  }

  const exportDocument = async (format: ExportFormat, requestedBaseName = exportFileName) => {
    if (!pagesRef.current) return
    const baseName = safeFileBaseName(requestedBaseName)
    const snapshotPages = pages.map(page => {
      const canvas = canvasMap.current.get(page.id)
      return canvas ? { ...page, canvasJson: canvasSnapshot(canvas) } : page
    })
    const html = snapshotPages.map((page) => page.html + canvasTextHtml(page.canvasJson)).join('<hr>')
    const text = snapshotPages.map((page, index) => `${t('page', { number: index + 1 })}\n${textFromHtml(page.html)}\n${canvasText(page.canvasJson)}`).join('\n\n')

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
        new Blob([wrapHtmlDocument(await renderPagesHtml(pages, (pageId) => canvasMap.current.get(pageId)?.toDataURL()), baseName, language)], {
          type: 'text/html;charset=utf-8',
        }),
        exportFileNameFor(baseName, format),
      )
      return
    }
    if (format === 'docx') {
      saveBlob(
        await createDocxBlobFromPages(snapshotPages),
        exportFileNameFor(baseName, format),
      )
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
    addImageToActivePage(src, maxWidth, maxWidth * ratio, file.name)
    moveLayerToFront('images')
    setTool('select-paint')
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

  const undoLast = () => { undoPagesChange() }
  const redoLast = () => { redoPagesChange() }

  shortcutHandlerRef.current = (event: KeyboardEvent) => {
      if (editor?.view.dom.closest('[hidden]')) return
      const target = event.target
      const typingTarget = isTypingTarget(target)
      const isUndo = (event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'z' && !event.shiftKey
      const isRedo = (event.ctrlKey || event.metaKey) && (event.key.toLowerCase() === 'y' || (event.shiftKey && event.key.toLowerCase() === 'z'))
      const isCopy = (event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'c'
      const isPaste = (event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'v'

      if (event.defaultPrevented) return
      if ((isUndo || isRedo) && target instanceof HTMLElement && target.closest('input,select,textarea') && !target.closest('.paper') && !target.hasAttribute('data-fabric')) return
      if (isUndo) {
        event.preventDefault()
        event.stopImmediatePropagation()
        undoLast()
        return
      }
      if (isRedo) {
        event.preventDefault()
        event.stopImmediatePropagation()
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
        if (tool === 'place-text' || tool === 'brush' || tool === 'eraser' || tool === 'area-eraser' || tool === 'bucket') setTool('select-paint')
      }
  }

  useEffect(() => {
    const handleShortcut = (event: KeyboardEvent) => shortcutHandlerRef.current(event)
    window.addEventListener('keydown', handleShortcut, true)
    return () => window.removeEventListener('keydown', handleShortcut, true)
  }, [])

  useEffect(() => {
    if (!editor) return
    const sync = () => {
      updateToolbar(value => value + 1)
      if (canvasMap.current.get(activePageId)?.getActiveObject()) return
      const style = editor.getAttributes('textStyle')
      setTextBold(editor.isActive('bold'))
      setTextItalic(editor.isActive('italic'))
      setTextUnderline(editor.isActive('underline'))
      setTextFontFamily(style.fontFamily || pages.find(page => page.id === activePageId)?.fontFamily || fontFamilies[0])
      setTextFontSize(style.fontSize || defaultFontSize)
    }
    editor.on('selectionUpdate', sync)
    editor.on('transaction', sync)
    sync()
    return () => { editor.off('selectionUpdate', sync); editor.off('transaction', sync) }
  }, [editor, activePageId, pages])

  useEffect(() => {
    const element = editor?.view.dom
    if (!element) return
    const measure = () => {
      const page = pages.find(page => page.id === activePageId)
      if (!page) return
      const height = Math.max(page.baseHeight ?? pageSizeFor(page.size).height, Math.ceil(element.scrollHeight))
      if (Math.abs(height - dimensionsFor(page).height) > 2) patchPage({ height })
    }
    const observer = new ResizeObserver(measure)
    observer.observe(element)
    return () => observer.disconnect()
  }, [editor, activePageId, pages, patchPage])

  const eraseArea = async (stroke:EraseStroke) => {
    if(!editor)return
    const page=currentPages.current.find(page=>page.id===activePageId)
    const canvas=canvasMap.current.get(activePageId)
    const paper=document.getElementById(`page-${activePageId}`)
    if(!page || !canvas || !paper)return
    const documentBefore=editor.state.doc
    const drawingBefore=canvasSnapshot(canvas)
    setStatusMessage('erasing')
    try {
      const bounds=paper.getBoundingClientRect()
      const size=dimensionsFor(page)
      const transaction=eraseTextTransaction(editor,stroke,bounds,bounds.width/size.width)
      const inlineImages=await eraseInlineImages(editor,stroke,bounds,bounds.width/size.width)
      for(const image of inlineImages)transaction.setNodeMarkup(image.pos,undefined,{...image.attrs,src:image.src})
      const applyDrawing=prepareDrawingErase(canvas,stroke)
      const pageImage=page.pageImage ? await eraseImage(page.pageImage,{left:0,top:0,width:size.width,height:size.height},0,stroke) : undefined
      const images:FloatingImage[]=[]
      for(const image of page.images) images.push({...image,src:await eraseImage(image.src,{left:image.x,top:image.y,width:image.width,height:image.height},image.rotation,stroke)})
      const latest=currentPages.current.find(item=>item.id===page.id)
      if(currentActiveId.current!==page.id || !latest || latest.pageImage!==page.pageImage || latest.images!==page.images || editor.state.doc!==documentBefore || canvasSnapshot(canvas)!==drawingBefore) {
        setStatusMessage('eraseStale');return
      }
      applyDrawing()
      editor.view.dispatch(transaction.setMeta('preventUpdate',true))
      patchPage({html:editor.getHTML(),canvasJson:canvasSnapshot(canvas),pageImage,images},true)
      setStatusMessage('erased')
    } catch(error) {setStatusMessage(error instanceof Error && error.message==='renderLimit' ? 'renderLimit' : 'eraseFailed')}
  }

  const textFontSizeNumber = Number.parseFloat(textFontSize)
  const defaultFontSizeNumber = Number.parseFloat(defaultFontSize)
  const visibleTextFontSize = Number.isFinite(textFontSizeNumber) ? Math.round(textFontSizeNumber) : defaultFontSizeNumber
  const selectedFontSizePreset = fontSizes.includes(`${visibleTextFontSize}px`) ? `${visibleTextFontSize}px` : ''

  return (
    <section className="mode-panel" style={textStyleVars}>
      <div className="editor-command-strip">
        <ToolButton icon={Upload} label={t('openFile')} onClick={() => openInputRef.current?.click()} />
        <ToolButton icon={Plus} label={t('newProject')} onClick={onNewProject} />
        <ToolButton hint={t('saveProjectHint')} icon={Save} label={t('saveProject')} onClick={saveProject} />
        <input
          accept=".akxr,.docx,.pdf,.txt,.md,.markdown,.html,.htm,.png,.jpg,.jpeg"
          className="hidden-input"
          onChange={(event) => {
            const file = event.target.files?.[0]
            if (file) {
              void openDocument(file)
                .then((nextPages) => {
                  if (!nextPages) return
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
        <div className="page-zoom" aria-label={t('zoomHint')}>
          <HelpHint text={t('zoomHint')}><button type="button" aria-label={t('zoomOut')} disabled={zoom<=0.35} onClick={()=>setZoom(value=>Math.max(0.35,Math.round((value-0.1)*100)/100))}><ZoomOut size={16}/></button></HelpHint>
          <HelpHint text={t('zoomHint')}><button type="button" aria-label={t('zoomReset')} onClick={()=>setZoom(1)}>{new Intl.NumberFormat(language,{style:'percent'}).format(zoom)}</button></HelpHint>
          <HelpHint text={t('zoomHint')}><button type="button" aria-label={t('zoomIn')} disabled={zoom>=2} onClick={()=>setZoom(value=>Math.min(2,Math.round((value+0.1)*100)/100))}><ZoomIn size={16}/></button></HelpHint>
        </div>
        {statusMessage && <span className="editor-status" role="status">{t(statusMessage)}</span>}
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
              <input aria-label={t('textColor')} type="color" onChange={(event) => updateTextColor(event.target.value)} value={textColor} />
            </label>
            <label className="range-control transparency-control" title={t('transparency')}>
              <span>{t('transparency')}</span>
              <input aria-label={`${t('textColor')} ${t('transparency')}`} min="0" max="100" onChange={(event) => updateTextTransparency(Number(event.target.value))} type="range" value={textTransparency} />
              <output>{textTransparency}%</output>
            </label>
            <label className="range-control" title={t('textGap')}>
              <Rows3 size={16} />
              <input aria-label={t('textGap')} min="1" max="2.6" onChange={(event) => updateTextLineGap(Number(event.target.value))} step="0.02" type="range" value={textLineGap} />
              <span>{textLineGap.toFixed(2)}</span>
            </label>
          </div>
          <div className="rail-section compact-icons" aria-label={t('formatting')}>
            <IconButton active={textBold} icon={Bold} label={t('bold')} onClick={toggleTextBold} />
            <IconButton active={textItalic} icon={Italic} label={t('italic')} onClick={toggleTextItalic} />
            <IconButton active={textUnderline} icon={UnderlineIcon} label={t('underline')} onClick={toggleTextUnderline} />
            <select
              aria-label={t('heading')}
              value={[1,2,3].find(level => editor?.isActive('heading', { level })) ?? 0}
              onChange={(event) => {
                const level = Number(event.target.value)
                if (selectedPaintText()) { updateSelectedPaintTextStyle({ fontSize: level ? [32, 24, 18][level - 1] : 12 }); return }
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
            <IconButton icon={AlignLeft} label={t('alignLeft')} onClick={() => alignText('left')} />
            <IconButton icon={AlignCenter} label={t('alignCenter')} onClick={() => alignText('center')} />
            <IconButton icon={AlignRight} label={t('alignRight')} onClick={() => alignText('right')} />
            <IconButton icon={List} label={t('bulletList')} onClick={() => selectedPaintText() ? addMovableText('- ', 'bullet') : editor?.chain().focus().toggleBulletList().run()} />
            <IconButton icon={ListOrdered} label={t('numberedList')} onClick={() => selectedPaintText() ? addMovableText('1. ', 'numbered') : editor?.chain().focus().toggleOrderedList().run()} />
            <IconButton icon={Table2} label={t('insertTable')} onClick={addHorizontalTable} />
            {editor?.isActive('table') && <>
              <button onClick={() => editor.chain().focus().addRowAfter().run()}>{t('addRow')}</button>
              <button onClick={() => editor.chain().focus().addColumnAfter().run()}>{t('addColumn')}</button>
              <button onClick={() => editor.chain().focus().deleteRow().run()}>{t('deleteRow')}</button>
              <button onClick={() => editor.chain().focus().deleteColumn().run()}>{t('deleteColumn')}</button>
              <button onClick={() => editor.chain().focus().deleteTable().run()}>{t('deleteTable')}</button>
            </>}
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
              <span>{page.section ? `${page.section} / ` : ''}{page.title || t('page', { number: index + 1 })}</span>
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
        {pages.filter(page => page.id === activePageId).map((page) => {
          const pageSize = dimensionsFor(page)
          return (
          <div className="paper-viewport" key="active-page" style={{width:pageSize.width*zoom,height:pageSize.height*zoom}}>
          <article
            dir={page.direction}
            className={page.id === activePageId ? 'paper active-paper' : 'paper'}
            id={`page-${page.id}`}
            key="active-page"
            style={{...pageStyle(page),transform:`scale(${zoom})`,transformOrigin:'top left'}}
            onClick={() => {
              activatePage(page.id)
              setSelectedImageId(null)
              setSelectedTarget(null)
              clearCanvasSelection(page.id)
            }}
          >
            {page.pageImage && <img alt={page.title ?? t('pageTitle')} className="page-image" draggable={false} src={page.pageImage} />}
            {page.id === activePageId && <PageCanvas key={page.id}
              active={page.id === activePageId}
              brushSize={brushSize}
              pressure={pressure}
              label={t('drawing')}
              penOnly={penOnly}
              canSelectPaint={canSelectPaint}
              canvasJson={page.canvasJson}
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
              textColor={effectiveTextColor}
              textBold={textBold}
              textFontFamily={textFontFamily}
              textFontSize={Number.parseFloat(textFontSize)}
              textItalic={textItalic}
              textLineGap={textLineGap}
              textUnderline={textUnderline}
              fill={effectiveFill}
              stroke={effectiveStroke}
              zIndex={layerZ.paint}
              tool={tool}
            />}
            {page.id === activePageId ? (
              <div
                className={textLayerClass(tool, canSelectText)}
                style={{ zIndex: layerZ.text }}
                onContextMenu={(event) => {
                  if (editor?.state.selection.empty) return
                  setSelectedImageId(null)
                  clearCanvasSelection(page.id)
                  openContextMenu(event, { kind: 'text', pageId: page.id })
                }}
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
            <div className={tool === 'select-paint' ? 'floating-layer interactive' : 'floating-layer'} style={{ zIndex: tool === 'select-paint' ? 10 : layerZ.images }}>
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
            {tool==='area-eraser' && <AreaEraser width={pageSize.width} height={pageSize.height} diameter={eraserSize} penOnly={penOnly} label={t('pixelEraser')} onErase={eraseArea}/> }
          </article>
          </div>
          )
        })}
      </div>
        </div>

        <aside className="editor-rail editor-rail-right" aria-label={t('drawing')}>
          <div className="rail-section">
<HelpHint text={t('penOnlyHint')}><label><input type="checkbox" checked={penOnly} onChange={event => setPenOnly(event.target.checked)} />{t('penOnly')}</label></HelpHint>
<HelpHint text={t('pressureHint')}><label><input type="checkbox" checked={pressure} onChange={event => setPressure(event.target.checked)} />{t('pressure')}</label></HelpHint>
            <ToolButton active={tool === 'select-paint'} icon={MousePointer2} label={t('selectPaint')} onClick={() => activateTool('select-paint')} />
            <IconButton active={tool === 'brush'} icon={Brush} label={t('brush')} onClick={() => activateTool('brush')} />
            <IconButton hint={t('objectEraserHint')} active={tool === 'eraser'} icon={Eraser} label={t('eraser')} onClick={() => activateTool('eraser')} />
            <ToolButton hint={t('pixelEraserHint')} active={tool==='area-eraser'} icon={Eraser} label={t('pixelEraser')} onClick={()=>activateTool('area-eraser')}/>
            {tool==='area-eraser' && <HelpHint text={t('eraserSizeHint')}><label className="range-control eraser-size"><span>{t('eraserSize')}</span><input aria-label={t('eraserSize')} min="4" max="120" type="range" value={eraserSize} onChange={event=>setEraserSize(Number(event.target.value))}/><output>{eraserSize}px</output></label></HelpHint>}
            <IconButton active={tool === 'bucket'} icon={PaintBucket} label={t('bucket')} onClick={() => activateTool('bucket')} />
            <label className="range-control">
              <PenLine size={16} />
              <input aria-label={t('brushWidth')} min="1" max="42" onChange={(event) => setBrushSize(Number(event.target.value))} type="range" value={brushSize} />
              <span>{brushSize}</span>
            </label>
            <label className="color-chip" title={t('strokeColor')}>
              <Highlighter size={16} />
              <input onChange={(event) => setStroke(event.target.value)} type="color" value={stroke} />
            </label>
            <label className="range-control transparency-control" title={`${t('strokeColor')} ${t('transparency')}`}>
              <span>{t('transparency')}</span>
              <input aria-label={`${t('strokeColor')} ${t('transparency')}`} min="0" max="100" onChange={(event) => setStrokeTransparency(clamp(Math.round(Number(event.target.value)), 0, 100))} type="range" value={strokeTransparency} />
              <output>{strokeTransparency}%</output>
            </label>
            <label className="color-chip" title={t('fillColor')}>
              <Shapes size={16} />
              <input onChange={(event) => setFill(event.target.value)} type="color" value={fill} />
            </label>
            <label className="range-control transparency-control" title={`${t('fillColor')} ${t('transparency')}`}>
              <span>{t('transparency')}</span>
              <input aria-label={`${t('fillColor')} ${t('transparency')}`} min="0" max="100" onChange={(event) => setFillTransparency(clamp(Math.round(Number(event.target.value)), 0, 100))} type="range" value={fillTransparency} />
              <output>{fillTransparency}%</output>
            </label>
          </div>
          <div className="rail-section compact-icons">
            <IconButton icon={RectangleHorizontal} label={t('rectangle')} onClick={() => addFabricObject(new Rect({ left: 160, top: 160, width: 190, height: 110, fill: effectiveFill, stroke: effectiveStroke, strokeWidth: 3, rx: 8, ry: 8 }))} />
            <IconButton icon={CircleIcon} label={t('circle')} onClick={() => addFabricObject(new Circle({ left: 220, top: 150, radius: 64, fill: effectiveFill, stroke: effectiveStroke, strokeWidth: 3 }))} />
            <IconButton
              icon={X}
              label={t('xShape')}
              onClick={() => {
                const size = 140
                const width = Math.max(3, brushSize)
                const firstStroke = new Line([0, 0, size, size], { stroke: effectiveStroke, strokeWidth: width, strokeLineCap: 'round' })
                const secondStroke = new Line([size, 0, 0, size], { stroke: effectiveStroke, strokeWidth: width, strokeLineCap: 'round' })
                addFabricObject(new Group([firstStroke, secondStroke], { left: 220, top: 160 }))
              }}
            />
            <IconButton icon={Rows3} label={t('line')} onClick={() => addFabricObject(new Line([150, 230, 470, 300], { stroke: effectiveStroke, strokeWidth: brushSize, strokeLineCap: 'round' }))} />
            <IconButton
              icon={ArrowRight}
              label={t('arrow')}
              onClick={() => {
                const line = new Line([0, 0, 240, 80], { stroke: effectiveStroke, strokeWidth: brushSize, strokeLineCap: 'round' })
                const head = new Triangle({ left: 240, top: 80, width: 24, height: 28, fill: effectiveStroke, angle: 110, originX: 'center', originY: 'center' })
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
            <ExportMenu fileName={exportFileName} formats={exportFormats} onExport={exportWithStatus} onFileNameChange={setExportFileName} t={t} />
          </div>
        </aside>
      </div>
      {pendingImageExport && (
        <ImageExportConfirmation
          format={pendingImageExport.format}
          pageCount={pendingImageExport.pageCount}
          onCancel={() => {
            setPendingImageExport(null)
            setStatusMessage('')
          }}
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
  label,
  pressure,
  penOnly,
  active,
  brushSize,
  canSelectPaint,
  canvasJson,
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
  label:string
  pressure: boolean
  penOnly: boolean
  active: boolean
  brushSize: number
  canSelectPaint: boolean
  canvasJson?: string
  pageHeight: number
  pageId: string
  pageWidth: number
  fillPageImageWithBucket: (pageId: string, x: number, y: number, fillColor: string) => Promise<boolean>
  onPageImageBucketFill: () => void
  recordCanvasHistory: (pageId: string, canvas: FabricCanvas) => void
  registerCanvas: (pageId: string, canvas: FabricCanvas | null, initialCanvasJson?: string) => void
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
  const initialSnapshot = useRef(canvasJson)
  const appliedSnapshot = useRef<string | undefined>(undefined)
  const restoring = useRef(false)
  const restoreSerial = useRef(Promise.resolve())
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
      enableRetinaScaling: false,
      enablePointerEvents: true,
      height: pageHeight,
      fireRightClick: true,
      preserveObjectStacking: true,
      perPixelTargetFind:true,
      targetFindTolerance:0,
      selection: true,
      selectionBorderColor: '#7c5cff',
      selectionColor: 'rgba(124, 92, 255, 0.08)',
      selectionDashArray: [5, 4],
      selectionLineWidth: 1.5,
      stopContextMenu: true,
      width: pageWidth,
    })
    fabricRef.current = canvas
    if (!initialSnapshot.current) appliedSnapshot.current = canvasSnapshot(canvas)
    let recordTimer: number | undefined
    const recordNow = () => {
      if (restoring.current) return
      const snapshot = canvasSnapshot(canvas)
      if (snapshot === appliedSnapshot.current) return
      appliedSnapshot.current = snapshot
      callbacksRef.current.recordCanvasHistory(pageId, canvas)
    }
    const record = () => { window.clearTimeout(recordTimer); recordTimer = window.setTimeout(recordNow, 100) }
    pendingCanvasWrites.add(recordNow)
    const styleAndRecord = (event: { target?: FabricObject }) => {
      if (event.target) { applyCanvasSelectionStyle(event.target); event.target.selectable=canSelectPaintRef.current || toolRef.current==='place-text';event.target.evented=['select-paint','place-text','bucket','eraser'].includes(toolRef.current) }
      record()
    }
    const constrain = (event: { target?: FabricObject }) => {
      if (event.target) keepObjectInsideCanvas(event.target, canvas)
    }
    canvas.on('object:added', styleAndRecord)
    canvas.on('object:modified', record)
    canvas.on('object:removed', record)
    canvas.on('text:changed', record)
    canvas.on('text:editing:exited', record)
    canvas.on('object:moving', constrain)
    canvas.on('object:scaling', constrain)
    canvas.on('object:rotating', constrain)
    const tableCellIndexFromEvent = (event: { e?: Event; target?: FabricObject; subTargets?: FabricObject[] }) => {
      if (!event.e) return null
      const target = event.target ?? canvas.getActiveObject()
      const table = horizontalTableFromTarget(target) ?? horizontalTableFromTarget(canvas.getActiveObject())
      return table ? horizontalTableCellIndexFromPointer(table, canvas.getScenePoint(event.e as MouseEvent | TouchEvent), event.subTargets ?? (target ? [target] : [])) : null
    }
    const selectDrawing = (event: { e?: Event; target?: FabricObject; subTargets?: FabricObject[] }) => {
      callbacksRef.current.onSelectDrawing(pageId, tableCellIndexFromEvent(event) ?? undefined)
    }
    const openDrawingMenu = (event: { e: Event; target?: FabricObject; subTargets?: FabricObject[] }) => {
      const canOpenPaintMenu = canSelectPaintRef.current || toolRef.current === 'place-text'
      if (!activeRef.current || !canOpenPaintMenu || !(event.e instanceof MouseEvent)) return
      const target = event.target ?? canvas.getActiveObject()
      const pointer = canvas.getScenePoint(event.e as MouseEvent | TouchEvent)
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
      if (!activeRef.current) return
      if (toolRef.current === 'eraser') { if (event.target) canvas.remove(event.target); return }
      if (toolRef.current === 'bucket') {
        const table = horizontalTableFromTarget(event.target)
        const tableCellIndex = table ? horizontalTableCellIndexFromPointer(table, canvas.getScenePoint(event.e as MouseEvent | TouchEvent), event.subTargets ?? []) : null
        if (table && tableCellIndex != null) {
          fillHorizontalTableCell(table, tableCellIndex, fillRef.current)
          canvas.setActiveObject(table)
          callbacksRef.current.onSelectDrawing(pageId, tableCellIndex)
        } else if (event.target) {
          applyPaintBucketFill(event.target, fillRef.current)
          canvas.setActiveObject(horizontalTableFromTarget(event.target) ?? event.target)
          callbacksRef.current.onSelectDrawing(pageId)
        } else {
          const pointer = canvas.getScenePoint(event.e as MouseEvent | TouchEvent)
          void callbacksRef.current
            .fillPageImageWithBucket(pageId, pointer.x, pointer.y, fillRef.current)
            .catch(error => { window.alert(error instanceof Error ? error.message : String(error)); return true })
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
        if (event.target instanceof IText) {
          canvas.setActiveObject(event.target)
          event.target.enterEditing()
          moveTextCursorToEnd(event.target)
          canvas.requestRenderAll()
        }
        callbacksRef.current.onSelectDrawing(pageId, tableCellIndexFromEvent(event) ?? undefined)
        return
      }
      const pointer = canvas.getScenePoint(event.e as MouseEvent | TouchEvent)
      const activeObject = canvas.getActiveObject()
      if (activeObject instanceof IText && activeObject.isEditing) activeObject.exitEditing()
      const text = new IText('', {
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
      // Match OneNote's placement behavior: a click commits the previous
      // textbox in place and opens a new empty textbox at the new point.
      configureManagedText(text, 'plain', canvas, { removeIfEmpty: true })
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
    canvas.on('mouse:move', event => { if (toolRef.current === 'eraser' && event.e instanceof PointerEvent && event.e.buttons === 1 && event.target) canvas.remove(event.target) })
    registerCanvas(pageId, canvas)
    return () => {
      window.clearTimeout(recordTimer)
      recordNow()
      pendingCanvasWrites.delete(recordNow)
      registerCanvas(pageId, null)
      canvas.dispose()
      fabricRef.current = null
    }
  }, [pageHeight, pageId, pageWidth, registerCanvas])

  useEffect(() => {
    const canvas = fabricRef.current
    if (!canvas || canvasJson === appliedSnapshot.current) return
    let cancelled = false
    restoring.current = true
    restoreSerial.current = restoreSerial.current.catch(() => undefined).then(async () => {
      if (cancelled) return
      try {
        await canvas.loadFromJSON(canvasJson ? JSON.parse(canvasJson) : { objects: [] })
        if (cancelled) return
        restoreManagedCanvasObjects(canvas)
        canvas.getObjects().forEach(object=>{object.selectable=canSelectPaintRef.current || toolRef.current==='place-text';object.evented=['select-paint','place-text','bucket','eraser'].includes(toolRef.current)})
        appliedSnapshot.current = canvasSnapshot(canvas)
        canvas.requestRenderAll()
      } catch (error) { console.error('Drawing could not be loaded', error) }
      finally { if (!cancelled) restoring.current = false }
    })
    return () => { cancelled = true }
  }, [canvasJson, pageId, pageHeight, pageWidth])

  useEffect(() => {
    const canvas = fabricRef.current
    if (!canvas) return
    const paintSelectionActive = active && (canSelectPaint || tool === 'bucket' || tool === 'eraser' || tool === 'place-text')
    canvas.isDrawingMode = active && tool === 'brush'
    canvas.selection = active && canSelectPaint
    if(tool==='eraser')canvas.discardActiveObject()
    canvas.skipTargetFind = !paintSelectionActive
    canvas.getObjects().forEach((object) => {
      object.selectable = active && (canSelectPaint || tool==='place-text')
      object.evented = paintSelectionActive
    })
    if (canvas.isDrawingMode) {
      const brush = pressure ? new PressureBrush(canvas) : new PencilBrush(canvas)
      brush.width = brushSize
      brush.color = stroke
      canvas.freeDrawingBrush = brush
    }
    canvas.requestRenderAll()
  }, [active, brushSize, canSelectPaint, stroke, tool, pressure])

  const interactive = active && (tool === 'brush' || tool === 'eraser' || tool === 'bucket' || canSelectPaint || tool === 'place-text')

  return (
    <div
      className={interactive ? 'drawing-layer interactive' : 'drawing-layer'}
      onClick={(event) => event.stopPropagation()}
      onPointerDown={(event) => event.stopPropagation()}
      onPointerDownCapture={event => { if (penOnly && event.pointerType === 'touch') { event.preventDefault(); event.stopPropagation() } }}
      onPointerMoveCapture={event => { if (penOnly && event.pointerType === 'touch') { event.preventDefault(); event.stopPropagation() } }}
      onPointerUpCapture={event => { if (penOnly && event.pointerType === 'touch') { event.preventDefault(); event.stopPropagation() } }}
      style={{ zIndex, touchAction: 'none' }}
    >
      <span className="sr-only">{canvasText(canvasJson)}</span>
      <canvas aria-label={canvasText(canvasJson) || label} role="img" height={pageHeight} ref={canvasEl} width={pageWidth} />
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
  const viewScale=()=>{const paper=document.querySelector('.paper');return paper ? paper.getBoundingClientRect().width/pageWidth : 1}
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
        x: clamp(origin.x + (moveEvent.clientX - startX)/viewScale(), 0, pageWidth - image.width),
        y: clamp(origin.y + (moveEvent.clientY - startY)/viewScale(), 0, pageHeight - image.height),
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
    const startY = event.clientY
    const ratio = image.height / image.width
    const move = (moveEvent: PointerEvent) => {
      const nextWidth = clamp(image.width + (moveEvent.clientX - startX)/viewScale() * Math.cos(image.rotation * Math.PI / 180) + (moveEvent.clientY - startY)/viewScale() * Math.sin(image.rotation * Math.PI / 180), 64, pageWidth - image.x)
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
    const startY = event.clientY
    const move = (moveEvent: PointerEvent) => {
      updateImage({
        width: clamp(image.width + (moveEvent.clientX - startX)/viewScale() * Math.cos(image.rotation * Math.PI / 180) + (moveEvent.clientY - startY)/viewScale() * Math.sin(image.rotation * Math.PI / 180), 64, pageWidth - image.x),
      })
    }
    bindPointerDrag(move)
  }

  const startVerticalResize = (event: ReactPointerEvent<HTMLButtonElement>) => {
    event.preventDefault()
    event.stopPropagation()
    onSelect()
    onChangeStart()
    const startX = event.clientX
    const startY = event.clientY
    const move = (moveEvent: PointerEvent) => {
      updateImage({
        height: clamp(image.height - (moveEvent.clientX - startX)/viewScale() * Math.sin(image.rotation * Math.PI / 180) + (moveEvent.clientY - startY)/viewScale() * Math.cos(image.rotation * Math.PI / 180), 48, pageHeight - image.y),
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
    const centerX = paperBounds.left + (image.x + image.width / 2)*viewScale()
    const centerY = paperBounds.top + (image.y + image.height / 2)*viewScale()
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
      <img alt={image.alt ?? t('addImage')} draggable={false} src={image.src} />
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
  const menuRef = useRef<HTMLDivElement>(null)
  useEffect(() => {
    const menu = menuRef.current
    const previous = document.activeElement as HTMLElement | null
    menu?.querySelectorAll('button').forEach(button => button.setAttribute('role', 'menuitem'))
    menu?.querySelector<HTMLButtonElement>('button')?.focus()
    return () => { if (menu?.contains(document.activeElement)) previous?.focus() }
  }, [])
  return (
    <div
      ref={menuRef}
      role="menu"
      onKeyDown={event => {
        const buttons = Array.from(menuRef.current?.querySelectorAll('button') ?? [])
        const index = buttons.indexOf(document.activeElement as HTMLButtonElement)
        if (['ArrowDown','ArrowUp','Home','End'].includes(event.key)) {
          event.preventDefault()
          buttons[event.key === 'Home' ? 0 : event.key === 'End' ? buttons.length - 1 : (index + (event.key === 'ArrowDown' ? 1 : -1) + buttons.length) % buttons.length]?.focus()
        }
      }}
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
  | { completed: number; kind: 'converting'; total: number }
  | { completed: number; failed: number; kind: 'converted'; limitation: boolean; target: ExportFormat; total: number }
  | { kind: 'error'; message?: string }

function ConvertMode({ t }: { t: Translate }) {
  const [files, setFiles] = useState<File[]>([])
  const [target, setTarget] = useState<ExportFormat>('pdf')
  const [message, setMessage] = useState<ConvertMessage>({ kind: 'prompt' })
  const displayMessage = useMemo(() => {
    if (message.kind === 'converting') return `${t('converting')} ${message.completed}/${message.total}`
    if (message.kind === 'converted') {
      const limitation = message.limitation ? ` ${t('pdfDocxExtractedOnly')}` : ''
      const summary = t('conversionComplete', { completed: message.completed, format: message.target.toUpperCase(), total: message.total })
      const failures = message.failed ? ` ${t('conversionFailed', { count: message.failed })}` : ''
      return `${summary}${failures}${limitation}`
    }
    if (message.kind === 'error') return localizedErrorMessage(message.message, t, 'unavailableConversion')
    return t('convertPrompt')
  }, [message, t])

  const addFiles = (nextFiles: FileList | File[]) => {
    if (busy.current) return
    const supportedFiles = Array.from(nextFiles).filter((item) => ['docx', 'pdf', 'txt', 'md', 'markdown', 'png', 'jpg', 'jpeg'].includes(extensionOf(item.name)))
    setFiles((currentFiles) => {
      const knownFiles = new Set(currentFiles.map((item) => `${item.name}:${item.size}:${item.lastModified}`))
      return [...currentFiles, ...supportedFiles.filter((item) => !knownFiles.has(`${item.name}:${item.size}:${item.lastModified}`))]
    })
    setMessage({ kind: 'prompt' })
  }

  const busy = useRef(false)
  const [errors, setErrors] = useState<{ fileName:string; message:string }[]>([])
  const convert = async () => {
    if (!files.length || busy.current) return
    busy.current = true
    setErrors([])
    let completed = 0
    let failed = 0
    let limitation = false
    for (const file of files) {
      setMessage({ completed, kind: 'converting', total: files.length })
      try {
        await convertFile(file, target)
        completed += 1
        limitation ||= extensionOf(file.name) === 'pdf' && target === 'docx'
      } catch (error) {
        setErrors(current => [...current, {fileName:file.name,message:error instanceof Error ? error.message : String(error)}])
        failed += 1
      }
    }
    busy.current = false
    setMessage({ completed, failed, kind: 'converted', limitation, target, total: files.length })
  }

  return (
    <section className="mode-panel">
      <Ribbon title={t('convertAFIle')}>
        <label className="upload-button">
          <Upload size={17} />
          <span>{files.length ? t('filesQueued', { count: files.length }) : t('chooseFile')}</span>
          <input
            accept=".docx,.pdf,.txt,.md,.markdown,.png,.jpg,.jpeg"
            className="hidden-input"
            onChange={(event) => {
              if (event.target.files) addFiles(event.target.files)
              event.currentTarget.value = ''
            }}
            multiple
            type="file"
          />
        </label>
        <select disabled={message.kind === 'converting'} aria-label={t('targetFormat')} onChange={(event) => setTarget(event.target.value as ExportFormat)} value={target}>
          {exportFormats.map((format) => (
            <option disabled={files.some(file => !canConvert(file, format))} key={format} value={format}>
              {format.toUpperCase()}
            </option>
          ))}
        </select>
        <button className="primary-action" disabled={!files.length || message.kind === 'converting' || files.some(file => !canConvert(file, target))} onClick={() => void convert()} type="button">
          <ScissorsLineDashed size={17} />
          {files.length > 1 ? t('convertFiles', { count: files.length }) : t('convert')}
        </button>
      </Ribbon>
      <div className="converter-stage">
        <div
          className="drop-zone"
          onDragOver={(event) => event.preventDefault()}
          onDrop={(event) => {
            event.preventDefault()
            addFiles(event.dataTransfer.files)
          }}
        >
          <Upload size={42} />
          <h2>{files.length ? t('filesQueued', { count: files.length }) : t('convertDropZone')}</h2>
          <p role="status">{displayMessage}</p>
          {errors.map(error => <p role="alert" key={`${error.fileName}:${error.message}`}>{error.fileName}: {localizedErrorMessage(error.message,t,'unavailableConversion')}</p>)}
          {files.length > 0 && (
            <ul className="conversion-file-list" aria-label={t('filesQueued', { count: files.length })}>
              {files.map((file) => (
                <li key={`${file.name}:${file.size}:${file.lastModified}`}>
                  <FileText size={16} />
                  <span title={file.name}>{file.name}</span>
                  <small>{extensionOf(file.name).toUpperCase()}</small>
                  <button aria-label={`${t('removeFile')}: ${file.name}`} onClick={() => setFiles((currentFiles) => currentFiles.filter((item) => item !== file))} type="button">
                    <X size={15} />
                  </button>
                </li>
              ))}
            </ul>
          )}
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
  hint,
  active,
  icon: Icon,
  label,
  onClick,
}: {
  hint?: string
  active?: boolean
  icon: React.ComponentType<{ size?: number }>
  label: string
  onClick: () => void
}) {
  const button = (
    <button aria-label={label} aria-pressed={active} className={active ? 'icon-button active' : 'icon-button'} onClick={onClick} title={hint ? undefined : label} type="button">
      <Icon size={17} />
    </button>
  )
  return hint ? <HelpHint text={hint}>{button}</HelpHint> : button
}

function ToolButton({
  hint,
  active,
  icon: Icon,
  label,
  onClick,
}: {
  hint?: string
  active?: boolean
  icon: React.ComponentType<{ size?: number }>
  label: string
  onClick: () => void
}) {
  const button = (
    <button aria-label={label} aria-pressed={active} className={active ? 'tool-button active' : 'tool-button'} onClick={onClick} title={hint ? undefined : label} type="button">
      <Icon size={16} />
      <span>{label}</span>
    </button>
  )
  return hint ? <HelpHint text={hint}>{button}</HelpHint> : button
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
  const dialogRef = useRef<HTMLDivElement>(null)
  useEffect(() => {
    const previous = document.activeElement as HTMLElement | null
    dialogRef.current?.querySelector<HTMLButtonElement>('button')?.focus()
    return () => previous?.focus()
  }, [])
  return (
    <div className="dialog-backdrop" ref={dialogRef} onKeyDown={event => {
      if (event.key === 'Escape') { event.stopPropagation(); onCancel() }
      if (event.key !== 'Tab') return
      const buttons = Array.from(dialogRef.current?.querySelectorAll('button') ?? [])
      const index = buttons.indexOf(document.activeElement as HTMLButtonElement)
      event.preventDefault()
      buttons[(index + (event.shiftKey ? -1 : 1) + buttons.length) % buttons.length]?.focus()
    }}>
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

function canConvert(file: File, target: ExportFormat) {
  return !['png','jpg','jpeg'].includes(extensionOf(file.name)) || ['png','jpg','pdf','html','docx'].includes(target)
}

async function convertFile(file: File, target: ExportFormat) {
  const source = extensionOf(file.name)
  if (file.size > 50 * 1024 * 1024) throw new Error('File exceeds 50 MB.')
  if (!canConvert(file, target)) throw new Error('Unsupported conversion without OCR.')
  if (source === target || source === 'jpeg' && target === 'jpg') { saveBlob(file, file.name); return }
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
    if (target === 'docx') { saveBlob(await createDocxBlobFromPages(await pagesFromFile(file)), `${baseName}.docx`); return }
    if (target === 'html') {
      saveBlob(new Blob([wrapHtmlDocument(`<img src="${dataUrl}" alt="${file.name}" />`)], { type: 'text/html;charset=utf-8' }), `${baseName}.html`)
      return
    }
    throw new Error('Image to text, Markdown, or DOCX needs OCR, which is not included in this client-side build.')
  }

  if (source === 'pdf' && !['txt','md'].includes(target)) {
    const pages = await pagesFromPdf(file, undefined, target === 'docx')
    if (target === 'docx') saveBlob(await createDocxBlobFromPages(pages), `${baseName}.docx`)
    else if (target === 'html') saveBlob(new Blob([wrapHtmlDocument(await renderPagesHtml(pages))], { type:'text/html;charset=utf-8' }), `${baseName}.html`)
    else if (target === 'png' || target === 'jpg') saveBlob(await exportPagesImageBlob(pages,target), `${baseName}.${pages.length > 1 ? 'zip' : target}`)
    return
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
    saveBlob(new Blob([wrapHtmlDocument(await renderPagesHtml(await pagesFromFile(file)))], { type: 'text/html;charset=utf-8' }), `${baseName}.html`)
    return
  }
  if (target === 'docx') {
    saveBlob(await createDocxBlobFromPages(pagesFromHtml(html || textToEditableHtml(text))), `${baseName}.docx`)
    return
  }
  if (target === 'pdf') {
    saveBlob(await exportPagesPdfBlob(await pagesFromFile(file)), `${baseName}.pdf`)
    return
  }
  const pages = await pagesFromFile(file)
  saveBlob(await exportPagesImageBlob(pages, target), `${baseName}.${pages.length > 1 ? 'zip' : target}`)
}

async function pagesFromFile(file: File, signal?: AbortSignal) {
  const source = extensionOf(file.name)
  if (source === 'akxr') return (await readProjectFile(file)).pages
  if (source === 'pdf') return pagesFromPdf(file, signal)
  if (source === 'docx') return pagesFromDocx(file, signal)
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

async function readProjectFile(file: File): Promise<SavedProject> {
  try {
    return normalizeSavedProject(JSON.parse(await file.text()))
  } catch {
    throw new Error('Could not open that file.')
  }
}

function readAutosavedProject(): SavedProject | null {
  try {
    const saved = localStorage.getItem(autosaveKey)
    return saved ? normalizeSavedProject(JSON.parse(saved)) : null
  } catch {
    return null
  }
}

function normalizeSavedProject(value: unknown): SavedProject {
  if (!value || typeof value !== 'object') throw new Error('Invalid project.')
  const candidate = value as Partial<SavedProject>
  if (candidate.version !== 1 || !Array.isArray(candidate.pages) || !candidate.pages.length || candidate.pages.length > 500) throw new Error('Invalid project pages.')
  const ids = new Set<string>()
  const normalizePage = (page: PageData): PageData => {
    if (!page || typeof page !== 'object' || typeof page.id !== 'string' || !page.id || typeof page.html !== 'string' || !Array.isArray(page.images)) throw new Error('Invalid page.')
    const images = page.images.map(image => {
      if (!image || typeof image.id !== 'string' || typeof image.src !== 'string' || !isSafeImportedUrl(image.src, true) || ![image.x,image.y,image.width,image.height,image.rotation].every(n => typeof n === 'number' && Number.isFinite(n) && Math.abs(n) < 100_000) || image.width <= 0 || image.height <= 0) throw new Error('Invalid image.')
      return { id: image.id, src: image.src, x: image.x, y: image.y, width: image.width, height: image.height, rotation: image.rotation, alt: typeof image.alt === 'string' ? image.alt : undefined }
    })
    if (page.pageImage && (typeof page.pageImage !== 'string' || !isSafeImportedUrl(page.pageImage, true))) throw new Error('Invalid page image.')
    const dimension = (value: unknown) => { if (value === undefined) return undefined; if (typeof value !== 'number' || !Number.isFinite(value) || value < 100 || value > 30000) throw new Error('Invalid dimensions.'); return value }
    const attachments = page.attachments?.map(attachment => {
      if (!attachment || typeof attachment.id !== 'string' || typeof attachment.name !== 'string' || typeof attachment.data !== 'string' || !/^data:[^,]*;base64,[a-z0-9+/=]*$/i.test(attachment.data)) throw new Error('Invalid attachment.')
      return { id: attachment.id, name: attachment.name, type: typeof attachment.type === 'string' ? attachment.type : '', data: attachment.data }
    })
    const text = (value: unknown) => typeof value === 'string' ? value.slice(0, 1000) : undefined
    return { direction: page.direction === 'rtl' || candidate.language === 'ar' ? 'rtl' : 'ltr', id: page.id, html: sanitizeImportedHtml(page.html), images, pageImage: page.pageImage,
      canvasJson: validateCanvasJson(page.canvasJson), size: page.size && page.size in pageSizes ? page.size : defaultPageSizeKey,
      margins: exactPageMargins(page.margins), fontFamily: typeof page.fontFamily === 'string' && /^[\p{L}\p{N} ,"'_-]+$/u.test(page.fontFamily) ? page.fontFamily : undefined, title: text(page.title), section: text(page.section), tags: text(page.tags),
      createdAt: text(page.createdAt), updatedAt: text(page.updatedAt), attachments,
      width: dimension(page.width), height: dimension(page.height), baseHeight: dimension(page.baseHeight),
      lineHeight: typeof page.lineHeight === 'number' && Number.isFinite(page.lineHeight) ? clamp(page.lineHeight, 1, 3) : 1.72,
      layerOrder: Array.isArray(page.layerOrder) && page.layerOrder.length === 3 && new Set(page.layerOrder).size === 3 && page.layerOrder.every(layer => ['text','images','paint'].includes(layer)) ? page.layerOrder : ['text','images','paint'] }
  }
  const pages = candidate.pages.map(page => {
    if (ids.has(page.id)) throw new Error('Duplicate page ID.')
    ids.add(page.id)
    return normalizePage(page)
  })
  if (candidate.trash && (!Array.isArray(candidate.trash) || candidate.trash.length > 500)) throw new Error('Invalid trash.')
  const trash = candidate.trash?.map(normalizePage) ?? []
  const activePageId = typeof candidate.activePageId === 'string' && ids.has(candidate.activePageId) ? candidate.activePageId : pages[0].id
  return { activePageId, language: candidate.language === 'ar' || candidate.language === 'de' || candidate.language === 'da' ? candidate.language : 'en', pages, trash, version: 1 }
}

function validateCanvasJson(value: unknown): string | undefined {
  if (value === undefined || value === '') return undefined
  if (typeof value !== 'string' || value.length > 25_000_000) throw new Error('Invalid drawing.')
  const data = JSON.parse(value)
  if (!data || !Array.isArray(data.objects)) throw new Error('Invalid drawing objects.')
  const types = new Set(['rect','circle','ellipse','line','path','polygon','polyline','triangle','group','text','i-text','itext','textbox','image','activeselection'])
  let count = 0
  const walk = (value: unknown, depth: number) => {
    if (depth > 30 || ++count > 200_000) throw new Error('Drawing too complex.')
    if (typeof value === 'number' && (!Number.isFinite(value) || Math.abs(value) > 1e8)) throw new Error('Invalid drawing number.')
    if (!value || typeof value !== 'object') return
    if (Array.isArray(value)) { value.forEach(item => walk(item, depth + 1)); return }
    const object = value as Record<string, unknown>
    for (const [key, item] of Object.entries(object)) {
      if (['__proto__','prototype','constructor'].includes(key)) throw new Error('Invalid drawing property.')
      if (key === 'src' && (typeof item !== 'string' || !isSafeImportedUrl(item, true))) throw new Error('Invalid drawing image.')
      if (typeof item === 'string' && /url\s*\(|expression\s*\(/i.test(item)) throw new Error('Invalid drawing style.')
      walk(item, depth + 1)
    }
    if ('objects' in object && (!Array.isArray(object.objects) || object.objects.length > 10000)) throw new Error('Invalid drawing group.')
  }
  const checkObjects = (objects: Record<string, unknown>[]) => objects.forEach(object => {
    if (!object || typeof object.type !== 'string' || !types.has(object.type.toLowerCase())) throw new Error('Unsupported drawing object.')
    const type = object.type.toLowerCase()
    if (['text','itext','i-text','textbox'].includes(type) && typeof object.text !== 'string') throw new Error('Invalid drawing text.')
    if (type === 'group' && !Array.isArray(object.objects)) throw new Error('Invalid drawing group.')
    if (type === 'path' && (!Array.isArray(object.path) || object.path.some(command => !Array.isArray(command) || typeof command[0] !== 'string' || !/^[MLCQZHVSTA]$/i.test(command[0]) || command.slice(1).some(value => typeof value !== 'number' || !Number.isFinite(value))))) throw new Error('Invalid drawing path.')
    if (type === 'image' && (typeof object.src !== 'string' || !isSafeImportedUrl(object.src, true))) throw new Error('Invalid drawing image.')
    for (const key of ['left','top','width','height','scaleX','scaleY','angle','fontSize','strokeWidth','opacity']) {
      if (object[key] !== undefined && (typeof object[key] !== 'number' || !Number.isFinite(object[key]))) throw new Error('Invalid drawing geometry.')
    }
    if (object.akxrTextMode && !['plain','bullet','numbered','table'].includes(String(object.akxrTextMode))) throw new Error('Invalid text mode.')
    if (object.objects) checkObjects(object.objects as Record<string, unknown>[])
    if (object.clipPath) checkObjects([object.clipPath as Record<string, unknown>])
  })
  walk(data, 0); checkObjects(data.objects)
  if (data.backgroundImage) checkObjects([data.backgroundImage])
  if (data.overlayImage) checkObjects([data.overlayImage])
  if (data.width && data.height && data.width * data.height > 50_000_000) throw new Error('Drawing exceeds rendering budget.')
  return JSON.stringify(data)
}

async function pagesFromDocx(file: File, signal?: AbortSignal): Promise<PageData[]> {
  const { default: JSZip } = await import('jszip')
  const mammoth = await import('mammoth')
  const data = await file.arrayBuffer()
  const zip = await JSZip.loadAsync(data)
  const source = await zip.file('word/document.xml')?.async('text')
  if (!source || source.length > 20_000_000) throw new Error('Invalid or oversized Word document.')
  const document = new DOMParser().parseFromString(source, 'application/xml')
  const body = xmlFirst(document, 'body')
  if (!body) throw new Error('Invalid Word document.')
  const groups: { elements: Element[]; section?: Element }[] = []
  let elements: Element[] = []
  for (const child of Array.from(body.children)) {
    if (child.localName === 'sectPr') { groups.push({ elements, section: child }); elements = []; continue }
    elements.push(child)
    const section = xmlFirst(child, 'sectPr')
    if (section) { groups.push({ elements, section }); elements = [] }
  }
  if (elements.length) groups.push({ elements })
  if (groups.length > 100) throw new Error('Word document exceeds 100 sections.')
  const fallback = await readDocxLayout(data)
  const pages: PageData[] = []
  for (const group of groups) {
    signal?.throwIfAborted()
    const sectionDocument = document.cloneNode(true) as Document
    const sectionBody = xmlFirst(sectionDocument, 'body')!
    sectionBody.replaceChildren(...group.elements.map(element => element.cloneNode(true)))
    sectionBody.querySelectorAll('*').forEach(element => { if (element.localName === 'sectPr') element.remove() })
    if (group.section) sectionBody.append(group.section.cloneNode(true))
    zip.file('word/document.xml', new XMLSerializer().serializeToString(sectionDocument))
    const buffer = await zip.generateAsync({ type: 'arraybuffer' })
    const converted = await mammoth.convertToHtml({ arrayBuffer: buffer }, { ignoreEmptyParagraphs: false })
    const layout = group.section ? { ...fallback, margins: marginsFromSectPr(group.section), size: pageSizeFromSectPr(group.section) } : fallback
    const dimensions = group.section ? xmlFirst(group.section, 'pgSz') : undefined
    const width = Number(xmlAttr(dimensions, 'w')) / 15, height = Number(xmlAttr(dimensions, 'h')) / 15
    const sectionPages = pagesFromHtml(converted.value, layout)
    for (const page of sectionPages) {
      if (width >= 100 && height >= 100) { page.width = Math.round(width); page.baseHeight = Math.round(height); page.height = Math.max(Math.round(height), page.height ?? height) }
      pages.push(page)
    }
  }
  return pages.length ? pages : pagesFromText('')
}

async function readDocxLayout(arrayBuffer: ArrayBuffer): Promise<Partial<Pick<PageData, 'fontFamily' | 'margins' | 'size'>>> {
  const { default: JSZip } = await import('jszip')
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
    top: twipsToPx(xmlAttr(pageMargins, 'top') === '' ? defaultPageMargins.top * 15 : Number(xmlAttr(pageMargins, 'top'))),
    right: twipsToPx(xmlAttr(pageMargins, 'right') === '' ? defaultPageMargins.right * 15 : Number(xmlAttr(pageMargins, 'right'))),
    bottom: twipsToPx(xmlAttr(pageMargins, 'bottom') === '' ? defaultPageMargins.bottom * 15 : Number(xmlAttr(pageMargins, 'bottom'))),
    left: twipsToPx(xmlAttr(pageMargins, 'left') === '' ? defaultPageMargins.left * 15 : Number(xmlAttr(pageMargins, 'left'))),
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

async function pagesFromPdf(file: File, signal?: AbortSignal, includeText = false) {
  if (file.size > 50 * 1024 * 1024) throw new Error('PDF exceeds 50 MB.')
  const pdfjsLib = await pdfLibrary()
  const task = pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) })
  const abort = () => { void task.destroy() }
  signal?.addEventListener('abort', abort, { once: true })
  const pages: PageData[] = []
  try {
    signal?.throwIfAborted()
    const pdf = await task.promise
    if (pdf.numPages > 150) throw new Error('PDF exceeds 150 pages. Split the document first.')
    let pixels = 0
    for (let number = 1; number <= pdf.numPages; number++) {
      signal?.throwIfAborted()
      const page = await pdf.getPage(number)
      const viewport = page.getViewport({ scale: 96 / 72 })
      pixels += viewport.width * viewport.height
      if (pixels > 100_000_000 || viewport.width > 8000 || viewport.height > 8000) throw new Error('PDF exceeds the rendering budget.')
      const canvas = document.createElement('canvas')
      canvas.width = Math.ceil(viewport.width); canvas.height = Math.ceil(viewport.height)
      const context = canvas.getContext('2d')
      if (!context) throw new Error('Canvas rendering is not available in this browser.')
      await page.render({ canvas, canvasContext: context, viewport }).promise
      const extracted = includeText ? (await page.getTextContent()).items.map(item => 'str' in item ? item.str + (item.hasEOL ? '\n' : ' ') : '').join('') : ''
      pages.push({ title: `${file.name} — ${number}`, id: createId(), html: extracted ? textToEditableHtml(extracted) : starterContent, images: [], width: Math.round(viewport.width), height: Math.round(viewport.height), baseHeight: Math.round(viewport.height), size: viewport.width > viewport.height ? 'a4Landscape' : 'a4', pageImage: canvas.toDataURL('image/jpeg', 0.9) })
      canvas.width = canvas.height = 0
      page.cleanup()
      await new Promise(resolve => window.setTimeout(resolve, 0))
    }
    return pages
  } finally { signal?.removeEventListener('abort', abort); await task.destroy() }
}

async function readTextualFile(file: File, source: string) {
  if (source === 'docx') {
    const mammoth = await import('mammoth')
    const arrayBuffer = await file.arrayBuffer()
    const layout = await readDocxLayout(arrayBuffer)
    const text = (await mammoth.extractRawText({ arrayBuffer })).value
    const html = (await mammoth.convertToHtml({ arrayBuffer }, { ignoreEmptyParagraphs: false })).value
    return { text, html, layout }
  }
  if (source === 'pdf') {
    const text = await extractPdfText(file)
    if (!text.trim()) throw new Error('No extractable PDF text. OCR is not included.')
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
    const html = sanitizeImportedHtml(body.innerHTML)
    return { text: textFromHtml(html), html }
  }
  throw new Error('Unsupported source file. Use DOCX, PDF, TXT, MD, HTML, PNG, or JPG.')
}

function pagesFromText(text: string) {
  return pagesFromHtml(textToEditableHtml(text))
}

function pagesFromHtml(html: string, layout: Partial<Pick<PageData, 'fontFamily' | 'margins' | 'size'>> = {}): PageData[] {
  const normalizedLayout = {
    ...layout,
    margins: normalizePageMargins(layout.margins),
    size: layout.size ?? defaultPageSizeKey,
  }
  const parser = new DOMParser()
  const document = parser.parseFromString(`<main>${sanitizeImportedHtml(html || starterContent)}</main>`, 'text/html')
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
    height: measuredHtmlHeight(pageHtml, normalizedLayout),
    baseHeight: pageSizeFor(normalizedLayout.size).height,
  }))
}

function htmlAtomsFromNode(node: ChildNode): string[] {
  if (node.nodeType === Node.TEXT_NODE) return node.textContent?.trim() ? [`<p>${escapeHtml(node.textContent)}</p>`] : []
  if (!(node instanceof HTMLElement)) return []
  if (['DIV','MAIN','SECTION','ARTICLE','BODY'].includes(node.tagName)) return Array.from(node.childNodes).flatMap(htmlAtomsFromNode)
  return [node.outerHTML]
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

function measuredHtmlHeight(html: string, layout: Partial<Pick<PageData, 'fontFamily' | 'margins' | 'size'>>) {
  const measure = createPageMeasurer(layout)
  try { return measure.height(html) } finally { measure.remove() }
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
    height(pageHtml: string) { content.style.height = 'auto'; content.style.minHeight = `${size.height}px`; content.innerHTML = pageHtml; return Math.max(size.height, content.scrollHeight) },
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
  return sanitizeImportedHtml(marked.parse(markdown, { async: false, gfm: true }) as string)
}

async function extractPdfText(file: File) {
  const pdfjsLib = await pdfLibrary()
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
  const document=new DOMParser().parseFromString(sanitizeImportedHtml(html),'text/html')
  document.querySelectorAll('[data-ink-cut]').forEach(element=>element.remove())
  return new TurndownService({ headingStyle: 'atx', bulletListMarker: '-' }).turndown(document.body.innerHTML)
}

async function createDocxBlobFromPages(pages: PageData[], _canvasDataUrlForPage?: (pageId: string) => string | undefined) {
  const { Document: DocxDocument, Paragraph, TextRun, Packer, Table: WordTable, TableRow: WordRow, TableCell: WordCell, ExternalHyperlink, ImageRun } = await import('docx')
  type Run = InstanceType<typeof TextRun> | InstanceType<typeof ExternalHyperlink> | InstanceType<typeof ImageRun>
  const runs = async (nodes: NodeListOf<ChildNode> | ChildNode[], style: { bold?: boolean; italics?: boolean; underline?: object; color?: string; size?: number; font?: string } = {}): Promise<Run[]> => {
    const result: Run[] = []
    for (const node of Array.from(nodes)) {
      if (node.nodeType === Node.TEXT_NODE) { result.push(new TextRun({ ...style, text: node.textContent ?? '' })); continue }
      if (!(node instanceof HTMLElement)) continue
      const cut=parseInkCut(node.getAttribute('data-ink-cut'))
      if(cut) {
        if(!fullyErased(cut)) { const blob=await maskedGlyphBlob(node,style);result.push(new ImageRun({type:'png',data:new Uint8Array(await blob.arrayBuffer()),transformation:{width:cut.width,height:cut.height}})) }
        continue
      }
      if (node.tagName === 'BR') { result.push(new TextRun({ break: 1 })); continue }
      const next = { ...style }
      if (['B','STRONG'].includes(node.tagName)) next.bold = true
      if (['I','EM'].includes(node.tagName)) next.italics = true
      if (node.tagName === 'U') next.underline = {}
      if (node.style.color) { const rgb = colorToRgba(node.style.color); next.color = rgb.slice(0,3).map(n => n.toString(16).padStart(2,'0')).join('') }
      if (node.style.fontSize) next.size = Math.round(parseFloat(node.style.fontSize) * 1.5)
      if (node.style.fontFamily) next.font = node.style.fontFamily
      if (node.tagName === 'IMG' && node.getAttribute('src')) {
        const image = await convertImageBlob(node.getAttribute('src')!, 'png')
        const loaded = await loadImage(node.getAttribute('src')!)
        const width = Math.min(600, loaded.naturalWidth)
        result.push(new ImageRun({ data: new Uint8Array(await image.arrayBuffer()), type: 'png', transformation: { width, height: width * loaded.naturalHeight / loaded.naturalWidth } }))
      } else if (node.tagName === 'A' && node.getAttribute('href')) {
        result.push(new ExternalHyperlink({ link: node.getAttribute('href')!, children: await runs(node.childNodes, next) }))
      } else result.push(...await runs(node.childNodes, next))
    }
    return result
  }
  const sections = []
  for (const page of pages) {
    const size = dimensionsFor(page)
    const doc = new DOMParser().parseFromString(sanitizeImportedHtml(page.html), 'text/html')
    const children: (InstanceType<typeof Paragraph> | InstanceType<typeof WordTable>)[] = []
    const paragraph = async (node: HTMLElement) => new Paragraph({
      children: await runs(node.childNodes),
      alignment: node.style.textAlign === 'justify' ? 'both' : ['left','center','right'].includes(node.style.textAlign) ? node.style.textAlign as 'left' | 'center' | 'right' : undefined,
      heading: /^H[1-6]$/.test(node.tagName) ? `Heading${node.tagName[1]}` as 'Heading1' : undefined,
      spacing: { after: 120, line: Math.round(240 * (page.lineHeight ?? 1.72)) },
    })
    for (const node of Array.from(doc.body.children)) {
      if (!(node instanceof HTMLElement)) continue
      if (node.tagName === 'TABLE') {
        children.push(new WordTable({ rows: await Promise.all(Array.from(node.querySelectorAll('tr')).map(async row => new WordRow({ children: await Promise.all(Array.from(row.children).map(async cell => new WordCell({ children: [await paragraph(cell as HTMLElement)] }))) }))) }))
      } else if (node.tagName === 'UL' || node.tagName === 'OL') {
        let index = 1
        for (const item of Array.from(node.querySelectorAll('li'))) children.push(new Paragraph({ children: [new TextRun(node.tagName === 'UL' ? '• ' : `${index++}. `), ...await runs(item.childNodes)] }))
      } else children.push(await paragraph(node))
    }
    if (canvasText(page.canvasJson)) children.push(new Paragraph({ children: [new TextRun(canvasText(page.canvasJson))] }))
    for (const source of [page.pageImage, ...page.images.map(image => image.src)].filter((src): src is string => Boolean(src))) {
      const blob = await convertImageBlob(source, 'png')
      const image = await loadImage(source)
      const width = Math.min(size.width - 100, image.naturalWidth)
      children.push(new Paragraph({ children: [new ImageRun({ type: 'png', data: new Uint8Array(await blob.arrayBuffer()), transformation: { width, height: width * image.naturalHeight / image.naturalWidth } })] }))
    }
    // Export canvas text once as editable text and rasterize only the remaining drawing objects.
    if (page.canvasJson) {
      const drawing = JSON.parse(page.canvasJson)
      const withoutText = (objects: Record<string, unknown>[]): Record<string, unknown>[] => objects.flatMap(object => {
        if (typeof object.text === 'string') return []
        if (!Array.isArray(object.objects)) return [object]
        const remaining = withoutText(object.objects)
        return remaining.length ? [{ ...object, objects: remaining }] : []
      })
      const objects = withoutText(drawing.objects)
      const visibleBackground = drawing.background && drawing.background !== 'transparent' && (typeof drawing.background !== 'string' || colorToRgba(drawing.background)[3] > 0)
      if (objects.length || visibleBackground || drawing.backgroundImage) {
        const source = await canvasDataFor({ ...page, canvasJson: JSON.stringify({ ...drawing, objects }) })
        if (source) children.push(new Paragraph({ children: [new ImageRun({ type: 'png', data: new Uint8Array(await (await fetch(source)).arrayBuffer()), transformation: { width: size.width - 100, height: (size.width - 100) * size.height / size.width } })] }))
      }
    }
    const margins = normalizePageMargins(page.margins)
    sections.push({ properties: { page: { size: { width: Math.round(size.width * 15), height: Math.round(size.height * 15) }, margin: { top: margins.top * 15, right: margins.right * 15, bottom: margins.bottom * 15, left: margins.left * 15 } } }, children })
  }
  return Packer.toBlob(new DocxDocument({ sections }))
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

async function renderPagesHtml(pages: PageData[], canvasDataUrlForPage?: (pageId: string) => string | undefined) {
  const html = []
  for (const page of pages) html.push(renderPageHtml(page, await canvasDataFor(page, canvasDataUrlForPage?.(page.id))))
  return html.join('')
}

function renderPageHtml(page: PageData, canvasDataUrl?: string) {
  const size = dimensionsFor(page)
  const margins = normalizePageMargins(page.margins)
  return `<section class="export-page" dir="${page.direction ?? 'ltr'}" data-page-size="${page.size ?? defaultPageSizeKey}" style="--paper-width:${size.width}px;--paper-height:${size.height}px;--page-padding-top:${margins.top}px;--page-padding-right:${margins.right}px;--page-padding-bottom:${margins.bottom}px;--page-padding-left:${margins.left}px;--page-font-family:${escapeHtml(fontStackFor(page.fontFamily))};--page-font-size:${defaultFontSize};--editor-line-height:${page.lineHeight ?? 1.72};">${
    page.pageImage ? `<img class="export-page-image" src="${page.pageImage}" alt="">` : ''
  }<div class="export-page-content" style="z-index:${(page.layerOrder ?? ['text','images','paint']).indexOf('text') + 1}">${page.html}</div>${
    canvasDataUrl ? `<img class="export-canvas-layer" style="z-index:${(page.layerOrder ?? ['text','images','paint']).indexOf('paint') + 1}" src="${canvasDataUrl}" alt="">` : ''
  }${page.images
    .map(
      (image) =>
        `<img class="export-floating-image" src="${image.src}" style="z-index:${(page.layerOrder ?? ['text','images','paint']).indexOf('images') + 1};left:${image.x}px;top:${image.y}px;width:${image.width}px;height:${image.height}px;transform:rotate(${image.rotation}deg);" alt="">`,
    )
    .join('')}</section>`
}

function wrapHtmlDocument(body: string, title = 'AKXREDITOR Export', language: AppLanguage = 'en') {
  return `<!doctype html><html lang="${language}" dir="${language === 'ar' ? 'rtl' : 'ltr'}"><head><meta charset="utf-8"><title>${escapeHtml(title)}</title><style>${exportDocumentCss()}</style></head><body>${body}</body></html>`
}

async function canvasDataFor(page: PageData, live?: string) {
  if (live || !page.canvasJson) return live
  const size = dimensionsFor(page)
  const canvas = new StaticCanvas(undefined, { width: size.width, height: size.height, enableRetinaScaling: false })
  try { await canvas.loadFromJSON(JSON.parse(page.canvasJson)); canvas.renderAll(); return canvas.toDataURL({ multiplier: 1, format: 'png' }) }
  finally { await canvas.dispose() }
}

async function exportPagesImageBlob(pages: PageData[], format: 'png' | 'jpg', canvasDataUrlForPage?: (pageId: string) => string | undefined) {
  if (pages.length > 1) {
    const { default: JSZip } = await import('jszip')
    const zip = new JSZip()
    for (let i = 0; i < pages.length; i++) {
      const canvas = await renderPageToCanvas(pages[i], canvasDataUrlForPage?.(pages[i].id), 1.5, format)
      zip.file(`page-${i + 1}.${format}`, await canvasToBlob(canvas, mimeForImageFormat(format), 0.95))
      canvas.width = canvas.height = 0
    }
    return zip.generateAsync({ type: 'blob' })
  }
  const canvas = await renderPageToCanvas(pages[0], canvasDataUrlForPage?.(pages[0].id), 2, format)
  try { return await canvasToBlob(canvas, mimeForImageFormat(format), 0.95) } finally { canvas.width = canvas.height = 0 }
}

async function exportPageImageBlobs(pages: PageData[], format: 'png' | 'jpg', canvasDataUrlForPage?: (pageId: string) => string | undefined) {
  const blobs = []
  for (const page of pages) {
    const canvas = await renderPageToCanvas(page, canvasDataUrlForPage?.(page.id), 2, format)
    blobs.push(await canvasToBlob(canvas, mimeForImageFormat(format), 0.95))
    canvas.width = canvas.height = 0
  }
  return blobs
}

async function exportPagesPdfBlob(pages: PageData[], canvasDataUrlForPage?: (pageId: string) => string | undefined) {
  const { jsPDF } = await import('jspdf')
  const first = dimensionsFor(pages[0])
  const pdf = new jsPDF({ compress: true, format: [first.width * 25.4 / 96, first.height * 25.4 / 96], orientation: first.width > first.height ? 'landscape' : 'portrait', unit: 'mm' })
  for (let i = 0; i < pages.length; i++) {
    const page = pages[i], size = dimensionsFor(page)
    const width = size.width * 25.4 / 96, height = size.height * 25.4 / 96
    if (i) pdf.addPage([width,height], width > height ? 'landscape' : 'portrait')
    const canvas = await renderPageToCanvas(page, canvasDataUrlForPage?.(page.id), 1.5, 'jpg')
    pdf.addImage(canvas.toDataURL('image/jpeg', 0.92), 'JPEG', 0, 0, width, height)
    canvas.width = canvas.height = 0
  }
  return pdf.output('blob')
}

async function renderPageToCanvas(page: PageData, canvasDataUrl: string | undefined, scale: number, format: 'png' | 'jpg') {
  const size = dimensionsFor(page)
  if (size.width * size.height * scale * scale > 50_000_000) throw new Error('Page exceeds image export size. Split the page first.')
  canvasDataUrl = await canvasDataFor(page, canvasDataUrl)
  const svgBody = xhtmlForSvg(`<style>${exportDocumentCss()}html,body{background:#fff;}.export-page{margin:0;box-shadow:none;}</style>${renderPageHtml(page, canvasDataUrl)}`)
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${size.width * scale}" height="${size.height * scale}" viewBox="0 0 ${size.width} ${size.height}"><foreignObject width="${size.width}" height="${size.height}">${svgBody}</foreignObject></svg>`
  const image = await loadImage(`data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`)
  const canvas = document.createElement('canvas'), context = canvas.getContext('2d')
  if (!context) throw new Error('Canvas export is not available in this browser.')
  canvas.width = size.width * scale; canvas.height = size.height * scale
  if (format === 'jpg') { context.fillStyle = '#ffffff'; context.fillRect(0,0,canvas.width,canvas.height) }
  context.drawImage(image, 0, 0, canvas.width, canvas.height)
  return canvas
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
.export-page-content{position:relative;z-index:1;height:100%;padding:var(--page-padding-top) var(--page-padding-right) var(--page-padding-bottom) var(--page-padding-left);overflow:hidden;font-family:var(--page-font-family);font-size:var(--page-font-size);line-height:var(--editor-line-height,1.72);text-align:start;}
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
  // Some browsers do not start an object-URL download synchronously. Keeping
  // it alive briefly prevents zero-byte or cancelled downloads.
  window.setTimeout(() => URL.revokeObjectURL(url), 2_000)
}

function saveDataUrl(url: string, fileName: string) {
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = fileName
  anchor.style.display = 'none'
  document.body.appendChild(anchor)
  anchor.click()
  anchor.remove()
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
    applyCanvasSelectionStyle(object)
    sharpenCanvasText(object)
    if (isHorizontalTableGroup(object)) configureHorizontalTableGroup(object, canvas)
    else if (object instanceof IText) configureManagedText(object, (object as ManagedIText).akxrTextMode ?? 'plain', canvas, { removeIfEmpty: true })
  })
}

function applyCanvasSelectionStyle(object: FabricObject) {
  object.set({
    borderColor: '#7c5cff',
    borderDashArray: [5, 4],
    cornerColor: '#f9f7ff',
    cornerSize: 11,
    cornerStrokeColor: '#7c5cff',
    cornerStyle: 'circle',
    padding: 0,
    perPixelTargetFind: true,
    transparentCorners: false,
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
  if (!value) return true
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
        originX: 'left',
        originY: 'top',
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
    const source = texts[index]
    if (source) { const { left, top, width, height } = text; text.set({ ...source.toObject(), left, top, width, height }) }
    const sourceRect = table.getObjects()[index * 2]
    const nextRect = nextTable.getObjects()[index * 2]
    if (sourceRect instanceof Rect && nextRect instanceof Rect) nextRect.set({ fill: sourceRect.fill, stroke: sourceRect.stroke })
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
    nextTexts[index]?.set({ styles: text.styles, textAlign: text.textAlign, backgroundColor: text.backgroundColor })
    const originalIndex = texts.indexOf(text)
    const previousRect = table.getObjects()[originalIndex * 2], nextRect = nextTable.getObjects()[index * 2]
    if (previousRect instanceof Rect && nextRect instanceof Rect) nextRect.set({ fill: previousRect.fill, stroke: previousRect.stroke })
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
  const fillColor = colorToRgba(fill)
  let imageData = context.getImageData(0, 0, canvas.width, canvas.height)
  const data = imageData.data
  const startIndex = (y * canvas.width + x) * 4
  const targetColor = [data[startIndex], data[startIndex + 1], data[startIndex + 2], data[startIndex + 3]]
  if (colorsClose(targetColor, fillColor, 0)) return dataUrl

  if (canvas.width * canvas.height > 16_000_000) throw new Error('Image exceeds fill budget.')
  const worker = new Worker(new URL('./fill.worker.ts', import.meta.url), { type: 'module' })
  try {
    const filled = await new Promise<ArrayBuffer>((resolve, reject) => {
      worker.onmessage = event => resolve(event.data)
      worker.onerror = event => reject(new Error(event.message))
      worker.postMessage({ pixels: data.buffer, width: canvas.width, height: canvas.height, x, y, fill: fillColor }, [data.buffer])
    })
    imageData = new ImageData(new Uint8ClampedArray(filled), canvas.width, canvas.height)
  } finally { worker.terminate() }
  context.putImageData(imageData, 0, 0)
  return canvas.toDataURL('image/png')
}

function colorWithTransparency(color: string, transparency: number) {
  const [red, green, blue] = colorToRgba(color)
  const alpha = 1 - clamp(transparency, 0, 100) / 100
  return `rgba(${red}, ${green}, ${blue}, ${alpha.toFixed(2)})`
}

function colorToRgba(color: string) {
  const rgba = color.trim().match(/^rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})(?:\s*,\s*(\d*(?:\.\d+)?))?\s*\)$/i)
  if (rgba) {
    return [
      clamp(Number.parseInt(rgba[1], 10), 0, 255),
      clamp(Number.parseInt(rgba[2], 10), 0, 255),
      clamp(Number.parseInt(rgba[3], 10), 0, 255),
      clamp(Math.round((rgba[4] == null ? 1 : Number.parseFloat(rgba[4])) * 255), 0, 255),
    ]
  }

  const hex = color.trim().replace(/^#/, '')
  if (/^[0-9a-f]{6}$/i.test(hex)) {
    return [
      Number.parseInt(hex.slice(0, 2), 16),
      Number.parseInt(hex.slice(2, 4), 16),
      Number.parseInt(hex.slice(4, 6), 16),
      255,
    ]
  }
  if (/^[0-9a-f]{8}$/i.test(hex)) {
    return [
      Number.parseInt(hex.slice(0, 2), 16),
      Number.parseInt(hex.slice(2, 4), 16),
      Number.parseInt(hex.slice(4, 6), 16),
      Number.parseInt(hex.slice(6, 8), 16),
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
  const doc = new DOMParser().parseFromString(sanitizeImportedHtml(html), 'text/html')
  doc.querySelectorAll<HTMLElement>('[data-ink-cut]').forEach(element=>{if(parseInkCut(element.getAttribute('data-ink-cut')))element.remove()})
  doc.querySelectorAll('br').forEach(element => element.replaceWith('\n'))
  doc.querySelectorAll('p,div,section,li,h1,h2,h3,h4,h5,h6,blockquote,pre,tr').forEach(element => element.append('\n'))
  doc.querySelectorAll('td,th').forEach(element => element.append('\t'))
  return (doc.body.textContent ?? '').replace(/\t\n/g, '\n').replace(/\n{3,}/g, '\n\n').trimEnd()
}

// Native selectionchange can arrive after a toolbar control receives focus.
// Preserve the actual highlighted range before commands restore editor focus.
function syncEditorSelection(editor: Editor) {
  const selection = editor.view.dom.ownerDocument.getSelection()
  if (!selection?.anchorNode || !selection.focusNode || !editor.view.dom.contains(selection.anchorNode) || !editor.view.dom.contains(selection.focusNode)) return
  const from = editor.view.posAtDOM(selection.anchorNode, selection.anchorOffset)
  const to = editor.view.posAtDOM(selection.focusNode, selection.focusOffset)
  if (editor.state.selection.anchor !== from || editor.state.selection.head !== to) editor.commands.setTextSelection({ from, to })
}

function sanitizeImportedHtml(html: string) {
  const clean = DOMPurify.sanitize(html, {
    USE_PROFILES: { html: true },
    FORBID_TAGS: ['style', 'form', 'input', 'button', 'textarea', 'select', 'video', 'audio'],
    FORBID_ATTR: ['srcset', 'formaction', 'action', 'id', 'name'],
  })
  const doc = new DOMParser().parseFromString(clean, 'text/html')
  const allowed = new Set(['background-color', 'color', 'font-family', 'font-size', 'font-style', 'font-weight', 'text-align', 'text-decoration', 'line-height'])
  doc.body.querySelectorAll<HTMLElement>('*').forEach(element => {
    if (element.hasAttribute('style')) {
      const styles = Array.from(element.style).filter(property => allowed.has(property))
        .filter(property => !/url\s*\(|expression|var\s*\(/i.test(element.style.getPropertyValue(property)))
        .map(property => `${property}:${element.style.getPropertyValue(property)}`)
      element.setAttribute('style', styles.join(';'))
    }
    for (const attr of ['href', 'src']) {
      const value = element.getAttribute(attr)
      if (value && !isSafeImportedUrl(value, attr === 'src')) element.removeAttribute(attr)
    }
    if (['P','H1','H2','H3','H4','H5','H6','LI'].includes(element.tagName) && ['color','font-size','font-family','font-weight','font-style','text-decoration'].some(property => element.style.getPropertyValue(property))) {
      const span = doc.createElement('span')
      span.setAttribute('style', Array.from(element.style).filter(property => property !== 'text-align').map(property => `${property}:${element.style.getPropertyValue(property)}`).join(';'))
      while (element.firstChild) span.append(element.firstChild)
      element.append(span)
    }
    if(element.hasAttribute('data-ink-cut'))restoreInkStyle(element)
    if (element.tagName === 'A') element.setAttribute('rel', 'noopener noreferrer')
  })
  return doc.body.innerHTML
}

function canvasText(canvasJson?: string): string {
  if (!canvasJson) return ''
  const collect = (objects: { text?: string; objects?: unknown[] }[]): string => objects.map(object => typeof object.text === 'string' ? object.text : object.objects ? collect(object.objects as typeof objects) : '').filter(Boolean).join('\n')
  try { return collect(JSON.parse(canvasJson).objects ?? []) } catch { return '' }
}
function canvasTextHtml(canvasJson?: string) { return `<pre>${escapeHtml(canvasText(canvasJson))}</pre>` }

function isSafeImportedUrl(value: string, isImageSource: boolean) {
  if (!value) return false
  if (isImageSource && /^data:image\/(?:png|jpe?g|gif|webp);base64,/i.test(value)) return true
  return /^(?:https?:|mailto:|tel:|#)/i.test(value)
}

function dimensionsFor(page: PageData) { return { ...pageSizeFor(page.size), width: page.width ?? pageSizeFor(page.size).width, height: page.height ?? pageSizeFor(page.size).height } }

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
  const size = dimensionsFor(page)
  const margins = normalizePageMargins(page.margins)
  return {
    '--paper-width': `${size.width}px`,
    '--paper-height': `${size.height}px`,
    '--paper-base-height': `${page.baseHeight ?? pageSizeFor(page.size).height}px`,
    '--editor-line-height': page.lineHeight ?? 1.72,
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
