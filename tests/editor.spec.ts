import { test, expect, type Page } from '@playwright/test'
import fs from 'node:fs/promises'
import JSZip from 'jszip'
import { jsPDF } from 'jspdf'
import { Document, Paragraph, Packer } from 'docx'

const project = (pages: object[], activePageId = 'one') => ({ version: 1, language: 'en', activePageId, pages })
const note = (id: string, html = '<p></p>', extra = {}) => ({ id, html, images: [], size: 'a4', ...extra })
async function open(page: Page, data: object) {
  await page.locator('input[accept^=".akxr"]').setInputFiles({ name: 'test.akxr', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(data)) })
  await page.waitForTimeout(200)
}
async function saved(page: Page) {
  const download = page.waitForEvent('download')
  await page.getByRole('button', { name: /^(Save project|Projekt speichern)$/ }).click()
  return JSON.parse(await fs.readFile(await (await download).path() as string, 'utf8'))
}
test.beforeEach(async ({ page }) => {
  page.on('dialog', dialog => dialog.accept())
  await page.goto('/')
  await expect(page.locator('.ProseMirror')).toBeVisible()
})

test('rejects malformed drawings atomically and sanitizes inactive HTML', async ({ page }) => {
  await open(page, project([note('one', '<p>FIRST</p>'), note('two', '<p>SECOND</p><img src="x" onerror="window.pwned=1"><script>window.pwned=2</script>')]))
  await expect(page.locator('.ProseMirror')).toContainText('FIRST')
  expect(await page.evaluate(() => (window as unknown as { pwned?: number }).pwned)).toBeUndefined()
  await open(page, project([note('bad', '<p>BAD</p>', { canvasJson: '{broken' })], 'bad'))
  await expect(page.locator('.ProseMirror')).toContainText('FIRST')
  expect((await saved(page)).pages[1].html).not.toContain('onerror')
})

test('undo is page aware, one shortcut is one action, converter retains history', async ({ page }) => {
  await open(page, project([note('one', '<p>FIRST</p>'), note('two', '<p>SECOND</p>')], 'two'))
  await expect(page.locator('.ProseMirror')).toContainText('SECOND')
  await page.locator('.ProseMirror').press('End')
  await page.locator('.ProseMirror').pressSequentially(' added')
  await expect(page.locator('.ProseMirror')).toContainText('SECOND added')
  await page.locator('.page-tab').filter({ hasText: 'Page 1' }).click()
  await page.getByRole('button', { name: 'Undo', exact: true }).click()
  await page.locator('.page-tab').filter({ hasText: 'Page 2' }).click()
  await expect(page.locator('.ProseMirror')).toHaveText('SECOND')
  await page.getByRole('button', { name: 'Redo', exact: true }).click()
  await expect(page.locator('.ProseMirror')).toHaveText('SECOND added')
  await page.locator('.mode-tabs button').nth(1).click()
  await page.locator('.mode-tabs button').nth(0).click()
  await page.locator('.ProseMirror').press('Control+z')
  await expect(page.locator('.ProseMirror')).toHaveText('SECOND')
})

test('drawings undo and redo update the project and mount only one canvas pair', async ({ page }) => {
  await open(page, project(Array.from({ length: 40 }, (_, i) => note(i ? `p${i}` : 'one'))))
  expect(await page.locator('canvas').count()).toBe(2)
  await page.getByRole('button', { name: 'Rectangle', exact: true }).click()
  await page.waitForTimeout(250)
  expect(JSON.parse((await saved(page)).pages[0].canvasJson).objects.length).toBe(1)
  await page.getByRole('button', { name: 'Undo', exact: true }).click()
  await page.waitForTimeout(250)
  expect(JSON.parse((await saved(page)).pages[0].canvasJson || '{"objects":[]}').objects.length).toBe(0)
  await page.getByRole('button', { name: 'Redo', exact: true }).click()
  await page.waitForTimeout(250)
  expect(JSON.parse((await saved(page)).pages[0].canvasJson).objects.length).toBe(1)
  expect(await page.locator('.page-strip').evaluate(element => element.clientHeight)).toBeLessThan(100)
})

test('reopening uses active page, text export preserves paragraphs and canvas text', async ({ page }) => {
  const data = project([note('one', '<p>FIRST</p>'), note('two', '<p>Line one</p><p>Line two</p>', { canvasJson: JSON.stringify({ objects: [{ type: 'IText', text: 'Canvas note', left: 100, top: 100, fontSize: 20 }] }) })], 'two')
  await open(page, data)
  await open(page, data)
  await expect(page.locator('.ProseMirror')).toContainText('Line two')
  const download = page.waitForEvent('download')
  await page.getByLabel('Export format', { exact: true }).selectOption('txt')
  const text = await fs.readFile(await (await download).path() as string, 'utf8')
  expect(text).toContain('Line one\nLine two')
  expect(text).toContain('Canvas note')
})

test('local autosave persists notes and rejects concurrent tab writes', async ({ page, context }) => {
  await open(page, project([note('one', '<p>Saved</p>')]))
  await expect(page.getByRole('status').first()).toContainText('Saved locally')
  const other = await context.newPage()
  other.on('dialog', dialog => dialog.accept())
  await other.goto('/')
  await expect(other.locator('.ProseMirror')).toContainText('Saved')
  await other.locator('.ProseMirror').press('End')
  await other.locator('.ProseMirror').pressSequentially(' other')
  await expect(other.getByRole('status').first()).toContainText('Saved locally')
  await page.locator('.ProseMirror').press('End')
  await page.locator('.ProseMirror').pressSequentially(' mine')
  await expect(page.locator('.save-indicator')).toContainText('Another tab')
  await other.reload()
  await expect(other.locator('.ProseMirror')).toContainText('Saved other')
})

test('editable DOCX and physically sized PDF exports', async ({ page }) => {
  await open(page, project([note('one', '<p>Editable text</p><table><tbody><tr><td><p>Cell</p></td><td><p>Other</p></td></tr></tbody></table>', {
    canvasJson: JSON.stringify({ background:'rgba(255,255,255,0)', objects:[{type:'IText',text:'Editable canvas note',left:100,top:200,fontSize:18}] }),
  })]))
  let download = page.waitForEvent('download')
  await page.getByLabel('Export format', { exact: true }).selectOption('docx')
  const zip = await JSZip.loadAsync(await fs.readFile(await (await download).path() as string))
  const xml = await zip.file('word/document.xml')!.async('text')
  expect(xml).toContain('Editable text')
  expect(xml).toContain('<w:tbl>')
  expect(xml.match(/Editable canvas note/g)?.length).toBe(1)
  expect(Object.keys(zip.files).filter(name => name.startsWith('word/media/') && !zip.files[name].dir)).toHaveLength(0)
  download = page.waitForEvent('download')
  await page.getByLabel('Export format', { exact: true }).selectOption('pdf')
  const pdf = await fs.readFile(await (await download).path() as string, 'latin1')
  const box = pdf.match(/\/MediaBox \[0 0 ([\d.]+) ([\d.]+)\]/)
  expect(box).toBeTruthy()
  expect(Number(box![1]) / 72 * 25.4).toBeCloseTo(210, 0)
  expect(Number(box![2]) / 72 * 25.4).toBeCloseTo(297, 0)
})

test('color affects only selected text and toolbar follows selection', async ({ page }) => {
  await open(page, project([note('one', '<p>First second</p>')]))
  const text = page.locator('.ProseMirror')
  await text.click()
  await page.waitForTimeout(60)
  await page.keyboard.press('Control+Home')
  await page.keyboard.press('Control+Shift+ArrowRight')
  await expect.poll(() => page.evaluate(() => window.getSelection()?.toString())).toBe('First ')
  await page.locator('input[type=color]').first().fill('#ff0000')
  await expect(page.locator('.ProseMirror span').first()).toHaveCSS('color', 'rgb(255, 0, 0)')
  const result = await saved(page)
  expect(result.pages[0].html).toMatch(/color: rgb\(255, 0, 0\)[^>]*>First/)
  expect(result.pages[0].html).toMatch(/<\/span>second/)
  await text.click()
  await page.waitForTimeout(60)
  await page.keyboard.press('Control+Home')
  await page.getByRole('button', { name: 'Bold', exact: true }).click()
  await expect(page.getByRole('button', { name: 'Bold', exact: true })).toHaveAttribute('aria-pressed', 'true')
  await text.click()
  await page.waitForTimeout(60)
  await page.keyboard.press('End')
  await expect(page.getByRole('button', { name: 'Bold', exact: true })).toHaveAttribute('aria-pressed', 'false')
})

test('long notes grow the sheet and page library restores deleted pages', async ({ page }) => {
  await open(page, project([note('one', '<p>Keep</p>'), note('two', '<p>Deleted</p>')]))
  await page.getByRole('button', { name: 'Delete page 2', exact: true }).click()
  await page.locator('.document-library summary').click()
  await page.getByRole('tab',{name:'Bring back earlier work',exact:true}).click()
  await page.getByRole('button', { name: 'Bring back page', exact: true }).click()
  expect(await page.locator('.page-tab-group').count()).toBe(2)
  await page.locator('.document-library summary').click()
  await page.locator('.ProseMirror').fill(Array.from({ length: 140 }, (_, i) => `Long note ${i}`).join('\n'))
  await expect.poll(() => page.locator('.paper').evaluate(element => element.clientHeight)).toBeGreaterThan(2000)
  const sizes = await page.locator('.paper').evaluate(element => ({ paper: element.clientHeight, text: element.querySelector('.ProseMirror')!.scrollHeight }))
  expect(sizes.paper).toBeGreaterThanOrEqual(sizes.text)
})

test('storage failures are visible and protect unsaved work', async ({ page }) => {
  await page.evaluate(() => {
    const original = IDBObjectStore.prototype.put
    IDBObjectStore.prototype.put = function (...args: Parameters<IDBObjectStore['put']>) {
      if (this.name === 'drafts') throw new DOMException('Storage full', 'QuotaExceededError')
      return original.apply(this, args)
    }
  })
  await page.locator('.ProseMirror').fill('Do not lose this')
  await expect(page.locator('.save-indicator')).toContainText('fehlgeschlagen')
  await page.getByLabel('Sprache', { exact: true }).selectOption('en')
  expect((await saved(page)).pages[0].html).toContain('Do not lose this')
})

test('PDF imports preserve landscape dimensions and DOCX imports preserve section margins', async ({ page }) => {
  const pdf = new jsPDF({ unit: 'mm', format: [300,200], orientation: 'landscape' })
  pdf.text('PDF page', 20, 20)
  await page.locator('input[accept^=".akxr"]').setInputFiles({ name: 'wide.pdf', mimeType: 'application/pdf', buffer: Buffer.from(pdf.output('arraybuffer')) })
  await expect(page.locator('.paper')).toHaveCSS('width', '1134px')
  expect((await saved(page)).pages[0].height).toBeCloseTo(756, 0)
  const document = new Document({ sections: [
    { properties: { page: { size: { width:11906,height:16838 }, margin: {top:0,bottom:0,left:0,right:0} } }, children: [new Paragraph('Section one')] },
    { properties: { page: { size: { width:11906,height:8391 }, margin: {top:300,bottom:300,left:300,right:300} } }, children: [new Paragraph('Section two')] },
  ] })
  await page.locator('input[accept^=".akxr"]').setInputFiles({ name: 'sections.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await Packer.toBuffer(document) })
  await expect(page.locator('.ProseMirror')).toContainText('Section one')
  const result = await saved(page)
  expect(result.pages.length).toBe(2)
  expect(result.pages[0].margins.left).toBe(0)
  expect(result.pages[1].size).toBe('a5Landscape')
  expect(result.pages[1].margins.left).toBe(20)
})

test('HTML retains inline style, links and pictures; real table supports rows', async ({ page }) => {
  const html = '<p style="color:#ff0000"><strong>Bold</strong> <a href="https://example.com">Link</a></p><img alt="Pixel" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9Z0sAAAAAASUVORK5CYII=">'
  await page.locator('input[accept^=".akxr"]').setInputFiles({ name: 'test.html', mimeType: 'text/html', buffer: Buffer.from(html) })
  await expect(page.locator('.ProseMirror strong')).toContainText('Bold')
  await expect(page.locator('.ProseMirror a')).toHaveAttribute('href','https://example.com')
  await expect(page.locator('.ProseMirror img')).toHaveAttribute('alt','Pixel')
  await expect(page.locator('.ProseMirror span').first()).toHaveCSS('color','rgb(255, 0, 0)')
  await page.getByLabel(/^(Language|Sprache)$/).selectOption('en')
  await page.getByRole('button', { name: 'Insert table', exact: true }).click()
  expect(await page.locator('.ProseMirror tr').count()).toBe(3)
  await page.getByRole('button', { name: 'Add row', exact: true }).click()
  expect(await page.locator('.ProseMirror tr').count()).toBe(4)
})

test('touch placement, pressure stroke and palm rejection', async ({ page }) => {
  await open(page, project([note('one')]))
  await page.getByRole('button', { name: 'Place text', exact: true }).click()
  const canvas = page.locator('canvas.upper-canvas')
  const rect = (await canvas.boundingBox())!
  const pointer = { pointerId: 1, pointerType: 'touch', isPrimary: true, clientX:rect.x+120,clientY:rect.y+130,buttons:1,button:0 }
  await canvas.dispatchEvent('pointerdown',pointer)
  await canvas.dispatchEvent('pointerup',{...pointer,buttons:0})
  await expect(page.locator('textarea[data-fabric]')).toBeAttached()
  await page.keyboard.type('Touch note')
  await page.keyboard.press('Escape')
  await page.getByRole('button', { name: 'Brush', exact: true }).click()
  await page.getByLabel('Pen only / palm rejection', { exact:true }).check()
  await canvas.dispatchEvent('pointerdown',pointer)
  await canvas.dispatchEvent('pointerup',{...pointer,buttons:0})
  const pen = {...pointer,pointerId:2,pointerType:'pen',pressure:0.2,clientY:rect.y+230}
  await canvas.dispatchEvent('pointerdown',pen)
  await canvas.dispatchEvent('pointermove',{...pen,clientX:rect.x+200,pressure:0.8})
  await canvas.dispatchEvent('pointerup',{...pen,clientX:rect.x+220,pressure:0,buttons:0})
  await page.waitForTimeout(200)
  const drawing = JSON.parse((await saved(page)).pages[0].canvasJson)
  expect(drawing.objects.length).toBe(2)
  expect(drawing.objects.some((object: { type:string }) => object.type === 'Path')).toBe(true)
})

test('production app reloads offline with local notes and export libraries', async ({ page, context }) => {
  await page.goto('http://127.0.0.1:4173')
  await expect(page.locator('.ProseMirror')).toBeVisible()
  await page.evaluate(async () => { await navigator.serviceWorker.ready })
  // First navigation may predate SW control; reload online to adopt the installed worker.
  await page.reload()
  await page.locator('.ProseMirror').fill('Offline note')
  await expect(page.locator('.save-indicator')).toContainText('Lokal gespeichert')
  await context.setOffline(true)
  await page.reload()
  await expect(page.locator('.ProseMirror')).toContainText('Offline note')
  await page.getByLabel('Sprache', { exact: true }).selectOption('en')
  const download = page.waitForEvent('download')
  await page.getByLabel('Export format', { exact: true }).selectOption('docx')
  const zip = await JSZip.loadAsync(await fs.readFile(await (await download).path() as string))
  expect(await zip.file('word/document.xml')!.async('text')).toContain('Offline note')
  await context.setOffline(false)
})

test('paper remains reachable on desktop and tablet; image selection wins over transparent canvas', async ({ page }) => {
  const pixel = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9Z0sAAAAAASUVORK5CYII='
  await open(page, project([note('one', '<p>Visible text</p>', { images:[{id:'picture',src:pixel,x:80,y:100,width:100,height:100,rotation:0,alt:'Picture'}] })]))
  for (const width of [1440,768]) {
    await page.setViewportSize({width,height:900})
    const geometry = await page.locator('.paper').evaluate(element => ({ left:element.getBoundingClientRect().left, stage:element.parentElement!.getBoundingClientRect().left, scroll:element.parentElement!.scrollLeft }))
    expect(geometry.left).toBeGreaterThanOrEqual(geometry.stage - 1)
  }
  await page.setViewportSize({width:1440,height:900})
  await page.getByRole('button', { name:'Select paint',exact:true }).click()
  await page.locator('.floating-image').click()
  await expect(page.locator('.floating-image')).toHaveClass(/active/)
})

test('metadata, canvas search, attachments and persistent versions', async ({ page }) => {
  await open(page, project([note('one','<p>Page body</p>', {canvasJson:JSON.stringify({objects:[{type:'IText',text:'Searchable canvas',left:100,top:100,fontSize:20}]})})]))
  await page.locator('.document-library summary').click()
  await page.getByRole('tab',{name:'Rename and organize',exact:true}).click()
  await page.getByLabel('Page title', {exact:true}).fill('Version A')
  await page.getByLabel('Section', {exact:true}).fill('Research')
  await page.getByLabel('Tags / metadata', {exact:true}).fill('local')
  await page.getByRole('tab',{name:'Files on this page',exact:true}).click()
  await page.locator('.document-library input[type=file]').setInputFiles({name:'attachment.txt',mimeType:'text/plain',buffer:Buffer.from('Attachment content')})
  await expect(page.getByRole('button',{name:'attachment.txt',exact:true})).toBeVisible()
  const attachment = page.waitForEvent('download')
  await page.getByRole('button',{name:'attachment.txt',exact:true}).click()
  expect(await fs.readFile(await (await attachment).path() as string,'utf8')).toBe('Attachment content')
  await page.getByRole('tab',{name:'Find a page',exact:true}).click()
  await page.getByLabel('Search all pages',{exact:true}).fill('Searchable canvas')
  await expect(page.getByRole('button',{name:'Version A Research · local',exact:true})).toBeVisible()
  await page.getByRole('tab',{name:'Bring back earlier work',exact:true}).click()
  await page.getByRole('button',{name:'Show previous versions',exact:true}).click()
  await expect(page.locator('.versions-list button').first()).toBeVisible()
  await page.getByRole('tab',{name:'Rename and organize',exact:true}).click()
  await page.getByLabel('Page title',{exact:true}).fill('Version B')
  await page.getByRole('tab',{name:'Bring back earlier work',exact:true}).click()
  await page.locator('.versions-list button').first().click()
  await expect(page.locator('.library-notice')).toContainText('Project returned to')
  await page.getByRole('tab',{name:'Rename and organize',exact:true}).click()
  await expect(page.getByLabel('Page title',{exact:true})).toHaveValue('Version A')
  await page.getByRole('button',{name:'Undo',exact:true}).click()
  await expect(page.getByLabel('Page title',{exact:true})).toHaveValue('Version B')
  await expect(page.locator('.save-indicator')).toContainText('Saved locally')
  await page.reload()
  await page.locator('.document-library summary').click()
  await page.getByRole('tab',{name:'Rename and organize',exact:true}).click()
  await expect(page.getByLabel('Page title',{exact:true})).toHaveValue('Version B')
  await page.getByRole('tab',{name:'Files on this page',exact:true}).click()
  await expect(page.getByRole('button',{name:'attachment.txt',exact:true})).toBeVisible()
  await page.getByRole('tab',{name:'Bring back earlier work',exact:true}).click()
  await page.getByRole('button',{name:'Show previous versions',exact:true}).click()
  await expect(page.locator('.versions-list button').first()).toBeVisible()
})

test('undo across drawing pages changes the action page and persists redo', async ({ page }) => {
  await open(page, project([note('one'),note('two')]))
  await page.getByRole('button',{name:'Rectangle',exact:true}).click()
  await page.locator('.page-tab').filter({hasText:'Page 2'}).click()
  await page.getByRole('button',{name:'Circle',exact:true}).click()
  await page.locator('.page-tab').filter({hasText:'Page 1'}).click()
  await page.getByRole('button',{name:'Undo',exact:true}).click()
  await expect(page.locator('.paper')).toHaveAttribute('id','page-two')
  const result = await saved(page)
  expect(JSON.parse(result.pages[0].canvasJson).objects.length).toBe(1)
  expect(JSON.parse(result.pages[1].canvasJson || '{"objects":[]}').objects.length).toBe(0)
  await page.getByRole('button',{name:'Redo',exact:true}).click()
  await page.waitForTimeout(150)
  expect(JSON.parse((await saved(page)).pages[1].canvasJson).objects.length).toBe(1)
})

test('a delayed import cannot overwrite a new project', async ({ page }) => {
  await page.evaluate(() => {
    const read = File.prototype.text
    File.prototype.text = async function() {
      if (this.name === 'slow.akxr') await new Promise(resolve => setTimeout(resolve,800))
      return read.call(this)
    }
  })
  await page.locator('input[accept^=".akxr"]').setInputFiles({name:'slow.akxr',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(project([note('one','<p>Outdated import</p>')])) )})
  await page.getByRole('button',{name:'Neues Projekt',exact:true}).click()
  await page.locator('.ProseMirror').fill('New work')
  await page.waitForTimeout(1000)
  await expect(page.locator('.ProseMirror')).toHaveText('New work')
})
