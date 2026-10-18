import { test, expect, type Page } from '@playwright/test'
import fs from 'node:fs/promises'
import { catalogs } from '../src/i18n'

const note=(html='<p></p>',extra={})=>({id:'one',html,images:[],size:'a4',...extra})
async function open(page:Page,data:object) {
  await page.locator('input[accept^=".akxr"]').setInputFiles({name:'update.akxr',mimeType:'application/json',buffer:Buffer.from(JSON.stringify({version:1,language:'en',activePageId:'one',pages:[data]}))})
  await expect(page.getByLabel('Language',{exact:true})).toHaveValue('en')
}
async function saved(page:Page) {
  const download=page.waitForEvent('download');await page.getByRole('button',{name:'Save project',exact:true}).click()
  return JSON.parse(await fs.readFile(await (await download).path() as string,'utf8'))
}
async function at(page:Page,x:number,y:number) {
  const rect=(await page.locator('.paper').boundingBox())!
  return {x:rect.x+x*rect.width/794,y:rect.y+y*rect.width/794}
}
async function erase(page:Page,x:number,y:number,endX=x,endY=y) {
  const start=await at(page,x,y),end=await at(page,endX,endY)
  await page.mouse.move(start.x,start.y);await page.mouse.down();await page.mouse.move(end.x,end.y,{steps:5});await page.mouse.up()
  await expect(page.locator('.editor-status')).toHaveText('Area erased. Undo can restore it.')
  await page.mouse.move(10,10)
}
async function rgba(page:Page,src:string,x:number,y:number) {
  return page.evaluate(async ({src,x,y})=>{const image=new Image();image.src=src;await image.decode();const canvas=document.createElement('canvas');canvas.width=image.naturalWidth;canvas.height=image.naturalHeight;const context=canvas.getContext('2d')!;context.drawImage(image,0,0);return [...context.getImageData(x,y,1,1).data]},{src,x,y})
}
test.beforeEach(async ({page})=>{
  page.on('dialog',dialog=>dialog.accept());await page.goto('/');await expect(page.locator('.ProseMirror')).toBeVisible()
})

test('all four language catalogs are complete and switch live, including recovery and hints',async({page})=>{
  for(const entries of Object.values(catalogs)) {
    expect(Object.keys(entries).sort()).toEqual(Object.keys(catalogs.en).sort())
    for(const value of Object.values(entries))expect(value).not.toMatch(/[ØÙÃ]/)
  }
  await open(page,note())
  await page.locator('.document-library summary').click()
  for(const language of ['de','ar','da','en'] as const) {
    await page.locator('.editor-language-picker select').selectOption(language)
    await expect(page.locator('.document-library summary strong')).toHaveText(catalogs[language].pageManagement)
    await page.getByRole('tab',{name:catalogs[language].recoverWork,exact:true}).click()
    await expect(page.getByRole('button',{name:catalogs[language].versions,exact:true})).toBeVisible()
    await page.getByRole('button',{name:catalogs[language].pixelEraser,exact:true}).hover()
    await expect(page.getByRole('tooltip')).toHaveText(catalogs[language].pixelEraserHint)
    await page.mouse.move(10,10)
    await expect(page.locator('.app-shell')).toHaveAttribute('lang',language)
  }
  await page.locator('.editor-language-picker select').selectOption('da')
  await expect(page.locator('.save-indicator')).toHaveText('Gemt lokalt')
  await page.reload()
  await expect(page.locator('.editor-language-picker select')).toHaveValue('da')
})

test('area eraser clears actual PDF/image pixels and undo restores them',async({page})=>{
  const red=await page.evaluate(()=>{const canvas=document.createElement('canvas');canvas.width=794;canvas.height=1123;const ctx=canvas.getContext('2d')!;ctx.fillStyle='#ff0000';ctx.fillRect(0,0,794,1123);return canvas.toDataURL()})
  await open(page,note('<p></p>',{pageImage:red}))
  await page.getByRole('button',{name:'Area eraser',exact:true}).click()
  await erase(page,130,150,180,150)
  let document=await saved(page)
  expect((await rgba(page,document.pages[0].pageImage,150,150))[3]).toBe(0)
  expect(await rgba(page,document.pages[0].pageImage,80,150)).toEqual([255,0,0,255])
  await page.getByRole('button',{name:'Undo',exact:true}).click()
  expect((await saved(page)).pages[0].pageImage).toBe(red)
  await page.getByRole('button',{name:'Redo',exact:true}).click()
  await expect(page.locator('.save-indicator')).toHaveText('Saved locally')
  await page.reload();await expect(page.locator('.ProseMirror')).toBeVisible()
  document=await saved(page)
  expect((await rgba(page,document.pages[0].pageImage,150,150))[3]).toBe(0)
})

test('area eraser removes only part of a drawing, leaving other ink selectable',async({page})=>{
  await open(page,note('<p></p>',{canvasJson:JSON.stringify({objects:[{type:'Line',x1:-70,y1:0,x2:70,y2:0,left:170,top:180,originX:'center',originY:'center',stroke:'#222222',strokeWidth:8}]})}))
  await page.getByRole('button',{name:'Area eraser',exact:true}).click()
  await erase(page,170,180)
  const document=await saved(page)
  const objects=JSON.parse(document.pages[0].canvasJson).objects
  expect(objects).toHaveLength(1)
  expect(objects[0].type).toBe('Image')
  const object=objects[0]
  expect((await rgba(page,object.src,Math.floor(170-object.left),Math.floor(180-object.top)))[3]).toBe(0)
  expect((await rgba(page,object.src,Math.floor(120-object.left),Math.floor(180-object.top)))[3]).toBeGreaterThan(0)
  await page.getByRole('button',{name:'Undo',exact:true}).click()
  expect(JSON.parse((await saved(page)).pages[0].canvasJson).objects[0].type).toBe('Line')
})

test('area eraser cuts text without moving the remaining text and survives reload/export',async({page})=>{
  await open(page,note('<p><span style="font-size:48px">MMMMM other text</span></p>'))
  await expect(page.locator('.ProseMirror')).toContainText('MMMMM')
  const box=await page.locator('.ProseMirror').evaluate(element=>{const walker=document.createTreeWalker(element,NodeFilter.SHOW_TEXT);const text=walker.nextNode()!;const range=document.createRange();range.setStart(text,0);range.setEnd(text,1);const rect=range.getBoundingClientRect();return{x:rect.x,y:rect.y,width:rect.width,height:rect.height}})
  const before=await page.locator('.ProseMirror').innerText()
  await page.getByRole('button',{name:'Area eraser',exact:true}).click()
  const beforePixels=await page.locator('.paper').screenshot()
  await page.mouse.move(box.x+box.width/2,box.y+box.height/2);await page.mouse.down();await page.mouse.up()
  await expect(page.locator('.editor-status')).toHaveText('Area erased. Undo can restore it.')
  await expect(page.locator('[data-ink-cut]').first()).toHaveCSS('mask-image',/data:image\/svg/)
  await page.mouse.move(10,10)
  const afterPixels=await page.locator('.paper').screenshot()
  const paper=(await page.locator('.paper').boundingBox())!
  const dark=async(buffer:Buffer)=>page.evaluate(async({data,box,paper})=>{
    const image=new Image();image.src=data;await image.decode();const canvas=document.createElement('canvas');canvas.width=image.naturalWidth;canvas.height=image.naturalHeight;const ctx=canvas.getContext('2d')!;ctx.drawImage(image,0,0)
    const pixels=ctx.getImageData(Math.floor(box.x-paper.x),Math.floor(box.y-paper.y),Math.ceil(box.width),Math.ceil(box.height)).data
    let count=0;for(let i=0;i<pixels.length;i+=4)if(pixels[i]+pixels[i+1]+pixels[i+2]<500)count++
    return count
  },{data:`data:image/png;base64,${buffer.toString('base64')}`,box,paper})
  expect(await dark(afterPixels)).toBeLessThan(await dark(beforePixels))
  await expect(page.locator('.ProseMirror')).toHaveText(before)
  const snapshot=await saved(page)
  expect(snapshot.pages[0].html).toContain('data-ink-cut')
  await page.getByRole('button',{name:'Undo',exact:true}).click()
  await expect(page.locator('[data-ink-cut]')).toHaveCount(0)
  await page.getByRole('button',{name:'Redo',exact:true}).click()
  await expect(page.locator('.save-indicator')).toHaveText('Saved locally')
  await page.reload();await expect(page.locator('[data-ink-cut]').first()).toHaveCSS('mask-image',/data:image\/svg/)
  const download=page.waitForEvent('download');await page.getByLabel('Export format',{exact:true}).selectOption('html')
  expect(await fs.readFile(await (await download).path() as string,'utf8')).toContain('mask-image')
})

test('tight stroke hit testing and object eraser never starts a selection rectangle',async({page})=>{
  await open(page,note('<p></p>',{canvasJson:JSON.stringify({objects:[{type:'Line',x1:-60,y1:-60,x2:60,y2:60,left:180,top:180,originX:'center',originY:'center',stroke:'#000000',strokeWidth:4}]})}))
  await page.getByRole('button',{name:'Select paint',exact:true}).click()
  let p=await at(page,130,220);await page.mouse.click(p.x,p.y)
  const state=()=>page.locator('.ProseMirror').evaluate(element=>{const paper=element.closest('.paper')!;const upper=paper.querySelector('canvas.upper-canvas') as HTMLCanvasElement;const context=upper.getContext('2d')!;return context.getImageData(0,0,upper.width,upper.height).data.some((value,index)=>index%4===3 && value>0)})
  expect(await state()).toBe(false)
  p=await at(page,180,180);await page.mouse.click(p.x,p.y)
  const selected=()=>page.locator('canvas.lower-canvas').evaluate(canvas=>{const element=canvas as HTMLCanvasElement;const data=element.getContext('2d')!.getImageData(0,0,element.width,element.height).data;for(let i=0;i<data.length;i+=4)if(data[i]>100&&data[i]<150&&data[i+1]>65&&data[i+1]<120&&data[i+2]>220&&data[i+3]>0)return true;return false})
  await expect.poll(selected).toBe(true)
  await page.getByRole('button',{name:'Object eraser',exact:true}).click()
  p=await at(page,90,100);await page.mouse.move(p.x,p.y);await page.mouse.down();p=await at(page,100,110);await page.mouse.move(p.x,p.y)
  expect(await state()).toBe(false)
  await page.mouse.up()
})

test('zoom changes only the page view and erases at the right coordinates when zoomed',async({page})=>{
  const background=await page.evaluate(()=>{const canvas=document.createElement('canvas');canvas.width=794;canvas.height=1123;const context=canvas.getContext('2d')!;context.fillStyle='#00ffff';context.fillRect(0,0,794,1123);return canvas.toDataURL()})
  await open(page,note('<p>Zoom test</p>',{pageImage:background}))
  const before=(await page.locator('.paper').boundingBox())!
  const rail=(await page.locator('.editor-rail-left').boundingBox())!
  await page.getByRole('button',{name:'Zoom out on page',exact:true}).click()
  const after=(await page.locator('.paper').boundingBox())!
  expect(after.width/before.width).toBeCloseTo(.9,2)
  expect((await page.locator('.editor-rail-left').boundingBox())!.width).toBeCloseTo(rail.width)
  expect((await saved(page)).pages[0].width).toBeUndefined()
  await page.getByRole('button',{name:'Area eraser',exact:true}).click()
  await erase(page,130,200)
  const snapshot=await saved(page)
  expect((await rgba(page,snapshot.pages[0].pageImage,130,200))[3]).toBe(0)
  expect((await rgba(page,snapshot.pages[0].pageImage,100,200))[3]).toBe(255)
  await page.getByRole('button',{name:'Reset page view to 100%',exact:true}).click()
  expect((await page.locator('.paper').boundingBox())!.width).toBeCloseTo(before.width)
})
