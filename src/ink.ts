import { Mark } from '@tiptap/core'
import type { Editor } from '@tiptap/react'
import { FabricImage, type Canvas as FabricCanvas, type FabricObject } from 'fabric'

export type InkPoint = { x:number; y:number }
export type EraseStroke = { points:InkPoint[]; diameter:number }
type InkCut = { width:number; height:number; strokes:EraseStroke[] }
type Box = { left:number; top:number; width:number; height:number }

export function parseInkCut(value: unknown): InkCut | null {
  if (typeof value !== 'string' || value.length > 300_000) return null
  try {
    const cut = JSON.parse(value) as InkCut
    if (!cut || ![cut.width,cut.height].every(n => Number.isFinite(n) && n > 0 && n < 100_000) || !Array.isArray(cut.strokes) || cut.strokes.length > 200) return null
    if (!cut.strokes.every(stroke => Number.isFinite(stroke.diameter) && stroke.diameter > 0 && stroke.diameter < 100_000 && Array.isArray(stroke.points) && stroke.points.length > 0 && stroke.points.length <= 5000 && stroke.points.every(p => p && Number.isFinite(p.x) && Number.isFinite(p.y) && Math.abs(p.x) < 1e6 && Math.abs(p.y) < 1e6))) return null
    return cut
  } catch { return null }
}
const number = (value:number) => Math.round(value * 1000) / 1000
export function strokePath(stroke:EraseStroke) {
  return stroke.points.map((point,index) => `${index ? 'L' : 'M'}${number(point.x)} ${number(point.y)}`).join(' ') + (stroke.points.length === 1 ? ' l0.001 0' : '')
}
export function inkMaskStyle(cut:InkCut) {
  const paths = cut.strokes.map(stroke => `<path d="${strokePath(stroke)}" fill="none" stroke="black" stroke-width="${number(stroke.diameter)}" stroke-linecap="round" stroke-linejoin="round"/>`).join('')
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${cut.width}" height="${cut.height}" viewBox="0 0 ${cut.width} ${cut.height}"><mask id="cut" maskUnits="userSpaceOnUse" x="0" y="0" width="${cut.width}" height="${cut.height}"><rect width="100%" height="100%" fill="white"/>${paths}</mask><rect width="100%" height="100%" fill="black" mask="url(#cut)"/></svg>`
  return `url("data:image/svg+xml,${encodeURIComponent(svg)}")`
}
export function restoreInkStyle(element:HTMLElement) {
  const cut = parseInkCut(element.getAttribute('data-ink-cut'))
  if (!cut) {element.removeAttribute('data-ink-cut');return}
  element.style.maskImage = inkMaskStyle(cut)
  element.style.maskSize = '100% 100%'
  element.style.maskRepeat = 'no-repeat'
  element.style.display = 'inline-block'
  element.style.height = `${cut.height}px`
  element.style.lineHeight = `${cut.height}px`
}
export const InkCutMark = Mark.create({
  name:'inkCut', inclusive:false,
  addAttributes() { return { cut: { default:null, parseHTML:element => {const value = element.getAttribute('data-ink-cut');return parseInkCut(value) ? value : null}, renderHTML:() => ({}) } } },
  parseHTML() {return [{tag:'span[data-ink-cut]',getAttrs:element => parseInkCut(element.getAttribute('data-ink-cut')) ? {} : false}]},
  renderHTML({ mark }) {
    const cut = parseInkCut(mark.attrs.cut)
    return cut ? ['span',{'data-ink-cut':JSON.stringify(cut),style:`display:inline-block;height:${cut.height}px;line-height:${cut.height}px;mask-image:${inkMaskStyle(cut)};mask-size:100% 100%;mask-repeat:no-repeat;`},0] : ['span',{},0]
  },
})

function pointSegmentDistance(point:InkPoint,a:InkPoint,b:InkPoint) {
  const dx=b.x-a.x,dy=b.y-a.y
  const t=Math.max(0,Math.min(1,((point.x-a.x)*dx+(point.y-a.y)*dy)/(dx*dx+dy*dy || 1)))
  return Math.hypot(point.x-a.x-t*dx,point.y-a.y-t*dy)
}
export function pointErased(point:InkPoint,stroke:EraseStroke) {
  return stroke.points.some((end,index) => pointSegmentDistance(point,stroke.points[Math.max(0,index-1)],end) <= stroke.diameter/2)
}
export function fullyErased(cut:InkCut) {
  for(let y=0;y<=4;y++) for(let x=0;x<=4;x++) if(!cut.strokes.some(stroke => pointErased({x:cut.width*x/4,y:cut.height*y/4},stroke))) return false
  return true
}
function intersects(box:Box,stroke:EraseStroke) {
  const radius=stroke.diameter/2
  const inside = (p:InkPoint) => p.x >= box.left-radius && p.x <= box.left+box.width+radius && p.y >= box.top-radius && p.y <= box.top+box.height+radius
  return stroke.points.some((point,index) => {
    const before=stroke.points[Math.max(0,index-1)]
    const steps=Math.max(1,Math.ceil(Math.hypot(point.x-before.x,point.y-before.y)/Math.max(1,radius)))
    for(let i=0;i<=steps;i++) if(inside({x:before.x+(point.x-before.x)*i/steps,y:before.y+(point.y-before.y)*i/steps})) return true
    return false
  })
}
export function eraseTextTransaction(editor:Editor,stroke:EraseStroke,paper:DOMRect,scale:number) {
  const tr=editor.state.tr
  const type=editor.schema.marks.inkCut
  let visited=0
  editor.state.doc.descendants((node,pos) => {
    if(!node.isText || !node.text) return
    if((visited+=node.text.length)>30_000) throw new Error('renderLimit')
    for(const {segment,index} of new Intl.Segmenter(undefined,{granularity:'grapheme'}).segment(node.text)) {
      const from=pos+index,to=from+segment.length
      const start=editor.view.domAtPos(from),end=editor.view.domAtPos(to)
      const range=document.createRange();range.setStart(start.node,start.offset);range.setEnd(end.node,end.offset)
      const rect=range.getBoundingClientRect()
      if(!rect.width || !rect.height) continue
      const box={left:(rect.left-paper.left)/scale,top:(rect.top-paper.top)/scale,width:rect.width/scale,height:rect.height/scale}
      if(!intersects(box,stroke)) continue
      const previous=parseInkCut(node.marks.find(mark => mark.type===type)?.attrs.cut)
      const sx=previous ? previous.width/box.width : 1,sy=previous ? previous.height/box.height : 1
      const local={diameter:stroke.diameter * Math.max(sx,sy),points:stroke.points.map(point => ({x:(point.x-box.left)*sx,y:(point.y-box.top)*sy}))}
      const cut:InkCut={width:previous?.width ?? box.width,height:previous?.height ?? box.height,strokes:[...(previous?.strokes ?? []),local]}
      if(!parseInkCut(JSON.stringify(cut))) throw new Error('renderLimit')
      tr.addMark(from,to,type.create({cut:JSON.stringify(cut)}))
    }
  })
  return tr
}

export function clearStroke(context:CanvasRenderingContext2D,stroke:EraseStroke,transform:(point:InkPoint)=>InkPoint=(point)=>point,diameter=stroke.diameter) {
  context.save();context.globalCompositeOperation='destination-out';context.lineWidth=diameter;context.strokeStyle='#000';context.fillStyle='#000';context.lineCap='round';context.lineJoin='round'
  context.beginPath()
  stroke.points.forEach((point,index) => {const p=transform(point);if(index)context.lineTo(p.x,p.y);else context.moveTo(p.x,p.y)})
  context.stroke()
  for(const point of [stroke.points[0],stroke.points.at(-1)!]) {const p=transform(point);context.beginPath();context.arc(p.x,p.y,diameter/2,0,Math.PI*2);context.fill()}
  context.restore()
}
async function imageElement(src:string) {
  const image=new Image();image.crossOrigin='anonymous'
  image.src=src;await image.decode();return image
}
export async function eraseImage(src:string,box:Box,rotation:number,stroke:EraseStroke) {
  const angle=rotation*Math.PI/180
  const transform=(point:InkPoint)=>{
    const dx=point.x-box.left-box.width/2,dy=point.y-box.top-box.height/2
    return {x:dx*Math.cos(angle)+dy*Math.sin(angle)+box.width/2,y:-dx*Math.sin(angle)+dy*Math.cos(angle)+box.height/2}
  }
  if(!intersects({left:0,top:0,width:box.width,height:box.height},{...stroke,points:stroke.points.map(transform)}))return src
  const image=await imageElement(src)
  if(image.naturalWidth*image.naturalHeight > 16_000_000) throw new Error('renderLimit')
  const canvas=document.createElement('canvas');canvas.width=image.naturalWidth;canvas.height=image.naturalHeight
  const context=canvas.getContext('2d')!
  context.drawImage(image,0,0)
  // Apply the inverse page transform so a circular eraser stays circular on rotated/scaled pictures.
  context.save();context.scale(canvas.width/box.width,canvas.height/box.height)
  clearStroke(context,stroke,transform)
  context.restore()
  const result=canvas.toDataURL('image/png');canvas.width=canvas.height=0
  return result
}
export function prepareDrawingErase(canvas:FabricCanvas,stroke:EraseStroke) {
  const replacements:{original:FabricObject;image:FabricImage;index:number}[]=[]
  canvas.getObjects().forEach((object,index) => {
    const box=object.getBoundingRect()
    if(!intersects(box,stroke)) return
    if(box.width*box.height > 16_000_000) throw new Error('renderLimit')
    const bitmap=object.toCanvasElement({withoutShadow:true,multiplier:1})
    const context=bitmap.getContext('2d')!
    const before=context.getImageData(0,0,bitmap.width,bitmap.height).data
    clearStroke(context,stroke,point=>({x:point.x-box.left,y:point.y-box.top}))
    const after=context.getImageData(0,0,bitmap.width,bitmap.height).data
    if(!before.some((value,index)=>index%4===3 && value!==after[index]))return
    const image=new FabricImage(bitmap,{left:box.left,top:box.top,originX:'left',originY:'top',perPixelTargetFind:true,objectCaching:false,padding:0})
    replacements.push({original:object,image,index})
  })
  let background:FabricImage|undefined
  if(canvas.backgroundColor && canvas.backgroundColor!=='rgba(255,255,255,0)' && canvas.backgroundColor!=='transparent') {
    const bitmap=document.createElement('canvas');bitmap.width=canvas.width;bitmap.height=canvas.height
    if(bitmap.width*bitmap.height > 16_000_000) throw new Error('renderLimit')
    const context=bitmap.getContext('2d')!;context.fillStyle=typeof canvas.backgroundColor==='string' ? canvas.backgroundColor : '#ffffff';context.fillRect(0,0,bitmap.width,bitmap.height);clearStroke(context,stroke)
    background=new FabricImage(bitmap,{left:0,top:0,originX:'left',originY:'top',perPixelTargetFind:true,padding:0})
  }
  return () => {
    for(const {original,image,index} of replacements) {canvas.remove(original);canvas.insertAt(index,image)}
    if(background) {canvas.backgroundColor='rgba(255,255,255,0)';canvas.insertAt(0,background)}
    canvas.discardActiveObject();canvas.requestRenderAll()
  }
}

export async function eraseInlineImages(editor:Editor,stroke:EraseStroke,paper:DOMRect,scale:number) {
  const images:{pos:number;attrs:Record<string,unknown>;src:string;box:Box}[]=[]
  editor.state.doc.descendants((node,pos)=>{
    if(node.type.name!=='image' || typeof node.attrs.src!=='string')return
    const dom=editor.view.nodeDOM(pos)
    if(!(dom instanceof HTMLElement))return
    const rect=dom.getBoundingClientRect()
    images.push({pos,attrs:node.attrs,src:node.attrs.src,box:{left:(rect.left-paper.left)/scale,top:(rect.top-paper.top)/scale,width:rect.width/scale,height:rect.height/scale}})
  })
  const updates=[]
  for(const image of images) {
    const src=await eraseImage(image.src,image.box,0,stroke)
    if(src!==image.src)updates.push({...image,src})
  }
  return updates
}

export async function maskedGlyphBlob(element:HTMLElement,font:{font?:string;size?:number;bold?:boolean;italics?:boolean;color?:string;underline?:object}) {
  const cut=parseInkCut(element.getAttribute('data-ink-cut'))!
  const container=document.createElement('div');container.setAttribute('xmlns','http://www.w3.org/1999/xhtml')
  container.style.cssText=`font-family:${font.font ?? 'Calibri'};font-size:${(font.size ?? 18)/1.5}px;line-height:${cut.height}px;font-weight:${font.bold?'bold':'normal'};font-style:${font.italics?'italic':'normal'};white-space:pre;`
  if(font.color)container.style.color=`#${font.color}`
  if(font.underline)container.style.textDecoration='underline'
  const glyph=element.cloneNode(true) as HTMLElement;restoreInkStyle(glyph);container.append(glyph)
  const svg=`<svg xmlns="http://www.w3.org/2000/svg" width="${cut.width*2}" height="${cut.height*2}" viewBox="0 0 ${cut.width} ${cut.height}"><foreignObject width="100%" height="100%">${new XMLSerializer().serializeToString(container)}</foreignObject></svg>`
  const image=await imageElement(`data:image/svg+xml,${encodeURIComponent(svg)}`)
  const canvas=document.createElement('canvas');canvas.width=Math.ceil(cut.width*2);canvas.height=Math.ceil(cut.height*2);canvas.getContext('2d')!.drawImage(image,0,0)
  return new Promise<Blob>((resolve,reject)=>canvas.toBlob(blob=>{canvas.width=canvas.height=0;if(blob)resolve(blob);else reject(new Error('couldNotExport'))},'image/png'))
}
