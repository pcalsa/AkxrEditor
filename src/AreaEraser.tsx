import { useLayoutEffect, useRef, useState } from 'react'
import { inkMaskStyle, type EraseStroke, type InkPoint } from './ink'

export function AreaEraser({width,height,diameter,penOnly,label,onErase}:{width:number;height:number;diameter:number;penOnly:boolean;label:string;onErase:(stroke:EraseStroke)=>Promise<void>}) {
  const drawing=useRef<{id:number;points:InkPoint[]}|null>(null)
  const overlay=useRef<SVGSVGElement>(null)
  const [points,setPoints]=useState<InkPoint[]>([])
  const [cursor,setCursor]=useState<InkPoint|null>(null)
  const [busy,setBusy]=useState(false)
  useLayoutEffect(()=>{
    if(!points.length)return
    const layers=Array.from(overlay.current?.parentElement?.children ?? []).filter((element):element is HTMLElement=>element instanceof HTMLElement)
    const previous=layers.map(layer=>({layer,image:layer.style.maskImage,size:layer.style.maskSize,repeat:layer.style.maskRepeat}))
    const mask=inkMaskStyle({width,height,strokes:[{diameter,points}]})
    for(const layer of layers){layer.style.maskImage=mask;layer.style.maskSize='100% 100%';layer.style.maskRepeat='no-repeat'}
    return ()=>{for(const {layer,image,size,repeat} of previous){layer.style.maskImage=image;layer.style.maskSize=size;layer.style.maskRepeat=repeat}}
  },[points,diameter,width,height])
  const point=(event:React.PointerEvent<SVGSVGElement>) => {const rect=event.currentTarget.getBoundingClientRect();return {x:(event.clientX-rect.left)*width/rect.width,y:(event.clientY-rect.top)*height/rect.height}}
  return <svg ref={overlay} className="area-eraser" role="img" aria-label={label} viewBox={`0 0 ${width} ${height}`} style={{width,height,cursor:busy?'wait':'none'}} onPointerDown={event=>{
    event.preventDefault();event.stopPropagation()
    if(busy || drawing.current || event.button!==0 || penOnly && event.pointerType==='touch')return
    const p=point(event);drawing.current={id:event.pointerId,points:[p]};setPoints([p]);setCursor(p);event.currentTarget.setPointerCapture(event.pointerId)
  }} onPointerMove={event=>{
    if(penOnly && event.pointerType==='touch')return
    const p=point(event);setCursor(p)
    if(drawing.current?.id!==event.pointerId)return
    const list=drawing.current.points
    if(Math.hypot(p.x-list.at(-1)!.x,p.y-list.at(-1)!.y)<0.7)return
    if(list.length < 5000)list.push(p)
    setPoints([...list])
  }} onPointerUp={event=>{
    if(drawing.current?.id!==event.pointerId)return
    const final=point(event)
    if(drawing.current.points.length<5000)drawing.current.points.push(final)
    const stroke={diameter,points:drawing.current.points};drawing.current=null
    if(event.currentTarget.hasPointerCapture(event.pointerId))event.currentTarget.releasePointerCapture(event.pointerId)
    setBusy(true)
    void onErase(stroke).finally(()=>{setBusy(false);setPoints([])})
  }} onPointerCancel={event=>{if(drawing.current?.id===event.pointerId){drawing.current=null;setPoints([])}}} onPointerLeave={()=>{if(!drawing.current)setCursor(null)}}>
    {cursor && <circle cx={cursor.x} cy={cursor.y} r={diameter/2} fill="rgba(137,111,255,0.12)" stroke="#7857df" strokeWidth="1" vectorEffect="non-scaling-stroke"/>}
  </svg>
}
