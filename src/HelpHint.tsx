import { cloneElement, useEffect, useId, useRef, useState, type ReactElement } from 'react'
import { createPortal } from 'react-dom'

export function HelpHint({ text, children }: { text: string; children: ReactElement<{ 'aria-describedby'?: string }> }) {
  const id = useId()
  const trigger = useRef<HTMLSpanElement>(null)
  const timer = useRef<number | undefined>(undefined)
  const [position, setPosition] = useState<{ left: number; top: number; direction: 'rtl' | 'ltr' } | null>(null)
  const close = () => { window.clearTimeout(timer.current); setPosition(null) }
  const open = () => {
    window.clearTimeout(timer.current)
    timer.current = window.setTimeout(() => {
      const bounds = trigger.current?.getBoundingClientRect()
      if (bounds) setPosition({ left: Math.max(12, Math.min(bounds.left, window.innerWidth - 312)), top: Math.max(12, Math.min(bounds.bottom + 9, window.innerHeight - 150)), direction: trigger.current?.closest('[dir="rtl"]') ? 'rtl' : 'ltr' })
    }, 650)
  }
  useEffect(() => () => window.clearTimeout(timer.current), [])
  useEffect(() => {
    if (!position) return
    const dismiss = () => setPosition(null)
    window.addEventListener('scroll', dismiss, true)
    return () => window.removeEventListener('scroll', dismiss, true)
  }, [position])
  return <span className="help-target" ref={trigger} onPointerEnter={open} onPointerLeave={close} onFocus={open} onBlur={close} onKeyDown={event => { if (event.key === 'Escape') close() }}>
    {cloneElement(children, { 'aria-describedby': position ? id : undefined })}
    {position && createPortal(<div className="help-popup" id={id} role="tooltip" dir={position.direction} style={{ left: position.left, top: position.top }}>{text}</div>, document.body)}
  </span>
}
