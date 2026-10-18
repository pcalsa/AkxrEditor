import { useState } from 'react'
import { Search, PencilLine, Paperclip, History, Trash2, Copy, Download, RotateCcw, CheckCircle2, ChevronDown, FileText, FolderOpen } from 'lucide-react'
import { HelpHint } from './HelpHint'
import type { AppLanguage, Translate, TranslationKey } from './i18n'
import type { LocalVersion } from './storage'
import type { FloatingImage, PageData, SavedProject } from './App'

type Notice = { key: TranslationKey; values?: Record<string, string | number>; error?: boolean; date?:number }
export function PageLibrary({ page, pages, trash, query, onQuery, results, onOpen, onPatch, onDuplicate, onAttach, onDownload, onRestorePage, onLoadVersions, onRestoreVersion, language, t }: {
  page: PageData; pages: PageData[]; trash: PageData[]; query: string; onQuery: (value: string) => void
  results: PageData[]; onOpen: (id: string) => void; onPatch: (patch: Partial<PageData>) => void
  onDuplicate: () => void; onAttach: (files: File[]) => Promise<void>; onDownload: (id: string) => Promise<void>
  onRestorePage: (id: string) => void; onLoadVersions: () => Promise<LocalVersion<SavedProject>[]>
  onRestoreVersion: (version: LocalVersion<SavedProject>) => Promise<void>; language: AppLanguage; t: Translate
}) {
  const [tab, setTab] = useState<'find' | 'edit' | 'files' | 'recover'>('find')
  const [versions, setVersions] = useState<LocalVersion<SavedProject>[] | null>(null)
  const [loading, setLoading] = useState(false)
  const [notice, setNotice] = useState<Notice | null>(null)
  const title = (item: PageData) => item.title || t('page', { number: Math.max(1, pages.findIndex(page => page.id === item.id) + 1) })
  const readableDate = (date: string | number | undefined) => date && Number.isFinite(new Date(date).getTime()) ? new Date(date).toLocaleString(language) : '—'
  const updateAlt = (image: FloatingImage, alt: string) => onPatch({ images: page.images.map(item => item.id === image.id ? { ...item, alt } : item) })
  return <details className="document-library">
    <summary><span className="library-summary-icon"><FolderOpen size={19} /></span><span><strong>{t('pageManagement')}</strong><small>{t('libraryIntro')}</small></span><ChevronDown className="library-chevron" size={18} /></summary>
    <div className="library-body">
      <div className="library-tabs" role="tablist" aria-label={t('pageManagement')}>
        {([{ id:'find', key:'findPage', icon:Search },{ id:'edit', key:'editPage', icon:PencilLine },{ id:'files', key:'pageFiles', icon:Paperclip },{ id:'recover', key:'recoverWork', icon:History }] as const).map(({ id,key,icon:Icon }) => <button key={id} type="button" role="tab" aria-selected={tab === id} aria-controls={`library-${id}`} id={`tab-${id}`} tabIndex={tab === id ? 0 : -1} onClick={() => { setTab(id); setNotice(null) }} onKeyDown={event => {
          if (!['ArrowRight','ArrowLeft','Home','End'].includes(event.key)) return
          event.preventDefault()
          const tabs = ['find','edit','files','recover'] as const
          const index = tabs.indexOf(tab)
          const direction=language==='ar' ? -1 : 1
          const next = event.key === 'Home' ? 0 : event.key === 'End' ? 3 : (index + (event.key === 'ArrowRight' ? direction : -direction) + 4) % 4
          setTab(tabs[next]); document.getElementById(`tab-${tabs[next]}`)?.focus()
        }}><Icon size={16} />{t(key)}</button>)}
      </div>
      <div id={`library-${tab}`} role="tabpanel" aria-labelledby={`tab-${tab}`} className={`library-panel library-${tab}`}>
        {tab === 'find' && <>
          <p className="library-explainer">{t('searchStart')}</p>
          <HelpHint text={t('searchHint')}><label className="library-search"><Search size={18} /><input aria-label={t('search')} type="search" value={query} placeholder={t('searchPlaceholder')} onChange={event => onQuery(event.target.value)} /></label></HelpHint>
          {query && <div className="library-results"><output>{t('searchCount', { count:results.length })}</output>{results.length ? results.map(item => <button className="library-result" key={item.id} type="button" onClick={() => onOpen(item.id)}><FileText size={17}/><span><strong>{title(item)}</strong><small>{item.section || '—'}{item.tags ? ` · ${item.tags}` : ''}</small></span></button>) : <p>{t('searchEmpty')}</p>}</div>}
        </>}
        {tab === 'edit' && <>
          <p className="library-explainer"><CheckCircle2 size={15}/>{t('librarySavedHint')}</p>
          <div className="library-form">
            <HelpHint text={t('pageNameHint')}><label>{t('pageTitle')}<input value={page.title ?? ''} onChange={event => onPatch({ title:event.target.value })} /></label></HelpHint>
            <HelpHint text={t('sectionHint')}><label>{t('section')}<input list="library-sections" value={page.section ?? ''} onChange={event => onPatch({ section:event.target.value })}/></label></HelpHint>
            <datalist id="library-sections">{Array.from(new Set(pages.map(page => page.section).filter(Boolean))).map(section => <option key={section} value={section} />)}</datalist>
            <HelpHint text={t('tagsHint')}><label>{t('tags')}<input value={page.tags ?? ''} onChange={event => onPatch({ tags:event.target.value })} /></label></HelpHint>
            <HelpHint text={t('duplicateHint')}><button className="library-action" type="button" onClick={() => { onDuplicate(); setNotice({key:'duplicateDone'}) }}><Copy size={16}/>{t('duplicate')}</button></HelpHint>
          </div>
          <small className="library-date">{t('modified')}: {readableDate(page.updatedAt ?? page.createdAt)}</small>
          {page.images.map(image => <HelpHint text={t('imageDescriptionHint')} key={image.id}><label className="library-image-alt">{t('imageDescription')}<input value={image.alt ?? ''} onChange={event => updateAlt(image,event.target.value)}/></label></HelpHint>)}
        </>}
        {tab === 'files' && <>
          <p className="library-explainer">{t('attachHint')}</p>
          <HelpHint text={t('attachHint')}><label className="library-action library-upload"><Paperclip size={17}/>{t('attachButton')}<input aria-label={t('attachments')} className="hidden-input" type="file" multiple onChange={event => {
            const files = Array.from(event.target.files ?? []); event.target.value = ''
            void onAttach(files).then(() => setNotice({key:'attachmentsDone',values:{count:files.length}})).catch(() => setNotice({key:'attachFailed',error:true}))
          }}/></label></HelpHint>
          {!page.attachments?.length && <p className="library-empty">{t('attachmentsEmpty')}</p>}
          <ul className="library-list">{page.attachments?.map(attachment => <li key={attachment.id}><HelpHint text={t('downloadAttachmentHint')}><button className="library-file" type="button" onClick={() => void onDownload(attachment.id).catch(() => setNotice({key:'attachFailed',error:true}))}><Download size={16}/>{attachment.name}</button></HelpHint><HelpHint text={t('removeAttachmentHint')}><button className="library-icon-danger" type="button" aria-label={`${t('delete')} ${attachment.name}`} onClick={() => { onPatch({attachments:page.attachments!.filter(item => item.id !== attachment.id)});setNotice({key:'attachmentRemoved'}) }}><Trash2 size={16}/></button></HelpHint></li>)}</ul>
        </>}
        {tab === 'recover' && <div className="recovery-cards">
          <section className="recovery-card trash-card"><header><Trash2 size={18}/><h3>{t('trash')}</h3><span>{trash.length}</span></header><p>{t('trashIntro')}</p>
            {!trash.length ? <div className="library-empty">{t('trashEmpty')}</div> : <ul className="library-list">{trash.map((item,index) => <li key={item.id}><span>{item.title || t('page',{number:index+1})}<small>{new DOMParser().parseFromString(item.html,'text/html').body.textContent?.slice(0,70)}</small></span><HelpHint text={t('restorePageHint')}><button className="library-action" type="button" onClick={() => { onRestorePage(item.id);setNotice({key:'pageRestored'}) }}><RotateCcw size={15}/>{t('restorePage')}</button></HelpHint></li>)}</ul>}
          </section>
          <section className="recovery-card versions-card"><header><History size={18}/><h3>{t('recoverWork')}</h3></header><p>{t('versionsIntro')}</p>
            <HelpHint text={t('versionsHint')}><button className="library-action" type="button" disabled={loading} onClick={async () => {
              setLoading(true);setNotice(null)
              try { const next = await onLoadVersions();setVersions(next);setNotice({key:next.length ? 'versionsFound' : 'versionsEmpty',values:{count:next.length}}) } catch {setNotice({key:'restoreFailed',error:true})} finally {setLoading(false)}
            }}><History size={16}/>{t(loading ? 'versionsLoading' : 'versions')}</button></HelpHint>
            {versions && <ul className="library-list versions-list">{versions.map(version => <li key={version.id}><span><strong>{readableDate(version.time)}</strong><small>{t('versionPages',{count:version.project.pages.length})}</small></span><HelpHint text={t('restoreVersionHint')}><button className="library-action" type="button" disabled={loading} onClick={async () => {
              const date = readableDate(version.time)
              if (!window.confirm(t('restoreVersionConfirm',{date}))) return
              setLoading(true)
              try {await onRestoreVersion(version);setNotice({key:'versionRestored',date:version.time})} catch {setNotice({key:'restoreFailed',error:true})} finally {setLoading(false)}
            }}><RotateCcw size={15}/>{t('restoreVersion')}</button></HelpHint></li>)}</ul>}
          </section>
          <p className="library-local-note">{t('recoveryLocal')}</p>
        </div>}
      </div>
      {notice && <p className={notice.error ? 'library-notice error' : 'library-notice'} role={notice.error ? 'alert' : 'status'}><CheckCircle2 size={16}/>{t(notice.key,notice.date ? {date:readableDate(notice.date)} : notice.values)}</p>}
    </div>
  </details>
}
