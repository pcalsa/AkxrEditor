// IndexedDB stores structured data without serializing embedded images on the UI thread.
const database = () => new Promise<IDBDatabase>((resolve, reject) => {
  const request = indexedDB.open('akxr-local', 1)
  request.onupgradeneeded = () => {
    request.result.createObjectStore('drafts')
    request.result.createObjectStore('versions', { keyPath: 'id' })
  }
  request.onsuccess = () => resolve(request.result)
  request.onerror = () => reject(request.error)
})
export type LocalVersion<T = unknown> = { id: string; time: number; project: T }
export class StorageConflict extends Error {}
export async function readDraft<T>(): Promise<{ revision: number; project: T } | undefined> {
  const db = await database()
  try {
    return await new Promise((resolve, reject) => {
      const request = db.transaction('drafts').objectStore('drafts').get('current')
      request.onsuccess = () => resolve(request.result)
      request.onerror = () => reject(request.error)
    })
  } finally { db.close() }
}
export async function writeDraft<T>(project: T, expectedRevision: number, checkpoint = false): Promise<number> {
  const db = await database()
  try {
    return await new Promise((resolve, reject) => {
      const transaction = db.transaction(['drafts', 'versions'], 'readwrite')
      const store = transaction.objectStore('drafts')
      const request = store.get('current')
      let conflict = false
      const revision = expectedRevision + 1
      request.onsuccess = () => {
        if ((request.result?.revision ?? 0) !== expectedRevision) {
          conflict = true
          transaction.abort()
          return
        }
        store.put({ revision, project }, 'current')
        if (checkpoint) {
          transaction.objectStore('versions').put({ id: crypto.randomUUID(), time: Date.now(), project })
          const versions = transaction.objectStore('versions').getAll()
          versions.onsuccess = () => {
            const sorted = versions.result.sort((a: LocalVersion, b: LocalVersion) => b.time - a.time)
            sorted.slice(50).forEach((version: LocalVersion) => transaction.objectStore('versions').delete(version.id))
          }
        }
      }
      transaction.oncomplete = () => resolve(revision)
      transaction.onabort = () => reject(conflict ? new StorageConflict('Another tab saved this draft.') : transaction.error)
      transaction.onerror = () => reject(transaction.error)
    })
  } finally { db.close() }
}
export async function localVersions<T>(): Promise<LocalVersion<T>[]> {
  const db = await database()
  try {
    return await new Promise((resolve, reject) => {
      const request = db.transaction('versions').objectStore('versions').getAll()
      request.onsuccess = () => resolve(request.result.sort((a: LocalVersion, b: LocalVersion) => b.time - a.time))
      request.onerror = () => reject(request.error)
    })
  } finally { db.close() }
}
