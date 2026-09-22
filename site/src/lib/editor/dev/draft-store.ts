export interface EditorDraft {
  id: string;
  projectId: string;
  baseHash: string | undefined;
  fileName: string;
  bytes: Uint8Array;
  version: number;
  updated: number;
}

/** Each editing session owns a separate copy, including across multiple tabs. */
export class DraftStore {
  #database: Promise<IDBDatabase>;
  constructor() {
    this.#database = new Promise((resolve, reject) => {
      const request = indexedDB.open('office-pptx-editor-drafts', 1);
      request.onupgradeneeded = () => {
        request.result
          .createObjectStore('drafts', { keyPath: 'id' })
          .createIndex('projectId', 'projectId');
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    void this.#database.catch(() => {});
  }
  async list(projectId: string): Promise<EditorDraft[]> {
    const db = await this.#database;
    return new Promise((resolve, reject) => {
      const request = db
        .transaction('drafts')
        .objectStore('drafts')
        .index('projectId')
        .getAll(projectId);
      request.onsuccess = () =>
        resolve(
          (request.result as EditorDraft[])
            .filter((d) => d.projectId === projectId)
            .sort((a, b) => b.updated - a.updated),
        );
      request.onerror = () => reject(request.error);
    });
  }
  async put(draft: EditorDraft): Promise<void> {
    const db = await this.#database;
    return new Promise((resolve, reject) => {
      const tx = db.transaction('drafts', 'readwrite');
      const store = tx.objectStore('drafts');
      const request = store.get(draft.id);
      request.onsuccess = () => {
        // An earlier serialization can finish after a newer edit's checkpoint.
        if (!request.result || request.result.version <= draft.version) store.put(draft);
      };
      tx.oncomplete = () => resolve();
      tx.onabort = () => reject(tx.error);
      tx.onerror = () => reject(tx.error);
    });
  }
  async remove(id: string, version?: number): Promise<void> {
    const db = await this.#database;
    return new Promise((resolve, reject) => {
      const tx = db.transaction('drafts', 'readwrite');
      const store = tx.objectStore('drafts');
      const request = store.get(id);
      request.onsuccess = () => {
        if (version === undefined || request.result?.version === version) store.delete(id);
      };
      tx.oncomplete = () => resolve();
      tx.onabort = () => reject(tx.error);
      tx.onerror = () => reject(tx.error);
    });
  }
}
