// db.js — Raimak Verizon LMS Local Storage Engine v2.0

const LocalDB = {
  dbName: "RaimakVZDB", // 🚀 Isolated from Frontier's "RaimakDB"
  version: 2, // 🚀 BUMPED TO VERSION 2 TO TRIGGER THE UPGRADE
  db: null,

  // 1. Open the connection and build the tables
  init: async function () {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(this.dbName, this.version);

      // This only runs the very first time the app loads, or if we change the version number
      request.onupgradeneeded = (event) => {
        const db = event.target.result;

        // Create our core CRM tables. 'id' is the primary key.
        if (!db.objectStoreNames.contains("activity_logs")) {
          db.createObjectStore("activity_logs", { keyPath: "id" });
        }
        if (!db.objectStoreNames.contains("leads")) {
          db.createObjectStore("leads", { keyPath: "id" });
        }

        // 🚀 THE FIX: Renamed from sms_messages to sms_history to match the new Delta Engine
        if (!db.objectStoreNames.contains("sms_history")) {
          db.createObjectStore("sms_history", { keyPath: "id" });
        }
      };

      request.onsuccess = (event) => {
        this.db = event.target.result;
        resolve(true);
      };

      request.onerror = (event) => {
        console.error("IndexedDB Error:", event.target.errorCode);
        reject(event.target.error);
      };
    });
  },

  // 2. Save a massive array of items instantly
  saveItems: async function (storeName, items) {
    return new Promise((resolve, reject) => {
      const transaction = this.db.transaction([storeName], "readwrite");
      const store = transaction.objectStore(storeName);

      items.forEach((item) => store.put(item));

      transaction.oncomplete = () => resolve(true);
      transaction.onerror = (event) => reject(event.target.error);
    });
  },

  // 3. Load the entire iceberg out of the hard drive
  getAllItems: async function (storeName) {
    return new Promise((resolve, reject) => {
      const transaction = this.db.transaction([storeName], "readonly");
      const store = transaction.objectStore(storeName);
      const request = store.getAll();

      request.onsuccess = (event) => resolve(event.target.result);
      request.onerror = (event) => reject(event.target.error);
    });
  },

  // 4. Delete a specific item by ID
  deleteItem: async function (storeName, id) {
    return new Promise((resolve, reject) => {
      const transaction = this.db.transaction([storeName], "readwrite");
      const store = transaction.objectStore(storeName);

      const request = store.delete(id);

      request.onsuccess = () => resolve(true);
      request.onerror = (event) => reject(event.target.error);
    });
  },
};
