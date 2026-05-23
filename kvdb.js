// kvdb.js - Simple SQLite(KV) store using sql.js + IndexedDB
// Exposes: window.KVDB.open(dbName) -> {get,set,del,close}
// Stores a single table kv(key TEXT PRIMARY KEY, value TEXT)

(function(){
  const SQLJS_VERSION = "1.9.0";
  const SQLJS_BASE = `https://cdn.jsdelivr.net/npm/sql.js@${SQLJS_VERSION}/dist/`;

  function idbOpen(){
    return new Promise((resolve,reject)=>{
      const req = indexedDB.open("TKB_SQLJS_DB", 1);
      req.onupgradeneeded = ()=> {
        const db = req.result;
        if (!db.objectStoreNames.contains("files")){
          db.createObjectStore("files");
        }
      };
      req.onsuccess = ()=> resolve(req.result);
      req.onerror = ()=> reject(req.error);
    });
  }

  function idbGet(db, key){
    return new Promise((resolve,reject)=>{
      const tx = db.transaction("files","readonly");
      const st = tx.objectStore("files");
      const req = st.get(key);
      req.onsuccess = ()=> resolve(req.result || null);
      req.onerror = ()=> reject(req.error);
    });
  }

  function idbPut(db, key, val){
    return new Promise((resolve,reject)=>{
      const tx = db.transaction("files","readwrite");
      const st = tx.objectStore("files");
      const req = st.put(val, key);
      req.onsuccess = ()=> resolve(true);
      req.onerror = ()=> reject(req.error);
    });
  }

  async function ensureSql(){
    if (window.__SQLJS_PROMISE__) return window.__SQLJS_PROMISE__;
    if (typeof initSqlJs !== "function"){
      throw new Error("sql.js chưa được load. Hãy thêm <script src='.../sql-wasm.js'></script> trước kvdb.js");
    }
    window.__SQLJS_PROMISE__ = initSqlJs({
      locateFile: (file)=> SQLJS_BASE + file
    });
    return window.__SQLJS_PROMISE__;
  }

  async function open(dbName){
    const SQL = await ensureSql();
    const idb = await idbOpen();
    const bin = await idbGet(idb, dbName);

    const db = bin ? new SQL.Database(new Uint8Array(bin)) : new SQL.Database();
    db.run("CREATE TABLE IF NOT EXISTS kv (key TEXT PRIMARY KEY, value TEXT);");

    let dirty = false;
    let flushTimer = null;

    async function flush(){
      if (!dirty) return;
      dirty = false;
      const data = db.export(); // Uint8Array
      await idbPut(idb, dbName, data);
    }

    function scheduleFlush(){
      dirty = true;
      if (flushTimer) return;
      flushTimer = setTimeout(async ()=>{
        flushTimer = null;
        try{ await flush(); }catch(e){ console.warn("KVDB flush failed:", e); }
      }, 250);
    }

    function get(key){
      const stmt = db.prepare("SELECT value FROM kv WHERE key=?");
      stmt.bind([key]);
      let val = null;
      if (stmt.step()){
        val = stmt.getAsObject().value;
      }
      stmt.free();
      return Promise.resolve(val);
    }

    function set(key, value){
      db.run("INSERT INTO kv(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value", [key, value]);
      scheduleFlush();
      return Promise.resolve(true);
    }

    function del(key){
      db.run("DELETE FROM kv WHERE key=?", [key]);
      scheduleFlush();
      return Promise.resolve(true);
    }

    async function close(){
      await flush();
      db.close();
      idb.close();
    }

    return { get, set, del, flush, close };
  }

  window.KVDB = { open };
})();