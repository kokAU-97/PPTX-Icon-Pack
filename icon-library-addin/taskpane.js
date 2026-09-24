/* My Icon Library — task pane logic
 * Stores icons in IndexedDB (local to this computer/browser profile) and
 * inserts them onto the current slide using Office.js.
 */

(function () {
  "use strict";

  var DB_NAME = "IconLibraryDB";
  var DB_VERSION = 1;
  var STORE = "icons";

  var dbPromise = null;
  var allIcons = [];      // in-memory cache, newest first
  var officeReady = false;

  // ---------------------------------------------------------------------
  // IndexedDB helpers
  // ---------------------------------------------------------------------

  function openDB() {
    if (dbPromise) return dbPromise;
    dbPromise = new Promise(function (resolve, reject) {
      var req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = function () {
        var db = req.result;
        if (!db.objectStoreNames.contains(STORE)) {
          var store = db.createObjectStore(STORE, { keyPath: "id" });
          store.createIndex("name", "name", { unique: false });
        }
      };
      req.onsuccess = function () { resolve(req.result); };
      req.onerror = function () { reject(req.error); };
    });
    return dbPromise;
  }

  function tx(mode) {
    return openDB().then(function (db) {
      return db.transaction(STORE, mode).objectStore(STORE);
    });
  }

  function dbGetAll() {
    return tx("readonly").then(function (store) {
      return new Promise(function (resolve, reject) {
        var req = store.getAll();
        req.onsuccess = function () { resolve(req.result); };
        req.onerror = function () { reject(req.error); };
      });
    });
  }

  function dbPut(record) {
    return tx("readwrite").then(function (store) {
      return new Promise(function (resolve, reject) {
        var req = store.put(record);
        req.onsuccess = function () { resolve(record); };
        req.onerror = function () { reject(req.error); };
      });
    });
  }

  function dbDelete(id) {
    return tx("readwrite").then(function (store) {
      return new Promise(function (resolve, reject) {
        var req = store.delete(id);
        req.onsuccess = function () { resolve(); };
        req.onerror = function () { reject(req.error); };
      });
    });
  }

  function dbClear() {
    return tx("readwrite").then(function (store) {
      return new Promise(function (resolve, reject) {
        var req = store.clear();
        req.onsuccess = function () { resolve(); };
        req.onerror = function () { reject(req.error); };
      });
    });
  }

  function makeId() {
    return "icon-" + Date.now() + "-" + Math.random().toString(36).slice(2, 9);
  }

  // ---------------------------------------------------------------------
  // DOM references
  // ---------------------------------------------------------------------

  var el = {};

  function cacheDom() {
    el.app = document.getElementById("app");
    el.loading = document.getElementById("loading");
    el.fileInput = document.getElementById("file-input");
    el.searchInput = document.getElementById("search-input");
    el.clearAllBtn = document.getElementById("clear-all-btn");
    el.status = document.getElementById("status");
    el.grid = document.getElementById("grid");
    el.emptyState = document.getElementById("empty-state");
    el.noResults = document.getElementById("no-results");
    el.noResultsQuery = document.getElementById("no-results-query");
  }

  function setStatus(message) {
    el.status.textContent = message || "";
  }

  // ---------------------------------------------------------------------
  // Adding icons
  // ---------------------------------------------------------------------

  function readFileAsDataUrl(file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onload = function () { resolve(reader.result); };
      reader.onerror = function () { reject(reader.error); };
      reader.readAsDataURL(file);
    });
  }

  function niceName(fileName) {
    var withoutExt = fileName.replace(/\.[^/.]+$/, "");
    return withoutExt.replace(/[-_]+/g, " ").trim() || fileName;
  }

  function handleFiles(fileList) {
    var files = Array.prototype.slice.call(fileList || []);
    if (!files.length) return;

    var accepted = files.filter(function (f) {
      return /\.(svg|png|jpe?g)$/i.test(f.name);
    });

    if (!accepted.length) {
      setStatus("Those files aren't supported. Use SVG, PNG or JPG.");
      return;
    }

    setStatus("Adding " + accepted.length + " icon" + (accepted.length > 1 ? "s" : "") + "…");

    Promise.all(
      accepted.map(function (file) {
        return readFileAsDataUrl(file).then(function (dataUrl) {
          return dbPut({
            id: makeId(),
            name: niceName(file.name),
            type: file.type || (/\.svg$/i.test(file.name) ? "image/svg+xml" : "image/png"),
            dataUrl: dataUrl,
            createdAt: Date.now()
          });
        });
      })
    )
      .then(function (added) {
        setStatus(added.length + " icon" + (added.length > 1 ? "s" : "") + " added.");
        return refresh();
      })
      .catch(function (err) {
        console.error(err);
        setStatus("Something went wrong while adding icons.");
      });
  }

  // ---------------------------------------------------------------------
  // Rendering
  // ---------------------------------------------------------------------

  function currentQuery() {
    return (el.searchInput.value || "").trim().toLowerCase();
  }

  function render() {
    var query = currentQuery();
    var list = query
      ? allIcons.filter(function (icon) { return icon.name.toLowerCase().indexOf(query) !== -1; })
      : allIcons;

    el.grid.innerHTML = "";

    var hasAny = allIcons.length > 0;
    el.emptyState.hidden = hasAny;
    el.noResults.hidden = !(hasAny && query && list.length === 0);
    if (!el.noResults.hidden) el.noResultsQuery.textContent = query;

    list.forEach(function (icon) {
      el.grid.appendChild(buildTile(icon));
    });
  }

  function buildTile(icon) {
    var tile = document.createElement("div");
    tile.className = "icon-tile";
    tile.tabIndex = 0;
    tile.setAttribute("role", "button");
    tile.title = "Insert \"" + icon.name + "\"";

    var img = document.createElement("img");
    img.src = icon.dataUrl;
    img.alt = icon.name;
    tile.appendChild(img);

    var label = document.createElement("span");
    label.className = "tile-name";
    label.textContent = icon.name;
    tile.appendChild(label);

    var del = document.createElement("button");
    del.className = "delete-btn";
    del.type = "button";
    del.setAttribute("aria-label", "Delete " + icon.name);
    del.textContent = "\u00D7";
    del.addEventListener("click", function (e) {
      e.stopPropagation();
      removeIcon(icon.id);
    });
    tile.appendChild(del);

    function activate() { insertIcon(tile, icon); }
    tile.addEventListener("click", activate);
    tile.addEventListener("keydown", function (e) {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        activate();
      }
    });

    return tile;
  }

  function refresh() {
    return dbGetAll().then(function (records) {
      records.sort(function (a, b) { return b.createdAt - a.createdAt; });
      allIcons = records;
      render();
    });
  }

  function removeIcon(id) {
    dbDelete(id)
      .then(function () {
        setStatus("Icon deleted.");
        return refresh();
      })
      .catch(function (err) {
        console.error(err);
        setStatus("Couldn't delete that icon.");
      });
  }

  function clearAll() {
    if (!allIcons.length) return;
    var ok = window.confirm("Remove all " + allIcons.length + " icons from your library? This can't be undone.");
    if (!ok) return;
    dbClear()
      .then(function () {
        setStatus("Library cleared.");
        return refresh();
      })
      .catch(function (err) {
        console.error(err);
        setStatus("Couldn't clear the library.");
      });
  }

  // ---------------------------------------------------------------------
  // Inserting into the slide
  // ---------------------------------------------------------------------

  // SVGs are rasterized to PNG the first time they're inserted, then the
  // rendered PNG is cached on the record so later inserts are instant.
  function rasterizeSvg(dataUrl) {
    return new Promise(function (resolve, reject) {
      var img = new Image();
      img.onload = function () {
        var target = 512;
        var w = img.naturalWidth || target;
        var h = img.naturalHeight || target;
        var scale = target / Math.max(w, h);
        var canvas = document.createElement("canvas");
        canvas.width = Math.max(1, Math.round(w * scale));
        canvas.height = Math.max(1, Math.round(h * scale));
        var ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
        try {
          resolve(canvas.toDataURL("image/png"));
        } catch (err) {
          reject(err);
        }
      };
      img.onerror = function () { reject(new Error("Could not rasterize SVG")); };
      img.src = dataUrl;
    });
  }

  function getInsertableDataUrl(icon) {
    if (icon.insertData) return Promise.resolve(icon.insertData);
    if (icon.type === "image/svg+xml" || /^data:image\/svg\+xml/.test(icon.dataUrl)) {
      return rasterizeSvg(icon.dataUrl).then(function (png) {
        icon.insertData = png;
        dbPut(icon).catch(function () {}); // best-effort cache
        return png;
      });
    }
    return Promise.resolve(icon.dataUrl);
  }

  function insertIcon(tileEl, icon) {
    if (!officeReady) {
      setStatus("Open this add-in inside PowerPoint to insert icons.");
      return;
    }

    tileEl.classList.add("inserting");
    setStatus("Inserting \"" + icon.name + "\"…");

    getInsertableDataUrl(icon)
      .then(function (dataUrl) {
        var base64 = dataUrl.split(",")[1] || "";
        Office.context.document.setSelectedDataAsync(
          base64,
          { coercionType: Office.CoercionType.Image },
          function (result) {
            tileEl.classList.remove("inserting");
            if (result.status === Office.AsyncResultStatus.Failed) {
              console.error(result.error);
              setStatus("Couldn't insert that icon: " + result.error.message);
            } else {
              setStatus("Inserted \"" + icon.name + "\".");
            }
          }
        );
      })
      .catch(function (err) {
        tileEl.classList.remove("inserting");
        console.error(err);
        setStatus("Couldn't prepare that icon for insertion.");
      });
  }

  // ---------------------------------------------------------------------
  // Wiring
  // ---------------------------------------------------------------------

  function init() {
    cacheDom();

    el.fileInput.addEventListener("change", function (e) {
      handleFiles(e.target.files);
      e.target.value = ""; // allow re-selecting the same file later
    });

    el.searchInput.addEventListener("input", render);
    el.clearAllBtn.addEventListener("click", clearAll);

    refresh().then(function () {
      el.loading.hidden = true;
      el.app.hidden = false;
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }

  // Office.js is only fully available once the add-in is sideloaded and
  // running inside PowerPoint. The library above still works as a plain
  // page (for previewing/managing icons) even before this resolves.
  if (window.Office && Office.onReady) {
    Office.onReady(function () {
      officeReady = true;
    });
  }
})();
