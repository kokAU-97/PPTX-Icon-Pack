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

  var clearAllPendingTimer = null;

  function resetClearAllButton() {
    if (clearAllPendingTimer) {
      clearTimeout(clearAllPendingTimer);
      clearAllPendingTimer = null;
    }
    el.clearAllBtn.textContent = "Clear all";
    el.clearAllBtn.classList.remove("btn-confirm");
  }

  function clearAll() {
    if (!allIcons.length) return;

    if (!clearAllPendingTimer) {
      // First click: arm it. Native window.confirm() dialogs are
      // unreliable inside Office task panes (some hosts silently block or
      // auto-dismiss them), so we ask for confirmation in the page itself.
      el.clearAllBtn.textContent = "Click again to confirm";
      el.clearAllBtn.classList.add("btn-confirm");
      clearAllPendingTimer = setTimeout(resetClearAllButton, 4000);
      return;
    }

    // Second click within the window: actually clear.
    resetClearAllButton();
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

  // Decodes a base64 payload to a UTF-8 string, with fallbacks for older
  // task pane rendering engines that lack TextDecoder.
  function base64ToUtf8(base64) {
    try {
      if (typeof TextDecoder !== "undefined") {
        var binary = atob(base64);
        var bytes = new Uint8Array(binary.length);
        for (var i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
        return new TextDecoder("utf-8").decode(bytes);
      }
    } catch (e) { /* fall through */ }
    try {
      return decodeURIComponent(escape(atob(base64)));
    } catch (e) {
      return atob(base64);
    }
  }

  // SVGs are rasterized to PNG the first time they're inserted, then the
  // rendered PNG is cached on the record so later inserts are instant.
  //
  // Simply loading an SVG data URL into an <img> and drawing it onto a
  // larger canvas produces blurry results in most browsers: the SVG gets
  // rasterized once at its own small intrinsic size (often the default
  // ~300x150, or whatever tiny width/height it declares) and THEN that
  // low-res bitmap gets stretched onto the bigger canvas. To get a crisp
  // result, we rewrite the SVG's own width/height attributes to the full
  // target resolution first, so the browser's vector renderer draws it
  // natively at that size instead of upscaling a small bitmap.
  function rasterizeSvg(dataUrl) {
    return new Promise(function (resolve, reject) {
      try {
        var target = 640;
        var base64 = dataUrl.split(",")[1] || "";
        var svgText = base64ToUtf8(base64);
        var doc = new DOMParser().parseFromString(svgText, "image/svg+xml");
        var svgEl = doc.documentElement;
        if (!svgEl || svgEl.nodeName.toLowerCase() !== "svg") {
          throw new Error("Not a valid SVG");
        }

        var w, h;
        var viewBox = svgEl.getAttribute("viewBox");
        if (viewBox) {
          var parts = viewBox.trim().split(/[\s,]+/).map(Number);
          w = parts[2]; h = parts[3];
        } else {
          w = parseFloat(svgEl.getAttribute("width"));
          h = parseFloat(svgEl.getAttribute("height"));
        }
        if (!w || !h) { w = target; h = target; }
        if (!svgEl.getAttribute("viewBox")) {
          svgEl.setAttribute("viewBox", "0 0 " + w + " " + h);
        }

        var scale = target / Math.max(w, h);
        var outW = Math.max(1, Math.round(w * scale));
        var outH = Math.max(1, Math.round(h * scale));
        svgEl.setAttribute("width", outW);
        svgEl.setAttribute("height", outH);

        var serialized = new XMLSerializer().serializeToString(doc);
        var blob = new Blob([serialized], { type: "image/svg+xml;charset=utf-8" });
        var url = URL.createObjectURL(blob);

        var img = new Image();
        img.onload = function () {
          var canvas = document.createElement("canvas");
          canvas.width = outW;
          canvas.height = outH;
          var ctx = canvas.getContext("2d");
          ctx.drawImage(img, 0, 0, outW, outH);
          URL.revokeObjectURL(url);
          try {
            resolve(canvas.toDataURL("image/png"));
          } catch (err) {
            reject(err);
          }
        };
        img.onerror = function () {
          URL.revokeObjectURL(url);
          reject(new Error("Could not rasterize SVG"));
        };
        img.src = url;
      } catch (err) {
        reject(err);
      }
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

  function isSvgIcon(icon) {
    return icon.type === "image/svg+xml" || /^data:image\/svg\+xml/.test(icon.dataUrl);
  }

  function insertIcon(tileEl, icon) {
    if (!officeReady) {
      setStatus("Open this add-in inside PowerPoint to insert icons.");
      return;
    }

    tileEl.classList.add("inserting");
    setStatus("Inserting \"" + icon.name + "\"…");

    if (isSvgIcon(icon)) {
      // Try handing PowerPoint the raw SVG first: dragging an SVG file in
      // natively gets you a special recolorable "SVG picture" (PowerPoint
      // can tint the whole icon one flat color via Graphics Fill), which a
      // pre-rasterized PNG can never get. If PowerPoint's insertion API
      // rejects raw SVG here, we fall back to the flattened PNG so the
      // insert doesn't just fail outright.
      var base64 = icon.dataUrl.split(",")[1] || "";
      Office.context.document.setSelectedDataAsync(
        base64,
        { coercionType: Office.CoercionType.Image },
        function (result) {
          if (result.status === Office.AsyncResultStatus.Failed) {
            console.warn("Raw SVG insert failed, falling back to PNG:", result.error);
            getInsertableDataUrl(icon)
              .then(function (png) {
                var pngBase64 = png.split(",")[1] || "";
                Office.context.document.setSelectedDataAsync(
                  pngBase64,
                  { coercionType: Office.CoercionType.Image },
                  function (fallbackResult) {
                    tileEl.classList.remove("inserting");
                    if (fallbackResult.status === Office.AsyncResultStatus.Failed) {
                      setStatus("Couldn't insert that icon: " + fallbackResult.error.message);
                    } else {
                      setStatus("Inserted \"" + icon.name + "\" (as a flattened image).");
                    }
                  }
                );
              })
              .catch(function () {
                tileEl.classList.remove("inserting");
                setStatus("Couldn't insert that icon.");
              });
          } else {
            tileEl.classList.remove("inserting");
            setStatus("Inserted \"" + icon.name + "\".");
          }
        }
      );
    } else {
      var rasterBase64 = icon.dataUrl.split(",")[1] || "";
      Office.context.document.setSelectedDataAsync(
        rasterBase64,
        { coercionType: Office.CoercionType.Image },
        function (result) {
          tileEl.classList.remove("inserting");
          if (result.status === Office.AsyncResultStatus.Failed) {
            setStatus("Couldn't insert that icon: " + result.error.message);
          } else {
            setStatus("Inserted \"" + icon.name + "\".");
          }
        }
      );
    }
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
