const $ = (s) => document.querySelector(s);

const els = {
  btnOpenFolder: $("#btnOpenFolder"),
  folderStatus: $("#folderStatus"),

  btnLoad: $("#btnLoad"),
  btnSave: $("#btnSave"),
  fileNote: $("#fileNote"),
  filePicker: $("#filePicker"),

  search: $("#search"),
  btnAddProduct: $("#btnAddProduct"),
  btnAddCategory: $("#btnAddCategory"),

  catSearch: $("#catSearch"),
  categoryList: $("#categoryList"),
  filterLine: $("#filterLine"),
  productGrid: $("#productGrid"),
  empty: $("#empty"),

  catDlg: $("#catDlg"),
  catForm: $("#catForm"),
  catTitle: $("#catTitle"),
  catName: $("#catName"),

  prodDlg: $("#prodDlg"),
  prodForm: $("#prodForm"),
  prodTitle: $("#prodTitle"),
  prodName: $("#prodName"),
  prodCtSize: $("#prodCtSize"),
  prodCategory: $("#prodCategory"),
  prodCategoriesBox: $("#prodCategoriesBox"),
  prodWordings: $("#prodWordings"),
  prodCodes: $("#prodCodes"),
  prodNotes: $("#prodNotes"),
  prodImageFileName: $("#prodImageFileName"),
  imgPreview: $("#imgPreview"),
  btnPickImage: $("#btnPickImage"),
  btnClearImage: $("#btnClearImage"),
  imgFilePicker: $("#imgFilePicker"),
  imgHint: $("#imgHint"),

  stockDlg: $("#stockDlg"),
  stockHeader: $("#stockHeader"),
  lotList: $("#lotList"),
  stockExpiry: $("#stockExpiry"),
stockUnknownExpiry: $("#stockUnknownExpiry"),
  stockQty: $("#stockQty"),
  btnStockAdd: $("#btnStockAdd"),
  btnStockWithdraw: $("#btnStockWithdraw"),
  btnStockOrder: $("#btnStockOrder"),
  stockMsg: $("#stockMsg"),

  infoDlg: $("#infoDlg"),
  infoTitle: $("#infoTitle"),
  infoSub: $("#infoSub"),
  infoWordings: $("#infoWordings"),
  infoCodes: $("#infoCodes"),
  infoLots: $("#infoLots"),
  
  btnCatCancel: $("#btnCatCancel"),
  btnProdCancel: $("#btnProdCancel"), 
  btnProdDuplicate: $("#btnProdDuplicate"),

  cardTpl: $("#cardTpl"),
};


// Stock unit selector (Units vs CTs)
stockUnitSelector.id = "stockUnitSelector";
stockUnitSelector.innerHTML = `
    <option value="units">Units</option>
    <option value="ct">Cartons (CT)</option>
`;



/* -------------------- Supabase + Portal Auth (username/password) -------------------- */
const SUPABASE_URL = window.SUPABASE_URL;
const SUPABASE_ANON_KEY = window.SUPABASE_ANON_KEY;
const SUPABASE_FN_NAME = window.SUPABASE_FN_NAME || "hyper-worker";
const CATALOGUE_ROW_ID = "main";

let supabaseClient = null;

const PORTAL_SESSION_KEY = "portal_session_v1";
let portalSession = null; // { token, role, username, exp }

function loadPortalSession() {
  try {
    const raw = localStorage.getItem(PORTAL_SESSION_KEY);
    if (!raw) return null;
    const s = JSON.parse(raw);
    if (!s?.token) return null;
    if (s.exp && Date.now() > (s.exp * 1000)) {
      localStorage.removeItem(PORTAL_SESSION_KEY);
      return null;
    }
    return s;
  } catch { return null; }
}
function savePortalSession(s) {
  portalSession = s;
  localStorage.setItem(PORTAL_SESSION_KEY, JSON.stringify(s));
  refreshAuthUI();
}
function clearPortalSession() {
  portalSession = null;
  localStorage.removeItem(PORTAL_SESSION_KEY);
  refreshAuthUI();
}
function isAdmin() { return portalSession?.role === "admin"; }

function refreshAuthUI() {
  const st = document.getElementById("authStatus");
  const btnLogin = document.getElementById("btnLogin");
  const btnLogout = document.getElementById("btnLogout");
  const u = document.getElementById("authUser");
  const p = document.getElementById("authPass");

  if (portalSession?.token) {
    if (st) st.textContent = `Logged as ${portalSession.username} (${portalSession.role})`;
    if (btnLogin) btnLogin.style.display = "none";
    if (btnLogout) btnLogout.style.display = "";
    if (u) u.style.display = "none";
    if (p) p.style.display = "none";
  } else {
    if (st) st.textContent = "Viewer mode";
    if (btnLogin) btnLogin.style.display = "";
    if (btnLogout) btnLogout.style.display = "none";
    if (u) u.style.display = "";
    if (p) p.style.display = "";
  }

  const editable = isAdmin();
  if (els?.btnSave) {
    els.btnSave.disabled = !editable;
    els.btnSave.title = editable ? "" : "Login as admin to save changes";
  }
  if (els?.btnAddProduct) els.btnAddProduct.style.display = editable ? "" : "none";
  if (els?.btnAddCategory) els.btnAddCategory.style.display = editable ? "" : "none";
}

async function initSupabase() {
  if (!SUPABASE_URL || !SUPABASE_ANON_KEY || String(SUPABASE_URL).includes("PASTE_")) {
    console.warn("Supabase not configured yet.");
    return;
  }
  supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
}

async function loadCatalogueOnline() {
  if (!supabaseClient) return false;
  const { data, error } = await supabaseClient
    .from("catalogue")
    .select("data, updated_at")
    .eq("id", CATALOGUE_ROW_ID)
    .single();

  if (error) {
    console.warn("Supabase load failed:", error.message);
    return false;
  }

  const obj = data?.data || {};
  validateAndNormalize(obj);
  state = obj;
  loadedFileName = "online:supabase/catalogue/main";
  setEnabled(true);
  setDirty(false);
  render();
  return true;
}

async function portalLogin(username, password) {
  if (!supabaseClient) {
    alert("Supabase not configured (URL/ANON key missing).");
    return;
  }
  const { data, error } = await supabaseClient.functions.invoke(SUPABASE_FN_NAME, {
    body: { action: "login", username, password }
  });
  if (error) { alert("Login failed: " + error.message); return; }
  if (!data?.ok) { alert(data?.error || "Invalid credentials"); return; }
  savePortalSession({ token: data.token, role: data.role, username: data.username, exp: data.exp });
}

async function portalSaveCatalogue() {
  if (!isAdmin()) { alert("Admin login required to save."); return; }
  if (!supabaseClient) { alert("Supabase not configured yet."); return; }
  const { data, error } = await supabaseClient.functions.invoke(SUPABASE_FN_NAME, {
    body: { action: "save", token: portalSession.token, catalogue: state }
  });
  if (error) { alert("Save failed: " + error.message); return; }
  if (!data?.ok) { alert(data?.error || "Save failed"); return; }
  setDirty(false);
  alert("Saved online ✅");
}
/* ------------------------------------------------------------------------------- */

let state = null;
let loadedFileName = "";

let ui = {
  selectedCategoryId: "__all__",
  search: "",
  catSearch: "",
  editingCategoryId: null,
  editingProductId: null,
  stockProductId: null,
};

const fs = {
  folderHandle: null,
  jsonFileName: "catalogue.json",
};

// Autosave debounce (folder mode only)
let autosaveTimer = null;
let autosaveInFlight = false;
let autosaveQueued = false;

function requestAutosave() {
  if (!fs.folderHandle) return;
  if (!state) return;

  if (autosaveTimer) clearTimeout(autosaveTimer);
  autosaveTimer = setTimeout(() => {
    doAutosave();
  }, 600);
}

async function doAutosave() {
  if (!fs.folderHandle || !state) return;

  if (autosaveInFlight) {
    autosaveQueued = true;
    return;
  }
  autosaveInFlight = true;
  autosaveQueued = false;

  try {
    await saveJsonToFolder();
    setDirty(false);
    if (els.folderStatus) els.folderStatus.textContent = "Folder mode: ON (saved)";
  } catch (e) {
    if (els.folderStatus) els.folderStatus.textContent = "Folder mode: ON (save failed)";
    console.error(e);
  } finally {
    autosaveInFlight = false;
    if (autosaveQueued) {
      autosaveQueued = false;
      doAutosave();
    }
  }
}

function wire() {
  // Folder mode
  if (els.btnOpenFolder) {
    els.btnOpenFolder.addEventListener("click", openCatalogueFolder);
  }

els.stockUnknownExpiry.addEventListener("change", () => {
  if (els.stockUnknownExpiry.checked) {
    els.stockExpiry.value = "";
    els.stockExpiry.disabled = true;
  } else {
    els.stockExpiry.disabled = false;
  }
});

  // Manual load
  els.btnLoad.addEventListener("click", () => els.filePicker.click());
  els.filePicker.addEventListener("change", loadJsonFromPicker);
  
  // Dialog cancel buttons
  els.btnCatCancel.addEventListener("click", () => {
    ui.editingCategoryId = null;
    els.catDlg.close();
  });

  els.btnProdCancel.addEventListener("click", () => {
    ui.editingProductId = null;
    els.prodDlg.close();
  });

  // Save button
  els.btnSave.addEventListener("click", async () => {
    if (supabaseClient) { await portalSaveCatalogue(); return; }

    if (!state) return;

    if (fs.folderHandle) {
      await saveJsonToFolder();
      setDirty(false);
      if (els.folderStatus) els.folderStatus.textContent = "Folder mode: ON (saved)";
      alert("Saved to data/catalogue.json");
    } else {
      downloadJSON(state, "catalogue.json");
    }
  });

  // Search
  els.search.addEventListener("input", () => {
    ui.search = els.search.value.trim();
    render();
  });

  els.btnAddCategory.addEventListener("click", () => openCategoryDlg(null));
  els.btnAddProduct.addEventListener("click", () => openProductDlg(null));
  
  els.catSearch.addEventListener("input", () => {
    ui.catSearch = els.catSearch.value.toLowerCase();
    renderCategories(); // Re-render the list as you type
  });

  // Category submit
  els.catForm.addEventListener("submit", (e) => {
  e.preventDefault();
  const name = (els.catName.value || "").trim();
  if (!name) return;

  if (ui.editingCategoryId) {
    const c = state.categories.find(x => x.id === ui.editingCategoryId);
    if (!c) return;
    c.name = name;
  } else {
    state.categories.push({ id: uid("cat"), name });
  }

  ui.editingCategoryId = null;
  els.catDlg.close();
  setDirty(true);
  render();
});

  // Product submit
  els.prodForm.addEventListener("submit", (e) => {
  e.preventDefault();
  const name = (els.prodName.value || "").trim();
  if (!name) return;

  const ctSize = parseInt($("#prodCtSize").value) || null;
  const selectedCategoryIds = Array
    .from(els.prodCategoriesBox.querySelectorAll('input[type="checkbox"]:checked'))
    .map(cb => cb.value);
  const wordings = splitLines(els.prodWordings.value);
  const codes = splitLines(els.prodCodes.value);
  const notes = (els.prodNotes.value || "").trim();
  const imageFileName = (els.prodImageFileName.value || "").trim();

  if (ui.editingProductId) {
    const p = state.products.find(x => x.id === ui.editingProductId);
    if (!p) return;
    p.name = name;
    p.ctSize = ctSize;
    p.categoryIds = selectedCategoryIds;
    delete p.categoryId;
    p.wordings = wordings;
    p.codes = codes;
    p.notes = notes;
    p.imageFileName = imageFileName;
    p.lots = normalizeLots(p.lots);
  } else {
    state.products.push({
      id: uid("prod"),
      name,
      ctSize,
      categoryIds: selectedCategoryIds,
      wordings,
      codes,
      notes,
      imageFileName,
      lots: [],
    });
  }

  ui.editingProductId = null;
  els.prodDlg.close();
  setDirty(true);
  render();
});

  els.btnProdDuplicate.addEventListener("click", () => {
    duplicateProductFromDialog();
  });

  // Stock controls
  els.btnStockAdd.addEventListener("click", () => adjustStock(+1));
  els.btnStockWithdraw.addEventListener("click", () => adjustStock(-1));
  els.btnStockOrder.addEventListener("click", () => createOrder());
  attachDMYMask(els.stockExpiry);

  // Image picker controls
  els.btnPickImage.addEventListener("click", () => els.imgFilePicker.click());
  els.btnClearImage.addEventListener("click", () => {
    els.prodImageFileName.value = "";
    els.imgHint.textContent = "Image cleared.";
    renderProductPreview(null);
    setDirty(true);
  });

  els.imgFilePicker.addEventListener("change", async () => {
    const file = els.imgFilePicker.files && els.imgFilePicker.files[0];
    els.imgFilePicker.value = "";
    if (!file) return;

    try {
      if (fs.folderHandle) {
        const storedName = await copyPickedImageToMedia(file);
        els.prodImageFileName.value = storedName;
        els.imgHint.textContent = `Copied to media/${storedName}`;
      } else {
        els.prodImageFileName.value = file.name;
        els.imgHint.textContent = `Now copy that file into media/ as: ${file.name}`;
      }
      renderProductPreview({ imageFileName: els.prodImageFileName.value });
      setDirty(true);
    } catch (e) {
      alert("Could not set image: " + (e?.message || e));
    }
  });

  // locked until loaded
  setEnabled(false);
  els.filterLine.textContent = "Tip: Chrome/Edge: use “Use catalogue folder…” for auto-load & auto-save. Firefox: use “Load JSON”.";

  if (els.folderStatus) {
    els.folderStatus.textContent = window.showDirectoryPicker
      ? "Folder mode: OFF"
      : "Folder mode: Not supported (use Chrome/Edge)";
  }
}

/* Folder mode (Chrome/Edge) */
async function openCatalogueFolder() {
  if (!window.showDirectoryPicker) {
    alert("Folder mode is not supported in this browser. Use Chrome or Edge.");
    return;
  }
  try {
    const dir = await window.showDirectoryPicker({
      id: "catalogueFolder",
      mode: "readwrite",
      startIn: "documents",
    });

    fs.folderHandle = dir;
    if (els.folderStatus) els.folderStatus.textContent = "Folder mode: ON (loading…)";

    await loadJsonFromFolder();

    if (els.folderStatus) els.folderStatus.textContent = "Folder mode: ON (ready)";
  } catch {
    // cancelled
  }
}

async function loadJsonFromFolder() {
  if (!fs.folderHandle) return;

  const dataDir = await fs.folderHandle.getDirectoryHandle("data", { create: false });
  const fileHandle = await dataDir.getFileHandle(fs.jsonFileName, { create: false });
  const file = await fileHandle.getFile();
  const text = await file.text();
  const obj = JSON.parse(text);

  validateAndNormalize(obj);

  state = obj;
  loadedFileName = "data/catalogue.json (folder mode)";
  ui.selectedCategoryId = "__all__";
  ui.search = "";
  els.search.value = "";

  setEnabled(true);
  setDirty(false);
  render();

  els.fileNote.textContent = "Using folder: data/catalogue.json";
}

async function saveJsonToFolder() {
  if (!fs.folderHandle || !state) return;

  const dataDir = await fs.folderHandle.getDirectoryHandle("data", { create: true });
  const fileHandle = await dataDir.getFileHandle(fs.jsonFileName, { create: true });
  const writable = await fileHandle.createWritable();

  const clean = JSON.parse(JSON.stringify(state));
  delete clean._dirty;

  await writable.write(JSON.stringify(clean, null, 2));
  await writable.close();
}

/* Load JSON manually */
async function loadJsonFromPicker() {
  const file = els.filePicker.files?.[0];
  if (!file) return;

  try {
    const text = await file.text();
    const obj = JSON.parse(text);
    validateAndNormalize(obj);

    state = obj;
    loadedFileName = file.name;
    ui.selectedCategoryId = "__all__";
    ui.search = "";
    els.search.value = "";

    setEnabled(true);
    setDirty(false);
    render();
  } catch (e) {
    alert("Failed to load JSON: " + (e?.message || "Unknown error"));
  } finally {
    els.filePicker.value = "";
  }
}

/* State + UI */
function setEnabled(on) {
  els.btnSave.disabled = !on;
  els.search.disabled = !on;
  els.btnAddCategory.disabled = !on;
  els.btnAddProduct.disabled = !on;
}

function setDirty(isDirty) {
  if (!state) return;
  state._dirty = !!isDirty;

  els.fileNote.textContent =
    (loadedFileName ? `Loaded: ${loadedFileName}` : "Loaded.")
    + (isDirty ? "  •  Unsaved changes" : "");

  // autosave in folder mode
  if (isDirty && fs.folderHandle) {
    if (els.folderStatus) els.folderStatus.textContent = "Folder mode: ON (saving…)";
    requestAutosave();
  }
}

function validateAndNormalize(obj) {
  if (!obj || typeof obj !== "object") throw new Error("Invalid JSON format");

  obj.version = 1;
  obj.categories = Array.isArray(obj.categories) ? obj.categories : [];
  obj.products  = Array.isArray(obj.products)  ? obj.products  : [];

  for (const c of obj.categories) {
    if (!c.id) c.id = uid("cat");
    if (!c.name) c.name = "Unnamed";
  }

  for (const p of obj.products) {
    if (!p.id) p.id = uid("prod");
    if (!p.name) p.name = "Unnamed product";

    // categories (multi)
    if (Array.isArray(p.categoryIds)) {
      p.categoryIds = p.categoryIds.filter(Boolean);
    } else if (typeof p.categoryId === "string" && p.categoryId.trim()) {
      p.categoryIds = [p.categoryId.trim()];
    } else {
      p.categoryIds = [];
    }
    delete p.categoryId;

    // misc fields
    p.wordings = Array.isArray(p.wordings) ? p.wordings : [];
    p.codes    = Array.isArray(p.codes)    ? p.codes    : [];
    p.notes = p.notes || "";
    p.imageFileName = p.imageFileName || "";

    // CT size normalization
    p.ctSize = (p.ctSize !== undefined && p.ctSize !== null && p.ctSize > 0) ? Number(p.ctSize) : null;

    // Lots normalization (ordered-aware; see Fix 4)
    p.lots = normalizeLots(Array.isArray(p.lots) ? p.lots : []);
  }
}

function render() {
  renderCategories();
  renderFilterLine();
  renderProducts();
}

function renderCategories() {
  els.categoryList.innerHTML = "";

  // 1. System filters (Always shown)
  const specials = [
    { id: "__all__", name: "All products" },
    { id: "__in__", name: "✅ In stock (Usable)" },
    { id: "__low__", name: "⚠ Low stock (<10)" },
    { id: "__out__", name: "❌ Out of stock (0)" },
    { id: "__exp__", name: "⏳ Expiring / Expired" }
  ];

  for (const s of specials) {
    els.categoryList.appendChild(catRow(s, true));
  }

  // 2. Filter user categories
  const cats = (state.categories || [])
    .slice()
    .sort((a, b) => a.name.localeCompare(b.name))
    .filter(c => c.name.toLowerCase().includes(ui.catSearch));

  // --- POINT 3 START ---
  // Only add the line if we aren't searching OR if there are matching categories
  if (cats.length > 0) {
    const hr = document.createElement("div");
    hr.className = "catSeparator";
    els.categoryList.appendChild(hr);
  }
  // --- POINT 3 END ---

  // 3. Render the categories
  for (const c of cats) {
    els.categoryList.appendChild(catRow(c, false));
  }
  
  if (ui.catSearch && cats.length === 0) {
    const note = document.createElement("div");
    note.className = "smallNote";
    note.style.textAlign = "center";
    note.textContent = "No matches found";
    els.categoryList.appendChild(note);
  }
}

function catRow(cat, pseudo) {
  const row = document.createElement("div");
  row.className = "cat" + (ui.selectedCategoryId === cat.id ? " active" : "");

  const name = document.createElement("div");
  name.className = "catName";
  name.textContent = cat.name;

  const btns = document.createElement("div");
  btns.className = "catBtns";

  if (!pseudo) {
    const edit = document.createElement("button");
    edit.className = "btn small ghost";
    edit.textContent = "Edit";
    edit.addEventListener("click", (e) => {
      e.stopPropagation();
      openCategoryDlg(cat.id);
    });

    const del = document.createElement("button");
    del.className = "btn small ghost danger";
    del.textContent = "Del";
    del.addEventListener("click", (e) => {
      e.stopPropagation();
      deleteCategory(cat.id);
    });

    btns.append(edit, del);
  }

  row.append(name, btns);
  row.addEventListener("click", () => {
    ui.selectedCategoryId = cat.id;
    render();
  });

  return row;
}

function renderFilterLine() {
  let label = "";
  if (ui.selectedCategoryId === "__all__") label = "All products";
  else if (ui.selectedCategoryId === "__in__") label = "In stock (Usable)";
  else if (ui.selectedCategoryId === "__low__") label = "Low stock (1-10)";
  else if (ui.selectedCategoryId === "__out__") label = "Out of stock";
  else if (ui.selectedCategoryId === "__exp__") label = "Expiring / Expired";
  else label = state.categories.find(c => c.id === ui.selectedCategoryId)?.name || "Uncategorized";

  els.filterLine.textContent =
    `Filter: ${label}` + (ui.search ? `  •  Search: "${ui.search}"` : "");
}

function renderProducts() {
  const list = filteredProducts();
  els.productGrid.innerHTML = "";
  els.empty.hidden = list.length !== 0;

  for (const p of list) els.productGrid.appendChild(productCard(p));
}

function filteredProducts() {
  const q = ui.search.toLowerCase();

  return state.products
    .filter(p => {
      if (ui.selectedCategoryId === "__all__") return true;

      const t = totalStock(p);
      
      if (ui.selectedCategoryId === "__in__") {
        const total = totalStock(p);
        if (total <= 0) return false;

        // Check if there are any valid lots (not expired and not expiring)
        const validLots = normalizeLots(p.lots).filter(l => {
          if (l.qty <= 0) return false;
          const s = lotStatus(l.expiry);
          return s !== "expired" && s !== "expiring";
        });

        return validLots.length > 0;
      }

      if (ui.selectedCategoryId === "__low__") {
        return t > 0 && t < 10; // strictly between 1 and 9
      }

      if (ui.selectedCategoryId === "__out__") {
        return t === 0; // strictly 0
      }

      if (ui.selectedCategoryId === "__exp__") {
        const lots = normalizeLots(p.lots).filter(l => l.qty > 0);
        return lots.some(l => {
          const s = lotStatus(l.expiry);
          return s === "expiring" || s === "expired" || s === "risky";
        });
      }

      if (state.categories.some(c => c.id === ui.selectedCategoryId)) {
        return (p.categoryIds || []).includes(ui.selectedCategoryId);
      }

      return true;
    })
    .filter(p => {
      if (!q) return true;
      const hay = [p.name, ...(p.wordings || []), ...(p.codes || [])].join(" ").toLowerCase();
      // return fuzzyMatch(q, hay);   // <‑‑ fuzzy search (disabilitata, riattivabile)
	return hay.includes(q);         // <‑‑ ricerca normale
    })
    .sort((a, b) => a.name.localeCompare(b.name));
}

function productCard(p) {
  const node = els.cardTpl.content.cloneNode(true);
  const cardEl = node.querySelector(".card");

  const img = node.querySelector(".thumbImg");
  const empty = node.querySelector(".thumbEmpty");
  const imgPath = p.imageFileName ? `media/${p.imageFileName}` : "";

  if (imgPath) {
    img.src = imgPath;
    img.alt = p.name;
    img.style.display = "block";
    empty.style.display = "none";
    img.onerror = () => {
      img.style.display = "none";
      empty.style.display = "block";
      empty.textContent = "Image missing";
    };
  } else {
    img.style.display = "none";
    empty.style.display = "block";
    empty.textContent = "No image";
  }

  node.querySelector(".pName").textContent = p.name;

  // ✅ show multiple categories
  const names = (p.categoryIds || [])
    .map(id => state.categories.find(c => c.id === id)?.name)
    .filter(Boolean);

  node.querySelector(".pCat").textContent = names.length ? names.join(", ") : "Uncategorized";

  // --- Stock line (safe, hoisted dotsContainer; no early dots) ---
const total = totalStock(p);
const stockLine = node.querySelector(".stockLine");
const stockTotalEl = node.querySelector(".stockTotal");

// Hoist so later code (card-coloring/dots rules) can reuse it safely
let dotsContainer = null;

if (stockTotalEl) {
  stockTotalEl.textContent = `Stock: ${total}`;

  // Create the (empty) dot container, but DO NOT populate it here.
  dotsContainer = document.createElement("div");
  dotsContainer.className = "statusDots";

  // Append dots right next to the stock text
  stockTotalEl.style.display = "flex";
  stockTotalEl.style.alignItems = "center";
  stockTotalEl.appendChild(dotsContainer);
}

  const ctContainer = node.querySelector("#ctPillContainer");
  if (ctContainer) {
    ctContainer.innerHTML = ""; // Clear previous pills
    // Only show pills if there is stock AND a valid CT size
    if (total > 0 && p.ctSize && p.ctSize > 0) {
      const cts = calculateCTs(p);
      if (cts.full > 0) ctContainer.appendChild(createCtPill(cts.full, "full"));
      if (cts.partial > 0) ctContainer.appendChild(createCtPill(cts.partial, "partial"));
    }
  }

  // --- Card coloring logic ---
cardEl.classList.remove(
  "cardLow",
  "cardExpiring",
  "cardExpired",
  "cardRisky",
  "cardOut",
  "cardOrdered"
);

const realLots = (p.lots || []).filter(l => !l.ordered && l.qty > 0);
const orderedLotsOnly = (p.lots || []).filter(l => l.ordered && l.qty > 0);
const statusesReal = realLots.map(l => lotStatus(l.expiry));
const totalReal = realLots.reduce((s, l) => s + l.qty, 0);
const totalOrdered = orderedLotsOnly.reduce((s, l) => s + l.qty, 0);

const hasOrdered = totalOrdered > 0;

// CASE 1: ONLY ordered lots → FULL GREEN CARD
if (hasOrdered && totalReal === 0) {
  cardEl.classList.add("cardOrdered");
}

// CASE 2: Real stock exists → evaluate like normal
else if (totalReal > 0) {
  const hasExpired = statusesReal.includes("expired");
  const hasExpiring = statusesReal.includes("expiring");
  const hasRisky = statusesReal.includes("risky");
  const hasOk = statusesReal.includes("ok");

  if (hasOk) {
    if (totalReal < 10) cardEl.classList.add("cardLow");
  } else {
    if (hasExpired) cardEl.classList.add("cardExpired");
    else if (hasExpiring) cardEl.classList.add("cardExpiring");
    else if (hasRisky) cardEl.classList.add("cardRisky");
    else if (totalReal < 10) cardEl.classList.add("cardLow");
  }
}

// CASE 3: No stock at all (rare now)
else {
  cardEl.classList.add("cardOut");
}

// --- Status dots ---
dotsContainer.innerHTML = "";

// Read back card state
const isOut       = cardEl.classList.contains("cardOut");
const isExpired   = cardEl.classList.contains("cardExpired");
const isExpiring  = cardEl.classList.contains("cardExpiring");
const isRiskyCard = cardEl.classList.contains("cardRisky");
const isOrderedOnly = cardEl.classList.contains("cardOrdered");

// Real-stock statuses again (already calculated)
const hasExpiredReal   = statusesReal.includes("expired");
const hasExpiringReal  = statusesReal.includes("expiring");
const hasRiskyReal     = statusesReal.includes("risky");
const hasOkReal        = statusesReal.includes("ok");

let showExpired = false;
let showExpiring = false;
let showRisky = false;
let showOrderedDot = false;

// Dots rules:
// ❌ Out / Expired / Expiring / Ordered-only → NO dots
if (!isOut && !isExpired && !isExpiring && !isOrderedOnly) {

  if (isRiskyCard) {
    // Risky card → ONLY expiring dot
    if (hasExpiringReal) showExpiring = true;
  } 
  else {
    // Neutral card → all dots from real-stock lots
    if (hasExpiredReal)  showExpired = true;
    if (hasExpiringReal) showExpiring = true;
    if (hasRiskyReal)    showRisky = true;
  }

  // Ordered dot → only when:
  // real stock > 0 AND ordered lots exist
  if (totalReal > 0 && hasOrdered) {
    showOrderedDot = true;
  }
}

// Render dots
if (showExpired) {
  const d = document.createElement("span");
  d.className = "dot expired";
  d.title = "Has expired lots";
  dotsContainer.appendChild(d);
}
if (showExpiring) {
  const d = document.createElement("span");
  d.className = "dot expiring";
  d.title = "Has expiring lots";
  dotsContainer.appendChild(d);
}
if (showRisky) {
  const d = document.createElement("span");
  d.className = "dot risky";
  d.title = "Has risky lots";
  dotsContainer.appendChild(d);
}
if (showOrderedDot) {
  const d = document.createElement("span");
  d.className = "dot ordered";
  d.title = "Ordered stock pending";
  dotsContainer.appendChild(d);
}

  node.querySelector('[data-act="info"]').addEventListener("click", () => openInfoDlg(p.id));
  node.querySelector('[data-act="stock"]').addEventListener("click", () => openStockDlg(p.id));
  node.querySelector('[data-act="edit"]').addEventListener("click", () => openProductDlg(p.id));
  node.querySelector('[data-act="del"]').addEventListener("click", () => {
    if (!confirm(`Delete "${p.name}"?`)) return;
    state.products = state.products.filter(x => x.id !== p.id);
    setDirty(true);
    render();
  });

  return node;
}

/* Dialogs */
function openCategoryDlg(id) {
  ui.editingCategoryId = id;

  if (id) {
    const c = state.categories.find(x => x.id === id);
    if (!c) return;
    els.catTitle.textContent = "Edit category";
    els.catName.value = c.name;
  } else {
    els.catTitle.textContent = "Add category";
    els.catName.value = "";
  }

  els.catDlg.showModal();
  setTimeout(() => els.catName.focus(), 50);
}

function deleteCategory(id) {
  const c = state.categories.find(x => x.id === id);
  if (!c) return;

  const used = state.products.some(p =>
    (p.categoryIds || []).includes(id)
  );

  const msg = used
    ? `Delete category "${c.name}"? Products in it will become Uncategorized.`
    : `Delete category "${c.name}"?`;

  if (!confirm(msg)) return;

  // Remove category from list
  state.categories = state.categories.filter(x => x.id !== id);

  // Remove category reference from all products
  for (const p of state.products) {
    p.categoryIds = (p.categoryIds || []).filter(cid => cid !== id);
  }

  // Reset filter if needed
  if (ui.selectedCategoryId === id) {
    ui.selectedCategoryId = "__all__";
  }

  setDirty(true);
  render();
}

function openProductDlg(id) {
  ui.editingProductId = id;

  // Build checkbox list
  els.prodCategoriesBox.innerHTML = "";
  const cats = state.categories.slice().sort((a,b)=>a.name.localeCompare(b.name));

  for (const c of cats) {
    const label = document.createElement("label");
    label.className = "checkItem";

    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.value = c.id;

    const text = document.createElement("span");
    text.textContent = c.name;

    label.append(cb, text);
    els.prodCategoriesBox.appendChild(label);
  }

  if (id) {
    const p = state.products.find(x => x.id === id);
    if (!p) return;

    els.prodTitle.textContent = "Edit product";
    els.prodName.value = p.name;
    $("#prodCtSize").value = p.ctSize || "";

    // Back-compat: support older single categoryId
    const selectedIds = Array.isArray(p.categoryIds)
      ? p.categoryIds
      : (p.categoryId ? [p.categoryId] : []);

    // check relevant boxes
    for (const cb of els.prodCategoriesBox.querySelectorAll('input[type="checkbox"]')) {
      cb.checked = selectedIds.includes(cb.value);
    }

    els.prodWordings.value = (p.wordings || []).join("\n");
    els.prodCodes.value = (p.codes || []).join("\n");
    els.prodNotes.value = p.notes || "";
    els.prodImageFileName.value = p.imageFileName || "";
    renderProductPreview(p);

  } else {
    els.prodTitle.textContent = "Add product";
    els.prodName.value = "";
    $("#prodCtSize").value = "";

    // optionally default-check current real category filter
    if (state.categories.some(c => c.id === ui.selectedCategoryId)) {
      const cb = els.prodCategoriesBox.querySelector(`input[value="${ui.selectedCategoryId}"]`);
      if (cb) cb.checked = true;
    }

    els.prodWordings.value = "";
    els.prodCodes.value = "";
    els.prodNotes.value = "";
    els.prodImageFileName.value = "";
    renderProductPreview(null);
  }

  els.prodDlg.showModal();
  setTimeout(() => els.prodName.focus(), 50);
}

function renderProductPreview(p) {
  els.imgPreview.innerHTML = "";
  const fileName = (p?.imageFileName || "").trim();

  if (!fileName) {
    els.imgPreview.textContent = "No image";
    return;
  }

  const img = document.createElement("img");
  img.src = `media/${fileName}`;
  img.alt = "Image";
  img.onerror = () => {
    els.imgPreview.textContent = "Image missing (check media/ + filename)";
  };

  els.imgPreview.appendChild(img);
}

function renderLotList(p) {
  const lotList = document.getElementById("lotList");
  lotList.innerHTML = "";

  const lots = (p.lots || []).slice();

  if (!lots.length) {
    lotList.innerHTML = `<div class="smallNote">No lots.</div>`;
    return;
  }

  for (const l of lots) {
    const row = document.createElement("div");
    row.className = "lotRow";

    const isOrdered = !!l.ordered;

    row.innerHTML = `
      <div class="lotMeta">
        <div class="lotExp">${formatDateDMY(l.expiry)}</div>
        <div class="lotQty">Qty: ${l.qty}</div>
      </div>
    `;

    const btns = document.createElement("div");

    // DELETE BUTTON
    const del = document.createElement("button");
    del.className = "btn small ghost danger";
    del.textContent = "Delete";

    del.onclick = (e) => {
      e.preventDefault();

      // Remove this lot
      p.lots = (p.lots || []).filter(x => x !== l);

      setDirty(true);
      render();            // refresh product cards
      renderLotList(p);    // refresh dialog
    };

    if (isOrdered) {
      // RECEIVED BUTTON
      const received = document.createElement("button");
      received.className = "btn small received";
      received.textContent = "Received";

      received.onclick = (e) => {
        e.preventDefault();

        // Remove the ordered lot
        p.lots = (p.lots || []).filter(x => x !== l);

        // Call the **exact same Add logic** as the Add button
        // Use product ID, expiry, and qty
        addLotViaUI(p.id, l.expiry, Number(l.qty));

        setDirty(true);
        render();            // refresh product cards
        renderLotList(p);    // refresh dialog
      };

      btns.append(received, del);
    } else {
      // NORMAL LOT: + - Delete
      const plus = document.createElement("button");
      plus.className = "btn small ghost";
      plus.textContent = "+";
      plus.onclick = (e) => {
        e.preventDefault();
        l.qty++;
        setDirty(true);
        render();
        renderLotList(p);
      };

      const minus = document.createElement("button");
      minus.className = "btn small ghost";
      minus.textContent = "-";
      minus.onclick = (e) => {
        e.preventDefault();
        if (l.qty > 1) l.qty--;
        setDirty(true);
        render();
        renderLotList(p);
      };

      btns.append(plus, minus, del);
    }

    row.appendChild(btns);
    lotList.appendChild(row);
  }
}

// helper to ensure all lots are created consistently
function createNormalLot(expiry, qty) {
  return { expiry, qty };
}

/* Stock */
function openStockDlg(productId) {
  ui.stockProductId = productId;

  const p = state.products.find((x) => x.id === productId);
  if (!p) return;

  els.stockMsg.textContent = "";
  els.stockExpiry.value = "";
  els.stockQty.value = "1";

  // Show CT option only if product has CT size
  if (p.ctSize && p.ctSize > 0) {
    $("#stockUnitSelector").style.display = "block";
  } else {
    $("#stockUnitSelector").value = "units";
    $("#stockUnitSelector").style.display = "none";
  }

  attachDMYMask(els.stockExpiry);

  // ✅ Update header / title if you have one
  if (els.stockTitle) {
    els.stockTitle.textContent = `Stock — ${p.name}`;
  }
  
  renderLots(p);

  els.stockDlg.showModal();
}

function renderLots(p) {
  els.lotList.innerHTML = "";

  const total = totalStock(p);
  let stockText = `${p.name} — Total: ${total}`;
  if (total > 0 && p.ctSize) {
    const cts = calculateCTs(p);
    stockText += ` (${cts.full} Full CT, ${cts.partial} Non-empty CT)`;
  }
  els.stockHeader.textContent = stockText;

  const lots = normalizeLots(p.lots).sort((a, b) => a.expiry.localeCompare(b.expiry));
  const displayLots = lots.filter(l => l.qty > 0);

  if (displayLots.length === 0) {
    const d = document.createElement("div");
    d.className = "smallNote";
    d.textContent = "No stock yet. Add with an expiry date.";
    els.lotList.appendChild(d);
    return;
  }

  for (const l of displayLots) {
    const isOrdered = !!l.ordered;
    const st = lotStatus(l.expiry);

    const row = document.createElement("div");
    let rowClass = "lotRow";
    if (isOrdered) rowClass += " lotOrdered";
    if (!isOrdered) {
      if (st === "expired") rowClass += " lotExpired";
      else if (st === "expiring") rowClass += " lotExpiring";
      else if (st === "risky") rowClass += " lotRisky";
    }
    row.className = rowClass;

    // LEFT SIDE
    const meta = document.createElement("div");
    meta.className = "lotMeta";

    const main = document.createElement("div");
    main.className = "lotMain";

    const qtyLine = document.createElement("div");
    qtyLine.className = "lotQtyBig";
    qtyLine.textContent = `Qty: ${l.qty}`;
    if (p.ctSize) {
      const full = Math.floor(l.qty / p.ctSize);
      const rem = l.qty % p.ctSize;
      if (full > 0) qtyLine.appendChild(createCtPill(full, "full"));
      if (rem > 0) qtyLine.appendChild(createCtPill(1, "partial"));
    }

    const dateLine = document.createElement("div");
    dateLine.className = "lotDateSmall";
    dateLine.textContent = formatDateDMY(l.expiry);

    const tag = document.createElement("span");
    tag.className = "lotTag";
    tag.textContent = isOrdered ? "ORDERED" : st.toUpperCase();

    dateLine.appendChild(tag);
    main.append(qtyLine, dateLine);
    meta.appendChild(main);

    // BUTTONS
    const btns = document.createElement("div");
    btns.style.display = "flex";
    btns.style.gap = "6px";

    // DEL BUTTON FIRST!
    const del = document.createElement("button");
    del.className = "btn small danger";
    del.textContent = "DEL";
    del.onclick = (e) => {
  e.preventDefault();

  const expiry = l.expiry;
  const isOrdered = !!l.ordered;

  p.lots = (p.lots || []).filter(x =>
    !(x.expiry === expiry && !!x.ordered === isOrdered)
  );

  setDirty(true);
  render();
  openStockDlg(p.id);
};

    if (isOrdered) {
  const received = document.createElement("button");
  received.className = "btn small received";
  received.textContent = "Received";
  received.onclick = (e) => {
    e.preventDefault();
    receiveOrderedLot(p, l);
  };
  btns.append(received, del);
} else {
      const plus = document.createElement("button");
      plus.className = "btn small";
      plus.textContent = "+";
      plus.onclick = (e) => {
        e.preventDefault();
        els.stockExpiry.value = formatDateDMY(l.expiry);
        els.stockQty.value = "1";
        adjustStock(+1);
      };

      const minus = document.createElement("button");
      minus.className = "btn small danger";
      minus.textContent = "-";
      minus.onclick = (e) => {
        e.preventDefault();
        els.stockExpiry.value = formatDateDMY(l.expiry);
        els.stockQty.value = "1";
        adjustStock(-1);
      };

      btns.append(plus, minus, del);
    }

    row.append(meta, btns);
    els.lotList.appendChild(row);
  }
}

function receiveOrderedLot(p, lot) {
  if (!p || !lot) return;

  const expiryISO = lot.expiry;
  const qty = Number(lot.qty || 0);

  // 1️⃣ Remove ordered lot
  p.lots = (p.lots || []).filter(l =>
    !(l.ordered && l.expiry === expiryISO)
  );

  // 2️⃣ Inject values into stock UI
  els.stockExpiry.value = formatDateDMY(expiryISO);
  els.stockQty.value = String(qty);

  // Force units mode (ordered lots already stored in units)
  const unitSel = $("#stockUnitSelector");
  if (unitSel) unitSel.value = "units";

  // 3️⃣ Use your existing stock logic
  const prevMode = unitSel?.value;
if (unitSel) unitSel.value = "units";

adjustStock(+1);

if (unitSel) unitSel.value = prevMode || "units";

  els.stockMsg.textContent =
    `Received ${qty} for ${formatDateDMY(expiryISO)}.`;
}

function adjustStock(dir) {
  const p = state.products.find(x => x.id === ui.stockProductId);
  if (!p) return;

  // --- Read expiry field with DMY mask already applied ---
  const expiryText = (els.stockExpiry.value || "").trim();
  let qty = Number(els.stockQty.value);

  // --- Determine real units (CT → units conversion) ---
  const unitMode = $("#stockUnitSelector").value;
  if (unitMode === "ct") {
    if (!p.ctSize || p.ctSize <= 0) {
      els.stockMsg.textContent = "This product has no CT size defined.";
      return;
    }
    qty = qty * p.ctSize;
  }

  // --- Convert DMY to ISO ---
  let expiryISO;

if (els.stockUnknownExpiry.checked) {
  expiryISO = "__unknown__";
} else {
  expiryISO = parseDMYToISO(expiryText);
  if (!expiryISO) {
    els.stockMsg.textContent =
      "Enter a valid date as DD/MM/YYYY (e.g. 08/02/2026).";
    return;
  }
}

  // --- Validate qty ---
  if (!Number.isInteger(qty) || qty <= 0) {
    els.stockMsg.textContent = "Quantity must be a whole number ≥ 1.";
    return;
  }

  // --- Normalize lots but KEEP ordered flag ---
  p.lots = normalizeLots(p.lots);

  // Only modify *real* lots, never ordered ones:
  let lot = p.lots.find(l => l.expiry === expiryISO && !l.ordered);

  // --- ADD ---
  if (dir > 0) {
    if (lot) {
      lot.qty += qty;
    } else {
      p.lots.push({ expiry: expiryISO, qty });
    }

    els.stockMsg.textContent =
      `Added ${qty} to ${formatDateDMY(expiryISO)}.`;
  }

  // --- WITHDRAW ---
  else {
    // Cannot withdraw from nonexistent real lot
    if (!lot) {
      els.stockMsg.textContent =
        `No lot for ${formatDateDMY(expiryISO)}.`;
      return;
    }
    if (lot.qty < qty) {
      els.stockMsg.textContent =
        `Cannot withdraw ${qty}. Only ${lot.qty} available for ${formatDateDMY(expiryISO)}.`;
      return;
    }

    lot.qty -= qty;
    els.stockMsg.textContent =
      `Withdrew ${qty} from ${formatDateDMY(expiryISO)}.`;
  }

  // --- Remove zero-qty lots & re‑normalize (keeps ordered flags) ---
  p.lots = normalizeLots(p.lots).filter(l => l.qty > 0);

  // --- Update UI ---
  setDirty(true);
  els.stockHeader.textContent =
    `${p.name} — Total: ${totalStock(p)}`;

  renderLots(p);
  render();
}

/* More info */
function openInfoDlg(productId) {
  const p = state.products.find(x => x.id === productId);
  if (!p) return;

  els.infoTitle.textContent = `More info: ${p.name}`;

  const catNames = (Array.isArray(p.categoryIds) ? p.categoryIds : [])
    .map(id => state.categories.find(c => c.id === id)?.name)
    .filter(Boolean);

  let ctInfo = "";
  if (p.ctSize && p.ctSize > 0) ctInfo = ` • CT Size: ${p.ctSize} units`;

  els.infoSub.textContent = `Category: ${catNames.length ? catNames.join(", ") : "Uncategorized"}${ctInfo}`;

  els.infoWordings.textContent = (p.wordings?.length ? p.wordings.join("\n") : "—");
  els.infoCodes.textContent = (p.codes?.length ? p.codes.join("\n") : "—");

  els.infoLots.innerHTML = "";
  const lots = normalizeLots(p.lots)
    .filter(l => l.qty > 0)
    .sort((a, b) => a.expiry.localeCompare(b.expiry));

  if (lots.length === 0) {
    els.infoLots.textContent = "No lots.";
  } else {
    for (const l of lots) {
    const st = lotStatus(l.expiry);

    const row = document.createElement("div");
    row.className = "lotItem";
    if (st === "expired") row.classList.add("lotExpired");
    else if (st === "expiring") row.classList.add("lotExpiring");
    else if (st === "risky") row.classList.add("lotRisky");

    const left = document.createElement("div");
    left.className = "lotMain";

    // --- Qty Line + CT pills ---
    const qtyLine = document.createElement("div");
    qtyLine.className = "lotQtyBig";
    qtyLine.style.display = "flex";
    qtyLine.style.alignItems = "center";
    qtyLine.style.gap = "8px";
    qtyLine.textContent = `Qty: ${l.qty}`;

    if (p.ctSize && p.ctSize > 0) {
        const full = Math.floor(l.qty / p.ctSize);
        const rem = l.qty % p.ctSize;

        if (full > 0) qtyLine.appendChild(createCtPill(full, "full"));
        if (rem > 0) qtyLine.appendChild(createCtPill(1, "partial"));
    }

    // --- Date + Status tag ---
    const dateLine = document.createElement("div");
    dateLine.className = "lotDateSmall";

    const dateText = document.createElement("span");
    dateText.textContent = formatDateDMY(l.expiry);

    const tag = document.createElement("span");
    tag.className = "lotTag";

    if (st === "expired") tag.textContent = "EXPIRED";
    else if (st === "expiring") tag.textContent = "EXPIRING";
    else if (st === "risky") tag.textContent = "RISKY";
    else tag.textContent = "OK";

    dateLine.append(dateText, tag);

    left.append(qtyLine, dateLine);

    // append row
    row.append(left, document.createElement("div"));
    els.infoLots.appendChild(row);
}
  }

  els.infoDlg.showModal();
}

/* Helpers */
function uid(prefix) {
  return `${prefix}_${Math.random().toString(16).slice(2)}_${Date.now().toString(16)}`;
}

function splitLines(text) {
  return (text || "").split(/\r?\n/).map(s => s.trim()).filter(Boolean);
}

function normalizeLots(lots) {
  const map = new Map(); // key = `${expiry}|${ordered?1:0}`

  for (const l of (lots || [])) {
    const expiry = String(l.expiry || "").trim();
    if (!expiry) continue;

    const qty = Math.max(0, Math.trunc(Number(l.qty) || 0));
    const ordered = !!l.ordered;

    const key = `${expiry}|${ordered ? 1 : 0}`;

    const prev = map.get(key);
    map.set(key, {
      expiry,
      qty: (prev?.qty || 0) + qty,
      ordered
    });
  }

  return [...map.values()];
}

function totalStock(p) {
  return normalizeLots(p.lots).reduce((s, l) => s + l.qty, 0);
}

function todayISO() {
  const d = new Date();
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const da = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${da}`;
}

function ym(iso) {
  return (iso || "").slice(0, 7);
}

function addMonthsYM(ymStr, monthsToAdd) {
  const [y0, m0] = ymStr.split("-").map(Number);
  let y = y0, m = m0 + monthsToAdd;
  while (m > 12) { m -= 12; y += 1; }
  while (m < 1) { m += 12; y -= 1; }
  return `${y}-${String(m).padStart(2, "0")}`;
}

function lotStatus(expiryISO) {
  if (!expiryISO || expiryISO === "__unknown__") {
    return "ok"; // unknown treated as neutral
  }

  const now = new Date();
  const exp = new Date(expiryISO);
  if (exp < now) return "expired";

  const diffMonths =
    (exp.getFullYear() - now.getFullYear()) * 12 +
    (exp.getMonth() - now.getMonth());

  if (diffMonths <= 3) return "expiring";
  if (diffMonths <= 4) return "risky";
  return "ok";
}

function duplicateProductFromDialog() {
  if (!state) return;

  // Read current fields from the dialog (even if user hasn't pressed Save)
  const name = (els.prodName.value || "").trim();
  if (!name) {
    alert("Give the product a name first (then duplicate).");
    return;
  }

  // Multiple categories checkbox list (prodCategoriesBox)
  const selectedCategoryIds = els.prodCategoriesBox
    ? Array.from(els.prodCategoriesBox.querySelectorAll('input[type="checkbox"]:checked')).map(cb => cb.value)
    : []; // fallback

  const wordings = splitLines(els.prodWordings.value);
  const codes = splitLines(els.prodCodes.value);
  const imageFileName = (els.prodImageFileName.value || "").trim();

  // Find the original (if editing an existing product)
  const original = ui.editingProductId
    ? state.products.find(p => p.id === ui.editingProductId)
    : null;

  // Decide what to clone:
  // - keep categories/wordings/codes/image
  // - DO NOT copy lots (new product should start at 0 stock)
  // - optionally copy notes if you still use it
  const newProduct = {
    id: uid("prod"),
    name: name,
    categoryIds: selectedCategoryIds,
    wordings: wordings,
    codes: codes,
    notes: (els.prodNotes?.value || "").trim(), // harmless if notes is unused
    imageFileName: imageFileName,
    lots: [], // ✅ start empty stock for the clone
  };

  // If you want to preserve some extra hidden fields from original in the future:
  // (none currently, but leaving hook)
  if (original) {
    // Example: if later you add vendor fields, etc.
  }

  state.products ||= [];
  state.products.push(newProduct);

  setDirty(true);

  // Close current dialog and immediately open edit for the new clone
  ui.editingProductId = null;
  els.prodDlg.close();

  // Open edit dialog for the cloned product
  openProductDlg(newProduct.id);
}

function parseDMYToISO(dmy) {
  const s = (dmy || "").trim();

  // Match DD/MM or DD/MM/YYYY
  const m = /^(\d{1,2})\/(\d{1,2})(?:\/(\d{4}))?$/.exec(s);
  if (!m) return null;

  const dd = Number(m[1]);
  const mm = Number(m[2]);

  // If year missing → use current year
  const yyyy = m[3] ? Number(m[3]) : new Date().getFullYear();

  if (mm < 1 || mm > 12) return null;
  if (dd < 1 || dd > 31) return null;

  const iso = `${String(yyyy).padStart(4,"0")}-${String(mm).padStart(2,"0")}-${String(dd).padStart(2,"0")}`;
  const test = new Date(`${iso}T00:00:00`);

  // Strict real-date validation (no 31 Feb nonsense)
  if (Number.isNaN(test.getTime())) return null;
  if (
    test.getFullYear() !== yyyy ||
    test.getMonth() + 1 !== mm ||
    test.getDate() !== dd
  ) return null;

  return iso;
}

function formatDateDMY(iso) {
  if (iso === "__unknown__") return "Unknown";
  if (!iso || typeof iso !== "string" || iso.length < 10) return iso || "";
  const [y, m, d] = iso.split("-");
  return `${d}/${m}/${y}`;
}

function attachDMYMask(input) {
  input.addEventListener("input", (e) => {
    const start = input.selectionStart;
    const oldValue = input.value;

    // Only digits
    let digits = oldValue.replace(/\D/g, "").slice(0, 8);

    let formatted = "";
    if (digits.length >= 1) formatted += digits.slice(0, 2);
    if (digits.length >= 3) formatted += "/" + digits.slice(2, 4);
    if (digits.length >= 5) formatted += "/" + digits.slice(4, 8);

    // If less than full segments
    if (digits.length <= 2) formatted = digits;
    else if (digits.length <= 4) formatted = digits.slice(0,2) + "/" + digits.slice(2);
    else formatted = digits.slice(0,2) + "/" + digits.slice(2,4) + "/" + digits.slice(4);

    const diff = formatted.length - oldValue.length;

    input.value = formatted;

    // Restore cursor properly
    input.setSelectionRange(start + diff, start + diff);
  });
}

function fuzzyMatch(query, text) {
  query = (query || "").toLowerCase().trim();
  if (!query) return true;

  const tokens = query.split(/\s+/).filter(Boolean);
  for (const t of tokens) {
    if (!fuzzyTokenMatch(t, text)) return false;
  }
  return true;
}

function fuzzyTokenMatch(token, text) {
  if (text.includes(token)) return true;
  let i = 0;
  for (let j = 0; j < text.length && i < token.length; j++) {
    if (text[j] === token[i]) i++;
  }
  return i === token.length;
}

function downloadJSON(obj, filename) {
  const clean = JSON.parse(JSON.stringify(obj));
  delete clean._dirty;

  const blob = new Blob([JSON.stringify(clean, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);

  setDirty(false);
  alert(
    "Saved as a download.\n\n" +
    "Now replace data/catalogue.json with the downloaded catalogue.json.\n" +
    "(Or save directly into the data folder if your browser prompts.)"
  );
}

function createOrder() {
  const p = state.products.find(x => x.id === ui.stockProductId);
  if (!p) return;

  const expiryText = (els.stockExpiry.value || "").trim();
const qtyRaw = Number(els.stockQty.value);

let expiryISO;

if (els.stockUnknownExpiry.checked) {
  expiryISO = "__unknown__";
} else {
  expiryISO = parseDMYToISO(expiryText);
  if (!expiryISO) {
    els.stockMsg.textContent = "Enter a valid date (DD/MM/YYYY).";
    return;
  }
}

  if (!Number.isInteger(qtyRaw) || qtyRaw <= 0) {
    els.stockMsg.textContent = "Quantity must be whole number ≥ 1.";
    return;
  }

  let qty = qtyRaw;
  const mode = $("#stockUnitSelector")?.value;

  if (mode === "ct") {
    if (!p.ctSize || p.ctSize <= 0) {
      els.stockMsg.textContent = "CT mode not available.";
      return;
    }
    qty = qty * p.ctSize;
  }

  p.lots.push({
    expiry: expiryISO,
    qty,
    ordered: true
  });

  els.stockMsg.textContent = `Ordered ${qty} for ${formatDateDMY(expiryISO)}.`;

  setDirty(true);
  renderLots(p);
  render();
}

/* Image copy (folder mode) */
function sanitizeFileName(name) {
  const s = (name || "").trim();
  if (!s) return "image";
  return s.replace(/[\\/:*?\"<>|]+/g, "_");
}

function splitBaseExt(name) {
  const n = name || "";
  const i = n.lastIndexOf(".");
  if (i <= 0) return { base: n, ext: "" };
  return { base: n.slice(0, i), ext: n.slice(i) };
}

async function copyPickedImageToMedia(file) {
  if (!fs.folderHandle) throw new Error("Folder mode not enabled.");

  const mediaDir = await fs.folderHandle.getDirectoryHandle("media", { create: true });

  const original = sanitizeFileName(file.name || "image");
  const { base, ext } = splitBaseExt(original);

  let candidate = `${base}${ext}`;
  for (let n = 1; n < 500; n++) {
    try {
      await mediaDir.getFileHandle(candidate, { create: false });
      candidate = `${base}_${n}${ext}`;
    } catch {
      break;
    }
  }

  const outHandle = await mediaDir.getFileHandle(candidate, { create: true });
  const writable = await outHandle.createWritable();
  await writable.write(await file.arrayBuffer());
  await writable.close();

  return candidate;
}

function calculateCTs(p) {
  let full = 0;
  let partial = 0;
  if (!p.ctSize || p.ctSize <= 0) return { full, partial };

  (p.lots || []).forEach(l => {
    const q = parseInt(l.qty) || 0;
    if (q <= 0) return;
    full += Math.floor(q / p.ctSize);
    if (q % p.ctSize > 0) partial += 1;
  });

  return { full, partial };
}

function createCtPill(count, type) {
  const span = document.createElement("span");
  // Use the classes defined in your CSS
  span.className = `ct-pill ${type === 'full' ? 'ct-pill-full' : 'ct-pill-partial'}`;
  span.textContent = `${count} CT`;
  return span;
}

wire();