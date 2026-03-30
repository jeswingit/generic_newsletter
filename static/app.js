/* ═══════════════════════════════════════════════════════════
   Aon Newsletter Builder — app.js
═══════════════════════════════════════════════════════════ */

// ─── State ────────────────────────────────────────────────
const state = {
  config: {
    meta: { newsletterName: "", subject: "", from: "", to: "" },
    theme: {
      primaryColor: "#E31837",
      backgroundColor: "#F5F5F5",
      fontFamily: "Arial, Helvetica, sans-serif",
      tableWidth: 700,
    },
    sections: [],
  },
  selectedId: null,
  pendingImageSectionId: null,
  pendingImagePropKey: null,
  sortable: null,
};

// ─── Section defaults ─────────────────────────────────────
const SECTION_LABELS = {
  header:       "Header",
  footer:       "Footer",
  bullet_list:  "Bullet List",
  event_list:   "Event / Save the Date",
  text_block:   "Text Block",
  product_card: "Product Card",
  image_block:  "Image",
  divider:      "Divider",
};

const SECTION_DEFAULTS = {
  header: {
    orgName: "Aon", tagline: "Newsletter", logoUrl: "",
    backgroundColor: "#E31837", textColor: "#ffffff",
  },
  footer: {
    orgName: "Aon", year: String(new Date().getFullYear()),
    backgroundColor: "#1A1A1A", textColor: "#aaaaaa",
  },
  bullet_list: {
    heading: "What's Going On",
    items: ["Item 1", "Item 2"],
    // bulletColor intentionally omitted — inherits theme.primaryColor via renderer
    backgroundColor: "#ffffff",
  },
  event_list: {
    heading: "Save the Date!",
    items: ["Event 1"],
    backgroundColor: "#EEF6F7",
  },
  text_block: {
    heading: "Section Title", body: "Body text.",
    backgroundColor: "#ffffff",
  },
  product_card: {
    title: "Product Name", body: "Description.", creator: "",
    imageUrl: "", backgroundColor: "#ffffff",
  },
  image_block: {
    imageUrl: "", altText: "", caption: "",
    backgroundColor: "#ffffff",
  },
  divider: { color: "#E0E0E0", spacing: 20 },
};

// ─── Aon brand colors ─────────────────────────────────────
const AON_COLORS = [
  { label: "Signature Red", value: "#EB0017" },
  { label: "Navy 01",       value: "#262836" },
  { label: "Gray 01",       value: "#46535E" },
  { label: "Gray 02",       value: "#5D6D78" },
  { label: "Gray 03",       value: "#82939A" },
  { label: "Gray 04",       value: "#ACC0C4" },
  { label: "Gray 05",       value: "#CDDBDE" },
  { label: "Gray 06",       value: "#E5EFF0" },
  { label: "Gray 07",       value: "#EEF6F7" },
  { label: "Gray 08",       value: "#F9FCFC" },
  { label: "White",         value: "#FFFFFF" },
];

// ═══════════════════════════════════════════════════════════
// INIT
// ═══════════════════════════════════════════════════════════
document.addEventListener("DOMContentLoaded", () => {
  loadState();
  bindMetaInputs();
  bindThemeInputs();
  bindSectionLibrary();
  bindTopBarActions();
  bindExportDropdown();
  bindFileInputs();
  renderCanvas();
  renderEmpty();
});

// ─── Bind left-panel meta fields ─────────────────────────
function bindMetaInputs() {
  const fields = ["meta-name", "meta-subject", "meta-from", "meta-to"];
  const keys   = ["newsletterName", "subject", "from", "to"];
  fields.forEach((id, i) => {
    const el = document.getElementById(id);
    if (!el) return;
    el.value = state.config.meta[keys[i]] || "";
    el.addEventListener("input", () => {
      state.config.meta[keys[i]] = el.value;
      saveState();
    });
  });
}

// ─── Bind theme controls ──────────────────────────────────
function bindThemeInputs() {
  bindColorPicker("theme-bg", "swatch-bg", "hex-bg", "backgroundColor");

  const widthEl = document.getElementById("theme-width");
  if (widthEl) {
    widthEl.value = state.config.theme.tableWidth;
    widthEl.addEventListener("input", () => {
      const v = parseInt(widthEl.value, 10);
      if (v >= 200) {
        state.config.theme.tableWidth = v;
        saveState();
      }
    });
  }
}

function bindColorPicker(inputId, swatchId, hexId, themeKey) {
  const input  = document.getElementById(inputId);
  const swatch = document.getElementById(swatchId);
  const hex    = document.getElementById(hexId);
  if (!input) return;

  const update = (val) => {
    const old = state.config.theme[themeKey];
    swatch.style.background = val;
    hex.textContent = val;
    state.config.theme[themeKey] = val;

    // When primary color changes, clear explicit bulletColor overrides that
    // matched the old primary so sections inherit the new theme color
    if (themeKey === "primaryColor") {
      for (const section of state.config.sections) {
        if (!section.props.bulletColor || section.props.bulletColor === old) {
          delete section.props.bulletColor;
        }
      }
    }
    saveState();
  };

  update(state.config.theme[themeKey] || input.value);
  swatch.addEventListener("click", () => input.click());
  input.addEventListener("input", () => update(input.value));
}

// ─── Section library ──────────────────────────────────────
function bindSectionLibrary() {
  // Click to add
  document.querySelectorAll(".section-type-btn").forEach(btn => {
    btn.addEventListener("click", () => addSection(btn.dataset.type));
  });

  // Drag from sidebar onto canvas
  Sortable.create(document.getElementById("section-library"), {
    group:     { name: "palette", pull: "clone", put: false },
    sort:      false,
    animation: 150,
    ghostClass: "sortable-ghost",
    // Cloned element dropped on canvas is handled by canvas onAdd
  });
}

// ─── Top bar buttons ──────────────────────────────────────
function bindTopBarActions() {
  document.getElementById("btn-preview")?.addEventListener("click", handlePreview);
  document.getElementById("btn-import-excel")?.addEventListener("click", () => {
    document.getElementById("import-excel-input").click();
  });
}

// ─── Export dropdown ──────────────────────────────────────
function bindExportDropdown() {
  const btn  = document.getElementById("btn-export");
  const menu = document.getElementById("export-menu");
  const dd   = btn?.closest(".dropdown");

  btn?.addEventListener("click", (e) => {
    e.stopPropagation();
    dd.classList.toggle("open");
  });

  menu?.querySelectorAll(".dropdown-item").forEach(item => {
    item.addEventListener("click", () => {
      dd.classList.remove("open");
      handleExport(item.dataset.format);
    });
  });

  document.addEventListener("click", () => dd?.classList.remove("open"));
}

// ─── File inputs ──────────────────────────────────────────
function bindFileInputs() {
  document.getElementById("import-excel-input")?.addEventListener("change", (e) => {
    const file = e.target.files[0];
    if (file) handleImportExcel(file);
    e.target.value = "";
  });

  document.getElementById("image-upload-input")?.addEventListener("change", (e) => {
    const file = e.target.files[0];
    if (file) handleImageUpload(file);
    e.target.value = "";
  });
}

// ═══════════════════════════════════════════════════════════
// SECTION CRUD
// ═══════════════════════════════════════════════════════════

function addSection(type) {
  const defaults = SECTION_DEFAULTS[type] || {};
  const newSection = {
    id: crypto.randomUUID(),
    type,
    props: { ...defaults, items: defaults.items ? [...defaults.items] : undefined },
  };
  // Remove undefined items key if not applicable
  if (newSection.props.items === undefined) delete newSection.props.items;

  state.config.sections.push(newSection);
  saveState();
  renderCanvas();
  selectSection(newSection.id);
  renderEmpty();
  // Scroll to bottom of canvas
  setTimeout(() => {
    const last = document.querySelector("#canvas-list .section-card:last-child");
    last?.scrollIntoView({ behavior: "smooth", block: "nearest" });
  }, 50);
}

function deleteSection(id) {
  state.config.sections = state.config.sections.filter(s => s.id !== id);
  if (state.selectedId === id) {
    state.selectedId = null;
    renderPropertiesPanel();
  }
  saveState();
  renderCanvas();
  renderEmpty();
}

function selectSection(id) {
  state.selectedId = id;
  document.querySelectorAll(".section-card").forEach(card => {
    card.classList.toggle("selected", card.dataset.id === id);
  });
  renderPropertiesPanel();
}

function getSection(id) {
  return state.config.sections.find(s => s.id === id);
}

function updateProp(id, key, value) {
  const section = getSection(id);
  if (section) {
    section.props[key] = value;
    // Update canvas card preview in place
    const preview = document.querySelector(`[data-id="${id}"] .section-card-preview`);
    if (preview) preview.innerHTML = cardPreviewHTML(section);
    saveState();
  }
}

// ═══════════════════════════════════════════════════════════
// CANVAS
// ═══════════════════════════════════════════════════════════

function renderCanvas() {
  const list = document.getElementById("canvas-list");
  list.innerHTML = "";

  state.config.sections.forEach(section => {
    const li = document.createElement("li");
    li.className = "section-card";
    li.dataset.id = section.id;
    li.innerHTML = `
      <div class="section-card-header">
        <span class="drag-handle" title="Drag to reorder">⠿</span>
        <span class="section-card-type">${SECTION_LABELS[section.type] || section.type}</span>
        <button class="section-card-delete" title="Remove section" data-id="${section.id}">&times;</button>
      </div>
      <div class="section-card-preview">${cardPreviewHTML(section)}</div>
    `;

    li.addEventListener("click", (e) => {
      if (!e.target.closest(".section-card-delete") && !e.target.closest(".drag-handle")) {
        selectSection(section.id);
      }
    });
    li.querySelector(".section-card-delete").addEventListener("click", (e) => {
      e.stopPropagation();
      deleteSection(section.id);
    });

    list.appendChild(li);
  });

  // Re-init or update SortableJS
  if (state.sortable) {
    state.sortable.destroy();
  }
  state.sortable = Sortable.create(list, {
    group:      { name: "sections", pull: true, put: ["palette", "sections"] },
    animation:  150,
    handle:     ".drag-handle",
    ghostClass: "sortable-ghost",
    dragClass:  "sortable-drag",
    // Reorder existing sections
    onEnd(evt) {
      if (evt.from === evt.to) {
        // Reorder within canvas
        const moved = state.config.sections.splice(evt.oldIndex, 1)[0];
        state.config.sections.splice(evt.newIndex, 0, moved);
        saveState();
      }
    },
    // Drop from sidebar palette
    onAdd(evt) {
      const type = evt.item.dataset.type;
      // Remove the cloned sidebar element — we'll render a real card
      evt.item.remove();
      if (!type) return;

      const defaults = SECTION_DEFAULTS[type] || {};
      const newSection = {
        id: crypto.randomUUID(),
        type,
        props: { ...defaults, items: defaults.items ? [...defaults.items] : undefined },
      };
      if (newSection.props.items === undefined) delete newSection.props.items;

      state.config.sections.splice(evt.newIndex, 0, newSection);
      saveState();
      renderCanvas();
      renderEmpty();
      selectSection(newSection.id);
    },
  });

  // Re-apply selected highlight
  if (state.selectedId) {
    document.querySelector(`[data-id="${state.selectedId}"]`)?.classList.add("selected");
  }
}

function renderEmpty() {
  const empty = document.getElementById("canvas-empty");
  const list  = document.getElementById("canvas-list");
  if (!empty) return;
  if (state.config.sections.length === 0) {
    empty.classList.add("visible");
    list.style.display = "none";
  } else {
    empty.classList.remove("visible");
    list.style.display = "";
  }
}

function cardPreviewHTML(section) {
  const p = section.props;
  switch (section.type) {
    case "header":
      return `<strong>${esc(p.orgName || "")}</strong> — ${esc(p.tagline || "")}`;
    case "footer":
      return `© ${esc(p.year || "")} ${esc(p.orgName || "")}`;
    case "bullet_list":
    case "event_list": {
      const items = (p.items || []).slice(0, 2).map(i => `• ${esc(i)}`).join("<br>");
      return `<strong>${esc(p.heading || "")}</strong><br>${items}${(p.items||[]).length > 2 ? "<br>…" : ""}`;
    }
    case "text_block":
      return `<strong>${esc(p.heading || "")}</strong><br>${esc((p.body || "").slice(0, 80))}`;
    case "product_card": {
      const img = p.imageUrl ? `<img src="${p.imageUrl}" class="preview-image" alt="" />` : "";
      return `${img}<strong>${esc(p.title || "")}</strong><br>${esc((p.body || "").slice(0, 60))}`;
    }
    case "image_block":
      return p.imageUrl
        ? `<img src="${p.imageUrl}" class="preview-image" alt="${esc(p.altText || "")}" /> ${esc(p.caption || "")}`
        : `<em style="color:#ccc">No image selected</em>`;
    case "divider":
      return `<hr style="border:none;border-top:1px solid ${p.color || "#E0E0E0"};margin:4px 0;">`;
    default:
      return "";
  }
}

// ═══════════════════════════════════════════════════════════
// PROPERTIES PANEL
// ═══════════════════════════════════════════════════════════

function renderPropertiesPanel() {
  const placeholder = document.getElementById("props-placeholder");
  const form        = document.getElementById("props-form");

  const section = state.selectedId ? getSection(state.selectedId) : null;

  if (!section) {
    placeholder.hidden = false;
    form.hidden = true;
    return;
  }

  placeholder.hidden = true;
  form.hidden = false;
  form.innerHTML = buildPropsForm(section);
  bindPropsForm(section);

  // Scroll the right panel back to the top so fields are visible immediately
  const panelRight = document.getElementById("panel-right");
  if (panelRight) panelRight.scrollTop = 0;
}

function buildPropsForm(section) {
  const p    = section.props;
  const type = section.type;
  let html   = `<span class="props-section-type">${SECTION_LABELS[type] || type}</span>`;

  switch (type) {
    case "header":
      html += textField("orgName", "Organisation Name", p.orgName);
      html += textField("tagline", "Tagline / Subtitle", p.tagline);
      html += imageField("logoUrl", "Logo Image (optional)", p.logoUrl);
      html += colorField("backgroundColor", "Accent Color", p.backgroundColor || state.config.theme.primaryColor || "#E31837");
      html += colorField("textColor", "Text Color", p.textColor || "#ffffff");
      break;

    case "footer":
      html += textField("orgName", "Organisation Name", p.orgName);
      html += textField("year", "Year", p.year);
      html += colorField("backgroundColor", "Background Color", p.backgroundColor || "#1A1A1A");
      html += colorField("textColor", "Text Color", p.textColor || "#aaaaaa");
      break;

    case "bullet_list":
      html += textField("heading", "Heading", p.heading);
      html += itemsField("items", "Items", p.items);
      html += colorField("bulletColor", "Bullet Color", p.bulletColor || state.config.theme.primaryColor || "#E31837");
      html += colorField("backgroundColor", "Background Color", p.backgroundColor || "#ffffff");
      break;

    case "event_list":
      html += textField("heading", "Heading", p.heading);
      html += itemsField("items", "Events", p.items);
      html += colorField("backgroundColor", "Background Color", p.backgroundColor || "#EEF6F7");
      break;

    case "text_block":
      html += textField("heading", "Heading", p.heading);
      html += textareaField("body", "Body Text", p.body);
      html += colorField("backgroundColor", "Background Color", p.backgroundColor || "#ffffff");
      break;

    case "product_card":
      html += textField("title", "Title", p.title);
      html += textareaField("body", "Description", p.body);
      html += textField("creator", "Creator / Author", p.creator);
      html += imageField("imageUrl", "Image", p.imageUrl);
      html += colorField("backgroundColor", "Background Color", p.backgroundColor || "#ffffff");
      break;

    case "image_block":
      html += imageField("imageUrl", "Image", p.imageUrl);
      html += textField("altText", "Alt Text", p.altText);
      html += textField("caption", "Caption (optional)", p.caption);
      html += colorField("backgroundColor", "Background Color", p.backgroundColor || "#ffffff");
      break;

    case "divider":
      html += colorField("color", "Line Color", p.color || "#E0E0E0");
      html += `<label class="field-label">Spacing (px)
        <input type="number" class="field-input" data-prop="spacing" value="${p.spacing || 20}" min="0" max="80" />
      </label>`;
      break;
  }

  return html;
}

function bindPropsForm(section) {
  const form = document.getElementById("props-form");

  // Text / number inputs
  form.querySelectorAll("input[data-prop]:not([type='color']), textarea[data-prop]").forEach(el => {
    el.addEventListener("input", () => {
      let val = el.type === "number" ? parseInt(el.value, 10) : el.value;
      if (el.dataset.prop === "items") {
        val = el.value.split("\n").map(s => s.trimEnd()).filter(s => s !== "");
      }
      updateProp(section.id, el.dataset.prop, val);
    });
  });

  // Color inputs
  form.querySelectorAll(".prop-color-swatch").forEach(swatch => {
    const propKey = swatch.dataset.prop;
    const input   = form.querySelector(`input[type='color'][data-prop='${propKey}']`);
    const hexSpan = form.querySelector(`.prop-color-hex[data-prop='${propKey}']`);

    swatch.style.background = input.value;

    swatch.addEventListener("click", () => input.click());
    input.addEventListener("input", () => {
      swatch.style.background = input.value;
      hexSpan.textContent = input.value;
      updateProp(section.id, propKey, input.value);
    });
  });

  // Link insertion buttons
  form.querySelectorAll(".link-btn").forEach(btn => {
    const prop    = btn.dataset.target;
    const dialog  = form.querySelector(`#ld-${prop}`);
    const urlInput = dialog?.querySelector(".link-url-input");

    btn.addEventListener("click", () => {
      const isOpen = !dialog.hidden;
      // Close all other open dialogs first
      form.querySelectorAll(".link-dialog").forEach(d => { d.hidden = true; });
      dialog.hidden = isOpen;
      if (!isOpen) urlInput?.focus();
    });

    dialog?.querySelector(".link-confirm")?.addEventListener("click", () => {
      const url = urlInput.value.trim();
      if (!url) { urlInput.focus(); return; }

      const ta    = form.querySelector(`#ta-${prop}`);
      const start = ta.selectionStart;
      const end   = ta.selectionEnd;
      const selected = ta.value.substring(start, end).trim();
      const linkText = selected || "link text";
      const markdown = `[${linkText}](${url})`;

      ta.setRangeText(markdown, start, end, "end");
      ta.dispatchEvent(new Event("input")); // trigger state update

      dialog.hidden = true;
      urlInput.value = "";
      ta.focus();
    });

    dialog?.querySelector(".link-cancel")?.addEventListener("click", () => {
      dialog.hidden = true;
      urlInput.value = "";
    });

    // Allow pressing Enter in URL field to confirm
    urlInput?.addEventListener("keydown", (e) => {
      if (e.key === "Enter") { e.preventDefault(); dialog.querySelector(".link-confirm")?.click(); }
      if (e.key === "Escape") { dialog.hidden = true; urlInput.value = ""; }
    });
  });

  // Aon brand color quick-pick swatches
  form.querySelectorAll(".color-quick-swatch").forEach(swatch => {
    swatch.addEventListener("click", () => {
      const propKey  = swatch.dataset.prop;
      const val      = swatch.dataset.value;
      const input    = form.querySelector(`input[type='color'][data-prop='${propKey}']`);
      const bigSwatch = form.querySelector(`.prop-color-swatch[data-prop='${propKey}']`);
      const hexSpan  = form.querySelector(`.prop-color-hex[data-prop='${propKey}']`);
      if (input)     input.value = val;
      if (bigSwatch) bigSwatch.style.background = val;
      if (hexSpan)   hexSpan.textContent = val;
      updateProp(section.id, propKey, val);
    });
  });

  // Image upload triggers
  form.querySelectorAll(".image-upload-area[data-prop]").forEach(area => {
    area.addEventListener("click", () => {
      state.pendingImageSectionId = section.id;
      state.pendingImagePropKey   = area.dataset.prop;
      document.getElementById("image-upload-input").click();
    });
  });
}

// ─── Field builders ───────────────────────────────────────

function textField(prop, label, value = "") {
  return `<label class="field-label">${label}
    <input type="text" class="field-input" data-prop="${prop}" value="${esc(value)}" />
  </label>`;
}

function textareaField(prop, label, value = "") {
  return `<label class="field-label">${label}
    <div class="link-toolbar">
      <button class="link-btn" type="button" data-target="${prop}">🔗 Insert Link</button>
    </div>
    <textarea class="field-input" data-prop="${prop}" id="ta-${prop}">${esc(value)}</textarea>
    <div class="link-dialog" id="ld-${prop}" hidden>
      <input type="url" class="link-url-input field-input" placeholder="https://example.com" />
      <div class="link-dialog-btns">
        <button class="btn btn-primary link-confirm" type="button" data-target="${prop}">Insert</button>
        <button class="btn btn-secondary link-cancel" type="button" data-target="${prop}">Cancel</button>
      </div>
    </div>
  </label>`;
}

function itemsField(prop, label, items = []) {
  return `<label class="field-label">${label}
    <span class="field-helper">One item per line · select a word then click Insert Link to add a hyperlink</span>
    <div class="link-toolbar">
      <button class="link-btn" type="button" data-target="${prop}">🔗 Insert Link</button>
    </div>
    <textarea class="field-input" data-prop="${prop}" id="ta-${prop}" rows="5">${esc((items || []).join("\n"))}</textarea>
    <div class="link-dialog" id="ld-${prop}" hidden>
      <input type="url" class="link-url-input field-input" placeholder="https://example.com" />
      <div class="link-dialog-btns">
        <button class="btn btn-primary link-confirm" type="button" data-target="${prop}">Insert</button>
        <button class="btn btn-secondary link-cancel" type="button" data-target="${prop}">Cancel</button>
      </div>
    </div>
  </label>`;
}

function colorField(prop, label, value = "#ffffff") {
  const swatches = AON_COLORS.map(c =>
    `<div class="color-quick-swatch" data-prop="${prop}" data-value="${c.value}"
          style="background:${c.value};" title="${c.label} — ${c.value}"></div>`
  ).join("");
  return `<label class="field-label">${label}
    <div class="color-field">
      <div class="color-swatch prop-color-swatch" data-prop="${prop}" style="background:${value};"></div>
      <input type="color" data-prop="${prop}" value="${value}" />
      <span class="color-hex prop-color-hex" data-prop="${prop}">${value}</span>
    </div>
    <div class="color-swatches">${swatches}</div>
  </label>`;
}

function imageField(prop, label, url = "") {
  const thumb = url
    ? `<img src="${url}" class="image-preview-thumb" alt="" />`
    : "";
  const hint = url
    ? `<strong>Change image</strong>`
    : `<strong>Click to upload</strong> an image`;
  return `<label class="field-label">${label}
    <div class="image-upload-area" data-prop="${prop}">
      ${thumb}
      <div class="image-upload-label">${hint}</div>
    </div>
  </label>`;
}

// ═══════════════════════════════════════════════════════════
// API ACTIONS
// ═══════════════════════════════════════════════════════════

async function handlePreview() {
  const btn = document.getElementById("btn-preview");
  setLoading(btn, true);
  try {
    const res  = await fetch("/api/preview", jsonPost(state.config));
    const data = await res.json();
    if (!res.ok) { showToast(data.error || "Preview failed", "error"); return; }

    const frame = document.getElementById("preview-frame");
    frame.srcdoc = data.html;
    document.getElementById("preview-modal").hidden = false;

    document.getElementById("preview-close").onclick = () => {
      document.getElementById("preview-modal").hidden = true;
    };
  } catch (e) {
    showToast("Network error: " + e.message, "error");
  } finally {
    setLoading(btn, false);
  }
}

async function handleExport(format) {
  const btn = document.getElementById("btn-export");
  setLoading(btn, true);
  try {
    const payload = { ...state.config, format };
    const res = await fetch("/api/generate", jsonPost(payload));
    if (!res.ok) {
      const data = await res.json().catch(() => ({}));
      showToast(data.error || "Export failed", "error");
      return;
    }
    const blob = await res.blob();
    const name = state.config.meta.newsletterName || "newsletter";
    const ext  = format === "html" ? ".html" : ".eml";
    downloadBlob(blob, name + ext);
    showToast(`Downloaded ${name}${ext}`, "success");
  } catch (e) {
    showToast("Network error: " + e.message, "error");
  } finally {
    setLoading(btn, false);
  }
}

async function handleImportExcel(file) {
  const btn = document.getElementById("btn-import-excel");
  setLoading(btn, true);
  try {
    const fd = new FormData();
    fd.append("xlsx_file", file);
    const res  = await fetch("/api/import-excel", { method: "POST", body: fd });
    const data = await res.json();
    if (!res.ok) { showToast(data.error || "Import failed", "error"); return; }

    state.config.sections = data.sections;
    state.selectedId = null;
    saveState();
    renderCanvas();
    renderEmpty();
    renderPropertiesPanel();
    showToast(`Imported ${data.sections.length} sections from Excel`, "success");
  } catch (e) {
    showToast("Network error: " + e.message, "error");
  } finally {
    setLoading(btn, false);
  }
}

async function handleImageUpload(file) {
  const sectionId = state.pendingImageSectionId;
  const propKey   = state.pendingImagePropKey;
  if (!sectionId || !propKey) return;

  try {
    const fd = new FormData();
    fd.append("image", file);
    const res  = await fetch("/api/upload-image", { method: "POST", body: fd });
    const data = await res.json();
    if (!res.ok) { showToast(data.error || "Upload failed", "error"); return; }

    updateProp(sectionId, propKey, data.url);

    // If still selected, re-render properties to show new thumbnail
    if (state.selectedId === sectionId) {
      renderPropertiesPanel();
      // Re-highlight selected card
      document.querySelectorAll(".section-card").forEach(c => {
        c.classList.toggle("selected", c.dataset.id === sectionId);
      });
    }
    // Also refresh the canvas card preview
    const preview = document.querySelector(`[data-id="${sectionId}"] .section-card-preview`);
    if (preview) preview.innerHTML = cardPreviewHTML(getSection(sectionId));

    showToast("Image uploaded", "success");
  } catch (e) {
    showToast("Upload error: " + e.message, "error");
  }
}

// ═══════════════════════════════════════════════════════════
// PERSISTENCE
// ═══════════════════════════════════════════════════════════

const LS_KEY = "aon_newsletter_builder_state";

function saveState() {
  try {
    localStorage.setItem(LS_KEY, JSON.stringify(state.config));
  } catch (_) {}
}

function loadState() {
  try {
    const raw = localStorage.getItem(LS_KEY);
    if (!raw) return;
    const saved = JSON.parse(raw);
    state.config = { ...state.config, ...saved };
    // Sync form fields
    syncMetaToDOM();
    syncThemeToDOM();
  } catch (_) {}
}

function syncMetaToDOM() {
  const m = state.config.meta;
  setVal("meta-name",    m.newsletterName || "");
  setVal("meta-subject", m.subject || "");
  setVal("meta-from",    m.from || "");
  setVal("meta-to",      m.to || "");
}

function syncThemeToDOM() {
  const t = state.config.theme;
  setVal("theme-bg",    t.backgroundColor || "#F5F5F5");
  setVal("theme-width", t.tableWidth || 700);
  updateSwatch("swatch-bg", "hex-bg", t.backgroundColor || "#F5F5F5");
}

function updateSwatch(swatchId, hexId, value) {
  const swatch = document.getElementById(swatchId);
  const hex    = document.getElementById(hexId);
  if (swatch) swatch.style.background = value;
  if (hex)    hex.textContent = value;
}

// ═══════════════════════════════════════════════════════════
// UTILITIES
// ═══════════════════════════════════════════════════════════

function esc(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function jsonPost(body) {
  return { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) };
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a   = document.createElement("a");
  a.href     = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 1000);
}

let toastTimer;
function showToast(msg, type = "success") {
  const toast = document.getElementById("toast");
  toast.textContent = msg;
  toast.className   = `toast ${type}`;
  toast.hidden      = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { toast.hidden = true; }, 3500);
}

function setLoading(btn, loading) {
  if (!btn) return;
  if (loading) {
    btn.dataset.origHtml = btn.innerHTML;
    btn.innerHTML = `<span class="spinner"></span> Working…`;
    btn.disabled  = true;
  } else {
    if (btn.dataset.origHtml) btn.innerHTML = btn.dataset.origHtml;
    btn.disabled = false;
  }
}

function setVal(id, value) {
  const el = document.getElementById(id);
  if (el) el.value = value;
}
