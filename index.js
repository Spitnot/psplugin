/**
 * Posterizer — index.js
 *
 * Applies a rugged glued-paper texture effect to the active Photoshop document.
 *
 * Folder layout (plugin root, not tracked in git — files are heavy):
 *   textures/001.jpg … 040.jpg   — paper texture (screen + multiply layers)
 *   displace/001.psd … 040.psd   — displacement maps (wrinkle / fold shape)
 *
 * Fix for UXP batchPlay: transform after placeEvent
 * ─────────────────────────────────────────────────
 * placeEvent leaves the Smart Object in Photoshop's internal free-transform
 * mode.  Any transform batchPlay issued while that mode is active succeeds
 * silently but has no effect on the layer.
 *
 * Solution: within the SAME modal as placeEvent, re-select the layer by its
 * ID.  That select action commits the pending free-transform so the layer is
 * in a clean state.  A second modal can then read real bounds and apply a
 * proper transform using pixelsUnit (absolute dimensions).
 */

/* global require */
const { app }                         = require("photoshop");
const { core, action }                = require("photoshop");
const { localFileSystem: fs }         = require("uxp").storage;

// ─── constants ───────────────────────────────────────────────────────────────

const TEXTURE_COUNT = 40;
const state = { lastTextureIndex: -1 };

// ─── helpers ─────────────────────────────────────────────────────────────────

function $(id) { return document.getElementById(id); }

function setStatus(msg, type) {
  const el = $("statusBar");
  if (!el) return;
  el.textContent = msg;
  el.className   = "status-bar" + (type ? " " + type : "");
  console.log("[posterizer]", msg);
}

function getParams() {
  return {
    displaceScale: parseFloat($("slDisplace").value),
    shadowOpacity: parseFloat($("slShadow").value),
    screenOpacity: parseFloat($("slScreen").value),
    baseOpacity:   parseFloat($("slBase").value),
  };
}

/** Pick a random texture index, avoiding immediate repeat. */
function pickRandom() {
  let idx;
  do {
    idx = Math.floor(Math.random() * TEXTURE_COUNT) + 1;
  } while (idx === state.lastTextureIndex && TEXTURE_COUNT > 1);
  state.lastTextureIndex = idx;
  return idx;
}

/**
 * Build session tokens for the texture jpg and displace psd that correspond
 * to the given numeric index.  Both folders live at the plugin root.
 */
async function getTokens(index) {
  const pluginFolder  = await fs.getPluginFolder();
  const paddedName    = String(index).padStart(3, "0");

  const textureFolder  = await pluginFolder.getEntry("textures");
  const displaceFolder = await pluginFolder.getEntry("displace");

  const textureEntry  = await textureFolder.getEntry(paddedName + ".jpg");
  const displaceEntry = await displaceFolder.getEntry(paddedName + ".psd");

  return {
    textureToken:  fs.createSessionToken(textureEntry),
    displaceToken: fs.createSessionToken(displaceEntry),
  };
}

// ─── core: place + scale ─────────────────────────────────────────────────────

/**
 * Place a file as a Smart Object and scale it to COVER the document canvas.
 *
 * Returns the new layer's ID.
 *
 * Two-modal pattern (required to work around the free-transform issue):
 *
 *  Modal 1 — placeEvent + select (commits free-transform)
 *  Modal 2 — read real bounds → compute cover scale → transform (pixelsUnit)
 */
async function placeAndFit(token, doc, name, blendMode, opacity) {
  let layerId;

  // ── Modal 1: place and immediately commit free-transform ──────────────────
  await core.executeAsModal(async () => {
    const [placeResult] = await action.batchPlay([{
      _obj: "placeEvent",
      null: { _path: token, _kind: "local" },
      freeTransformCenterState: { _enum: "quadCenterState", _value: "QCSAverage" },
      offset: {
        _obj: "offset",
        horizontal: { _unit: "pixelsUnit", _value: 0 },
        vertical:   { _unit: "pixelsUnit", _value: 0 },
      },
      _options: { dialogOptions: "dontDisplay" },
    }], {});

    layerId = placeResult.ID;

    // KEY FIX — selecting the layer by ID in the same modal commits the
    // pending free-transform that placeEvent leaves behind.  Without this,
    // any transform in a subsequent modal is a silent no-op.
    await action.batchPlay([{
      _obj: "select",
      _target: [{ _ref: "layer", _id: layerId }],
      makeVisible: false,
      _options: { dialogOptions: "dontDisplay" },
    }], {});
  }, { commandName: "Place Smart Object" });

  // ── Modal 2: read bounds, compute cover-scale, transform, set props ───────
  await core.executeAsModal(async () => {
    const layer = doc.layers.find((l) => l.id === layerId);
    if (!layer) throw new Error(`Layer ${layerId} not found after place.`);

    const b = layer.bounds;                      // top/left/bottom/right in px
    const layerW = b.right  - b.left;
    const layerH = b.bottom - b.top;

    if (layerW === 0 || layerH === 0) {
      throw new Error("Placed layer has zero dimensions; cannot scale.");
    }

    // Cover strategy: pick the scale factor that fills the canvas in both axes.
    const scale  = Math.max(doc.width / layerW, doc.height / layerH);
    const newW   = Math.round(layerW * scale);
    const newH   = Math.round(layerH * scale);

    // Transform using absolute pixel dimensions.
    // pixelsUnit is more reliable than percentUnit immediately after a place.
    await action.batchPlay([{
      _obj: "transform",
      _target: [{ _ref: "layer", _id: layerId }],
      freeTransformCenterState: { _enum: "quadCenterState", _value: "QCSAverage" },
      width:  { _unit: "pixelsUnit", _value: newW },
      height: { _unit: "pixelsUnit", _value: newH },
      interfaceIconFrameDimmed: {
        _enum: "interpolationType",
        _value: "bicubicAutomatic",
      },
      _options: { dialogOptions: "dontDisplay" },
    }], {});

    // Set blend mode, opacity, and name in one descriptor.
    await action.batchPlay([{
      _obj: "set",
      _target: [{ _ref: "layer", _id: layerId }],
      to: {
        _obj: "layer",
        name,
        mode:    { _enum: "blendMode",   _value: blendMode },
        opacity: { _unit: "percentUnit", _value: opacity   },
      },
      _options: { dialogOptions: "dontDisplay" },
    }], {});
  }, { commandName: "Scale and Configure Layer" });

  return layerId;
}

// ─── core: full posterizer effect ────────────────────────────────────────────

async function applyPosterEffect(textureToken, displaceToken, params) {
  const doc = app.activeDocument;
  if (!doc) throw new Error("No active document.");

  // Remove any previous "Posterizer" group.
  setStatus("Cleaning previous effect…", "working");
  await core.executeAsModal(async () => {
    const old = doc.layers.filter((l) => l.name === "Posterizer");
    for (const layer of old) await layer.delete();
  }, { commandName: "Remove Previous Posterizer" });

  // Place texture layers.
  setStatus("Placing highlights layer…", "working");
  const brillosId = await placeAndFit(
    textureToken, doc, "textura — brillos", "screen",   params.screenOpacity
  );

  setStatus("Placing shadows layer…", "working");
  const sombrasId = await placeAndFit(
    textureToken, doc, "textura — sombras", "multiply", params.shadowOpacity
  );

  // Duplicate the background, convert to Smart Object, apply Displace filter.
  setStatus("Applying displacement…", "working");
  let imgId;
  await core.executeAsModal(async () => {
    const orig = doc.layers[doc.layers.length - 1];
    const dup  = await orig.duplicate();
    await dup.moveBefore(doc.layers[0]);

    // Convert to Smart Object so the Displace filter is non-destructive.
    await action.batchPlay([{
      _obj: "newPlacedLayer",
      _options: { dialogOptions: "dontDisplay" },
    }], {});

    await action.batchPlay([{
      _obj: "displace",
      horizontalScale: params.displaceScale,
      verticalScale:   params.displaceScale,
      displacementMap: { _enum: "displacementMap", _value: "stretchToFit" },
      undefinedArea:   { _enum: "undefinedArea",   _value: "repeatEdgePixels" },
      displaceFile:    { _path: displaceToken, _kind: "local" },
      _options: { dialogOptions: "dontDisplay" },
    }], {});

    imgId = doc.activeLayers[0].id;
    await action.batchPlay([{
      _obj: "set",
      _target: [{ _ref: "layer", _enum: "ordinal", _value: "targetEnum" }],
      to: { _obj: "layer", name: "imagen — desplazada" },
      _options: { dialogOptions: "dontDisplay" },
    }], {});
  }, { commandName: "Displacement Map" });

  // Place the base paper layer (normal blend, semi-transparent).
  setStatus("Placing base paper layer…", "working");
  const papelId = await placeAndFit(
    textureToken, doc, "papel — base", "normal", params.baseOpacity
  );

  // Select all four layers and group them.
  setStatus("Grouping layers…", "working");
  await core.executeAsModal(async () => {
    await action.batchPlay([{
      _obj: "select",
      _target: [{ _ref: "layer", _id: brillosId }],
      makeVisible: false,
      _options: { dialogOptions: "dontDisplay" },
    }], {});

    for (const lid of [sombrasId, imgId, papelId]) {
      await action.batchPlay([{
        _obj: "select",
        _target: [{ _ref: "layer", _id: lid }],
        selectionModifier: {
          _enum: "_selectionModifierType",
          _value: "addToSelection",
        },
        makeVisible: false,
        _options: { dialogOptions: "dontDisplay" },
      }], {});
    }

    await action.batchPlay([{
      _obj: "make",
      _target: [{ _ref: "layerSection" }],
      from:  { _ref: "layer", _enum: "ordinal", _value: "targetEnum" },
      using: { _obj: "layerSection", name: "Posterizer" },
      _options: { dialogOptions: "dontDisplay" },
    }], {});
  }, { commandName: "Group Layers" });

  const padded = String(state.lastTextureIndex).padStart(3, "0");
  setStatus(`Texture #${padded} applied ✓`, "ok");
  $("previewLabel").textContent = `Textura ${padded}.jpg`;
  $("previewSeed").textContent  = `Displace: ${params.displaceScale}px`;
  $("btnRegen").disabled = $("btnSave").disabled = $("btnFlatten").disabled = false;
}

// ─── save helpers ─────────────────────────────────────────────────────────────

async function saveAsCopy() {
  try {
    await action.batchPlay([{
      _obj: "exportDocumentAs",
      _options: { dialogOptions: "display" },
    }], { synchronousExecution: false });
    setStatus("Exported ✓", "ok");
  } catch {
    setStatus("Export cancelled.", "");
  }
}

async function flattenAndSave() {
  const doc = app.activeDocument;
  try {
    await core.executeAsModal(async () => {
      await action.batchPlay([{
        _obj: "flattenImage",
        _options: { dialogOptions: "dontDisplay" },
      }], {});
    }, { commandName: "Flatten" });
    await doc.save();
    $("btnRegen").disabled = $("btnSave").disabled = $("btnFlatten").disabled = true;
    $("previewLabel").textContent = "Saved ✓";
    $("previewSeed").textContent  = "—";
    setStatus("Saved ✓", "ok");
  } catch (e) {
    setStatus(`Error: ${e.message}`, "error");
  }
}

// ─── UI wiring ────────────────────────────────────────────────────────────────

// Sync slider display values.
[
  ["slDisplace", "valDisplace"],
  ["slShadow",   "valShadow"  ],
  ["slScreen",   "valScreen"  ],
  ["slBase",     "valBase"    ],
].forEach(([sliderId, labelId]) => {
  const slider = $(sliderId);
  const label  = $(labelId);
  if (slider && label) {
    slider.addEventListener("input", () => { label.textContent = slider.value; });
  }
});

async function runEffect() {
  setStatus("Selecting texture…", "working");
  try {
    const idx = pickRandom();
    const { textureToken, displaceToken } = await getTokens(idx);
    await applyPosterEffect(textureToken, displaceToken, getParams());
  } catch (e) {
    setStatus(`Error: ${e.message}`, "error");
    console.error(e);
  }
}

$("btnApply").addEventListener("click", async () => {
  $("btnApply").disabled = true;
  try   { await runEffect(); }
  finally { $("btnApply").disabled = false; }
});

$("btnRegen").addEventListener("click", async () => {
  $("btnRegen").disabled = $("btnApply").disabled = true;
  try   { await runEffect(); }
  finally { $("btnRegen").disabled = $("btnApply").disabled = false; }
});

$("btnSave").addEventListener("click", () => saveAsCopy());
$("btnFlatten").addEventListener("click", () => flattenAndSave());
