/**
 * Place & Scale — UXP Photoshop Plugin
 *
 * Fix for: transform batchPlay after placeEvent has no effect
 *
 * Root cause
 * ----------
 * placeEvent leaves the new Smart Object layer in Photoshop's internal
 * "free transform" mode.  Even when executed in a separate executeAsModal
 * call, that pending free-transform context prevents any subsequent
 * `transform` descriptor from taking effect (the call succeeds silently
 * but the layer dimensions are unchanged).
 *
 * Fix
 * ---
 * 1. In the SAME modal as placeEvent, immediately re-select the newly
 *    placed layer by its ID.  That `select` action commits / cancels the
 *    pending free-transform context and leaves the layer in a clean state.
 * 2. In a second modal, read the actual document and layer pixel dimensions
 *    from the DOM API, compute the "cover" scale factor, then apply it via
 *    a `transform` descriptor that targets the layer explicitly by ID.
 *    Using pixelsUnit (absolute dimensions) is more reliable than
 *    percentUnit for this operation.
 */

/* global require */
const { app, constants } = require("photoshop");
const { core, action } = require("photoshop");
const { localFileSystem: fs } = require("uxp").storage;

// ─── helpers ────────────────────────────────────────────────────────────────

/** Write a line to the status <div> in the panel. */
function setStatus(msg) {
  const el = document.getElementById("status");
  if (el) el.textContent = msg;
  console.log("[place-and-scale]", msg);
}

/**
 * Pick a local file using the UXP file picker.
 * Returns a UXP File entry or null if cancelled.
 */
async function pickFile() {
  const entry = await fs.getFileForOpening({
    allowMultiple: false,
    types: ["jpg", "jpeg", "png", "tif", "tiff", "psb", "psd"],
  });
  return entry || null;
}

// ─── core logic ─────────────────────────────────────────────────────────────

/**
 * Place a file as a Smart Object and scale it to cover the document canvas.
 *
 * @param {import("uxp").storage.File} fileEntry  The file to place.
 */
async function placeAndScale(fileEntry) {
  const doc = app.activeDocument;
  if (!doc) throw new Error("No active document.");

  // UXP token for the chosen file (required by placeEvent)
  const sessionToken = fs.createSessionToken(fileEntry);

  let layerId;

  // ── Modal 1: place the Smart Object and immediately commit free-transform ──
  //
  // After placeEvent, Photoshop keeps the layer in free-transform mode.
  // Re-selecting the same layer by ID within the same modal commits that
  // pending state so the layer is "clean" before we leave this modal.
  await core.executeAsModal(
    async () => {
      const [placeResult] = await action.batchPlay(
        [
          {
            _obj: "placeEvent",
            null: { _path: sessionToken, _kind: "local" },
            // Place centred on the canvas origin; we will resize next.
            freeTransformCenterState: {
              _enum: "quadCenterState",
              _value: "QCSAverage",
            },
            offset: {
              _obj: "offset",
              horizontal: { _unit: "pixelsUnit", _value: 0 },
              vertical: { _unit: "pixelsUnit", _value: 0 },
            },
            _options: { dialogOptions: "dontDisplay" },
          },
        ],
        {}
      );

      layerId = placeResult.ID;
      setStatus(`Placed layer ID ${layerId}. Committing free-transform…`);

      // KEY FIX ─ select the layer by ID to commit the free-transform.
      // Without this, transform in any subsequent modal is a no-op.
      await action.batchPlay(
        [
          {
            _obj: "select",
            _target: [{ _ref: "layer", _id: layerId }],
            makeVisible: false,
            _options: { dialogOptions: "dontDisplay" },
          },
        ],
        {}
      );
    },
    { commandName: "Place Smart Object" }
  );

  // ── Modal 2: compute cover-scale and apply transform ──────────────────────
  await core.executeAsModal(
    async () => {
      // Read dimensions from the DOM API (reliable, no batchPlay needed).
      const docWidth = doc.width;   // pixels
      const docHeight = doc.height; // pixels

      // Find the placed layer by ID.
      const layer = doc.layers.find((l) => l.id === layerId);
      if (!layer) throw new Error(`Layer ${layerId} not found after place.`);

      const b = layer.bounds; // { top, left, bottom, right } in pixels
      const layerWidth  = b.right  - b.left;
      const layerHeight = b.bottom - b.top;

      if (layerWidth === 0 || layerHeight === 0) {
        throw new Error("Placed layer has zero dimensions; cannot scale.");
      }

      // "Cover" strategy: scale so the layer fills the canvas in both axes,
      // preserving aspect ratio (the larger scale factor wins).
      const scaleX = docWidth  / layerWidth;
      const scaleY = docHeight / layerHeight;
      const scale  = Math.max(scaleX, scaleY);

      const newWidth  = Math.round(layerWidth  * scale);
      const newHeight = Math.round(layerHeight * scale);

      setStatus(
        `Scaling ${layerWidth}×${layerHeight} → ${newWidth}×${newHeight} ` +
          `(canvas ${docWidth}×${docHeight})`
      );

      // Apply the transform.
      // - _target by layer ID ensures the correct layer is transformed even
      //   if selection changed between modals.
      // - pixelsUnit is more reliable than percentUnit here because
      //   Photoshop's % is relative to the layer's CURRENT size, which can
      //   be ambiguous immediately after a place operation.
      // - interfaceIconFrameDimmed sets the interpolation method
      //   (bicubicAutomatic is the PS default for enlargements).
      const [transformResult] = await action.batchPlay(
        [
          {
            _obj: "transform",
            _target: [{ _ref: "layer", _id: layerId }],
            freeTransformCenterState: {
              _enum: "quadCenterState",
              _value: "QCSAverage",
            },
            width:  { _unit: "pixelsUnit", _value: newWidth  },
            height: { _unit: "pixelsUnit", _value: newHeight },
            interfaceIconFrameDimmed: {
              _enum: "interpolationType",
              _value: "bicubicAutomatic",
            },
            _options: { dialogOptions: "dontDisplay" },
          },
        ],
        {}
      );

      if (transformResult == null) {
        // This should no longer happen after the free-transform commit above,
        // but guard just in case.
        throw new Error(
          "transform returned undefined — free-transform may still be active."
        );
      }
    },
    { commandName: "Scale to Canvas" }
  );

  setStatus("Done — layer placed and scaled to cover the canvas.");
}

// ─── UI wiring ───────────────────────────────────────────────────────────────

document.getElementById("btnPlace").addEventListener("click", async () => {
  try {
    setStatus("Picking file…");
    const file = await pickFile();
    if (!file) {
      setStatus("Cancelled.");
      return;
    }
    setStatus(`Placing ${file.name}…`);
    await placeAndScale(file);
  } catch (err) {
    setStatus(`Error: ${err.message}`);
    console.error(err);
  }
});
