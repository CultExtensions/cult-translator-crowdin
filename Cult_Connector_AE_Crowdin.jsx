//@target aftereffects
(function CrowdinAE_Plugin(thisObj){

  var SERVER_BASE = "https://ae-translate-proxy.onrender.com";
  // Crowdin: proxy uses Storage then screenshot API. First run can be slow (proxy cold).
  var EP_OAUTH_START    = SERVER_BASE + "/integrations/crowdin/oauth/start.json";
  var EP_OAUTH_STATUS   = SERVER_BASE + "/integrations/crowdin/oauth/status";
  var EP_PROJECTS       = SERVER_BASE + "/integrations/crowdin/projects";
  var EP_SELECT_PROJECT = SERVER_BASE + "/integrations/crowdin/select-project";
  var EP_LANGS          = SERVER_BASE + "/integrations/crowdin/project/languages";
  var EP_STRINGS        = SERVER_BASE + "/integrations/crowdin/ae/strings";
  var EP_SCAN_FRAME     = SERVER_BASE + "/integrations/crowdin/ae/scan-frame";
  // Crowdin docs: support.crowdin.com/developer/automating-screenshot-management — API uses Storage then add screenshot.
  // Tagging: server should use PUT .../screenshots/{id}/tags with one array of { stringId, position } per screenshot (not one Add Tag per string). See VELOCITY_NOTES.
  var EP_PULL           = SERVER_BASE + "/integrations/crowdin/ae/pull";

  var IS_WIN = ($.os && $.os.indexOf("Windows") === 0);
  var CURL   = IS_WIN ? "C:\\Windows\\System32\\curl.exe" : "/usr/bin/curl";
  var TEST_LICENSE = "TEST";

  // Plugin version (used for update checks via GitHub Releases).
  var PLUGIN_VERSION = "0.1.1";
  var UPDATE_GITHUB_OWNER = "CultExtensions";
  var UPDATE_GITHUB_REPO  = "cult-translator-crowdin";
  // Preferred release asset name; if not found we fall back to the first `.jsx` asset.
  var UPDATE_ASSET_NAME_PREFERRED = "Cult_Connector_AE_Crowdin.jsx";

  // ✅ knobs
  var MIN_OPACITY = 10;           // 10% minimum (allow export); prefer 100%
  var PREFERRED_OPACITY = 100;    // prefer frame when opacity is 100%
  var MIN_SCALE = 0.5;            // minimum scale (50%) when scale animates; prefer 100%
  var MIN_IN_RATIO = 0.35;        // strict: 35% bbox inside
  var STABLE_FRAMES = 3;          // typewriter stability frames
  var CAPTURE_DELAY_FRAMES = 2;   // capture a few frames later
  var FALLBACK_RATIO = 0.05;      // fallback: 5% bbox inside
  // Screenshot resolution for Crowdin (1 = full comp size; 2 = half; 4 = quarter). Set EXPORT_SCALE_UP to render larger than comp (e.g. 2 = double dimensions).
  var SCREENSHOT_RES_FACTOR = (typeof SCREENSHOT_RES_FACTOR_OVERRIDE !== "undefined" && SCREENSHOT_RES_FACTOR_OVERRIDE > 0) ? SCREENSHOT_RES_FACTOR_OVERRIDE : 1;
  // Timeline-scan resolution: matches SCREENSHOT_RES_FACTOR by default.
  var SCAN_RES_FACTOR = (typeof SCAN_RES_FACTOR_OVERRIDE !== "undefined" && SCAN_RES_FACTOR_OVERRIDE > 0) ? SCAN_RES_FACTOR_OVERRIDE : SCREENSHOT_RES_FACTOR;
  // Scan export resolution for Crowdin: 1 = Full, 0.5 = Half, 0.25 = Quarter. Half = smaller uploads. Override: SCAN_EXPORT_RESOLUTION_OVERRIDE.
  var SCAN_EXPORT_RESOLUTION = (typeof SCAN_EXPORT_RESOLUTION_OVERRIDE !== "undefined" && SCAN_EXPORT_RESOLUTION_OVERRIDE > 0 && SCAN_EXPORT_RESOLUTION_OVERRIDE <= 1) ? SCAN_EXPORT_RESOLUTION_OVERRIDE : 0.5;
  // Export at 3x comp dimensions for Crowdin context (larger + sharper). 1 = comp size; 2 = double; 3 = triple. Override: EXPORT_SCALE_UP_OVERRIDE.
  var EXPORT_SCALE_UP = (typeof EXPORT_SCALE_UP_OVERRIDE !== "undefined" && EXPORT_SCALE_UP_OVERRIDE >= 1) ? EXPORT_SCALE_UP_OVERRIDE : 3;

  // ✅ NEW rescue bbox size (used only if geometry fails)
  var RESCUE_BBOX_W = 280;
  var RESCUE_BBOX_H = 100;
  // Shrink sourceRect-based bbox so highlight better matches visible text (AE often returns slightly large bounds). 1 = no change, 0.96 = 2% inset each side.
  var BBOX_TIGHTEN_RATIO = (typeof BBOX_TIGHTEN_RATIO_OVERRIDE !== "undefined" && BBOX_TIGHTEN_RATIO_OVERRIDE >= 0.5 && BBOX_TIGHTEN_RATIO_OVERRIDE <= 1) ? BBOX_TIGHTEN_RATIO_OVERRIDE : 0.96;
  // Crowdin API expects tag positions in 480×270. Server scales from our export dimensions to 480×270.
  var CROWDIN_TAG_W = 480;
  var CROWDIN_TAG_H = 270;
  // Max export size: resolution factor scales comp to fit inside this (no comp resize). Override: SCAN_EXPORT_RESOLUTION_OVERRIDE (0-1).
  var CROWDIN_EXPORT_MAX_W = 1920;
  var CROWDIN_EXPORT_MAX_H = 1080;
  // Half quality: resolution/downsample factor applied on top of fit scale (0.5 = half res; 1 = full). Override: CROWDIN_EXPORT_DOWNSAMPLE_OVERRIDE (0.25–1).
  var CROWDIN_EXPORT_DOWNSAMPLE_FACTOR = (typeof CROWDIN_EXPORT_DOWNSAMPLE_OVERRIDE !== "undefined" && CROWDIN_EXPORT_DOWNSAMPLE_OVERRIDE >= 0.25 && CROWDIN_EXPORT_DOWNSAMPLE_OVERRIDE <= 1) ? CROWDIN_EXPORT_DOWNSAMPLE_OVERRIDE : 0.5;
  // Crowdin displays screenshots larger in the context view when the filename ends with @2x or @3x (e.g. name@2x.png). 1 = normal name (no suffix).
  var CROWDIN_DISPLAY_SCALE = (typeof CROWDIN_DISPLAY_SCALE_OVERRIDE !== "undefined" && CROWDIN_DISPLAY_SCALE_OVERRIDE >= 1 && CROWDIN_DISPLAY_SCALE_OVERRIDE <= 3) ? Math.floor(CROWDIN_DISPLAY_SCALE_OVERRIDE) : 1;
  // PNG quality for scan export on Mac only: "high"=90-100, "normal"=75-92, "low"=50-75, "min"=35-55 (smallest). Requires pngquant. Windows: server compresses. Override: SCAN_PNG_QUALITY_OVERRIDE.
  var SCAN_PNG_QUALITY = (typeof SCAN_PNG_QUALITY_OVERRIDE !== "undefined") ? SCAN_PNG_QUALITY_OVERRIDE : "min";
  // Track matte: set DEBUG_MATTE_LOG = true before running (e.g. in another script that #includes this) to write Crowdin_matte_debug.txt to Documents with per-frame ref bbox and ratio.
  // Typewriter: set DEBUG_TYPEWRITER_LOG = true before running to write Crowdin_typewriter_debug.txt with effect/animator discovery and result time.
  var DEBUG_TYPEWRITER_LOG = false;  // Log files disabled; set true only for debugging typewriter/smart scan.

  /** Marker comment used for Snapshot Marker (preferred screenshot time per layer). One marker per layer; Smart Scan uses this time when present. */
  var SNAPSHOT_MARKER_COMMENT = "Crowdin Snapshot";

  var STATE = {
    connected: false,
    projectId: "",
    projectName: "",
    compId: "",
    fileKey: "",
    projects: [],
    languages: [],
    // Segmentation is on by default; user can uncheck to minimize segmentation.
    useSegmentation: true,
    // Compositions to send: array of comp ids (numbers). When non-empty, Send uses this list instead of panel selection.
    compsToSend: []
  };

  /** Per-layer cache for Blinking Cursor full-text bbox dimensions (w,h). Cleared at start of smartScanTimeline so we never modify effects and Ctrl+Z restores animation. */
  var __blinkFullTextSizeCache = {};

  var compCheckboxes = [];

  /** Scale bbox from export dimensions (ssW×ssH) to Crowdin tag space (480×270) so API accepts positions. */
  function scaleBboxToCrowdin(bb, ssW, ssH) {
    if (!bb || ssW <= 0 || ssH <= 0) return bb;
    var sx = CROWDIN_TAG_W / ssW, sy = CROWDIN_TAG_H / ssH;
    var x = Math.round(bb.x * sx), y = Math.round(bb.y * sy);
    var w = Math.max(1, Math.round(bb.w * sx)), h = Math.max(1, Math.round(bb.h * sy));
    x = Math.min(Math.max(0, x), CROWDIN_TAG_W - 1);
    y = Math.min(Math.max(0, y), CROWDIN_TAG_H - 1);
    w = Math.min(w, CROWDIN_TAG_W - x);
    h = Math.min(h, CROWDIN_TAG_H - y);
    return { x: x, y: y, w: Math.max(1, w), h: Math.max(1, h) };
  }

  /** Map comp-space bbox to export (PNG) pixel space. Offset compensates for AE half-res sampling so highlight centers on text (tune CROWDIN_BBOX_OFFSET_* if needed). */
  function compBboxToExportBbox(bb, scale) {
    if (!bb || scale <= 0) return bb;
    var ox = (typeof CROWDIN_BBOX_OFFSET_X === "number") ? CROWDIN_BBOX_OFFSET_X : -5;
    var oy = (typeof CROWDIN_BBOX_OFFSET_Y === "number") ? CROWDIN_BBOX_OFFSET_Y : -3;
    var x1 = Math.floor(bb.x * scale + ox);
    var y1 = Math.floor(bb.y * scale + oy);
    var x2 = Math.floor((bb.x + bb.w) * scale + ox);
    var y2 = Math.floor((bb.y + bb.h) * scale + oy);
    return {
      x: Math.max(0, x1),
      y: Math.max(0, y1),
      w: Math.max(1, x2 - x1),
      h: Math.max(1, y2 - y1)
    };
  }

  /** Build screenshot name for Crowdin; @2x/@3x suffix makes context view display larger. */
  function crowdinScreenshotName(baseNameNoExt) {
    return baseNameNoExt + (CROWDIN_DISPLAY_SCALE >= 2 ? "@" + CROWDIN_DISPLAY_SCALE + "x" : "") + ".jpg";
  }

  function trim(s){ return (s||"").replace(/^[\s\r\n\t]+|[\s\r\n\t]+$/g,""); }
  /** Strip Blinking Cursor Typewriter cursor chars from start/end (expression uses reveal + c[d-1], c = ["|","_","—","<",">","«","»","^"]). */
  function stripBlinkingCursorCursor(s) {
    if (!s || typeof s !== "string") return "";
    s = s.replace(/[\|\u005f\u2014<>«»\u005e]\s*$/g, "").replace(/^\s*[\|\u005f\u2014<>«»\u005e]/g, "");
    return trim(s);
  }
  function alertIf(s){ try{ alert(s); }catch(e){} }
  function run(cmd){ try{ return system.callSystem(cmd) || ""; }catch(e){ return ""; } }
  function tryCompressPngForUpload(pngFile, highQualityOrProfile){
    if (IS_WIN || !pngFile || !pngFile.exists) return;
    var profile = (highQualityOrProfile === true) ? "skip" : (highQualityOrProfile === false ? "normal" : (highQualityOrProfile || "normal"));
    if (profile === "skip") return;
    var qBand = (profile === "high") ? "90-100" : (profile === "low" ? "50-75" : (profile === "min" ? "35-55" : "75-92"));
    var tmpPath = pngFile.fsName + ".q.png";
    try {
      run("optipng -o2 \"" + pngFile.fsName + "\" 2>/dev/null");
    } catch (e) {}
    try {
      run("pngquant -f --quality " + qBand + " -o \"" + tmpPath + "\" \"" + pngFile.fsName + "\" 2>/dev/null");
      var tmpF = new File(tmpPath);
      if (tmpF.exists && tmpF.length > 0) {
        pngFile.remove();
        run("mv \"" + tmpPath + "\" \"" + pngFile.fsName + "\"");
      } else { try { tmpF.remove(); } catch(e){} }
    } catch (e) {}
  }
  function readTextFile(f){
    try{ f.encoding="UTF-8"; if(!f.open("r")) return ""; var t=f.read(); f.close(); return t; }
    catch(e){ try{f.close();}catch(_){} return ""; }
  }
  function writeTextFile(f, txt){
    try{ f.encoding="UTF-8"; f.lineFeed="Unix"; if(!f.open("w")) return false; f.write(txt); f.close(); return true; }
    catch(e){ try{f.close();}catch(_){} return false; }
  }

  function jsonEscape(s){
    if (s===null || s===undefined) return "";
    s = ""+s;
    s = s.replace(/\\/g, "\\\\");
    s = s.replace(/"/g, "\\\"");
    s = s.replace(/\r/g, "\\r");
    s = s.replace(/\n/g, "\\n");
    s = s.replace(/\t/g, "\\t");
    return s;
  }

  function normalizeHttpCode(raw){
    raw = trim(raw || "");
    var m = raw.match(/(\d{3})(?!.*\d{3})/);
    if (m && m[1]) return m[1];
    m = raw.match(/(\d{3})/);
    if (m && m[1]) return m[1];
    return raw;
  }

  function openUrl(url){
    if (!url) return;
    try{
      if (IS_WIN) run('cmd /c start "" "' + url + '"');
      else run('/usr/bin/open "' + url + '"');
    }catch(e){
      alertIf("Open this URL in your browser:\n\n" + url);
    }
  }

  function getDeviceId(){
    var name;
    if (IS_WIN) name = trim(run("hostname"));
    else {
      name = trim(run("/usr/sbin/scutil --get ComputerName"));
      if (!name) name = trim(run("/bin/hostname"));
    }
    if (!name) name = "unknown-device";
    return name.replace(/[\r\n\t"'\`]/g, " ");
  }

  function ensureCurl(){
    var cv = run(CURL + " -V");
    if (!cv || cv.toLowerCase().indexOf("curl") === -1) {
      alertIf("curl not found.\nCheck AE Prefs: 'Allow Scripts to Write Files and Access Network'.");
      return false;
    }
    return true;
  }

  function getActiveComp(){
    var c = app.project && app.project.activeItem;
    if (!c || !(c instanceof CompItem)) { alertIf("Open/select a composition first."); return null; }
    return c;
  }

  /** Collect all compositions from the project (root folder, recursive). */
  function getAllCompsInProject() {
    var out = [];
    try {
      function collect(folder) {
        if (!folder || typeof folder.numItems !== "number") return;
        for (var i = 1; i <= folder.numItems; i++) {
          try {
            var item = folder.item(i);
            if (!item) continue;
            if (item instanceof CompItem) out.push(item);
            if (item instanceof FolderItem) collect(item);
          } catch (e) {}
        }
      }
      collect(app.project.rootFolder);
    } catch (e) {}
    return out;
  }

  /**
   * Try to bring the Project panel to front so app.project.selection is populated.
   * When a ScriptUI panel has focus, app.project.selection is often empty; executing
   * the "Project" window command then sleeping briefly can restore selection for reading.
   * No-op if executeCommand is unavailable or fails (getSelectedComps will fall back to activeItem).
   */
  function tryFocusProjectPanelForSelection() {
    try {
      if (typeof app.findMenuCommandId !== "function" || typeof app.executeCommand !== "function") return;
      var cmdId = app.findMenuCommandId("Project");
      if (cmdId && typeof cmdId === "number" && cmdId > 0) {
        app.executeCommand(cmdId);
        $.sleep(150);
      }
    } catch (e) {}
  }

  /** Read compositions from Project panel: selection (CompItems only), or activeItem if none selected. Does not use compsToSend list. */
  function getSelectionFromProjectPanel() {
    tryFocusProjectPanelForSelection();
    var sel = [];
    try {
      if (app.project.selection && app.project.selection.length > 0) {
        for (var i = 0; i < app.project.selection.length; i++) {
          if (app.project.selection[i] instanceof CompItem) sel.push(app.project.selection[i]);
        }
      }
    } catch (e) {}
    if (sel.length === 0) {
      try {
        var c = app.project && app.project.activeItem;
        if (c && c instanceof CompItem) sel.push(c);
      } catch (e2) {}
    }
    return sel;
  }

  /** Return compositions to export: compsToSend list if non-empty (resolved by id), else selection/active from Project panel. */
  function getSelectedComps() {
    if (typeof compCheckboxes !== "undefined" && compCheckboxes && compCheckboxes.length > 0) {
      var out = [];
      for (var i = 0; i < compCheckboxes.length; i++) {
        if (compCheckboxes[i].cb.value === true) out.push(compCheckboxes[i].comp);
      }
      if (out.length > 0) return out;
    }
    if (STATE.compsToSend && STATE.compsToSend.length > 0) {
      var resolved = [];
      for (var j = 0; j < STATE.compsToSend.length; j++) {
        var comp = findCompById(STATE.compsToSend[j]);
        if (comp && comp instanceof CompItem) resolved.push(comp);
      }
      if (resolved.length > 0) return resolved;
    }
    return getSelectionFromProjectPanel();
  }

  /** Safe file key for Crowdin (used as filename base). AE comp.name → safe string; else comp_<id>. */
  function safeFileKeyForComp(comp) {
    try {
      var nameRaw = (comp && comp.name != null) ? String(comp.name) : "";
      var name = nameRaw.replace(/^\s+|\s+$/g, "");
      if (name === "") return "comp_" + comp.id;
      var s = name.replace(/[^A-Za-z0-9_.-]+/g, "_").replace(/_+/g, "_").replace(/^_|_$/g, "");
      return s.length > 0 ? s : "comp_" + comp.id;
    } catch (e) { return comp ? "comp_" + comp.id : ""; }
  }

  /** Find a composition in the project by id (number or string). Searches root, folders, and flat list. Returns comp or null. */
  function findCompById(compId) {
    try {
      if (compId == null || compId === "") return null;
      var id = parseInt(compId, 10);
      if (!isFinite(id)) return null;
      function searchFolder(folder) {
        try {
          if (!folder || typeof folder.numItems !== "number") return null;
          for (var i = 1; i <= folder.numItems; i++) {
            var item = folder.item(i);
            if (!item) continue;
            if (item instanceof CompItem && item.id == id) return item;
            if (item instanceof FolderItem) {
              var found = searchFolder(item);
              if (found) return found;
            }
          }
        } catch (e2) {}
        return null;
      }
      var found = searchFolder(app.project.rootFolder);
      if (found) return found;
      var n = app.project.numItems;
      if (typeof n === "number") {
        for (var j = 1; j <= n; j++) {
          try {
            var it = app.project.item(j);
            if (it && (it instanceof CompItem) && it.id == id) return it;
          } catch (e3) {}
        }
      }
      if (app.project.items && typeof app.project.items.length === "number") {
        for (var k = 0; k < app.project.items.length; k++) {
          try {
            var it2 = app.project.items[k] || app.project.items[k + 1];
            if (it2 && (it2 instanceof CompItem) && it2.id == id) return it2;
          } catch (e4) {}
        }
      }
      return null;
    } catch (e) { return null; }
  }

  /** Same logic as File > Scripts > Scale Composition.jsx: parent all unparented layers to newParent. */
  function makeParentLayerOfAllUnparented(theComp, newParent) {
    for (var i = 1; i <= theComp.numLayers; i++) {
      var curLayer = theComp.layer(i);
      if (curLayer !== newParent && curLayer.parent === null) curLayer.parent = newParent;
    }
  }

  /** Scale every camera zoom by scaleBy (same as Scale Composition.jsx). */
  function scaleAllCameraZooms(theComp, scaleBy) {
    for (var i = 1; i <= theComp.numLayers; i++) {
      var curLayer = theComp.layer(i);
      if (curLayer.matchName === "ADBE Camera Layer") {
        var curZoom = curLayer.zoom;
        if (curZoom.numKeys === 0) curZoom.setValue(curZoom.value * scaleBy);
        else for (var j = 1; j <= curZoom.numKeys; j++) curZoom.setValueAtKey(j, curZoom.keyValue(j) * scaleBy);
      }
    }
  }

  /**
   * Scale composition by factor — same algorithm as File > Scripts > Scale Composition.jsx
   * (temp null parent, resize comp, scale null, then remove null). Uses beginUndoGroup/endUndoGroup
   * so the whole operation is one undo step (Ctrl+Z reverts the scale without ever showing the null).
   * Uses try/finally so the temp null is always removed.
   */
  function scaleCompositionByFactor(comp, scaleFactor) {
    if (!comp || !(comp instanceof CompItem) || scaleFactor <= 0 || scaleFactor === 1) return false;
    var newW = Math.floor(comp.width * scaleFactor), newH = Math.floor(comp.height * scaleFactor);
    if (newW < 1 || newH < 1 || newW > 30000 || newH > 30000) return false;
    var null3DLayer = null;
    var undoStarted = false;
    try {
      if (typeof app.beginUndoGroup === "function") {
        app.beginUndoGroup("Scale Composition");
        undoStarted = true;
      }
      null3DLayer = comp.layers.addNull();
      null3DLayer.threeDLayer = true;
      null3DLayer.position.setValue([0, 0, 0]);
      makeParentLayerOfAllUnparented(comp, null3DLayer);
      comp.width = newW;
      comp.height = newH;
      scaleAllCameraZooms(comp, scaleFactor);
      var superParentScale = null3DLayer.scale.value;
      superParentScale[0] = superParentScale[0] * scaleFactor;
      superParentScale[1] = superParentScale[1] * scaleFactor;
      superParentScale[2] = superParentScale[2] * scaleFactor;
      null3DLayer.scale.setValue(superParentScale);
      return true;
    } catch (e) {
      return false;
    } finally {
      if (null3DLayer != null) {
        try { null3DLayer.remove(); } catch (eRemove) {}
      }
      if (undoStarted && typeof app.endUndoGroup === "function") app.endUndoGroup();
    }
  }

  function clamp(v, a, b){ return Math.max(a, Math.min(b, v)); }

  /**
   * Effective time window to scan in a comp.
   * If the user shortened the work area, we only scan inside it; otherwise we use the full duration.
   */
  function getCompScanWindow(comp) {
    var duration = 0;
    try { duration = Math.max(0, Number(comp.duration || 0)); } catch (e) { duration = 0; }

    var start = 0;
    var end = duration > 0 ? duration : 0;

    try {
      var waStart = Number(comp.workAreaStart);
      var waDur   = Number(comp.workAreaDuration);
      if (isFinite(waStart) && isFinite(waDur) && waDur > 0.01) {
        var waEnd = waStart + waDur;
        if (duration > 0) {
          if (waStart < 0) waStart = 0;
          if (waEnd > duration) waEnd = duration;
        }
        // Only treat work area as a limit when it is actually shorter than the full comp.
        if (duration <= 0 || (waEnd - waStart) < (duration - 0.01)) {
          start = waStart;
          end = waEnd;
        }
      }
    } catch (e2) {}

    if (end < start + 0.01) end = start + 0.01;
    return { start: start, end: end };
  }

  // MINI JSON stringify
  function jsonStringifyMini(v){
    var t = typeof v;
    if (v === null || v === undefined) return "null";
    if (t === "string") return '"' + jsonEscape(v) + '"';
    if (t === "number") return isFinite(v) ? String(v) : "null";
    if (t === "boolean") return v ? "true" : "false";
    if (v instanceof Array) {
      var a = [];
      for (var i=0;i<v.length;i++) a[a.length] = jsonStringifyMini(v[i]);
      return "[" + a.join(",") + "]";
    }
    if (t === "object") {
      var parts = [];
      for (var k in v) {
        if (!v.hasOwnProperty(k)) continue;
        parts[parts.length] = '"' + jsonEscape(k) + '":' + jsonStringifyMini(v[k]);
      }
      return "{" + parts.join(",") + "}";
    }
    return "null";
  }

  function extractJsonField(body, field){
    try{
      var re = new RegExp('"' + field + '"\\s*:\\s*"([^"]*)"', "i");
      var m = (String(body||"")).match(re);
      return (m && m[1]) ? m[1] : "";
    }catch(e){}
    return "";
  }

  function parseProjects(body){
    body = String(body || "");
    var out = [];
    var re = /"id"\s*:\s*(\d+)\s*,\s*"name"\s*:\s*"([^"]*)"/g;
    var m;
    while ((m = re.exec(body)) !== null) out[out.length] = { id: String(m[1]), name: String(m[2]) };
    return out;
  }

  function parseLanguages(body){
    body = String(body || "");
    var out = [];
    var re = /"id"\s*:\s*"([^"]+)"[^}]*"name"\s*:\s*"([^"]*)"/g;
    var m;
    while ((m = re.exec(body)) !== null) out[out.length] = { id: String(m[1]), name: String(m[2]) };
    return out;
  }

  function parsePullItems(body){
    body = String(body || "");
    var out = [];
    var re = /"id"\s*:\s*"([^"]+)"[^}]*"translatedText"\s*:\s*"([^"]*)"/g;
    var m;
    while ((m = re.exec(body)) !== null) {
      var txt = m[2];
      txt = txt.replace(/\\"/g, '"').replace(/\\n/g, "\n").replace(/\\r/g, "\r").replace(/\\t/g, "\t").replace(/\\\\/g, "\\");
      out[out.length] = { id: String(m[1]), translatedText: txt };
    }
    return out;
  }

  // HTTP (curl)
  function curlGet(url){
    var TMP = Folder.temp;
    var TS  = "" + (new Date().getTime());
    var RES = new File(TMP.fsName + "/ct_get_" + TS + ".txt");
    var HTTP= new File(TMP.fsName + "/ct_get_" + TS + ".http.txt");
    var ERR = new File(TMP.fsName + "/ct_get_" + TS + ".err.txt");

    var DEVICE_ID = getDeviceId();

    var cmd = CURL + ' -4 --http1.1 --noproxy "*" -sS ' +
              '--connect-timeout 10 --max-time 60 ' +
              '-H "x-license-key: ' + TEST_LICENSE + '" ' +
              '-H "x-device-id: ' + DEVICE_ID + '" ' +
              '-H "x-device-name: ' + DEVICE_ID + '" ' +
              '-o "' + RES.fsName + '" "' + url + '" ' +
              '-w "%{http_code}" > "' + HTTP.fsName + '" 2> "' + ERR.fsName + '"';

    run(cmd);

    var http = normalizeHttpCode(readTextFile(HTTP));
    var body = readTextFile(RES);

    try{ RES.remove(); }catch(e){}
    try{ HTTP.remove(); }catch(e){}
    try{ ERR.remove(); }catch(e){}

    return { http:http, body:body };
  }

  function curlPostJson(url, jsonBody){
    var TMP = Folder.temp;
    var TS  = "" + (new Date().getTime());
    var REQ = new File(TMP.fsName + "/ct_post_" + TS + ".json");
    var RES = new File(TMP.fsName + "/ct_post_" + TS + ".txt");
    var HTTP= new File(TMP.fsName + "/ct_post_" + TS + ".http.txt");
    var ERR = new File(TMP.fsName + "/ct_post_" + TS + ".err.txt");

    if (!writeTextFile(REQ, jsonBody)) return { http:"", body:"" };

    var DEVICE_ID = getDeviceId();

    var cmd = CURL + ' -4 --http1.1 --noproxy "*" -sS ' +
              '--connect-timeout 10 --max-time 90 ' +
              '-X POST ' +
              '-H "x-license-key: ' + TEST_LICENSE + '" ' +
              '-H "x-device-id: ' + DEVICE_ID + '" ' +
              '-H "x-device-name: ' + DEVICE_ID + '" ' +
              '-H "Content-Type: application/json" ' +
              '--data-binary @"' + REQ.fsName + '" ' +
              '-o "' + RES.fsName + '" "' + url + '" ' +
              '-w "%{http_code}" > "' + HTTP.fsName + '" 2> "' + ERR.fsName + '"';

    run(cmd);

    var http = normalizeHttpCode(readTextFile(HTTP));
    var body = readTextFile(RES);

    try{ REQ.remove(); }catch(e){}
    try{ RES.remove(); }catch(e){}
    try{ HTTP.remove(); }catch(e){}
    try{ ERR.remove(); }catch(e){}

    return { http:http, body:body };
  }

  function curlPostMultipart(url, fields, files){
    var TMP = Folder.temp;
    var TS  = "" + (new Date().getTime());
    var RES = new File(TMP.fsName + "/ct_mp_" + TS + ".txt");
    var HTTP= new File(TMP.fsName + "/ct_mp_" + TS + ".http.txt");
    var ERR = new File(TMP.fsName + "/ct_mp_" + TS + ".err.txt");

    var DEVICE_ID = getDeviceId();

    var cmd = CURL + ' -4 --http1.1 --noproxy "*" -sS ' +
              '--connect-timeout 10 --max-time 180 ' +
              '-X POST ' +
              '-H "x-license-key: ' + TEST_LICENSE + '" ' +
              '-H "x-device-id: ' + DEVICE_ID + '" ' +
              '-H "x-device-name: ' + DEVICE_ID + '" ';

    var i;
    for (i=0;i<fields.length;i++){
      cmd += '-F "' + fields[i].name + '=' + jsonEscape(fields[i].value) + '" ';
    }
    for (i=0;i<files.length;i++){
      var part = files[i].name + '=@' + files[i].path;
      if (files[i].filename) part += ';filename=' + files[i].filename;
      part += ';type=' + (files[i].mime||'application/octet-stream');
      cmd += '-F "' + part + '" ';
    }

    cmd += '-o "' + RES.fsName + '" "' + url + '" ' +
           '-w "%{http_code}" > "' + HTTP.fsName + '" 2> "' + ERR.fsName + '"';

    run(cmd);

    var http = normalizeHttpCode(readTextFile(HTTP));
    var body = readTextFile(RES);

    try{ RES.remove(); }catch(e){}
    try{ HTTP.remove(); }catch(e){}
    try{ ERR.remove(); }catch(e){}

    return { http:http, body:body };
  }

  // -----------------------------
  // Updates (GitHub Releases)
  // -----------------------------
  function parseSemver(v) {
    v = String(v || "").replace(/^\s+|\s+$/g, "");
    if (v.charAt(0) === "v" || v.charAt(0) === "V") v = v.substring(1);
    var m = v.match(/^(\d+)\.(\d+)\.(\d+)/);
    if (!m) return { major: 0, minor: 0, patch: 0, raw: v, ok: false };
    return { major: parseInt(m[1], 10), minor: parseInt(m[2], 10), patch: parseInt(m[3], 10), raw: v, ok: true };
  }

  function compareSemver(a, b) {
    // Returns -1 if a<b, 0 if equal, 1 if a>b
    if (!a || !a.ok) a = parseSemver(a && a.raw ? a.raw : a);
    if (!b || !b.ok) b = parseSemver(b && b.raw ? b.raw : b);
    if (a.major !== b.major) return a.major < b.major ? -1 : 1;
    if (a.minor !== b.minor) return a.minor < b.minor ? -1 : 1;
    if (a.patch !== b.patch) return a.patch < b.patch ? -1 : 1;
    return 0;
  }

  // GitHub update requests should not include proxy/license headers.
  function curlGetPlain(url, extraHeaders) {
    var TMP = Folder.temp;
    var TS  = "" + (new Date().getTime());
    var RES = new File(TMP.fsName + "/ct_uget_" + TS + ".txt");
    var HTTP= new File(TMP.fsName + "/ct_uget_" + TS + ".http.txt");
    var ERR = new File(TMP.fsName + "/ct_uget_" + TS + ".err.txt");

    var cmd = CURL + ' -4 --http1.1 --noproxy "*" -sS ' +
              '--connect-timeout 10 --max-time 60 ';
    if (extraHeaders && extraHeaders.length) {
      for (var i = 0; i < extraHeaders.length; i++) {
        cmd += '-H "' + String(extraHeaders[i]).replace(/"/g, '\\"') + '" ';
      }
    }
    cmd += '-o "' + RES.fsName + '" "' + url + '" ' +
           '-w "%{http_code}" > "' + HTTP.fsName + '" 2> "' + ERR.fsName + '"';

    run(cmd);

    var http = normalizeHttpCode(readTextFile(HTTP));
    var body = readTextFile(RES);
    var err  = readTextFile(ERR);
    try{ RES.remove(); }catch(e){}
    try{ HTTP.remove(); }catch(e){}
    try{ ERR.remove(); }catch(e){}
    return { http: http, body: body, err: err };
  }

  function curlDownloadPlain(url, outFile, extraHeaders) {
    if (!outFile) return { http: "", err: "No outFile" };
    var TMP = Folder.temp;
    var TS  = "" + (new Date().getTime());
    var HTTP= new File(TMP.fsName + "/ct_udl_" + TS + ".http.txt");
    var ERR = new File(TMP.fsName + "/ct_udl_" + TS + ".err.txt");

    var cmd = CURL + ' -4 --http1.1 --noproxy "*" -sS -L ' +
              '--connect-timeout 10 --max-time 180 ';
    if (extraHeaders && extraHeaders.length) {
      for (var i = 0; i < extraHeaders.length; i++) {
        cmd += '-H "' + String(extraHeaders[i]).replace(/"/g, '\\"') + '" ';
      }
    }
    cmd += '-o "' + outFile.fsName + '" "' + url + '" ' +
           '-w "%{http_code}" > "' + HTTP.fsName + '" 2> "' + ERR.fsName + '"';

    run(cmd);

    var http = normalizeHttpCode(readTextFile(HTTP));
    var err  = readTextFile(ERR);
    try{ HTTP.remove(); }catch(e){}
    try{ ERR.remove(); }catch(e){}
    return { http: http, err: err };
  }

  function pickGithubReleaseAsset(releaseJson) {
    try {
      var assets = releaseJson && releaseJson.assets ? releaseJson.assets : [];
      if (!assets || !assets.length) return null;

      var i;
      for (i = 0; i < assets.length; i++) {
        var a = assets[i];
        if (!a) continue;
        if (String(a.name || "") === UPDATE_ASSET_NAME_PREFERRED) return a;
      }
      for (i = 0; i < assets.length; i++) {
        var a2 = assets[i];
        var nm = String(a2 && a2.name ? a2.name : "");
        if (nm.toLowerCase().match(/\.jsx$/)) return a2;
      }
      return null;
    } catch (e) {
      return null;
    }
  }

  function getLatestReleaseInfo() {
    var apiUrl = "https://api.github.com/repos/" + encodeURIComponent(UPDATE_GITHUB_OWNER) + "/" + encodeURIComponent(UPDATE_GITHUB_REPO) + "/releases/latest";
    var headers = [
      "Accept: application/vnd.github+json"
    ];
    var r = curlGetPlain(apiUrl, headers);
    if (r.http !== "200") return { ok: false, http: r.http, err: r.err, body: r.body };

    var j;
    try { j = JSON.parse(r.body || "{}"); } catch (e) { return { ok: false, http: r.http, err: "Invalid JSON from GitHub" }; }

    var tag = String(j.tag_name || "");
    var ver = tag || String(j.name || "");
    ver = ver.replace(/^\s+|\s+$/g, "");
    if (ver.charAt(0) === "v" || ver.charAt(0) === "V") ver = ver.substring(1);

    var asset = pickGithubReleaseAsset(j);
    if (!asset) return { ok: false, http: r.http, err: "No .jsx asset found in latest release." };

    return {
      ok: true,
      version: ver,
      tag: String(j.tag_name || ""),
      htmlUrl: String(j.html_url || ""),
      assetName: String(asset.name || ""),
      downloadUrl: String(asset.browser_download_url || "")
    };
  }

  function isValidUpdateScriptText(txt) {
    txt = String(txt || "");
    if (txt.length < 2000) return false;
    if (txt.indexOf("//@target aftereffects") === -1) return false;
    if (txt.indexOf("(function") === -1) return false;
    if (txt.indexOf("var SERVER_BASE") === -1) return false;
    return true;
  }

  function tryConfirm(msg) {
    try { return confirm(msg); } catch (e) {}
    try { return Window.confirm(msg); } catch (e2) {}
    return false;
  }

  function ensureFolder(folder) {
    try {
      if (!folder) return false;
      if (folder.exists) return true;
      return folder.create();
    } catch (e) { return false; }
  }

  function installUpdateFromRelease(info, setStatus) {
    if (!info || !info.downloadUrl) return { ok: false, reason: "Missing download URL." };
    var ver = String(info.version || "").replace(/^\s+|\s+$/g, "");
    var tmpFile = new File(Folder.temp.fsName + "/CultConnector_update_" + (ver ? ver : ("" + new Date().getTime())) + ".jsx");
    if (tmpFile.exists) { try { tmpFile.remove(); } catch (e0) {} }

    if (setStatus) setStatus("Downloading update…");
    var dl = curlDownloadPlain(info.downloadUrl, tmpFile, []);
    if (dl.http !== "200") {
      try { if (tmpFile.exists) tmpFile.remove(); } catch (eRm) {}
      return { ok: false, reason: "Download failed (HTTP " + dl.http + "). " + (dl.err || "") };
    }

    var newText = readTextFile(tmpFile);
    if (!isValidUpdateScriptText(newText)) {
      try { if (tmpFile.exists) tmpFile.remove(); } catch (eRm2) {}
      return { ok: false, reason: "Downloaded file did not look like a valid .jsx update." };
    }

    // Attempt overwrite of the currently running script file.
    var currentPath = $.fileName;
    if (!currentPath) currentPath = "";
    var currentFile = new File(currentPath);
    var canOverwrite = (currentFile && currentFile.exists);

    if (canOverwrite) {
      if (setStatus) setStatus("Installing update…");
      // Backup existing script.
      var bakPath = currentPath + ".bak";
      var bak = new File(bakPath);
      if (bak.exists) bak = new File(bakPath + "." + (new Date().getTime()));
      var oldText = readTextFile(currentFile);
      if (oldText && oldText.length > 0) {
        writeTextFile(bak, oldText);
      }

      // Overwrite
      var okWrite = writeTextFile(currentFile, newText);
      if (okWrite) {
        var verify = readTextFile(currentFile);
        if (isValidUpdateScriptText(verify)) {
          try { if (tmpFile.exists) tmpFile.remove(); } catch (eRm3) {}
          return { ok: true, installedPath: currentPath, backupPath: bak.fsName, mode: "overwrite" };
        }
      }
    }

    // Fallback: save to a writable folder for manual install.
    if (setStatus) setStatus("Saving update for manual install…");
    var outDir = new Folder(Folder.myDocuments.fsName + "/CultConnector_Update");
    if (!ensureFolder(outDir)) {
      return { ok: false, reason: "Could not create update folder in Documents." };
    }
    var outName = "CultConnector_AE_Crowdin_" + (ver ? ver : "update") + ".jsx";
    var outFile = new File(outDir.fsName + "/" + outName);
    var okOut = writeTextFile(outFile, newText);
    try { if (tmpFile.exists) tmpFile.remove(); } catch (eRm4) {}
    if (!okOut) return { ok: false, reason: "Could not write update file to Documents." };
    return { ok: true, installedPath: outFile.fsName, backupPath: "", mode: "manual" };
  }

  function runUpdateCheck(setStatus) {
    var local = parseSemver(PLUGIN_VERSION);
    if (setStatus) setStatus("Checking for updates…");
    var info = getLatestReleaseInfo();
    if (!info.ok) {
      if (setStatus) setStatus("Update check failed.");
      alertIf("Could not check for updates.\n\nHTTP " + (info.http || "?") + "\n" + (info.err || ""));
      return false;
    }

    var remote = parseSemver(info.version);
    if (!remote.ok) {
      if (setStatus) setStatus("Update check failed.");
      alertIf("Latest release version is not a valid semver: " + (info.tag || info.version));
      return false;
    }

    if (compareSemver(remote, local) <= 0) {
      if (setStatus) setStatus("Up to date (v" + PLUGIN_VERSION + ").");
      return true;
    }

    var msg = "Update available:\n\nCurrent: v" + PLUGIN_VERSION + "\nLatest:  v" + info.version + "\n\nInstall now? (After Effects restart required)";
    if (!tryConfirm(msg)) {
      if (setStatus) setStatus("Update cancelled.");
      return true;
    }

    var res = installUpdateFromRelease(info, setStatus);
    if (!res.ok) {
      if (setStatus) setStatus("Update failed.");
      alertIf("Update failed.\n\n" + (res.reason || "Unknown error."));
      return false;
    }

    if (res.mode === "overwrite") {
      if (setStatus) setStatus("Updated to v" + info.version + ". Restart After Effects.");
      alertIf("Update installed.\n\nInstalled: " + res.installedPath + (res.backupPath ? ("\nBackup: " + res.backupPath) : "") + "\n\nPlease restart After Effects.");
    } else {
      if (setStatus) setStatus("Update downloaded. Restart After Effects after replacing the script.");
      alertIf("Update downloaded.\n\nSaved to:\n" + res.installedPath + "\n\nReplace your installed ScriptUI panel .jsx with this file, then restart After Effects.");
    }
    return true;
  }

  // Build one multipart curl command string (for parallel uploads). Returns { cmd, httpPath }.
  function curlPostMultipartBuild(url, fields, files, suffix){
    var TMP = Folder.temp;
    var RES = new File(TMP.fsName + "/ct_mp_" + suffix + ".txt");
    var HTTP= new File(TMP.fsName + "/ct_mp_" + suffix + ".http.txt");
    var ERR = new File(TMP.fsName + "/ct_mp_" + suffix + ".err.txt");
    var DEVICE_ID = getDeviceId();
    var cmd = CURL + ' -4 --http2 --noproxy "*" -sS --connect-timeout 10 --max-time 120 -X POST ' +
              '-H "x-license-key: ' + TEST_LICENSE + '" -H "x-device-id: ' + DEVICE_ID + '" -H "x-device-name: ' + DEVICE_ID + '" ';
    var i;
    for (i=0;i<fields.length;i++) cmd += '-F "' + fields[i].name + '=' + jsonEscape(fields[i].value) + '" ';
    for (i=0;i<files.length;i++){
      var part = files[i].name + '=@' + files[i].path;
      if (files[i].filename) part += ';filename=' + files[i].filename;
      part += ';type=' + (files[i].mime||'application/octet-stream');
      cmd += '-F "' + part + '" ';
    }
    cmd += '-o "' + RES.fsName + '" "' + url + '" -w "%{http_code}" > "' + HTTP.fsName + '" 2> "' + ERR.fsName + '"';
    return { cmd: cmd, httpPath: HTTP.fsName };
  }

  // Run multiple upload commands in parallel (Mac/Linux: shell background; Windows: start /b then poll).
  // If options.keepBodies is true, returns { codes: string[], bodies: string[] }; otherwise returns codes only (for backward compat).
  function runParallelScanUploads(commands, options){
    var keepBodies = options && options.keepBodies;
    if (!commands.length) return keepBodies ? { codes: [], bodies: [] } : [];
    if (IS_WIN) {
      if (commands.length === 1) {
        run(commands[0].cmd);
        var httpF = new File(commands[0].httpPath);
        var code = normalizeHttpCode(readTextFile(httpF));
        var body = "";
        if (keepBodies) {
          var resF = new File(commands[0].httpPath.replace(/\.http\.txt$/, ".txt"));
          if (resF.exists) body = readTextFile(resF);
          try { if (resF.exists) resF.remove(); } catch(e2){}
        } else {
          try { var resF = new File(commands[0].httpPath.replace(/\.http\.txt$/, ".txt")); if (resF.exists) resF.remove(); } catch(e2){}
        }
        try { if (httpF.exists) httpF.remove(); } catch(e){}
        try { var errF = new File(commands[0].httpPath.replace(/\.http\.txt$/, ".err.txt")); if (errF.exists) errF.remove(); } catch(e3){}
        return keepBodies ? { codes: [code], bodies: [body] } : [code];
      }
      var batDir = Folder.temp;
      var batFiles = [];
      var i;
      for (i = 0; i < commands.length; i++) {
        var bat = new File(batDir.fsName + "/ct_scan_par_" + i + ".bat");
        writeTextFile(bat, "@echo off\r\n" + commands[i].cmd);
        batFiles.push(bat);
        run('start /b cmd /c "' + bat.fsName + '"');
      }
      var deadline = (new Date()).getTime() + 120000;
      while ((new Date()).getTime() < deadline) {
        var allDone = true;
        for (i = 0; i < commands.length; i++) {
          var hf = new File(commands[i].httpPath);
          if (!hf.exists || hf.length === 0) { allDone = false; break; }
        }
        if (allDone) break;
        $.sleep(150);
      }
      var out = [];
      var bodies = [];
      for (i = 0; i < commands.length; i++) {
        var hf = new File(commands[i].httpPath);
        out.push(normalizeHttpCode(readTextFile(hf)));
        if (keepBodies) {
          var resF = new File(commands[i].httpPath.replace(/\.http\.txt$/, ".txt"));
          bodies.push(resF.exists ? readTextFile(resF) : "");
          try { if (resF.exists) resF.remove(); } catch(e2){}
        } else {
          try { var resF = new File(commands[i].httpPath.replace(/\.http\.txt$/, ".txt")); if (resF.exists) resF.remove(); } catch(e2){}
        }
        try { if (hf.exists) hf.remove(); } catch(e){}
        try { var errF = new File(commands[i].httpPath.replace(/\.http\.txt$/, ".err.txt")); if (errF.exists) errF.remove(); } catch(e3){}
        try { if (batFiles[i].exists) batFiles[i].remove(); } catch(e4){}
      }
      return keepBodies ? { codes: out, bodies: bodies } : out;
    }
    if (commands.length === 1) {
      run(commands[0].cmd);
      var httpF = new File(commands[0].httpPath);
      var code = normalizeHttpCode(readTextFile(httpF));
      var body = "";
      if (keepBodies) {
        var resF = new File(commands[0].httpPath.replace(/\.http\.txt$/, ".txt"));
        if (resF.exists) body = readTextFile(resF);
        try { if (resF.exists) resF.remove(); } catch(e2){}
      } else {
        try { var resF = new File(commands[0].httpPath.replace(/\.http\.txt$/, ".txt")); if (resF.exists) resF.remove(); } catch(e2){}
      }
      try { if (httpF.exists) httpF.remove(); } catch(e){}
      try { var errF = new File(commands[0].httpPath.replace(/\.http\.txt$/, ".err.txt")); if (errF.exists) errF.remove(); } catch(e3){}
      return keepBodies ? { codes: [code], bodies: [body] } : [code];
    }
    var par = "";
    for (var p = 0; p < commands.length; p++) {
      if (p > 0) par += " ";
      par += "(" + commands[p].cmd + ") &";
    }
    par += " wait";
    run(par);
    var out = [];
    var bodies = [];
    for (var o = 0; o < commands.length; o++) {
      var hf = new File(commands[o].httpPath);
      var code = normalizeHttpCode(readTextFile(hf));
      out.push(code);
      if (keepBodies) {
        var resF = new File(commands[o].httpPath.replace(/\.http\.txt$/, ".txt"));
        bodies.push(resF.exists ? readTextFile(resF) : "");
        try { if (resF.exists) resF.remove(); } catch(e2){}
      } else {
        try { var resF = new File(commands[o].httpPath.replace(/\.http\.txt$/, ".txt")); if (resF.exists) resF.remove(); } catch(e2){}
      }
      try { if (hf.exists) hf.remove(); } catch(e){}
      try { var errF = new File(commands[o].httpPath.replace(/\.http\.txt$/, ".err.txt")); if (errF.exists) errF.remove(); } catch(e3){}
    }
    return keepBodies ? { codes: out, bodies: bodies } : out;
  }

  // Run one upload command in background (for pipelining: don't wait). Mac: ( cmd ) & ; Windows: .bat + start /b.
  function runUploadInBackground(cmdObj, batFilesToRemove){
    var cmd = cmdObj.cmd;
    if (IS_WIN) {
      var bat = new File(Folder.temp.fsName + "/ct_bg_" + (new Date().getTime()) + "_" + (batFilesToRemove.length) + ".bat");
      writeTextFile(bat, "@echo off\r\n" + cmd);
      batFilesToRemove.push(bat);
      run('start /b cmd /c "' + bat.fsName + '"');
    } else {
      run('( ' + cmd + ' ) &');
    }
  }

  // Wait for all background uploads to finish (poll .http.txt files), return array of http codes in same order as httpPaths.
  // If options.keepBodies is true, returns { codes: string[], bodies: string[] }; otherwise returns codes only.
  function waitForBackgroundUploads(httpPaths, timeoutMs, options){
    var keepBodies = options && options.keepBodies;
    if (!httpPaths.length) return keepBodies ? { codes: [], bodies: [] } : [];
    var deadline = (new Date()).getTime() + (timeoutMs || 120000);
    while ((new Date()).getTime() < deadline) {
      var allDone = true;
      for (var i = 0; i < httpPaths.length; i++) {
        var f = new File(httpPaths[i]);
        if (!f.exists || f.length === 0) { allDone = false; break; }
      }
      if (allDone) break;
      $.sleep(100);
    }
    var out = [];
    var bodies = [];
    for (var i = 0; i < httpPaths.length; i++) {
      var hf = new File(httpPaths[i]);
      out.push(normalizeHttpCode(readTextFile(hf)));
      if (keepBodies) {
        var resF = new File(httpPaths[i].replace(/\.http\.txt$/, ".txt"));
        bodies.push(resF.exists ? readTextFile(resF) : "");
        try { if (resF.exists) resF.remove(); } catch(e2){}
      } else {
        try { var resF = new File(httpPaths[i].replace(/\.http\.txt$/, ".txt")); if (resF.exists) resF.remove(); } catch(e2){}
      }
      try { if (hf.exists) hf.remove(); } catch(e){}
      try { var errF = new File(httpPaths[i].replace(/\.http\.txt$/, ".err.txt")); if (errF.exists) errF.remove(); } catch(e3){}
    }
    return keepBodies ? { codes: out, bodies: bodies } : out;
  }

  // OAuth connect
  function oauthConnect(setStatus){
    if (!ensureCurl()) return false;

    setStatus("Starting Crowdin login…");
    var r = curlGet(EP_OAUTH_START);

    var state = extractJsonField(r.body, "state");
    var url   = extractJsonField(r.body, "url");

    if (!state || !url) {
      setStatus("Login failed.");
      alertIf("OAuth start failed.\nHTTP " + (r.http||"(none)") + "\n\n" + (r.body||""));
      return false;
    }

    openUrl(url);
    setStatus("Waiting for authorization…");

    var startMs = (new Date()).getTime();
    var TIMEOUT = 240000;

    while (((new Date()).getTime() - startMs) < TIMEOUT) {
      $.sleep(1200);
      var s = curlGet(EP_OAUTH_STATUS + "?state=" + encodeURIComponent(state));
      if (!s.body) continue;
      if (String(s.body).toLowerCase().indexOf('"done":true') !== -1) {
        STATE.connected = true;
        setStatus("Connected.");
        return true;
      }
    }

    setStatus("Login timed out.");
    alertIf("Crowdin login timed out.\nIf browser says Connected, click Connect again.");
    return false;
  }

  function loadProjects(setStatus){
    setStatus("Loading projects…");
    var r = curlGet(EP_PROJECTS);
    if (r.http !== "200") {
      setStatus("Projects failed.");
      alertIf("Projects failed.\nHTTP " + r.http + "\n\n" + (r.body||""));
      return [];
    }
    var ps = parseProjects(r.body);
    STATE.projects = ps;
    setStatus("Projects loaded.");
    return ps;
  }

  function selectProject(projectId, projectName, setStatus){
    STATE.projectId = String(projectId);
    STATE.projectName = projectName || ("Project " + projectId);
    var payload = '{"projectId":"' + jsonEscape(STATE.projectId) + '"}';
    curlPostJson(EP_SELECT_PROJECT, payload);
    setStatus("Selected: " + STATE.projectName);
  }

  function loadLanguages(setStatus){
    if (!STATE.projectId) { setStatus("Select a project first."); return []; }
    setStatus("Loading languages…");
    var r = curlGet(EP_LANGS + "?projectId=" + encodeURIComponent(STATE.projectId));
    if (r.http !== "200") {
      setStatus("Languages failed.");
      alertIf("Languages failed.\nHTTP " + r.http + "\n\n" + (r.body||""));
      return [];
    }
    var langs = parseLanguages(r.body);
    langs.sort(function(a,b){
      var A=(a.name||a.id||"").toLowerCase(), B=(b.name||b.id||"").toLowerCase();
      if (A<B) return -1; if (A>B) return 1; return 0;
    });
    STATE.languages = langs;
    setStatus("Languages loaded.");
    return langs;
  }

  // AE collect selected text layers
  var TEXT_PROPS_MATCHNAME="ADBE Text Properties";
  var TEXT_DOC_MATCHNAME="ADBE Text Document";

  function getSourceTextProp(layer){
    if (!layer || layer.matchName !== "ADBE Text Layer") return null;
    var tp = layer.property(TEXT_PROPS_MATCHNAME);
    if (!tp) return null;
    return tp.property(TEXT_DOC_MATCHNAME) || null;
  }

  function makeStringKey(comp, layer){
    return "comp_" + comp.id + "__layer_" + layer.index;
  }

  /** Snapshot Marker: preferred screenshot time. Returns time (in layer comp) if layer has a marker with SNAPSHOT_MARKER_COMMENT, else null. Uses layer.marker or property("Marker") per AE scripting docs. */
  function getSnapshotMarkerTime(layer) {
    try {
      var prevActive = null;
      try {
        if (layer && layer.comp && app.project && app.project.activeItem !== layer.comp) {
          prevActive = app.project.activeItem;
          app.project.activeItem = layer.comp;
        }
      } catch (eSwitch) {}
      try {
        var mp = (typeof layer.marker !== "undefined" && layer.marker != null) ? layer.marker : (layer.property("Marker") || layer.property("ADBE Marker"));
        if (!mp || typeof mp.numKeys !== "number") return null;
        for (var k = 1; k <= mp.numKeys; k++) {
          var kv = mp.keyValue(k);
          if (!kv) continue;
          var cmt = (typeof kv.comment !== "undefined") ? String(kv.comment).replace(/^\s+|\s+$/g, "") : "";
          if (cmt === SNAPSHOT_MARKER_COMMENT) return mp.keyTime(k);
        }
      } finally {
        if (prevActive && app.project) { try { app.project.activeItem = prevActive; } catch (eRestore) {} }
      }
    } catch (e) {}
    return null;
  }

  /** Add or update Snapshot Marker on the given layer at the given time (in the layer's comp). Removes any existing Snapshot Marker first so there is at most one. Uses layer.marker or property("Marker"). Does not create any layers or null objects. */
  function setSnapshotMarkerAtTime(layer, time) {
    try {
      var mp = (typeof layer.marker !== "undefined" && layer.marker != null) ? layer.marker : (layer.property("Marker") || layer.property("ADBE Marker"));
      if (!mp || typeof mp.setValueAtTime !== "function") return false;
      var n = (mp.numKeys != null) ? mp.numKeys : 0;
      for (var k = n; k >= 1; k--) {
        try {
          var kv = mp.keyValue(k);
          if (!kv) continue;
          var cmt = (typeof kv.comment !== "undefined") ? String(kv.comment) : "";
          if (cmt === SNAPSHOT_MARKER_COMMENT) mp.removeKey(k);
        } catch (eKey) {}
      }
      var mv = new MarkerValue(SNAPSHOT_MARKER_COMMENT);
      if (!mv) return false;
      mp.setValueAtTime(time, mv);
      return true;
    } catch (e) { return false; }
  }

  /** True if the layer has at least one frame where it is visible in the comp. For track-matte layers, "comp" means the matte bounds. */
  function hasVisibleFrameInComp(layer, comp) {
    try {
      var a = Math.max(0, Number(layer.inPoint || 0));
      var b = Math.max(a, Number(layer.outPoint || 0));
      if (b - a < 0.01) return false;
      var step = comp.frameDuration || (1/24);
      if (!isFinite(step) || step <= 0) step = 1/24;
      var samples = Math.min(15, Math.max(1, Math.floor((b - a) / step)));
      var timesToTry = [];
      for (var i = 0; i <= samples; i++) {
        timesToTry.push(a + (b - a) * (i / Math.max(1, samples)));
      }
      // For typewriter (animator/range selector), also sample the end of the layer so we hit the fully-revealed portion
      if (b - a > 0.5) {
        for (var r = 0.8; r <= 1; r += 0.05) timesToTry.push(a + (b - a) * r);
      } else if (b - a > 0.05) {
        timesToTry.push(a + (b - a) * 0.9);
      }
      for (var idx = 0; idx < timesToTry.length; idx++) {
        var t = timesToTry[idx];
        try { if (!layer.activeAtTime(t)) continue; } catch (e) { continue; }
        if (!layerEligible(layer, t)) continue;
        var info = bboxForLayer(layer, comp, t);
        if (!info || !info.bbox) continue;
        var ref = getEffectiveBoundsForLayerAtTime(layer, comp, t);
        if (!bboxIntersectsRect(info.bbox, ref)) continue;
        if (intersectionRatioRects(info.bbox, ref) < MIN_IN_RATIO) continue;
        return true;
      }
    } catch (e) {}
    return false;
  }

  /** Get all text layers from comp c and nested precomps. Only includes layers that are visible: in their own comp (hasVisibleFrameInComp) and, if in a precomp, visible in rootComp (hasVisibleFrameInMainComp). rootComp = main comp we're exporting. */
  function getTextLayersFromComp(c, visited, rootComp, debugCollect) {
    if (!c || !(c instanceof CompItem)) return [];
    if (visited[c.id]) return [];
    visited[c.id] = true;
    var out = [];
    try {
      var layers = c.layers;
      var cName = (c.name != null) ? String(c.name) : "";
      for (var i = 1; i <= layers.length; i++) {
        var layer = layers[i];
        if (!layer) continue;
        if (layer.matchName === "ADBE Text Layer" && layer.enabled) {
          var visibleInOwnComp = hasVisibleFrameInComp(layer, c);
          if (debugCollect) debugCollect.push("  [" + cName + "] '" + (layer.name || "") + "': visibleInOwnComp=" + visibleInOwnComp);
          if (c === rootComp) {
            var inP = Number(layer.inPoint || 0);
            var outP = Number(layer.outPoint || 0);
            var hasDuration = (outP - inP) >= 0.01;
            if (visibleInOwnComp || hasDuration) {
              out[out.length] = { layer: layer, comp: c };
              if (debugCollect) debugCollect.push("    -> included (root)");
            } else if (debugCollect) debugCollect.push("    -> skipped (root, not visible)");
          } else {
            var visibleInMain = rootComp ? hasVisibleFrameInMainComp(layer, c, rootComp) : false;
            if (debugCollect) debugCollect.push("    visibleInMainComp=" + visibleInMain);
            if (visibleInOwnComp && rootComp && visibleInMain) {
              out[out.length] = { layer: layer, comp: c };
              if (debugCollect) debugCollect.push("    -> included (nested)");
            } else if (debugCollect) debugCollect.push("    -> skipped (nested)");
          }
        }
        if (layer.enabled && layer.source && (layer.source instanceof CompItem)) {
          var nested = getTextLayersFromComp(layer.source, visited, rootComp, debugCollect);
          for (var j = 0; j < nested.length; j++) out[out.length] = nested[j];
        }
      }
    } catch (e) { if (debugCollect) debugCollect.push("  getTextLayersFromComp error: " + e.toString()); }
    return out;
  }

  /** Full comp: all text layers in comp and nested precomps (enabled only, visible in comp / in main). */
  function getTextLayersIncludingPrecomps(comp, debugCollect) {
    return getTextLayersFromComp(comp, {}, comp, debugCollect || null);
  }

  /** If layers are selected: only those layers (and text inside selected precomps). If none selected: full comp. Returns [{ layer, comp }, ...]. */
  function getTextLayersForExport(comp) {
    var sel = comp.selectedLayers || [];
    var debugCollect = (typeof DEBUG_TYPEWRITER_LOG !== "undefined" && DEBUG_TYPEWRITER_LOG) ? [] : null;
    if (!sel.length) {
      var out = getTextLayersFromComp(comp, {}, comp, debugCollect);
      if (debugCollect && debugCollect.length) {
        try {
          var f = new File(Folder.myDocuments.fsName + "/Crowdin_layers_debug.txt");
          if (f.open("w")) {
            f.write("=== getTextLayersForExport (full comp: " + (comp.name || "") + ") ===\r\n");
            for (var i = 0; i < debugCollect.length; i++) f.write(debugCollect[i] + "\r\n");
            f.write("Total included: " + out.length + "\r\n");
            f.close();
          }
        } catch (e) {}
      }
      return out;
    }
    var out = [];
    var visited = {};
    if (debugCollect) {
      debugCollect.push("=== getTextLayersForExport (selected layers only) ===");
      debugCollect.push("Comp: " + (comp.name || ""));
      for (var s = 0; s < sel.length; s++) debugCollect.push("  Selected: '" + (sel[s] && sel[s].name ? sel[s].name : "") + "'");
    }
    for (var i = 0; i < sel.length; i++) {
      var L = sel[i];
      if (!L) continue;
      if (L.matchName === "ADBE Text Layer" && L.enabled && hasVisibleFrameInComp(L, comp)) {
        out[out.length] = { layer: L, comp: comp };
        if (debugCollect) debugCollect.push("  Included text: '" + (L.name || "") + "'");
      } else if (L.matchName === "ADBE Text Layer" && L.enabled && debugCollect) {
        debugCollect.push("  Skipped text (not visible in comp): '" + (L.name || "") + "'");
      }
      if (L.enabled && L.source && (L.source instanceof CompItem)) {
        var nested = getTextLayersFromComp(L.source, visited, comp, debugCollect);
        for (var j = 0; j < nested.length; j++) out[out.length] = nested[j];
      }
    }
    if (debugCollect && debugCollect.length) {
      try {
        var f = new File(Folder.myDocuments.fsName + "/Crowdin_layers_debug.txt");
        if (f.open("w")) {
          for (var k = 0; k < debugCollect.length; k++) f.write(debugCollect[k] + "\r\n");
          f.write("Total included: " + out.length + "\r\n");
          f.close();
        }
      } catch (e) {}
    }
    return out;
  }

  function getSelectedTextLayers(comp){
    var sel = comp.selectedLayers || [];
    var out = [];
    for (var i=0;i<sel.length;i++){
      if (sel[i] && sel[i].matchName === "ADBE Text Layer") out[out.length] = sel[i];
    }
    return out;
  }

  function collectText(setStatus, comp){
    comp = (comp != null && comp instanceof CompItem) ? comp : null;
    if (!comp) return null;

    var layerComps = getTextLayersForExport(comp);
    if (!layerComps.length) {
      setStatus("No text layers in selection or comp.");
      alertIf("Select one or more TEXT or PRECOMP layers to send only those, or leave nothing selected to send the entire composition.");
      return null;
    }

    var items = [];
    for (var i=0;i<layerComps.length;i++){
      var L = layerComps[i].layer;
      var C = layerComps[i].comp;
      var sp = getSourceTextProp(L);
      if (!sp) continue;
      var txt = getCompletedTextForLayer(L, C);
      if (!txt) continue;
      items[items.length] = { id: makeStringKey(C, L), text: txt };
    }

    STATE.compId = "" + comp.id;
    STATE.fileKey = safeFileKeyForComp(comp);
    // Store comp name in first item so upload always has it (only when we have at least one item)
    if (items.length > 0) {
      try {
        var compNameRaw = (comp && comp.name != null) ? String(comp.name) : "";
        var compName = compNameRaw.replace(/^\s+|\s+$/g, ""); // trim without .trim() for ExtendScript
        if (compName.length > 0) {
          var safe = compName.replace(/[^A-Za-z0-9_.-]+/g, "_").replace(/_+/g, "_").replace(/^_|_$/g, "");
          if (safe.length > 0 && safe.length <= 200) items[0].__compName = safe;
        }
      } catch (e) { /* skip __compName on any error so collection still works */ }
    }

    setStatus("Collected " + items.length + " strings.");
    return items;
  }

  function uploadStrings(items, setStatus, targetLang){
    if (!items || !items.length) return false;
    if (!STATE.projectId) { alertIf("Select a project first."); return false; }

    // Ensure we have comp id from the items we're uploading (e.g. "comp_21__layer_1" -> 21)
    if (!STATE.compId && items[0] && items[0].id) {
      var match = String(items[0].id).match(/^comp_(\d+)__/);
      if (match) STATE.compId = match[1];
    }

    // Prefer comp name embedded at collect time (most reliable)
    if (items[0] && items[0].__compName && String(items[0].__compName).length > 0) {
      STATE.fileKey = String(items[0].__compName);
    }
    // Else discover composition name from comp id (avoid getActiveComp so multi-comp export does not alert).
    if (!STATE.fileKey || STATE.fileKey === "comp_" + (STATE.compId || "")) {
      var foundComp = findCompById(STATE.compId);
      if (foundComp) STATE.fileKey = safeFileKeyForComp(foundComp);
    }
    if (!STATE.fileKey) STATE.fileKey = "comp_" + (STATE.compId || "");

    setStatus("Uploading strings…");

    var TMP = Folder.temp;
    // Write file with composition name so curl sends that filename
    var exportFileName = STATE.fileKey + ".json";
    var fStrings = new File(TMP.fsName + "/" + exportFileName);

    // Strip internal __compName before sending; send fileKey inside JSON so server has the composition name.
    var itemsToSend = [];
    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      var copy = { id: it.id, text: it.text };
      itemsToSend.push(copy);
    }
    var payloadObj = { fileKey: STATE.fileKey, items: itemsToSend };
    var payload = jsonStringifyMini(payloadObj);
    if (!writeTextFile(fStrings, payload)) {
      setStatus("Failed writing temp strings file.");
      return false;
    }

    var resp = curlPostMultipart(
      EP_STRINGS,
      [
        { name:"projectId", value: STATE.projectId },
        { name:"compId", value: STATE.compId },
        { name:"fileKey", value: STATE.fileKey },
        // Pass segmentation preference explicitly (1/0) for compatibility with existing API.
        { name:"useSegmentation", value: STATE.useSegmentation ? "1" : "0" }
      ].concat((targetLang && targetLang !== "all") ? [{ name:"targetLanguage", value: targetLang }] : []),
      [
        { name:"strings", path: fStrings.fsName, mime:"application/json" }
      ]
    );

    try{ fStrings.remove(); }catch(e){}

    if (resp.http !== "200") {
      setStatus("Upload failed.");
      alertIf("Upload strings failed.\nHTTP " + resp.http + "\n\n" + (resp.body||""));
      return false;
    }

    var statusMsg = "Uploaded!";
    try {
      var r = JSON.parse(resp.body || "{}");
      if (r.fileName) {
        var displayName = (r.displayFileName != null && String(r.displayFileName).length > 0) ? String(r.displayFileName) : String(r.fileName).replace(/\.(json|xml)$/i, "");
        statusMsg = "Uploaded as " + displayName;
      }
      if (r._receivedFilename != null) statusMsg += " (received: " + String(r._receivedFilename).replace(/\.(json|xml)$/i, "") + ")";
    } catch(e){}
    setStatus(statusMsg);
    return true;
  }

  // =========================
  // SMART SCAN (strict + fallback + rescue bbox)
  // =========================

  function layerOpacityAt(layer, t){
    try{
      var tr = layer.property("ADBE Transform Group");
      if (!tr) return 100;
      var op = tr.property("ADBE Opacity");
      if (!op) return 100;
      return Number(op.valueAtTime(t, false)); // 0..100
    }catch(e){ return 100; }
  }

  /** Scale at time t (min of X,Y). Returns 1 if no scale or error. Do not use || 100 so that scale 0 is preserved. */
  function layerScaleAt(layer, t){
    try{
      var tr = layer.property("ADBE Transform Group");
      if (!tr) return 1;
      var scaleProp = tr.property("ADBE Scale");
      if (!scaleProp) return 1;
      var v = scaleProp.valueAtTime(t, false);
      if (!v || v.length < 2) return 1;
      var sx = Number(v[0]);
      var sy = Number(v[1]);
      if (!isFinite(sx)) sx = 100;
      if (!isFinite(sy)) sy = 100;
      return Math.min(sx, sy) / 100;
    }catch(e){ return 1; }
  }

  function layerEligible(layer, t){
    try{
      if (!layer.enabled) return false;
      if (t < layer.inPoint || t > layer.outPoint) return false;
      if (layerOpacityAt(layer, t) < MIN_OPACITY) return false;
      if (layerScaleAt(layer, t) < MIN_SCALE) return false;
      return true;
    }catch(e){ return false; }
  }

  function toCompSafe(layer, xy){
    try{
      if (layer.threeDLayer) return layer.toComp([xy[0], xy[1], 0]);
      return layer.toComp([xy[0], xy[1]]);
    }catch(e){
      return null;
    }
  }

  /** Transform a point from layer's source comp space to layer's containing comp space at time t, using valueAtTime (does not set comp.time). */
  function layerPointToCompAtTime(layer, point, t) {
    try {
      var tr = layer.property("ADBE Transform Group");
      if (!tr) return null;
      var pos = tr.property("ADBE Position");
      var anc = tr.property("ADBE Anchor Point");
      var scl = tr.property("ADBE Scale");
      var rot = tr.property("ADBE Rotate Z");
      if (!pos || !anc || !scl || !rot) return null;
      var pv = pos.valueAtTime(t, false);
      var av = anc.valueAtTime(t, false);
      var sv = scl.valueAtTime(t, false);
      var rv = rot.valueAtTime(t, false);
      if (!pv || pv.length < 2) return null;
      var px = Number(pv[0]), py = Number(pv[1]);
      var ax = av && av.length >= 2 ? Number(av[0]) : 0;
      var ay = av && av.length >= 2 ? Number(av[1]) : 0;
      var sx = (sv && sv[0] != null) ? Number(sv[0]) / 100 : 1;
      var sy = (sv && sv[1] != null) ? Number(sv[1]) / 100 : 1;
      var r = (rv != null) ? Number(rv) * Math.PI / 180 : 0;
      var dx = point[0] - ax, dy = point[1] - ay;
      dx *= sx; dy *= sy;
      var cos = Math.cos(r), sin = Math.sin(r);
      return [px + (dx * cos - dy * sin), py + (dx * sin + dy * cos)];
    } catch (e) { return null; }
  }

  function pointInComp(comp, p, margin){
    margin = (margin == null) ? 10 : margin;
    if (!comp || !p) return false;
    return (p[0] >= -margin && p[0] <= comp.width + margin && p[1] >= -margin && p[1] <= comp.height + margin);
  }

  function originInComp(layer, comp){
    var p0 = toCompSafe(layer, [0,0]);
    return (p0 && pointInComp(comp, p0, 10));
  }

  /** Position [x, y] at time t (layer space). Returns null if not available. */
  function getLayerPositionAtTime(layer, t) {
    try {
      var tr = layer.property("ADBE Transform Group");
      if (!tr) return null;
      var pos = tr.property("ADBE Position");
      if (!pos) return null;
      var v = pos.valueAtTime(t, false);
      if (!v || v.length < 2) return null;
      return [Number(v[0]), Number(v[1])];
    } catch (e) { return null; }
  }

  /**
   * True if the layer is "paused" at t: position (and scale) are stable over a short window.
   * Prefer the first moment of a hold: require stability forward (t to t+step). Only require
   * stability backward (t-step to t) when t-step is within layer bounds, so we don't fail at the very first frame of a hold.
   */
  var PAUSE_POSITION_TOLERANCE = 1;
  var PAUSE_SCALE_TOLERANCE = 0.005;
  function isLayerPausedAt(layer, comp, t) {
    try {
      var step = comp.frameDuration || (1 / 24);
      if (!isFinite(step) || step <= 0) step = 1 / 24;
      var tPrev = t - step;
      var tNext = t + step;
      if (tPrev < layer.inPoint) tPrev = layer.inPoint;
      if (tNext > layer.outPoint) tNext = layer.outPoint;
      var p0 = getLayerPositionAtTime(layer, t);
      if (!p0) return false;
      var pPrev = getLayerPositionAtTime(layer, tPrev);
      var pNext = getLayerPositionAtTime(layer, tNext);
      // Always require stable going forward (so we're at the start of a hold)
      if (pNext && (Math.abs(p0[0] - pNext[0]) > PAUSE_POSITION_TOLERANCE || Math.abs(p0[1] - pNext[1]) > PAUSE_POSITION_TOLERANCE)) return false;
      // Require stable from previous frame only when we have a valid previous frame (not at layer start)
      if (tPrev >= layer.inPoint && pPrev && (Math.abs(p0[0] - pPrev[0]) > PAUSE_POSITION_TOLERANCE || Math.abs(p0[1] - pPrev[1]) > PAUSE_POSITION_TOLERANCE)) return false;
      var s0 = layerScaleAt(layer, t);
      var sPrev = layerScaleAt(layer, tPrev);
      var sNext = layerScaleAt(layer, tNext);
      if (Math.abs(s0 - sNext) > PAUSE_SCALE_TOLERANCE) return false;
      if (tPrev >= layer.inPoint && Math.abs(s0 - sPrev) > PAUSE_SCALE_TOLERANCE) return false;
      return true;
    } catch (e) { return false; }
  }

  /** Transform a point from layer's source comp space to layer's containing comp space at time t, using valueAtTime (does not set comp.time). */
  function layerPointToCompAtTime(layer, point, t) {
    try {
      var tr = layer.property("ADBE Transform Group");
      if (!tr) return null;
      var pos = tr.property("ADBE Position");
      var anc = tr.property("ADBE Anchor Point");
      var scl = tr.property("ADBE Scale");
      var rot = tr.property("ADBE Rotate Z");
      if (!pos || !anc || !scl || !rot) return null;
      var pv = pos.valueAtTime(t, false);
      var av = anc.valueAtTime(t, false);
      var sv = scl.valueAtTime(t, false);
      var rv = rot.valueAtTime(t, false);
      if (!pv || pv.length < 2) return null;
      var px = Number(pv[0]), py = Number(pv[1]);
      var ax = av && av.length >= 2 ? Number(av[0]) : 0;
      var ay = av && av.length >= 2 ? Number(av[1]) : 0;
      var sx = (sv && sv[0] != null) ? Number(sv[0]) / 100 : 1;
      var sy = (sv && sv[1] != null) ? Number(sv[1]) / 100 : 1;
      var r = (rv != null) ? Number(rv) * Math.PI / 180 : 0;
      var dx = point[0] - ax, dy = point[1] - ay;
      dx *= sx; dy *= sy;
      var cos = Math.cos(r), sin = Math.sin(r);
      return [px + (dx * cos - dy * sin), py + (dx * sin + dy * cos)];
    } catch (e) { return null; }
  }

  /** Position offset [dx, dy] in layer space from Transform/Position effects at time t (so bbox matches rendered text). Returns null if none. */
  function getEffectPositionOffsetAtTime(layer, t) {
    try {
      var parade = layer.property("ADBE Effect Parade");
      if (!parade || !parade.numProperties) return null;
      function findPositionOffset(grp) {
        if (!grp || !grp.numProperties) return null;
        for (var i = 1; i <= grp.numProperties; i++) {
          var p = grp.property(i);
          if (!p) continue;
          var posProp = null, anchorProp = null;
          if (p.numProperties != null && p.numProperties > 0) {
            for (var j = 1; j <= p.numProperties; j++) {
              var sub = p.property(j);
              if (!sub) continue;
              var name = (sub.name || "").toString();
              var mn = (sub.matchName || "").toString();
              if (name === "Position" || mn.indexOf("Position") >= 0) posProp = sub;
              if (name === "Anchor Point" || mn.indexOf("Anchor") >= 0) anchorProp = sub;
            }
            if (posProp) {
              var pv = posProp.valueAtTime(t, false);
              if (!pv || pv.length < 2) continue;
              var dx = Number(pv[0]), dy = Number(pv[1]);
              if (anchorProp) {
                try {
                  var av = anchorProp.valueAtTime(t, false);
                  if (av && av.length >= 2) {
                    dx -= Number(av[0]);
                    dy -= Number(av[1]);
                  }
                } catch (eA) {}
              }
              return [dx, dy];
            }
            var fromGroup = findPositionOffset(p);
            if (fromGroup) return fromGroup;
          }
        }
        return null;
      }
      return findPositionOffset(parade);
    } catch (e) { return null; }
  }

  /** Single rule for screenshot time: second keyframe if any property has >= 2 keyframes, else midpoint of layer. Returns time in [a,b]. allowSnapshotMarker: when true (default), Snapshot Marker is used for main comp layers; when false, only keyframe/midpoint (for precomp layers). */
  function getScreenshotTimeForLayer(layer, a, b, allowSnapshotMarker) {
    if (allowSnapshotMarker === undefined) allowSnapshotMarker = true;
    var list = getPreferredScreenshotTimes(layer, a, b, allowSnapshotMarker);
    return (list && list.length > 0) ? list[0] : ((Number(a) + Number(b)) / 2);
  }

  /** All preferred times for this layer: second keyframes first, then all other keyframe times + midpoint, sorted, clamped to [a,b]. allowSnapshotMarker: only use Snapshot Marker when true (main comp); precomp layers pass false. */
  function getPreferredScreenshotTimes(layer, a, b, allowSnapshotMarker) {
    if (allowSnapshotMarker === undefined) allowSnapshotMarker = true;
    try {
      var aNum = Number(a);
      var bNum = Number(b);
      if (!isFinite(aNum) || !isFinite(bNum) || bNum - aNum < 0.01) return [(aNum + bNum) / 2];
      var midpoint = (aNum + bNum) / 2;
      var step = 1 / 24;
      try { if (layer.comp) step = layer.comp.frameDuration || step; } catch (e) {}

      var primary = midpoint;
      var primarySet = false;
      if (allowSnapshotMarker) {
        var snapshotTime = getSnapshotMarkerTime(layer);
        if (snapshotTime != null && isFinite(snapshotTime) && snapshotTime >= aNum - 0.001 && snapshotTime <= bNum + 0.001) {
          primary = snapshotTime;
          primarySet = true;
        }
      }

      // Precomp layers: set active comp to the layer's comp so keyframe reads (Transform, Effect Parade, Scale) work correctly.
      var layerComp = null;
      var prevActive = null;
      try {
        layerComp = layer.comp;
        if (layerComp && app.project && app.project.activeItem !== layerComp) {
          prevActive = app.project.activeItem;
          app.project.activeItem = layerComp;
          try { layerComp.time = aNum; } catch (eTime) {}
        }
      } catch (eSwitch) {}

      function allKeyTimesFromProp(prop, out) {
        if (!prop) return;
        try {
          var n = (prop.numKeys != null) ? prop.numKeys : 0;
          for (var k = 1; k <= n; k++) {
            var kt = prop.keyTime(k);
            if (kt != null && isFinite(kt)) out.push(kt);
          }
        } catch (e) {}
        if (prop.numProperties != null && prop.numProperties >= 1) {
          for (var d = 1; d <= Math.min(prop.numProperties, 5); d++) {
            try {
              var dim = prop.property(d);
              if (dim && dim.numKeys != null) {
                for (var kd = 1; kd <= dim.numKeys; kd++) {
                  var ktd = dim.keyTime(kd);
                  if (ktd != null && isFinite(ktd)) out.push(ktd);
                }
              }
            } catch (ed) {}
          }
        }
      }
      function collectAllKeyTimes(grp, out) {
        if (!grp || !grp.numProperties) return;
        for (var i = 1; i <= grp.numProperties; i++) {
          try {
            var p = grp.property(i);
            if (!p) continue;
            allKeyTimesFromProp(p, out);
            if (p.numProperties != null && p.numProperties > 0) collectAllKeyTimes(p, out);
          } catch (e) {}
        }
      }
      var times = [];
      var transform = layer.property("ADBE Transform Group");
      if (transform) collectAllKeyTimes(transform, times);
      var parade = layer.property("ADBE Effect Parade");
      if (parade) collectAllKeyTimes(parade, times);
      if (layer.matchName === "ADBE Text Layer") {
        var tp = layer.property("ADBE Text Properties");
        if (tp) collectAllKeyTimes(tp, times);
      }
      times.push(midpoint);
      // Simple typewriter (no keyframes): add tail times so multiple typewriter layers each get a slot when text is fully revealed.
      if (times.length <= 1) {
        for (var tail = 0.7; tail <= 1; tail += 0.1) times.push(aNum + (bNum - aNum) * tail);
        times.push(bNum);
      }
      var seen = {};
      var unique = [];
      for (var j = 0; j < times.length; j++) {
        var t = Math.max(aNum, Math.min(bNum, times[j]));
        var k = Math.round(t / step).toString();
        if (!seen[k]) { seen[k] = true; unique.push(t); }
      }
      unique.sort(function (x, y) { return x - y; });
      if (unique.length === 0) return [midpoint];
      // Keep primary/primarySet from Snapshot Marker above; only set from second keyframe or fallback when not already set.
      function keyTime2FromProp(prop) {
        if (!prop) return null;
        try {
          if (prop.numKeys != null && prop.numKeys >= 2) return prop.keyTime(2);
        } catch (e) {}
        if (prop.numProperties != null && prop.numProperties >= 1) {
          for (var d = 1; d <= Math.min(prop.numProperties, 5); d++) {
            try {
              var dim = prop.property(d);
              if (dim && dim.numKeys != null && dim.numKeys >= 2) return dim.keyTime(2);
            } catch (ed) {}
          }
        }
        return null;
      }
      // Collect every property's keyTime(2) from Transform, Effect Parade, Text (recursive).
      function collectAllKeyTime2(grp, out) {
        if (!grp || !grp.numProperties) return;
        for (var i = 1; i <= grp.numProperties; i++) {
          try {
            var p = grp.property(i);
            if (!p) continue;
            var kt = keyTime2FromProp(p);
            if (kt != null && isFinite(kt)) out.push(kt);
            if (p.numProperties > 0) collectAllKeyTime2(p, out);
          } catch (e) {}
        }
      }
      var allKeyTime2 = [];
      if (transform) collectAllKeyTime2(transform, allKeyTime2);
      if (parade) collectAllKeyTime2(parade, allKeyTime2);
      if (layer.matchName === "ADBE Text Layer") {
        var tpp = layer.property("ADBE Text Properties");
        if (tpp) collectAllKeyTime2(tpp, allKeyTime2);
      }
      // Rule: use the second keyframe that appears first on the timeline (earliest keyTime(2)).
      if (!primarySet && allKeyTime2.length > 0) {
        var earliest = allKeyTime2[0];
        for (var ei = 1; ei < allKeyTime2.length; ei++) {
          if (allKeyTime2[ei] < earliest) earliest = allKeyTime2[ei];
        }
        primary = Math.max(aNum, Math.min(bNum, earliest));
        primarySet = true;
      }
      if (!primarySet) {
        primary = aNum + (bNum - aNum) * 0.9;
        if (primary > bNum) primary = bNum;
        if (primary < aNum) primary = aNum;
      }
      var result = [primary];
      var others = [];
      for (var r = 0; r < unique.length; r++) {
        if (Math.abs(unique[r] - primary) > 0.001) others.push(unique[r]);
      }
      others.sort(function (x, y) { return x - y; });
      for (var o = 0; o < others.length; o++) result.push(others[o]);
      return result.length ? result : [midpoint];
    } catch (e) {
      var an = Number(a), bn = Number(b);
      return [(isFinite(an) && isFinite(bn)) ? (an + bn) / 2 : 0];
    } finally {
      if (prevActive && app.project) {
        try { app.project.activeItem = prevActive; } catch (eRestore) {}
      }
    }
  }

  function getSecondKeyframeTimeForEffectPosition(layer) {
    try {
      var parade = layer.property("ADBE Effect Parade");
      if (!parade || !parade.numProperties) return null;
      function keyTime2FromProp(prop) {
        if (!prop) return null;
        try {
          if (prop.numKeys != null && prop.numKeys >= 2) return prop.keyTime(2);
        } catch (e) {}
        if (prop.numProperties != null && prop.numProperties >= 1) {
          for (var d = 1; d <= Math.min(prop.numProperties, 3); d++) {
            var dim = prop.property(d);
            if (dim && dim.numKeys != null && dim.numKeys >= 2) {
              try {
                return dim.keyTime(2);
              } catch (ed) {}
            }
          }
        }
        return null;
      }
      function findPosPropSecondKey(grp) {
        if (!grp || !grp.numProperties) return null;
        for (var i = 1; i <= grp.numProperties; i++) {
          var p = grp.property(i);
          if (!p) continue;
          if (p.numProperties != null && p.numProperties > 0) {
            for (var j = 1; j <= p.numProperties; j++) {
              var sub = p.property(j);
              if (!sub) continue;
              var name = (sub.name || "").toString();
              var mn = (sub.matchName || "").toString();
              if (name === "Position" || mn.indexOf("Position") >= 0) {
                var kt2 = keyTime2FromProp(sub);
                if (kt2 != null && isFinite(kt2)) return kt2;
              }
            }
            var fromGroup = findPosPropSecondKey(p);
            if (fromGroup != null) return fromGroup;
          }
        }
        return null;
      }
      return findPosPropSecondKey(parade);
    } catch (e) { return null; }
  }

  /** Source rect in comp space (layer transform only). Used for bbox/highlight. */
  function bboxFromSourceRect(layer, t) {
    try{
      var r = null;
      try { r = layer.sourceRectAtTime(t, false); } catch(e0){ r = null; }
      if (!r || r.width <= 0.1 || r.height <= 0.1) { try { r = layer.sourceRectAtTime(t, true); } catch(e1){ r = null; } }
      if (!r || r.width <= 0.1 || r.height <= 0.1) return null;

      var x1 = r.left, y1 = r.top;
      var x2 = r.left + r.width, y2 = r.top + r.height;

      var offset = getEffectPositionOffsetAtTime(layer, t);
      if (offset && offset.length >= 2) {
        x1 += Number(offset[0]); y1 += Number(offset[1]);
        x2 += Number(offset[0]); y2 += Number(offset[1]);
      }

      var p1 = layerPointToCompAtTime(layer, [x1, y1], t);
      var p2 = layerPointToCompAtTime(layer, [x2, y1], t);
      var p3 = layerPointToCompAtTime(layer, [x1, y2], t);
      var p4 = layerPointToCompAtTime(layer, [x2, y2], t);

      if (!p1 || !p2 || !p3 || !p4) return null;

      var minX = Math.min(p1[0], p2[0], p3[0], p4[0]);
      var maxX = Math.max(p1[0], p2[0], p3[0], p4[0]);
      var minY = Math.min(p1[1], p2[1], p3[1], p4[1]);
      var maxY = Math.max(p1[1], p2[1], p3[1], p4[1]);

      var w = maxX - minX, h = maxY - minY;
      if (w < 0.1 || h < 0.1) return null;

      if (BBOX_TIGHTEN_RATIO < 1 && w >= 4 && h >= 4) {
        var cx = (minX + maxX) / 2, cy = (minY + maxY) / 2;
        var halfW = w / 2, halfH = h / 2;
        halfW *= BBOX_TIGHTEN_RATIO;
        halfH *= BBOX_TIGHTEN_RATIO;
        minX = cx - halfW; maxX = cx + halfW;
        minY = cy - halfH; maxY = cy + halfH;
        w = maxX - minX; h = maxY - minY;
      }

      return { x: minX, y: minY, w: w, h: h };
    }catch(e){
      return null;
    }
  }

  function bboxEstimateFromTextDoc(layer, t){
    try{
      var sp = getSourceTextProp(layer);
      if (!sp) return null;

      var doc = sp.valueAtTime(t, false);
      var text = String(doc.text || "");
      text = text.replace(/\r/g, "\n");

      var lines = text.split("\n");
      var lineCount = Math.max(1, lines.length);

      var maxLen = 1;
      for (var i=0;i<lines.length;i++){
        if (lines[i].length > maxLen) maxLen = lines[i].length;
      }

      var fs = Number(doc.fontSize || 40);
      if (!isFinite(fs) || fs <= 0) fs = 40;

      var w = clamp(fs * 0.52 * maxLen, 40, 1600);
      var h = clamp(fs * 1.22 * lineCount, 30, 900);

      var p = toCompSafe(layer, [0,0]);
      if (!p) return null;

      return { x: p[0] - w/2, y: p[1] - h/2, w: w, h: h };
    }catch(e){
      return null;
    }
  }

  /** Like bboxEstimateFromTextDoc but uses the completed/full text (cursor stripped, max length over layer)
   * so the highlight matches the fully formed text. Use when capturing at typewriter full-reveal with
   * blinking cursor. When comp is passed, dimensions are based on getCompletedTextForLayer so the box
   * is never minimal (e.g. when source at t is just "|"). */
  function bboxEstimateFromTextDocNoCursor(layer, t, comp) {
    try {
      var sp = getSourceTextProp(layer);
      if (!sp) return null;

      var doc = sp.valueAtTime(t, false);
      var fs = Number(doc.fontSize || 40);
      if (!isFinite(fs) || fs <= 0) fs = 40;

      var text;
      if (comp) {
        text = getCompletedTextForLayer(layer, comp);
        if (!text) text = String(doc.text || "");
      } else {
        text = String(doc.text || "");
        text = stripBlinkingCursorCursor(text);
      }
      text = text.replace(/\r/g, "\n");
      var lines = text.split("\n");
      var lineCount = Math.max(1, lines.length);
      var maxLen = 1;
      for (var i = 0; i < lines.length; i++) {
        if (lines[i].length > maxLen) maxLen = lines[i].length;
      }

      var w = clamp(fs * 0.52 * maxLen, 40, 1600);
      var h = clamp(fs * 1.22 * lineCount, 30, 900);

      var p = toCompSafe(layer, [0, 0]);
      if (!p) return null;

      return { x: p[0] - w / 2, y: p[1] - h / 2, w: w, h: h };
    } catch (e) {
      return null;
    }
  }

  // ✅ NEW: last-resort bbox from Position (very reliable)
  function bboxFromPositionRescue(layer, comp, t){
    try{
      var tr = layer.property("ADBE Transform Group");
      if (!tr) return null;
      var pos = tr.property("ADBE Position");
      if (!pos) return null;

      var v = pos.valueAtTime(t, false);
      if (!v || v.length < 2) return null;

      var x = Number(v[0]), y = Number(v[1]);
      if (!isFinite(x) || !isFinite(y)) return null;

      var w = Math.min(RESCUE_BBOX_W, comp.width);
      var h = Math.min(RESCUE_BBOX_H, comp.height);

      return { x: x - w/2, y: y - h/2, w: w, h: h };
    }catch(e){
      return null;
    }
  }

  // Detect native Blinking Cursor Typewriter preset by inspecting the Source Text expression
  // (more reliable than effect names). Match multiple possible phrasings so all layers get the fix.
  function hasBlinkingCursorTypewriterEffect(layer) {
    try {
      if (!layer || layer.matchName !== "ADBE Text Layer") return false;
      var sp = getSourceTextProp(layer);
      if (!sp) return false;
      var expr = "";
      try { expr = String(sp.expression || ""); } catch (eExpr) { expr = ""; }
      if (!expr) return false;
      var e = expr.replace(/\s+/g, " ");
      if (e.indexOf("Blinking Cursor Typewriter") >= 0) return true;
      if (e.indexOf("Cursor Shape") >= 0 && (e.indexOf("Animation") >= 0 || e.indexOf("linear(") >= 0) && e.indexOf("reveal") >= 0) return true;
      if (e.indexOf("effect(") >= 0 && e.indexOf("Animation") >= 0 && (e.indexOf("slice(") >= 0 || e.indexOf("reveal") >= 0)) return true;
    } catch (eOuter) {}
    return false;
  }

  // For Blinking Cursor preset, measure full-text bbox at t WITHOUT modifying any effect (read-only).
  // Uses cached full-text dimensions per layer; position comes from sourceRect at t (left/top of visible text) so highlight aligns correctly.
  function getBlinkFullTextBboxReadOnly(layer, comp, tCapture) {
    try {
      if (!layer || layer.matchName !== "ADBE Text Layer") return null;
      if (!hasBlinkingCursorTypewriterEffect(layer)) return null;

      var ck = layer.id;
      if (__blinkFullTextSizeCache[ck] === undefined) {
        var fullBox = bboxEstimateFromTextDocNoCursor(layer, Math.max(0, Number(layer.inPoint || 0)), comp);
        __blinkFullTextSizeCache[ck] = fullBox ? { w: fullBox.w, h: fullBox.h } : null;
      }
      var cached = __blinkFullTextSizeCache[ck];
      if (!cached || cached.w < 2 || cached.h < 2) return null;

      var a = Math.max(0, Number(layer.inPoint || 0));
      var b = Math.max(a, Number(layer.outPoint || 0));
      var t = (tCapture != null && tCapture >= a && tCapture <= b) ? tCapture : (a + (b - a) * 0.8);

      // Use visible text bounds at t for position (left/top) so highlight aligns with actual text, not anchor.
      var bbSource = bboxFromSourceRect(layer, t);
      if (!bbSource) return null;

      var x = bbSource.x;
      var y = bbSource.y;
      var w = cached.w;
      var h = cached.h;
      // At full-reveal frame, sourceRect may already be the full text; use it when it's close to full size.
      if (bbSource.w >= cached.w * 0.85 && bbSource.h >= cached.h * 0.85) {
        w = bbSource.w;
        h = bbSource.h;
      }
      x -= Math.min(w * 0.08, 12);
      if (x < 0) x = 0;
      return { x: x, y: y, w: w, h: h };
    } catch (eOuter) {
      return null;
    }
  }

  // For Blinking Cursor preset, measure full-text geometry at capture time t with the
  // \"Animation\" control temporarily forced to 100, so sourceRect reflects the full text
  // and the bbox position matches the frame we're capturing.
  function getBlinkFullTextBbox(layer, comp, tCapture) {
    try {
      if (!layer || layer.matchName !== "ADBE Text Layer") return null;
      if (!hasBlinkingCursorTypewriterEffect(layer)) return null;
      var eff = null;
      try { eff = layer.effect("Animation"); } catch (e0) {}
      if (!eff) {
        try {
          var parade = layer.property("ADBE Effect Parade");
          if (parade && parade.numProperties) {
            for (var i = 1; i <= parade.numProperties; i++) {
              var e = null;
              try { e = parade.property(i); } catch (e1) {}
              if (e && e.name && String(e.name) === "Animation") { eff = e; break; }
            }
          }
        } catch (eFind) {}
      }
      if (!eff) {
        try {
          var animIn = layer.effect("Animate In");
          if (animIn && animIn.numProperties) {
            for (var k = 1; k <= animIn.numProperties; k++) {
              var sub = null;
              try { sub = animIn.property(k); } catch (eSub) {}
              if (!sub || !sub.numProperties) continue;
              for (var j = 1; j <= sub.numProperties; j++) {
                var prop = null;
                try { prop = sub.property(j); } catch (eP2) {}
                if (prop && prop.name && String(prop.name) === "Animation") { eff = sub; break; }
              }
              if (eff) break;
            }
          }
        } catch (eAnimIn) {}
      }
      if (!eff) return null;
      var ctrl = null;
      try { ctrl = eff.property(1); } catch (eP) {}
      if (!ctrl || (ctrl.name && String(ctrl.name) !== "Animation")) {
        try {
          var n = eff.numProperties || 0;
          for (var idx = 1; idx <= n; idx++) {
            var p = eff.property(idx);
            if (p && p.name && String(p.name) === "Animation") { ctrl = p; break; }
          }
        } catch (eSearch) {}
      }
      if (!ctrl) return null;

      var a = Math.max(0, Number(layer.inPoint || 0));
      var b = Math.max(a, Number(layer.outPoint || 0));
      var t = (tCapture != null && tCapture >= a && tCapture <= b) ? tCapture : (a + (b - a) * 0.8);

      var hadKeys = false, tempIndex = 0, backupVal = null;
      var result = null;
      try {
        if (ctrl.numKeys > 0) {
          hadKeys = true;
          tempIndex = ctrl.addKey(t);
          ctrl.setValueAtKey(tempIndex, 100);
        } else {
          backupVal = ctrl.value;
          ctrl.setValue(100);
        }

        var r = null;
        try { r = layer.sourceRectAtTime(t, false); } catch (eSR0) { r = null; }
        if (!r || r.width <= 0.1 || r.height <= 0.1) {
          try { r = layer.sourceRectAtTime(t, true); } catch (eSR1) { r = null; }
        }
        if (!r || r.width <= 0.1 || r.height <= 0.1) return null;

        var x1 = r.left, y1 = r.top;
        var x2 = r.left + r.width, y2 = r.top + r.height;
        var p1 = layerPointToCompAtTime(layer, [x1, y1], t);
        var p2 = layerPointToCompAtTime(layer, [x2, y2], t);
        var p3 = layerPointToCompAtTime(layer, [x1, y2], t);
        var p4 = layerPointToCompAtTime(layer, [x2, y2], t);
        if (!p1 || !p2 || !p3 || !p4) return null;

        var minX = Math.min(p1[0], p2[0], p3[0], p4[0]);
        var maxX = Math.max(p1[0], p2[0], p3[0], p4[0]);
        var minY = Math.min(p1[1], p2[1], p3[1], p4[1]);
        var maxY = Math.max(p1[1], p2[1], p3[1], p4[1]);
        var w = maxX - minX, h = maxY - minY;
        if (w < 0.1 || h < 0.1) return null;
        minX -= Math.min(w * 0.08, 12);
        if (minX < 0) minX = 0;
        w = maxX - minX;
        result = { x: minX, y: minY, w: w, h: h };
      } finally {
        try {
          if (hadKeys && tempIndex > 0 && tempIndex <= (ctrl.numKeys || 0)) ctrl.removeKey(tempIndex);
          else if (!hadKeys && backupVal != null) ctrl.setValue(backupVal);
        } catch (eR) {}
      }
      return result;
    } catch (eOuter) {
      return null;
    }
  }

  function bboxForLayer(layer, comp, t){
    // Only for Typewriter Blinking Cursor: use full-text size (completed text at 100%) so highlight isn't just the cursor "|".
    if (layer && layer.matchName === "ADBE Text Layer" && hasBlinkingCursorTypewriterEffect(layer)) {
      var bbFull = bboxEstimateFromTextDocNoCursor(layer, t, comp);
      if (bbFull && bbFull.w >= 2 && bbFull.h >= 2) return { bbox: bbFull, source: "blinkFullText" };
    }
    var bb = bboxFromSourceRect(layer, t);
    if (bb) return { bbox: bb, source: 'sourceRect' };

    bb = bboxEstimateFromTextDoc(layer, t);
    if (bb) return { bbox: bb, source: 'estimate' };

    bb = bboxFromPositionRescue(layer, comp, t);
    if (bb) return { bbox: bb, source: 'rescue' };

    return null;
  }

  function bboxIntersectsComp(comp, bb){
    if (!comp || !bb) return false;
    if (bb.w < 2 || bb.h < 2) return false;

    var x1 = bb.x, y1 = bb.y, x2 = bb.x + bb.w, y2 = bb.y + bb.h;
    var ix1 = Math.max(0, x1);
    var iy1 = Math.max(0, y1);
    var ix2 = Math.min(comp.width,  x2);
    var iy2 = Math.min(comp.height, y2);

    return (ix2 - ix1) >= 2 && (iy2 - iy1) >= 2;
  }

  /** True if bb overlaps refRect by at least 2px. Use when refRect is effective comp (full comp or matte). */
  function bboxIntersectsRect(bb, refRect) {
    if (!bb || !refRect || bb.w < 2 || bb.h < 2) return false;
    var ix1 = Math.max(refRect.x, bb.x);
    var iy1 = Math.max(refRect.y, bb.y);
    var ix2 = Math.min(refRect.x + refRect.w, bb.x + bb.w);
    var iy2 = Math.min(refRect.y + refRect.h, bb.y + bb.h);
    return (ix2 - ix1) >= 2 && (iy2 - iy1) >= 2;
  }

  function intersectionRatio(comp, bb){
    if (!comp || !bb) return 0;
    var x1 = bb.x, y1 = bb.y, x2 = bb.x + bb.w, y2 = bb.y + bb.h;

    var ix1 = Math.max(0, x1);
    var iy1 = Math.max(0, y1);
    var ix2 = Math.min(comp.width,  x2);
    var iy2 = Math.min(comp.height, y2);

    var iw = Math.max(0, ix2 - ix1);
    var ih = Math.max(0, iy2 - iy1);
    var inter = iw * ih;
    var area = Math.max(1, bb.w * bb.h);

    return inter / area;
  }

  /** Overlap area of two comp-space bboxes. */
  function bboxOverlapArea(bb1, bb2) {
    if (!bb1 || !bb2) return 0;
    var x1 = Math.max(bb1.x, bb2.x);
    var y1 = Math.max(bb1.y, bb2.y);
    var x2 = Math.min(bb1.x + bb1.w, bb2.x + bb2.w);
    var y2 = Math.min(bb1.y + bb1.h, bb2.y + bb2.h);
    if (x2 <= x1 || y2 <= y1) return 0;
    return (x2 - x1) * (y2 - y1);
  }

  /** Ratio of bboxA's area that lies inside bboxB (0–1). Used for "text inside matte" check. */
  function intersectionRatioRects(bboxA, bboxB) {
    if (!bboxA || !bboxB || bboxA.w < 1 || bboxA.h < 1) return 0;
    var iw = Math.max(0, Math.min(bboxA.x + bboxA.w, bboxB.x + bboxB.w) - Math.max(bboxA.x, bboxB.x));
    var ih = Math.max(0, Math.min(bboxA.y + bboxA.h, bboxB.y + bboxB.h) - Math.max(bboxA.y, bboxB.y));
    var inter = iw * ih;
    var area = bboxA.w * bboxA.h;
    return area > 0 ? inter / area : 0;
  }

  /** Matte layer is the layer directly above (index - 1). NO_TRACK_MATTE = 1 in AE. */
  function getMatteLayerAbove(layer, comp) {
    try {
      if (!layer || !comp || layer.index <= 1) return null;
      var tt = layer.trackMatteType;
      if (tt == null || tt === 1) return null;
      if (typeof TrackMatteType !== "undefined" && tt === TrackMatteType.NO_TRACK_MATTE) return null;
      return comp.layer(layer.index - 1);
    } catch (e) { return null; }
  }

  /** True if layer has no track matte, or its bbox at t is sufficiently inside the matte layer's bbox. */
  function textVisibleInsideTrackMatte(layer, comp, t) {
    var matte = getMatteLayerAbove(layer, comp);
    if (!matte) return true;
    try {
      var textInfo = bboxForLayer(layer, comp, t);
      var matteBbox = getMatteBboxAtTime(matte, comp, t);
      if (!textInfo || !textInfo.bbox || !matteBbox) return false;
      return intersectionRatioRects(textInfo.bbox, matteBbox) >= MIN_IN_RATIO;
    } catch (e) { return false; }
  }

  /** Ratio of text bbox inside matte at t (0–1). Returns 1 if no track matte, so scoring is unchanged. */
  function getMatteRatioAt(layer, comp, t) {
    var matte = getMatteLayerAbove(layer, comp);
    if (!matte) return 1;
    try {
      var textInfo = bboxForLayer(layer, comp, t);
      var matteBbox = getMatteBboxAtTime(matte, comp, t);
      if (!textInfo || !textInfo.bbox || !matteBbox) return 0;
      return intersectionRatioRects(textInfo.bbox, matteBbox);
    } catch (e) { return 0; }
  }

  /** Matte layer bbox at t: use mask bounds when the matte has masks (e.g. solid+mask small box), else layer source rect (no tighten). */
  function getMatteBboxAtTime(matte, comp, t) {
    try {
      if (matte.mask && matte.mask.numProperties && matte.mask.numProperties > 0) {
        var maskBounds = getMaskShapeBoundsAtTime(matte, t);
        if (maskBounds) return maskBounds;
      }
      var bb = bboxFromSourceRectNoTighten(matte, t);
      if (bb) return bb;
      var info = bboxForLayer(matte, comp, t);
      return (info && info.bbox) ? info.bbox : null;
    } catch (e) { return null; }
  }

  /** Source rect in comp space without BBOX_TIGHTEN (for matte layer so we get full shape bounds). */
  function bboxFromSourceRectNoTighten(layer, t) {
    try {
      var r = null;
      try { r = layer.sourceRectAtTime(t, false); } catch (e0) { r = null; }
      if (!r || r.width <= 0.1 || r.height <= 0.1) { try { r = layer.sourceRectAtTime(t, true); } catch (e1) { r = null; } }
      if (!r || r.width <= 0.1 || r.height <= 0.1) return null;
      var x1 = r.left, y1 = r.top, x2 = r.left + r.width, y2 = r.top + r.height;
      var offset = getEffectPositionOffsetAtTime(layer, t);
      if (offset && offset.length >= 2) {
        x1 += Number(offset[0]); y1 += Number(offset[1]);
        x2 += Number(offset[0]); y2 += Number(offset[1]);
      }
      var p1 = layerPointToCompAtTime(layer, [x1, y1], t);
      var p2 = layerPointToCompAtTime(layer, [x2, y1], t);
      var p3 = layerPointToCompAtTime(layer, [x1, y2], t);
      var p4 = layerPointToCompAtTime(layer, [x2, y2], t);
      if (!p1 || !p2 || !p3 || !p4) return null;
      var minX = Math.min(p1[0], p2[0], p3[0], p4[0]);
      var maxX = Math.max(p1[0], p2[0], p3[0], p4[0]);
      var minY = Math.min(p1[1], p2[1], p3[1], p4[1]);
      var maxY = Math.max(p1[1], p2[1], p3[1], p4[1]);
      var w = maxX - minX, h = maxY - minY;
      if (w < 0.1 || h < 0.1) return null;
      return { x: minX, y: minY, w: w, h: h };
    } catch (e) { return null; }
  }

  /** Bbox of the first mask's shape at t in comp space. Tries layer.mask(1).maskPath then ADBE Mask Parade. */
  function getMaskShapeBoundsAtTime(layer, t) {
    try {
      var shapeProp = null;
      if (layer.mask && layer.mask.numProperties >= 1) {
        var firstMask = layer.mask(1);
        if (firstMask) {
          if (firstMask.maskPath) shapeProp = firstMask.maskPath;
          else if (firstMask.property("ADBE Mask Shape")) shapeProp = firstMask.property("ADBE Mask Shape");
        }
      }
      if (!shapeProp && layer.property("ADBE Mask Parade")) {
        var maskGroup = layer.property("ADBE Mask Parade");
        if (maskGroup.numProperties >= 1) {
          var maskAtom = maskGroup.property(1);
          if (maskAtom) shapeProp = maskAtom.property("ADBE Mask Shape");
        }
      }
      if (!shapeProp) return null;
      var shape = shapeProp.valueAtTime(t, false);
      if (!shape) return null;
      var verts = shape.vertices;
      var n = (verts && verts.length != null && typeof verts.length === "number") ? verts.length : 0;
      if (n < 2) return null;
      var minX = 1e9, maxX = -1e9, minY = 1e9, maxY = -1e9;
      for (var i = 0; i < n; i++) {
        var v = verts[i];
        if (!v) continue;
        var x = (v[0] !== undefined) ? Number(v[0]) : (v.x != null ? Number(v.x) : NaN);
        var y = (v[1] !== undefined) ? Number(v[1]) : (v.y != null ? Number(v.y) : NaN);
        if (!isFinite(x) || !isFinite(y)) continue;
        var pt = layerPointToCompAtTime(layer, [x, y], t);
        if (pt && pt.length >= 2) {
          if (pt[0] < minX) minX = pt[0]; if (pt[0] > maxX) maxX = pt[0];
          if (pt[1] < minY) minY = pt[1]; if (pt[1] > maxY) maxY = pt[1];
        }
      }
      if (minX > maxX || minY > maxY) return null;
      return { x: minX, y: minY, w: Math.max(1, maxX - minX), h: Math.max(1, maxY - minY) };
    } catch (e) { return null; }
  }

  /**
   * For track-matte layers: the visible "comp" is the matte. Returns the bbox to use as the composition bounds for visibility/scoring.
   * - No track matte: comp bounds { x:0, y:0, w: comp.width, h: comp.height }.
   * - Track matte: matte layer's bbox at t (the small box). Scan then treats the masked layer as if the comp size were the matte size.
   */
  function getEffectiveBoundsForLayerAtTime(layer, comp, t) {
    try {
      var matte = getMatteLayerAbove(layer, comp);
      if (matte) {
        var matteBbox = getMatteBboxAtTime(matte, comp, t);
        if (matteBbox && matteBbox.w >= 2 && matteBbox.h >= 2) return matteBbox;
      }
      return { x: 0, y: 0, w: comp.width || 1920, h: comp.height || 1080 };
    } catch (e) { return { x: 0, y: 0, w: comp.width || 1920, h: comp.height || 1080 }; }
  }

  /** True if at time t any layer above (lower index) largely covers the text bbox. */
  function isCoveredByLayerDirectlyAbove(textLayer, comp, t, textBbox) {
    try {
      var textIdx = textLayer.index;
      if (textIdx <= 1) return false;
      var textArea = Math.max(1, textBbox.w * textBbox.h);
      for (var i = 1; i < textIdx; i++) {
        var L = comp.layer(i);
        if (!L || !L.enabled) continue;
        if (t < L.inPoint || t > L.outPoint) continue;
        if (layerOpacityAt(L, t) < 25) continue;
        var aboveBb = bboxFromSourceRect(L, t);
        if (!aboveBb || aboveBb.w < 2 || aboveBb.h < 2) {
          aboveBb = { x: 0, y: 0, w: comp.width || 1920, h: comp.height || 1080 };
        }
        var overlap = bboxOverlapArea(textBbox, aboveBb);
        if (overlap >= 0.5 * textArea) return true;
      }
      return false;
    } catch (e) { return false; }
  }

  function textLenAt(layer, t){
    try{
      var sp = getSourceTextProp(layer);
      if (!sp) return 0;
      var doc = sp.valueAtTime(t, false);
      return String(doc.text || "").length;
    }catch(e){ return 0; }
  }

  /** Width of visible text at time t (sourceRectAtTime). Used to detect full reveal for simple/path-trim typewriter. */
  function sourceRectWidthAt(layer, t) {
    try {
      var r = null;
      try { r = layer.sourceRectAtTime(t, false); } catch (e0) { r = null; }
      if (!r || r.width <= 0) try { r = layer.sourceRectAtTime(t, true); } catch (e1) { r = null; }
      return (r && r.width > 0) ? r.width : 0;
    } catch (e) { return 0; }
  }

  /** For simple typewriter (path trim): text length is constant but visible portion grows. Returns time when rect width is max, or null. */
  function getTimeOfFullVisualReveal(layer, comp) {
    try {
      var a = Math.max(0, Number(layer.inPoint || 0));
      var b = Math.max(a, Number(layer.outPoint || 0));
      if (b - a < 0.01) return null;
      var step = comp.frameDuration || (1 / 24);
      if (!isFinite(step) || step <= 0) step = 1 / 24;
      var bestT = a;
      var maxW = 0;
      for (var t = a; t <= b; t += step) {
        if (!layerEligible(layer, t)) continue;
        var w = sourceRectWidthAt(layer, t);
        if (w > maxW) { maxW = w; bestT = t; }
      }
      return (maxW > 0) ? { t: bestT, maxW: maxW } : null;
    } catch (e) { return null; }
  }

  /** For constant-length typewriter: prefer time in last portion of layer (preset often reveals by end). Returns time >= start of "tail". */
  var TYPEWRITER_TAIL_RATIO = 0.70;
  /** For animator+range-selector typewriter (rect doesn't grow): use later tail so we're in the fully-revealed portion. */
  var TYPEWRITER_TAIL_RATIO_ANIMATOR = 0.90;
  function getMinTimeInTypewriterTail(layer, forAnimatorTypewriter) {
    try {
      var a = Math.max(0, Number(layer.inPoint || 0));
      var b = Math.max(a, Number(layer.outPoint || 0));
      if (b - a < 0.05) return a;
      var ratio = (forAnimatorTypewriter === true) ? TYPEWRITER_TAIL_RATIO_ANIMATOR : TYPEWRITER_TAIL_RATIO;
      return a + (b - a) * ratio;
    } catch (e) { return 0; }
  }

  /** True if source rect width grows over the layer (real typewriter reveal). Plain text has constant rect. */
  function rectWidthGrowsOverDuration(layer, comp) {
    try {
      var a = Math.max(0, Number(layer.inPoint || 0));
      var b = Math.max(a, Number(layer.outPoint || 0));
      if (b - a < 0.05) return false;
      var step = comp.frameDuration || (1 / 24);
      if (!isFinite(step) || step <= 0) step = 1 / 24;
      var minW = 1e9;
      var maxW = 0;
      for (var t = a; t <= b; t += step) {
        if (!layerEligible(layer, t)) continue;
        var w = sourceRectWidthAt(layer, t);
        if (w > 0) { if (w < minW) minW = w; if (w > maxW) maxW = w; }
      }
      return (minW > 0 && minW < 1e8 && maxW >= minW * 1.15);
    } catch (e) { return false; }
  }

  /** True if text length is constant over the layer (simple/path-trim typewriter). */
  function isConstantLengthTypewriter(layer, comp, maxLen) {
    if (!maxLen || maxLen < 2) return false;
    try {
      var a = Math.max(0, Number(layer.inPoint || 0));
      var b = Math.max(a, Number(layer.outPoint || 0));
      if (b - a < 0.01) return false;
      var samples = [a, a + (b - a) * 0.25, a + (b - a) * 0.5, a + (b - a) * 0.75, b];
      for (var i = 0; i < samples.length; i++) {
        var t = samples[i];
        if (t > b) continue;
        if (textLenAt(layer, t) !== maxLen) return false;
      }
      return true;
    } catch (e) { return false; }
  }

  /** Animator + range selector typewriter: constant length but sourceRect does not grow (AE reports full text bounds). Prefer end of layer and only 1 stable frame. */
  function isAnimatorTypewriter(layer, comp, maxLen) {
    if (!maxLen || maxLen < 2) return false;
    try {
      var a = Math.max(0, Number(layer.inPoint || 0));
      var b = Math.max(a, Number(layer.outPoint || 0));
      if (b - a < 0.5) return false;
      return isConstantLengthTypewriter(layer, comp, maxLen) && !rectWidthGrowsOverDuration(layer, comp);
    } catch (e) { return false; }
  }

  /**
   * First time in [inPoint, outPoint] when ANY text animator RANGE SELECTOR's Start reaches 100%.
   * This follows the AE UI directly (the Range Selector "Start" slider) instead of trying to infer
   * coverage. Works regardless of animator name/preset and for any number of animators/selectors.
   * Returns null if no selector has a Start property or it never reaches ~100 in the scan window.
   */
  function getTimeWhenRangeSelectorStartAt100(layer, comp, maxLen) {
    try {
      if (!layer || layer.matchName !== "ADBE Text Layer") return null;
      var tp = layer.property("ADBE Text Properties");
      if (!tp) return null;
      var animatorsGroup = tp.property("ADBE Text Animators");
      if (!animatorsGroup || animatorsGroup.numProperties < 1) return null;

      var a = Math.max(0, Number(layer.inPoint || 0));
      var b = Math.max(a, Number(layer.outPoint || 0));
      var step = comp.frameDuration || (1 / 24);
      if (!isFinite(step) || step <= 0) step = 1 / 24;
      if (b - a < step) return null;

      function numVal(prop, t) {
        if (!prop) return NaN;
        try {
          var v = prop.valueAtTime(t, false);
          if (v != null && typeof v === "number" && isFinite(v)) return v;
          if (v && typeof v.length === "number" && v.length > 0) return Number(v[0]);
          v = prop.valueAtTime(t, true);
          if (v != null && typeof v === "number" && isFinite(v)) return v;
          if (v && typeof v.length === "number" && v.length > 0) return Number(v[0]);
        } catch (e) {}
        return NaN;
      }

      // Earliest moment when Start is ~100. Prefer the SECOND keyframe of the selector's Start
      // (common AE pattern: key1=0, key2=100) and fall back to other keyframes / sampling with
      // a short "hold" period if that is not usable. Keyframes may be on the main property or on a dimension (X/Y).
      function propWithKeys(prop) {
        if (!prop) return null;
        if (prop.numKeys != null && prop.numKeys >= 2) return prop;
        if (prop.numProperties != null) {
          for (var d = 1; d <= Math.min(prop.numProperties, 5); d++) {
            try {
              var sub = prop.property(d);
              if (sub && sub.numKeys != null && sub.numKeys >= 2) return sub;
            } catch (ed) {}
          }
        }
        return prop;
      }
      function stableStartAt100Time(startProp, a, b, baseStep) {
        if (!startProp) return null;
        var keyProp = propWithKeys(startProp);
        var maxSamples = 80;
        var span = b - a;
        if (span <= 0) return null;
        var stepLocal = baseStep;
        if (!isFinite(stepLocal) || stepLocal <= 0) stepLocal = 1 / 24;
        if (span / stepLocal > maxSamples) stepLocal = span / maxSamples;

        var threshold = 99;
        var holdThreshold = 90;
        var minHold = Math.max(stepLocal * 2, 0.05); // tiny pause at ~100%

        // 1) Hard preference: if the SECOND keyframe of Start is ~100 within [a,b], use it.
        try {
          if (keyProp.numKeys && keyProp.numKeys >= 2) {
            var kt2 = keyProp.keyTime(2);
            if (kt2 >= a - 0.001 && kt2 <= b + 0.001) {
              var kv2;
              try { kv2 = keyProp.keyValue(2); } catch (eKV2) { kv2 = numVal(startProp, kt2); }
              if (isFinite(kv2) && kv2 >= threshold) return kt2;
            }
          }
        } catch (eSecond) {}

        // 2) General keyframe-first path: any key in [a,b] that is ~100 and stays high for a bit.
        try {
          if (keyProp.numKeys && keyProp.numKeys >= 2) {
            for (var ki = 1; ki <= keyProp.numKeys; ki++) {
              var kt = keyProp.keyTime(ki);
              if (kt < a - 0.001 || kt > b + 0.001) continue;
              var kv;
              try { kv = keyProp.keyValue(ki); } catch (eKV) { kv = numVal(startProp, kt); }
              if (!isFinite(kv) || kv < threshold) continue;

              // If there is a following key, ensure we stay high at least halfway to it.
              var tEndHold = null;
              if (ki < keyProp.numKeys) {
                var ktNext = keyProp.keyTime(ki + 1);
                tEndHold = Math.min(b, kt + Math.max(minHold, (ktNext - kt) * 0.5));
              } else {
                tEndHold = Math.min(b, kt + minHold);
              }

              var okKF = true;
              for (var tfKF = kt; tfKF <= tEndHold + 0.0001; tfKF += stepLocal) {
                var vfKF = numVal(startProp, tfKF);
                if (!isFinite(vfKF) || vfKF < holdThreshold) { okKF = false; break; }
              }
              if (okKF) return kt;
            }
          }
        } catch (eKF) {}

        // 3) Fallback sampling path if keyframes are not available or inconclusive.
        var t;
        for (t = a; t <= b + 0.0001; t += stepLocal) {
          var v = numVal(startProp, t);
          if (!isFinite(v) || v < threshold) continue;

          var tEnd = Math.min(b, t + minHold);
          var ok = true;
          var tf;
          for (tf = t; tf <= tEnd + 0.0001; tf += stepLocal) {
            var vf = numVal(startProp, tf);
            if (!isFinite(vf) || vf < holdThreshold) { ok = false; break; }
          }
          if (ok) return t;
        }
        return null;
      }

      var bestT = null;
      var ai, animator, selectorsGroup, si, selector, startProp;
      for (ai = 1; ai <= animatorsGroup.numProperties; ai++) {
        animator = animatorsGroup.property(ai);
        if (!animator) continue;
        selectorsGroup = animator.property("ADBE Text Selectors");
        if (!selectorsGroup || selectorsGroup.numProperties < 1) continue;

        for (si = 1; si <= selectorsGroup.numProperties; si++) {
          selector = selectorsGroup.property(si);
          if (!selector) continue;
          startProp = null;
          try { startProp = selector.property("ADBE Text Percent Start"); } catch (eSP) {}
          if (!startProp) continue;
          var tSel = stableStartAt100Time(startProp, a, b, step);
          if (tSel != null && (bestT == null || tSel < bestT)) bestT = tSel;
        }
      }
      return bestT;
    } catch (eOuter) {
      return null;
    }
  }

  /** First time when a "Typewriter" (or similar) layer EFFECT reaches full reveal. Effects live under ADBE Effect Parade; can be in groups (e.g. "Animate In" > "Typewriter"). Scans effect params for 0-100 or 0-1 sliders. */
  function getTimeWhenEffectTypewriterAt100(layer, comp) {
    try {
      if (!layer) return null;
      var step = comp.frameDuration || (1 / 24);
      if (!isFinite(step) || step <= 0) step = 1 / 24;
      var a = Math.max(0, Number(layer.inPoint || 0));
      var b = Math.max(a, Number(layer.outPoint || 0));
      if (b - a < step) return null;
      var debugLog = (typeof DEBUG_TYPEWRITER_LOG !== "undefined" && DEBUG_TYPEWRITER_LOG) ? [] : null;
      var compName = (comp && comp.name != null) ? String(comp.name) : "";
      if (debugLog) debugLog.push("--- Layer: '" + (layer.name || "") + "' (comp: " + compName + ") ---");

      function effectNumVal(prop, t) {
        if (!prop) return NaN;
        try {
          var v = prop.valueAtTime(t, false);
          if (v != null && typeof v === "number" && isFinite(v)) return v;
          if (v && typeof v.length === "number" && v.length > 0) return Number(v[0]);
          v = prop.valueAtTime(t, true);
          if (v != null && typeof v === "number" && isFinite(v)) return v;
          if (v && typeof v.length === "number" && v.length > 0) return Number(v[0]);
          try { v = prop.value; if (v != null && typeof v === "number" && isFinite(v)) return v; if (v && v.length && v.length > 0) return Number(v[0]); } catch (e) {}
          return NaN;
        } catch (e) { return NaN; }
      }
      function firstTimeParamAtOrAbove(prop, a, b, step, threshold) {
        if (!prop) return null;
        function fromPropKeyframes(keyProp, valueProp) {
          if (!keyProp) return null;
          var p = valueProp || keyProp;
          try {
            if (keyProp.numKeys && keyProp.numKeys > 0) {
              for (var ki = 1; ki <= keyProp.numKeys; ki++) {
                var kt = keyProp.keyTime(ki);
                if (kt >= a - 0.001 && kt <= b + 0.001) {
                  var v = effectNumVal(p, kt);
                  if (isFinite(v) && (v >= threshold || (threshold >= 98 && v >= 0.98 && v <= 1.02))) return kt;
                }
              }
            }
          } catch (e) {}
          return null;
        }
        var tFromMain = fromPropKeyframes(prop, prop);
        if (tFromMain != null) {
          if (prop.numProperties != null && prop.numProperties >= 1) {
            for (var d = 1; d <= Math.min(prop.numProperties, 5); d++) {
              try {
                var sub = prop.property(d);
                var tSub = fromPropKeyframes(sub, prop);
                if (tSub != null && tSub < tFromMain) tFromMain = tSub;
              } catch (ed) {}
            }
          }
          return tFromMain;
        }
        if (prop.numProperties != null && prop.numProperties >= 1) {
          for (var d = 1; d <= Math.min(prop.numProperties, 5); d++) {
            try {
              var sub = prop.property(d);
              var tSub = fromPropKeyframes(sub, prop);
              if (tSub != null) return tSub;
            } catch (ed) {}
          }
        }
        for (var t = a; t <= b; t += step) {
          var v = effectNumVal(prop, t);
          if (isFinite(v) && (v >= threshold || (threshold >= 98 && v >= 0.98 && v <= 1.02))) return t;
        }
        return null;
      }
      function scanEffectParams(effectObj, effectName) {
        if (!effectObj) return;
        var n = 0;
        try { n = effectObj.numProperties; } catch (e) {}
        if (n == null || n < 1) return;
        try {
          for (var j = 1; j <= n; j++) {
            var param = null;
            try { param = effectObj.property(j); } catch (e1) {}
            if (!param) try { param = effectObj.param(j); } catch (e2) {}
            if (!param) continue;
            var nsub = 0;
            try { nsub = param.numProperties; } catch (e) {}
            if (nsub != null && nsub > 0) {
              scanEffectParams(param, (param.name || "").toString());
              continue;
            }
            var pname = (param.name || "").toString();
            // Blinking Cursor "Animation" effect: only the main Slider controls reveal; ignore Effect Opacity / GPU Rendering so we capture when text is fully revealed (e.g. Slider=9.961 not Opacity=8.880).
            if (effectName === "Animation" && pname !== "Slider" && pname !== "Animation") continue;
            var t98 = firstTimeParamAtOrAbove(param, a, b, step, 98);
            if (t98 != null) {
              if (bestT == null || t98 < bestT) bestT = t98;
              if (debugLog) debugLog.push("  param[" + j + "] '" + pname + "' => t=" + t98.toFixed(3));
            }
            var t098 = firstTimeParamAtOrAbove(param, a, b, step, 0.98);
            if (t098 != null && (bestT == null || t098 < bestT)) bestT = t098;
          }
        } catch (e) { if (debugLog) debugLog.push("  scanEffectParams err: " + e.toString()); }
      }

      var bestT = null;

      try {
        var effDirect = layer.effect("Typewriter");
        if (effDirect && effDirect.numProperties != null) {
          if (debugLog) debugLog.push("layer '" + (layer.name || "") + "': found effect('Typewriter')");
          scanEffectParams(effDirect, "Typewriter");
        }
      } catch (e) {}

      try {
        var animIn = layer.effect("Animate In");
        if (animIn && animIn.numProperties != null) {
          if (debugLog) debugLog.push("layer '" + (layer.name || "") + "': found effect('Animate In'), numProperties=" + animIn.numProperties);
          scanEffectParams(animIn, "Animate In");
          for (var k = 1; k <= animIn.numProperties; k++) {
            var sub = animIn.property(k);
            if (!sub) continue;
            var subName = (sub.name || "").toString();
            if (subName.indexOf("Typewriter") >= 0 && sub.numProperties != null) {
              if (debugLog) debugLog.push("  sub '" + subName + "' (Typewriter-like)");
              scanEffectParams(sub, subName);
            }
          }
        }
      } catch (e) {}

      try {
        var animEffect = layer.effect("Animation");
        if (animEffect && animEffect.numProperties != null && animEffect.numProperties > 0) {
          if (debugLog) debugLog.push("layer '" + (layer.name || "") + "': found effect('Animation')");
          scanEffectParams(animEffect, "Animation");
        }
      } catch (e) {}

      if (bestT == null && debugLog) {
        for (var ei = 1; ei <= 30; ei++) {
          try {
            var effByIndex = layer.effect(ei);
            if (!effByIndex) continue;
            var ename = (effByIndex.name || "").toString();
            var nprop = (effByIndex.numProperties != null) ? effByIndex.numProperties : 0;
            debugLog.push("layer.effect(" + ei + ") = '" + ename + "' numProperties=" + nprop);
            var isTypewriterLike = (ename.indexOf("Typewriter") >= 0 || ename.indexOf("Animate In") >= 0 || ename.indexOf("Animate") >= 0 || ename === "Animation");
            if (isTypewriterLike && nprop > 0) {
              scanEffectParams(effByIndex, ename);
            }
            if (nprop > 0 && ename.indexOf("Typewriter") < 0 && (ename.indexOf("Animate") >= 0 || ename === "Animation")) {
              for (var sk = 1; sk <= nprop; sk++) {
                try {
                  var subEff = effByIndex.property(sk);
                  if (!subEff) continue;
                  var subName = (subEff.name || "").toString();
                  if (subName.indexOf("Typewriter") >= 0 && subEff.numProperties != null) {
                    debugLog.push("  -> sub '" + subName + "'");
                    scanEffectParams(subEff, subName);
                  }
                } catch (e3) {}
              }
            }
          } catch (e2) {}
        }
      }

      var effectsGroup = layer.property("ADBE Effect Parade");
      if (effectsGroup && effectsGroup.numProperties) {
        function scanEffectOrGroup(grp) {
          if (!grp || !grp.numProperties) return;
          try {
            for (var i = 1; i <= grp.numProperties; i++) {
              var p = grp.property(i);
              if (!p) continue;
              var name = (p.name || "").toString();
              var matchName = (p.matchName || "").toString();
              var isTypewriterLike = (name.indexOf("Typewriter") >= 0 || name.indexOf("Type writer") >= 0 || matchName.indexOf("Typewriter") >= 0);
              var isAnimateIn = (name.indexOf("Animate In") >= 0 || name.indexOf("Animate") >= 0);
              if (p.numProperties != null && p.numProperties > 0) {
                if (isTypewriterLike) {
                  if (debugLog) debugLog.push("layer '" + (layer.name || "") + "': parade group '" + name + "' (Typewriter-like)");
                  scanEffectParams(p, name);
                } else if (isAnimateIn || name.length > 0) {
                  scanEffectOrGroup(p);
                } else {
                  scanEffectOrGroup(p);
                }
              }
            }
          } catch (e) {}
        }
        scanEffectOrGroup(effectsGroup);
      } else if (debugLog) {
        var ep = layer.property("ADBE Effect Parade");
        debugLog.push("layer '" + (layer.name || "") + "': no ADBE Effect Parade or empty (effectsGroup=" + (ep ? "exists" : "null") + (ep && ep.numProperties != null ? ", numProperties=" + ep.numProperties : "") + ")");
      }

      if (debugLog && debugLog.length > 0) {
        debugLog.push("=> getTimeWhenEffectTypewriterAt100 result: " + (bestT != null ? bestT.toFixed(3) : "null"));
        try {
          var f = new File(Folder.myDocuments.fsName + "/Crowdin_typewriter_debug.txt");
          f.open("a");
          f.write(debugLog.join("\n") + "\n");
          f.close();
        } catch (ex) {}
      }
      return bestT;
    } catch (e) { return null; }
  }

  /** First time when typewriter is at full reveal: from text animator range selector OR from Typewriter (effect). */
  function getTimeWhenTypewriterFullReveal(layer, comp, maxLen) {
    var tAnim = getTimeWhenRangeSelectorStartAt100(layer, comp, maxLen);
    var tEffect = getTimeWhenEffectTypewriterAt100(layer, comp);
    // Blinking Cursor: prefer effect Slider time (full reveal) so we never modify the effect and sourceRect at t is correct.
    if (hasBlinkingCursorTypewriterEffect(layer) && tEffect != null) return tEffect;
    if (tEffect != null && (tAnim == null || tEffect < tAnim)) return tEffect;
    return tAnim;
  }

  /** Return the completed/full text for Crowdin (e.g. after typewriter or without blinking cursor). Samples layer duration and returns text at a time when length is max. */
  function getCompletedTextForLayer(layer, comp) {
    try {
      var sp = getSourceTextProp(layer);
      if (!sp) return "";
      var a = Math.max(0, Number(layer.inPoint || 0));
      var b = Math.max(a, Number(layer.outPoint || 0));
      if (b - a < 0.01) return trim(String(sp.value.text || ""));
      var step = comp.frameDuration || (1/24);
      if (!isFinite(step) || step <= 0) step = 1/24;
      var maxLen = 0;
      var bestT = a;
      for (var t = a; t <= b; t += step) {
        if (!layerEligible(layer, t)) continue;
        var len = textLenAt(layer, t);
        if (len > maxLen) { maxLen = len; bestT = t; }
      }
      var doc = sp.valueAtTime(bestT, false);
      var txt = trim(String(doc.text || ""));
      txt = stripBlinkingCursorCursor(txt);
      txt = txt.replace(/\s+$/, "").replace(/^\s+/, "");
      return trim(txt);
    } catch (e) { return ""; }
  }

  function exportCompPngAtTime(comp, t, outFile, resolutionAlreadySet){
    var oldRes = null;
    try{
      if (!/\.png$/i.test(outFile.fsName)) outFile = new File(outFile.fsName + ".png");
      try { if (outFile.exists) outFile.remove(); } catch(e0){}

      try { app.project.activeItem = comp; } catch(eActive){}
      try { comp.time = t; } catch(e2){}

      if (resolutionAlreadySet) {
        // Crowdin path: force Half resolution (AE uses divisor: 2 = Half, 4 = Quarter; Full = 1)
        try { comp.resolutionFactor = [2, 2]; } catch(eHalf){}
      } else {
        try { oldRes = comp.resolutionFactor; } catch(eOld){}
        if (SCREENSHOT_RES_FACTOR && SCREENSHOT_RES_FACTOR > 1) {
          try { comp.resolutionFactor = [SCREENSHOT_RES_FACTOR, SCREENSHOT_RES_FACTOR]; } catch(eRF){}
        }
      }

      if (typeof SKIP_REFRESH_EVERY_FRAME === "undefined" || !SKIP_REFRESH_EVERY_FRAME) {
        try { app.refresh(); } catch(e3){}
      }

      // Give AE time to update comp to time t before capture (saveFrameToPng can be async and may use current comp state).
      try { $.sleep(150); } catch (eSleep) {}

      comp.saveFrameToPng(t, outFile);

      var tries = 0;
      while (!outFile.exists && tries < 60) { $.sleep(50); tries++; }
      return outFile.exists;
    }catch(e){
      return false;
    } finally {
      if (!resolutionAlreadySet && oldRes && oldRes.length === 2) {
        try { comp.resolutionFactor = oldRes; } catch(eBack){}
      }
    }
  }

  // ✅ Single rule: screenshot at second keyframe or midpoint. Snapshot Marker (when present and valid) is the top-level preference.
  function findBestTime(layer, comp){
    try{
      var win = getCompScanWindow(comp);
      var winStart = win.start;
      var winEnd = win.end;
      var a = Math.max(winStart, Math.max(0, Number(layer.inPoint || 0)));
      var b = Math.min(winEnd, Math.max(a, Number(layer.outPoint || 0)));
      if ((b - a) < 0.01) {
        var layerIn = Math.max(0, Number(layer.inPoint || 0));
        var layerOut = Math.max(layerIn, Number(layer.outPoint || 0));
        var compDur = 0;
        try { compDur = Math.max(0, Number(comp.duration || 0)); } catch (e) {}
        if (compDur > 0 && layerIn < compDur && layerOut > 0 && (layerOut - layerIn) >= 0.01) {
          a = Math.max(0, layerIn);
          b = Math.min(compDur, layerOut);
        }
      }
      if ((b - a) < 0.01) return null;

      var tSnap = getSnapshotMarkerTime(layer);
      if (tSnap != null && isFinite(tSnap) && tSnap >= a - 0.001 && tSnap <= b + 0.001 && layerEligible(layer, tSnap)) {
        var info = bboxForLayer(layer, comp, tSnap);
        var bb = info ? info.bbox : null;
        if (!bb) bb = bboxFromPositionRescue(layer, comp, tSnap);
        if (bb && bb.w >= 1 && bb.h >= 1)
          return { t: tSnap, bbox: bb, source: (info && info.source) ? info.source : "snapshot-marker" };
      }

      var t = getScreenshotTimeForLayer(layer, a, b);
      if (!layerEligible(layer, t)) return null;
      var ref = getEffectiveBoundsForLayerAtTime(layer, comp, t);
      var info = bboxForLayer(layer, comp, t);
      var bb = info ? info.bbox : null;
      if (!bb) bb = bboxFromPositionRescue(layer, comp, t);
      if (!bb || !bboxIntersectsRect(bb, ref)) return null;
      if ((ref ? intersectionRatioRects(bb, ref) : 1) < MIN_IN_RATIO) return null;
      if (isCoveredByLayerDirectlyAbove(layer, comp, t, bb)) return null;
      return { t: t, bbox: bb, source: (info && info.source) || "screenshot-time" };
    }catch(e){}
    return null;
  }

  // Last-resort: same single rule (second keyframe or midpoint), then ensure we return a valid bbox.
  function getFallbackTimeAndBbox(layer, comp) {
    try {
      var win = getCompScanWindow(comp);
      var winStart = win.start;
      var winEnd = win.end;
      var a = Math.max(winStart, Math.max(0, Number(layer.inPoint || 0)));
      var b = Math.min(winEnd, Math.max(a, Number(layer.outPoint || 0)));
      if (b - a < 0.01) {
        a = winStart;
        b = winEnd;
      }
      var step = comp.frameDuration || (1/24);
      if (!isFinite(step) || step <= 0) step = 1/24;

      var t = getScreenshotTimeForLayer(layer, a, b);
      for (var j = 0; j <= 10; j++) {
        var tTry = (j === 0) ? t : (t + (j % 2 === 1 ? 1 : -1) * Math.ceil(j / 2) * step);
        if (tTry < a) tTry = a;
        if (tTry > b) tTry = b;
        if (!layerEligible(layer, tTry)) continue;
        var refTry = getEffectiveBoundsForLayerAtTime(layer, comp, tTry);
        var info = bboxForLayer(layer, comp, tTry);
        if (info && info.bbox && bboxIntersectsRect(info.bbox, refTry) && intersectionRatioRects(info.bbox, refTry) >= FALLBACK_RATIO)
          return { t: tTry, bbox: info.bbox, source: info.source };
        var bbR = bboxFromPositionRescue(layer, comp, tTry);
        if (bbR && bboxIntersectsRect(bbR, refTry) && intersectionRatioRects(bbR, refTry) >= FALLBACK_RATIO)
          return { t: tTry, bbox: bbR, source: 'rescue' };
      }
      for (var k = 0; k <= 20; k++) {
        var tk = a + (b - a) * (k / 20);
        if (!layerEligible(layer, tk)) continue;
        var refK = getEffectiveBoundsForLayerAtTime(layer, comp, tk);
        var bb2 = bboxFromPositionRescue(layer, comp, tk);
        if (bb2 && bboxIntersectsRect(bb2, refK) && intersectionRatioRects(bb2, refK) >= FALLBACK_RATIO) return { t: tk, bbox: bb2, source: 'rescue' };
      }
      if (layerEligible(layer, a)) {
        var refA = getEffectiveBoundsForLayerAtTime(layer, comp, a);
        var bb3 = bboxFromPositionRescue(layer, comp, a);
        if (bb3 && bboxIntersectsRect(bb3, refA) && intersectionRatioRects(bb3, refA) >= FALLBACK_RATIO) return { t: a, bbox: bb3, source: 'rescue' };
        var cx = (refA.w * 0.5) + refA.x - 40;
        var cy = (refA.h * 0.5) + refA.y - 20;
        var cxClamp = Math.max(refA.x, Math.min(refA.x + refA.w - 80, cx));
        var cyClamp = Math.max(refA.y, Math.min(refA.y + refA.h - 40, cy));
        return { t: a, bbox: { x: cxClamp, y: cyClamp, w: 80, h: 40 }, source: 'rescue' };
      }
    } catch (e) {}
    return null;
  }

  /** When findBestTime and getFallbackTimeAndBbox both fail (e.g. simple typewriter at keyframe time has tiny bbox), use a time in the tail (90%+) so text is fully revealed and we still send a candidate for this layer. */
  function getRescueTailTimeAndBbox(layer, comp) {
    try {
      var win = getCompScanWindow(comp);
      var a = Math.max(win.start, Math.max(0, Number(layer.inPoint || 0)));
      var b = Math.min(win.end, Math.max(a, Number(layer.outPoint || 0)));
      if (b - a < 0.01) return null;
      var ratios = [0.9, 0.95, 1.0];
      for (var ri = 0; ri < ratios.length; ri++) {
        var t = a + (b - a) * ratios[ri];
        if (t > b) t = b;
        if (!layerEligible(layer, t)) continue;
        var ref = getEffectiveBoundsForLayerAtTime(layer, comp, t);
        var info = bboxForLayer(layer, comp, t);
        var bb = info ? info.bbox : null;
        if (!bb) bb = bboxFromPositionRescue(layer, comp, t);
        if (!bb || bb.w < 1 || bb.h < 1) continue;
        if (ref && !bboxIntersectsRect(bb, ref)) continue;
        return { t: t, bbox: bb, source: (info && info.source) || "rescue-tail" };
      }
    } catch (e) {}
    return null;
  }

  /** Find chain of layers from mainComp down to targetComp. Returns [{ layer, comp }, ...] where path[0].comp = mainComp, path[i].layer.source = path[i+1].comp, path[path.length-1].layer.source = targetComp. */
  function getPathToComp(mainComp, targetComp) {
    if (!mainComp || !targetComp || mainComp === targetComp) return [];
    var targetId = targetComp.id != null ? String(targetComp.id) : null;
    try {
      var layers = mainComp.layers;
      for (var i = 1; i <= layers.length; i++) {
        var L = layers[i];
        if (!L || !L.source) continue;
        var src = L.source;
        var match = (src === targetComp) || (targetId && src.id != null && String(src.id) === targetId);
        if (match) return [{ layer: L, comp: mainComp }];
        if (src instanceof CompItem) {
          var sub = getPathToComp(src, targetComp);
          if (sub.length) return [{ layer: L, comp: mainComp }].concat(sub);
        }
      }
    } catch (e) {}
    return [];
  }

  /** True if this precomp text layer is visible in the main comp at some frame. Respects temporal cropping: only considers main comp times that fall within the precomp layer's in/out and the text layer's in/out. */
  function hasVisibleFrameInMainComp(layer, layerComp, mainComp) {
    var path = getPathToComp(mainComp, layerComp);
    if (!path.length) return false;
    try {
      var step = mainComp.frameDuration || (1/24);
      if (!isFinite(step) || step <= 0) step = 1/24;
      function getPrecompTime(tMain) {
        var t = tMain;
        for (var p = 0; p < path.length; p++) t = t - (Number(path[p].layer.startTime) || 0);
        return t;
      }
      function precompLayerVisibleAt(tMain) {
        var t = tMain;
        for (var p = 0; p < path.length; p++) {
          var pl = path[p].layer;
          try { if (!pl.activeAtTime(t)) return false; } catch (e) { return false; }
          if (layerOpacityAt(pl, t) < MIN_OPACITY) return false;
          if (layerScaleAt(pl, t) < MIN_SCALE) return false;
          t = t - (Number(pl.startTime) || 0);
        }
        return true;
      }
      var firstLayer = path[0].layer;
      var mainWin = getCompScanWindow(mainComp);
      var aMain = Math.max(mainWin.start, Math.max(0, Number(firstLayer.inPoint || 0)));
      var bMain = Math.min(mainWin.end, Math.min(mainComp.duration || 1, Number(firstLayer.outPoint || 1)));
      if (bMain - aMain < 0.01) return false;
      var layerIn = Number(layer.inPoint || 0);
      var layerOut = Number(layer.outPoint || 0);

      // When layer has a track matte in layerComp, require paused + visible in matte so we don't exclude layers that only have valid frames when paused
      var matteInPrecomp = getMatteLayerAbove(layer, layerComp);
      if (matteInPrecomp) {
        var stepMainMatte = Math.min(step, (bMain - aMain) / 80);
        if (!isFinite(stepMainMatte) || stepMainMatte <= 0) stepMainMatte = step;
        for (var tMain = aMain; tMain <= bMain; tMain += stepMainMatte) {
          if (!precompLayerVisibleAt(tMain)) continue;
          var tPre = getPrecompTime(tMain);
          try { if (!layer.activeAtTime(tPre)) continue; } catch (e) { continue; }
          if (tPre < layerIn || tPre > layerOut) continue;
          if (!layerEligible(layer, tPre)) continue;
          if (!isLayerPausedAt(layer, layerComp, tPre)) continue;
          var refPre = getEffectiveBoundsForLayerAtTime(layer, layerComp, tPre);
          var info = bboxForLayer(layer, layerComp, tPre);
          if (!info || !info.bbox || info.bbox.w < 2 || info.bbox.h < 2) continue;
          if (!bboxIntersectsRect(info.bbox, refPre) || intersectionRatioRects(info.bbox, refPre) < MIN_IN_RATIO) continue;
          var times = [];
          var tt = tMain;
          for (var q = 0; q < path.length; q++) {
            times[q] = tt;
            tt = tt - (Number(path[q].layer.startTime) || 0);
          }
          var bb = info.bbox;
          var corners = [[bb.x, bb.y], [bb.x + bb.w, bb.y], [bb.x + bb.w, bb.y + bb.h], [bb.x, bb.y + bb.h]];
          var minX = 1e9, maxX = -1e9, minY = 1e9, maxY = -1e9;
          for (var c = 0; c < corners.length; c++) {
            var pt = corners[c];
            for (var p = path.length - 1; p >= 0; p--) {
              var tr = layerPointToCompAtTime(path[p].layer, [pt[0], pt[1]], times[p]);
              if (!tr || tr.length < 2) break;
              pt = [tr[0], tr[1]];
            }
            if (pt[0] < minX) minX = pt[0]; if (pt[0] > maxX) maxX = pt[0];
            if (pt[1] < minY) minY = pt[1]; if (pt[1] > maxY) maxY = pt[1];
          }
          var bboxMain = { x: minX, y: minY, w: maxX - minX, h: maxY - minY };
          if (bboxIntersectsComp(mainComp, bboxMain) && intersectionRatio(mainComp, bboxMain) >= MIN_IN_RATIO) return true;
        }
      }

      var samples = Math.min(20, Math.max(1, Math.floor((bMain - aMain) / step)));
      for (var i = 0; i <= samples; i++) {
        var tMain = aMain + (bMain - aMain) * (i / Math.max(1, samples));
        if (!precompLayerVisibleAt(tMain)) continue;
        var tPre = getPrecompTime(tMain);
        try { if (!layer.activeAtTime(tPre)) continue; } catch (e) { continue; }
        if (tPre < layerIn || tPre > layerOut) continue;
        if (!layerEligible(layer, tPre)) continue;
        var info = bboxForLayer(layer, layerComp, tPre);
        if (!info || !info.bbox || info.bbox.w < 2 || info.bbox.h < 2) continue;
        var refPre = getEffectiveBoundsForLayerAtTime(layer, layerComp, tPre);
        if (!bboxIntersectsRect(info.bbox, refPre) || intersectionRatioRects(info.bbox, refPre) < MIN_IN_RATIO) continue;
        var times = [];
        var tt = tMain;
        for (var q = 0; q < path.length; q++) {
          times[q] = tt;
          tt = tt - (Number(path[q].layer.startTime) || 0);
        }
        var bb = info.bbox;
        var corners = [[bb.x, bb.y], [bb.x + bb.w, bb.y], [bb.x + bb.w, bb.y + bb.h], [bb.x, bb.y + bb.h]];
        var minX = 1e9, maxX = -1e9, minY = 1e9, maxY = -1e9;
        for (var c = 0; c < corners.length; c++) {
          var pt = corners[c];
          for (var p = path.length - 1; p >= 0; p--) {
            var tr = layerPointToCompAtTime(path[p].layer, [pt[0], pt[1]], times[p]);
            if (!tr || tr.length < 2) break;
            pt = [tr[0], tr[1]];
          }
          if (pt[0] < minX) minX = pt[0]; if (pt[0] > maxX) maxX = pt[0];
          if (pt[1] < minY) minY = pt[1]; if (pt[1] > maxY) maxY = pt[1];
        }
        var bboxMain = { x: minX, y: minY, w: maxX - minX, h: maxY - minY };
        if (bboxIntersectsComp(mainComp, bboxMain) && intersectionRatio(mainComp, bboxMain) >= MIN_IN_RATIO) return true;
      }
    } catch (e) {}
    return false;
  }

  /** For a text layer inside a precomp: find main comp time and bbox in main comp space. Returns { t, bbox, source } or null. */
  function findBestTimeInMainCompForNestedLayer(layer, layerComp, mainComp) {
    var path = getPathToComp(mainComp, layerComp);
    if (!path.length) return null;

    var step = mainComp.frameDuration || (1/24);
    if (!isFinite(step) || step <= 0) step = 1/24;

    // Time in layerComp at main time t_main: walk path and subtract startTimes
    function getPrecompTime(tMain) {
      var t = tMain;
      for (var p = 0; p < path.length; p++) {
        var pl = path[p].layer;
        t = t - (Number(pl.startTime) || 0);
      }
      return t;
    }

    function precompLayerVisibleAt(tMain) {
      var t = tMain;
      for (var p = 0; p < path.length; p++) {
        var pl = path[p].layer;
        try { if (!pl.activeAtTime(t)) return false; } catch (e) { return false; }
        if (layerOpacityAt(pl, t) < MIN_OPACITY) return false;
        if (layerScaleAt(pl, t) < MIN_SCALE) return false;
        t = t - (Number(pl.startTime) || 0);
      }
      return true;
    }

    // Find a main comp time range where the precomp chain is visible
    var firstLayer = path[0].layer;
    var mainWin = getCompScanWindow(mainComp);
    var aMain = Math.max(mainWin.start, Math.max(0, Number(firstLayer.inPoint || 0)));
    var bMain = Math.min(mainWin.end, Math.min(mainComp.duration || 1, Number(firstLayer.outPoint || 1)));
    if (bMain - aMain < 0.01) return null;

    var best = null;
    var layerIn = Number(layer.inPoint || 0);
    var layerOut = Number(layer.outPoint || 0);
    var sumStart = 0;
    for (var si = 0; si < path.length; si++) sumStart += Number(path[si].layer.startTime || 0);

    // Snapshot Marker on the text layer inside the precomp: use it as preferred time (marker time is in precomp time)
    var tPreSnap = getSnapshotMarkerTime(layer);
    if (tPreSnap != null && isFinite(tPreSnap) && tPreSnap >= layerIn - 0.001 && tPreSnap <= layerOut + 0.001 && layerEligible(layer, tPreSnap)) {
      var tMainSnap = tPreSnap + sumStart;
      if (tMainSnap >= aMain - 0.001 && tMainSnap <= bMain + 0.001 && precompLayerVisibleAt(tMainSnap)) {
        try {
          var infoSnap = bboxForLayer(layer, layerComp, tPreSnap);
          var bbSnap = infoSnap ? infoSnap.bbox : null;
          if (!bbSnap) bbSnap = bboxFromPositionRescue(layer, layerComp, tPreSnap);
          if (bbSnap && bbSnap.w >= 2 && bbSnap.h >= 2) {
            var refPreSnap = getEffectiveBoundsForLayerAtTime(layer, layerComp, tPreSnap);
            if (bboxIntersectsRect(bbSnap, refPreSnap) && intersectionRatioRects(bbSnap, refPreSnap) >= MIN_IN_RATIO) {
              var timesSnap = [];
              var tt = tMainSnap;
              for (var q = 0; q < path.length; q++) {
                timesSnap[q] = tt;
                tt = tt - (Number(path[q].layer.startTime) || 0);
              }
              var cornersSnap = [[bbSnap.x, bbSnap.y], [bbSnap.x + bbSnap.w, bbSnap.y], [bbSnap.x + bbSnap.w, bbSnap.y + bbSnap.h], [bbSnap.x, bbSnap.y + bbSnap.h]];
              var minXS = 1e9, maxXS = -1e9, minYS = 1e9, maxYS = -1e9;
              for (var cs = 0; cs < cornersSnap.length; cs++) {
                var pt = cornersSnap[cs].slice();
                for (var ps = path.length - 1; ps >= 0; ps--) {
                  var tr = layerPointToCompAtTime(path[ps].layer, [pt[0], pt[1]], timesSnap[ps]);
                  if (!tr || tr.length < 2) break;
                  pt[0] = tr[0]; pt[1] = tr[1];
                }
                if (pt[0] < minXS) minXS = pt[0]; if (pt[0] > maxXS) maxXS = pt[0];
                if (pt[1] < minYS) minYS = pt[1]; if (pt[1] > maxYS) maxYS = pt[1];
              }
              var bboxMainSnap = { x: minXS, y: minYS, w: maxXS - minXS, h: maxYS - minYS };
              if (bboxIntersectsComp(mainComp, bboxMainSnap) && intersectionRatio(mainComp, bboxMainSnap) >= MIN_IN_RATIO) {
                try { mainComp.time = tMainSnap; app.project.activeItem = mainComp; } catch (eR) {}
                return { t: tMainSnap, bbox: bboxMainSnap, source: (infoSnap && infoSnap.source) ? infoSnap.source : "snapshot-marker" };
              }
            }
          }
        } catch (eSnap) {}
      }
    }

    // Single rule: use second keyframe or midpoint in precomp, then map to main comp. Use text layer's Snapshot marker when present (allowSnapshotMarker true).
    var tPre = getScreenshotTimeForLayer(layer, layerIn, layerOut, true);
    var tMainSingle = tPre + sumStart;
    if (tMainSingle >= aMain - 0.001 && tMainSingle <= bMain + 0.001 && precompLayerVisibleAt(tMainSingle)) {
      try {
        var tPreCheck = getPrecompTime(tMainSingle);
        if (Math.abs(tPreCheck - tPre) < 0.02 && layerEligible(layer, tPre)) {
          var infoSingle = bboxForLayer(layer, layerComp, tPre);
          var bbSingle = infoSingle ? infoSingle.bbox : null;
          if (!bbSingle) bbSingle = bboxFromPositionRescue(layer, layerComp, tPre);
          if (bbSingle && bbSingle.w >= 2 && bbSingle.h >= 2) {
            var refPreSingle = getEffectiveBoundsForLayerAtTime(layer, layerComp, tPre);
            if (bboxIntersectsRect(bbSingle, refPreSingle) && intersectionRatioRects(bbSingle, refPreSingle) >= MIN_IN_RATIO) {
              var timesSingle = [];
              var ttS = tMainSingle;
              for (var q = 0; q < path.length; q++) {
                timesSingle[q] = ttS;
                ttS = ttS - (Number(path[q].layer.startTime) || 0);
              }
              var cornersSingle = [[bbSingle.x, bbSingle.y], [bbSingle.x + bbSingle.w, bbSingle.y], [bbSingle.x + bbSingle.w, bbSingle.y + bbSingle.h], [bbSingle.x, bbSingle.y + bbSingle.h]];
              var minXS = 1e9, maxXS = -1e9, minYS = 1e9, maxYS = -1e9;
              for (var c = 0; c < cornersSingle.length; c++) {
                var pt = cornersSingle[c].slice();
                for (var p = path.length - 1; p >= 0; p--) {
                  var tr = layerPointToCompAtTime(path[p].layer, [pt[0], pt[1]], timesSingle[p]);
                  if (!tr || tr.length < 2) break;
                  pt[0] = tr[0]; pt[1] = tr[1];
                }
                if (pt[0] < minXS) minXS = pt[0]; if (pt[0] > maxXS) maxXS = pt[0];
                if (pt[1] < minYS) minYS = pt[1]; if (pt[1] > maxYS) maxYS = pt[1];
              }
              var bboxMainSingle = { x: minXS, y: minYS, w: maxXS - minXS, h: maxYS - minYS };
              if (bboxIntersectsComp(mainComp, bboxMainSingle) && intersectionRatio(mainComp, bboxMainSingle) >= MIN_IN_RATIO) {
                try { mainComp.time = tMainSingle; app.project.activeItem = mainComp; } catch (eR) {}
                return { t: tMainSingle, bbox: bboxMainSingle, source: (infoSingle && infoSingle.source) || 'nested' };
              }
            }
          }
        }
      } catch (eSingle) {}
    }

    // Typewriter: max text length in precomp so we prefer full-reveal frames
    var precompStep = layerComp.frameDuration || (1/24);
    if (!isFinite(precompStep) || precompStep <= 0) precompStep = 1/24;
    var maxLen = 0;
    for (var tPreScan = layerIn; tPreScan <= layerOut; tPreScan += precompStep) {
      if (!layerEligible(layer, tPreScan)) continue;
      var len = textLenAt(layer, tPreScan);
      if (len > maxLen) maxLen = len;
    }
    // Simple typewriter: only when rect grows (plain text has constant rect). Long constant-length = built-in preset, use tail.
    var tMinPreForFullReveal = layerIn;
    if (maxLen > 0 && isConstantLengthTypewriter(layer, layerComp, maxLen)) {
      if (rectWidthGrowsOverDuration(layer, layerComp)) {
        var fullVisualPre = getTimeOfFullVisualReveal(layer, layerComp);
        if (fullVisualPre && fullVisualPre.t > layerIn) tMinPreForFullReveal = fullVisualPre.t;
        var tailStartPre = getMinTimeInTypewriterTail(layer, false);
        if (tailStartPre > tMinPreForFullReveal) tMinPreForFullReveal = tailStartPre;
      } else if ((layerOut - layerIn) >= 2) {
        var tailStartPre = getMinTimeInTypewriterTail(layer, isAnimatorTypewriter(layer, layerComp, maxLen));
        tMinPreForFullReveal = tailStartPre;
      }
      if (isAnimatorTypewriter(layer, layerComp, maxLen)) {
        var tStart100Pre = getTimeWhenTypewriterFullReveal(layer, layerComp, maxLen);
        if (tStart100Pre != null && tStart100Pre >= layerIn) tMinPreForFullReveal = Math.min(layerOut, Math.max(layerIn, tStart100Pre));
        else {
          var tailAnimatorPre = getMinTimeInTypewriterTail(layer, true);
          if (tailAnimatorPre > tMinPreForFullReveal) tMinPreForFullReveal = tailAnimatorPre;
        }
      }
    }

    // Dedicated track-matte + pause path for nested: when the precomp text layer has a matte, prefer first main-comp time where the layer is paused (in precomp) and visible in matte.
    var matteLayerNested = getMatteLayerAbove(layer, layerComp);
    if (matteLayerNested) {
      // Sample at least every main-comp frame so we don't miss short pauses in nested precomps
      var sampleStepMain = Math.min(mainComp.frameDuration || step, Math.max(step, (bMain - aMain) / 120));
      if (!isFinite(sampleStepMain) || sampleStepMain <= 0) sampleStepMain = step;
      for (var tMain = aMain; tMain <= bMain; tMain += sampleStepMain) {
        if (!precompLayerVisibleAt(tMain)) continue;
        var tPre = getPrecompTime(tMain);
        try { if (!layer.activeAtTime(tPre)) continue; } catch (e) { continue; }
        if (tPre < layerIn || tPre > layerOut) continue;
        if (!layerEligible(layer, tPre)) continue;
        if (!isLayerPausedAt(layer, layerComp, tPre)) continue;
        var refPre = getEffectiveBoundsForLayerAtTime(layer, layerComp, tPre);
        var info = bboxForLayer(layer, layerComp, tPre);
        if (!info || !info.bbox || info.bbox.w < 2 || info.bbox.h < 2) continue;
        var bb = info.bbox;
        if (!bboxIntersectsRect(bb, refPre) || intersectionRatioRects(bb, refPre) < MIN_IN_RATIO) continue;
        if (maxLen > 0 && textLenAt(layer, tPre) < maxLen) continue;
        try {
          var times = [];
          var tt = tMain;
          for (var q = 0; q < path.length; q++) {
            times[q] = tt;
            tt = tt - (Number(path[q].layer.startTime) || 0);
          }
          var corners = [[bb.x, bb.y], [bb.x + bb.w, bb.y], [bb.x + bb.w, bb.y + bb.h], [bb.x, bb.y + bb.h]];
          var minX = 1e9, maxX = -1e9, minY = 1e9, maxY = -1e9;
          for (var c = 0; c < corners.length; c++) {
            var pt = corners[c];
            for (var p = path.length - 1; p >= 0; p--) {
              var tr = layerPointToCompAtTime(path[p].layer, [pt[0], pt[1]], times[p]);
              if (!tr || tr.length < 2) break;
              pt = [tr[0], tr[1]];
            }
            if (pt[0] < minX) minX = pt[0]; if (pt[0] > maxX) maxX = pt[0];
            if (pt[1] < minY) minY = pt[1]; if (pt[1] > maxY) maxY = pt[1];
          }
          var bboxMain = { x: minX, y: minY, w: maxX - minX, h: maxY - minY };
          if (!bboxIntersectsComp(mainComp, bboxMain) || intersectionRatio(mainComp, bboxMain) < MIN_IN_RATIO) continue;
          try { mainComp.time = tMain; app.project.activeItem = mainComp; } catch (eRestore) {}
          return { t: tMain, bbox: bboxMain, source: (info && info.source) || 'nested' };
        } catch (eInner) {}
      }
    }

    var animatorTwNested = maxLen > 0 && isAnimatorTypewriter(layer, layerComp, maxLen);
    var mainStart = aMain, mainEnd = bMain, mainStep = step;
    if (tMinPreForFullReveal > layerIn) { mainStart = bMain; mainEnd = aMain; mainStep = -step; }
    for (var tMain = mainStart; tMinPreForFullReveal > layerIn ? (tMain >= mainEnd) : (tMain <= mainEnd); tMain += mainStep) {
      if (!precompLayerVisibleAt(tMain)) continue;
      var tPre = getPrecompTime(tMain);
      try { if (!layer.activeAtTime(tPre)) continue; } catch (e) { continue; }
      if (tPre < layerIn || tPre > layerOut) continue;
      if (tPre < tMinPreForFullReveal) continue;
      if (!layerEligible(layer, tPre)) continue;
      if (maxLen > 0 && textLenAt(layer, tPre) < maxLen) continue;
      var stable = true;
      var stableFramesRequiredNested = animatorTwNested ? 1 : STABLE_FRAMES;
      for (var k = 1; k <= stableFramesRequiredNested; k++) {
        var tPreK = tPre + k * precompStep;
        if (tPreK > layerOut) break;
        if (!layerEligible(layer, tPreK) || textLenAt(layer, tPreK) < maxLen) { stable = false; break; }
        var infoK = bboxForLayer(layer, layerComp, tPreK);
        var refPreK = getEffectiveBoundsForLayerAtTime(layer, layerComp, tPreK);
        if (!infoK || !infoK.bbox || !bboxIntersectsRect(infoK.bbox, refPreK) || intersectionRatioRects(infoK.bbox, refPreK) < MIN_IN_RATIO) { stable = false; break; }
      }
      if (!stable) continue;

      var info = bboxForLayer(layer, layerComp, tPre);
      var bb = info ? info.bbox : null;
      if (!bb || bb.w < 2 || bb.h < 2) continue;
      var refPre = getEffectiveBoundsForLayerAtTime(layer, layerComp, tPre);
      if (!bboxIntersectsRect(bb, refPre) || intersectionRatioRects(bb, refPre) < MIN_IN_RATIO) continue;

      // Transform bbox from layerComp to mainComp using valueAtTime only (do not set comp.time)
      try {
        var times = [];
        var tt = tMain;
        for (var q = 0; q < path.length; q++) {
          times[q] = tt;
          tt = tt - (Number(path[q].layer.startTime) || 0);
        }
        var corners = [
          [bb.x, bb.y],
          [bb.x + bb.w, bb.y],
          [bb.x + bb.w, bb.y + bb.h],
          [bb.x, bb.y + bb.h]
        ];
        var mainCorners = [];
        for (var c = 0; c < corners.length; c++) {
          var pt = corners[c];
          for (var p = path.length - 1; p >= 0; p--) {
            var tr = layerPointToCompAtTime(path[p].layer, [pt[0], pt[1]], times[p]);
            if (!tr || tr.length < 2) break;
            pt = [tr[0], tr[1]];
          }
          mainCorners.push(pt);
        }
        var minX = mainCorners[0][0], maxX = mainCorners[0][0], minY = mainCorners[0][1], maxY = mainCorners[0][1];
        for (var m = 1; m < mainCorners.length; m++) {
          minX = Math.min(minX, mainCorners[m][0]); maxX = Math.max(maxX, mainCorners[m][0]);
          minY = Math.min(minY, mainCorners[m][1]); maxY = Math.max(maxY, mainCorners[m][1]);
        }
        var bboxMain = { x: minX, y: minY, w: maxX - minX, h: maxY - minY };
        if (!bboxIntersectsComp(mainComp, bboxMain)) continue;
        if (intersectionRatio(mainComp, bboxMain) < MIN_IN_RATIO) continue;

        var score = intersectionRatio(mainComp, bboxMain);
        var op = layerOpacityAt(layer, tPre) / 100;
        var sc = Math.min(1, layerScaleAt(layer, tPre));
        score *= (0.5 + 0.5 * op);
        score *= (0.3 + 0.7 * sc);
        if (!best || score > best.score) {
          best = { t: tMain, bbox: bboxMain, source: (info && info.source) || 'nested', score: score };
          if (tMinPreForFullReveal > layerIn) break;
        }
      } catch (e) {}
    }

    if (best) {
      try { mainComp.time = best.t; app.project.activeItem = mainComp; } catch(eRestore){}
      return { t: best.t, bbox: best.bbox, source: best.source };
    }

    // Pass 2: same as strict but without stability (typewriter may not have 3 stable frames). When typewriter tail, prefer latest.
    var mainStart2 = aMain, mainEnd2 = bMain, mainStep2 = step;
    if (tMinPreForFullReveal > layerIn) { mainStart2 = bMain; mainEnd2 = aMain; mainStep2 = -step; }
    for (var tMain2 = mainStart2; tMinPreForFullReveal > layerIn ? (tMain2 >= mainEnd2) : (tMain2 <= mainEnd2); tMain2 += mainStep2) {
      if (!precompLayerVisibleAt(tMain2)) continue;
      var tPre2 = getPrecompTime(tMain2);
      try { if (!layer.activeAtTime(tPre2)) continue; } catch (e) { continue; }
      if (tPre2 < layerIn || tPre2 > layerOut) continue;
      if (tPre2 < tMinPreForFullReveal) continue;
      if (!layerEligible(layer, tPre2)) continue;
      if (maxLen > 0 && textLenAt(layer, tPre2) < maxLen) continue;
      var info2 = bboxForLayer(layer, layerComp, tPre2);
      var bb2 = info2 ? info2.bbox : null;
      if (!bb2 || bb2.w < 2 || bb2.h < 2) continue;
      var refPre2 = getEffectiveBoundsForLayerAtTime(layer, layerComp, tPre2);
      if (!bboxIntersectsRect(bb2, refPre2) || intersectionRatioRects(bb2, refPre2) < MIN_IN_RATIO) continue;
      try {
        var times2 = [];
        var tt2 = tMain2;
        for (var q = 0; q < path.length; q++) {
          times2[q] = tt2;
          tt2 = tt2 - (Number(path[q].layer.startTime) || 0);
        }
        var corners2 = [
          [bb2.x, bb2.y],
          [bb2.x + bb2.w, bb2.y],
          [bb2.x + bb2.w, bb2.y + bb2.h],
          [bb2.x, bb2.y + bb2.h]
        ];
        var mainCorners2 = [];
        for (var c = 0; c < corners2.length; c++) {
          var pt = corners2[c];
          for (var p = path.length - 1; p >= 0; p--) {
            var tr = layerPointToCompAtTime(path[p].layer, [pt[0], pt[1]], times2[p]);
            if (!tr || tr.length < 2) break;
            pt = [tr[0], tr[1]];
          }
          mainCorners2.push(pt);
        }
        var minX2 = mainCorners2[0][0], maxX2 = mainCorners2[0][0], minY2 = mainCorners2[0][1], maxY2 = mainCorners2[0][1];
        for (var m = 1; m < mainCorners2.length; m++) {
          minX2 = Math.min(minX2, mainCorners2[m][0]); maxX2 = Math.max(maxX2, mainCorners2[m][0]);
          minY2 = Math.min(minY2, mainCorners2[m][1]); maxY2 = Math.max(maxY2, mainCorners2[m][1]);
        }
        var bboxMain2 = { x: minX2, y: minY2, w: maxX2 - minX2, h: maxY2 - minY2 };
        if (!bboxIntersectsComp(mainComp, bboxMain2) || intersectionRatio(mainComp, bboxMain2) < MIN_IN_RATIO) continue;
        var score2 = intersectionRatio(mainComp, bboxMain2);
        var op2 = layerOpacityAt(layer, tPre2) / 100;
        var sc2 = Math.min(1, layerScaleAt(layer, tPre2));
        score2 *= (0.5 + 0.5 * op2);
        score2 *= (0.3 + 0.7 * sc2);
        if (!best || score2 > best.score) {
          best = { t: tMain2, bbox: bboxMain2, source: (info2 && info2.source) || 'nested', score: score2 };
          if (tMinPreForFullReveal > layerIn) break;
        }
      } catch (e) {}
    }

    if (best) {
      try { mainComp.time = best.t; app.project.activeItem = mainComp; } catch(eRestore){}
      return { t: best.t, bbox: best.bbox, source: best.source };
    }

    // Relaxed pass: still require layerEligible, precomp chain visible, and full typewriter reveal
    var sampleCount = Math.min(8, Math.max(1, Math.floor((bMain - aMain) / step)));
    for (var n = 0; n < sampleCount; n++) {
      var tMain = aMain + (bMain - aMain) * (n + 1) / (sampleCount + 1);
      if (!precompLayerVisibleAt(tMain)) continue;
      var tPre = getPrecompTime(tMain);
      try { if (!layer.activeAtTime(tPre)) continue; } catch (e) { continue; }
      if (tPre < layerIn || tPre > layerOut) continue;
      if (tPre < tMinPreForFullReveal) continue;
      if (!layerEligible(layer, tPre)) continue;
      if (maxLen > 0 && textLenAt(layer, tPre) < maxLen) continue;
      var info = bboxForLayer(layer, layerComp, tPre);
      var bb = info ? info.bbox : null;
      if (!bb || bb.w < 2 || bb.h < 2) continue;
      var refPre = getEffectiveBoundsForLayerAtTime(layer, layerComp, tPre);
      if (!bboxIntersectsRect(bb, refPre) || intersectionRatioRects(bb, refPre) < MIN_IN_RATIO) continue;
      try {
        var times = [];
        var tt = tMain;
        for (var q = 0; q < path.length; q++) {
          times[q] = tt;
          tt = tt - (Number(path[q].layer.startTime) || 0);
        }
        var corners = [
          [bb.x, bb.y],
          [bb.x + bb.w, bb.y],
          [bb.x + bb.w, bb.y + bb.h],
          [bb.x, bb.y + bb.h]
        ];
        var mainCorners = [];
        for (var c = 0; c < corners.length; c++) {
          var pt = corners[c];
          for (var p = path.length - 1; p >= 0; p--) {
            var tr = layerPointToCompAtTime(path[p].layer, [pt[0], pt[1]], times[p]);
            if (!tr || tr.length < 2) break;
            pt = [tr[0], tr[1]];
          }
          mainCorners.push(pt);
        }
        var minX = mainCorners[0][0], maxX = mainCorners[0][0], minY = mainCorners[0][1], maxY = mainCorners[0][1];
        for (var m = 1; m < mainCorners.length; m++) {
          minX = Math.min(minX, mainCorners[m][0]); maxX = Math.max(maxX, mainCorners[m][0]);
          minY = Math.min(minY, mainCorners[m][1]); maxY = Math.max(maxY, mainCorners[m][1]);
        }
        var bboxMain = { x: minX, y: minY, w: maxX - minX, h: maxY - minY };
        if (bboxIntersectsComp(mainComp, bboxMain) && intersectionRatio(mainComp, bboxMain) >= MIN_IN_RATIO) {
          try { mainComp.time = tMain; app.project.activeItem = mainComp; } catch(eR){}
          return { t: tMain, bbox: bboxMain, source: (info && info.source) || 'nested' };
        }
      } catch (e) {}
    }

    // Fallback: use precomp's fallback time/bbox and map to main comp; require scale/opacity meet minimums
    try {
      var fall = getFallbackTimeAndBbox(layer, layerComp);
      if (!fall || !fall.bbox) return null;
      var tPre = fall.t;
      if (layerScaleAt(layer, tPre) < MIN_SCALE) return null;
      if (layerOpacityAt(layer, tPre) < MIN_OPACITY) return null;
      var sumStart = 0;
      for (var s = 0; s < path.length; s++) sumStart += Number(path[s].layer.startTime || 0);
      var tMainFall = tPre + sumStart;
      if (tMainFall < aMain || tMainFall > bMain) return null;
      if (!precompLayerVisibleAt(tMainFall)) return null;

      var times = [];
      var tt = tMainFall;
      for (var q = 0; q < path.length; q++) {
        times[q] = tt;
        tt = tt - (Number(path[q].layer.startTime) || 0);
      }
      var bb = fall.bbox;
      var corners = [
        [bb.x, bb.y],
        [bb.x + bb.w, bb.y],
        [bb.x + bb.w, bb.y + bb.h],
        [bb.x, bb.y + bb.h]
      ];
      var mainCorners = [];
      for (var c = 0; c < corners.length; c++) {
        var pt = corners[c];
        for (var p = path.length - 1; p >= 0; p--) {
          var tr = layerPointToCompAtTime(path[p].layer, [pt[0], pt[1]], times[p]);
          if (!tr || tr.length < 2) break;
          pt = [tr[0], tr[1]];
        }
        mainCorners.push(pt);
      }
      var minX = mainCorners[0][0], maxX = mainCorners[0][0], minY = mainCorners[0][1], maxY = mainCorners[0][1];
      for (var m = 1; m < mainCorners.length; m++) {
        minX = Math.min(minX, mainCorners[m][0]); maxX = Math.max(maxX, mainCorners[m][0]);
        minY = Math.min(minY, mainCorners[m][1]); maxY = Math.max(maxY, mainCorners[m][1]);
      }
      var bboxMain = { x: minX, y: minY, w: Math.max(2, maxX - minX), h: Math.max(2, maxY - minY) };
      if (bboxIntersectsComp(mainComp, bboxMain) && intersectionRatio(mainComp, bboxMain) >= MIN_IN_RATIO) {
        try { mainComp.time = tMainFall; app.project.activeItem = mainComp; } catch(eR){}
        return { t: tMainFall, bbox: bboxMain, source: fall.source || 'nested' };
      }
    } catch (e) {}

    return null;
  }

  function smartScanTimeline(setStatus, setProgress, compOptional){
    if (!STATE.projectId) return false;

    var comp = (compOptional && compOptional instanceof CompItem) ? compOptional : getActiveComp();
    if (!comp) return false;

    if (!STATE.fileKey) STATE.fileKey = safeFileKeyForComp(comp);

    /** Clear Blinking Cursor dimension cache so each scan starts fresh. */
    __blinkFullTextSizeCache = {};

    // Scale comp to height 1080 first (Scale Composition logic) so all Crowdin screenshots same height; bbox and PNG in same space. Restore in finally.
    var scaleFactorApplied = null;
    if (Math.abs(comp.height - CROWDIN_EXPORT_MAX_H) > 0.5) {
      var scaleToHeight = CROWDIN_EXPORT_MAX_H / comp.height;
      if (scaleCompositionByFactor(comp, scaleToHeight)) scaleFactorApplied = scaleToHeight;
    }

    try {
    var layerComps = getTextLayersForExport(comp);
    if (typeof DEBUG_TYPEWRITER_LOG !== "undefined" && DEBUG_TYPEWRITER_LOG) {
      try {
        var dbgFile = new File(Folder.myDocuments.fsName + "/Crowdin_typewriter_debug.txt");
        if (dbgFile.exists) dbgFile.remove();
      } catch (e) {}
    }
    if (!layerComps.length) {
      setStatus("Smart Scan skipped (no text layers in selection or comp).");
      if (setProgress) setProgress(0, 0, "Ready");
      return false;
    }

    var scanStartMs = new Date().getTime();
    function elapsedSec() { return ((new Date().getTime() - scanStartMs) / 1000).toFixed(1); }
    function statusWithTime(msg) { setStatus(msg); }
    function progressWithTime(current, total, msg) { if (setProgress) setProgress(current, total, msg || "Ready"); }

    setStatus("Timeline Scan: Analyzing...");
    progressWithTime(0, layerComps.length, "Analyzing…");

    var TMP = Folder.temp;
    var step = comp.frameDuration || (1 / 24);
    if (!isFinite(step) || step <= 0) step = 1 / 24;

    // First pass: get best time + bbox for each layer
    var candidates = [];
    for (var i = 0; i < layerComps.length; i++) {
      progressWithTime(i, layerComps.length, "Layer " + (i + 1) + " of " + layerComps.length);
      var L = layerComps[i].layer;
      var layerComp = layerComps[i].comp;
      if (L.matchName !== "ADBE Text Layer") continue;
      var id = makeStringKey(layerComp, L);
      var best = null;
      if (layerComp === comp) {
        best = findBestTime(L, comp);
        if (!best) best = getFallbackTimeAndBbox(L, comp);
        if (!best) best = getRescueTailTimeAndBbox(L, comp);
      } else {
        best = findBestTimeInMainCompForNestedLayer(L, layerComp, comp);
      }
      if (!best) continue;
      // Typewriter Blinking Cursor: force bbox to full-text size at 100% so the workaround always applies (findBestTime may have used another path).
      if (hasBlinkingCursorTypewriterEffect(L)) {
        var stepFrame = (layerComp.frameDuration || 1/24);
        if (isFinite(stepFrame) && stepFrame > 0) best.t = Math.round(best.t / stepFrame) * stepFrame;
        var bbBlink = bboxEstimateFromTextDocNoCursor(L, best.t, layerComp);
        if (bbBlink && bbBlink.w >= 2 && bbBlink.h >= 2)
          best = { t: best.t, bbox: bbBlink, source: "blinkFullText" };
      }
      if (layerComp !== comp) {
        var pathCheck = getPathToComp(comp, layerComp);
        if (pathCheck.length && pathCheck[0].layer && !pathCheck[0].layer.activeAtTime(best.t)) continue;
        var tPreCheck = best.t;
        for (var pi = 0; pi < pathCheck.length; pi++) tPreCheck -= Number(pathCheck[pi].layer.startTime || 0);
        try { if (!L.activeAtTime(tPreCheck)) continue; } catch (e) { continue; }
      }
      var layerText = "";
      try { var sp = getSourceTextProp(L); if (sp) layerText = getCompletedTextForLayer(L, layerComp); } catch (e) {}
      candidates.push({ layer: L, layerComp: layerComp, id: id, best: best, layerText: layerText });
    }

    // Ensure each layer gets a unique capture time (and thus a different screenshot) so Crowdin receives distinct context per string.
    var stepFrame = comp.frameDuration || (1 / 24);
    if (!isFinite(stepFrame) || stepFrame <= 0) stepFrame = 1 / 24;
    var compEnd = 0;
    try { compEnd = Number(comp.duration) || 0; } catch (e) {}
    var usedTime = {};
    for (var ci = 0; ci < candidates.length; ci++) {
      var c = candidates[ci];
      var tOriginal = c.best.t;
      var layerInMain = 0;
      var layerOutMain = compEnd;
      if (c.layerComp === comp) {
        layerInMain = Math.max(0, Number(c.layer.inPoint) || 0);
        layerOutMain = Math.min(compEnd, Math.max(0, Number(c.layer.outPoint) || compEnd));
      } else {
        var pathOut = getPathToComp(comp, c.layerComp);
        var sumStart = 0;
        for (var ps = 0; ps < pathOut.length; ps++) sumStart += Number(pathOut[ps].layer.startTime || 0);
        layerInMain = Math.max(0, (Number(c.layer.inPoint) || 0) + sumStart);
        layerOutMain = Math.min(compEnd, Math.max(0, (Number(c.layer.outPoint) || 0) + sumStart));
      }
      var preferredTimesMain = [];
      if (c.layerComp === comp) {
        preferredTimesMain = getPreferredScreenshotTimes(c.layer, layerInMain, layerOutMain, true);
      } else {
        var pathOutPref = getPathToComp(comp, c.layerComp);
        var sumStartPref = 0;
        for (var psp = 0; psp < pathOutPref.length; psp++) sumStartPref += Number(pathOutPref[psp].layer.startTime || 0);
        var prefsPre = getPreferredScreenshotTimes(c.layer, Number(c.layer.inPoint) || 0, Number(c.layer.outPoint) || compEnd, true);
        for (var pp = 0; pp < prefsPre.length; pp++) preferredTimesMain.push(prefsPre[pp] + sumStartPref);
      }
      var tRound = Math.round(tOriginal / stepFrame) * stepFrame;
      var key = tRound.toFixed(6);
      var foundSlot = false;
      for (var idx = 0; idx < preferredTimesMain.length; idx++) {
        var tTry = Math.round(preferredTimesMain[idx] / stepFrame) * stepFrame;
        if (tTry < layerInMain - 0.001 || tTry > layerOutMain + 0.001) continue;
        var kTry = tTry.toFixed(6);
        if (!usedTime[kTry]) {
          tRound = tTry;
          key = kTry;
          foundSlot = true;
          break;
        }
      }
      if (!foundSlot) {
        while (usedTime[key] && (tRound + stepFrame <= layerOutMain + 0.001)) {
          tRound += stepFrame;
          key = tRound.toFixed(6);
        }
        if (usedTime[key] && (tRound > layerOutMain + 0.001) && (tRound - stepFrame >= layerInMain - 0.001)) {
          tRound = Math.round(tOriginal / stepFrame) * stepFrame - stepFrame;
          while (tRound >= layerInMain - 0.001) {
            key = tRound.toFixed(6);
            if (!usedTime[key]) break;
            tRound -= stepFrame;
          }
        }
        key = tRound.toFixed(6);
      }
      usedTime[key] = true;
      var shouldApplyNudge = (Math.abs(tRound - tOriginal) > 0.001);
      if (shouldApplyNudge) {
        c.best.t = tRound;
        var inRange = (tRound >= layerInMain - 0.001 && tRound <= layerOutMain + 0.001);
        if (inRange) {
          if (c.layerComp === comp) {
            var infoNudge = bboxForLayer(c.layer, c.layerComp, tRound);
            if (infoNudge && infoNudge.bbox && infoNudge.bbox.w >= 2 && infoNudge.bbox.h >= 2)
              c.best = { t: tRound, bbox: infoNudge.bbox, source: infoNudge.source || c.best.source };
            else
              c.best = { t: tRound, bbox: c.best.bbox, source: c.best.source };
          } else {
            var pathNudge = getPathToComp(comp, c.layerComp);
            var tPreNudge = tRound;
            for (var pj = 0; pj < pathNudge.length; pj++) tPreNudge -= Number(pathNudge[pj].layer.startTime || 0);
            var infoPre = bboxForLayer(c.layer, c.layerComp, tPreNudge);
            if (infoPre && infoPre.bbox && pathNudge.length > 0) {
              var bbPre = infoPre.bbox;
              var timesNudge = [tRound];
              for (var pk = 1; pk < pathNudge.length; pk++) timesNudge.push(timesNudge[timesNudge.length - 1] - (Number(pathNudge[pk - 1].layer.startTime) || 0));
              var cornersNudge = [[bbPre.x, bbPre.y], [bbPre.x + bbPre.w, bbPre.y], [bbPre.x + bbPre.w, bbPre.y + bbPre.h], [bbPre.x, bbPre.y + bbPre.h]];
              var mainCornersNudge = [];
              for (var pc = 0; pc < cornersNudge.length; pc++) {
                var ptN = cornersNudge[pc].slice();
                for (var pd = pathNudge.length - 1; pd >= 0; pd--) {
                  var trN = layerPointToCompAtTime(pathNudge[pd].layer, ptN, timesNudge[pd]);
                  if (!trN || trN.length < 2) break;
                  ptN = trN;
                }
                if (ptN && ptN.length >= 2) mainCornersNudge.push(ptN);
              }
              if (mainCornersNudge.length >= 2) {
                var minXN = mainCornersNudge[0][0], maxXN = mainCornersNudge[0][0], minYN = mainCornersNudge[0][1], maxYN = mainCornersNudge[0][1];
                for (var pe = 1; pe < mainCornersNudge.length; pe++) {
                  minXN = Math.min(minXN, mainCornersNudge[pe][0]); maxXN = Math.max(maxXN, mainCornersNudge[pe][0]);
                  minYN = Math.min(minYN, mainCornersNudge[pe][1]); maxYN = Math.max(maxYN, mainCornersNudge[pe][1]);
                }
                c.best = { t: tRound, bbox: { x: minXN, y: minYN, w: Math.max(2, maxXN - minXN), h: Math.max(2, maxYN - minYN) }, source: c.best.source };
              } else {
                c.best = { t: tRound, bbox: c.best.bbox, source: c.best.source };
              }
            } else {
              c.best = { t: tRound, bbox: c.best.bbox, source: c.best.source };
            }
          }
        } else {
          c.best = { t: tRound, bbox: c.best.bbox, source: c.best.source };
        }
      }
    }

    if (!candidates.length) {
      setStatus("Timeline Scan: nothing to capture.");
      progressWithTime(0, layerComps.length, "Ready");
      alertIf(
        "No on-screen moments found.\n\n" +
        "This usually means geometry failed for all layers (e.g. off-screen or collapsed).\n" +
        "Check that text layers are visible and on-screen in their comp (or precomp)."
      );
      return false;
    }

    // Group by (comp, layer) + candidate index so each text layer gets its own screenshot at ITS best time.
    var groups = {};
    for (var g = 0; g < candidates.length; g++) {
      var k = comp.id + "_" + candidates[g].id + "_" + g;
      groups[k] = [candidates[g]];
    }

    // Deterministic order: by candidate index (0,1,2...) so export order matches candidates and uniqueExportTimeForGroup(gi) aligns.
    var groupKeys = [];
    for (var g = 0; g < candidates.length; g++) {
      var k = comp.id + "_" + candidates[g].id + "_" + g;
      groupKeys.push(k);
    }

    // Ensure we export at a unique time per group so each string gets a different screenshot (even if dedupe assigned the same frame).
    // Use a spread step (e.g. 0.5s) so layers with same text/opacity don't get identical frames (one-frame nudge often looks the same).
    var stepFrameExport = comp.frameDuration || (1 / 24);
    if (!isFinite(stepFrameExport) || stepFrameExport <= 0) stepFrameExport = 1 / 24;
    var exportSpreadSec = 0.5;
    var exportSpreadFrames = Math.max(1, Math.round(exportSpreadSec / stepFrameExport));
    var stepExportUnique = exportSpreadFrames * stepFrameExport;
    var compEndExport = 0;
    try { compEndExport = Number(comp.duration) || 0; } catch (e) {}
    var usedExportTime = {};
    function uniqueExportTimeForGroup(bestT, groupIndex, layerInMain, layerOutMain) {
      // We already assigned unique capture times in the dedupe loop (second keyframe or midpoint per layer).
      // Use bestT as-is and only nudge when that frame is already taken by another export, so we stay on the keyframe.
      var t = Math.round(bestT / stepFrameExport) * stepFrameExport;
      if (t < 0) t = 0;
      var layerIn = (layerInMain != null && isFinite(layerInMain)) ? layerInMain : 0;
      var layerOut = (layerOutMain != null && isFinite(layerOutMain)) ? layerOutMain : (compEndExport > 0 ? compEndExport : 1e6);
      if (compEndExport > 0 && layerOut > compEndExport) layerOut = compEndExport;
      if (t < layerIn - 0.001 || t > layerOut + 0.001) {
        t = Math.max(layerIn, Math.min(layerOut, (layerIn + layerOut) / 2));
        t = Math.round(t / stepFrameExport) * stepFrameExport;
      }
      var key = t.toFixed(6);
      while (usedExportTime[key]) {
        t += stepExportUnique;
        t = Math.round(t / stepFrameExport) * stepFrameExport;
        if (t > layerOut + 0.001 && layerOut >= layerIn) t = layerIn;
        if (t < layerIn - 0.001) t = layerIn;
        key = t.toFixed(6);
      }
      usedExportTime[key] = true;
      return t;
    }

    var okCount = 0;
    var captured = 0;
    var totalExportMs = 0;
    var totalUploadMs = 0;
    var allPendingHttpPaths = [];
    var allBackgroundCommands = [];
    var pngFilesToRemove = [];
    var boxesFilesToRemoveAll = [];
    var batFilesToRemove = [];
    var tFirstUploadStart = null;

    // Always force Half resolution for Crowdin export (AE dropdown: 2 = Half). All comp sizes.
    var CROWDIN_FORCE_HALF_RESOLUTION = 2;  // AE resolution divisor: 2 = Half (1 = Full, 4 = Quarter)
    var compResBefore = null;
    try { compResBefore = comp.resolutionFactor; } catch(eR0){}
    var scanResolutionFactor = 1 / CROWDIN_FORCE_HALF_RESOLUTION;  // 0.5 for bbox/ssW/ssH math
    try { comp.resolutionFactor = [CROWDIN_FORCE_HALF_RESOLUTION, CROWDIN_FORCE_HALF_RESOLUTION]; } catch(eR1){}

    // Warm server (e.g. Render.com cold start) so first scan-frame upload isn't slow
    try { curlGet(SERVER_BASE + "/"); } catch(eWarm){}

    try {
    if (typeof EXPORT_ONLY_AND_MANIFEST !== "undefined" && EXPORT_ONLY_AND_MANIFEST) {
      var manifest = [];
      var batchPngsToRemove = [];
      var batchBoxesToRemove = [];
      for (var gi = 0; gi < groupKeys.length; gi++) {
        var gk = groupKeys[gi];
        var group = groups[gk];
        setStatus("Timeline Scan: Export " + (gi + 1) + "/" + groupKeys.length + "...");
        if (setProgress) setProgress(gi, groupKeys.length, "Export " + (gi + 1) + "/" + groupKeys.length);
        try { app.refresh(); } catch (eEx) {}
        var first = group[0];
        var best = first.best;
        var TS = "" + (new Date().getTime()) + "_" + gi;
        var pngFile = new File(TMP.fsName + "/ct_scan_" + TS + ".png");
        var safeGroupKey = String(gk).replace(/[^\w.\-]+/g, "_");
        var layerInExp = 0, layerOutExp = compEndExport;
        if (first.layerComp === comp) {
          layerInExp = Math.max(0, Number(first.layer.inPoint) || 0);
          layerOutExp = Math.min(compEndExport, Math.max(0, Number(first.layer.outPoint) || compEndExport));
        } else {
          var pathOutE = getPathToComp(comp, first.layerComp);
          var sumStartE = 0;
          for (var pse = 0; pse < pathOutE.length; pse++) sumStartE += Number(pathOutE[pse].layer.startTime || 0);
          layerInExp = Math.max(0, (Number(first.layer.inPoint) || 0) + sumStartE);
          layerOutExp = Math.min(compEndExport, Math.max(0, (Number(first.layer.outPoint) || 0) + sumStartE));
        }
        var tExport = uniqueExportTimeForGroup(best.t, gi, layerInExp, layerOutExp);
        try { comp.time = tExport; app.project.activeItem = comp; } catch (eForce) {}
        try { app.refresh(); } catch (ePre) {}
        var okPng = exportCompPngAtTime(comp, tExport, pngFile, true);
        if (!okPng) {
          try { app.refresh(); } catch (eRetry) {}
          okPng = exportCompPngAtTime(comp, tExport, pngFile, true);
        }
        if (!okPng) { try { if (pngFile.exists) pngFile.remove(); } catch (e0) {} continue; }
        tryCompressPngForUpload(pngFile, SCAN_PNG_QUALITY);
        var scale = scanResolutionFactor;
        var ssW = Math.round(comp.width * scale);
        var ssH = Math.round(comp.height * scale);
        var seenId = {};
        for (var bi = 0; bi < group.length; bi++) {
          var c = group[bi];
          if (seenId[c.id]) continue;
          seenId[c.id] = true;
          var bb = c.best.bbox;
          var bbExport = compBboxToExportBbox(bb, scale);
          var boxes = [{ id: c.id, stringIdentifier: c.id, layerName: (c.layer && c.layer.name) ? String(c.layer.name).replace(/[\r\n]+/g, " ") : null, bbox: bbExport, source: c.best.source || null, text: c.layerText || null }];
          var safeId = String(c.id).replace(/[^\w.\-]+/g, "_");
          var ssNameForLayer = crowdinScreenshotName(STATE.fileKey + "__" + safeGroupKey + "__" + safeId + "__t" + Math.round(tExport * 1000));
          var fBoxes = new File(TMP.fsName + "/ct_boxes_" + TS + "_" + bi + ".json");
          writeTextFile(fBoxes, jsonStringifyMini(boxes));
          manifest.push({ projectId: STATE.projectId, fileKey: STATE.fileKey, t: "" + Math.round(best.t * 1000), ssName: ssNameForLayer, ssWidth: "" + ssW, ssHeight: "" + ssH, pngPath: pngFile.fsName, boxesPath: fBoxes.fsName });
          batchBoxesToRemove.push(fBoxes);
        }
        batchPngsToRemove.push(pngFile);
      }
      if (manifest.length === 0) {
        for (var i = 0; i < batchPngsToRemove.length; i++) try { if (batchPngsToRemove[i].exists) batchPngsToRemove[i].remove(); } catch(e){}
        for (var i = 0; i < batchBoxesToRemove.length; i++) try { if (batchBoxesToRemove[i].exists) batchBoxesToRemove[i].remove(); } catch(e){}
        setStatus("Timeline Scan: nothing to upload.");
        if (setProgress) setProgress(0, 0, "Ready");
        if (compResBefore && compResBefore.length === 2) try { comp.resolutionFactor = compResBefore; } catch(eR2){}
        return false;
      }
      setStatus("Timeline Scan: uploading " + manifest.length + " item(s) in parallel...");
      if (setProgress) setProgress(groupKeys.length, groupKeys.length, "Uploading...");
      try { app.refresh(); } catch(e){}
      var uploadCommands = [];
      for (var mi = 0; mi < manifest.length; mi++) {
        var item = manifest[mi];
        var suffix = "batch_" + (new Date().getTime()) + "_" + mi;
        var one = curlPostMultipartBuild(
          EP_SCAN_FRAME,
          [
            { name: "projectId", value: item.projectId },
            { name: "fileKey", value: item.fileKey },
            { name: "t", value: item.t },
            { name: "ssName", value: item.ssName },
            { name: "ssWidth", value: item.ssWidth },
            { name: "ssHeight", value: item.ssHeight }
          ],
          [
            { name: "png", path: item.pngPath, mime: "image/png" },
            { name: "boxes", path: item.boxesPath, mime: "application/json" }
          ],
          suffix
        );
        uploadCommands.push(one);
      }
      // Run manifest uploads sequentially instead of in parallel so that the
      // server processes each /ae/scan-frame in order for this comp/fileKey.
      var codes = [];
      var responseBodies = [];
      for (var mi = 0; mi < uploadCommands.length; mi++) {
        var oneCmd = uploadCommands[mi];
        var res = runParallelScanUploads([oneCmd], { keepBodies: true });
        var cArr = res.codes || res;
        var bArr = res.bodies || [];
        var code = (cArr && cArr.length > 0) ? cArr[0] : "0";
        var body = (bArr && bArr.length > 0) ? bArr[0] : "";
        codes.push(code);
        responseBodies.push(body);
        if (code !== "200") {
          // Best-effort single retry for this command
          $.sleep(300);
          var retryRes = runParallelScanUploads([oneCmd], { keepBodies: true });
          var rCodes = retryRes.codes || retryRes;
          var rBodies = retryRes.bodies || [];
          var rCode = (rCodes && rCodes.length > 0) ? rCodes[0] : code;
          var rBody = (rBodies && rBodies.length > 0) ? rBodies[0] : body;
          codes[codes.length - 1] = rCode;
          responseBodies[responseBodies.length - 1] = rBody;
        }
      }
      var batchOkCount = 0;
      for (var h = 0; h < codes.length; h++) {
        if (codes[h] === "200") batchOkCount++;
      }
      var responsesPath = Folder.myDocuments.fsName + "/Crowdin_scan_responses.json";
      try {
        var arr = [];
        for (var r = 0; r < codes.length; r++) {
          arr.push({ index: r, httpCode: codes[r], body: (responseBodies[r] != null) ? responseBodies[r] : "" });
        }
        writeTextFile(new File(responsesPath), jsonStringifyMini(arr));
      } catch (eSave) {}
      for (var i = 0; i < batchPngsToRemove.length; i++) try { if (batchPngsToRemove[i].exists) batchPngsToRemove[i].remove(); } catch(e){}
      for (var i = 0; i < batchBoxesToRemove.length; i++) try { if (batchBoxesToRemove[i].exists) batchBoxesToRemove[i].remove(); } catch(e){}
      setStatus("Timeline Scan complete (batch) " + batchOkCount + "/" + manifest.length + " uploaded. Responses: " + responsesPath);
      if (setProgress) setProgress(layerComps.length, layerComps.length, "Complete");
      if (compResBefore && compResBefore.length === 2) try { comp.resolutionFactor = compResBefore; } catch(eR2){}
      return batchOkCount > 0;
    }

    for (var gi = 0; gi < groupKeys.length; gi++) {
      var gk = groupKeys[gi];
      var group = groups[gk];
      setStatus("Timeline Scan: Screenshot " + (gi + 1) + "/" + groupKeys.length);
      if (setProgress) setProgress(gi, groupKeys.length, "Screenshot " + (gi + 1) + " of " + groupKeys.length + " (" + group.length + " layer(s))");
      try { app.refresh(); } catch (eRefreshStart) {}
      var first = group[0];
      var best = first.best;
      captured += group.length;

      var TS = "" + (new Date().getTime()) + "_" + gi;
      var pngFile = new File(TMP.fsName + "/ct_scan_" + TS + ".png");
      var safeGroupKey = String(gk).replace(/[^\w.\-]+/g, "_");

      var layerInExp = 0, layerOutExp = compEndExport;
      if (first.layerComp === comp) {
        layerInExp = Math.max(0, Number(first.layer.inPoint) || 0);
        layerOutExp = Math.min(compEndExport, Math.max(0, Number(first.layer.outPoint) || compEndExport));
      } else {
        var pathOutE = getPathToComp(comp, first.layerComp);
        var sumStartE = 0;
        for (var pse = 0; pse < pathOutE.length; pse++) sumStartE += Number(pathOutE[pse].layer.startTime || 0);
        layerInExp = Math.max(0, (Number(first.layer.inPoint) || 0) + sumStartE);
        layerOutExp = Math.min(compEndExport, Math.max(0, (Number(first.layer.outPoint) || 0) + sumStartE));
      }
      var tExport = uniqueExportTimeForGroup(best.t, gi, layerInExp, layerOutExp);
      try { comp.time = tExport; app.project.activeItem = comp; } catch (eForce) {}
      var tExport0 = (new Date()).getTime();
      var okPng = exportCompPngAtTime(comp, tExport, pngFile, true);
      var tExport1 = (new Date()).getTime();
      if (!okPng) { try { if (pngFile.exists) pngFile.remove(); } catch (e0) {} continue; }
      tryCompressPngForUpload(pngFile, SCAN_PNG_QUALITY);
      var scale = scanResolutionFactor;
      var ssW = Math.round(comp.width * scale);
      var ssH = Math.round(comp.height * scale);

      // One upload per layer
      var seenId = {};
      var uploadCommands = [];
      var boxesFilesToRemove = [];
      var uploadIdx = 0;
      for (var bi = 0; bi < group.length; bi++) {
        var c = group[bi];
        if (seenId[c.id]) continue;
        seenId[c.id] = true;
        var bb = c.best.bbox;
        var bbExport = compBboxToExportBbox(bb, scale);
        var boxes = [{
          id: c.id,
          stringIdentifier: c.id,
          layerName: (c.layer && c.layer.name) ? String(c.layer.name).replace(/[\r\n]+/g, " ") : null,
          bbox: bbExport,
          source: c.best.source || null,
          text: c.layerText || null
        }];
        var safeId = String(c.id).replace(/[^\w.\-]+/g, "_");
        var ssNameForLayer = crowdinScreenshotName(STATE.fileKey + "__" + safeGroupKey + "__" + safeId + "__t" + Math.round(tExport * 1000));
        var fBoxes = new File(TMP.fsName + "/ct_boxes_" + TS + "_" + bi + ".json");
        writeTextFile(fBoxes, jsonStringifyMini(boxes));
        boxesFilesToRemove.push(fBoxes);

        var suffix = TS + "_u" + uploadIdx;
        uploadIdx++;
        var one = curlPostMultipartBuild(
          EP_SCAN_FRAME,
          [
            { name: "projectId", value: STATE.projectId },
            { name: "fileKey", value: STATE.fileKey },
            { name: "t", value: "" + Math.round(best.t * 1000) },
            { name: "ssName", value: ssNameForLayer },
            { name: "ssWidth", value: "" + ssW },
            { name: "ssHeight", value: "" + ssH }
          ],
          [
            { name: "png", path: pngFile.fsName, mime: "image/png" },
            { name: "boxes", path: fBoxes.fsName, mime: "application/json" }
          ],
          suffix
        );
        uploadCommands.push(one);
      }

      // Pipeline: start uploads in background and continue to next frame.
      if (uploadCommands.length > 0) {
        for (var ui = 0; ui < uploadCommands.length; ui++) {
          if (tFirstUploadStart === null) tFirstUploadStart = (new Date()).getTime();
          runUploadInBackground(uploadCommands[ui], batFilesToRemove);
          allPendingHttpPaths.push(uploadCommands[ui].httpPath);
          allBackgroundCommands.push(uploadCommands[ui]);
        }
        pngFilesToRemove.push(pngFile);
        for (var br = 0; br < boxesFilesToRemove.length; br++) boxesFilesToRemoveAll.push(boxesFilesToRemove[br]);
      } else {
        try { if (pngFile.exists) pngFile.remove(); } catch (e0) {}
      }

      var exportSec = ((tExport1 - tExport0) / 1000).toFixed(1);
      totalExportMs += (tExport1 - tExport0);
      setStatus("Timeline Scan: Screenshot " + (gi + 1) + "/" + groupKeys.length + " (uploading in background)");
      if (setProgress) setProgress(gi, groupKeys.length, "Screenshot " + (gi + 1) + " of " + groupKeys.length + " | uploading…");
      try { app.refresh(); } catch (eRefresh) {}
    }

    } finally {
      if (compResBefore && compResBefore.length === 2) {
        try { comp.resolutionFactor = compResBefore; } catch(eR2){}
      }
    }

    if (allPendingHttpPaths.length > 0) {
      setStatus("Timeline Scan: Waiting for uploads…");
      var totalUploads = allPendingHttpPaths.length;
      var deadline = (new Date()).getTime() + 120000;
      while ((new Date()).getTime() < deadline) {
        var doneCount = 0;
        var allDone = true;
        for (var wi = 0; wi < allPendingHttpPaths.length; wi++) {
          var wf = new File(allPendingHttpPaths[wi]);
          if (wf.exists && wf.length > 0) {
            doneCount++;
          } else {
            allDone = false;
          }
        }
        if (setProgress) setProgress(doneCount, totalUploads, "Uploading images…");
        if (allDone) break;
        $.sleep(120);
        try { app.refresh(); } catch (eWLoop) {}
      }
      var waitResult = waitForBackgroundUploads(allPendingHttpPaths, null, { keepBodies: true });
      var httpCodes = waitResult.codes || waitResult;
      var responseBodies = waitResult.bodies || [];
      // Retry all 500s in one parallel batch (one short delay)
      if (allBackgroundCommands.length === httpCodes.length) {
        var retryCommands = [];
        var retryIndices = [];
        for (var ri = 0; ri < httpCodes.length; ri++) {
          if (httpCodes[ri] === "500") {
            retryCommands.push(allBackgroundCommands[ri]);
            retryIndices.push(ri);
          }
        }
        if (retryCommands.length > 0) {
          $.sleep(300);
          var retryRes = runParallelScanUploads(retryCommands, { keepBodies: true });
          var retryCodes = retryRes.codes || retryRes;
          var retryBodies = retryRes.bodies || [];
          for (var rj = 0; rj < retryIndices.length; rj++) {
            if (retryCodes[rj] === "200") {
              httpCodes[retryIndices[rj]] = "200";
              if (retryBodies[rj] != null) responseBodies[retryIndices[rj]] = retryBodies[rj];
            }
          }
        }
      }
      var tAllUploadsDone = (new Date()).getTime();
      if (tFirstUploadStart !== null) totalUploadMs = tAllUploadsDone - tFirstUploadStart;
      for (var h = 0; h < httpCodes.length; h++) {
        if (httpCodes[h] === "200") okCount++;
      }
      var responsesPath = Folder.myDocuments.fsName + "/Crowdin_scan_responses.json";
      try {
        var arr = [];
        for (var r = 0; r < httpCodes.length; r++) {
          arr.push({ index: r, httpCode: httpCodes[r], body: (responseBodies[r] != null) ? responseBodies[r] : "" });
        }
        writeTextFile(new File(responsesPath), jsonStringifyMini(arr));
      } catch (eSave) {}
      for (var b = 0; b < boxesFilesToRemoveAll.length; b++) {
        try { if (boxesFilesToRemoveAll[b].exists) boxesFilesToRemoveAll[b].remove(); } catch(eB){}
      }
      for (var p = 0; p < pngFilesToRemove.length; p++) {
        try { if (pngFilesToRemove[p].exists) pngFilesToRemove[p].remove(); } catch(eP){}
      }
      for (var bf = 0; bf < batFilesToRemove.length; bf++) {
        try { if (batFilesToRemove[bf].exists) batFilesToRemove[bf].remove(); } catch(eBf){}
      }
    }

    var totalExportS = (totalExportMs / 1000).toFixed(1);
    var totalUploadS = (totalUploadMs / 1000).toFixed(1);
    setStatus("Timeline Scan complete. " + okCount + " string(s), " + groupKeys.length + " frame(s). Responses: " + (Folder.myDocuments.fsName + "/Crowdin_scan_responses.json"));
    if (setProgress) setProgress(layerComps.length, layerComps.length, "Complete");
    return okCount > 0;
  } finally {
    if (scaleFactorApplied) { try { scaleCompositionByFactor(comp, 1 / scaleFactorApplied); } catch(eRestore){} }
  }
  }

  // Import translations: duplicate source comp, name as [comp name]_[language], apply translations to the new comp.
  function applyTranslationsToComp(comp, map, sourceCompId) {
    var updated = 0;
    for (var i = 1; i <= comp.numLayers; i++) {
      var L = comp.layer(i);
      var key = "comp_" + sourceCompId + "__layer_" + L.index;
      if (!(key in map)) continue;
      var sp = getSourceTextProp(L);
      if (!sp) continue;
      var doc = sp.value;
      doc.text = String(map[key] == null ? "" : map[key]);
      sp.setValue(doc);
      updated++;
    }
    return updated;
  }

  function importText(sourceComp, langId, setStatus, langDisplayName){
    langId = trim(langId||"");
    if (!langId) return false;
    if (!sourceComp || !(sourceComp instanceof CompItem)) return false;

    if (!STATE.projectId) { alertIf("Select a project first."); return false; }
    var fileKey = safeFileKeyForComp(sourceComp);

    setStatus("Importing (" + String(langId).toUpperCase() + ")…");

    var body = '{' +
      '"projectId":"' + jsonEscape(STATE.projectId) + '",' +
      '"fileKey":"' + jsonEscape(fileKey) + '",' +
      '"targetLang":"' + jsonEscape(langId) + '"' +
    '}';

    var r = curlPostJson(EP_PULL, body);
    if (r.http !== "200") {
      setStatus("Import failed.");
      alertIf("Pull failed.\nHTTP " + r.http + "\n\n" + (r.body||""));
      return false;
    }

    var items = parsePullItems(r.body);
    if (!items.length) {
      setStatus("Invalid pull response.");
      alertIf("Invalid pull response:\n\n" + (r.body||""));
      return false;
    }

    var map = {};
    for (var i=0;i<items.length;i++) map[items[i].id] = items[i].translatedText;

    var compNameBase = (sourceComp.name != null && String(sourceComp.name).replace(/^\s+|\s+$/g, "").length > 0)
      ? String(sourceComp.name).replace(/^\s+|\s+$/g, "")
      : "Comp";
    var langPart = (langDisplayName != null && String(langDisplayName).replace(/^\s+|\s+$/g, "").length > 0)
      ? String(langDisplayName).replace(/^\s+|\s+$/g, "")
      : langId;
    var nameForComp = compNameBase + "_" + langPart;

    app.beginUndoGroup("Crowdin Import: " + nameForComp);

    var newComp;
    try {
      newComp = sourceComp.duplicate();
    } catch (e) {
      app.endUndoGroup();
      setStatus("Duplicate failed.");
      alertIf("Could not duplicate composition.\n" + (e.message || e));
      return false;
    }

    newComp.name = nameForComp;
    var sourceCompId = sourceComp.id;
    var updated = applyTranslationsToComp(newComp, map, sourceCompId);

    app.endUndoGroup();

    setStatus("Imported " + updated + " layers into \"" + nameForComp + "\".");
    return true;
  }

  // UI (Collect triggers scan)
  // ScriptUI: uses native controls (Mac/Windows). helpTip on controls aids accessibility; labels before controls support screen readers. For localization, move UI strings to a single table.
  function buildUI(thisObj){
    var pal = (thisObj instanceof Panel)
      ? thisObj
      : new Window("palette", "Cult Connector - AE + Crowdin", undefined, {resizeable:true});

    // Fixed content width when run as Script UI panel (AE): avoids layout shift, dropdown stretch, margin changes on resize.
    var IS_PANEL = (thisObj instanceof Panel);
    var CONTENT_W = IS_PANEL ? 300 : -1;

    pal.orientation = "column";
    pal.alignChildren = IS_PANEL ? ["left", "top"] : ["fill", "fill"];
    pal.margins = [8, 0, 8, 8];
    pal.spacing = 4;
    if (IS_PANEL) {
      pal.preferredSize = [-1, -1];
    } else {
      pal.preferredSize = [320, 280];
    }
    if (pal instanceof Window) {
      pal.minimumSize = [280, 180];
    }
    // Resize: only resize(), not layout(true), to avoid resetting alignment/sizes (AE panel best practice).
    pal.onResizing = pal.onResize = function() {
      try { pal.layout.resize(); } catch (e) {}
    };

    // ---------- Onboarding (Connect → Choose project); hidden once main panel is active ----------
    // Use a stack so onboarding and main share one slot: content is centered, window has one content height.
    var contentStack = pal.add("group");
    contentStack.orientation = "stack";
    contentStack.alignChildren = ["fill", "fill"];
    contentStack.preferredSize = CONTENT_W > 0 ? [CONTENT_W, -1] : [-1, -1];

    var onboardingGroup = contentStack.add("group");
    onboardingGroup.orientation = "column";
    onboardingGroup.alignChildren = ["center", "center"];
    onboardingGroup.spacing = 0;
    onboardingGroup.preferredSize = CONTENT_W > 0 ? [CONTENT_W, 260] : [-1, 260];

    var onboardingPanel = onboardingGroup.add("panel", undefined, "Get started");
    onboardingPanel.orientation = "column";
    onboardingPanel.alignChildren = ["center", "center"];
    onboardingPanel.margins = [8, 8, 8, 8];
    onboardingPanel.spacing = 8;
    onboardingPanel.preferredSize = CONTENT_W > 0 ? [CONTENT_W, -1] : [340, -1];
    onboardingPanel.alignment = ["center", "center"];

    var stepStack = onboardingPanel.add("group");
    stepStack.orientation = "stack";
    stepStack.alignChildren = ["center", "center"];

    var step1Group = stepStack.add("group");
    step1Group.orientation = "column";
    step1Group.alignChildren = ["center", "center"];
    step1Group.spacing = 8;
    step1Group.alignment = ["center", "center"];
    var step1Hint = step1Group.add("statictext", undefined, "Sign in with your Crowdin account", {multiline:false});
    var btnConnectOnboard = step1Group.add("button", undefined, "Connect to Crowdin");
    btnConnectOnboard.preferredSize = [200, 30];

    var step2Group = stepStack.add("group");
    step2Group.orientation = "column";
    step2Group.alignChildren = ["center", "center"];
    step2Group.spacing = 8;
    step2Group.alignment = ["center", "center"];
    var step2Message = step2Group.add("statictext", undefined, "Choose a project to start", {multiline:false});
    var ddProjOnboard = step2Group.add("dropdownlist", undefined, []);
    ddProjOnboard.preferredSize = [320, 24];
    ddProjOnboard.helpTip = "Select the Crowdin project to use for this session.";
    step2Group.visible = false;

    // ---------- Main panel (Pages + Settings tabs); shown after onboarding complete ----------
    var mainGroup = contentStack.add("group");
    mainGroup.orientation = "column";
    mainGroup.alignChildren = ["fill", "fill"];
    mainGroup.preferredSize = CONTENT_W > 0 ? [CONTENT_W, -1] : [-1, -1];

    var tabs = mainGroup.add("tabbedpanel");
    tabs.alignChildren = ["fill", "fill"];
    tabs.preferredSize = CONTENT_W > 0 ? [CONTENT_W, -1] : [-1, -1];
    tabs.margins = [8, 0, 8, 0];

    var tabComposition = tabs.add("tab", undefined, "Composition");
    var tabSettings = tabs.add("tab", undefined, "Settings");

    function setStatus(s){
      if (typeof progressLabel !== "undefined") progressLabel.text = (s && s.length > 0) ? s : "Ready";
      try { pal.layout.resize(); } catch (e) {}
    }

    // (Legacy bottom progress bar removed in favor of popup progress.)
    var progressLabel = { text: "Ready" };

    function setProgress(current, total, message){
      // No-op: progress now shown in popup only.
    }

    tabComposition.orientation = "column";
    tabComposition.alignChildren = ["fill", "fill"];
    tabComposition.preferredSize = CONTENT_W > 0 ? [CONTENT_W, -1] : [-1, -1];
    tabComposition.margins = [8, 4, 8, 8];
    tabComposition.spacing = 6;

    // Snapshot Marker: small square only, top right below tab names. Single child + right alignment = no full-width spacer (per Adobe layout).
    var rowSnapshot = tabComposition.add("group");
    rowSnapshot.orientation = "row";
    rowSnapshot.alignChildren = ["right", "center"];
    rowSnapshot.alignment = ["fill", "top"];
    rowSnapshot.preferredSize = [-1, 22];
    rowSnapshot.minimumSize = [-1, 22];
    var btnSnapshotMarker = rowSnapshot.add("button", undefined, "\u25B2");
    btnSnapshotMarker.preferredSize = [22, 22];
    btnSnapshotMarker.minimumSize = [22, 22];
    btnSnapshotMarker.alignment = ["right", "center"];
    btnSnapshotMarker.helpTip = "Snapshot Marker - Select a text layer (works in main comp or precomp), set the playhead to the desired frame, then click to set as the preferred time for Smart Scan.";
    var snapshotMarkerMsg = "Snapshot Marker - Select a text layer, set the playhead to the desired frame on the timeline, then click to set as the preferred time for Smart Scan.";
    var snapshotMarkerWrongLayerMsg = "Select a text layer in this composition. Set the playhead to the desired frame, then click the Snapshot Marker button.";
    btnSnapshotMarker.onClick = function() {
      var comp = app.project && app.project.activeItem;
      if (!comp || !(comp instanceof CompItem)) {
        alertIf(snapshotMarkerMsg);
        return;
      }
      var sel = comp.selectedLayers;
      if (!sel || sel.length !== 1) {
        alertIf(snapshotMarkerMsg);
        return;
      }
      var layer = sel[0];
      if (layer.matchName !== "ADBE Text Layer") {
        alertIf(snapshotMarkerWrongLayerMsg);
        return;
      }
      var t = comp.time;
      if (setSnapshotMarkerAtTime(layer, t)) {
        if (typeof setStatus === "function") setStatus("Snapshot marker set at current time.");
        else alertIf("Snapshot marker set. Smart Scan will use this frame for this layer.");
      } else {
        alertIf("Could not add the marker.");
      }
    };

    var spacerBeforeExport = tabComposition.add("group");
    spacerBeforeExport.preferredSize = [-1, 1];

    // Action buttons stack: Export then Import.
    var rowMainBtns = tabComposition.add("group");
    rowMainBtns.orientation = "column";
    rowMainBtns.alignChildren = ["fill", "top"];
    rowMainBtns.alignment = ["fill", "top"];
    rowMainBtns.spacing = 12;
    var btnWidth = 220;
    var btnHeight = 26;
    var exportLabel = rowMainBtns.add("statictext", undefined, "After Effects \u2192 Crowdin");
    exportLabel.alignment = ["center", "top"];
    exportLabel.graphics = exportLabel.graphics || {};
    try { exportLabel.graphics.font = ScriptUI.newFont(exportLabel.graphics.font.name, ScriptUI.FontStyle.PLAIN, 11); } catch(e) {}

    var btnSendCompositions = rowMainBtns.add("button", undefined, "Send Selected Compositions");
    btnSendCompositions.alignment = ["fill", "top"];
    btnSendCompositions.minimumSize = [0, btnHeight];
    btnSendCompositions.preferredSize = [-1, btnHeight];
    btnSendCompositions.helpTip = "Export selected compositions to Crowdin for the chosen target language(s).";

    var spacerAfterSend = rowMainBtns.add("group");
    spacerAfterSend.preferredSize = [-1, 5];

    // Separator line centered between export and import sections.
    var sepLine = rowMainBtns.add("panel", undefined, "");
    sepLine.alignment = ["fill", "center"];
    sepLine.preferredSize = [-1, 2];
    sepLine.margins = [0, 0, 0, 0];

    var spacerAfterSep = rowMainBtns.add("group");
    spacerAfterSep.preferredSize = [-1, 5];

    // Update test marker: helps confirm the updater pulled the new build.
    var importLabel = rowMainBtns.add("statictext", undefined, "After Effects - Cult");
    importLabel.alignment = ["center", "top"];
    importLabel.graphics = importLabel.graphics || {};
    try { importLabel.graphics.font = ScriptUI.newFont(importLabel.graphics.font.name, ScriptUI.FontStyle.PLAIN, 11); } catch(e2) {}

    var btnImportSelected = rowMainBtns.add("button", undefined, "Import Selected");
    btnImportSelected.alignment = ["fill", "top"];
    btnImportSelected.minimumSize = [0, btnHeight];
    btnImportSelected.preferredSize = [-1, btnHeight];
    btnImportSelected.helpTip = "Create new compositions from Crowdin translations for the selected composition(s) and language(s). New comps are named [comp name]_[language].";

    var spacerAfterImport = rowMainBtns.add("group");
    spacerAfterImport.preferredSize = [-1, 12];

    // Languages section in its own panel. Fills remaining vertical space; fixed width when panel.
    var panelLang = tabComposition.add("panel", undefined, "Languages");
    panelLang.orientation = "column";
    panelLang.alignChildren = ["fill", "top"];
    panelLang.alignment = ["fill", "fill"];
    panelLang.preferredSize = CONTENT_W > 0 ? [CONTENT_W, -1] : [-1, -1];
    panelLang.margins = [8, 8, 8, 10];
    panelLang.spacing = 6;

    var langHeader = panelLang.add("group");
    langHeader.orientation = "row";
    langHeader.alignChildren = ["left", "center"];
    langHeader.margins = [8, 0, 0, 0];

    var langHeaderSpacer = langHeader.add("group");
    langHeaderSpacer.alignment = ["fill", "center"];
    var txtLangPage = langHeader.add("statictext", undefined, "", {multiline:false});
    txtLangPage.alignment = ["right", "center"];
    var btnLangPrev = langHeader.add("button", undefined, "Back");
    btnLangPrev.preferredSize = [36, 18];
    var btnLangNext = langHeader.add("button", undefined, "Next");
    btnLangNext.preferredSize = [36, 18];

    // "All languages" in its own row so each column below has max 4 languages.
    var rowLangAll = panelLang.add("group");
    rowLangAll.orientation = "row";
    rowLangAll.alignChildren = ["left", "center"];
    rowLangAll.margins = [8, 4, 0, 4];
    var cbLangAll = rowLangAll.add("checkbox", undefined, "All languages");
    cbLangAll.value = true;
    cbLangAll.helpTip = "When checked, all languages are used for export and import.";

    // Languages list: paged grid of checkboxes. 8 per page = 4 per column (no empty spaces).
    var langColumnsRow = panelLang.add("group");
    langColumnsRow.orientation = "row";
    langColumnsRow.alignChildren = ["fill", "top"];
    langColumnsRow.margins = [8, 0, 0, 0];

    var langColLeft = langColumnsRow.add("group");
    langColLeft.orientation = "column";
    langColLeft.alignChildren = ["left", "top"];
    langColLeft.margins = [4, 0, 0, 0];
    langColLeft.alignment = ["fill", "top"];

    var langColRight = langColumnsRow.add("group");
    langColRight.orientation = "column";
    langColRight.alignChildren = ["left", "top"];
    langColRight.margins = [0, 0, 0, 0];
    langColRight.alignment = ["fill", "top"];

    // Language checkboxes: 8 per page, 4 per column (left column up to 4, right column the rest).
    var LANGS_PER_PAGE = 8;
    var LANGS_PER_COLUMN = 4;
    var langPageIndex = 0; // 0-based
    if (!STATE.languageSelections) STATE.languageSelections = {};

    // For per-language checkboxes we add them directly into langColLeft / langColRight,
    // so they have the same focus ring behavior as "All languages".
    var langCheckGroup = langColLeft;   // left column list host
    var langRightGroup = langColRight;  // right column list host

    // Minimal gap above Readme/Check for updates so footer sits higher.
    var spacerBeforeCompFooter = tabComposition.add("group");
    spacerBeforeCompFooter.preferredSize = [-1, 0];

    // Composition footer: Readme and Check for updates (same placement as Settings).
    var compositionFooter = tabComposition.add("group");
    compositionFooter.orientation = "row";
    compositionFooter.alignChildren = ["left", "center"];
    compositionFooter.margins = [0, 0, 0, 0];
    compositionFooter.spacing = 6;

    var btnReadmeComp = compositionFooter.add("statictext", undefined, "Readme");
    btnReadmeComp.alignment = ["left", "center"];
    btnReadmeComp.helpTip = "Open plugin readme and data security info.";
    var compFooterSpacer = compositionFooter.add("group");
    compFooterSpacer.alignment = ["fill", "center"];
    var btnCheckUpdatesComp = compositionFooter.add("statictext", undefined, "Check for updates");
    btnCheckUpdatesComp.alignment = ["right", "center"];
    btnCheckUpdatesComp.helpTip = "Check for and install the latest version from GitHub Releases (restart required).";
    try {
      var linkBlueComp = btnReadmeComp.graphics.newPen(btnReadmeComp.graphics.PenType.SOLID_COLOR, [0.2, 0.55, 1, 1], 1);
      btnReadmeComp.graphics.foregroundColor = linkBlueComp;
      btnCheckUpdatesComp.graphics.foregroundColor = linkBlueComp;
    } catch (eLinkComp) {}

    tabSettings.orientation = "column";
    tabSettings.alignChildren = ["fill", "top"];
    tabSettings.margins = [8, 10, 8, 8];
    tabSettings.spacing = 6;

    var panelProj = tabSettings.add("panel", undefined, "Crowdin Project");
    panelProj.orientation = "column";
    panelProj.alignChildren = ["fill", "top"];
    panelProj.margins = 8;
    panelProj.spacing = 8;

    var rowProj = panelProj.add("group");
    rowProj.orientation = "row";
    rowProj.alignChildren = ["fill", "center"];
    rowProj.spacing = 8;
    rowProj.margins = [0, 8, 0, 0];

    var ddProj = rowProj.add("dropdownlist", undefined, []);
    // Fixed width when panel: prevents dropdown list from stretching/shrinking on resize (AE best practice).
    ddProj.preferredSize = CONTENT_W > 0 ? [CONTENT_W - 24, 24] : [-1, 24];
    ddProj.minimumSize = [180, 24];
    ddProj.alignment = ["left", "center"];
    ddProj.helpTip = "Current Crowdin project. Change to load a different project's languages.";

    var rowOpenCrowdin = panelProj.add("group");
    rowOpenCrowdin.orientation="row";
    rowOpenCrowdin.alignChildren=["left","center"];
    rowOpenCrowdin.spacing=6;

    // Blue text-style button for Open in Crowdin (clickable).
    var btnOpenCrowdin = rowOpenCrowdin.add("statictext", undefined, "Open in Crowdin");
    btnOpenCrowdin.alignment = ["left", "center"];
    try {
      var crowdinBlue = btnOpenCrowdin.graphics.newPen(btnOpenCrowdin.graphics.PenType.SOLID_COLOR, [0.2, 0.55, 1, 1], 1);
      btnOpenCrowdin.graphics.foregroundColor = crowdinBlue;
    } catch (eOC) {}

    var spacerCrowdin = rowOpenCrowdin.add("group");
    spacerCrowdin.alignment = ["fill", "center"];
    spacerCrowdin.preferredSize = [-1, 18];

    // Compact actions to the right of "Open in Crowdin".
    var btnRefreshProj = rowOpenCrowdin.add("button", undefined, "Refresh");
    btnRefreshProj.preferredSize = [60, 18];
    btnRefreshProj.helpTip = "Reload projects and languages from Crowdin.";

    var btnDisconnect = rowOpenCrowdin.add("button", undefined, "Disconnect");
    btnDisconnect.preferredSize=[70,18];

    var rowSegmentation = panelProj.add("group");
    rowSegmentation.orientation="row";
    rowSegmentation.alignChildren=["left","center"];
    rowSegmentation.spacing=6;
    rowSegmentation.margins = [4, 0, 0, 0];
    var cbSegmentation = rowSegmentation.add("checkbox", undefined, "Content Segmentation");
    cbSegmentation.alignment = ["left", "center"];
    // Default: segmentation on; user can uncheck to minimize segmentation.
    cbSegmentation.value = true;
    STATE.useSegmentation = true;
    cbSegmentation.onClick = function () {
      STATE.useSegmentation = (cbSegmentation.value === true);
    };

    // ----- Settings footer: blue text buttons (Readme, Check for updates) -----
    var settingsFooter = tabSettings.add("group");
    settingsFooter.orientation = "row";
    settingsFooter.alignChildren = ["left", "center"];
    settingsFooter.margins = [0, 4, 0, 0];
    settingsFooter.spacing = 6;

    var btnReadme = settingsFooter.add("statictext", undefined, "Readme");
    btnReadme.alignment = ["left", "center"];
    btnReadme.helpTip = "Open plugin readme and data security info.";

    var footerSpacer = settingsFooter.add("group");
    footerSpacer.alignment = ["fill", "center"];

    var btnCheckUpdates = settingsFooter.add("statictext", undefined, "Check for updates");
    btnCheckUpdates.alignment = ["right", "center"];
    btnCheckUpdates.helpTip = "Check for and install the latest version from GitHub Releases (restart required).";

    // Style footer links as blue text to match AE link-style labels.
    try {
      var linkBlue = btnReadme.graphics.newPen(btnReadme.graphics.PenType.SOLID_COLOR, [0.2, 0.55, 1, 1], 1);
      btnReadme.graphics.foregroundColor = linkBlue;
      btnCheckUpdates.graphics.foregroundColor = linkBlue;
    } catch (eLink) {}

    function showReadmeDialog(){
      try {
        var dlg = new Window("dialog", "Cult Connector - AE + Crowdin");
        dlg.orientation = "column";
        dlg.alignChildren = ["fill", "top"];
        dlg.margins = [12, 12, 12, 8];
        dlg.spacing = 6;

        var msg = dlg.add("statictext", undefined,
          "Cult Connector - AE + Crowdin\n" +
          "Version 1.0.0" + PLUGIN_VERSION + "\n\n" +
          "Data security:\n\n" +
          "• Text and screenshots are sent securely over HTTPS to Crowdin.\n" +
          "• Translations are stored and managed by your Crowdin project.\n" +
          "• Cult Connecto does not store translations.\n\n" +
          "All trademarks and brands are the property of Cult Extensions.\n" +
          "© 2026 Cult Extensions. All rights reserved.\n\n" +
          "Support: contact@cultextensions.com",
          { multiline: true }
        );
        msg.alignment = ["fill", "top"];
        msg.preferredSize = [320, -1];

        var rowBtns = dlg.add("group");
        rowBtns.orientation = "row";
        rowBtns.alignChildren = ["right", "center"];
        var ok = rowBtns.add("button", undefined, "OK", { name: "ok" });
        ok.alignment = ["right", "center"];

        dlg.minimumSize = [340, 160];
        dlg.layout.layout(true);
        dlg.center();
        dlg.show();
      } catch (e) {
        alertIf("Readme\n\nCult Connecto - AE + Crowdin\nVersion " + PLUGIN_VERSION + "\n\nData security:\n" +
                "- Text and screenshots are sent securely over HTTPS to Crowdin.\n" +
                "- Translations are stored and managed by your Crowdin project.\n" +
                "- Cult Connecto does not store translations.\n\n" +
                "All trademarks and brands are the property of Cult Extensions.\n" +
                "© 2026 Cult Extensions. All rights reserved.\n\n" +
                "Support: contact@cultextensions.com");
      }
    }

    // Simple centered popup progress bar for AE → Crowdin operations.
    function showAeCrowdinProgressPopup() {
      try {
        var w = new Window("palette", "After Effects \u2192 Crowdin");
        w.orientation = "column";
        w.alignChildren = ["fill", "top"];
        w.margins = 12;
        var title = w.add("statictext", undefined, "After Effects \u2192 Crowdin", { multiline: false });
        title.alignment = ["center", "top"];
        var statusLabel = w.add("statictext", undefined, "Preparing…", { multiline: false });
        statusLabel.preferredSize = [380, 22];
        statusLabel.alignment = ["fill", "top"];
        var bar = w.add("progressbar", undefined, 0, 100);
        bar.preferredSize = [380, 12];
        bar.value = 0;
        try { app.refresh(); } catch (eR0) {}
        w.layout.layout(true);
        w.center();
        w.show();
        return { win: w, bar: bar, statusLabel: statusLabel };
      } catch (e) {
        return null;
      }
    }

    // Centered popup progress bar for Crowdin → AE imports.
    function showCrowdinAeProgressPopup() {
      try {
        var w = new Window("palette", "Crowdin \u2192 After Effects");
        w.orientation = "column";
        w.alignChildren = ["fill", "top"];
        w.margins = 12;
        var title = w.add("statictext", undefined, "Crowdin \u2192 After Effects", { multiline: false });
        title.alignment = ["center", "top"];
        var statusLabel = w.add("statictext", undefined, "Preparing…", { multiline: false });
        statusLabel.preferredSize = [380, 22];
        statusLabel.alignment = ["fill", "top"];
        var bar = w.add("progressbar", undefined, 0, 100);
        bar.preferredSize = [380, 12];
        bar.value = 0;
        try { app.refresh(); } catch (eR1) {}
        w.layout.layout(true);
        w.center();
        w.show();
        return { win: w, bar: bar, statusLabel: statusLabel };
      } catch (e) {
        return null;
      }
    }

    function updateLangPager(){
      var L = STATE.languages || [];
      var total = L.length;
      if (total <= 0) {
        txtLangPage.text = "No languages.";
        btnLangPrev.enabled = false;
        btnLangNext.enabled = false;
        btnLangPrev.visible = true;
        txtLangPage.visible = true;
        btnLangNext.visible = true;
        return;
      }
      var totalPages = Math.max(1, Math.ceil(total / LANGS_PER_PAGE));
      if (langPageIndex < 0) langPageIndex = 0;
      if (langPageIndex > totalPages - 1) langPageIndex = totalPages - 1;
      var start = langPageIndex * LANGS_PER_PAGE;
      var end = Math.min(total, start + LANGS_PER_PAGE);
      txtLangPage.text = (start + 1) + "–" + end + " of " + total;
      btnLangPrev.enabled = (langPageIndex > 0);
      btnLangNext.enabled = (langPageIndex < totalPages - 1);
      var showPager = (totalPages > 1);
      btnLangPrev.visible = showPager;
      txtLangPage.visible = showPager;
      btnLangNext.visible = showPager;
    }

    function fillProjects(ps){
      ddProj.removeAll();
      for (var i = 0; i < ps.length; i++) {
        var it = ddProj.add("item", ps[i].name);
        it._project = ps[i];
      }
      if (ddProj.items.length) {
        var targetIdx = 0;
        if (STATE.projectId && ps && ps.length) {
          for (var i = 0; i < ps.length; i++) {
            if (String(ps[i].id) === String(STATE.projectId)) { targetIdx = i; break; }
          }
        }
        ddProj.selection = targetIdx;
      }
      pal.layout.layout(true);
    }

    function fillLangs(langs){
      // Always prefer freshly loaded languages when provided; otherwise fall back to existing STATE.languages.
      if (langs && langs.length) {
        STATE.languages = langs;
        // Reset page and clear previous selections when a new project loads.
        langPageIndex = 0;
        STATE.languageSelections = {};
      } else if (!STATE.languages) {
        STATE.languages = [];
      }
      if (typeof langCheckGroup !== "undefined") {
        // Clear existing language checkboxes in both columns (All languages is in rowLangAll).
        while (langCheckGroup.children.length > 0) langCheckGroup.remove(langCheckGroup.children[0]);
        if (typeof langRightGroup !== "undefined") {
          while (langRightGroup.children.length > 0) langRightGroup.remove(langRightGroup.children[0]);
        }
        var L = STATE.languages || [];
        if (!L.length) {
          var emptyMsg = langCheckGroup.add("statictext", undefined, "No languages loaded. Refresh the project.", {multiline:false});
          emptyMsg.enabled = false;
          pal.layout.layout(true);
          updateLangPager();
          return;
        }
        var total = L.length;
        var totalPages = Math.max(1, Math.ceil(total / LANGS_PER_PAGE));
        if (langPageIndex < 0) langPageIndex = 0;
        if (langPageIndex > totalPages - 1) langPageIndex = totalPages - 1;
        var start = langPageIndex * LANGS_PER_PAGE;
        var end = Math.min(total, start + LANGS_PER_PAGE);
        var defaultChecked = cbLangAll && cbLangAll.value === true;

        // Split current page: left column up to 4, right column the rest (no empty slots).
        var count = end - start;
        var mid = Math.min(LANGS_PER_COLUMN, count);
        for (var offset = 0; offset < count; offset++) {
          var j = start + offset;
          var lang = L[j];
          var targetCol = (offset < mid) ? langCheckGroup : langRightGroup;
          var baseLabel = (lang.name || lang.id) + " (" + String(lang.id || "").toUpperCase() + ")";
          var cb = targetCol.add("checkbox", undefined, baseLabel);
          cb._langId = lang.id;
          cb._langName = (lang.name != null && String(lang.name).length > 0) ? String(lang.name).replace(/^\s+|\s+$/g, "") : (lang.id || "");
          // Use stored selection if available; otherwise inherit from All checkbox.
          var stored = (STATE.languageSelections && STATE.languageSelections.hasOwnProperty(lang.id)) ? STATE.languageSelections[lang.id] : null;
          cb.value = (stored !== null) ? stored : defaultChecked;
          cb.onClick = function(){
            if (!this._langId) return;
            if (!STATE.languageSelections) STATE.languageSelections = {};
            STATE.languageSelections[this._langId] = (this.value === true);
            if (!cbLangAll) return;
            var allOn = true;
            var list = STATE.languages || [];
            if (!list.length) allOn = false;
            for (var i = 0; i < list.length; i++) {
              var lid = list[i].id;
              var val;
              if (STATE.languageSelections && STATE.languageSelections.hasOwnProperty(lid)) {
                val = STATE.languageSelections[lid];
              } else {
                val = defaultChecked;
              }
              if (!val) { allOn = false; break; }
            }
            cbLangAll.value = allOn;
          };
        }
        pal.layout.layout(true);
        updateLangPager();
      }
      pal.layout.layout(true);
    }

    // Keep \"All languages\" checkbox in sync with all language selections when toggled directly.
    cbLangAll.onClick = function() {
      var L = STATE.languages || [];
      if (!STATE.languageSelections) STATE.languageSelections = {};
      var v = cbLangAll.value === true;
      for (var i = 0; i < L.length; i++) {
        STATE.languageSelections[L[i].id] = v;
      }
      fillLangs(null); // re-render current page with updated checkboxes
    };

    // Footer link behavior
    btnReadme.addEventListener("click", function () {
      showReadmeDialog();
    });
    btnReadmeComp.addEventListener("click", function () {
      showReadmeDialog();
    });
    // Make "Open in Crowdin" blue label clickable as well.
    btnOpenCrowdin.addEventListener("click", function () {
      if (!STATE.projectId) { alertIf("No project selected."); return; }
      var name = (STATE.projectName || "").toString().replace(/^\s+|\s+$/g, "");
      if (name.length === 0) {
        alertIf("No project name.");
        return;
      }
      var slug = name.toLowerCase().replace(/\s+/g, "-").replace(/[^a-z0-9-]/g, "");
      if (slug.length === 0) slug = String(STATE.projectId);
      var url = "https://crowdin.com/project/" + encodeURIComponent(slug);
      openUrl(url);
    });
    btnCheckUpdates.addEventListener("click", function () {
      runUpdateCheck(setStatus);
    });
    btnCheckUpdatesComp.addEventListener("click", function () {
      runUpdateCheck(setStatus);
    });

    btnDisconnect.onClick = function(){
      STATE.projectId = null;
      STATE.projectName = null;
      STATE.projects = null;
      STATE.languages = null;
      STATE.connected = false;
      onboardingGroup.visible = true;
      mainGroup.visible = false;
      step1Group.visible = true;
      step2Group.visible = false;
      ddProjOnboard.removeAll();
      contentStack.preferredSize = CONTENT_W > 0 ? [CONTENT_W, 170] : [-1, 170];
      onboardingGroup.preferredSize = CONTENT_W > 0 ? [CONTENT_W, 170] : [-1, 170];
      if (typeof progressLabel !== "undefined") progressLabel.text = "Disconnected. Connect to get started.";
      pal.layout.layout(true);
      if (pal.layout.resize) pal.layout.resize();
      if (pal instanceof Window) pal.size = [560, 250];
    };

    btnOpenCrowdin.onClick = function(){
      if (!STATE.projectId) { alertIf("No project selected."); return; }
      var name = (STATE.projectName || "").toString().replace(/^\s+|\s+$/g, "");
      if (name.length === 0) {
        alertIf("No project name.");
        return;
      }
      var slug = name.toLowerCase().replace(/\s+/g, "-").replace(/[^a-z0-9-]/g, "");
      if (slug.length === 0) slug = String(STATE.projectId);
      var url = "https://crowdin.com/project/" + encodeURIComponent(slug);
      var cmd = IS_WIN ? 'start "" "' + url + '"' : 'open "' + url.replace(/"/g, '\\"') + '"';
      try { system.callSystem(cmd); } catch(e) { alertIf("Could not open browser: " + (e.message || e)); }
    };

    ddProj.onChange = function(){
      var sel = ddProj.selection;
      var p = (sel && sel._project) ? sel._project : null;
      if (!p) return;
      selectProject(p.id, p.name, setStatus);
      fillLangs(loadLanguages(setStatus));
    };

    btnRefreshProj.onClick = function(){
      if (!STATE.connected) { alertIf("Connect first."); return; }
      setStatus("Reloading projects…");
      STATE.projects = loadProjects(setStatus);
      if (STATE.projects && STATE.projects.length) fillProjects(STATE.projects);
      if (STATE.projectId) {
        setStatus("Reloading languages…");
        fillLangs(loadLanguages(setStatus));
      }
      setStatus("Projects and languages refreshed.");
    };

    btnLangPrev.onClick = function() {
      if (!STATE.languages || !STATE.languages.length) return;
      if (langPageIndex <= 0) return;
      langPageIndex--;
      fillLangs(null);
    };

    btnLangNext.onClick = function() {
      var L = STATE.languages || [];
      if (!L.length) return;
      var totalPages = Math.max(1, Math.ceil(L.length / LANGS_PER_PAGE));
      if (langPageIndex >= totalPages - 1) return;
      langPageIndex++;
      fillLangs(null);
    };

    btnSendCompositions.onClick = function(){
      if (!STATE.connected) { alertIf("Connect first."); return; }
      if (!STATE.projectId) { alertIf("Select a project first."); return; }
      var uploadTargets = [];
      if (cbLangAll && cbLangAll.value === true) {
        uploadTargets.push("all");
      } else {
        var L = STATE.languages || [];
        if (L.length && STATE.languageSelections) {
          for (var i = 0; i < L.length; i++) {
            var id = L[i].id;
            if (STATE.languageSelections[id] === true) uploadTargets.push(id);
          }
        }
      }
      if (!uploadTargets.length) {
        alertIf("Select at least one language.");
        return;
      }
      STATE.compsToSend = [];
      var popup = showAeCrowdinProgressPopup();
      function popupStatus(msg) {
        if (popup && popup.statusLabel) popup.statusLabel.text = msg || "";
      }
      function popupSetProgress(current, total, message) {
        if (!popup || !popup.bar) return;
        var max = (total && total > 0) ? total : 1;
        var frac = current / max;
        if (frac < 0) frac = 0;
        if (frac > 1) frac = 1;
        var base = 60;
        var end = 100;
        popup.bar.value = Math.round(base + (end - base) * frac);
        try { app.refresh(); } catch (eR2) {}
      }
      function statusProxy(msg) {
        popupStatus(msg);
        setStatus(msg);
      }
      popupStatus("Preparing…");
      var comps = getSelectedComps();
      if (!comps || comps.length === 0) {
        alertIf("Select one or more compositions in the Project panel, or open a composition in the timeline.");
        if (popup && popup.win) try { popup.win.close(); } catch(eC0) {}
        pal.layout.layout(true);
        STATE.compsToSend = [];
        return;
      }
      STATE.useSegmentation = (typeof cbSegmentation !== "undefined" && cbSegmentation.value === true);
      var allOk = true;
      for (var t = 0; t < uploadTargets.length; t++) {
        var currentTarget = uploadTargets[t];
        if (popup && popup.bar) {
          var fracUpload = uploadTargets.length ? (t / uploadTargets.length) : 0;
          var startU = 20;
          var endU = 60;
          if (fracUpload < 0) fracUpload = 0;
          if (fracUpload > 1) fracUpload = 1;
          popup.bar.value = Math.round(startU + (endU - startU) * fracUpload);
          try { app.refresh(); } catch (eR3) {}
        }
        popupStatus("Uploading: " + (currentTarget === "all" ? "All Languages" : String(currentTarget).toUpperCase()));
        for (var c = 0; c < comps.length; c++) {
          var comp = comps[c];
          var compName = (comp.name != null) ? String(comp.name) : ("comp " + (c + 1));
          if (comps.length > 1) popupStatus("Uploading: " + (currentTarget === "all" ? "All Languages" : String(currentTarget).toUpperCase()) + " — " + compName);
          var items = collectText(statusProxy, comp);
          if (!items || !items.length) continue;
          var ok = uploadStrings(items, statusProxy, currentTarget);
          if (!ok) { allOk = false; break; }
        }
        if (!allOk) break;
      }
      if (!allOk) {
        if (popup && popup.win) try { popup.win.close(); } catch(eC2) {}
        pal.layout.layout(true);
        STATE.compsToSend = [];
        return;
      }
      statusProxy("Waiting for Crowdin…");
      for (var w = 0; w < 6; w++) {
        if (w > 0) $.sleep(80);
        try { app.refresh(); } catch (eR) {}
      }
      popupStatus("Scanning timeline…");
      for (var sc = 0; sc < comps.length; sc++) {
        try { app.project.activeItem = comps[sc]; } catch (eAct) {}
        STATE.fileKey = safeFileKeyForComp(comps[sc]);
        if (comps.length > 1) popupStatus("Scanning timeline: " + (comps[sc].name || STATE.fileKey) + "…");
        smartScanTimeline(statusProxy, popupSetProgress, comps[sc]);
      }
      if (popup && popup.bar) {
        popup.bar.value = 100;
        try { app.refresh(); } catch (eR4) {}
      }
      if (popup && popup.win) try { popup.win.close(); } catch(eC3) {}
      pal.layout.layout(true);
      STATE.compsToSend = [];
    };

    btnImportSelected.onClick = function(){
      if (!STATE.connected) { alertIf("Connect first."); return; }
      if (!STATE.projectId) { alertIf("Select a project first."); return; }
      var comps = getSelectedComps();
      if (!comps || comps.length === 0) {
        alertIf("Select at least one composition (Composition tab checkboxes, or comps in Project panel).");
        return;
      }
      var selected = [];
      if (cbLangAll && cbLangAll.value === true && STATE.languages && STATE.languages.length) {
        for (var k = 0; k < STATE.languages.length; k++) {
          var lang = STATE.languages[k];
          selected.push({ id: lang.id, name: (lang.name != null && String(lang.name).length > 0) ? String(lang.name).replace(/^\s+|\s+$/g, "") : lang.id });
        }
      } else if (STATE.languages && STATE.languages.length && STATE.languageSelections) {
        var L = STATE.languages;
        for (var i = 0; i < L.length; i++) {
          var id = L[i].id;
          if (STATE.languageSelections[id] === true) {
            var name = (L[i].name != null && String(L[i].name).length > 0) ? String(L[i].name).replace(/^\s+|\s+$/g, "") : id;
            selected.push({ id: id, name: name });
          }
        }
      }
      if (selected.length === 0) {
        alertIf("Select at least one language.");
        return;
      }
      var popup = showCrowdinAeProgressPopup();
      function popupStatus(msg) {
        if (popup && popup.statusLabel) popup.statusLabel.text = msg || "";
      }
      function popupSetProgress(current, total) {
        if (!popup || !popup.bar) return;
        var max = (total && total > 0) ? total : 1;
        var frac = current / max;
        if (frac < 0) frac = 0;
        if (frac > 1) frac = 1;
        popup.bar.value = Math.round(100 * frac);
        try { app.refresh(); } catch (eR5) {}
      }
      function statusProxy(msg) {
        popupStatus(msg);
        setStatus(msg);
      }
      var totalWork = comps.length * selected.length;
      var done = 0;
      for (var c = 0; c < comps.length; c++) {
        var comp = comps[c];
        var compName = (comp.name != null) ? String(comp.name) : ("comp " + (c + 1));
        for (var d = 0; d < selected.length; d++) {
          var displayName = selected[d].name || selected[d].id;
          statusProxy("Importing " + compName + " \u2192 " + displayName + " (" + String(selected[d].id).toUpperCase() + ")…");
          importText(comp, selected[d].id, statusProxy, displayName);
          done++;
          popupSetProgress(done, totalWork);
        }
      }
      statusProxy("Import complete. " + done + " composition(s) created.");
      if (popup && popup.bar) {
        popup.bar.value = 100;
        try { app.refresh(); } catch (eR6) {}
      }
      if (popup && popup.win) try { popup.win.close(); } catch(eC4) {}
    };

    // ---------- Onboarding: fill project dropdown and show step 2 ----------
    function fillProjectsOnboard(ps) {
      ddProjOnboard.removeAll();
      var placeholder = ddProjOnboard.add("item", "— Choose project —");
      placeholder._project = null;
      for (var i = 0; i < ps.length; i++) {
        var it = ddProjOnboard.add("item", ps[i].name);
        it._project = ps[i];
      }
      ddProjOnboard.selection = 0;
      step1Group.visible = false;
      step2Group.visible = true;
      pal.layout.layout(true);
    }

    btnConnectOnboard.onClick = function() {
      var ok = oauthConnect(setStatus);
      if (!ok) return;
      STATE.projects = loadProjects(setStatus);
      if (!STATE.projects || !STATE.projects.length) {
        setStatus("No projects found.");
        return;
      }
      fillProjectsOnboard(STATE.projects);
      setStatus("Choose a project.");
    };

    ddProjOnboard.onChange = function() {
      var sel = ddProjOnboard.selection;
      var p = (sel && sel._project) ? sel._project : null;
      if (!p) return;
      setStatus("Loading project…");
      selectProject(p.id, p.name, setStatus);
      fillLangs(loadLanguages(setStatus));
      onboardingGroup.visible = false;
      mainGroup.visible = true;
      tabs.selection = 0;
      contentStack.preferredSize = CONTENT_W > 0 ? [CONTENT_W, 260] : [-1, 260];
      onboardingGroup.preferredSize = CONTENT_W > 0 ? [CONTENT_W, 260] : [-1, 260];
      pal.layout.layout(true);
      if (pal.layout.resize) pal.layout.resize();
      if (pal instanceof Window) pal.size = [560, 290];
      fillProjects(STATE.projects);
      fillLangs(STATE.languages);
      setStatus("Ready.");
      pal.layout.layout(true);
    };

    // ---------- Initial view: onboarding vs main ----------
    if (STATE.connected && STATE.projectId) {
      onboardingGroup.visible = false;
      mainGroup.visible = true;
      tabs.selection = 0;
      contentStack.preferredSize = CONTENT_W > 0 ? [CONTENT_W, 260] : [-1, 260];
      onboardingGroup.preferredSize = CONTENT_W > 0 ? [CONTENT_W, 260] : [-1, 260];
      STATE.projects = loadProjects(setStatus);
      if (STATE.projects && STATE.projects.length) fillProjects(STATE.projects);
      fillLangs(loadLanguages(setStatus));
      setStatus("Ready.");
      pal.layout.layout(true);
      if (pal.layout.resize) pal.layout.resize();
      if (pal instanceof Window) pal.size = [560, 290];
    } else {
      onboardingGroup.visible = true;
      mainGroup.visible = false;
      setStatus("Connect to get started.");
      contentStack.preferredSize = CONTENT_W > 0 ? [CONTENT_W, 170] : [-1, 170];
      onboardingGroup.preferredSize = CONTENT_W > 0 ? [CONTENT_W, 170] : [-1, 170];
      if (pal instanceof Window) pal.size = [560, 250];
    }

    // Force layout so onboarding content is centered on first paint (before show).
    pal.layout.layout(true);
    if (pal.layout.resize) pal.layout.resize();
    if (onboardingGroup.visible && onboardingPanel.layout) {
      onboardingPanel.layout.layout(true);
      if (stepStack.layout) stepStack.layout.layout(true);
    }

    if (pal instanceof Window) { pal.center(); pal.show(); }
    return pal;
  }

  buildUI(thisObj);

})(this);