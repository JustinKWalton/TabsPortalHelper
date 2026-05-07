#!/usr/bin/env node
/**
 * TABS Portal helper — content-stream cropper
 *
 * Standalone Node port of the v22-v30 inline cropper that used to run inside
 * the generate-markup-summary edge function. Moved here so the heavy CPU work
 * happens on the desktop, where there's no 2-second budget.
 *
 * Input:  a marked-up plan PDF + the user_id and plan_pdf_hash to key output.
 * Output: a folder of self-contained 1-page PDFs (one per markup) + manifest.json
 *         that the edge function reads at report-generation time.
 *
 * USAGE
 * -----
 *   node cropper.js <plan-pdf-path> <output-dir> [--max-decoded-bytes N]
 *
 * EXAMPLE
 * -------
 *   node cropper.js ./plans/2026004168.pdf ./out/_cropped/USERID/HASH/
 *
 * The output directory will contain:
 *   - manifest.json
 *   - markup_<nm>.pdf  (one per Polygon annotation with /BSIColumnData)
 *
 * The C# tray helper invokes this script after every plan upload and uploads
 * the output folder to:
 *   tdlr-reports/_cropped/{user_id}/{plan_pdf_hash}/...
 *
 * EXIT CODES
 * ----------
 *   0  ok (manifest written, may include zero markups if none found)
 *   1  invalid args / I/O error
 *   2  PDF parse error (file is not a valid PDF or unsupported encryption)
 *   3  unrecoverable crop error (logged in stderr; partial manifest still written)
 */

'use strict';

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const {
  PDFDocument,
  PDFName,
  PDFArray,
  PDFDict,
  PDFNumber,
  PDFString,
  PDFHexString,
  PDFRef,
  PDFStream,
  PDFRawStream,
} = require('pdf-lib');

const pako = require('pako');

// -----------------------------------------------------------------------------
// Constants — same memory bounds as the original v24, but the desktop has lots
// more headroom so we can crank these up without risking an OOM. Kept generous
// rather than tight: a 100 MB plan with 100 MB inflated content is fine on a
// modern dev machine, and we'd rather succeed than fall through.
// -----------------------------------------------------------------------------
const MAX_DECODED_FOR_V24 = 200_000_000;       // 200 MB inflated cap (was 2.5 MB in edge)
const MAX_TOKENS_FOR_V24 = 100_000_000;        // 100 M tokens cap
const THUMB_PAD_PT = 120;                      // crop padding around markup vertices/rect

// Token type constants (same as edge function)
const TT_NUM = 1;
const TT_NAME = 2;
const TT_OP = 3;
const TT_STRING = 4;
const TT_HEX = 5;
const TT_ARRAY = 6;
const TT_DICT = 7;
const TT_INLINE_IMAGE = 8;

const OP_UNKNOWN = 0;
const OP_q = 1; const OP_Q = 2; const OP_cm = 3;
const OP_m = 4; const OP_l = 5; const OP_c = 6; const OP_v = 7; const OP_y = 8; const OP_h = 9; const OP_re = 10;
const OP_S = 11; const OP_s = 12; const OP_f = 13; const OP_F = 14; const OP_f_star = 15;
const OP_B = 16; const OP_B_star = 17; const OP_b = 18; const OP_b_star = 19; const OP_n = 20;
const OP_w = 21; const OP_J = 22; const OP_j = 23; const OP_M = 24; const OP_d = 25;
const OP_ri = 26; const OP_i = 27; const OP_gs = 28;
const OP_cs = 29; const OP_CS = 30; const OP_sc = 31; const OP_SC = 32; const OP_scn = 33; const OP_SCN = 34;
const OP_g = 35; const OP_G = 36; const OP_rg = 37; const OP_RG = 38; const OP_k = 39; const OP_K = 40;

const OP_MAP = new Map([
  ['q', OP_q], ['Q', OP_Q], ['cm', OP_cm],
  ['m', OP_m], ['l', OP_l], ['c', OP_c], ['v', OP_v], ['y', OP_y], ['h', OP_h], ['re', OP_re],
  ['S', OP_S], ['s', OP_s], ['f', OP_f], ['F', OP_F], ['f*', OP_f_star],
  ['B', OP_B], ['B*', OP_B_star], ['b', OP_b], ['b*', OP_b_star], ['n', OP_n],
  ['w', OP_w], ['J', OP_J], ['j', OP_j], ['M', OP_M], ['d', OP_d],
  ['ri', OP_ri], ['i', OP_i], ['gs', OP_gs],
  ['cs', OP_cs], ['CS', OP_CS], ['sc', OP_sc], ['SC', OP_SC], ['scn', OP_scn], ['SCN', OP_SCN],
  ['g', OP_g], ['G', OP_G], ['rg', OP_rg], ['RG', OP_RG], ['k', OP_k], ['K', OP_K],
]);

// -----------------------------------------------------------------------------
// Logging helpers
// -----------------------------------------------------------------------------
const t0 = Date.now();
function log(label, value) {
  const ms = Date.now() - t0;
  if (value === undefined) {
    process.stderr.write(`[+${ms}ms] ${label}\n`);
  } else {
    const v = typeof value === 'string' ? value : JSON.stringify(value);
    process.stderr.write(`[+${ms}ms] ${label}: ${v}\n`);
  }
}

// -----------------------------------------------------------------------------
// PDF reading helpers (mirrored from edge function)
// -----------------------------------------------------------------------------
function decodePdfString(v) {
  if (v instanceof PDFString || v instanceof PDFHexString) return v.decodeText();
  return String(v ?? '');
}

function getDictStr(d, name) {
  const ref = d.get(PDFName.of(name));
  if (!ref) return '';
  const resolved = d.context.lookup(ref);
  return decodePdfString(resolved);
}

function getDictArray(d, name) {
  const ref = d.get(PDFName.of(name));
  if (!ref) return null;
  const resolved = d.context.lookup(ref);
  if (resolved instanceof PDFArray) return resolved;
  return null;
}

function getDictNumberArray(d, name) {
  const arr = getDictArray(d, name);
  if (!arr) return [];
  const out = [];
  for (let i = 0; i < arr.size(); i++) {
    const el = arr.context.lookup(arr.get(i));
    if (el instanceof PDFNumber) out.push(el.asNumber());
  }
  return out;
}

// -----------------------------------------------------------------------------
// Cropper algorithm — direct port of edge-function v22-v30
// -----------------------------------------------------------------------------

/**
 * Tokenize a PDF content stream into parallel TypedArrays. Pre-parses numeric
 * values and maps known operator strings to integer opcodes for fast filtering.
 */
function tokenize(buf) {
  const capacity = Math.max(1024, Math.ceil(buf.length / 4));
  const types = new Int32Array(capacity);
  const starts = new Int32Array(capacity);
  const ends = new Int32Array(capacity);
  const nums = new Float64Array(capacity);
  const opcodes = new Int32Array(capacity);
  let tc = 0;
  const len = buf.length;
  let i = 0;

  while (i < len) {
    const c = buf[i];
    if (c === 0 || c === 9 || c === 10 || c === 12 || c === 13 || c === 32) { i++; continue; }
    if (c === 0x25) { while (i < len && buf[i] !== 10 && buf[i] !== 13) i++; continue; }

    // String
    if (c === 0x28) {
      const start = i; let d = 1; i++;
      while (i < len && d > 0) {
        if (buf[i] === 0x5c) { i += 2; continue; }
        if (buf[i] === 0x28) d++; else if (buf[i] === 0x29) d--;
        i++;
      }
      types[tc] = TT_STRING; starts[tc] = start; ends[tc] = i; tc++;
      continue;
    }
    // Dict or hex string
    if (c === 0x3c) {
      const start = i;
      if (buf[i + 1] === 0x3c) {
        let d = 1; i += 2;
        while (i < len && d > 0) {
          if (buf[i] === 0x3c && buf[i + 1] === 0x3c) { d++; i += 2; }
          else if (buf[i] === 0x3e && buf[i + 1] === 0x3e) { d--; i += 2; }
          else i++;
        }
        types[tc] = TT_DICT; starts[tc] = start; ends[tc] = i; tc++;
        continue;
      }
      while (i < len && buf[i] !== 0x3e) i++;
      i++;
      types[tc] = TT_HEX; starts[tc] = start; ends[tc] = i; tc++;
      continue;
    }
    // Array
    if (c === 0x5b) {
      const start = i; let d = 1; i++;
      while (i < len && d > 0) {
        if (buf[i] === 0x5b) d++; else if (buf[i] === 0x5d) d--;
        i++;
      }
      types[tc] = TT_ARRAY; starts[tc] = start; ends[tc] = i; tc++;
      continue;
    }
    // Name
    if (c === 0x2f) {
      const start = i; i++;
      while (i < len) {
        const b = buf[i];
        if (b === 0 || b === 9 || b === 10 || b === 12 || b === 13 || b === 32 ||
            b === 0x2f || b === 0x28 || b === 0x3c || b === 0x5b || b === 0x5d) break;
        i++;
      }
      types[tc] = TT_NAME; starts[tc] = start; ends[tc] = i; tc++;
      continue;
    }
    // Number
    if ((c >= 0x30 && c <= 0x39) || c === 0x2d || c === 0x2e || c === 0x2b) {
      const start = i; i++;
      while (i < len) {
        const b = buf[i];
        if ((b >= 0x30 && b <= 0x39) || b === 0x2e || b === 0x2d ||
            b === 0x65 || b === 0x45 || b === 0x2b) { i++; continue; }
        break;
      }
      types[tc] = TT_NUM; starts[tc] = start; ends[tc] = i;
      let s = '';
      for (let j = start; j < i; j++) s += String.fromCharCode(buf[j]);
      nums[tc] = parseFloat(s);
      tc++;
      continue;
    }
    // Operator
    if ((c >= 0x41 && c <= 0x5a) || (c >= 0x61 && c <= 0x7a) || c === 0x27 || c === 0x22) {
      const opStart = i; i++;
      while (i < len) {
        const b = buf[i];
        if ((b >= 0x41 && b <= 0x5a) || (b >= 0x61 && b <= 0x7a) ||
            b === 0x2a || b === 0x27 || b === 0x22) { i++; continue; }
        break;
      }
      let opStr = '';
      for (let j = opStart; j < i; j++) opStr += String.fromCharCode(buf[j]);

      // Inline image: BI ... ID ... EI as a single opaque token
      if (opStr === 'BI') {
        let idPos = -1;
        for (let j = i; j < len - 1; j++) {
          if (buf[j] === 0x49 && buf[j + 1] === 0x44) {
            const b0 = j === 0 ? 32 : buf[j - 1];
            const b2 = j + 2 >= len ? 32 : buf[j + 2];
            const before = b0 === 0 || b0 === 9 || b0 === 10 || b0 === 12 || b0 === 13 || b0 === 32;
            const after = b2 === 0 || b2 === 9 || b2 === 10 || b2 === 12 || b2 === 13 || b2 === 32;
            if (before && after) { idPos = j; break; }
          }
        }
        if (idPos < 0) {
          types[tc] = TT_OP; starts[tc] = opStart; ends[tc] = i;
          opcodes[tc] = OP_UNKNOWN; tc++;
          continue;
        }
        let dictStr = '';
        for (let j = i; j < idPos; j++) dictStr += String.fromCharCode(buf[j]);
        let dataLen = -1;
        const lm = dictStr.match(/\/L\s+(\d+)/);
        if (lm) dataLen = parseInt(lm[1], 10);

        let dataEnd = -1;
        if (dataLen >= 0) {
          let ds = idPos + 2;
          if (ds < len) {
            const bs = buf[ds];
            if (bs === 0 || bs === 9 || bs === 10 || bs === 12 || bs === 13 || bs === 32) ds++;
          }
          dataEnd = ds + dataLen;
          let ei = dataEnd;
          while (ei < len) {
            const be = buf[ei];
            if (be === 0 || be === 9 || be === 10 || be === 12 || be === 13 || be === 32) ei++;
            else break;
          }
          if (ei + 1 < len && buf[ei] === 0x45 && buf[ei + 1] === 0x49) {
            dataEnd = ei + 2;
          } else dataEnd = -1;
        }
        if (dataEnd < 0) {
          let scan = idPos + 2;
          if (scan < len) {
            const bs = buf[scan];
            if (bs === 0 || bs === 9 || bs === 10 || bs === 12 || bs === 13 || bs === 32) scan++;
          }
          dataEnd = len;
          while (scan < len - 1) {
            if (buf[scan] === 0x45 && buf[scan + 1] === 0x49) {
              const b0 = scan > 0 ? buf[scan - 1] : 32;
              const b2 = scan + 2 >= len ? 32 : buf[scan + 2];
              const beforeWS = b0 === 0 || b0 === 9 || b0 === 10 || b0 === 12 || b0 === 13 || b0 === 32;
              const afterWS = b2 === 0 || b2 === 9 || b2 === 10 || b2 === 12 || b2 === 13 || b2 === 32;
              if (beforeWS && afterWS) { dataEnd = scan + 2; break; }
            }
            scan++;
          }
        }
        types[tc] = TT_INLINE_IMAGE; starts[tc] = opStart; ends[tc] = dataEnd; tc++;
        i = dataEnd;
        continue;
      }

      const oc = OP_MAP.get(opStr) ?? OP_UNKNOWN;
      types[tc] = TT_OP; starts[tc] = opStart; ends[tc] = i; opcodes[tc] = oc; tc++;
      continue;
    }
    i++;
  }

  return {
    types: types.subarray(0, tc),
    starts: starts.subarray(0, tc),
    ends: ends.subarray(0, tc),
    nums: nums.subarray(0, tc),
    opcodes: opcodes.subarray(0, tc),
    buf,
    length: tc,
  };
}

/**
 * Walk the token list once, producing N keep-bitmaps simultaneously — one per
 * crop rect. Tracks CTM through q/Q/cm and per-path bbox in page coords.
 */
function cropMulti(toks, rects) {
  const { types, opcodes, nums, length: n } = toks;
  const M = rects.length;
  const keeps = [];
  for (let r = 0; r < M; r++) keeps.push(new Uint8Array(n));
  const cx0 = new Float64Array(M), cy0 = new Float64Array(M);
  const cx1 = new Float64Array(M), cy1 = new Float64Array(M);
  for (let r = 0; r < M; r++) {
    cx0[r] = rects[r][0]; cy0[r] = rects[r][1];
    cx1[r] = rects[r][2]; cy1[r] = rects[r][3];
  }

  let ma = 1, mb = 0, mc = 0, md = 1, me = 0, mf = 0;
  const stack = [];

  let pMinX = Infinity, pMinY = Infinity, pMaxX = -Infinity, pMaxY = -Infinity;
  let operandStart = 0;
  const buildBuf = [];

  for (let i = 0; i < n; i++) {
    if (types[i] !== TT_OP) continue;
    const oc = opcodes[i];

    if (oc === OP_q) {
      stack.push([ma, mb, mc, md, me, mf]);
      for (let r = 0; r < M; r++) keeps[r][i] = 1;
      operandStart = i + 1;
      continue;
    }
    if (oc === OP_Q) {
      const s = stack.pop();
      if (s) { ma = s[0]; mb = s[1]; mc = s[2]; md = s[3]; me = s[4]; mf = s[5]; }
      else { ma = 1; mb = 0; mc = 0; md = 1; me = 0; mf = 0; }
      for (let r = 0; r < M; r++) keeps[r][i] = 1;
      operandStart = i + 1;
      continue;
    }
    if (oc === OP_cm) {
      let a2 = 0, b2 = 0, c2 = 0, d2 = 0, e2 = 0, f2 = 0;
      let found = 0;
      for (let j = operandStart; j < i; j++) {
        if (types[j] === TT_NUM) {
          const v = nums[j];
          switch (found) {
            case 0: a2 = v; break; case 1: b2 = v; break;
            case 2: c2 = v; break; case 3: d2 = v; break;
            case 4: e2 = v; break; case 5: f2 = v; break;
          }
          found++;
        }
      }
      if (found === 6) {
        const na = ma * a2 + mc * b2;
        const nb = mb * a2 + md * b2;
        const nc = ma * c2 + mc * d2;
        const nd = mb * c2 + md * d2;
        const ne = ma * e2 + mc * f2 + me;
        const nf = mb * e2 + md * f2 + mf;
        ma = na; mb = nb; mc = nc; md = nd; me = ne; mf = nf;
      }
      for (let r = 0; r < M; r++) {
        keeps[r][i] = 1;
        for (let j = operandStart; j < i; j++) keeps[r][j] = 1;
      }
      operandStart = i + 1;
      continue;
    }

    // Path builds
    if (oc === OP_m || oc === OP_l) {
      let x = 0, y = 0, count = 0;
      for (let j = operandStart; j < i; j++) {
        if (types[j] === TT_NUM) {
          if (count === 0) x = nums[j]; else y = nums[j];
          count++;
        }
      }
      if (count >= 2) {
        const px = ma * x + mc * y + me;
        const py = mb * x + md * y + mf;
        if (px < pMinX) pMinX = px; if (px > pMaxX) pMaxX = px;
        if (py < pMinY) pMinY = py; if (py > pMaxY) pMaxY = py;
      }
      buildBuf.push(i);
      operandStart = i + 1;
      continue;
    }
    if (oc === OP_c || oc === OP_v || oc === OP_y) {
      const expect = oc === OP_c ? 6 : 4;
      const vals = [];
      for (let j = operandStart; j < i && vals.length < expect; j++) {
        if (types[j] === TT_NUM) vals.push(nums[j]);
      }
      for (let k = 0; k + 1 < vals.length; k += 2) {
        const x = vals[k], y = vals[k + 1];
        const px = ma * x + mc * y + me;
        const py = mb * x + md * y + mf;
        if (px < pMinX) pMinX = px; if (px > pMaxX) pMaxX = px;
        if (py < pMinY) pMinY = py; if (py > pMaxY) pMaxY = py;
      }
      buildBuf.push(i);
      operandStart = i + 1;
      continue;
    }
    if (oc === OP_re) {
      let x = 0, y = 0, w = 0, h = 0, count = 0;
      for (let j = operandStart; j < i; j++) {
        if (types[j] === TT_NUM) {
          if (count === 0) x = nums[j];
          else if (count === 1) y = nums[j];
          else if (count === 2) w = nums[j];
          else if (count === 3) h = nums[j];
          count++;
        }
      }
      const corners = [[x, y], [x + w, y], [x + w, y + h], [x, y + h]];
      for (const [xx, yy] of corners) {
        const px = ma * xx + mc * yy + me;
        const py = mb * xx + md * yy + mf;
        if (px < pMinX) pMinX = px; if (px > pMaxX) pMaxX = px;
        if (py < pMinY) pMinY = py; if (py > pMaxY) pMaxY = py;
      }
      buildBuf.push(i);
      operandStart = i + 1;
      continue;
    }
    if (oc === OP_h) {
      buildBuf.push(i);
      operandStart = i + 1;
      continue;
    }

    // Paint operators — decide keep per crop rect
    if (oc >= OP_S && oc <= OP_n) {
      for (let r = 0; r < M; r++) {
        let keepThis = false;
        if (pMaxX > -Infinity) {
          keepThis = !(pMaxX < cx0[r] || pMinX > cx1[r] || pMaxY < cy0[r] || pMinY > cy1[r]);
        }
        if (keepThis) {
          const kr = keeps[r];
          kr[i] = 1;
          for (let j = operandStart; j < i; j++) kr[j] = 1;
          for (let bi = 0; bi < buildBuf.length; bi++) {
            const bidx = buildBuf[bi];
            kr[bidx] = 1;
            for (let j = bidx - 1; j >= 0; j--) {
              if (types[j] === TT_OP) break;
              kr[j] = 1;
            }
          }
        }
      }
      pMinX = Infinity; pMinY = Infinity; pMaxX = -Infinity; pMaxY = -Infinity;
      buildBuf.length = 0;
      operandStart = i + 1;
      continue;
    }

    // State / text / Do / unknown — always keep
    for (let r = 0; r < M; r++) {
      keeps[r][i] = 1;
      for (let j = operandStart; j < i; j++) keeps[r][j] = 1;
    }
    operandStart = i + 1;
  }

  // Inline images: keep all
  for (let i = 0; i < n; i++) {
    if (types[i] === TT_INLINE_IMAGE) {
      for (let r = 0; r < M; r++) keeps[r][i] = 1;
    }
  }

  return keeps;
}

/** Reassemble kept tokens back into a content stream byte buffer. */
function emitKept(toks, keep) {
  const { starts, ends, buf, length: n } = toks;
  let total = 0;
  for (let i = 0; i < n; i++) {
    if (keep[i]) total += (ends[i] - starts[i]) + 1;
  }
  const out = new Uint8Array(total);
  let p = 0;
  for (let i = 0; i < n; i++) {
    if (!keep[i]) continue;
    const s = starts[i], e = ends[i];
    for (let j = s; j < e; j++) out[p++] = buf[j];
    out[p++] = 0x20;
  }
  return out;
}

/**
 * Concatenate all of a page's /Contents streams into one decoded byte buffer.
 */
function decodeFullContent(srcPage) {
  const ctx = srcPage.node.context;
  const contentsRef = srcPage.node.get(PDFName.of('Contents'));
  if (!contentsRef) return null;
  const resolved = ctx.lookup(contentsRef);

  const collectStream = (s) => {
    if (!(s instanceof PDFRawStream || s instanceof PDFStream)) return null;
    if (s instanceof PDFRawStream) {
      const bytes = s.contents;
      const filter = s.dict.get(PDFName.of('Filter'));
      const filterStr = filter ? filter.toString() : '';
      if (filterStr.includes('FlateDecode')) {
        try {
          return pako.inflate(bytes);
        } catch (err) {
          log('inflate error', String(err));
          return null;
        }
      }
      if (!filter) return bytes;
      return null;
    }
    if (s.contents instanceof Uint8Array) return s.contents;
    return null;
  };

  if (resolved instanceof PDFArray) {
    const parts = [];
    for (let i = 0; i < resolved.size(); i++) {
      const sref = resolved.get(i);
      const s = sref instanceof PDFRef ? ctx.lookup(sref) : sref;
      const decoded = collectStream(s);
      if (!decoded) return null;
      parts.push(decoded);
      parts.push(new Uint8Array([0x20]));
    }
    let total = 0;
    for (const p of parts) total += p.length;
    const out = new Uint8Array(total);
    let off = 0;
    for (const p of parts) { out.set(p, off); off += p.length; }
    return out;
  }
  return collectStream(resolved);
}

// -----------------------------------------------------------------------------
// Standalone-PDF generator: take a cropped content-stream buffer and produce a
// 1-page PDF with that content. The page's MediaBox = source MediaBox so the
// kept operators land in the right place. /Resources comes from the source
// page (copied across so the cropped PDF is self-contained — fonts, images,
// extgstate, colorspaces all travel with it).
//
// The edge function uses outDoc.embedPdf(cropDoc, [0]) on this PDF. pdf-lib
// extracts the page as a Form XObject, which is the correct shape for
// drawing in our markup-summary thumbnails.
// -----------------------------------------------------------------------------
async function buildCroppedPdf(srcDoc, srcPageIdx, cropContent, mediaBox) {
  const cropDoc = await PDFDocument.create();
  // Copy the source page (incl. Resources) into the new doc
  const [copiedPage] = await cropDoc.copyPages(srcDoc, [srcPageIdx]);

  // Replace its /Contents with our cropped stream. pdf-lib's PDFPage doesn't
  // expose a clean setter, so we operate on the underlying dict.
  const ctx = cropDoc.context;
  const compressed = pako.deflate(cropContent);
  const newStreamDict = PDFDict.withContext(ctx);
  newStreamDict.set(PDFName.of('Filter'), PDFName.of('FlateDecode'));
  newStreamDict.set(PDFName.of('Length'), PDFNumber.of(compressed.length));
  const newStream = PDFRawStream.of(newStreamDict, compressed);
  const newStreamRef = ctx.register(newStream);

  // Replace /Contents on the copied page with our new stream ref
  copiedPage.node.set(PDFName.of('Contents'), newStreamRef);

  // Strip /Annots — annotations on a page aren't visible when the page is
  // embedded as a Form XObject, but their /AP appearance-stream references
  // pull entire fonts/colorspaces into the file via copyPages. Clearing
  // /Annots lets GC drop those resources.
  copiedPage.node.delete(PDFName.of('Annots'));

  // ----- Prune unused /Resources -----
  // copyPages brings the source page's COMPLETE /Resources, which on
  // architecture plans includes 10+ MB of embedded fonts that the cropped
  // content doesn't use. Walk the cropped content for actually-referenced
  // resource names (/F1 Tf, /Im0 Do, /GS1 gs etc) and prune everything else.
  // After pruning + GC, the embedded font streams that nothing references
  // get cleaned out, dropping cropped-PDF size from ~10 MB to ~50-200 KB.
  pruneUnusedResources(copiedPage, cropContent);

  // Add the page to the doc tree (copyPages registered the page object but
  // didn't add it to the tree)
  cropDoc.addPage(copiedPage);

  // pdf-lib's save() leaves the original /Contents stream as an orphan in the
  // file — it was copied by copyPages but is no longer referenced after we
  // overwrote /Contents above. Resource pruning above also produces orphans.
  // Save once, reload, walk reachables from the trailer, delete unreachable
  // objects, save again. Mirror of the old saveWithGC.
  const dirty = await cropDoc.save();
  return await garbageCollectPdf(dirty);
}

/**
 * Scan the cropped content stream for references to resource names — the
 * tokens that appear immediately before resource-using operators. For each
 * resource type, build a set of names that are actually used by the cropped
 * content, then prune the page's /Resources dict so only those names remain.
 *
 * Resource-using operator forms (PDF 32000-1 §8.4.4-8.4.5, §9.2.4):
 *   /Name Tf        — font (used in text-state)
 *   /Name Do        — XObject (image or form)
 *   /Name gs        — graphics state ExtGState
 *   /Name sh        — shading
 *   /Name cs / /Name CS — colorspace
 *   /Name scn / /Name SCN — pattern colorspace (last operand of scn/SCN)
 *   /Name BMC / /Name BDC — marked content / properties
 *
 * Conservative on edge cases (sh / properties): we keep all if any are
 * referenced, since these are rare and keeping them costs little.
 */
function pruneUnusedResources(page, cropContentBuf) {
  const cropStr = Buffer.from(cropContentBuf).toString('latin1');

  // Collect resource references by category.
  // RegExes that match `/Name <op>` patterns inside the cropped content.
  const collect = (regex) => {
    const found = new Set();
    let m;
    while ((m = regex.exec(cropStr)) !== null) {
      found.add(m[1]);
    }
    return found;
  };

  // Names of fonts referenced via Tf operator. Tf takes TWO operands —
  // /Name <size> Tf — so we have to match a number between the name and Tf.
  const usedFonts = collect(/\/([A-Za-z0-9_.+-]+)\s+[\d.+-]+\s+Tf\b/g);
  // Names of XObjects (forms or images) referenced via Do
  const usedXObjects = collect(/\/([A-Za-z0-9_.+-]+)\s+Do\b/g);
  // Names of ExtGState entries referenced via gs
  const usedExtGState = collect(/\/([A-Za-z0-9_.+-]+)\s+gs\b/g);
  // Shadings via sh
  const usedShading = collect(/\/([A-Za-z0-9_.+-]+)\s+sh\b/g);
  // Colorspaces via cs / CS
  const usedColorSpace = collect(/\/([A-Za-z0-9_.+-]+)\s+(?:cs|CS)\b/g);
  // Patterns are referenced via scn/SCN — the LAST name in operand list,
  // but for simplicity we collect any /Name preceding scn/SCN.
  const usedPattern = collect(/\/([A-Za-z0-9_.+-]+)\s+(?:scn|SCN)\b/g);
  // Properties dict — referenced via BDC / BMC (rare)
  const usedProperties = collect(/\/([A-Za-z0-9_.+-]+)\s+(?:BDC|BMC)\b/g);

  const ctx = page.node.context;
  const resourcesRef = page.node.get(PDFName.of('Resources'));
  if (!resourcesRef) return;
  const resources = resourcesRef instanceof PDFRef
    ? ctx.lookup(resourcesRef)
    : resourcesRef;
  if (!(resources instanceof PDFDict)) return;

  // Helper: prune entries from a sub-dict whose keys aren't in the keep set.
  // Keys with names like "F1", "GS0" are stored as PDFName, so we compare
  // the underlying string.
  const pruneSubDict = (key, keep) => {
    const subRef = resources.get(PDFName.of(key));
    if (!subRef) return;
    const sub = subRef instanceof PDFRef ? ctx.lookup(subRef) : subRef;
    if (!(sub instanceof PDFDict)) return;
    // Collect entries to delete first (mutating during iteration is fragile)
    const toDelete = [];
    for (const [k] of sub.entries()) {
      // PDFName keys: skip leading slash via decodeText / encode roundtrip
      const name = k instanceof PDFName ? k.toString().replace(/^\//, '') : String(k);
      if (!keep.has(name)) toDelete.push(k);
    }
    for (const k of toDelete) sub.delete(k);
    // If sub-dict is now empty, remove the entry from /Resources entirely
    let count = 0;
    for (const _ of sub.entries()) count++;
    if (count === 0) resources.delete(PDFName.of(key));
  };

  pruneSubDict('Font', usedFonts);
  pruneSubDict('XObject', usedXObjects);
  pruneSubDict('ExtGState', usedExtGState);
  pruneSubDict('Shading', usedShading);
  pruneSubDict('ColorSpace', usedColorSpace);
  pruneSubDict('Pattern', usedPattern);
  pruneSubDict('Properties', usedProperties);

  // /ProcSet is always small (an array of names) — leave it untouched.
}

/**
 * Reload a PDF, walk reachable objects from the catalog/trailer, delete
 * everything else, and resave. Removes orphan streams that copyPages leaves
 * behind when we overwrite /Contents.
 */
async function garbageCollectPdf(dirtyBytes) {
  const reDoc = await PDFDocument.load(dirtyBytes);
  const ctx = reDoc.context;
  const reachable = new Set();

  const visit = (obj) => {
    if (!obj) return;
    if (obj instanceof PDFRef) {
      const key = obj.toString();
      if (reachable.has(key)) return;
      reachable.add(key);
      visit(ctx.lookup(obj));
      return;
    }
    if (obj instanceof PDFDict) {
      for (const [, v] of obj.entries()) visit(v);
      return;
    }
    if (obj instanceof PDFArray) {
      for (const v of obj.asArray()) visit(v);
      return;
    }
    if (obj instanceof PDFStream) {
      visit(obj.dict);
      return;
    }
  };

  const ti = ctx.trailerInfo;
  if (ti?.Root) visit(ti.Root);
  if (ti?.Info) visit(ti.Info);
  if (ti?.ID) visit(ti.ID);
  if (reachable.size === 0) {
    for (const [ref, obj] of ctx.enumerateIndirectObjects()) {
      if (obj === reDoc.catalog) { visit(ref); break; }
    }
  }

  const allRefs = [];
  for (const [r] of ctx.enumerateIndirectObjects()) allRefs.push(r);
  let deleted = 0;
  for (const ref of allRefs) {
    if (!reachable.has(ref.toString())) {
      ctx.delete(ref);
      deleted++;
    }
  }

  const cleaned = await reDoc.save();
  log('GC walk', {
    dirtyBytes: dirtyBytes.length,
    cleanBytes: cleaned.length,
    totalObjects: allRefs.length,
    reachable: reachable.size,
    deletedOrphans: deleted,
    reduction: `${((1 - cleaned.length / dirtyBytes.length) * 100).toFixed(1)}%`,
  });
  return cleaned;
}

// -----------------------------------------------------------------------------
// Markup extraction (mirrored from edge function — keep in sync)
// -----------------------------------------------------------------------------
function extractMarkups(pdfDoc) {
  const out = [];
  const pages = pdfDoc.getPages();
  for (let pi = 0; pi < pages.length; pi++) {
    const page = pages[pi];
    const annotsRef = page.node.get(PDFName.of('Annots'));
    if (!annotsRef) continue;
    const annots = page.node.context.lookup(annotsRef);
    if (!(annots instanceof PDFArray)) continue;

    for (let i = 0; i < annots.size(); i++) {
      const aRef = annots.get(i);
      const annot = annots.context.lookup(aRef);
      if (!(annot instanceof PDFDict)) continue;
      const subtypeObj = annot.get(PDFName.of('Subtype'));
      if (!(subtypeObj instanceof PDFName)) continue;
      if (subtypeObj.toString() !== '/Polygon') continue;

      // Skip grouped children — parent already covers them
      if (annot.has(PDFName.of('IRT'))) continue;

      const bsi = getDictArray(annot, 'BSIColumnData');
      if (!bsi) continue;

      const nm = getDictStr(annot, 'NM');
      const rectArr = getDictNumberArray(annot, 'Rect');
      const rect = rectArr.length === 4 ? rectArr : [0, 0, 0, 0];

      // Also pull /Vertices - Bluebeam writes /Rect as the polygon's
      // control-point bbox, which sits INSIDE the visible scalloped cloud.
      // Computing bbox from vertices captures the actual visible extent
      // (scallop tips bulge outward past the rect by ~chord/2 per edge).
      const vertArr = getDictNumberArray(annot, 'Vertices');
      const vertices = [];
      for (let vi = 0; vi + 1 < vertArr.length; vi += 2) {
        vertices.push([vertArr[vi], vertArr[vi + 1]]);
      }

      out.push({ page_index: pi, nm, rect, vertices });
    }
  }
  return out;
}

// -----------------------------------------------------------------------------
// Per-page processing
// -----------------------------------------------------------------------------
async function processPage(srcDoc, srcPageIdx, markups) {
  const srcPage = srcDoc.getPages()[srcPageIdx];
  if (!srcPage) {
    log('processPage: page not found', srcPageIdx);
    return [];
  }
  const rotation = (srcPage.getRotation().angle) % 360;
  if (rotation !== 0) {
    log('processPage: skipping rotated page', { srcPageIdx, rotation });
    return markups.map((m) => ({ nm: m.nm, status: 'skipped_rotated' }));
  }

  const decoded = decodeFullContent(srcPage);
  if (!decoded) {
    log('processPage: cannot decode content', srcPageIdx);
    return markups.map((m) => ({ nm: m.nm, status: 'skipped_undecodable' }));
  }

  if (decoded.length > MAX_DECODED_FOR_V24) {
    log('processPage: content too large', { srcPageIdx, decodedBytes: decoded.length });
    return markups.map((m) => ({ nm: m.nm, status: 'skipped_too_large' }));
  }

  const tokTime = Date.now();
  const toks = tokenize(decoded);
  log('processPage: tokenized', {
    srcPageIdx,
    decodedBytes: decoded.length,
    tokens: toks.length,
    ms: Date.now() - tokTime,
  });

  if (toks.length > MAX_TOKENS_FOR_V24) {
    log('processPage: too many tokens', { srcPageIdx, tokens: toks.length });
    return markups.map((m) => ({ nm: m.nm, status: 'skipped_too_many_tokens' }));
  }

  // Compute crop rects with THUMB_PAD_PT padding, clamped to MediaBox
  const mediaBox = srcPage.getMediaBox();
  const minX = mediaBox.x;
  const minY = mediaBox.y;
  const maxX = mediaBox.x + mediaBox.width;
  const maxY = mediaBox.y + mediaBox.height;
  const cropRects = markups.map((m) => {
    let x0, y0, x1, y1;
    // Prefer /Vertices bbox - captures the visible cloud's actual extent
    // including scallop tips. Falls back to /Rect when vertices are missing
    // (non-PolygonCloud annotations) or degenerate.
    if (m.vertices && m.vertices.length >= 3) {
      let minVx = Infinity, minVy = Infinity, maxVx = -Infinity, maxVy = -Infinity;
      for (const [vx, vy] of m.vertices) {
        if (vx < minVx) minVx = vx;
        if (vy < minVy) minVy = vy;
        if (vx > maxVx) maxVx = vx;
        if (vy > maxVy) maxVy = vy;
      }
      x0 = minVx; y0 = minVy; x1 = maxVx; y1 = maxVy;
    } else {
      [x0, y0, x1, y1] = m.rect;
    }
    if (!isFinite(x0) || x1 <= x0 || y1 <= y0) {
      x0 = minX; y0 = minY; x1 = maxX; y1 = maxY;
    }
    return [
      Math.max(minX, x0 - THUMB_PAD_PT),
      Math.max(minY, y0 - THUMB_PAD_PT),
      Math.min(maxX, x1 + THUMB_PAD_PT),
      Math.min(maxY, y1 + THUMB_PAD_PT),
    ];
  });

  const cropTime = Date.now();
  const keeps = cropMulti(toks, cropRects);
  log('processPage: cropped', {
    srcPageIdx,
    markups: markups.length,
    ms: Date.now() - cropTime,
  });

  const results = [];
  for (let i = 0; i < markups.length; i++) {
    const m = markups[i];
    try {
      const cropped = emitKept(toks, keeps[i]);
      results.push({
        nm: m.nm,
        status: 'ok',
        cropContent: cropped,
        cropRect: cropRects[i],
        srcPageIdx,
      });
    } catch (err) {
      log('processPage: emit failed', { nm: m.nm, err: String(err) });
      results.push({ nm: m.nm, status: 'error', error: String(err) });
    }
  }
  return results;
}

// -----------------------------------------------------------------------------
// Main
// -----------------------------------------------------------------------------
async function main() {
  const args = process.argv.slice(2);
  if (args.length < 2) {
    console.error('Usage: node cropper.js <plan-pdf-path> <output-dir>');
    process.exit(1);
  }
  const [planPath, outDir] = args;

  if (!fs.existsSync(planPath)) {
    console.error(`Plan PDF not found: ${planPath}`);
    process.exit(1);
  }
  fs.mkdirSync(outDir, { recursive: true });

  const planBytes = fs.readFileSync(planPath);
  log('plan loaded', { path: planPath, bytes: planBytes.length });

  // Compute hash for manifest validation
  const hash = crypto.createHash('sha256').update(planBytes).digest('hex').substring(0, 16);
  log('plan hash', hash);

  let srcDoc;
  try {
    srcDoc = await PDFDocument.load(planBytes);
  } catch (err) {
    console.error(`Failed to parse PDF: ${err.message}`);
    process.exit(2);
  }

  const allMarkups = extractMarkups(srcDoc);
  log('markups found', allMarkups.length);

  if (allMarkups.length === 0) {
    // Still write a manifest — edge function checks existence, not content
    const manifest = {
      version: 1,
      plan_pdf_hash: hash,
      generated_at: new Date().toISOString(),
      markups: [],
    };
    fs.writeFileSync(path.join(outDir, 'manifest.json'), JSON.stringify(manifest, null, 2));
    log('manifest (empty) written');
    process.exit(0);
  }

  // Group by page for batched cropping
  const byPage = new Map();
  for (const m of allMarkups) {
    if (!byPage.has(m.page_index)) byPage.set(m.page_index, []);
    byPage.get(m.page_index).push(m);
  }

  const manifestEntries = [];
  let okCount = 0;
  let skipCount = 0;
  let errCount = 0;

  for (const [pageIdx, markups] of byPage.entries()) {
    const results = await processPage(srcDoc, pageIdx, markups);
    for (const r of results) {
      if (r.status === 'ok') {
        // Write self-contained PDF for this markup
        const filename = `markup_${sanitizeNm(r.nm)}.pdf`;
        const outPath = path.join(outDir, filename);
        try {
          const cropPdfBytes = await buildCroppedPdf(
            srcDoc, r.srcPageIdx, r.cropContent, srcDoc.getPages()[r.srcPageIdx].getMediaBox()
          );
          fs.writeFileSync(outPath, cropPdfBytes);
          manifestEntries.push({
            nm: r.nm,
            page_index: r.srcPageIdx,
            crop_pdf_path: filename,
            crop_rect: r.cropRect,
          });
          okCount++;
        } catch (err) {
          log('build failed', { nm: r.nm, err: String(err) });
          errCount++;
        }
      } else {
        log('markup skipped', { nm: r.nm, status: r.status });
        skipCount++;
      }
    }
  }

  // Write manifest
  const manifest = {
    version: 1,
    plan_pdf_hash: hash,
    generated_at: new Date().toISOString(),
    markups: manifestEntries,
  };
  fs.writeFileSync(path.join(outDir, 'manifest.json'), JSON.stringify(manifest, null, 2));

  log('done', { ok: okCount, skipped: skipCount, error: errCount, totalMs: Date.now() - t0 });

  // Exit 0 even if some markups had errors — partial success is still useful.
  // Edge function falls back to placeholder for any /NM not in the manifest.
  process.exit(0);
}

/**
 * The /NM key may contain characters unsafe for filenames (PDF doesn't restrict
 * it; common values are like "abc123-def456" UUIDs, but Bluebeam can use any
 * string). Sanitize to a filesystem-safe slug — keep alphanumerics and dashes,
 * hash everything else. The edge function reads filenames from the manifest,
 * so the on-disk name only has to roundtrip through the manifest.
 */
function sanitizeNm(nm) {
  if (!nm) return 'unknown';
  if (/^[A-Za-z0-9\-_.]+$/.test(nm) && nm.length <= 64) return nm;
  return crypto.createHash('sha256').update(nm).digest('hex').substring(0, 16);
}

main().catch((err) => {
  console.error('FATAL', err);
  process.exit(3);
});
