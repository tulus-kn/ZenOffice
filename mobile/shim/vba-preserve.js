/**
 * vba-preserve.js — keep Excel macro assets alive when the in-app spreadsheet
 * engine rebuilds the OOXML package.
 *
 * The engine (ExcelJS / SheetJS) re-serialises the workbook from a cell
 * snapshot, so every part it does not model is dropped on save:
 *   - xl/vbaProject.bin (+ signature)      → the macro code itself
 *   - customUI/customUI.xml                → the macro ribbon / toolbar
 *   - xl/drawings/vmlDrawing*.vml          → form-control buttons bound to macros
 *
 * This module puts those parts back and repairs the declarations that make
 * them reachable (content types, root rels, sheet rels). The produced package
 * is then re-parsed; when a stage cannot be proven consistent it is discarded
 * instead of risking a corrupt workbook.
 *
 * Macros are preserved, never executed: there is no VBA engine here.
 */
(function () {
  'use strict'

  var REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  var VBA_REL = 'http://schemas.microsoft.com/office/2006/relationships/vbaProject'
  var UI_REL = 'http://schemas.microsoft.com/office/2006/relationships/ui/extensibility'
  var VML_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing'

  var MACRO_MAIN_CT = 'application/vnd.ms-excel.sheet.macroEnabled.main+xml'
  var PLAIN_MAIN_CT =
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
  var VML_CT = 'application/vnd.openxmlformats-officedocument.vmlDrawing'
  var CUSTOM_UI_CT = 'application/vnd.ms-office.customUI+xml'

  var PART_TYPES = [
    [/^xl\/vbaProject\.bin$/i, 'application/vnd.ms-office.vbaProject'],
    [/^xl\/vbaProjectSignature\.bin$/i, 'application/vnd.ms-office.vbaProjectSignature'],
    [/^xl\/vbaProjectSignatureAgile\.bin$/i, 'application/vnd.ms-office.vbaProjectSignature'],
    [/^xl\/vbaProjectSignatureV3\.bin$/i, 'application/vnd.ms-office.vbaProjectSignature'],
    [/^customUI\/customUI\.xml$/i, CUSTOM_UI_CT],
  ]

  // ---------------------------------------------------------------- helpers

  function isMacroPart(name) {
    return /^xl\/vbaProject/i.test(name)
  }

  function contentTypeFor(name) {
    for (var i = 0; i < PART_TYPES.length; i++) {
      if (PART_TYPES[i][0].test(name)) return PART_TYPES[i][1]
    }
    return null
  }

  function escapeXml(value) {
    return String(value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
  }

  function unescapeXml(value) {
    return String(value)
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&amp;/g, '&')
  }

  function dirname(path) {
    var i = path.lastIndexOf('/')
    return i < 0 ? '' : path.slice(0, i)
  }

  function basename(path) {
    var i = path.lastIndexOf('/')
    return i < 0 ? path : path.slice(i + 1)
  }

  /** Resolve an OPC relationship target against the directory that owns it. */
  function resolveTarget(baseDir, target) {
    var clean = String(target || '').replace(/\\/g, '/')
    if (!clean) return ''
    var segments = (baseDir ? baseDir.split('/') : []).concat(clean.split('/'))
    var out = []
    for (var i = 0; i < segments.length; i++) {
      var s = segments[i]
      if (s === '' || s === '.') continue
      if (s === '..') {
        if (out.length) out.pop()
        continue
      }
      out.push(s)
    }
    return out.join('/')
  }

  function relsPathFor(part) {
    return dirname(part) + '/_rels/' + basename(part) + '.rels'
  }

  /** Relative target from xl/worksheets/ to a part such as xl/drawings/x.vml */
  function fromSheetDir(part) {
    return '../' + String(part).replace(/^xl\//, '')
  }

  async function readText(zip, path) {
    var file = zip.file(path)
    return file ? await file.async('string') : null
  }

  function relMapOf(relsXml) {
    var map = {}
    if (!relsXml) return map
    var re = /<Relationship\b[^>]*>/g
    var m
    while ((m = re.exec(relsXml))) {
      var id = /\bId="([^"]+)"/.exec(m[0])
      var target = /\bTarget="([^"]+)"/.exec(m[0])
      if (id && target) map[id[1]] = target[1]
    }
    return map
  }

  function nextRelId(relsXml, prefix) {
    var used = []
    var re = /\bId="([^"]+)"/g
    var m
    while ((m = re.exec(relsXml || ''))) used.push(m[1])
    var id = prefix || 'rIdZen'
    while (used.indexOf(id) >= 0) id += 'X'
    return id
  }

  function addRelationship(relsXml, id, type, target) {
    var entry = '<Relationship Id="' + id + '" Type="' + type + '" Target="' + target + '"/>'
    if (!relsXml || !/<\/Relationships>/.test(relsXml)) {
      return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        entry + '</Relationships>'
    }
    return relsXml.replace(/<\/Relationships>\s*$/, entry + '</Relationships>')
  }

  function addOverride(ct, part, type) {
    if (ct.indexOf('PartName="/' + part + '"') >= 0) return ct
    var entry = '<Override PartName="/' + escapeXml(part) + '" ContentType="' + type + '"/>'
    return ct.replace(/<\/Types>\s*$/, entry + '</Types>')
  }

  function addDefault(ct, ext, type) {
    if (new RegExp('Extension="' + ext + '"', 'i').test(ct)) return ct
    var entry = '<Default Extension="' + ext + '" ContentType="' + type + '"/>'
    return ct.replace(/<\/Types>\s*$/, entry + '</Types>')
  }

  /** Sheets as [{name, part}] taken from the workbook part + its rels. */
  async function sheetParts(zip) {
    var wb = await readText(zip, 'xl/workbook.xml')
    if (!wb) return []
    var rels = relMapOf(await readText(zip, 'xl/_rels/workbook.xml.rels'))
    var out = []
    var re = /<sheet\b[^>]*>/g
    var m
    while ((m = re.exec(wb))) {
      var name = /\bname="([^"]*)"/.exec(m[0])
      var rid = /\br:id="([^"]+)"/.exec(m[0])
      if (!name || !rid || !rels[rid[1]]) continue
      out.push({ name: unescapeXml(name[1]), part: resolveTarget('xl', rels[rid[1]]) })
    }
    return out
  }

  // ------------------------------------------------------- macro inspection

  async function inspectMacros(bytes) {
    if (!window.JSZip || !bytes) return { hasMacro: false, parts: [] }
    try {
      var zip = await window.JSZip.loadAsync(bytes)
      var names = Object.keys(zip.files)
      var parts = names.filter(isMacroPart)
      var ui = names.filter(function (n) {
        return /^customUI\//i.test(n) && !zip.files[n].dir
      })
      var vml = names.filter(function (n) {
        return /^xl\/drawings\/vmlDrawing.*\.vml$/i.test(n)
      })
      return { hasMacro: parts.length > 0, parts: parts, customUi: ui, vml: vml }
    } catch (e) {
      return { hasMacro: false, parts: [], error: String((e && e.message) || e) }
    }
  }

  /**
   * The engine cannot parse VML drawing parts, and a worksheet that declares
   * one without a matching comments part makes `ExcelJS.xlsx.load()` throw —
   * that is exactly the shape of a workbook with form-control buttons bound to
   * macros. Feed the engine a VML-free copy so opening/saving keeps working;
   * the untouched original stays in session.bytes and is merged back on save.
   */
  async function prepareEngineBytes(bytes) {
    if (!window.JSZip || !bytes) return bytes
    try {
      var zip = await window.JSZip.loadAsync(bytes)
      var names = Object.keys(zip.files)
      var hasVml = names.some(function (n) {
        return /^xl\/drawings\/vmlDrawing.*\.vml$/i.test(n)
      })
      if (!hasVml) return bytes

      for (var i = 0; i < names.length; i++) {
        var n = names[i]
        if (/^xl\/drawings\/vmlDrawing.*\.vml$/i.test(n)) zip.remove(n)
        else if (/^xl\/drawings\/_rels\/vmlDrawing.*\.vml\.rels$/i.test(n)) zip.remove(n)
      }
      for (var r = 0; r < names.length; r++) {
        var relName = names[r]
        if (!/^xl\/worksheets\/_rels\/.*\.rels$/i.test(relName)) continue
        var relText = await readText(zip, relName)
        if (!relText || relText.indexOf('vmlDrawing') < 0) continue
        zip.file(relName, relText.replace(/<Relationship\b[^>]*vmlDrawing[^>]*\/>/g, ''))
      }
      for (var s = 0; s < names.length; s++) {
        var sheetName = names[s]
        if (!/^xl\/worksheets\/sheet\d+\.xml$/i.test(sheetName)) continue
        var sheetText = await readText(zip, sheetName)
        if (!sheetText || sheetText.indexOf('<legacyDrawing') < 0) continue
        zip.file(sheetName, sheetText.replace(/<legacyDrawing\b[^>]*\/>/g, ''))
      }
      return await zip.generateAsync({ type: 'uint8array' })
    } catch (e) {
      console.warn('[vba-preserve] could not strip VML for the engine', e)
      return bytes
    }
  }

  // ------------------------------------------------------------ macro parts

  async function copyMacroParts(out, src) {
    var parts = Object.keys(src.files).filter(function (n) {
      return isMacroPart(n) && !src.files[n].dir
    })
    for (var i = 0; i < parts.length; i++) {
      out.file(parts[i], await src.file(parts[i]).async('uint8array'))
    }
    if (parts.length) {
      var ct = await readText(out, '[Content_Types].xml')
      if (ct) {
        ct = ct.split(PLAIN_MAIN_CT).join(MACRO_MAIN_CT)
        for (var p = 0; p < parts.length; p++) {
          var type = contentTypeFor(parts[p])
          if (type) ct = addOverride(ct, parts[p], type)
        }
        out.file('[Content_Types].xml', ct)
      }
      var relsPath = 'xl/_rels/workbook.xml.rels'
      var rels = await readText(out, relsPath)
      if (rels && rels.indexOf('vbaProject.bin') < 0) {
        out.file(
          relsPath,
          addRelationship(rels, nextRelId(rels, 'rIdZenVba'), VBA_REL, 'vbaProject.bin'),
        )
      }
    }
    return parts
  }

  // ------------------------------------------------ customUI (macro ribbon)

  async function copyCustomUi(out, src) {
    var parts = Object.keys(src.files).filter(function (n) {
      return /^customUI\//i.test(n) && !src.files[n].dir
    })
    if (!parts.length) return []
    for (var i = 0; i < parts.length; i++) {
      out.file(parts[i], await src.file(parts[i]).async('uint8array'))
    }
    var ct = await readText(out, '[Content_Types].xml')
    if (ct) {
      for (var p = 0; p < parts.length; p++) {
        var type = contentTypeFor(parts[p])
        if (type) {
          ct = addOverride(ct, parts[p], type)
        } else {
          var ext = parts[p].split('.').pop().toLowerCase()
          if (ext === 'png') ct = addDefault(ct, 'png', 'image/png')
          if (ext === 'jpg' || ext === 'jpeg') ct = addDefault(ct, 'jpeg', 'image/jpeg')
        }
      }
      out.file('[Content_Types].xml', ct)
    }
    var root = await readText(out, '_rels/.rels')
    if (root && root.indexOf('customUI/customUI.xml') < 0) {
      out.file(
        '_rels/.rels',
        addRelationship(root, nextRelId(root, 'rIdZenUi'), UI_REL, 'customUI/customUI.xml'),
      )
    }
    return parts
  }

  // ------------------------------ VML drawings (macro-bound form buttons)

  async function copyLegacyDrawings(out, src) {
    var srcSheets = await sheetParts(src)
    var outSheets = await sheetParts(out)
    var byName = {}
    for (var o = 0; o < outSheets.length; o++) byName[outSheets[o].name] = outSheets[o].part

    var copied = []
    for (var i = 0; i < srcSheets.length; i++) {
      var sheet = srcSheets[i]
      var outPart = byName[sheet.name]
      if (!outPart) continue

      var srcXml = await readText(src, sheet.part)
      if (!srcXml) continue
      var legacy = /<legacyDrawing\b[^>]*\/>/.exec(srcXml)
      if (!legacy) continue
      var ridMatch = /\br:id="([^"]+)"/.exec(legacy[0])
      if (!ridMatch) continue

      var srcSheetRels = relMapOf(await readText(src, relsPathFor(sheet.part)))
      var target = srcSheetRels[ridMatch[1]]
      if (!target) continue
      var vmlPart = resolveTarget(dirname(sheet.part), target)
      if (!src.file(vmlPart)) continue

      out.file(vmlPart, await src.file(vmlPart).async('uint8array'))
      copied.push(vmlPart)

      // images referenced by the drawing
      var vmlRelsPath = relsPathFor(vmlPart)
      var vmlRels = await readText(src, vmlRelsPath)
      if (vmlRels) {
        out.file(vmlRelsPath, vmlRels)
        var map = relMapOf(vmlRels)
        var ids = Object.keys(map)
        for (var r = 0; r < ids.length; r++) {
          var mediaPart = resolveTarget(dirname(vmlPart), map[ids[r]])
          if (src.file(mediaPart)) {
            out.file(mediaPart, await src.file(mediaPart).async('uint8array'))
            copied.push(mediaPart)
          }
        }
      }

      var ct = await readText(out, '[Content_Types].xml')
      if (ct) out.file('[Content_Types].xml', addDefault(ct, 'vml', VML_CT))

      // link it from the rewritten sheet with a fresh relationship id
      var outXml = await readText(out, outPart)
      if (!outXml || /<legacyDrawing\b/.test(outXml)) continue
      var outRelsPath = relsPathFor(outPart)
      var outRels = (await readText(out, outRelsPath)) || ''
      var newId = nextRelId(outRels, 'rIdZenVml')
      out.file(outRelsPath, addRelationship(outRels, newId, VML_REL, fromSheetDir(vmlPart)))
      outXml = outXml.replace(
        /<\/worksheet>\s*$/,
        '<legacyDrawing xmlns:r="' + REL_NS + '" r:id="' + newId + '"/></worksheet>',
      )
      out.file(outPart, outXml)
    }
    return copied
  }

  // ------------------------------------------------------------- validation

  function parseXml(text) {
    try {
      var doc = new DOMParser().parseFromString(text, 'application/xml')
      return doc.getElementsByTagName('parsererror').length ? null : doc
    } catch (e) {
      return null
    }
  }

  async function validate(bytes) {
    var errors = []
    var zip = await window.JSZip.loadAsync(bytes)
    var names = Object.keys(zip.files).filter(function (n) {
      return !zip.files[n].dir
    })
    var ct = await readText(zip, '[Content_Types].xml')
    if (!ct || !parseXml(ct)) errors.push('content types unreadable')
    for (var i = 0; i < names.length; i++) {
      var n = names[i]
      if (!/\.(xml|rels|vml)$/i.test(n)) continue
      var text = await readText(zip, n)
      if (!text) continue
      if (!parseXml(text)) errors.push('bad xml: ' + n)
    }
    return errors
  }

  /** every r:id injected into a sheet must resolve in that sheet's rels */
  async function validateLinks(bytes) {
    var errors = []
    var zip = await window.JSZip.loadAsync(bytes)
    var sheets = await sheetParts(zip)
    for (var i = 0; i < sheets.length; i++) {
      var part = sheets[i].part
      var xml = await readText(zip, part)
      if (!xml) continue
      var rels = relMapOf(await readText(zip, relsPathFor(part)))
      var re = /<legacyDrawing\b[^>]*\br:id="([^"]+)"/g
      var m
      while ((m = re.exec(xml))) {
        if (!rels[m[1]]) errors.push('dangling legacyDrawing r:id in ' + part)
        else if (!zip.file(resolveTarget(dirname(part), rels[m[1]]))) {
          errors.push('missing drawing part for ' + part)
        }
      }
    }
    return errors
  }

  // ------------------------------------------------------------------- API

  async function preserveMacroParts(savedBytes, originalBytes) {
    if (!window.JSZip) throw new Error('JSZip is unavailable')
    if (!savedBytes || !originalBytes) {
      return { bytes: savedBytes, preserved: false, parts: [], extras: {} }
    }

    var src = await window.JSZip.loadAsync(originalBytes)
    var macroParts = Object.keys(src.files).filter(function (n) {
      return isMacroPart(n) && !src.files[n].dir
    })
    var hasCustomUi = Object.keys(src.files).some(function (n) {
      return /^customUI\/customUI\.xml$/i.test(n)
    })
    var hasVml = Object.keys(src.files).some(function (n) {
      return /^xl\/drawings\/vmlDrawing.*\.vml$/i.test(n)
    })
    if (!macroParts.length && !hasCustomUi && !hasVml) {
      return { bytes: savedBytes, preserved: false, parts: [], extras: {} }
    }

    try {
      var out = await window.JSZip.loadAsync(savedBytes)
      var extracted = { parts: [], extras: { customUi: [], vml: [] } }

      if (macroParts.length) extracted.parts = await copyMacroParts(out, src)
      if (hasCustomUi) extracted.extras.customUi = await copyCustomUi(out, src)
      if (hasVml) extracted.extras.vml = await copyLegacyDrawings(out, src)

      var bytes = await out.generateAsync({
        type: 'uint8array',
        compression: 'DEFLATE',
        mimeType: 'application/vnd.ms-excel.sheet.macroEnabled.12',
      })

      var errors = (await validate(bytes)).concat(await validateLinks(bytes))
      if (errors.length) {
        console.warn('[vba-preserve] validation failed, keeping the engine output', errors.slice(0, 5))
        return { bytes: savedBytes, preserved: false, parts: [], extras: {}, errors: errors.slice(0, 5) }
      }

      return { bytes: bytes, preserved: true, parts: extracted.parts, extras: extracted.extras }
    } catch (e) {
      console.warn('[vba-preserve] could not rebuild the macro package', e)
      return {
        bytes: savedBytes,
        preserved: false,
        parts: [],
        extras: {},
        error: String((e && e.message) || e),
      }
    }
  }


  // --------------------------------------- non-cell parts: charts, pivots, ActiveX
  //
  // The engine rebuilds the package from cell data, so every part that is not a
  // cell (charts, pivot tables and their caches, slicers, ActiveX controls,
  // drawings, embedded objects) disappears. These helpers copy those parts back
  // byte for byte, restore the worksheet elements that reference them, and keep
  // [Content_Types].xml and the relationship files consistent. On any doubt the
  // merge bails out and the caller keeps the engine output instead.

  var EXTRA_DIRS = /^xl\/(charts|drawings|media|activeX|ctrlProps|embeddings|slicers|slicerCaches|pivotTables|pivotCache)\//i
  var EXTRA_SKIP = /vmlDrawing/i

  var EXTRA_TYPES = [
    [/^xl\/charts\//i, 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'],
    [/^xl\/drawings\//i, 'application/vnd.openxmlformats-officedocument.drawing+xml'],
    [/^xl\/activeX\//i, 'application/vnd.ms-office.activeX+xml'],
    [/^xl\/ctrlProps\//i, 'application/vnd.ms-office.activeXProperties+xml'],
    [/^xl\/pivotTables\//i, 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml'],
    [/^xl\/pivotCache\/pivotCacheDefinition/i, 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml'],
    [/^xl\/pivotCache\/pivotCacheRecords/i, 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml'],
    [/^xl\/slicers\//i, 'application/vnd.ms-excel.slicer+xml'],
    [/^xl\/slicerCaches\//i, 'application/vnd.ms-excel.slicerCache+xml'],
    [/^xl\/embeddings\//i, 'application/vnd.openxmlformats-officedocument.oleObject'],
  ]

  var MEDIA_TYPES = {
    png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif',
    emf: 'image/x-emf', wmf: 'image/x-wmf', bmp: 'image/bmp', tiff: 'image/tiff', svg: 'image/svg+xml',
  }

  function isExtraPart(name) {
    return EXTRA_DIRS.test(name) && !EXTRA_SKIP.test(name)
  }

  function extraType(name) {
    for (var i = 0; i < EXTRA_TYPES.length; i++) {
      if (EXTRA_TYPES[i][0].test(name)) return EXTRA_TYPES[i][1]
    }
    return null
  }

  /** [{id, type, target}] for every relationship in a rels document. */
  function relEntries(relsXml) {
    var list = []
    if (!relsXml) return list
    var re = /<Relationship\b[^>]*>/g
    var m
    while ((m = re.exec(relsXml))) {
      var id = /\bId="([^"]+)"/.exec(m[0])
      var type = /\bType="([^"]+)"/.exec(m[0])
      var target = /\bTarget="([^"]+)"/.exec(m[0])
      if (id && type && target) list.push({ id: id[1], type: type[1], target: target[1] })
    }
    return list
  }

  function blocksOf(text, re) {
    var list = []
    var m
    while ((m = re.exec(text))) list.push(m[0])
    return list
  }

  function ensureRNamespace(text) {
    if (/xmlns:r=/.test(text)) return text
    return text.replace(/<worksheet\b/, '<worksheet xmlns:r="' + REL_NS + '"')
  }

  function insertBefore(text, block, tokens, closeTag) {
    var at = -1
    for (var i = 0; i < tokens.length; i++) {
      var index = text.indexOf('<' + tokens[i])
      if (index >= 0 && (at < 0 || index < at)) at = index
    }
    if (at < 0) at = text.lastIndexOf(closeTag || '</worksheet>')
    if (at < 0) return text
    return text.slice(0, at) + block + text.slice(at)
  }

  async function copyExtraParts(out, src) {
    var copied = []
    var names = Object.keys(src.files)
    for (var i = 0; i < names.length; i++) {
      var name = names[i]
      if (!isExtraPart(name) || src.files[name].dir) continue
      // a plain drawing part may have been generated by the engine for images;
      // keep that one and only restore parts the engine never writes
      if (/^xl\/drawings\/drawing\d+\.xml$/i.test(name) && out.files[name]) continue
      out.file(name, await src.files[name].async('uint8array'))
      copied.push(name)
    }
    if (!copied.length) return copied

    var ct = await readText(out, '[Content_Types].xml')
    if (ct) {
      for (var c = 0; c < copied.length; c++) {
        var part = copied[c]
        if (/^xl\/media\//i.test(part)) {
          var ext = (part.split('.').pop() || '').toLowerCase()
          var mediaType = MEDIA_TYPES[ext]
          if (mediaType) ct = addDefault(ct, ext, mediaType)
          continue
        }
        var type = extraType(part)
        if (type) ct = addOverride(ct, part, type)
      }
      out.file('[Content_Types].xml', ct)
    }
    return copied
  }

  async function mergeSheetExtras(out, src) {
    var srcSheets = await sheetParts(src)
    var outSheets = await sheetParts(out)
    var report = { conditionalFormatting: 0, drawing: 0, oleObjects: 0, controls: 0, relations: 0 }
    var count = Math.min(srcSheets.length, outSheets.length)

    for (var i = 0; i < count; i++) {
      var srcPart = srcSheets[i].part
      var outPart = outSheets[i].part
      var srcText = await readText(src, srcPart)
      var outText = await readText(out, outPart)
      if (!srcText || !outText) continue

      var outRelsPath = relsPathFor(outPart)
      var outRelsXml = await readText(out, outRelsPath)
      var outRels = relMapOf(outRelsXml)
      var srcRelsXml = await readText(src, relsPathFor(srcPart))
      var srcRels = relEntries(srcRelsXml)

      var cf = blocksOf(srcText, /<conditionalFormatting[\s\S]*?<\/conditionalFormatting>/g)
      var draw = blocksOf(srcText, /<drawing\b[^>]*\/>/g)
      var ole = blocksOf(srcText, /<oleObjects[\s\S]*?<\/oleObjects>/g)
      var ctl = blocksOf(srcText, /<controls[\s\S]*?<\/controls>/g)

      // relationships the restored blocks point at
      var relsChanged = false
      var referenced = []
      ;[draw, ole, ctl].forEach(function (group) {
        group.forEach(function (block) {
          var re = /r:id="([^"]+)"/g
          var m
          while ((m = re.exec(block))) referenced.push(m[1])
        })
      })
      referenced.forEach(function (id) {
        if (outRels[id]) return
        var entry = null
        for (var r = 0; r < srcRels.length; r++) if (srcRels[r].id === id) entry = srcRels[r]
        if (!entry) return
        if (!out.files[resolveTarget(dirname(srcPart), entry.target)]) return
        outRelsXml = addRelationship(outRelsXml, id, entry.type, entry.target)
        outRels[id] = entry.target
        relsChanged = true
        report.relations += 1
      })
      if (relsChanged) out.file(outRelsPath, outRelsXml)
      if ((draw.length || ole.length || ctl.length) && /r:id=/.test(draw.concat(ole, ctl).join(''))) {
        outText = ensureRNamespace(outText)
      }

      function resolvable(block) {
        var re = /r:id="([^"]+)"/g
        var m
        while ((m = re.exec(block))) if (!outRels[m[1]]) return false
        return true
      }

      if (cf.length && outText.indexOf('<conditionalFormatting') < 0) {
        var cfText = cf.join('')
        outText = insertBefore(outText, cfText, [
          'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter',
          'rowBreaks', 'colBreaks', 'drawing', 'legacyDrawing', 'picture', 'oleObjects', 'controls',
          'tableParts', 'extLst',
        ])
        report.conditionalFormatting += cf.length
      }
      if (draw.length && outText.indexOf('<drawing') < 0) {
        var drawText = draw.filter(resolvable).join('')
        if (drawText) {
          outText = insertBefore(outText, drawText, ['legacyDrawing', 'picture', 'oleObjects', 'controls', 'tableParts', 'extLst'])
          report.drawing += draw.length
        }
      }
      if (ole.length && outText.indexOf('<oleObjects') < 0) {
        var oleText = ole.filter(resolvable).join('')
        if (oleText) {
          outText = insertBefore(outText, oleText, ['controls', 'tableParts', 'extLst'])
          report.oleObjects += ole.length
        }
      }
      if (ctl.length && outText.indexOf('<controls') < 0) {
        var ctlText = ctl.filter(resolvable).join('')
        if (ctlText) {
          outText = insertBefore(outText, ctlText, ['tableParts', 'extLst'])
          report.controls += ctl.length
        }
      }
      out.file(outPart, outText)
    }
    return report
  }

  async function mergeWorkbookExtras(out, src) {
    var srcWb = await readText(src, 'xl/workbook.xml')
    var outWb = await readText(out, 'xl/workbook.xml')
    if (!srcWb || !outWb) return { pivotCaches: 0 }
    var block = /<pivotCaches[\s\S]*?<\/pivotCaches>/.exec(srcWb)
    if (!block || outWb.indexOf('<pivotCaches') >= 0) return { pivotCaches: 0 }

    var srcRels = relEntries(await readText(src, 'xl/_rels/workbook.xml.rels'))
    var outRelsXml = await readText(out, 'xl/_rels/workbook.xml.rels')
    var outRels = relMapOf(outRelsXml)
    var added = 0
    var ok = true

    var re = /r:id="([^"]+)"/g
    var m
    while ((m = re.exec(block[0]))) {
      var id = m[1]
      if (outRels[id]) continue
      var entry = null
      for (var i = 0; i < srcRels.length; i++) if (srcRels[i].id === id) entry = srcRels[i]
      if (!entry || !out.files[resolveTarget('xl', entry.target)]) { ok = false; break }
      outRelsXml = addRelationship(outRelsXml, id, entry.type, entry.target)
      outRels[id] = entry.target
      added += 1
    }
    if (!ok) return { pivotCaches: 0 }
    if (added) out.file('xl/_rels/workbook.xml.rels', outRelsXml)
    outWb = insertBefore(
      outWb,
      block[0],
      ['calcPr', 'oleSize', 'customWorkbookViews', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'extLst'],
      '</workbook>',
    )
    out.file('xl/workbook.xml', outWb)
    return { pivotCaches: 1 }
  }

  async function preserveExtraParts(savedBytes, originalBytes) {
    if (!window.JSZip || !savedBytes || !originalBytes) return { bytes: savedBytes, preserved: false, parts: [] }
    try {
      var src = await window.JSZip.loadAsync(originalBytes)
      var out = await window.JSZip.loadAsync(savedBytes)
      var copied = await copyExtraParts(out, src)
      var sheet = await mergeSheetExtras(out, src)
      var book = await mergeWorkbookExtras(out, src)
      var touched =
        copied.length + sheet.conditionalFormatting + sheet.drawing + sheet.oleObjects + sheet.controls + book.pivotCaches
      if (!touched) return { bytes: savedBytes, preserved: false, parts: [], sheet: sheet }
      var bytes = await out.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
      var errors = await validate(bytes)
      if (errors.length) {
        console.warn('[vba-preserve] extra-part merge failed validation', errors.slice(0, 5))
        return { bytes: savedBytes, preserved: false, parts: [], sheet: sheet, errors: errors.slice(0, 5) }
      }
      return { bytes: bytes, preserved: true, parts: copied, sheet: sheet }
    } catch (e) {
      console.warn('[vba-preserve] could not restore non-cell parts', e)
      return { bytes: savedBytes, preserved: false, parts: [], error: String((e && e.message) || e) }
    }
  }

  async function preserveAll(savedBytes, originalBytes) {
    var macro = await preserveMacroParts(savedBytes, originalBytes)
    var base = macro && macro.preserved ? macro.bytes : savedBytes
    var extra = await preserveExtraParts(base, originalBytes)
    var bytes = extra && extra.preserved ? extra.bytes : base
    return {
      bytes: bytes,
      preserved: !!((macro && macro.preserved) || (extra && extra.preserved)),
      macro: macro,
      extras: extra,
    }
  }


  // --------------------------------------------- form-control buttons (VML)
  //
  // Form-control buttons live in vmlDrawing*.vml and point at a macro through
  // <x:FmlaMacro>. The grid engine never draws those shapes, so the app lists
  // them as buttons in the macro panel instead of pretending they are clickable
  // on the sheet.

  function stripTags(text) {
    return String(text)
      .replace(/<[^>]*>/g, ' ')
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/\s+/g, ' ')
      .trim()
  }

  function normalizeMacroTarget(raw) {
    var text = String(raw || '').trim()
    var bang = text.lastIndexOf('!')
    if (bang >= 0) text = text.slice(bang + 1)
    return text.replace(/^['"]|['"]$/g, '').trim()
  }

  async function readFormButtons(bytes) {
    var buttons = []
    if (!window.JSZip || !bytes) return buttons
    try {
      var zip = await window.JSZip.loadAsync(bytes)
      var names = Object.keys(zip.files).filter(function (n) {
        return /^xl\/drawings\/vmlDrawing.*\.vml$/i.test(n) && !zip.files[n].dir
      })
      for (var i = 0; i < names.length; i++) {
        var vml = await readText(zip, names[i])
        if (!vml) continue
        var re = /<v:shape\b[^>]*>[\s\S]*?<\/v:shape>/g
        var m
        while ((m = re.exec(vml))) {
          var block = m[0]
          var macroMatch = /<x:FmlaMacro>([\s\S]*?)<\/x:FmlaMacro>/.exec(block)
          if (!macroMatch) continue
          var idMatch = /id="([^"]+)"/.exec(block)
          var caption = stripTags((/<v:textbox[^>]*>([\s\S]*?)<\/v:textbox>/.exec(block) || [])[1] || '')
          buttons.push({
            id: idMatch ? idMatch[1] : '',
            caption: caption,
            macro: normalizeMacroTarget(macroMatch[1]),
            part: names[i],
          })
        }
      }
      return buttons.filter(function (b) { return !!b.macro })
    } catch (e) {
      console.warn('[vba-preserve] could not read form buttons', e)
      return []
    }
  }


  // --------------------------------------------------- ActiveX controls
  //
  // ActiveX controls are declared in the sheet (<controls><control name="..."/>).
  // Their VBA handlers follow the usual <name>_Click convention. The grid never
  // renders the controls themselves, so the panel lists them and runs the handler.

  async function readActiveXControls(bytes) {
    var controls = []
    if (!window.JSZip || !bytes) return controls
    try {
      var zip = await window.JSZip.loadAsync(bytes)
      var names = Object.keys(zip.files).filter(function (n) {
        return /^xl\/worksheets\/sheet\d+\.xml$/i.test(n) && !zip.files[n].dir
      })
      for (var i = 0; i < names.length; i++) {
        var text = await readText(zip, names[i])
        if (!text) continue
        var blocks = text.match(/<controls[\s\S]*?<\/controls>/g) || []
        for (var b = 0; b < blocks.length; b++) {
          var re = /<control\b[^>]*>/g
          var m
          while ((m = re.exec(blocks[b]))) {
            var nameMatch = /\bname="([^"]+)"/.exec(m[0])
            if (!nameMatch) continue
            var ids = /\bshapeId="([^"]+)"/.exec(m[0])
            controls.push({
              name: nameMatch[1],
              shapeId: ids ? ids[1] : '',
              sheetPart: names[i],
            })
          }
        }
      }
      return controls
    } catch (e) {
      console.warn('[vba-preserve] could not read ActiveX controls', e)
      return []
    }
  }

  window.__zenVba = {
    readActiveXControls: readActiveXControls,
    readFormButtons: readFormButtons,
    inspectMacros: inspectMacros,
    preserveAll: preserveAll,
    preserveExtraParts: preserveExtraParts,
    preserveMacroParts: preserveMacroParts,
    prepareEngineBytes: prepareEngineBytes,
    validate: validate,
  }
})()
