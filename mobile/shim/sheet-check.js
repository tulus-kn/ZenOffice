/**
 * sheet-check.js — "Periksa" panel for the Sheets view.
 *
 * Applies the machine-checkable rules from the bundled xlsx skill:
 *   - Zero formula errors  : find #REF!, #DIV/0!, #NAME?, #VALUE!, #N/A cells
 *   - Formulas, never hardcoded results : flags formulas that bake numbers in
 *   - Recalculation needed : formulas with no cached value
 *
 * Everything runs locally on the SheetJS workbook already held by the session.
 * Font auditing is not included: it needs cell styles, which the sheet engine
 * does not keep for every cell.
 */
(function () {
  'use strict'

  var ERR_TYPES = ['#REF!', '#DIV/0!', '#NAME?', '#VALUE!', '#N/A', '#NUM!', '#NULL!']

  function inspect(wb, names) {
    var XLSX = window.XLSX
    var result = {
      ok: true,
      sheets: [],
      errors: [],
      hardcoded: [],
      formulaCells: 0,
      uncached: 0,
      scannedCells: 0,
    }
    if (!wb || !XLSX) return { ok: false, error: 'Workbook tidak tersedia' }
    var sheetNames = names && names.length ? names : wb.SheetNames || []
    for (var s = 0; s < sheetNames.length; s++) {
      var name = sheetNames[s]
      var ws = wb.Sheets[name]
      if (!ws || !ws['!ref']) {
        result.sheets.push({ name: name, cells: 0, empty: true })
        continue
      }
      var range = XLSX.utils.decode_range(ws['!ref'])
      var cells = 0
      for (var R = range.s.r; R <= range.e.r; R++) {
        for (var C = range.s.c; C <= range.e.c; C++) {
          var addr = XLSX.utils.encode_cell({ r: R, c: C })
          var cell = ws[addr]
          if (!cell) continue
          cells++
          result.scannedCells++
          if (cell.t === 'e') {
            if (result.errors.length < 200) {
              result.errors.push({
                sheet: name,
                cell: addr,
                error: cell.w || cell.v || '#ERROR',
                formula: cell.f || null,
              })
            } else {
              result.errors.push(null)
            }
          }
          if (cell.f) {
            result.formulaCells++
            if (cell.v === undefined && cell.w === undefined) result.uncached++
            // strip strings and cell references; anything numeric left is a baked constant
            var expr = String(cell.f)
              .replace(/"[^"]*"/g, '""')
              .replace(/\$?[A-Za-z]{1,3}\$?\d+/g, 'R')
            if (/[0-9]/.test(expr) && result.hardcoded.length < 100) {
              result.hardcoded.push({ sheet: name, cell: addr, formula: cell.f })
            }
          }
        }
      }
      result.sheets.push({ name: name, cells: cells, empty: false })
    }
    result.errorCount = result.errors.filter(Boolean).length
    result.errorsTruncated = result.errors.length !== result.errorCount
    result.errors = result.errors.filter(Boolean)
    return result
  }

  function fmt(n) { return String(n) }

  function mountUi() {
    if (window.__tknSheetCheckMounted) return
    if (!/\/sheets\//.test(window.location.pathname)) return
    window.__tknSheetCheckMounted = true

    var css = document.createElement('style')
    css.textContent = [
      '.tkn-check-fab{position:fixed;right:14px;bottom:142px;z-index:9998;height:38px;padding:0 16px;border:0;border-radius:19px;',
      'background:#0f172a;color:#fff;font:600 13px/1 -apple-system,Roboto,sans-serif;display:flex;align-items:center;gap:8px;',
      'box-shadow:0 6px 18px rgba(15,23,42,.32);cursor:pointer}',
      '.tkn-check-card{position:fixed;right:12px;bottom:142px;left:12px;max-height:72vh;z-index:9999;display:flex;flex-direction:column;',
      'background:#fff;color:#1b1d21;border-radius:16px;box-shadow:0 18px 48px rgba(9,16,38,.28);overflow:hidden;',
      'font:400 13px/1.45 -apple-system,Roboto,sans-serif}',
      '.tkn-check-head{display:flex;align-items:center;justify-content:space-between;padding:12px 14px;background:#0f172a;color:#fff}',
      '.tkn-check-head b{font-size:13px}.tkn-check-head span{opacity:.72;font-size:11px;display:block;font-weight:400;margin-top:2px}',
      '.tkn-check-x{background:rgba(255,255,255,.16);border:0;color:#fff;width:26px;height:26px;border-radius:13px;font-size:15px}',
      '.tkn-check-body{padding:10px 12px 12px;overflow:auto}',
      '.tkn-check-row{display:flex;align-items:baseline;gap:10px;padding:9px 10px;border:1px solid #e6e8ec;border-radius:10px;margin-bottom:8px}',
      '.tkn-check-row.bad{border-color:#f2c7c3;background:#fdf2f1}',
      '.tkn-check-row.ok{border-color:#cfe7d6;background:#f2f9f4}',
      '.tkn-check-k{font-weight:600;flex:1}',
      '.tkn-check-v{font-weight:700;font-variant-numeric:tabular-nums}',
      '.tkn-check-item{font:400 11px/1.5 ui-monospace,Menlo,monospace;color:#4b5563;padding:6px 9px;background:#f6f7f9;border-radius:8px;margin-bottom:6px;word-break:break-word}',
      '.tkn-check-note{font-size:11px;color:#6b7280;margin:8px 0 0}',
      '[data-theme=dark] .tkn-check-card{background:#1e1e1e;color:#e4e4e4}',
      '[data-theme=dark] .tkn-check-row{border-color:#3a3a3a;background:#232323}',
      '[data-theme=dark] .tkn-check-item{background:#2a2a2a;color:#c8c8c8}',
    ].join('')
    ;(document.head || document.documentElement).appendChild(css)

    var fab = document.createElement('button')
    fab.type = 'button'
    fab.className = 'tkn-check-fab'
    fab.innerHTML = '<span style="font-size:14px">✓</span><span>Periksa</span>'

    var card = null

    function close() {
      if (card && card.parentNode) card.parentNode.removeChild(card)
      card = null
      fab.style.display = ''
    }

    function line(label, value, kind) {
      return '<div class="tkn-check-row ' + (kind || '') + '"><span class="tkn-check-k">' + label +
        '</span><span class="tkn-check-v">' + value + '</span></div>'
    }

    function render(data) {
      var body = card.querySelector('.tkn-check-body')
      if (!data || data.ok === false) {
        body.innerHTML = '<div class="tkn-check-item">' + ((data && data.error) || 'Tidak ada workbook terbuka') + '</div>'
        return
      }
      var html = ''
      html += line('Error rumus', fmt(data.errorCount), data.errorCount ? 'bad' : 'ok')
      html += line('Sel berumus', fmt(data.formulaCells), '')
      html += line('Rumus belum dihitung', fmt(data.uncached), data.uncached ? 'bad' : '')
      html += line('Rumus dengan konstanta', fmt(data.hardcoded.length), data.hardcoded.length ? 'bad' : '')
      html += '<div class="tkn-check-note">' + data.scannedCells + ' sel dipindai di ' + data.sheets.length + ' sheet.</div>'
      if (data.errorCount) {
        html += '<div class="tkn-check-note">Lokasi error (' + (data.errorsTruncated ? '200 pertama' : 'semua') + '):</div>'
        data.errors.slice(0, 40).forEach(function (e) {
          html += '<div class="tkn-check-item">' + e.sheet + '!' + e.cell + ' → ' + e.error +
            (e.formula ? '   [' + e.formula + ']' : '') + '</div>'
        })
        if (data.errorCount > 40) html += '<div class="tkn-check-note">… ' + (data.errorCount - 40) + ' lainnya</div>'
      }
      if (data.hardcoded.length) {
        html += '<div class="tkn-check-note">Angka yang ditanam di rumus (contoh):</div>'
        data.hardcoded.slice(0, 20).forEach(function (h) {
          html += '<div class="tkn-check-item">' + h.sheet + '!' + h.cell + ' = ' + h.formula + '</div>'
        })
      }
      html += '<div class="tkn-check-note">Aturan dari skill xlsx: “Zero formula errors”, “Use formulas, never hardcoded results”. Font tidak diperiksa (butuh style per sel).</div>'
      body.innerHTML = html
    }

    async function open() {
      var api = window.__genofficeCheck
      if (!api) return
      fab.style.display = 'none'
      card = document.createElement('div')
      card.className = 'tkn-check-card'
      card.innerHTML =
        '<div class="tkn-check-head"><div><b>Periksa spreadsheet</b><span>skill xlsx · berjalan lokal</span></div>' +
        '<button class="tkn-check-x" type="button">×</button></div>' +
        '<div class="tkn-check-body"><div class="tkn-check-item">Memeriksa…</div></div>'
      document.body.appendChild(card)
      card.querySelector('.tkn-check-x').addEventListener('click', close)
      var data = null
      try {
        data = await api.run()
      } catch (e) {
        data = { ok: false, error: String((e && e.message) || e) }
      }
      render(data)
    }

    fab.addEventListener('click', open)
    ;(document.body || document.documentElement).appendChild(fab)
  }

  window.__tknSheetCheck = { inspect: inspect, mountUi: mountUi }
  mountUi()
})()
