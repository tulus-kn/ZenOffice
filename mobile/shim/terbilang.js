// TERBILANG — angka menjadi huruf (bahasa Indonesia), dijalankan di perangkat.
// Tidak memakai layanan luar dan tidak butuh rumus tambahan di mesin grid:
// sel yang berisi angka (atau rumus =TERBILANG(...)) diubah menjadi teks.
(function () {
  var ONES = ['', 'satu', 'dua', 'tiga', 'empat', 'lima', 'enam', 'tujuh', 'delapan', 'sembilan', 'sepuluh', 'sebelas']

  function under100(n) {
    if (n < 12) return ONES[n]
    if (n < 20) return ONES[n - 10] + ' belas'
    var tens = Math.floor(n / 10)
    var rest = n % 10
    return ONES[tens] + ' puluh' + (rest ? ' ' + ONES[rest] : '')
  }

  function under1000(n) {
    if (n < 100) return under100(n)
    var hundreds = Math.floor(n / 100)
    var rest = n % 100
    var head = hundreds === 1 ? 'seratus' : ONES[hundreds] + ' ratus'
    return head + (rest ? ' ' + under100(rest) : '')
  }

  var SCALES = [
    [1e12, 'triliun'],
    [1e9, 'miliar'],
    [1e6, 'juta'],
    [1e3, 'ribu'],
  ]

  function integerWords(n) {
    if (n === 0) return 'nol'
    var parts = []
    var rest = n
    for (var i = 0; i < SCALES.length; i++) {
      var unit = SCALES[i][0]
      var label = SCALES[i][1]
      var count = Math.floor(rest / unit)
      if (!count) continue
      rest = rest % unit
      if (count === 1 && label === 'ribu') parts.push('seribu')
      else parts.push(integerWords(count) + ' ' + label)
    }
    if (rest) parts.push(under1000(rest))
    return parts.join(' ')
  }

  // numbers only, optional decimals: 1234.56 -> "... koma lima enam"
  function words(value, options) {
    var opts = options || {}
    var ready = String(value === null || value === undefined ? '' : value).trim()
    ready = ready.replace(/^(rp|idr)\.?\s*/i, '').replace(/\s*rupiah\s*$/i, '').replace(/\s/g, '')
    if (!ready) return ''
    // Indonesian/Excel grouping: 1.234.567,89 — strip thousand dots, comma is decimal.
    // Excel/Univer menyimpan angka sangat besar sebagai teks ilmiah ("1E+12").
    if (/^[-+]?\d+(?:[.,]\d+)?[eE][-+]?\d+$/.test(ready)) {
      var big = Number(ready.replace(',', '.'))
      if (isFinite(big)) ready = big.toFixed(0)
    }
    if (/^[-+]?\d{1,3}(\.\d{3})+(,\d+)?$/.test(ready)) {
      ready = ready.replace(/\./g, '').replace(',', '.')
    } else if (/^[-+]?\d+(,\d+)$/.test(ready)) {
      ready = ready.replace(',', '.')
    } else {
      ready = ready.replace(/,/g, '')
    }
    var negative = false
    if (ready.charAt(0) === '-') {
      negative = true
      ready = ready.slice(1)
    }
    if (!/^[0-9]+(\.[0-9]+)?$/.test(ready)) return ''
    var whole = ready.split('.')[0].replace(/^0+/, '') || '0'
    var decimals = ready.split('.')[1] || ''
    var text = integerWords(Number(whole))
    if (decimals) {
      var digits = []
      for (var i = 0; i < decimals.length; i++) {
        var d = Number(decimals.charAt(i))
        digits.push(ONES[d] === '' ? 'nol' : ONES[d])
      }
      text = text + ' koma ' + digits.join(' ')
    }
    if (negative) text = 'minus ' + text
    if (opts.rupiah) text = text + ' rupiah'
    return applyCase(text, opts.case || 'title')
  }

  function applyCase(text, mode) {
    if (mode === 'upper') return text.toUpperCase()
    if (mode === 'lower') return text
    return text.replace(/\S+/g, function (word) { return word.charAt(0).toUpperCase() + word.slice(1) })
  }

  function activeSheet() {
    var api = window.__univerAPI
    var wb = api && api.getActiveWorkbook && api.getActiveWorkbook()
    return wb ? wb.getActiveSheet() : null
  }

  function columnName(col) {
    var letters = ''
    var n = col + 1
    while (n > 0) {
      var rem = (n - 1) % 26
      letters = String.fromCharCode(65 + rem) + letters
      n = Math.floor((n - 1) / 26)
    }
    return letters
  }

  // Rewrites cells that hold =TERBILANG(...) with the computed words.
  function convertFormulas(mode) {
    var sheet = activeSheet()
    if (!sheet) return { ok: false, error: 'Spreadsheet belum siap' }
    var used = sheet.getRange ? sheet.getMaxRows && sheet.getMaxColumns ? null : null : null
    var values = sheet.getDataRange ? null : null
    var maxRow = sheet.getMaxRows ? sheet.getMaxRows() : 0
    var maxCol = sheet.getMaxColumns ? sheet.getMaxColumns() : 0
    var count = 0
    var placeRight = mode === 'right'
    var last = { row: Math.min(maxRow, 1000), col: Math.min(maxCol, 60) }
    for (var r = 1; r <= last.row; r++) {
      for (var c = 1; c <= last.col; c++) {
        var range = sheet.getRange(columnName(c - 1) + r)
        var formula = range.getFormula ? range.getFormula() : ''
        if (!formula) continue
        var source = String(formula).trim()
        if (!/terbilang\(/i.test(source)) continue
        // Replace every TERBILANG(arg) with its quoted text, then keep only a
        // pure string concatenation (e.g. TERBILANG(D7)&" Rupiah").
        var evaluated = source.replace(/terbilang\(\s*([^()]+?)\s*\)/gi, function (_all, argument) {
          var raw = argument
          if (/^[A-Za-z]{1,3}[0-9]{1,7}$/.test(argument.trim())) raw = sheet.getRange(argument.trim()).getValue()
          var text = words(raw, {})
          return text ? '"' + text.replace(/"/g, '""') + '"' : '"' + '"'
        })
        if (!/^=\s*"([^"]|"")*"(\s*&\s*"([^"]|"")*")*\s*$/.test(evaluated)) continue
        var plain = evaluated.replace(/^=\s*/, '').split('&').map(function (piece) {
          var body = piece.trim().replace(/^"|"$/g, '')
          return body.replace(/""/g, '"')
        }).join('')
        if (!plain) continue
        if (placeRight) sheet.getRange(columnName(c) + r).setValue(plain)
        else range.setValue(plain)
        count++
      }
    }
    return { ok: true, count: count }
  }

  // Writes the words of the selected number(s) into the column on the right.
  function convertSelection() {
    var sheet = activeSheet()
    if (!sheet) return { ok: false, error: 'Spreadsheet belum siap' }
    var selection = sheet.getSelection && sheet.getSelection()
    var range = selection && selection.getActiveRange && selection.getActiveRange()
    if (!range) return { ok: false, error: 'Pilih sel angka lebih dulu' }
    var start = range.getRow ? range.getRow() : 0
    var end = range.getLastRow ? range.getLastRow() : start
    var col = range.getColumn ? range.getColumn() : 0
    var count = 0
    for (var r = start; r <= end; r++) {
      var source = sheet.getRange(columnName(col) + (r + 1)).getValue()
      var text = words(source, {})
      if (!text) continue
      sheet.getRange(columnName(col + 1) + (r + 1)).setValue(text)
      count++
    }
    return { ok: true, count: count }
  }

  // Counts cells that still hold =TERBILANG(...) so the UI can hint the user.
  function countPending() {
    var sheet = activeSheet()
    if (!sheet) return 0
    var maxRow = Math.min(sheet.getMaxRows ? sheet.getMaxRows() : 0, 1000)
    var maxCol = Math.min(sheet.getMaxColumns ? sheet.getMaxColumns() : 0, 60)
    var pending = 0
    for (var r = 1; r <= maxRow; r++) {
      for (var c = 1; c <= maxCol; c++) {
        var formula = sheet.getRange(columnName(c - 1) + r).getFormula ?
          sheet.getRange(columnName(c - 1) + r).getFormula() : ''
        if (formula && /terbilang\(/i.test(String(formula)) && String(formula).charAt(0) === '=') pending++
      }
    }
    return pending
  }

  function toast(message) {
    var box = document.createElement('div')
    box.textContent = message
    box.style.cssText =
      'position:fixed;left:50%;bottom:120px;transform:translateX(-50%);z-index:9999;background:#0b1526;color:#fff;' +
      'padding:10px 14px;border-radius:10px;font:500 13px/1.4 -apple-system,Roboto,sans-serif;max-width:86vw;text-align:center'
    document.body.appendChild(box)
    setTimeout(function () { if (box.parentNode) box.parentNode.removeChild(box) }, 3200)
  }

  var MODE_STORE = 'tkn_terbilang_mode'

  function savedMode() {
    try { return localStorage.getItem(MODE_STORE) || '' } catch (e) { return '' }
  }

  function rememberMode(mode) {
    try { localStorage.setItem(MODE_STORE, mode) } catch (e) {}
  }

  function forgetMode() {
    try { localStorage.removeItem(MODE_STORE) } catch (e) {}
  }

  function runMode(mode) {
    var result = convertFormulas(mode)
    if (!result.ok) { toast(result.error); return }
    if (result.count) {
      toast(mode === 'right'
        ? result.count + ' sel: hasil ditulis di kolom kanan, rumus tetap'
        : result.count + ' sel rumus TERBILANG diisi')
      return
    }
    var sel = convertSelection()
    if (sel.ok && sel.count) toast(sel.count + ' sel angka diubah menjadi huruf (kolom kanan)')
    else toast('Tidak ada =TERBILANG(...) maupun angka terpilih. Pilih sel angka, lalu ketuk lagi.')
  }

  function showModeDialog(pending) {
    var overlay = document.createElement('div')
    overlay.style.cssText = 'position:fixed;inset:0;z-index:10002;background:rgba(9,16,38,.45);display:flex;align-items:center;justify-content:center;padding:18px'
    var card = document.createElement('div')
    card.style.cssText = 'background:#fff;border-radius:14px;padding:16px;width:100%;max-width:420px;box-sizing:border-box;' +
      'font:400 13px/1.5 -apple-system,Roboto,sans-serif;color:#111827;box-shadow:0 16px 40px rgba(9,16,38,.3)'
    card.innerHTML = '<div style="font:700 15px/1.2 inherit;margin-bottom:4px">Terbilang — ' + pending +
      ' sel rumus</div><div style="color:#6b7280;font-size:12px;margin-bottom:12px">Pilih cara mengisi hasil.</div>'

    var state = { remember: false }
    var remember = document.createElement('input')
    remember.type = 'checkbox'
    remember.addEventListener('change', function () { state.remember = !!remember.checked })

    function option(label, detail, mode, primary) {
      var row = document.createElement('button')
      row.type = 'button'
      row.style.cssText = 'width:100%;text-align:left;border:1px solid ' + (primary ? '#1d4ed8' : '#e1e5ec') +
        ';background:' + (primary ? '#eff4ff' : '#fff') + ';border-radius:10px;padding:11px 12px;margin-bottom:8px;font:inherit;color:inherit'
      row.innerHTML = '<div style="font-weight:600">' + label + '</div><div style="color:#6b7280;font-size:11.5px;margin-top:2px">' + detail + '</div>'
      row.addEventListener('click', function () {
        var remembered = !!remember.checked || state.remember
        if (remembered) rememberMode(mode)
        close()
        runMode(mode)
      })
      return row
    }

    card.appendChild(option('Ganti di sel (disarankan)', 'Rumus diganti hasil teks. Cocok untuk kwitansi/laporan seperti file Anda.', 'replace', true))
    card.appendChild(option('Tulis di kolom kanan', 'Rumus =TERBILANG(...) tetap utuh, hasil ditulis di sel sebelah kanan.', 'right', false))

    var rememberRow = document.createElement('label')
    rememberRow.style.cssText = 'display:flex;gap:7px;align-items:center;color:#374151;font-size:12px;margin:2px 0 10px'
    rememberRow.appendChild(remember)
    rememberRow.appendChild(document.createTextNode('Ingat pilihan saya (bisa diubah nanti di Tentang)'))

    var closeBtn = document.createElement('button')
    closeBtn.type = 'button'
    closeBtn.textContent = 'Batal'
    closeBtn.style.cssText = 'width:100%;border:0;background:#f1f3f6;color:#111827;border-radius:10px;padding:11px;font:600 13px/1 inherit'
    closeBtn.addEventListener('click', close)

    card.appendChild(rememberRow)
    card.appendChild(closeBtn)
    overlay.appendChild(card)
    overlay.addEventListener('click', function (event) { if (event.target === overlay) close() })
    document.body.appendChild(overlay)

    function close() { if (overlay.parentNode) overlay.parentNode.removeChild(overlay) }
  }

  function mountUi() {
    if (!/\/sheets\//.test(window.location.pathname)) return
    if (window.__tknTerbilangMounted) return
    window.__tknTerbilangMounted = true
    var button = document.createElement('button')
    button.type = 'button'
    button.textContent = 'Terbilang'
    button.style.cssText =
      'position:fixed;right:16px;bottom:188px;z-index:9000;border:0;border-radius:12px;padding:11px 14px;' +
      'background:#0f766e;color:#fff;font:600 13px/1 -apple-system,Roboto,sans-serif;box-shadow:0 8px 22px rgba(9,16,38,.28)'
    button.addEventListener('click', function () {
      var pending = countPending()
      if (pending > 0) {
        var mode = savedMode()
        if (mode) runMode(mode)
        else showModeDialog(pending)
        return
      }
      var sel = convertSelection()
      if (sel.ok && sel.count) toast(sel.count + ' sel angka diubah menjadi huruf (kolom kanan)')
      else toast('Tidak ada =TERBILANG(...) maupun angka terpilih. Pilih sel angka, lalu ketuk lagi.')
    })
    document.body.appendChild(button)
  }

  window.__tknTerbilang = {
    words: words,
    countPending: countPending,
    forgetMode: forgetMode,
    savedMode: savedMode,
    convertFormulas: convertFormulas,
    convertSelection: convertSelection,
    mountUi: mountUi,
  }
})()
