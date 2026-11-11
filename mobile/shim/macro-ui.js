/**
 * macro-ui.js — floating VBA panel for the Sheets view.
 *
 * Lists the procedures found in the workbook's vbaProject.bin and runs one on
 * the live sheet. Everything stays on the device: no external service, no
 * Excel process. Constructs the interpreter does not support are reported with
 * their keyword and source line instead of being half-executed.
 */
(function () {
  'use strict'

  if (window.__zenMacroUiInstalled) return
  window.__zenMacroUiInstalled = true

  var api = window.__genofficeMacros
  if (!api) return

  var state = { open: false, macros: null, busy: false, key: '', consented: Object.create(null) }

  var css = document.createElement('style')
  css.textContent = [
    '.tkn-macro-fab{position:fixed;right:14px;bottom:96px;z-index:9998;height:38px;padding:0 16px;border:0;border-radius:19px;',
    'background:linear-gradient(135deg,#4894FF,#0A2E96);color:#fff;font:600 13px/1 -apple-system,Roboto,sans-serif;',
    'box-shadow:0 6px 18px rgba(10,46,150,.34);display:flex;align-items:center;gap:8px;cursor:pointer}',
    '.tkn-macro-fab:active{transform:scale(.97)}',
    '.tkn-macro-card{position:fixed;right:12px;bottom:96px;left:12px;max-height:70vh;z-index:9999;display:flex;flex-direction:column;',
    'background:#fff;color:#1b1d21;border-radius:16px;box-shadow:0 18px 48px rgba(9,16,38,.28);overflow:hidden;',
    'font:400 13px/1.45 -apple-system,Roboto,sans-serif}',
    '.tkn-macro-head{display:flex;align-items:center;justify-content:space-between;padding:12px 14px;background:#0A2E96;color:#fff}',
    '.tkn-macro-head b{font-size:13px;letter-spacing:.2px}',
    '.tkn-macro-head span{opacity:.75;font-size:11px;display:block;font-weight:400;margin-top:2px}',
    '.tkn-macro-x{background:rgba(255,255,255,.16);border:0;color:#fff;width:26px;height:26px;border-radius:13px;font-size:15px;cursor:pointer}',
    '.tkn-macro-body{padding:10px 12px 12px;overflow:auto}',
    '.tkn-macro-note{font-size:11px;color:#6b7280;margin:0 0 8px}',
    '.tkn-macro-row{display:flex;align-items:center;gap:10px;padding:10px;border:1px solid #e6e8ec;border-radius:10px;margin-bottom:8px;cursor:pointer}',
    '.tkn-macro-row:active{background:#f3f6ff;border-color:#c7d7ff}',
    '.tkn-macro-tag{font-size:10px;font-weight:700;padding:2px 6px;border-radius:5px;background:#eaf1ff;color:#0A2E96;text-transform:uppercase}',
    '.tkn-macro-name{font-weight:600;flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}',
    '.tkn-macro-mod{font-size:10px;color:#8a9099}',
    '.tkn-macro-out{margin-top:10px;padding:9px 10px;border-radius:9px;background:#f5f7fa;font:400 12px/1.5 ui-monospace,Menlo,monospace;white-space:pre-wrap;word-break:break-word;max-height:180px;overflow:auto}',
    '.tkn-macro-out.err{background:#fdeceb;color:#a11d12}',
    '.tkn-macro-out.ok{background:#eaf7ef;color:#136c37}',
    '.tkn-macro-empty{color:#6b7280;font-size:12px;padding:6px 2px}',
    '[data-theme=dark] .tkn-macro-card{background:#1e1e1e;color:#e4e4e4}',
    '[data-theme=dark] .tkn-macro-row{border-color:#3a3a3a}',
    '[data-theme=dark] .tkn-macro-out{background:#2a2a2a}',
    '[data-theme=dark] .tkn-macro-out.ok{background:#20342a;color:#8fd8a9}',
    '[data-theme=dark] .tkn-macro-out.err{background:#3a2020;color:#ffb4ab}',
  ].join('')
  ;(document.head || document.documentElement).appendChild(css)

  var fab = document.createElement('button')
  fab.className = 'tkn-macro-fab'
  fab.type = 'button'
  fab.innerHTML = '<span style="font-size:15px">ƒ</span><span>Makro</span>'

  var card = null
  var listBox = null
  var outputBox = null

  function close() {
    state.open = false
    if (card && card.parentNode) card.parentNode.removeChild(card)
    card = null
    listBox = null
    outputBox = null
    fab.style.display = ''
  }

  function showOutput(text, kind) {
    if (!outputBox) return
    outputBox.className = 'tkn-macro-out' + (kind ? ' ' + kind : '')
    outputBox.textContent = text
  }

  function askConfirm(label, onYes) {
    var layer = document.createElement('div')
    layer.style.cssText =
      'position:fixed;inset:0;z-index:10000;background:rgba(9,16,38,.45);display:flex;align-items:center;justify-content:center;padding:24px'

    var box = document.createElement('div')
    box.style.cssText =
      'background:#fff;color:#1b1d21;border-radius:14px;padding:18px 16px;max-width:320px;width:100%;' +
      'box-shadow:0 18px 48px rgba(9,16,38,.35);font:400 13px/1.5 -apple-system,Roboto,sans-serif'

    var title = document.createElement('div')
    title.style.cssText = 'font-weight:700;font-size:14px;margin-bottom:6px'
    title.textContent = 'Jalankan makro dari file ini?'

    var body = document.createElement('div')
    body.style.cssText = 'color:#4b5563;margin-bottom:14px;white-space:pre-wrap'
    body.textContent = label + '\n\nMakro bisa mengubah data di lembar ini. Hanya jalankan file yang Anda percaya.'

    var row = document.createElement('div')
    row.style.cssText = 'display:flex;gap:8px;justify-content:flex-end'

    var cancel = document.createElement('button')
    cancel.type = 'button'
    cancel.textContent = 'Batal'
    cancel.style.cssText =
      'border:1px solid #d7dbe0;background:#fff;color:#1b1d21;border-radius:9px;padding:8px 14px;font-weight:600'

    var go = document.createElement('button')
    go.type = 'button'
    go.textContent = 'Jalankan'
    go.style.cssText =
      'border:0;background:#0A2E96;color:#fff;border-radius:9px;padding:8px 16px;font-weight:600'

    function closeLayer() {
      if (layer.parentNode) layer.parentNode.removeChild(layer)
    }
    cancel.addEventListener('click', closeLayer)
    go.addEventListener('click', function () {
      closeLayer()
      onYes()
    })
    layer.addEventListener('click', function (event) {
      if (event.target === layer) closeLayer()
    })

    row.appendChild(cancel)
    row.appendChild(go)
    box.appendChild(title)
    box.appendChild(body)
    box.appendChild(row)
    layer.appendChild(box)
    document.body.appendChild(layer)
  }

  function requestRun(macro, label) {
    if (state.busy) return
    if (state.consented[state.key || 'default']) {
      runOne(macro)
      return
    }
    askConfirm(label || (macro.name + '()'), function () {
      state.consented[state.key || 'default'] = true
      runOne(macro)
    })
  }

  function autoEntryMacro(macros) {
    if (!macros || !macros.length) return null
    var wanted = ['workbook_open', 'auto_open']
    for (var i = 0; i < wanted.length; i++) {
      for (var k = 0; k < macros.length; k++) {
        if (String(macros[k].name).toLowerCase() === wanted[i]) return macros[k]
      }
    }
    return null
  }

  function renderList(result) {
    if (!listBox) return
    listBox.textContent = ''
    if (result && result.name) state.key = result.name
    if (result && result.ok) {
      var auto = autoEntryMacro(result.macros)
      if (auto && state.key && !state.autoRan) state.autoRan = {}
      if (auto && state.key && !state.autoRan[state.key]) {
        state.autoRan[state.key] = true
        setTimeout(function () {
          requestRun(auto, auto.name + '() berjalan otomatis saat file dibuka')
        }, 400)
      }
    }
    if (!result || !result.ok) {
      var err = document.createElement('div')
      err.className = 'tkn-macro-empty'
      err.textContent = 'Tidak bisa membaca makro: ' + ((result && result.error) || 'unknown')
      listBox.appendChild(err)
      showOutput('', '')
      return
    }
    if (!result.macros || !result.macros.length) {
      var note0 = document.createElement('div')
      note0.className = 'tkn-macro-empty'
      note0.textContent = 'Workbook ini tidak punya prosedur VBA yang bisa dibaca.'
      listBox.appendChild(note0)
      showOutput('', '')
      return
    }
    result.macros.forEach(function (macro) {
      var row = document.createElement('div')
      row.className = 'tkn-macro-row'

      var tag = document.createElement('span')
      tag.className = 'tkn-macro-tag'
      tag.textContent = macro.kind === 'function' ? 'Fn' : 'Sub'

      var name = document.createElement('div')
      name.className = 'tkn-macro-name'
      name.textContent = macro.name + (macro.params.length ? '(' + macro.params.join(', ') + ')' : '()')

      var mod = document.createElement('div')
      mod.className = 'tkn-macro-mod'
      mod.textContent = macro.module

      row.appendChild(tag)
      row.appendChild(name)
      row.appendChild(mod)
      row.addEventListener('click', function () {
        requestRun(macro)
      })
      listBox.appendChild(row)
    })
    if (result.buttons && result.buttons.length) {
      var btnTitle = document.createElement('p')
      btnTitle.className = 'tkn-macro-note'
      btnTitle.style.marginTop = '12px'
      btnTitle.textContent = 'Tombol form di file ini'
      listBox.appendChild(btnTitle)

      result.buttons.forEach(function (button) {
        var row = document.createElement('div')
        row.className = 'tkn-macro-row'

        var tag = document.createElement('span')
        tag.className = 'tkn-macro-tag'
        tag.textContent = 'Btn'

        var name = document.createElement('div')
        name.className = 'tkn-macro-name'
        name.textContent = button.caption || 'Tombol'

        var mod = document.createElement('div')
        mod.className = 'tkn-macro-mod'
        mod.textContent = button.macro

        row.appendChild(tag)
        row.appendChild(name)
        row.appendChild(mod)
        row.addEventListener('click', function () {
          requestRun({ name: button.macro, kind: 'sub' }, 'Tombol "' + (button.caption || button.macro) + '" akan menjalankan ' + button.macro + '()')
        })
        listBox.appendChild(row)
      })
    }


    if (result.activeX && result.activeX.length) {
      var axTitle = document.createElement('p')
      axTitle.className = 'tkn-macro-note'
      axTitle.style.marginTop = '12px'
      axTitle.textContent = 'Kontrol ActiveX'
      listBox.appendChild(axTitle)

      result.activeX.forEach(function (control) {
        var row = document.createElement('div')
        row.className = 'tkn-macro-row'

        var tag = document.createElement('span')
        tag.className = 'tkn-macro-tag'
        tag.textContent = 'AX'

        var name = document.createElement('div')
        name.className = 'tkn-macro-name'
        name.textContent = control.name

        var mod = document.createElement('div')
        mod.className = 'tkn-macro-mod'
        mod.textContent = control.hasHandler ? control.handler + '()' : 'tanpa handler'

        row.appendChild(tag)
        row.appendChild(name)
        row.appendChild(mod)
        if (control.hasHandler) {
          row.addEventListener('click', function () {
            requestRun(
              { name: control.handler, kind: 'sub' },
              'Kontrol ' + control.name + ' akan menjalankan ' + control.handler + '()',
            )
          })
        } else {
          row.style.opacity = '0.55'
        }
        listBox.appendChild(row)
      })
    }

    showOutput(
      'Ketuk nama makro untuk menjalankannya di lembar yang terbuka.\n' +
        'Berjalan lokal di perangkat · tanpa mesin Excel · hasil langsung masuk ke sheet.\n' +
        'Sekali Anda menyetujui file ini, makro lain di file yang sama tidak bertanya lagi.',
    )
  }

  async function runOne(macro) {
    if (state.busy) return
    state.busy = true
    showOutput('Menjalankan ' + macro.name + '…')
    try {
      var result = await api.run(macro.name)
      state.busy = false
      if (!result.ok) {
        var lines = ['⚠ ' + result.error]
        if (result.line) lines.push('baris sumber: ' + result.line)
        if (result.changes && result.changes.length) {
          lines.push('sel yang sempat berubah: ' + result.changes.length)
        }
        showOutput(lines.join('\n'), 'err')
        return
      }
      var okLines = []
      if (result.output && result.output.length) okLines.push(result.output.join('\n'))
      okLines.push('✓ ' + result.changes.length + ' sel diperbarui · ' + result.steps + ' langkah')
      okLines.push('Simpan file untuk menyimpan hasilnya.')
      showOutput(okLines.join('\n'), 'ok')
    } catch (e) {
      state.busy = false
      showOutput('⚠ ' + String((e && e.message) || e), 'err')
    }
  }

  async function open() {
    state.open = true
    fab.style.display = 'none'

    card = document.createElement('div')
    card.className = 'tkn-macro-card'

    var head = document.createElement('div')
    head.className = 'tkn-macro-head'
    var title = document.createElement('div')
    title.innerHTML = '<b>Makro VBA</b><span>Dijalankan lokal · tanpa mesin Excel</span>'
    var closeBtn = document.createElement('button')
    closeBtn.className = 'tkn-macro-x'
    closeBtn.type = 'button'
    closeBtn.textContent = '×'
    closeBtn.addEventListener('click', close)
    head.appendChild(title)
    head.appendChild(closeBtn)

    var body = document.createElement('div')
    body.className = 'tkn-macro-body'

    var note = document.createElement('p')
    note.className = 'tkn-macro-note'
    note.textContent =
      'Sub/Fungsi di bawah dibaca dari vbaProject.bin. Perulangan, kondisi, Range/Cells dan MsgBox didukung; UserForm, ActiveX, chart dan panggilan sistem tidak.'

    listBox = document.createElement('div')
    outputBox = document.createElement('div')
    outputBox.className = 'tkn-macro-out'

    body.appendChild(note)
    body.appendChild(listBox)
    body.appendChild(outputBox)
    card.appendChild(head)
    card.appendChild(body)
    document.body.appendChild(card)

    showOutput('Membaca proyek VBA…')
    try {
      var list = await api.list()
      renderList(list)
    } catch (e) {
      renderList({ ok: false, error: String((e && e.message) || e) })
    }
  }

  fab.style.display = 'none'

  async function refreshAvailability() {
    if (state.open) return
    try {
      var list = await api.list()
      fab.style.display = list && list.ok && list.macros && list.macros.length ? '' : 'none'
    } catch (e) {
      fab.style.display = 'none'
    }
  }
  setTimeout(refreshAvailability, 900)
  setInterval(refreshAvailability, 4000)

  fab.addEventListener('click', function () {
    if (state.open) close()
    else open()
  })
  document.body.appendChild(fab)

  window.__zenMacroUi = { open: open, close: close }
})()
