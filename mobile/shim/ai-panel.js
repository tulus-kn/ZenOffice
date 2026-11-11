// Panel AI — memakai kunci API milik pengguna sendiri (Google AI Studio /
// Gemini Flash). Kunci disimpan hanya di localStorage perangkat ini.
(function () {
  var KEY_STORE = 'tkn_ai_key'
  var MODEL_STORE = 'tkn_ai_model'
  var DEFAULT_MODEL = 'gemini-3.8-flash'

  function readStore(name, fallback) {
    try {
      var value = localStorage.getItem(name)
      return value === null ? fallback : value
    } catch (e) { return fallback }
  }

  function apiKey() { return readStore(KEY_STORE, '') }
  // Models get retired by Google; an old saved name is upgraded automatically.
  var RETIRED = ['gemini-2.0-flash', 'gemini-2.0-flash-001', 'gemini-1.5-flash']
  function modelName() {
    var stored = readStore(MODEL_STORE, DEFAULT_MODEL) || DEFAULT_MODEL
    if (RETIRED.indexOf(stored) >= 0) {
      try { localStorage.setItem(MODEL_STORE, DEFAULT_MODEL) } catch (e) {}
      return DEFAULT_MODEL
    }
    return stored
  }

  function saveKey(key, model) {
    try {
      localStorage.setItem(KEY_STORE, String(key || '').trim())
      if (model) localStorage.setItem(MODEL_STORE, String(model).trim())
      return true
    } catch (e) { return false }
  }

  function clearKey() {
    try {
      localStorage.removeItem(KEY_STORE)
      return true
    } catch (e) { return false }
  }

  function sheetContext() {
    var api = window.__univerAPI
    var wb = api && api.getActiveWorkbook && api.getActiveWorkbook()
    if (!wb) return ''
    var sheet = wb.getActiveSheet()
    var selection = sheet.getSelection && sheet.getSelection()
    var range = selection && selection.getActiveRange && selection.getActiveRange()
    if (!range || !range.getValues) return ''
    var values = range.getValues()
    if (!values || !values.length) return ''
    var rows = values.slice(0, 30).map(function (row) {
      return row.slice(0, 12).map(function (cell) {
        if (cell === null || cell === undefined) return ''
        return String(cell).slice(0, 40)
      }).join('\t')
    })
    return 'Data terpilih (tab-separated, maksimal 30 baris):\n' + rows.join('\n')
  }

  // POST JSON; bila preflight CORS ditolak, ulangi sebagai text/plain
  // (permintaan sederhana tanpa preflight) — Gemini tetap membaca body JSON.
  function postJson(url, body) {
    var payload = JSON.stringify(body)
    function once(contentType) {
      return fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': contentType },
        body: payload,
      })
    }
    return once('application/json').then(function (response) {
      return response
    }).catch(function (firstError) {
      return once('text/plain;charset=UTF-8').catch(function () { throw firstError })
    })
  }

  function ask(prompt, withContext) {
    var key = apiKey()
    if (!key) return Promise.resolve({ ok: false, error: 'Kunci API belum diisi. Buka Tentang → AI Assistant → Isi kunci API.' })
    var text = String(prompt || '').trim()
    if (!text) return Promise.resolve({ ok: false, error: 'Tulis pertanyaan lebih dulu' })
    var body = {
      contents: [{ role: 'user', parts: [{ text: withContext ? sheetContext() + '\n\n' + text : text }] }],
    }
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(modelName()) +
      ':generateContent?key=' + encodeURIComponent(key)
    function send(model) {
      var target = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) +
        ':generateContent?key=' + encodeURIComponent(key)
      return postJson(target, body)
    }

    function parse(response, model) {
      return response.json().then(function (data) {
        if (!response.ok) {
          var detail = data && data.error && data.error.message ? data.error.message : 'HTTP ' + response.status
          return { ok: false, error: 'Permintaan ditolak: ' + detail, status: response.status, detail: detail }
        }
        var candidate = data && data.candidates && data.candidates[0]
        var parts = candidate && candidate.content && candidate.content.parts || []
        var answer = parts.map(function (part) { return part.text || '' }).join('\n').trim()
        return { ok: true, answer: answer || '(jawaban kosong)', model: model }
      })
    }

    return send(modelName()).then(function (response) {
      return parse(response, modelName()).then(function (result) {
        if (result.ok) return result
        // 404 + "no longer available ... use models/xxx" → retry once with the suggested model.
        var hint = /(?:use|to use)\s+models\/([a-zA-Z0-9._-]+)/i.exec(result.detail || '')
        var all = (result.detail || '').match(/models\/([a-zA-Z0-9._-]+)/g) || []
        var suggested = hint ? [null, hint[1]] : (all.length ? [null, all[all.length - 1].split('/')[1]] : null)
        if (result.status === 404 && suggested && suggested[1] !== modelName()) {
          var retryModel = suggested[1]
          return send(retryModel).then(function (next) {
            return parse(next, retryModel).then(function (second) {
              if (second.ok) {
                try { localStorage.setItem(MODEL_STORE, retryModel) } catch (e) {}
                second.note = 'Model dialihkan ke ' + retryModel
              }
              return second
            })
          })
        }
        delete result.status
        delete result.detail
        return result
      })
    }).catch(function (error) {
      var detail = String((error && error.message) || error)
      var hint = ''
      if (error && error.name === 'TypeError') {
        hint = navigator.onLine === false
          ? ' — HP sedang offline. Sambungkan internet lalu coba lagi.'
          : ' — koneksi ke generativelanguage.googleapis.com tidak tercapai. Periksa internet, lalu coba jaringan lain (beberapa jaringan memblokir domain ini).'
      }
      return { ok: false, error: 'Tidak bisa menghubungi server: ' + detail + hint }
    })
  }

  function insertIntoCell(text) {
    var api = window.__univerAPI
    var wb = api && api.getActiveWorkbook && api.getActiveWorkbook()
    if (!wb) return false
    var sheet = wb.getActiveSheet()
    var selection = sheet.getSelection && sheet.getSelection()
    var range = selection && selection.getActiveRange && selection.getActiveRange()
    var target = range || sheet.getRange('A1')
    target.setValue(String(text || ''))
    return true
  }

  var overlay = null

  function closePanel() {
    if (overlay && overlay.parentNode) overlay.parentNode.removeChild(overlay)
    overlay = null
  }

  function showPanel() {
    if (overlay) return
    overlay = document.createElement('div')
    overlay.style.cssText =
      'position:fixed;inset:0;z-index:10001;background:rgba(9,16,38,.46);display:flex;align-items:flex-end;justify-content:center;padding:0'

    var card = document.createElement('div')
    card.style.cssText =
      'background:#fff;color:#111827;border-radius:16px 16px 0 0;width:100%;max-width:560px;padding:16px;' +
      'box-sizing:border-box;font:400 13px/1.5 -apple-system,Roboto,sans-serif;max-height:86vh;display:flex;flex-direction:column'

    var head = document.createElement('div')
    head.style.cssText = 'display:flex;align-items:center;justify-content:space-between;margin-bottom:10px'
    var title = document.createElement('div')
    title.innerHTML = '<b style="font-size:15px">AI Assistant</b><div style="color:#6b7280;font-size:11px" id="tkn-ai-model-label"></div>'
    head.appendChild(title)
    var closeBtn = document.createElement('button')
    closeBtn.type = 'button'
    closeBtn.textContent = '×'
    closeBtn.style.cssText = 'border:0;background:#f1f3f6;color:#111827;width:30px;height:30px;border-radius:15px;font-size:18px;line-height:1'
    closeBtn.addEventListener('click', closePanel)
    head.appendChild(closeBtn)
    card.appendChild(head)

    var modelLabel = title.querySelector('#tkn-ai-model-label')
    modelLabel.textContent = apiKey() ? ('Gemini · ' + modelName() + ' · kunci tersimpan di perangkat') : 'Belum ada kunci API'

    var note = document.createElement('div')
    note.style.cssText = 'background:#f5f7fb;border:1px solid #e3e8f0;border-radius:10px;padding:9px 11px;color:#374151;margin-bottom:10px'
    if (!apiKey()) {
      note.textContent = 'Isi kunci API dulu di Tentang → AI Assistant → “Isi kunci API”. Kunci hanya disimpan di HP ini dan dipakai langsung untuk memanggil Gemini.'
      note.style.borderColor = '#f5c2c0'
      note.style.background = '#fff6f5'
    } else {
      note.textContent = 'Pertanyaan dikirim ke Gemini memakai kunci Anda. Jangan kirim data rahasia.'
    }
    card.appendChild(note)

    var prompt = document.createElement('textarea')
    prompt.placeholder = 'Contoh: ringkas data terpilih, atau buat rumus untuk kolom Total'
    prompt.style.cssText =
      'width:100%;box-sizing:border-box;min-height:76px;border:1px solid #e1e5ec;border-radius:10px;padding:10px;font:inherit;resize:vertical'
    card.appendChild(prompt)

    var row = document.createElement('div')
    row.style.cssText = 'display:flex;gap:8px;align-items:center;margin:10px 0'
    var contextToggle = document.createElement('input')
    contextToggle.type = 'checkbox'
    var contextLabel = document.createElement('label')
    contextLabel.textContent = 'Sertakan data terpilih'
    contextLabel.style.cssText = 'color:#374151;display:flex;gap:6px;align-items:center'
    contextLabel.insertBefore(contextToggle, contextLabel.firstChild)
    row.appendChild(contextLabel)
    card.appendChild(row)

    var actions = document.createElement('div')
    actions.style.cssText = 'display:flex;gap:8px'
    var sendBtn = document.createElement('button')
    sendBtn.type = 'button'
    sendBtn.textContent = 'Tanya'
    sendBtn.style.cssText = 'flex:1;border:0;background:#0a2e96;color:#fff;border-radius:10px;padding:11px;font:600 13px/1 inherit'
    var insertBtn = document.createElement('button')
    insertBtn.type = 'button'
    insertBtn.textContent = 'Sisipkan ke sel'
    insertBtn.disabled = true
    insertBtn.style.cssText = 'flex:1;border:1px solid #d7dbe0;background:#fff;color:#111827;border-radius:10px;padding:11px;font:600 13px/1 inherit'
    actions.appendChild(sendBtn)
    actions.appendChild(insertBtn)
    card.appendChild(actions)

    var output = document.createElement('div')
    output.style.cssText =
      'margin-top:10px;background:#f8fafc;border:1px solid #e5e9f0;border-radius:10px;padding:10px;white-space:pre-wrap;' +
      'max-height:34vh;overflow:auto;display:none'
    card.appendChild(output)

    sendBtn.addEventListener('click', function () {
      sendBtn.disabled = true
      output.style.display = 'block'
      output.textContent = 'Menghubungi Gemini…'
      ask(prompt.value, contextToggle.checked).then(function (result) {
        sendBtn.disabled = false
        if (!result.ok) {
          output.textContent = result.error
          insertBtn.disabled = true
          return
        }
        output.textContent = result.answer + (result.note ? '\n\n(' + result.note + ')' : '')
        insertBtn.disabled = false
        insertBtn.dataset.answer = result.answer
      })
    })

    insertBtn.addEventListener('click', function () {
      if (!insertBtn.dataset.answer) return
      if (insertIntoCell(insertBtn.dataset.answer)) {
        output.textContent += '\n\n— disisipkan ke sel aktif —'
      }
    })

    overlay.addEventListener('click', function (event) {
      if (event.target === overlay) closePanel()
    })
    overlay.appendChild(card)
    document.body.appendChild(overlay)
    prompt.focus()
  }

  function mountUi() {
    if (!/\/sheets\//.test(window.location.pathname)) return
    if (window.__tknAiPanelMounted) return
    window.__tknAiPanelMounted = true
    var button = document.createElement('button')
    button.type = 'button'
    button.textContent = 'AI'
    button.style.cssText =
      'position:fixed;right:16px;bottom:234px;z-index:9000;border:0;border-radius:12px;padding:11px 16px;' +
      'background:#4338ca;color:#fff;font:700 13px/1 -apple-system,Roboto,sans-serif;box-shadow:0 8px 22px rgba(9,16,38,.28)'
    button.addEventListener('click', showPanel)
    document.body.appendChild(button)
  }

  window.__tknAi = {
    show: showPanel,
    ask: ask,
    insert: insertIntoCell,
    hasKey: function () { return !!apiKey() },
    saveKey: saveKey,
    clearKey: clearKey,
    model: modelName,
    mountUi: mountUi,
  }
})()
