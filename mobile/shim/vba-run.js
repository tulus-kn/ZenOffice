/**
 * vba-run.js — a small, self-contained VBA interpreter for Andi Kore Office.
 *
 * It reads the module source out of xl/vbaProject.bin (CFB + MS-OVBA
 * decompression), parses the procedures it understands and executes them
 * against a workbook adapter. Anything it cannot represent is reported as a
 * named error with the source line, never executed approximately.
 *
 * Scope is deliberately honest: this is an interpreter for everyday data
 * macros (cells, loops, conditions, MsgBox, simple functions), NOT a VBA
 * runtime. UserForms, ActiveX, charts, external calls, error handlers and
 * late-bound COM are out of scope and say so.
 */
(function () {
  'use strict'

  // ------------------------------------------------------------ MS-OVBA

  function u16(bytes, i) {
    return bytes[i] | (bytes[i + 1] << 8)
  }

  /** Port of MS-OVBA 2.4.1.3, cross-checked against oletools' decompress. */
  function decompressStream(bytes) {
    if (!bytes || !bytes.length || bytes[0] !== 0x01) {
      throw new Error('not a compressed VBA stream')
    }
    var out = []
    var cursor = 1
    while (cursor < bytes.length) {
      var header = u16(bytes, cursor)
      var chunkSize = (header & 0x0fff) + 3
      var signature = (header >> 12) & 0x07
      var flag = (header >> 15) & 0x01
      if (signature !== 0x03) throw new Error('invalid VBA chunk signature')
      var start = cursor
      cursor += 2
      var end = Math.min(bytes.length, start + chunkSize)
      if (flag === 0) {
        for (var r = 0; r < 4096 && cursor < bytes.length; r++) out.push(bytes[cursor++])
      } else {
        var chunkStart = out.length
        while (cursor < end) {
          var flags = bytes[cursor++]
          for (var bit = 0; bit < 8 && cursor < end; bit++) {
            if (((flags >> bit) & 1) === 0) {
              out.push(bytes[cursor++])
            } else {
              var token = u16(bytes, cursor)
              cursor += 2
              var difference = out.length - chunkStart
              var bitCount = Math.max(Math.ceil(Math.log2(Math.max(difference, 1))), 4)
              var lengthMask = 0xffff >> bitCount
              var length = (token & lengthMask) + 3
              var offset = ((token & (~lengthMask & 0xffff)) >> (16 - bitCount)) + 1
              var source = out.length - offset
              if (source < 0) throw new Error('invalid VBA copy token')
              for (var k = 0; k < length; k++) out.push(out[source + k])
            }
          }
        }
        cursor = end
      }
    }
    return Uint8Array.from(out)
  }

  function looksLikeSource(text) {
    if (!text) return false
    var head = text.slice(0, 400)
    // Real modules start with Attribute lines or a declaration; random bytes that
    // happen to contain "Sub" must not pass, otherwise we lock onto garbage.
    var goodStart = /^\s*Attribute\s+VB_Name|^\s*(Option\s|Private\s|Public\s|Friend\s|Static\s|Sub\s|Function\s|Dim\s|Const\s|Rem\s|')/m.test(head)
    var printable = 0
    var sample = text.slice(0, 4000)
    for (var i = 0; i < sample.length; i++) {
      var c = sample.charCodeAt(i)
      if (c === 9 || c === 10 || c === 13 || (c >= 32 && c < 127) || c >= 160) printable++
    }
    var ratio = sample.length ? printable / sample.length : 0
    if (!goodStart || ratio < 0.9) return false
    return /Attribute\s+VB_Name|\bSub\s+\w+|\bFunction\s+\w+/.test(text)
  }

  /** Module streams carry an optional performance cache before the container. */
  function extractSource(stream) {
    var bytes = stream instanceof Uint8Array ? stream : new Uint8Array(stream)
    for (var offset = 0; offset < Math.min(bytes.length, 4096); offset++) {
      if (bytes[offset] !== 0x01) continue
      var header = u16(bytes, offset + 1)
      if (((header >> 12) & 0x07) !== 0x03) continue
      try {
        var text = new TextDecoder('utf-8').decode(decompressStream(bytes.subarray(offset)))
        if (looksLikeSource(text)) return text.replace(/\u0000+$/g, '')
      } catch (e) {
        /* try the next candidate */
      }
    }
    // Legacy .xls (BIFF) keeps a big performance cache before the code, so the
    // real TextOffset can sit far beyond the first pages. Scan from the end
    // backwards for the first offset that decodes into real source.
    for (var back = bytes.length - 2; back >= 0; back--) {
      if (bytes[back] !== 0x01) continue
      var bh = u16(bytes, back + 1)
      if (((bh >> 12) & 0x07) !== 0x03) continue
      try {
        var btext = new TextDecoder('utf-8').decode(decompressStream(bytes.subarray(back)))
        if (looksLikeSource(btext)) return btext.replace(/\u0000+$/g, '')
      } catch (e) {
        /* keep looking */
      }
    }
    // Some tools store plain text; accept that instead of failing outright.
    try {
      var plain = new TextDecoder('utf-8').decode(bytes).replace(/\u0000+$/g, '')
      if (looksLikeSource(plain)) return plain
    } catch (e) {
      /* ignore */
    }
    return null
  }

  function readModules(bytes) {
    if (!window.XLSX || !window.XLSX.CFB) throw new Error('CFB reader unavailable')
    var cfb = window.XLSX.CFB.read(bytes, { type: 'array' })
    var modules = []
    var entries = cfb.FileIndex || []
    var fullPaths = cfb.FullPaths || []
    for (var i = 0; i < entries.length; i++) {
      var entry = entries[i]
      if (!entry || entry.type !== 2 || !entry.name || !entry.content) continue
      var full = String(fullPaths[i] || entry.name)
      var slash = full.indexOf('/')
      var name = (slash >= 0 ? full.slice(slash + 1) : full).replace(/^\//, '')
      if (!/^VBA\//i.test(name)) continue
      var leaf = name.slice(name.lastIndexOf('/') + 1)
      if (!leaf || /^dir$/i.test(leaf) || /^_VBA_PROJECT$/i.test(leaf) || /^__SRP_/i.test(leaf)) continue
      var content = entry.content
      var source = extractSource(new Uint8Array(content))
      if (source) modules.push({ name: leaf, source: source })
    }
    return modules
  }

  // --------------------------------------------------------- statements

  function stripComment(line) {
    var inString = false
    for (var i = 0; i < line.length; i++) {
      var ch = line[i]
      if (ch === '"') {
        if (inString && line[i + 1] === '"') {
          i++
          continue
        }
        inString = !inString
      } else if (ch === "'" && !inString) {
        return line.slice(0, i)
      }
    }
    return line
  }

  function splitStatements(line) {
    var parts = []
    var current = ''
    var inString = false
    var parenDepth = 0
    for (var i = 0; i < line.length; i++) {
      var ch = line[i]
      if (!inString && ch === '(') parenDepth++
      if (!inString && ch === ')') parenDepth = Math.max(0, parenDepth - 1)
      if (ch === '"') {
        if (inString && line[i + 1] === '"') {
          current += '""'
          i++
          continue
        }
        inString = !inString
        current += ch
        continue
      }
      // ":" separates statements, but "Name:=value" inside a call is a named
      // argument, and a colon inside parentheses is never a separator.
      if (ch === ':' && !inString && line[i + 1] !== '=' && parenDepth === 0) {
        // "Foo:" at the start of a line is a label, not a statement break.
        if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(current.trim())) {
          current += ':'
          continue
        }
        parts.push(current)
        current = ''
        continue
      }
      current += ch
    }
    parts.push(current)
    return parts.map(function (p) { return p.trim() }).filter(Boolean)
  }

  function preprocess(source) {
    var raw = source.split(/\r?\n/)
    var lines = []
    for (var i = 0; i < raw.length; i++) {
      var line = raw[i]
      while (/_$/.test(line.replace(/\s+$/, '')) && i + 1 < raw.length) {
        line = line.replace(/_$/, '') + ' ' + raw[++i]
      }
      line = stripComment(line).trim()
      if (!line) continue
      if (/^(Attribute\s|Option\s|Declare\s)/i.test(line)) continue
      var pieces = splitStatements(line)
      for (var p = 0; p < pieces.length; p++) lines.push({ text: pieces[p], line: i + 1 })
    }
    return lines
  }

  // --------------------------------------------------------- expressions

  function tokenize(text) {
    var tokens = []
    var i = 0
    while (i < text.length) {
      var ch = text[i]
      if (/\s/.test(ch)) {
        i++
        continue
      }
      if (ch === '"') {
        var value = ''
        i++
        while (i < text.length) {
          if (text[i] === '"') {
            if (text[i + 1] === '"') {
              value += '"'
              i += 2
              continue
            }
            i++
            break
          }
          value += text[i++]
        }
        tokens.push({ type: 'str', value: value })
        continue
      }
      if (/[0-9]/.test(ch) || (ch === '.' && /[0-9]/.test(text[i + 1] || ''))) {
        var num = ''
        while (i < text.length && /[0-9.]/.test(text[i])) num += text[i++]
        if (/[eE]/.test(text[i] || '') && /[0-9+-]/.test(text[i + 1] || '')) {
          num += text[i++]
          if (/[+-]/.test(text[i])) num += text[i++]
          while (i < text.length && /[0-9]/.test(text[i])) num += text[i++]
        }
        tokens.push({ type: 'num', value: Number(num) })
        continue
      }
      if (/[A-Za-z_]/.test(ch)) {
        var word = ''
        while (i < text.length && /[A-Za-z0-9_]/.test(text[i])) word += text[i++]
        tokens.push({ type: 'id', value: word })
        continue
      }
      var two = text.slice(i, i + 2)
      if (two === '<=' || two === '>=' || two === '<>') {
        tokens.push({ type: 'op', value: two })
        i += 2
        continue
      }
      if ('+-*/\\^&=<>(),.:'.indexOf(ch) >= 0) {
        tokens.push({ type: 'op', value: ch })
        i++
        continue
      }
      throw new Error('unexpected character "' + ch + '"')
    }
    return tokens
  }

  var PRECEDENCE = {
    Or: 2,
    And: 3,
    '=': 5, '<>': 5, '<': 5, '>': 5, '<=': 5, '>=': 5, Is: 5,
    '&': 6,
    '+': 7, '-': 7,
    Mod: 8,
    '*': 9, '/': 9,
    '^': 11,
  }

  function parseExpression(tokens, pos, minPrec) {
    var token = tokens[pos]
    if (!token) throw new Error('unexpected end of expression')
    var left
    if (token.type === 'op' && (token.value === '-' || token.value === '+')) {
      var operand = parseExpression(tokens, pos + 1, 10)
      left = { type: 'unary', op: token.value, operand: operand.node }
      pos = operand.pos
    } else if (token.type === 'id' && /^Not$/i.test(token.value)) {
      var notOperand = parseExpression(tokens, pos + 1, 4)
      left = { type: 'unary', op: 'Not', operand: notOperand.node }
      pos = notOperand.pos
    } else {
      var primary = parsePrimary(tokens, pos)
      left = primary.node
      pos = primary.pos
    }

    while (pos < tokens.length) {
      var op = tokens[pos]
      var name = op.value
      if (op.type === 'id' && /^(And|Or|Mod|Is)$/i.test(name)) name = name[0].toUpperCase() + name.slice(1).toLowerCase()
      if (op.type !== 'op' && !(op.type === 'id' && /^(And|Or|Mod|Is)$/i.test(op.value))) break
      var prec = PRECEDENCE[name]
      if (!prec || prec < minPrec) break
      var rhs = parseExpression(tokens, pos + 1, name === '^' ? prec : prec + 1)
      left = { type: 'binary', op: name, left: left, right: rhs.node }
      pos = rhs.pos
    }
    return { node: left, pos: pos }
  }

  function parsePrimary(tokens, pos) {
    var token = tokens[pos]
    if (!token) throw new Error('unexpected end of expression')
    if (token.type === 'op' && token.value === '(') {
      var inner = parseExpression(tokens, pos + 1, 0)
      if (!tokens[inner.pos] || tokens[inner.pos].value !== ')') throw new Error('expected ")"')
      return { node: { type: 'group', inner: inner.node }, pos: inner.pos + 1 }
    }
    var node
    if (token.type === 'op' && token.value === '.') {
      var withMember = tokens[pos + 1]
      if (!withMember || withMember.type !== 'id') throw new Error('expected a member name after "."')
      node = { type: 'member', object: { type: 'var', name: '__with' }, name: withMember.value }
      pos += 1
    } else if (token.type === 'num') node = { type: 'num', value: token.value }
    else if (token.type === 'str') node = { type: 'str', value: token.value }
    else if (token.type === 'id') {
      var upper = token.value.toUpperCase()
      if (upper === 'TRUE' || upper === 'FALSE') node = { type: 'bool', value: upper === 'TRUE' }
      else if (upper === 'NOTHING' || upper === 'EMPTY' || upper === 'NULL') node = { type: 'null' }
      else node = { type: 'var', name: token.value }
    } else throw new Error('unexpected token "' + token.value + '"')
    pos++

    while (pos < tokens.length) {
      var next = tokens[pos]
      if (next.type === 'op' && next.value === '.') {
        var member = tokens[pos + 1]
        if (!member || member.type !== 'id') throw new Error('expected a member name after "."')
        node = { type: 'member', object: node, name: member.value }
        pos += 2
        continue
      }
      if (next.type === 'op' && next.value === '(') {
        var args = []
        pos++
        if (tokens[pos] && tokens[pos].value === ')') {
          pos++
        } else {
          while (pos < tokens.length) {
            if (tokens[pos] && tokens[pos].type === 'id' && tokens[pos + 1] && tokens[pos + 1].value === ':' &&
                tokens[pos + 2] && tokens[pos + 2].value === '=') {
              pos += 3 // named argument: Name:=value, the name is ignored
            }
            var arg = parseExpression(tokens, pos, 2)
            args.push(arg.node)
            pos = arg.pos
            if (tokens[pos] && tokens[pos].value === ',') {
              pos++
              continue
            }
            if (tokens[pos] && tokens[pos].value === ')') {
              pos++
              break
            }
            throw new Error('expected "," or ")" in an argument list')
          }
        }
        node = { type: 'call', callee: node, args: args }
        continue
      }
      break
    }
    return { node: node, pos: pos }
  }

  function parseExprText(text) {
    var tokens = tokenize(text)
    var result = parseExpression(tokens, 0, 0)
    if (result.pos !== tokens.length) {
      throw new Error('unexpected trailing text: ' + text)
    }
    return result.node
  }

  // --------------------------------------------------------- block parse

  var UNSUPPORTED_KEYWORDS = [
    [/^Type\b/i, 'Type declaration'],
    [/^Enum\b/i, 'Enum'],
    [/^Declare\b/i, 'DLL Declare'],
    [/^ReDim\b/i, 'ReDim'],
    [/^On\s+\w+\s+GoTo\b/i, 'On <event> GoTo'],
    [/^RaiseEvent\b/i, 'RaiseEvent'],
    [/^Erase\b/i, 'Erase'],
    [/^Open\b\s/i, 'file I/O'],
    [/^Print\s*#/i, 'file I/O'],
  ]

  function parseBlock(lines, index, terminators) {
    var nodes = []
    while (index < lines.length) {
      var entry = lines[index]
      var text = entry.text
      var upper = text.toUpperCase()

      if (terminators && terminators.some(function (t) { return t.test(text) })) {
        return { nodes: nodes, index: index }
      }

      // label:  Name:
      var labelMatch = /^([A-Za-z_][A-Za-z0-9_]{0,39}):\s*(.*)$/.exec(text)
      if (labelMatch && !/^(Case|Else|ElseIf|Default|Next|Loop|End)\b/i.test(labelMatch[1])) {
        nodes.push({ type: 'label', name: labelMatch[1], line: entry.line })
        if (!labelMatch[2].trim()) { index++; continue }
        parseInline(labelMatch[2], entry.line).forEach(function (n) { nodes.push(n) })
        index++
        continue
      }

      var onError = /^On\s+Error\s+(GoTo\s+0|Resume\s+Next|GoTo\s+([A-Za-z_]\w*))/i.exec(text)
      if (onError) {
        var onKind = /GoTo\s+0/i.test(onError[1]) ? 'clear' : /Resume/i.test(onError[1]) ? 'resumeNext' : 'goto'
        nodes.push({ type: 'onerror', kind: onKind, label: onError[2] || null, line: entry.line })
        index++
        continue
      }

      var goToMatch = /^GoTo\s+([A-Za-z_]\w*)/i.exec(text)
      if (goToMatch) {
        nodes.push({ type: 'goto', label: goToMatch[1], line: entry.line })
        index++
        continue
      }

      var forEachMatch = /^For\s+Each\s+([A-Za-z_]\w*)\s+In\s+(.+)$/i.exec(text)
      if (forEachMatch) {
        var feBody = parseBlock(lines, index + 1, [/^Next\b/i])
        nodes.push({ type: 'foreach', name: forEachMatch[1], collection: parseExprText(forEachMatch[2]), body: feBody.nodes, line: entry.line })
        index = feBody.index
        if (index < lines.length && /^Next\b/i.test(lines[index].text)) index++
        continue
      }

      var selectMatch = /^Select\s+Case\s+(.+)$/i.exec(text)
      if (selectMatch) {
        var cases = []
        var elseBody = []
        index++
        while (index < lines.length && !/^End\s+Select\b/i.test(lines[index].text)) {
          var caseMatch = /^Case\s+(.+)$/i.exec(lines[index].text)
          if (!caseMatch) throw new Error('malformed Select Case near line ' + lines[index].line)
          var tests = parseCaseTests(caseMatch[1])
          var caseBody = parseBlock(lines, index + 1, [/^Case\b/i, /^End\s+Select\b/i])
          if (tests.elseCase) elseBody = caseBody.nodes
          else cases.push({ tests: tests.list, body: caseBody.nodes })
          index = caseBody.index
        }
        if (index < lines.length) index++
        nodes.push({ type: 'select', subject: parseExprText(selectMatch[1]), cases: cases, elseBody: elseBody, line: entry.line })
        continue
      }

      // If ... Then (single line or block)
      var ifMatch = /^If\s+(.+?)\s+Then\s*(.*)$/i.exec(text)
      if (ifMatch) {
        var condition = parseExprText(ifMatch[1])
        var inline = ifMatch[2].trim()
        if (inline) {
          var thenNodes = parseInline(inline, entry.line)
          nodes.push({ type: 'if', cond: condition, body: thenNodes, elifs: [], elseBody: [], line: entry.line })
          index++
          continue
        }
        var body = parseBlock(lines, index + 1, [/^ElseIf\b/i, /^Else\b/i, /^End\s+If\b/i])
        var elifs = []
        var elseBody = []
        index = body.index
        while (index < lines.length) {
          var t = lines[index].text
          var elifMatch = /^ElseIf\s+(.+?)\s+Then\s*(.*)$/i.exec(t)
          if (elifMatch) {
            var elifCond = parseExprText(elifMatch[1])
            var elifInline = elifMatch[2].trim()
            if (elifInline) {
              elifs.push({ cond: elifCond, body: parseInline(elifInline, lines[index].line) })
              index++
              continue
            }
            var elifBody = parseBlock(lines, index + 1, [/^ElseIf\b/i, /^Else\b/i, /^End\s+If\b/i])
            elifs.push({ cond: elifCond, body: elifBody.nodes })
            index = elifBody.index
            continue
          }
          if (/^Else\b/i.test(t)) {
            var elseBlock = parseBlock(lines, index + 1, [/^End\s+If\b/i])
            elseBody = elseBlock.nodes
            index = elseBlock.index
            continue
          }
          if (/^End\s+If\b/i.test(t)) {
            index++
            break
          }
          throw new Error('malformed If block near line ' + entry.line)
        }
        nodes.push({ type: 'if', cond: condition, body: body.nodes, elifs: elifs, elseBody: elseBody, line: entry.line })
        continue
      }

      var forMatch = /^For\s+([A-Za-z_]\w*)\s*=\s*(.+?)\s+To\s+(.+?)(?:\s+Step\s+(.+))?$/i.exec(text)
      if (forMatch) {
        var body2 = parseBlock(lines, index + 1, [/^Next\b/i])
        nodes.push({
          type: 'for',
          name: forMatch[1],
          from: parseExprText(forMatch[2]),
          to: parseExprText(forMatch[3]),
          step: forMatch[4] ? parseExprText(forMatch[4]) : null,
          body: body2.nodes,
          line: entry.line,
        })
        index = body2.index
        if (index < lines.length && /^Next\b/i.test(lines[index].text)) index++
        continue
      }

      var doMatch = /^Do\s*(While|Until)?\s*(.*)$/i.exec(text)
      if (doMatch && /^Do\b/i.test(text)) {
        var doCond = doMatch[2] ? parseExprText(doMatch[2]) : null
        var doUntil = /Until/i.test(doMatch[1] || '')
        var doBody = parseBlock(lines, index + 1, [/^Loop\b/i])
        var loopLine = doBody.index < lines.length ? lines[doBody.index].text : ''
        var postMatch = /^Loop\s*(While|Until)\s+(.+)$/i.exec(loopLine)
        nodes.push({
          type: 'do',
          cond: doCond,
          until: doUntil,
          postCond: postMatch ? parseExprText(postMatch[2]) : null,
          postUntil: postMatch ? /Until/i.test(postMatch[1]) : false,
          body: doBody.nodes,
          line: entry.line,
        })
        index = doBody.index + 1
        continue
      }

      var whileMatch = /^While\s+(.+)$/i.exec(text)
      if (whileMatch) {
        var whileBody = parseBlock(lines, index + 1, [/^Wend\b/i])
        nodes.push({ type: 'do', cond: parseExprText(whileMatch[1]), until: false, body: whileBody.nodes, line: entry.line })
        index = whileBody.index
        if (index < lines.length && /^Wend\b/i.test(lines[index].text)) index++
        continue
      }

      if (/^With\s+/i.test(text)) {
        var withBody = parseBlock(lines, index + 1, [/^End\s+With\b/i])
        nodes.push({ type: 'with', expr: parseExprText(text.replace(/^With\s+/i, '')), body: withBody.nodes, line: entry.line })
        index = withBody.index
        if (index < lines.length && /^End\s+With\b/i.test(lines[index].text)) index++
        continue
      }

      var unsupported = UNSUPPORTED_KEYWORDS.find(function (pair) { return pair[0].test(text) })
      if (unsupported) {
        nodes.push({ type: 'unsupported', keyword: unsupported[1], text: text, line: entry.line })
        index++
        continue
      }

      if (/^(End\s+Sub|End\s+Function|End\s+Property)\b/i.test(text)) {
        return { nodes: nodes, index: index }
      }

      if (/^Next\b/i.test(text) || /^Loop\b/i.test(text) || /^End\s+If\b/i.test(text)) {
        return { nodes: nodes, index: index }
      }

      try {
        nodes.push(parseSimple(text, entry.line))
      } catch (error) {
        if (error && !error.lineNumber) {
          error.message = error.message + ' (line ' + entry.line + ': ' + text.slice(0, 60) + ')'
          error.lineNumber = entry.line
        }
        throw error
      }
      index++
    }
    return { nodes: nodes, index: index }
  }

  function parseCaseTests(text) {
    if (/^Else\s*$/i.test(String(text).trim())) return { elseCase: true, list: [] }
    var list = []
    splitArgs(text).forEach(function (raw) {
      var t = raw.trim()
      var isMatch = /^Is\s*(=|<>|<|>|<=|>=)\s*(.+)$/i.exec(t)
      if (isMatch) { list.push({ op: isMatch[1], expr: parseExprText(isMatch[2]) }); return }
      var rangeMatch = /^(.+?)\s+To\s+(.+)$/i.exec(t)
      if (rangeMatch) { list.push({ from: parseExprText(rangeMatch[1]), to: parseExprText(rangeMatch[2]) }); return }
      list.push({ expr: parseExprText(t) })
    })
    return { elseCase: false, list: list }
  }

  function parseSimple(text, line) {
    var exitMatch = /^Exit\s+(Sub|Function|For|Do)$/i.exec(text)
    if (exitMatch) return { type: 'exit', kind: exitMatch[1].toLowerCase(), line: line }

    if (/^Dim\b|^Const\b|^Static\b/i.test(text)) {
      var decl = text.replace(/^(Dim|Const|Static)\s+/i, '')
      var names = []
      var depth = 0
      var chunk = ''
      for (var i = 0; i < decl.length; i++) {
        var ch = decl[i]
        if (ch === '(') depth++
        if (ch === ')') depth--
        if (ch === ',' && depth === 0) {
          names.push(chunk)
          chunk = ''
          continue
        }
        chunk += ch
      }
      names.push(chunk)
      return {
        type: 'dim',
        names: names.map(function (n) { return n.trim().split(/\s+As\s+/i)[0].replace(/[()]/g, '').trim() }).filter(Boolean),
        line: line,
      }
    }

    var msgMatch = /^MsgBox\s+([\s\S]+)$/i.exec(text)
    if (msgMatch) {
      var args = splitArgs(msgMatch[1])
      return { type: 'msgbox', expr: parseExprText(args[0]), line: line }
    }

    var printMatch = /^Debug\.Print\s*(.*)$/i.exec(text)
    if (printMatch) {
      return { type: 'print', expr: printMatch[1] ? parseExprText(printMatch[1]) : null, line: line }
    }

    if (/^Set\s+/i.test(text)) {
      var setText = text.replace(/^Set\s+/i, '')
      var setAssign = splitAssignment(setText)
      if (setAssign) {
        return { type: 'set', target: parseExprText(setAssign[0]), expr: parseExprText(setAssign[1]), line: line }
      }
      return { type: 'unsupported', keyword: 'Set (object assignment)', text: text, line: line }
    }

    // assignment:  lhs = rhs   (top level "=" only)
    var assign = splitAssignment(text)
    if (assign) {
      return { type: 'assign', target: parseExprText(assign[0]), expr: parseExprText(assign[1]), line: line }
    }

    var inlineGoTo = /^GoTo\s+([A-Za-z_]\w*)$/i.exec(text)
    if (inlineGoTo) return { type: 'goto', label: inlineGoTo[1], line: line }

    var resumeMatch = /^Resume\s*(?:(Next)|([A-Za-z_]\w*))?$/i.exec(text)
    if (resumeMatch) {
      return { type: 'resume', next: !!resumeMatch[1], label: resumeMatch[2] || null, line: line }
    }


    // bare call: Name arg, arg
    var callMatch = /^([A-Za-z_]\w*)\s*(.*)$/.exec(text)
    if (callMatch) {
      var rest = callMatch[2].trim()
      if (!rest) return { type: 'callstmt', callee: { type: 'var', name: callMatch[1] }, args: [], line: line }
      if (!rest.startsWith('(')) {
        return {
          type: 'callstmt',
          callee: { type: 'var', name: callMatch[1] },
          args: splitArgs(rest).map(parseExprText),
          line: line,
        }
      }
    }
    return { type: 'callstmt', callee: parseExprText(text), args: [], line: line }
  }

  function parseInline(text, line) {
    var elseIndex = findTopLevelElse(text)
    if (elseIndex >= 0) {
      var thenPart = text.slice(0, elseIndex).trim()
      var elsePart = text.slice(elseIndex + 4).trim()
      return [
        { type: 'if', cond: { type: 'bool', value: true }, body: splitStatements(thenPart).map(function (s) { return parseSimple(s, line) }), elifs: [], elseBody: splitStatements(elsePart).map(function (s) { return parseSimple(s, line) }), line: line },
      ]
    }
    return splitStatements(text).map(function (s) { return parseSimple(s, line) })
  }

  function findTopLevelElse(text) {
    var inString = false
    for (var i = 0; i < text.length - 4; i++) {
      if (text[i] === '"') inString = !inString
      if (inString) continue
      if (/^\sElse\s/i.test(text.slice(i, i + 6)) && /\sElse\s/i.test(text.slice(i - 1, i + 6))) return i
    }
    return -1
  }

  function splitAssignment(text) {
    var inString = false
    var depth = 0
    for (var i = 0; i < text.length; i++) {
      var ch = text[i]
      if (ch === '"') {
        if (inString && text[i + 1] === '"') {
          i++
          continue
        }
        inString = !inString
        continue
      }
      if (inString) continue
      if (ch === '(') depth++
      if (ch === ')') depth--
      if (ch === '=' && depth === 0) {
        var prev = text[i - 1]
        var next = text[i + 1]
        if (prev === '<' || prev === '>' || prev === ':' || next === '=') continue
        var lhs = text.slice(0, i).trim()
        if (!lhs || !/^[A-Za-z_.]/.test(lhs)) continue
        return [lhs, text.slice(i + 1).trim()]
      }
    }
    return null
  }

  function splitArgs(text) {
    var out = []
    var inString = false
    var depth = 0
    var current = ''
    for (var i = 0; i < text.length; i++) {
      var ch = text[i]
      if (ch === '"') {
        if (inString && text[i + 1] === '"') {
          current += '""'
          i++
          continue
        }
        inString = !inString
        current += ch
        continue
      }
      if (!inString) {
        if (ch === '(') depth++
        if (ch === ')') depth--
        if (ch === ',' && depth === 0) {
          out.push(current.trim())
          current = ''
          continue
        }
      }
      current += ch
    }
    if (current.trim()) out.push(current.trim())
    return out
  }

  // ------------------------------------------------------------- modules

  var MODULE_VAR_DEFAULTS = {
    boolean: false, integer: 0, long: 0, single: 0, double: 0, currency: 0, byte: 0,
    string: '', variant: null, range: null, object: null, worksheet: null, workbook: null,
  }

  function moduleLevelVars(lines, stopAt) {
    var vars = {}
    for (var i = 0; i < stopAt; i++) {
      var text = lines[i].text
      var decl = /^(?:Public|Private|Global|Dim|Static)\s+(.+)$/i.exec(text)
      if (!decl) continue
      splitArgs(decl[1]).forEach(function (raw) {
        var name = raw.split(/\s+As\s+/i)[0].replace(/[()]/g, '').trim()
        if (!/^[A-Za-z_]\w*$/.test(name)) return
        var typeMatch = /\s+As\s+([A-Za-z]+)/i.exec(raw)
        var type = typeMatch ? typeMatch[1].toLowerCase() : 'variant'
        vars[name.toLowerCase()] = MODULE_VAR_DEFAULTS[type] !== undefined ? MODULE_VAR_DEFAULTS[type] : null
      })
    }
    return vars
  }

  function parseProcedures(module) {
    var lines = preprocess(module.source)
    var procedures = []
    var firstHeader = -1
    for (var i = 0; i < lines.length; i++) {
      var header = /^(?:Public\s+|Private\s+|Friend\s+|Static\s+)*?(Sub|Function)\s+([A-Za-z_]\w*)\s*(\(([^)]*)\))?/i.exec(lines[i].text)
      if (!header) continue
      if (firstHeader < 0) firstHeader = i
      var kind = header[1].toLowerCase()
      var name = header[2]
      var params = []
      if (header[4]) {
        splitArgs(header[4]).forEach(function (raw) {
          var clean = raw.replace(/^(ByVal|ByRef|Optional|ParamArray)\s+/i, '').trim()
          var pname = clean.split(/\s+As\s+/i)[0].replace(/[()]/g, '').trim()
          if (!pname) return
          var def = null
          var eq = pname.indexOf('=')
          if (eq > 0) {
            def = pname.slice(eq + 1)
            pname = pname.slice(0, eq).trim()
          }
          params.push({ name: pname, default: def })
        })
      }
      var body = parseBlock(lines, i + 1, [/^End\s+(Sub|Function)\b/i])
      procedures.push({
        module: module.name,
        name: name,
        kind: kind,
        params: params,
        body: body.nodes,
        line: lines[i].line,
        endIndex: body.index,
        moduleVars: moduleLevelVars(lines, firstHeader),
      })
      i = body.index
    }
    return procedures
  }

  function listProcedures(bytes) {
    var modules = readModules(bytes)
    var all = []
    var moduleErrors = []
    modules.forEach(function (module) {
      try {
        parseProcedures(module).forEach(function (p) {
          all.push({ name: p.name, kind: p.kind, module: p.module, params: p.params.map(function (x) { return x.name }), line: p.line })
        })
      } catch (error) {
        // A UserForm module or an unusual construct must not hide the rest of
        // the project's macros, so keep going and report the module.
        moduleErrors.push({ module: module.name, error: String((error && error.message) || error) })
      }
    })
    return { modules: modules.map(function (m) { return m.name }), procedures: all, moduleErrors: moduleErrors }
  }

  // --------------------------------------------------------- interpreter

  function toNumber(value) {
    if (typeof value === 'number') return isFinite(value) ? value : 0
    if (typeof value === 'boolean') return value ? -1 : 0
    if (value === null || value === undefined) return 0
    var n = Number(String(value).trim().replace(',', '.'))
    return isFinite(n) ? n : 0
  }

  function toStr(value) {
    if (value === null || value === undefined) return ''
    if (typeof value === 'number') return Number.isInteger(value) ? String(value) : String(value)
    if (typeof value === 'boolean') return value ? 'True' : 'False'
    if (value instanceof Date) return value.toISOString().slice(0, 10)
    return String(value)
  }

  function truthy(value) {
    if (typeof value === 'boolean') return value
    if (typeof value === 'number') return value !== 0
    if (typeof value === 'string') return value !== ''
    return !!value
  }

  function formatNumber(value, pattern) {
    var n = toNumber(value)
    if (pattern === '0.00') return n.toFixed(2)
    if (pattern === '0.0') return n.toFixed(1)
    if (pattern === '0') return String(Math.round(n))
    if (pattern === '#,##0' || pattern === '#,##0.00') {
      var digits = pattern.indexOf('.') >= 0 ? 2 : 0
      var parts = n.toFixed(digits).split('.')
      parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',')
      return parts.join('.')
    }
    return toStr(value)
  }

  function vbaDate(serial) {
    var ms = Date.UTC(1899, 11, 30) + Math.round(serial * 86400000)
    var d = new Date(ms)
    var fake = new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), d.getUTCHours(), d.getUTCMinutes(), d.getUTCSeconds())
    return fake
  }

  function integerWords(n) {
    var ONES = ['', 'satu', 'dua', 'tiga', 'empat', 'lima', 'enam', 'tujuh', 'delapan', 'sembilan', 'sepuluh', 'sebelas']
    function under100(value) {
      if (value < 12) return ONES[value]
      if (value < 20) return ONES[value - 10] + ' belas'
      var tens = Math.floor(value / 10)
      var rest = value % 10
      return ONES[tens] + ' puluh' + (rest ? ' ' + ONES[rest] : '')
    }
    function under1000(value) {
      if (value < 100) return under100(value)
      var hundreds = Math.floor(value / 100)
      var rest = value % 100
      return (hundreds === 1 ? 'seratus' : ONES[hundreds] + ' ratus') + (rest ? ' ' + under100(rest) : '')
    }
    if (n === 0) return 'nol'
    var parts = []
    var rest = n
    var scales = [[1e12, 'triliun'], [1e9, 'miliar'], [1e6, 'juta'], [1e3, 'ribu']]
    for (var i = 0; i < scales.length; i++) {
      var unit = scales[i][0]
      var label = scales[i][1]
      var count = Math.floor(rest / unit)
      if (!count) continue
      rest = rest % unit
      parts.push(count === 1 && label === 'ribu' ? 'seribu' : integerWords(count) + ' ' + label)
    }
    if (rest) parts.push(under1000(rest))
    return parts.join(' ')
  }

  var BUILTINS = {
    terbilang: function (value) {
      var number = toNumber(value)
      if (!isFinite(number)) return ''
      var whole = Math.floor(Math.abs(number))
      var sign = number < 0 ? 'MINUS ' : ''
      var text = integerWords(whole)
      var cents = Math.round((Math.abs(number) - whole) * 100)
      if (cents) text = text + ' KOMA ' + integerWords(cents)
      return (sign + text + ' RUPIAH').toUpperCase()
    },
    now: function () { return vbaDate(Date.now() / 86400000 + 25569) },
    date: function () { var d = vbaDate(Date.now() / 86400000 + 25569); d.setHours(0, 0, 0, 0); return d },
    time: function () { return vbaDate(Date.now() / 86400000 + 25569) },
    len: function (v) { return toStr(v).length },
    ucase: function (v) { return toStr(v).toUpperCase() },
    lcase: function (v) { return toStr(v).toLowerCase() },
    trim: function (v) { return toStr(v).trim() },
    ltrim: function (v) { return toStr(v).replace(/^\s+/, '') },
    rtrim: function (v) { return toStr(v).replace(/\s+$/, '') },
    left: function (v, n) { return toStr(v).slice(0, Math.max(0, toNumber(n))) },
    right: function (v, n) {
      var s = toStr(v)
      var k = Math.max(0, toNumber(n))
      return k === 0 ? '' : s.slice(-k)
    },
    mid: function (v, start, length) {
      var s = toStr(v)
      var from = Math.max(1, toNumber(start)) - 1
      if (length === undefined || length === null) return s.slice(from)
      return s.slice(from, from + Math.max(0, toNumber(length)))
    },
    replace: function (v, find, repl) { return toStr(v).split(toStr(find)).join(toStr(repl === undefined || repl === null ? '' : repl)) },
    instr: function (start, v, find) {
      if (find === undefined) {
        find = v
        v = start
        start = 1
      }
      var idx = toStr(v).indexOf(toStr(find), Math.max(1, toNumber(start)) - 1)
      return idx < 0 ? 0 : idx + 1
    },
    cstr: toStr,
    cint: function (v) { return Math.round(toNumber(v)) },
    clng: function (v) { return Math.round(toNumber(v)) },
    cdbl: toNumber,
    csng: toNumber,
    cbool: truthy,
    abs: function (v) { return Math.abs(toNumber(v)) },
    int: function (v) { return Math.floor(toNumber(v)) },
    fix: function (v) { return Math.trunc(toNumber(v)) },
    round: function (v, d) { var p = Math.pow(10, toNumber(d || 0)); return Math.round(toNumber(v) * p) / p },
    sqr: function (v) { return Math.sqrt(Math.max(0, toNumber(v))) },
    val: function (v) { var m = /^[-+]?[0-9]*\.?[0-9]+/.exec(toStr(v).trim()); return m ? Number(m[0]) : 0 },
    isnumeric: function (v) { return toStr(v).trim() !== '' && isFinite(Number(toStr(v).trim().replace(',', '.'))) },
    isempty: function (v) { return v === null || v === undefined || v === '' },
    chr: function (v) { return String.fromCharCode(toNumber(v)) },
    asc: function (v) { return toStr(v).charCodeAt(0) || 0 },
    space: function (v) { return ' '.repeat(Math.max(0, toNumber(v))) },
    now: function () { return new Date() },
    date: function () { return new Date() },
    time: function () { return new Date() },
    year: function (v) { return (v instanceof Date ? v : new Date()).getFullYear() },
    month: function (v) { return (v instanceof Date ? v : new Date()).getMonth() + 1 },
    day: function (v) { return (v instanceof Date ? v : new Date()).getDate() },
    format: function (v, pattern) { return formatNumber(v, toStr(pattern)) },
    ubound: function (v) { return Array.isArray(v) ? v.length - 1 : -1 },
    lbound: function (v) { return Array.isArray(v) ? 0 : -1 },
    msgbox: function (v) { return v },
  }

  function cellRef(adapter, sheetName, row, col, row2, col2) {
    var endRow = row2 || row
    var endCol = col2 || col
    return {
      __kind: 'range',
      sheet: sheetName,
      row: row,
      col: col,
      row2: endRow,
      col2: endCol,
      area: endRow !== row || endCol !== col,
      address: toAddress(row, col) + (endRow !== row || endCol !== col ? ':' + toAddress(endRow, endCol) : ''),
    }
  }

  function toAddress(row, col) {
    var letters = ''
    var c = col
    while (c > 0) {
      var rem = (c - 1) % 26
      letters = String.fromCharCode(65 + rem) + letters
      c = Math.floor((c - 1) / 26)
    }
    return letters + row
  }

  function singleCell(address) {
    var m = /^([A-Za-z]{1,3})([0-9]{1,7})$/.exec(String(address).replace(/\$/g, '').trim())
    if (!m) throw new Error('not a cell address: ' + address)
    var col = 0
    var letters = m[1].toUpperCase()
    for (var i = 0; i < letters.length; i++) col = col * 26 + (letters.charCodeAt(i) - 64)
    return { row: Number(m[2]), col: col }
  }

  function parseAddress(address) {
    var raw = String(address).replace(/\$/g, '').trim()
    var sheet = null
    var bang = raw.indexOf('!')
    if (bang >= 0) {
      sheet = raw.slice(0, bang).replace(/^'|'$/g, '')
      raw = raw.slice(bang + 1)
    }
    var parts = raw.split(':')
    var a = singleCell(parts[0])
    var b = parts[1] ? singleCell(parts[1]) : a
    return {
      sheet: sheet,
      row: Math.min(a.row, b.row),
      col: Math.min(a.col, b.col),
      row2: Math.max(a.row, b.row),
      col2: Math.max(a.col, b.col),
    }
  }

  function makeObjects(adapter, state) {
    function sheetProxy(name) {
      return { __kind: 'sheet', name: name }
    }
    function resolveSheet(arg) {
      if (arg && arg.__kind === 'sheet') return arg.name
      if (typeof arg === 'number') return adapter.sheetNames[arg - 1]
      var requested = toStr(arg)
      var found = adapter.sheetNames.find(function (n) { return n.toLowerCase() === requested.toLowerCase() })
      if (!found) throw new Error('worksheet not found: ' + requested)
      return found
    }
    return {
      sheetProxy: sheetProxy,
      resolveSheet: resolveSheet,
      rangeFor: function (sheetName, row, col) { return cellRef(adapter, sheetName, row, col) },
      state: state,
    }
  }

  var VBA_CONSTANTS = {
    xlvalues: -4163, xlwhole: 1, xlpart: 2, xlformulas: -4123, xlcomments: -4144,
    xlnothing: 0, xlsimple: 0, xlscreentip: 1,
    xldown: -4121, xlup: -4162, xltoright: -4161, xltoleft: -4159,
    vbokonly: 0, vbokcancel: 1, vbabortyretryignore: 2, vbyesnocancel: 3, vbyesno: 4, vbokcancelretry: 5,
    vbinformation: 64, vbcritical: 16, vbexclamation: 48, vbquestion: 32,
    vbok: 1, vbcancel: 2, vbyes: 6, vbno: 7,
    vbnullstring: '', vbtab: '\t', vbcrlf: '\r\n', vbcr: '\r', vblf: '\n',
    vbboolean: 11, vblong: 3, vbinteger: 2, vbstring: 8, vbdouble: 5, vbvariant: 12,
  }

  function evaluate(node, env) {
    switch (node.type) {
      case 'num':
      case 'str':
      case 'bool':
        return node.value
      case 'null':
        return null
      case 'group':
        return evaluate(node.inner, env)
      case 'var': {
        if (node.name === '__with') {
          if (!env.withTarget) throw new Error('With block is not active here')
          return env.withTarget
        }
        var key = node.name.toLowerCase()
        if (env.vars.has(key)) return env.vars.get(key)
        if (key === 'activesheet') return { __kind: 'sheet', name: env.activeSheet }
        if (key === 'thisworkbook' || key === 'activeworkbook' || key === 'application' || key === 'workbook') {
          return { __kind: 'app' }
        }
        if (BUILTINS[key] && node.args === undefined) return BUILTINS[key]()
        var constant = VBA_CONSTANTS[key]
        if (constant !== undefined) return constant
        if (key === 'err') {
          return {
            __kind: 'err',
            get description() { return env.lastError ? String(env.lastError.message || env.lastError) : '' },
            number: 0,
            clear: function () { env.lastError = null },
          }
        }
        if (/^[A-Za-z]{1,3}[0-9]{1,7}$/.test(node.name)) {
          return cellValueAt(env, env.activeSheet, parseAddress(node.name))
        }
        throw new Error('undefined variable: ' + node.name)
      }
      case 'unary': {
        var value = evaluate(node.operand, env)
        if (node.op === '-') return -toNumber(value)
        if (node.op === 'Not') return !truthy(value)
        return value
      }
      case 'binary':
        return binaryOp(node, env)
      case 'member':
        return memberOf(node, env)
      case 'call':
        return callNode(node, env)
      default:
        throw new Error('unsupported expression node ' + node.type)
    }
  }

  function cellValueAt(env, sheet, address) {
    return env.adapter.getCell(sheet, address.row, address.col)
  }

  function memberOf(node, env) {
    var target = evaluate(node.object, env)
    var name = node.name
    var lower = name.toLowerCase()
    if (target && target.__kind === 'range') {
      if (lower === 'value' || lower === 'value2' || lower === 'formula' || lower === 'formular1c1') {
        return env.adapter.getCell(target.sheet, target.row, target.col)
      }
      if (lower === 'text') return toStr(env.adapter.getCell(target.sheet, target.row, target.col))
      if (lower === 'address') return target.address.replace(/([A-Z]+)(\d+)/, '$1$$$2')
      if (lower === 'row') return target.row
      if (lower === 'column') return target.col
      if (lower === 'rows' || lower === 'columns') {
        return {
          __kind: 'range',
          sheet: target.sheet, row: target.row, col: target.col, row2: target.row2, col2: target.col2,
          area: target.area, address: target.address,
          count: lower === 'rows' ? (target.row2 - target.row + 1) : (target.col2 - target.col + 1),
        }
      }
      if (lower === 'count' || lower === 'cells') {
        return target.area ? (target.row2 - target.row + 1) * (target.col2 - target.col + 1) : 1
      }
      if (lower === 'offset') return function (dr, dc) {
        var r = target.row + toNumber(dr === undefined ? 0 : dr)
        var c = target.col + toNumber(dc === undefined ? 0 : dc)
        return cellRef(env.adapter, target.sheet, r, c)
      }
      if (lower === 'item') return function (index) {
        var i = toNumber(index)
        var width = target.col2 - target.col + 1
        var r = target.row + Math.floor((i - 1) / width)
        var c = target.col + ((i - 1) % width)
        return cellRef(env.adapter, target.sheet, r, c)
      }
      if (lower === 'activate' || lower === 'select') return function () {
        env.activeSheet = target.sheet
        return null
      }
      if (lower === 'end') return function (direction) {
        var dir = toNumber(direction)
        var r = target.row
        var c = target.col
        var last = { row: r, col: c }
        var stepRow = dir === -4121 ? 1 : dir === -4162 ? -1 : 0
        var stepCol = dir === -4161 ? 1 : dir === -4159 ? -1 : 0
        if (!stepRow && !stepCol) return target
        while (r > 0 && c > 0) {
          var nextValue = env.adapter.getCell(target.sheet, r + stepRow, c + stepCol)
          if (nextValue === null || nextValue === '' || nextValue === undefined) break
          r += stepRow
          c += stepCol
          last = { row: r, col: c }
        }
        return cellRef(env.adapter, target.sheet, last.row, last.col)
      }
      if (lower === 'find') return function (needle) {
        for (var r = target.row; r <= target.row2; r++) {
          for (var c = target.col; c <= target.col2; c++) {
            if (looseEquals(env.adapter.getCell(target.sheet, r, c), needle)) {
              return cellRef(env.adapter, target.sheet, r, c)
            }
          }
        }
        return null
      }
      throw new Error('range property not supported: ' + name)
    }
    if (target && target.__kind === 'sheet') {
      if (lower === 'name') return target.name
      if (lower === 'range') return function (address) {
        var parsed = parseAddress(address)
        return cellRef(env.adapter, target.name, parsed.row, parsed.col)
      }
      if (lower === 'cells') return function (row, col) {
        return cellRef(env.adapter, target.name, toNumber(row), toNumber(col))
      }
      if (lower === 'activate' || lower === 'select') return function () {
        env.activeSheet = target.name
        return null
      }
      throw new Error('worksheet member not supported: ' + name)
    }
    if (target && target.__kind === 'app') {
      if (lower === 'sheets' || lower === 'worksheets') return function (arg) {
        return { __kind: 'sheet', name: env.resolveSheet(arg) }
      }
      if (lower === 'activesheet') return { __kind: 'sheet', name: env.activeSheet }
      if (lower === 'range') return function (address) {
        return resolveRangeAddress(env, address)
      }
      if (lower === 'screenupdating' || lower === 'enableevents' || lower === 'displayalerts' ||
          lower === 'calculation' || lower === 'visible' || lower === 'asktoupdate') {
        return true
      }
      throw new Error('workbook member not supported: ' + name)
    }
    if (target && target.__kind === 'err') {
      if (lower === 'description') return target.description
      if (lower === 'number') return target.number
      if (lower === 'clear') return function () { target.clear() }
      throw new Error('Err member not supported: ' + name)
    }
    throw new Error('cannot read "' + name + '" from this value')
  }

  function binaryOp(node, env) {
    var op = node.op
    if (op === 'And') return truthy(evaluate(node.left, env)) && truthy(evaluate(node.right, env))
    if (op === 'Or') return truthy(evaluate(node.left, env)) || truthy(evaluate(node.right, env))
    var a = evaluate(node.left, env)
    var b = evaluate(node.right, env)
    if (op === 'Is') {
      if (b === null || b === undefined) return a === null || a === undefined
      if (a === null || a === undefined) return false
      return a === b
    }
    switch (op) {
      case '+':
        if (typeof a === 'string' || typeof b === 'string') {
          // VBA prefers numeric addition when both sides look numeric
          if (isFinite(Number(a)) && isFinite(Number(b)) && toStr(a).trim() !== '' && toStr(b).trim() !== '') {
            return toNumber(a) + toNumber(b)
          }
          return toStr(a) + toStr(b)
        }
        return toNumber(a) + toNumber(b)
      case '-': return toNumber(a) - toNumber(b)
      case '*': return toNumber(a) * toNumber(b)
      case '/': {
        var divisor = toNumber(b)
        if (divisor === 0) throw new Error('division by zero')
        return toNumber(a) / divisor
      }
      case '\\': {
        var d = toNumber(b)
        if (d === 0) throw new Error('division by zero')
        return Math.trunc(toNumber(a) / d)
      }
      case 'Mod': {
        var m = toNumber(b)
        if (m === 0) throw new Error('division by zero')
        return toNumber(a) % m
      }
      case '^': return Math.pow(toNumber(a), toNumber(b))
      case '&': return toStr(a) + toStr(b)
      case '=': return looseEquals(a, b)
      case '<>': return !looseEquals(a, b)
      case '<': return compare(a, b) < 0
      case '>': return compare(a, b) > 0
      case '<=': return compare(a, b) <= 0
      case '>=': return compare(a, b) >= 0
      default: throw new Error('unsupported operator ' + op)
    }
  }

  function looseEquals(a, b) {
    if (typeof a === 'number' || typeof b === 'number') return toNumber(a) === toNumber(b)
    return toStr(a) === toStr(b)
  }

  function compare(a, b) {
    if (typeof a === 'number' || typeof b === 'number') return toNumber(a) - toNumber(b)
    var x = toStr(a)
    var y = toStr(b)
    return x < y ? -1 : x > y ? 1 : 0
  }

  function resolveRangeAddress(env, address) {
    var raw = toStr(address)
    if (/^[A-Za-z]{1,3}\$?[0-9]{1,7}(:\$?[A-Za-z]{1,3}\$?[0-9]{1,7})?$/.test(raw.replace(/\$/g, '')) === false) {
      if (env.adapter.namedRange) {
        var named = env.adapter.namedRange(raw)
        if (named) return cellRef(env.adapter, named.sheet, named.row, named.col, named.row2, named.col2)
      }
      throw new Error('named range not found: ' + raw)
    }
    var parsed = parseAddress(raw)
    var sheet = parsed.sheet || env.activeSheet
    return cellRef(env.adapter, sheet, parsed.row, parsed.col, parsed.row2, parsed.col2)
  }

  function callNode(node, env) {
    var callee = node.callee
    var args = node.args || []
    if (callee.type === 'var') {
      var name = callee.name
      var lower = name.toLowerCase()
      if (lower === 'range') {
        var first = args.length ? evaluate(args[0], env) : ''
        if (first && first.__kind === 'range' && args.length > 1) {
          var second = evaluate(args[1], env)
          if (second && second.__kind === 'range') {
            return cellRef(env.adapter, first.sheet, Math.min(first.row, second.row), Math.min(first.col, second.col),
              Math.max(first.row, second.row), Math.max(first.col, second.col))
          }
          return first
        }
        var address = toStr(first)
        if (args.length > 1) address = address + ':' + toStr(evaluate(args[1], env))
        return resolveRangeAddress(env, address)
      }
      if (lower === 'cells') {
        var sheetName = env.activeSheet
        var rowArg = args[0]
        var colArg = args[1]
        // Cells(r, c) with a sheet receiver is handled in memberOf
        return cellRef(env.adapter, sheetName, toNumber(evaluate(rowArg, env)), toNumber(evaluate(colArg, env)))
      }
      if (lower === 'worksheets' || lower === 'sheets') {
        return { __kind: 'sheet', name: env.resolveSheet(evaluate(args[0], env)) }
      }
      var builtin = BUILTINS[lower]
      if (builtin) {
        var values = args.map(function (arg) { return evaluate(arg, env) })
        return builtin.apply(null, values)
      }
      var procedure = env.procedures[name.toLowerCase()]
      if (procedure) {
        var callArgs = args.map(function (arg) { return evaluate(arg, env) })
        return invokeProcedure(procedure, callArgs, env)
      }
      throw new Error('unknown function or procedure: ' + name)
    }
    if (callee.type === 'member') {
      var receiver = evaluate(callee.object, env)
      var member = memberOf(callee, env)
      if (typeof member === 'function') {
        var memberArgs = args.map(function (arg) { return evaluate(arg, env) })
        return member.apply(null, memberArgs)
      }
      return member
    }
    var value = evaluate(callee, env)
    if (typeof value === 'function') {
      return value.apply(null, args.map(function (arg) { return evaluate(arg, env) }))
    }
    throw new Error('cannot call this value')
  }

  function lvalue(node, env) {
    if (node.type === 'var') {
      var key = node.name.toLowerCase()
      if (/^[A-Za-z]{1,3}[0-9]{1,7}$/.test(node.name) && !env.vars.has(key)) {
        var address = parseAddress(node.name)
        return {
          get: function () { return cellValueAt(env, env.activeSheet, address) },
          set: function (v) { env.adapter.setCell(env.activeSheet, address.row, address.col, v) },
        }
      }
      return {
        get: function () { return env.vars.get(key) },
        set: function (v) {
          // a bare name that matches the enclosing Function is the return value
          env.vars.set(key, v)
        },
      }
    }
    if (node.type === 'member') {
      var target = evaluate(node.object, env)
      var lower = node.name.toLowerCase()
      if (target && target.__kind === 'app') {
        // Application.ScreenUpdating = False and friends: accepted, nothing to apply
        return { get: function () { return true }, set: function () {} }
      }
      if (target && target.__kind === 'range' && ['value', 'value2', 'formula', 'formular1c1', 'text'].indexOf(lower) >= 0) {
        if (lower === 'formula' || lower === 'formular1c1') {
          return {
            get: function () {
              var ws = env.adapter.getSheet ? env.adapter.getSheet(target.sheet) : null
              var cell = ws && ws[target.address]
              return cell && cell.f ? '=' + cell.f : env.adapter.getCell(target.sheet, target.row, target.col)
            },
            set: function (v) {
              if (env.adapter.setFormula) env.adapter.setFormula(target.sheet, target.row, target.col, v)
              else env.adapter.setCell(target.sheet, target.row, target.col, v)
            },
          }
        }
        return {
          get: function () { return env.adapter.getCell(target.sheet, target.row, target.col) },
          set: function (v) { env.adapter.setCell(target.sheet, target.row, target.col, v) },
        }
      }
      throw new Error('cannot assign to "' + node.name + '"')
    }
    throw new Error('unsupported assignment target')
  }

  function labelIndex(nodes, label) {
    var want = String(label).toLowerCase()
    for (var i = 0; i < nodes.length; i++) {
      if (nodes[i].type === 'label' && String(nodes[i].name).toLowerCase() === want) return i
    }
    return -1
  }

  function collectionItems(env, value) {
    if (value && value.__kind === 'range' && value.row2 && (value.row !== value.row2 || value.col !== value.col2)) {
      var items = []
      for (var r = value.row; r <= value.row2; r++) {
        for (var c = value.col; c <= value.col2; c++) {
          items.push(cellRef(env.adapter, value.sheet, r, c))
        }
      }
      return items
    }
    if (Array.isArray(value)) return value
    if (value && value.__kind === 'range') return [value]
    throw new Error('For Each needs a range or an array')
  }

  function executeNodes(nodes, env) {
    var i = 0
    while (i < nodes.length) {
      try {
        var flow = executeNode(nodes[i], env)
        if (flow) {
          if (typeof flow === 'object' && flow.goto) {
            var target = labelIndex(nodes, flow.goto)
            if (target >= 0) {
              env.inHandler = false
              i = target + 1
              continue
            }
            return flow
          }
          return flow
        }
        i++
      } catch (error) {
        if (error && !error.sourceLine && nodes[i] && nodes[i].line) {
          error.sourceLine = nodes[i].line
          error.message = error.message + ' (line ' + nodes[i].line + ')'
        }
        // On Error Resume Next: swallow and move to the following statement
        if (env.errorLabel === '__resume-next') {
          env.lastError = error
          i++
          continue
        }
        if (env.errorLabel) {
          var handler = labelIndex(nodes, env.errorLabel)
          env.lastError = error
          env.inHandler = true
          if (handler >= 0) {
            i = handler + 1
            continue
          }
          return { goto: env.errorLabel }
        }
        throw error
      }
    }
    return null
  }

  function executeNode(node, env) {
    if (++env.steps > env.limits.steps) throw new Error('macro stopped: step limit exceeded')
    switch (node.type) {
      case 'dim':
        node.names.forEach(function (n) {
          var key = n.toLowerCase()
          if (!env.vars.has(key)) env.vars.set(key, null)
        })
        return null
      case 'assign': {
        var ref = lvalue(node.target, env)
        var value = evaluate(node.expr, env)
        ref.set(value)
        return null
      }
      case 'msgbox':
        env.output.push(toStr(evaluate(node.expr, env)))
        return null
      case 'print':
        env.output.push(node.expr ? toStr(evaluate(node.expr, env)) : '')
        return null
      case 'callstmt': {
        var callee = node.callee
        if (callee.type === 'var') {
          var name = callee.name
          var args = node.args.map(function (arg) { return evaluate(arg, env) })
          var builtin = BUILTINS[name.toLowerCase()]
          if (builtin) {
            builtin.apply(null, args)
            return null
          }
          var procedure = env.procedures[name.toLowerCase()]
          if (procedure) {
            invokeProcedure(procedure, args, env)
            return null
          }
          throw new Error('unknown procedure: ' + name)
        }
        callNode({ callee: callee, args: node.args }, env)
        return null
      }
      case 'if': {
        if (truthy(evaluate(node.cond, env))) return executeNodes(node.body, env)
        for (var e = 0; e < node.elifs.length; e++) {
          if (truthy(evaluate(node.elifs[e].cond, env))) return executeNodes(node.elifs[e].body, env)
        }
        return executeNodes(node.elseBody, env)
      }
      case 'for': {
        var start = toNumber(evaluate(node.from, env))
        var end = toNumber(evaluate(node.to, env))
        var step = node.step ? toNumber(evaluate(node.step, env)) : 1
        if (step === 0) throw new Error('For loop with Step 0')
        env.vars.set(node.name.toLowerCase(), start)
        for (var v = start; step > 0 ? v <= end : v >= end; v += step) {
          env.vars.set(node.name.toLowerCase(), v)
          var flow = executeNodes(node.body, env)
          if (flow === 'exitfor') return null
          if (flow === 'exitsub' || flow === 'exitfunction') return flow
          if (++env.steps > env.limits.steps) throw new Error('macro stopped: step limit exceeded')
        }
        return null
      }
      case 'do': {
        var guard = 0
        while (true) {
          if (node.cond) {
            var cond = truthy(evaluate(node.cond, env))
            if (node.until ? cond : !cond) break
          }
          var bodyFlow = executeNodes(node.body, env)
          if (bodyFlow === 'exitdo') break
          if (bodyFlow === 'exitsub' || bodyFlow === 'exitfunction') return bodyFlow
          if (node.postCond) {
            var post = truthy(evaluate(node.postCond, env))
            if (node.postUntil ? post : !post) break
          }
          if (++env.steps > env.limits.steps) throw new Error('macro stopped: step limit exceeded')
          if (++guard > env.limits.steps) throw new Error('macro stopped: loop limit exceeded')
        }
        return null
      }
      case 'with': {
        var previous = env.withTarget
        env.withTarget = evaluate(node.expr, env)
        var withFlow = executeNodes(node.body, env)
        env.withTarget = previous
        return withFlow
      }
      case 'exit':
        if (node.kind === 'sub') return 'exitsub'
        if (node.kind === 'function') return 'exitfunction'
        if (node.kind === 'for') return 'exitfor'
        if (node.kind === 'do') return 'exitdo'
        return null
      case 'label':
        return null
      case 'goto':
        return { goto: node.label }
      case 'onerror':
        if (node.kind === 'goto') env.errorLabel = node.label
        else if (node.kind === 'clear') env.errorLabel = null
        else env.errorLabel = '__resume-next'
        return null
      case 'resume':
        if (node.label) return { goto: node.label }
        if (node.next) return { resume: 'next' }
        return { resume: 'plain' }
      case 'set': {
        var setTarget = lvalue(node.target, env)
        setTarget.set(evaluate(node.expr, env))
        return null
      }
      case 'select': {
        var subject = evaluate(node.subject, env)
        for (var si = 0; si < node.cases.length; si++) {
          var entry = node.cases[si]
          for (var ti = 0; ti < entry.tests.length; ti++) {
            if (caseMatches(entry.tests[ti], subject, env)) return executeNodes(entry.body, env)
          }
        }
        return executeNodes(node.elseBody, env)
      }
      case 'foreach': {
        var collection = collectionItems(env, evaluate(node.collection, env))
        for (var fi = 0; fi < collection.length; fi++) {
          var item = collection[fi]
          env.vars.set(node.name.toLowerCase(), item && item.__kind === 'range' ? item : item)
          var feFlow = executeNodes(node.body, env)
          if (feFlow === 'exitfor') return null
          if (feFlow === 'exitsub' || feFlow === 'exitfunction') return feFlow
          if (++env.steps > env.limits.steps) throw new Error('macro stopped: step limit exceeded')
        }
        return null
      }
      case 'unsupported':
        throw new MacroUnsupported(node)
      default:
        throw new Error('unsupported statement at line ' + (node.line || '?'))
    }
  }

  function MacroUnsupported(node) {
    var error = new Error('unsupported VBA construct: ' + node.keyword + ' (line ' + node.line + ')')
    error.code = 'unsupported'
    error.keyword = node.keyword
    error.line = node.line
    return error
  }

  function caseMatches(test, subject, env) {
    if (test.expr) return looseEquals(evaluate(test.expr, env), subject)
    if (test.op) {
      var rhs = evaluate(test.expr, env)
      switch (test.op) {
        case '=': return looseEquals(subject, rhs)
        case '<>': return !looseEquals(subject, rhs)
        case '<': return compare(subject, rhs) < 0
        case '>': return compare(subject, rhs) > 0
        case '<=': return compare(subject, rhs) <= 0
        case '>=': return compare(subject, rhs) >= 0
        default: return false
      }
    }
    var from = evaluate(test.from, env)
    var to = evaluate(test.to, env)
    return compare(subject, from) >= 0 && compare(subject, to) <= 0
  }

  function invokeProcedure(procedure, args, parentEnv) {
    var env = {
      adapter: parentEnv.adapter,
      procedures: parentEnv.procedures,
      sheetNames: parentEnv.sheetNames,
      activeSheet: parentEnv.activeSheet,
      resolveSheet: parentEnv.resolveSheet,
      output: parentEnv.output,
      steps: parentEnv.steps,
      limits: parentEnv.limits,
      vars: new Map(parentEnv.vars),
      moduleVars: parentEnv.moduleVars,
      isProcedureScope: true,
      withTarget: null,
      procedure: procedure,
      activeSheet: parentEnv.activeSheet,
      getCell: parentEnv.getCell,
      setCell: parentEnv.setCell,
      errorLabel: null,
      inHandler: false,
      lastError: parentEnv.lastError || null,
    }
    procedure.params.forEach(function (param, index) {
      env.vars.set(param.name.toLowerCase(), index < args.length ? args[index] : null)
    })
    var flow = executeNodes(procedure.body, env)
    parentEnv.steps = env.steps
    parentEnv.lastError = env.lastError
    // Locals stay local (ByVal must not leak), module variables are shared
    if (parentEnv.moduleVars) {
      Object.keys(parentEnv.moduleVars).forEach(function (key) {
        if (env.vars.has(key)) parentEnv.vars.set(key, env.vars.get(key))
      })
    }
    if (flow && typeof flow === 'object' && flow.goto) throw new Error('label not found: ' + flow.goto)
    if (procedure.kind === 'function') {
      var nameKey = procedure.name.toLowerCase()
      return env.vars.has(nameKey) ? env.vars.get(nameKey) : null
    }
    if (flow === 'exitsub' || flow === 'exitfunction') return null
    return null
  }

  function runMacro(name, adapter, options) {
    var opts = options || {}
    var modules = readModules(opts.bytes)
    var procedures = []
    modules.forEach(function (module) {
      procedures = procedures.concat(parseProcedures(module))
    })
    var index = {}
    procedures.forEach(function (p) { index[p.name.toLowerCase()] = p })
    var procedure = index[String(name).toLowerCase()]
    if (!procedure) throw new Error('macro not found: ' + name)

    var moduleVars = {}
    procedures.forEach(function (proc) {
      Object.keys(proc.moduleVars || {}).forEach(function (key) {
        if (!(key in moduleVars)) moduleVars[key] = proc.moduleVars[key]
      })
    })
    var env = {
      adapter: adapter,
      procedures: index,
      moduleVars: moduleVars,
      sheetNames: adapter.sheetNames.slice(),
      activeSheet: adapter.activeSheet || adapter.sheetNames[0],
      resolveSheet: function (arg) {
        if (arg && arg.__kind === 'sheet') return arg.name
        if (typeof arg === 'number') return adapter.sheetNames[arg - 1]
                var requested = toStr(arg)
        var found = adapter.sheetNames.find(function (n) { return n.toLowerCase() === requested.toLowerCase() })
        if (!found) throw new Error('worksheet not found: ' + requested)
        return found
      },
      output: [],
      steps: 0,
      limits: { steps: Number(opts.maxSteps) || 200000 },
      vars: new Map(),
      withTarget: null,
      procedure: procedure,
      getCell: adapter.getCell,
      setCell: adapter.setCell,
      errorLabel: null,
      inHandler: false,
      lastError: null,
    }
    var shared = opts.state && typeof opts.state === 'object' ? opts.state : null
    Object.keys(moduleVars).forEach(function (key) { env.vars.set(key, moduleVars[key]) })
    if (shared) {
      Object.keys(shared).forEach(function (key) { env.vars.set(key, shared[key]) })
    }
    if (opts.args && opts.args.length) {
      procedure.params.forEach(function (param, index) {
        env.vars.set(param.name.toLowerCase(), index < opts.args.length ? opts.args[index] : null)
      })
    }
    var flow = executeNodes(procedure.body, env)
    if (flow && typeof flow === 'object' && flow.goto) throw new Error('label not found: ' + flow.goto)
    if (shared) {
      Object.keys(moduleVars).forEach(function (key) {
        if (env.vars.has(key)) shared[key] = env.vars.get(key)
      })
    }
    return {
      ok: true,
      name: procedure.name,
      value: procedure.kind === 'function' && env.vars.has(procedure.name.toLowerCase()) ? env.vars.get(procedure.name.toLowerCase()) : null,
      output: env.output,
      steps: env.steps,
      changes: adapter.changes || [],
      flow: flow || 'normal',
    }
  }

  // --------------------------------------------------- SheetJS adapter

  function contextFromSheetJS(wb, activeSheetName) {
    var names = (wb.SheetNames || []).slice()
    var changes = []
    var cellKey = function (sheet, row, col) { return sheet + '!' + row + ':' + col }
    var seen = Object.create(null)
    var named = Object.create(null)
    try {
      ((wb.Workbook && wb.Workbook.Names) || []).forEach(function (item) {
        if (!item || !item.Name || !item.Ref) return
        var ref = String(item.Ref)
        if (ref.indexOf('#') >= 0 || ref.indexOf('!') < 0) return
        try {
          var parsed = parseAddress(ref)
          named[String(item.Name).toLowerCase()] = {
            sheet: parsed.sheet || names[0],
            row: parsed.row,
            col: parsed.col,
            row2: parsed.row2,
            col2: parsed.col2,
          }
        } catch (e) { /* skip unparsable names */ }
      })
    } catch (e) { /* older workbooks may have no Names table */ }

    var adapter = {
      sheetNames: names,
      activeSheet: activeSheetName || names[0],
      changes: changes,
      namedRange: function (name) { return named[String(name).toLowerCase()] || null },
      getSheet: function (sheet) { return wb.Sheets[sheet] || null },
      setFormula: function (sheet, row, col, formulaText) {
        var ws = wb.Sheets[sheet]
        if (!ws) return
        var addr = toAddress(row, col)
        var text = String(formulaText).replace(/^=/, '')
        ws[addr] = { t: 'n', f: text }
        var key = cellKey(sheet, row, col)
        if (!seen[key]) {
          seen[key] = { sheet: sheet, row: row, col: col, address: addr }
          changes.push(seen[key])
        }
        seen[key].value = '=' + text
      },
      getCell: function (sheet, row, col) {
        var ws = wb.Sheets[sheet]
        if (!ws) return null
        var addr = toAddress(row, col)
        var cell = ws[addr]
        if (!cell) return null
        return cell.v === undefined ? null : cell.v
      },
      setCell: function (sheet, row, col, value) {
        var ws = wb.Sheets[sheet]
        if (!ws) return
        var addr = toAddress(row, col)
        if (value === null || value === undefined || value === '') {
          delete ws[addr]
        } else if (typeof value === 'number') {
          ws[addr] = { t: 'n', v: value }
        } else if (typeof value === 'boolean') {
          ws[addr] = { t: 'b', v: value }
        } else {
          ws[addr] = { t: 's', v: toStr(value) }
        }
        var key = cellKey(sheet, row, col)
        if (!seen[key]) {
          seen[key] = { sheet: sheet, row: row, col: col, address: addr }
          changes.push(seen[key])
        }
        seen[key].value = value
      },
    }
    return adapter
  }

  window.__zenVbaMacros = {
    listProcedures: listProcedures,
    runMacro: runMacro,
    contextFromSheetJS: contextFromSheetJS,
    decompressStream: decompressStream,
    readModules: readModules,
    parseProcedures: parseProcedures,
    tokenize: tokenize,
    parseExprText: parseExprText,
  }
})()
