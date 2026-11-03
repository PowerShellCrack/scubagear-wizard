/* Small helpers: quoting, validation, safe inline rendering. */
(function (global) {
  'use strict';

  const SMART_QUOTES = /[\u2018\u2019\u201A\u201B]/g;
  const CONTROL_CHARS = /[\u0000-\u001F\u007F]/g;

  function clean(value) {
    return String(value == null ? '' : value).replace(CONTROL_CHARS, ' ').trim();
  }

  // PowerShell single-quoted literal; smart single quotes also terminate strings, so double them too.
  function psq(value) {
    return "'" + clean(value).replace(/'/g, "''").replace(SMART_QUOTES, (m) => m + m) + "'";
  }

  // YAML single-quoted scalar (no backslash escaping needed for Windows paths).
  function yq(value) {
    return "'" + clean(value).replace(/'/g, "''") + "'";
  }

  const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  const THUMB = /^[0-9a-f]{40}$/i;
  const DOMAIN = /^[a-z0-9]([a-z0-9-]*[a-z0-9])?(\.[a-z0-9]([a-z0-9-]*[a-z0-9])?)+$/i;
  const LABEL = /^[a-z0-9]([a-z0-9-]*[a-z0-9])?$/i;
  const BAD_PATH = /[<>"|?*\u0000-\u001F]/;

  const VALIDATORS = {
    any: (v) => v.length > 0 && v.length <= 200,
    guid: (v) => GUID.test(v),
    thumb: (v) => THUMB.test(v),
    domain: (v) => DOMAIN.test(v),
    label: (v) => LABEL.test(v),
    path: (v) => v.length > 0 && v.length <= 240 && !BAD_PATH.test(v),
  };

  const MESSAGES = {
    guid: 'Expected a GUID like 12345678-1234-1234-1234-123456789012.',
    guidlist: 'Each entry must be a GUID. Separate several with commas or new lines.',
    thumb: 'Expected a 40-character hexadecimal thumbprint.',
    domain: 'Expected a domain like contoso.onmicrosoft.com.',
    label: 'Expected just the tenant name, like contoso.',
    path: 'That does not look like a valid Windows path.',
    any: 'Please shorten this value.',
  };

  // Returns { ok, value, message }. Empty input is "ok" with an empty value (fields are optional).
  function validate(kind, raw) {
    const text = clean(raw);
    if (!text) return { ok: true, value: kind === 'guidlist' ? [] : '', empty: true };
    if (kind === 'guidlist') {
      const items = text.split(/[\s,;]+/).filter(Boolean);
      const ok = items.every((i) => GUID.test(i));
      return { ok, value: ok ? items : [], message: ok ? '' : MESSAGES.guidlist };
    }
    const fn = VALIDATORS[kind] || VALIDATORS.any;
    const ok = fn(text);
    return { ok, value: ok ? text : '', message: ok ? '' : MESSAGES[kind] || MESSAGES.any };
  }

  // Build a multi-line PowerShell command using backtick continuations.
  function cmd(name, parts) {
    const p = parts.filter(Boolean);
    return p.length ? name + ' `\n  ' + p.join(' `\n  ') : name;
  }

  function el(tag, attrs, children) {
    const node = document.createElement(tag);
    if (attrs) {
      Object.keys(attrs).forEach((k) => {
        const v = attrs[k];
        if (v == null || v === false) return;
        if (k === 'class') node.className = v;
        else if (k === 'text') node.textContent = v;
        else if (k.slice(0, 2) === 'on' && typeof v === 'function') node.addEventListener(k.slice(2), v);
        else node.setAttribute(k, v === true ? '' : v);
      });
    }
    (children || []).forEach((c) => {
      if (c == null) return;
      node.appendChild(typeof c === 'string' ? document.createTextNode(c) : c);
    });
    return node;
  }

  const INLINE = /(`[^`]+`)|(\*\*[^*]+\*\*)|(\[[^\]]+\]\(https:\/\/[^)\s]+\))/g;

  // Minimal, DOM-only inline formatter: `code`, **bold**, [text](https://...). No innerHTML.
  function inline(parent, text) {
    let last = 0;
    String(text).replace(INLINE, (match, code, bold, link, offset) => {
      if (offset > last) parent.appendChild(document.createTextNode(text.slice(last, offset)));
      if (code) parent.appendChild(el('code', { text: code.slice(1, -1) }));
      else if (bold) parent.appendChild(el('strong', { text: bold.slice(2, -2) }));
      else {
        const m = /^\[([^\]]+)\]\((https:\/\/[^)\s]+)\)$/.exec(link);
        parent.appendChild(el('a', { href: m[2], text: m[1], target: '_blank', rel: 'noopener noreferrer' }));
      }
      last = offset + match.length;
      return match;
    });
    if (last < text.length) parent.appendChild(document.createTextNode(text.slice(last)));
    return parent;
  }

  function plainText(text) {
    return String(text).replace(/\[([^\]]+)\]\((https:\/\/[^)\s]+)\)/g, '$1 ($2)').replace(/\*\*([^*]+)\*\*/g, '$1');
  }

  function download(filename, text, mime) {
    const blob = new Blob([text], { type: (mime || 'text/plain') + ';charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = el('a', { href: url, download: filename });
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  async function copy(text) {
    try {
      await navigator.clipboard.writeText(text);
      return true;
    } catch (e) {
      const ta = el('textarea', { 'aria-hidden': 'true', class: 'sr-copy' });
      ta.value = text;
      document.body.appendChild(ta);
      ta.select();
      let ok = false;
      try { ok = document.execCommand('copy'); } catch (err) { ok = false; }
      ta.remove();
      return ok;
    }
  }

  global.U = { clean, psq, yq, validate, cmd, el, inline, plainText, download, copy };
})(window);
