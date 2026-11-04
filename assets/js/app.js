/* UI: routing, forms, plan rendering. */
(function () {
  'use strict';
  const { el, inline, validate, copy, download, plainText } = window.U;
  const SG = window.SG;
  const STATE_KEY = 'sgw-state-v1';
  const DONE_KEY = 'sgw-done-v1';
  const main = document.getElementById('main');

  /* ---------- state ---------- */
  function readStore(key, fallback) {
    try { return JSON.parse(sessionStorage.getItem(key)) || fallback; } catch (e) { return fallback; }
  }
  function writeStore(key, value) {
    try { sessionStorage.setItem(key, JSON.stringify(value)); } catch (e) { /* storage unavailable */ }
  }
  let state = readStore(STATE_KEY, {});
  let done = readStore(DONE_KEY, {});
  const save = () => writeStore(STATE_KEY, state);

  function applyDefaults(fields) {
    fields.forEach((f) => { if (state[f.key] === undefined && f.def !== undefined) state[f.key] = f.def; });
  }

  /* ---------- routing ---------- */
  function route() {
    const parts = (location.hash || '#/').replace(/^#\/?/, '').split('/').filter(Boolean);
    main.replaceChildren();
    window.scrollTo(0, 0);
    if (parts[0] === 'guided') renderGuided(parts[1]);
    else if (parts[0] === 'q' && parts[1]) renderRecipe(parts[1]);
    else {
      renderHome();
      if (parts[0] === 'quick') document.getElementById('quick').scrollIntoView();
    }
    document.title = 'ScubaGear Wizard' + (parts[0] ? ' - ' + (parts[0] === 'guided' ? 'Guided setup' : 'Quick answer') : '');
    main.focus({ preventScroll: true });
  }
  window.addEventListener('hashchange', route);

  /* ---------- home ---------- */
  function renderHome() {
    const hero = el('section', { class: 'hero' }, [
      el('h1', { text: 'Get ScubaGear running, and passing' }),
      inline(el('p', { class: 'lead' }), 'Answer a few questions about your tenant and this wizard tells you **what to install, what to configure, and exactly what to run**: from first install through fixing failed controls for CISA BOD 25-01.'),
      el('div', { class: 'hero-actions' }, [
        el('a', { class: 'btn btn-primary', href: '#/guided/1', text: 'Start guided setup' }),
        el('a', { class: 'btn', href: '#/quick', text: 'I just need one thing' }),
      ]),
      el('p', { class: 'small muted', text: 'Runs entirely in your browser. Nothing you type is sent anywhere.' }),
    ]);

    const flow = el('section', { class: 'card', 'aria-labelledby': 'flow-h' }, [
      el('h2', { id: 'flow-h', text: 'How the tools fit together' }),
      el('p', { class: 'muted', text: 'Most people move left to right, then loop back through the analyzer until failures are explained.' }),
      toolchainSvg(),
      el('h3', { text: 'Prefer automation? Add ScubaConnect' }),
      el('p', { class: 'muted', text: 'ScubaConnect runs ScubaGear for you on a schedule, hosted by CISA or in your own Azure subscription. You still build and maintain the YAML config.' }),
      automationSvg(),
    ]);

    const grid = el('div', { class: 'grid' }, SG.RECIPES.map((r) => el('a', { class: 'tile', href: '#/q/' + r.id }, [
      el('span', { class: 'tile-title', text: r.title }),
      el('code', { class: 'tile-cmd', text: r.cmd }),
    ])));
    const quick = el('section', { id: 'quick', class: 'card', 'aria-labelledby': 'quick-h' }, [
      el('h2', { id: 'quick-h', text: 'Quick answers' }),
      el('p', { class: 'muted', text: 'Pick the situation that sounds like yours.' }),
      grid,
    ]);

    const bod = el('section', { class: 'card' }, [
      el('h2', { text: 'Before you begin' }),
      list([
        '**Windows and Windows PowerShell 5.1** are required to run ScubaGear.',
        'Know your tenant type (Commercial, GCC, GCC High, or DoD) and which products you must assess.',
        'Have an account with at least **Global Reader**, or plan to create a service principal for unattended runs.',
        'Exclusions, omissions, and annotations must be approved by your organization\'s risk process.',
      ]),
    ]);
    main.append(hero, flow, quick, bod);
  }

  function list(items) {
    return el('ul', null, items.map((t) => inline(el('li'), t)));
  }

  function svgKit(viewBox, ariaLabel, markerId) {
    const NS = 'http://www.w3.org/2000/svg';
    const svg = document.createElementNS(NS, 'svg');
    svg.setAttribute('viewBox', viewBox);
    svg.setAttribute('class', 'toolchain');
    svg.setAttribute('role', 'img');
    svg.setAttribute('aria-label', ariaLabel);
    const mk = (tag, attrs, text) => {
      const n = document.createElementNS(NS, tag);
      Object.keys(attrs).forEach((k) => n.setAttribute(k, attrs[k]));
      if (text) n.textContent = text;
      return n;
    };
    const defs = mk('defs', {});
    const marker = mk('marker', { id: markerId, viewBox: '0 0 10 10', refX: '9', refY: '5', markerWidth: '7', markerHeight: '7', orient: 'auto-start-reverse' });
    marker.appendChild(mk('path', { d: 'M0,0 L10,5 L0,10 z', class: 'tc-arrow' }));
    defs.appendChild(marker);
    svg.appendChild(defs);
    const box = (x, y, w, lines, dashed) => {
      const g = mk('g', {});
      g.appendChild(mk('rect', { x, y, width: w, height: 48, rx: 4, class: dashed ? 'tc-box tc-dashed' : 'tc-box' }));
      lines.forEach((t, i) => g.appendChild(mk('text', { x: x + w / 2, y: y + (lines.length === 1 ? 29 : 21 + i * 15), 'text-anchor': 'middle', class: 'tc-text' }, t)));
      svg.appendChild(g);
    };
    const path = (d) => svg.appendChild(mk('path', { d, class: 'tc-line', 'marker-end': 'url(#' + markerId + ')' }));
    const label = (x, y, t) => svg.appendChild(mk('text', { x, y, 'text-anchor': 'middle', class: 'tc-label' }, t));
    return { svg, box, path, label };
  }

  function automationSvg() {
    const k = svgKit('0 0 1050 150', 'Your YAML config is uploaded to ScubaConnect, which is CISA-hosted or self-hosted. ScubaConnect scans on a schedule and writes ScubaResults.json to CyberScope or your Azure storage. From there you can analyze the results with Start-SCuBAConfigAnalyzer or Invoke-SCuBADiff, and self-hosted deployments can feed a Power BI report.', 'arr2');
    k.path('M160,70 L230,70'); k.label(195, 60, 'upload YAML');
    k.path('M440,70 L510,70'); k.label(475, 60, 'scheduled scans');
    k.path('M760,60 L830,34'); k.path('M760,80 L830,110'); k.label(795, 76, 'results');
    k.box(10, 46, 150, ['Config YAML']);
    k.box(230, 46, 210, ['ScubaConnect', 'CISA-hosted or self-hosted']);
    k.box(510, 46, 250, ['ScubaResults.json', 'CyberScope or your Azure storage']);
    k.box(830, 10, 200, ['Config Analyzer / Diff']);
    k.box(830, 86, 200, ['Power BI report', '(self-hosted)'], true);
    return k.svg;
  }

  function toolchainSvg() {
    const k = svgKit('0 0 1050 270', 'Show-SCuBABaselinePolicyViewer informs Start-SCuBAConfigApp or New-SCuBAConfig, which produces a config YAML for Invoke-SCuBA. Invoke-SCuBA produces an HTML report and ScubaResults.json for Start-SCuBAConfigAnalyzer, which feeds justified exclusions back to the config tool. Two results files can be compared with Invoke-SCuBADiff.', 'arr');
    const { svg, box, path, label } = k;
    path('M220,130 L300,130'); label(260, 120, 'look up a policy');
    path('M490,130 L580,130'); label(535, 120, 'config YAML');
    path('M720,112 L860,38'); label(795, 60, 'BaselineReports.html');
    path('M720,130 L860,130'); label(790, 122, 'ScubaResults.json');
    path('M720,148 L860,232'); label(795, 205, 'two results files');
    path('M860,150 C760,240 520,240 400,156'); label(630, 244, 'justified exclusions');
    box(10, 106, 210, ['Show-SCuBABaselinePolicyViewer']);
    box(300, 106, 190, ['Start-SCuBAConfigApp', 'or New-SCuBAConfig']);
    box(580, 106, 140, ['Invoke-SCuBA']);
    box(860, 14, 180, ['HTML Report']);
    box(860, 106, 180, ['Start-SCuBAConfigAnalyzer']);
    box(860, 206, 180, ['Invoke-SCuBADiff']);
    return svg;
  }

  /* ---------- form fields ---------- */
  function fieldNode(f, onChange) {
    const wrap = el('div', { class: 'field field-' + f.type, 'data-key': f.key });
    const helpId = 'help-' + f.key;
    const help = f.help ? el('p', { class: 'help', id: helpId, text: f.help }) : null;

    if (f.type === 'cards' || f.type === 'multi') {
      const fs = el('fieldset', { 'aria-describedby': help ? helpId : null }, [el('legend', { text: f.label })]);
      if (help) fs.appendChild(help);
      const box = el('div', { class: 'choices' });
      f.options.forEach((o, i) => {
        const id = 'f-' + f.key + '-' + i;
        const multi = f.type === 'multi';
        const input = el('input', { type: multi ? 'checkbox' : 'radio', name: f.key, id, value: o.value });
        input.checked = multi ? (state[f.key] || []).includes(o.value) : state[f.key] === o.value;
        input.addEventListener('change', () => {
          if (multi) {
            const set = new Set(state[f.key] || []);
            if (input.checked) set.add(o.value); else set.delete(o.value);
            state[f.key] = Array.from(set);
          } else {
            state[f.key] = o.value;
          }
          save();
          onChange(f);
        });
        box.appendChild(el('label', { class: 'choice', for: id }, [
          input,
          el('span', { class: 'choice-body' }, [el('span', { class: 'choice-title', text: o.label }), o.desc ? el('span', { class: 'choice-desc', text: o.desc }) : null]),
        ]));
      });
      fs.appendChild(box);
      wrap.appendChild(fs);
    } else if (f.type === 'toggle') {
      const id = 'f-' + f.key;
      const input = el('input', { type: 'checkbox', id, 'aria-describedby': help ? helpId : null });
      input.checked = !!state[f.key];
      input.addEventListener('change', () => { state[f.key] = input.checked; save(); onChange(f); });
      wrap.append(el('label', { class: 'toggle', for: id }, [input, el('span', { text: f.label })]));
      if (help) wrap.appendChild(help);
    } else {
      const id = 'f-' + f.key;
      const msgId = 'msg-' + f.key;
      const input = f.multiline ? el('textarea', { rows: 3 }) : el('input', { type: 'text', autocomplete: 'off', spellcheck: 'false' });
      input.id = id;
      input.name = f.key;
      input.placeholder = f.placeholder || '';
      input.setAttribute('aria-describedby', [help ? helpId : '', msgId].filter(Boolean).join(' '));
      input.value = state[f.key] || '';
      const msg = el('p', { class: 'error', id: msgId, 'aria-live': 'polite' });
      const check = () => {
        const r = validate(f.kind || 'any', input.value);
        msg.textContent = r.ok ? '' : r.message;
        input.toggleAttribute('aria-invalid', !r.ok);
      };
      input.addEventListener('input', () => { state[f.key] = input.value; save(); check(); onChange(f, true); });
      wrap.append(el('label', { for: id, text: f.label }), input);
      if (help) wrap.appendChild(help);
      wrap.appendChild(msg);
      check();
    }
    return wrap;
  }

  function buildForm(fields, onChange) {
    applyDefaults(fields);
    const form = el('form', { class: 'form', novalidate: true, onsubmit: (e) => e.preventDefault() });
    const nodes = fields.map((f) => ({ f, node: fieldNode(f, onChange2) }));
    function updateVisibility() {
      nodes.forEach(({ f, node }) => { node.hidden = f.when ? !f.when(state) : false; });
    }
    function onChange2(f, typing) { updateVisibility(); onChange(f, typing); }
    nodes.forEach(({ node }) => form.appendChild(node));
    updateVisibility();
    return form;
  }

  /* ---------- guided wizard ---------- */
  function renderGuided(stepArg) {
    const visible = () => SG.GUIDED.filter((s) => !s.when || s.when(state));
    if (stepArg === 'plan') return renderPlanPage(SG.buildGuided(state), '#/guided/' + visible().length, 'Change my answers');
    const first = visible();
    const n = Math.min(Math.max(parseInt(stepArg, 10) || 1, 1), first.length);
    const step = first[n - 1];

    const stepper = el('ol', { class: 'stepper', 'aria-label': 'Progress' });
    const eyebrow = el('p', { class: 'eyebrow' });
    const nav = el('div', { class: 'nav' });
    // Later steps appear or disappear as answers change, so rebuild the chrome from the current state.
    const refreshChrome = () => {
      const steps = visible();
      const idx = Math.max(steps.indexOf(step), 0);
      const pos = idx + 1;
      const last = pos === steps.length;
      stepper.replaceChildren(...steps.map((s, i) => el('li', { class: i === idx ? 'current' : i < idx ? 'past' : '', 'aria-current': i === idx ? 'step' : null }, [
        el('a', { href: '#/guided/' + (i + 1) }, [el('span', { class: 'dot', text: String(i + 1) }), el('span', { class: 'step-name', text: s.title })]),
      ])));
      eyebrow.textContent = 'Step ' + pos + ' of ' + steps.length;
      nav.replaceChildren(
        pos > 1 ? el('a', { class: 'btn', href: '#/guided/' + (pos - 1), text: 'Back' }) : el('a', { class: 'btn', href: '#/', text: 'Home' }),
        el('a', { class: 'btn btn-primary', href: last ? '#/guided/plan' : '#/guided/' + (pos + 1), text: last ? 'Build my plan' : 'Next' }),
      );
    };

    const form = buildForm(step.fields, refreshChrome);
    refreshChrome();
    main.append(
      stepper,
      el('section', { class: 'card' }, [
        eyebrow,
        el('h1', { text: step.title }),
        el('p', { class: 'lead', text: step.intro }),
        form, nav,
      ]),
    );
  }

  /* ---------- recipe page ---------- */
  function renderRecipe(id) {
    const r = SG.RECIPES.find((x) => x.id === id);
    if (!r) { location.hash = '#/'; return; }
    const result = el('div', { 'aria-live': 'polite' });
    let timer;
    const refresh = () => {
      result.replaceChildren(planNode(r.build(state), { prefix: 'q-' + r.id }));
    };
    const form = r.fields.length ? buildForm(r.fields, (f, typing) => {
      clearTimeout(timer);
      timer = setTimeout(refresh, typing ? 250 : 0);
    }) : null;
    main.append(
      el('p', { class: 'crumbs' }, [el('a', { href: '#/', text: 'Home' }), ' / Quick answers']),
      el('section', { class: 'card' }, [
        el('h1', { text: r.title }),
        el('code', { class: 'tile-cmd', text: r.cmd }),
        form ? el('p', { class: 'muted', text: 'Answer these to tailor the commands below. All fields are optional.' }) : null,
        form,
      ]),
      result,
      el('p', null, [el('a', { class: 'btn', href: '#/guided/1', text: 'Need the whole journey? Start guided setup' })]),
    );
    refresh();
  }

  /* ---------- plan rendering ---------- */
  function renderPlanPage(plan, editHref, editLabel) {
    const fed = el('input', { type: 'checkbox', id: 'plan-federal' });
    fed.checked = !!state.federal;
    fed.addEventListener('change', () => {
      state.federal = fed.checked;
      save();
      route();
      const again = document.getElementById('plan-federal');
      if (again) again.focus();
    });
    const bar = el('div', { class: 'toolbar no-print' }, [
      el('a', { class: 'btn', href: editHref, text: editLabel }),
      el('button', { class: 'btn', type: 'button', onclick: () => download('scubagear-plan.md', toMarkdown(plan), 'text/markdown'), text: 'Download plan (.md)' }),
      el('button', { class: 'btn', type: 'button', onclick: () => window.print(), text: 'Print' }),
      el('button', { class: 'btn', type: 'button', onclick: () => { state = {}; done = {}; save(); writeStore(DONE_KEY, done); location.hash = '#/'; }, text: 'Start over' }),
      el('label', { class: 'toggle', for: 'plan-federal' }, [fed, el('span', { text: 'Federal agency (adds the quarterly CyberScope step)' })]),
    ]);
    main.append(bar, planNode(plan, { prefix: 'g', toc: true }));
    if (plan.jumpTo) {
      const t = document.getElementById('g-' + plan.jumpTo);
      if (t) t.scrollIntoView();
    }
  }

  function planNode(plan, opts) {
    const root = el('div', { class: 'plan' });
    const header = el('header', { class: 'plan-head card' }, [
      el('h1', { text: plan.title }),
      plan.intro ? el('p', { class: 'lead', text: plan.intro }) : null,
    ]);
    if (plan.jumpTo) {
      const sec = plan.sections.find((s) => s.id === plan.jumpTo);
      header.appendChild(el('p', { class: 'note note-info' }, [
        'You already have failing results. ',
        el('a', { href: '#g-' + sec.id, onclick: (e) => { e.preventDefault(); document.getElementById('g-' + sec.id).scrollIntoView(); }, text: 'Jump to "' + sec.title + '"' }),
        '.',
      ]));
    }
    root.appendChild(header);

    const body = el('div', { class: 'plan-body' });
    plan.sections.forEach((s, i) => {
      const sid = opts.prefix + '-' + s.id;
      const doneKey = opts.prefix + ':' + s.id;
      const cb = el('input', { type: 'checkbox', id: 'done-' + sid, 'aria-label': 'Mark "' + s.title + '" as done' });
      cb.checked = !!done[doneKey];
      const section = el('section', { class: 'card step' + (s.highlight ? ' highlight' : ''), id: sid, 'aria-labelledby': 'h-' + sid }, [
        el('div', { class: 'step-head' }, [
          el('span', { class: 'step-num', text: String(i + 1) }),
          el('h2', { id: 'h-' + sid, text: s.title }),
          el('label', { class: 'done-label no-print', for: 'done-' + sid }, [cb, 'Done']),
        ]),
        s.highlight ? el('p', { class: 'badge', text: s.highlight }) : null,
        ...s.blocks.map(blockNode),
      ]);
      cb.addEventListener('change', () => {
        done[doneKey] = cb.checked;
        writeStore(DONE_KEY, done);
        section.classList.toggle('is-done', cb.checked);
      });
      section.classList.toggle('is-done', cb.checked);
      body.appendChild(section);
    });

    if (opts.toc) {
      const toc = el('nav', { class: 'toc no-print', 'aria-label': 'Plan sections' }, [
        el('p', { class: 'toc-title', text: 'On this page' }),
        el('ol', null, plan.sections.map((s) => el('li', null, [el('a', { href: '#' + opts.prefix + '-' + s.id, onclick: (e) => { e.preventDefault(); document.getElementById(opts.prefix + '-' + s.id).scrollIntoView(); }, text: s.title })]))),
      ]);
      root.appendChild(el('div', { class: 'plan-layout' }, [toc, body]));
    } else {
      root.appendChild(body);
    }
    return root;
  }

  function codeNode(text, caption, filename) {
    const placeholder = /<[A-Za-z][\w.-]*>|\bTODO\b|vX\.Y\.Z/.test(text);
    const status = el('span', { class: 'copy-status', role: 'status' });
    const btn = el('button', { class: 'btn btn-small no-print', type: 'button', text: 'Copy', 'aria-label': 'Copy command to clipboard' });
    btn.addEventListener('click', async () => {
      const ok = await copy(text);
      status.textContent = ok ? 'Copied' : 'Copy failed';
      setTimeout(() => { status.textContent = ''; }, 1800);
    });
    const actions = [btn];
    if (filename) actions.push(el('button', { class: 'btn btn-small no-print', type: 'button', text: 'Download', onclick: () => download(filename, text, 'text/yaml') }));
    return el('figure', { class: 'code' }, [
      el('div', { class: 'code-bar' }, [
        el('span', { class: 'code-lang', text: filename || 'PowerShell' }),
        placeholder ? el('span', { class: 'badge badge-warn', text: 'Replace the <placeholders>' }) : null,
        el('span', { class: 'spacer' }), status, ...actions,
      ]),
      el('pre', { tabindex: '0' }, [el('code', { text })]),
      caption ? inline(el('figcaption'), caption) : null,
    ]);
  }

  function blockNode(b) {
    switch (b.type) {
      case 'p': return inline(el('p'), b.text);
      case 'note': return inline(el('div', { class: 'note note-' + b.kind, role: 'note' }, [el('strong', { class: 'note-label', text: b.kind === 'warn' ? 'Important: ' : 'Note: ' })]), b.text);
      case 'code': return codeNode(b.text, b.caption);
      case 'file': return codeNode(b.text, b.caption, b.filename);
      case 'list': return list(b.items);
      case 'steps': return el('ol', { class: 'num' }, b.items.map((t) => inline(el('li'), t)));
      case 'table': return el('div', { class: 'table-wrap' }, [el('table', null, [
        el('thead', null, [el('tr', null, b.head.map((h) => el('th', { scope: 'col', text: h })))]),
        el('tbody', null, b.rows.map((r) => el('tr', null, r.map((c) => inline(el('td'), c))))),
      ])]);
      case 'links': return el('ul', { class: 'links' }, b.items.map((l) => el('li', null, [el('a', { href: l.url, target: '_blank', rel: 'noopener noreferrer', text: l.label + ' (docs)' })])));
      case 'details': return el('details', null, [el('summary', { text: b.summary }), ...b.blocks.map(blockNode)]);
      default: return document.createTextNode('');
    }
  }

  /* ---------- markdown export ---------- */
  function toMarkdown(plan) {
    const out = ['# ' + plan.title, '', plan.intro || '', ''];
    const render = (b) => {
      switch (b.type) {
        case 'p': return plainText(b.text) + '\n';
        case 'note': return '> **' + (b.kind === 'warn' ? 'Important' : 'Note') + ':** ' + plainText(b.text) + '\n';
        case 'code': return '```powershell\n' + b.text + '\n```\n' + (b.caption ? plainText(b.caption) + '\n' : '');
        case 'file': return '`' + b.filename + '`\n\n```yaml\n' + b.text + '```\n';
        case 'list': return b.items.map((t) => '- ' + plainText(t)).join('\n') + '\n';
        case 'steps': return b.items.map((t, i) => (i + 1) + '. ' + plainText(t)).join('\n') + '\n';
        case 'table': return '| ' + b.head.join(' | ') + ' |\n| ' + b.head.map(() => '---').join(' | ') + ' |\n' + b.rows.map((r) => '| ' + r.map(plainText).join(' | ') + ' |').join('\n') + '\n';
        case 'links': return b.items.map((l) => '- [' + l.label + '](' + l.url + ')').join('\n') + '\n';
        case 'details': return '**' + b.summary + '**\n\n' + b.blocks.map(render).join('\n');
        default: return '';
      }
    };
    plan.sections.forEach((s, i) => {
      out.push('## ' + (i + 1) + '. ' + s.title, '');
      s.blocks.forEach((b) => out.push(render(b)));
    });
    return out.join('\n');
  }

  route();
})();
