(function () {
  'use strict';

  var I18N = {
    en: {
      skip: 'Skip to builder',
      heroTitle: 'Install official Office LTSC in any language with one command',
      heroLead: 'Pick the version, language and apps, then copy the PowerShell command. Office is downloaded straight from Microsoft servers using the official Office Deployment Tool (ODT), after its digital signature is verified.',
      quickLabel: 'Fastest way (interactive menu):',
      copy: 'Copy',
      copied: 'Copied',
      builderTitle: 'Command builder',
      version: 'Version',
      language: 'Language',
      langSearch: 'Search by name or code... (arab, fr-fr, Deutsch)',
      primary: 'Main language:',
      addExtra: '+ Add as extra language',
      setMain: 'Set as main language',
      remove: 'Remove',
      arch: 'Architecture',
      edition: 'Edition',
      addons: 'Extra products',
      exclude: 'Exclude apps (they will not be installed)',
      existing: 'Office already on this PC',
      rmNone: 'Keep it as is',
      rmOdt: 'Remove it with Microsoft\'s tool before installing',
      rmScrub: 'Deep clean (OfficeScrubber), then remove',
      warnOdt: 'Every Click-to-Run and MSI version of Office on the PC (including Microsoft 365) will be removed before the install. Save your work first.',
      warnScrub: 'OfficeScrubber (by abbodi1406, open source on GitHub) runs first. It force-removes broken or stubborn Office installs and also removes installed Office licenses and keys. It is downloaded from a pinned commit and every file is checked by SHA-256 before it runs. It can take 10-30 minutes.',
      advanced: 'Advanced options',
      noUpdates: 'Turn off automatic updates',
      silent: 'Silent install (no Office setup window)',
      downloadOnly: 'Download only to a folder (for offline install)',
      yourCmd: 'Your command',
      keyNote: '🔑 The product key is never put in the command. Run the simple command above and the script asks for the key with hidden input, or add -ProductKey yourself inside the PowerShell window.',
      howTitle: 'How to run',
      how1: 'Right-click the Start button and choose "Terminal (Admin)" or "Windows PowerShell (Admin)".',
      how2: 'Paste the command and press Enter.',
      how3: 'If the window is not elevated, the script opens a new Administrator window after you approve the UAC prompt.',
      how4: 'Wait for the install to finish. A full log is saved on the desktop as OfficeLTSC_Install_Log.txt.',
      offTitle: 'Offline install',
      offText: 'Download the Office files once into a folder, then copy it to any PC (or burn it to a disc/ISO) and install from it:',
      offText2: 'The folder will contain setup.exe and configuration.xml. On the other PC, run this from an Administrator window inside the folder:',
      verTitle: 'Verify the script',
      verText: 'SHA-256 of the currently published install.ps1:',
      verText2: 'Download the file and compare the hash and the digital signature:',
      viewSource: 'View the formatted source code on GitHub',
      viewSecurity: 'What the script downloads and how it verifies it',
      cacheNote: 'Note: a new script release can take up to 10 minutes to appear because of GitHub Pages caching.',
      footer: 'Independent tool. Not affiliated with Microsoft. Office is downloaded directly from Microsoft\'s servers. Microsoft and Office are trademarks of Microsoft Corporation.',
      repo: 'Source repository',
      toggle: 'العربية',
      matchOs: 'Match Windows language'
    },
    ar: {
      copied: 'تم النسخ',
      remove: 'إزالة',
      warnOdt: 'ستُزال كل إصدارات Office الموجودة على الجهاز (Click-to-Run وMSI، بما فيها Microsoft 365) قبل التثبيت. احفظ عملك أولًا.',
      warnScrub: 'تعمل أداة OfficeScrubber (من abbodi1406، مفتوحة المصدر على GitHub) أولًا لإزالة نسخ Office التالفة أو العالقة بالقوة، وتزيل أيضًا تراخيص ومفاتيح Office المثبتة. تُنزَّل من نسخة مثبّتة (commit) ويُتحقق من بصمة SHA-256 لكل ملف قبل التشغيل. قد تستغرق 10–30 دقيقة.',
      toggle: 'English',
      matchOs: 'لغة Windows الحالية'
    }
  };

  var state = {
    ui: 'ar',
    products: null,
    languages: [],
    version: '2024',
    language: 'MatchOS',
    extra: []
  };

  var $ = function (id) { return document.getElementById(id); };
  var base = (location.origin && location.origin.indexOf('http') === 0) ? location.origin : 'https://q123qrc-source.github.io';

  // Remember the Arabic strings that are written in the HTML.
  document.querySelectorAll('[data-i18n]').forEach(function (el) {
    var k = el.getAttribute('data-i18n');
    if (!(k in I18N.ar)) { I18N.ar[k] = el.textContent; }
  });
  document.querySelectorAll('[data-i18n-ph]').forEach(function (el) {
    var k = el.getAttribute('data-i18n-ph');
    if (!(k in I18N.ar)) { I18N.ar[k] = el.getAttribute('placeholder'); }
  });

  function t(k) { return (I18N[state.ui] && I18N[state.ui][k]) || I18N.ar[k] || k; }

  function applyUi() {
    var html = document.documentElement;
    html.lang = state.ui;
    html.dir = state.ui === 'ar' ? 'rtl' : 'ltr';
    document.querySelectorAll('[data-i18n]').forEach(function (el) { el.textContent = t(el.getAttribute('data-i18n')); });
    document.querySelectorAll('[data-i18n-ph]').forEach(function (el) { el.setAttribute('placeholder', t(el.getAttribute('data-i18n-ph'))); });
    $('langToggle').textContent = t('toggle');
    renderLanguages();
    renderChips();
    updateRemoveWarn();
    try { localStorage.setItem('ui', state.ui); } catch (e) { /* storage unavailable */ }
  }

  function psQuote(s) { return "'" + String(s).replace(/'/g, "''") + "'"; }

  function radioValue(name) {
    var el = document.querySelector('input[name="' + name + '"]:checked');
    return el ? el.value : '';
  }

  function buildCommand() {
    var args = ['-Version ' + state.version, '-Language ' + state.language];
    if (state.extra.length) { args.push('-ExtraLanguages ' + state.extra.join(',')); }
    var arch = radioValue('arch');
    if (arch !== '64') { args.push('-Arch ' + arch); }
    var ed = radioValue('edition');
    if (ed !== 'ProPlus') { args.push('-Edition ' + ed); }
    if ($('visio').checked) { args.push('-Visio'); }
    if ($('project').checked) { args.push('-Project'); }
    var ex = Array.prototype.slice.call(document.querySelectorAll('#appsGroup input:checked')).map(function (i) { return i.value; });
    if (ex.length) { args.push('-ExcludeApps ' + ex.join(',')); }
    var rm = $('remove').value;
    var dl = $('downloadOnly').checked;
    if (!dl && rm === 'odt') { args.push('-RemoveExisting'); }
    if (!dl && rm === 'scrub') { args.push('-RemoveExisting -Scrub'); }
    if ($('noUpdates').checked) { args.push('-NoUpdates'); }
    if ($('silent').checked) { args.push('-Silent'); }
    if (dl) {
      var p = $('sourcePath').value.trim() || 'D:\\OfficeSource';
      args.push('-DownloadOnly -SourcePath ' + psQuote(p));
    }
    return '& ([scriptblock]::Create((irm ' + base + '/ps))) ' + args.join(' ');
  }

  function update() {
    $('cmdSimple').textContent = 'irm ' + base + '/ps | iex';
    $('cmdCustom').textContent = buildCommand();
    $('cmdOffline').textContent = '& ([scriptblock]::Create((irm ' + base + '/ps))) -Version ' + state.version +
      ' -Language ' + (state.language === 'MatchOS' ? 'en-us' : state.language) + ' -DownloadOnly -SourcePath ' + psQuote($('sourcePath').value.trim() || 'D:\\OfficeSource');
    $('cmdVerify').textContent =
      'irm ' + base + '/install.txt -OutFile install.ps1\n' +
      'Get-FileHash .\\install.ps1 -Algorithm SHA256\n' +
      'Get-AuthenticodeSignature .\\install.ps1 | Format-List Status, SignerCertificate';
    $('langPrimary').textContent = state.language;
  }

  function renderVersions() {
    var g = $('versionGroup');
    g.textContent = '';
    state.products.versions.forEach(function (v) {
      var label = document.createElement('label');
      var input = document.createElement('input');
      input.type = 'radio'; input.name = 'version'; input.value = v.version;
      input.checked = v.version === state.version;
      input.addEventListener('change', function () { state.version = v.version; renderApps(); update(); });
      var span = document.createElement('span');
      span.textContent = v.name;
      label.appendChild(input); label.appendChild(span);
      g.appendChild(label);
    });
  }

  function renderApps() {
    var g = $('appsGroup');
    var checked = {};
    g.querySelectorAll('input:checked').forEach(function (i) { checked[i.value] = true; });
    g.textContent = '';
    var v = state.products.versions.filter(function (x) { return x.version === state.version; })[0];
    v.apps.forEach(function (a) {
      var label = document.createElement('label');
      label.className = 'check';
      var input = document.createElement('input');
      input.type = 'checkbox'; input.value = a; input.checked = !!checked[a];
      input.addEventListener('change', update);
      var span = document.createElement('span');
      span.textContent = a === 'Lync' ? 'Skype for Business (Lync)' : a;
      label.appendChild(input); label.appendChild(span);
      g.appendChild(label);
    });
  }

  function langLabel(l) {
    if (l.id === 'MatchOS') { return t('matchOs') + '  ·  MatchOS'; }
    var name = state.ui === 'ar' ? l.name : l.english;
    var other = state.ui === 'ar' ? l.english : l.name;
    return name + (other && other !== name ? ' — ' + other : '') + '  ·  ' + l.id;
  }

  function renderLanguages() {
    var sel = $('langSelect');
    if (!state.languages.length) { return; }
    var q = $('langSearch').value.trim().toLowerCase();
    var selected = sel.value;
    sel.textContent = '';
    state.languages.forEach(function (l) {
      var hay = (l.id + ' ' + l.name + ' ' + l.english + (l.id === 'MatchOS' ? ' ' + t('matchOs') : '')).toLowerCase();
      if (q && hay.indexOf(q) === -1) { return; }
      var o = document.createElement('option');
      o.value = l.id;
      o.textContent = langLabel(l);
      o.selected = l.id === (selected || state.language);
      sel.appendChild(o);
    });
  }

  function renderChips() {
    var box = $('extraChips');
    box.textContent = '';
    state.extra.forEach(function (id) {
      var chip = document.createElement('span');
      chip.className = 'chip';
      chip.dir = 'ltr';
      chip.appendChild(document.createTextNode(id));
      var b = document.createElement('button');
      b.type = 'button';
      b.setAttribute('aria-label', t('remove') + ' ' + id);
      b.textContent = '×';
      b.addEventListener('click', function () {
        state.extra = state.extra.filter(function (x) { return x !== id; });
        renderChips(); update();
      });
      chip.appendChild(b);
      box.appendChild(chip);
    });
  }

  function updateRemoveWarn() {
    var v = $('remove').value;
    var w = $('removeWarn');
    if (v === 'none') { w.hidden = true; return; }
    w.hidden = false;
    w.textContent = v === 'scrub' ? t('warnScrub') + ' ' + t('warnOdt') : t('warnOdt');
  }

  function toast(msg) {
    var el = $('toast');
    el.textContent = msg;
    el.classList.add('show');
    clearTimeout(toast.timer);
    toast.timer = setTimeout(function () { el.classList.remove('show'); }, 1600);
  }

  function copyText(text) {
    if (navigator.clipboard && window.isSecureContext) {
      return navigator.clipboard.writeText(text);
    }
    return new Promise(function (resolve, reject) {
      var ta = document.createElement('textarea');
      ta.value = text; ta.setAttribute('readonly', '');
      ta.style.position = 'fixed'; ta.style.opacity = '0';
      document.body.appendChild(ta); ta.select();
      try { document.execCommand('copy') ? resolve() : reject(new Error('copy failed')); } catch (e) { reject(e); }
      document.body.removeChild(ta);
    });
  }

  function wire() {
    document.querySelectorAll('[data-copy]').forEach(function (b) {
      b.addEventListener('click', function () {
        copyText($(b.getAttribute('data-copy')).textContent).then(function () { toast(t('copied')); }, function () { /* ignore */ });
      });
    });
    $('langToggle').addEventListener('click', function () { state.ui = state.ui === 'ar' ? 'en' : 'ar'; applyUi(); });
    $('langSearch').addEventListener('input', renderLanguages);
    var setMain = function () {
      var id = $('langSelect').value;
      if (!id) { return; }
      state.language = id;
      state.extra = state.extra.filter(function (x) { return x !== id; });
      renderChips(); update();
    };
    $('setMain').addEventListener('click', setMain);
    $('langSelect').addEventListener('dblclick', setMain);
    $('langSelect').addEventListener('keydown', function (e) { if (e.key === 'Enter') { e.preventDefault(); setMain(); } });
    $('addExtra').addEventListener('click', function () {
      var id = $('langSelect').value;
      if (!id || id === 'MatchOS' || id === state.language || state.extra.indexOf(id) !== -1) { return; }
      state.extra.push(id); renderChips(); update();
    });
    ['visio', 'project', 'noUpdates', 'silent'].forEach(function (id) { $(id).addEventListener('change', update); });
    document.querySelectorAll('input[name="arch"], input[name="edition"]').forEach(function (i) { i.addEventListener('change', update); });
    $('remove').addEventListener('change', function () { updateRemoveWarn(); update(); });
    $('downloadOnly').addEventListener('change', function () { $('sourcePath').hidden = !this.checked; update(); });
    $('sourcePath').addEventListener('input', update);
  }

  function load(url) {
    return fetch(url, { cache: 'no-cache' }).then(function (r) {
      if (!r.ok) { throw new Error(url + ' ' + r.status); }
      return r.json();
    });
  }

  // Arabic is the default; remember the visitor's choice when storage works.
  try { var saved = localStorage.getItem('ui'); if (saved === 'en' || saved === 'ar') { state.ui = saved; } } catch (e) { /* ignore */ }

  if ($('sha').textContent.indexOf('{{') === 0) { $('sha').textContent = '(computed at deploy time)'; }

  wire();
  Promise.all([load('products.json'), load('languages.json')]).then(function (res) {
    state.products = res[0];
    state.languages = res[1];
    renderVersions();
    renderApps();
    applyUi();
    update();
  }).catch(function (err) {
    $('cmdCustom').textContent = 'Error loading data: ' + err.message;
  });
  applyUi();
  update();
})();
