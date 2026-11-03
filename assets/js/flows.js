/* Content and logic: questions, recipes, and plan generation. All commands come from the ScubaGear docs. */
(function (global) {
  'use strict';
  const { psq, yq, validate, cmd } = global.U;

  const DOCS = 'https://github.com/cisagov/ScubaGear/blob/main/docs/';
  const link = (path, label) => ({ label, url: DOCS + path });

  const ENVIRONMENTS = [
    { value: 'commercial', label: 'Commercial', desc: 'Standard Microsoft 365 (non-government cloud).' },
    { value: 'gcc', label: 'GCC', desc: 'Government Community Cloud.' },
    { value: 'gcchigh', label: 'GCC High', desc: 'Government Community Cloud High.' },
    { value: 'dod', label: 'DoD', desc: 'Department of Defense tenants.' },
  ];

  const PRODUCTS = [
    { id: 'aad', label: 'Microsoft Entra ID', def: true, role: 'Global Reader' },
    { id: 'securitysuite', label: 'Security Suite', def: true, role: 'Global Reader (or Exchange Administrator)' },
    { id: 'exo', label: 'Exchange Online', def: true, role: 'Global Reader (or Exchange Administrator)' },
    { id: 'sharepoint', label: 'SharePoint Online and OneDrive', def: true, role: 'Global Reader (or SharePoint Administrator)' },
    { id: 'teams', label: 'Microsoft Teams', def: true, role: 'Global Reader (or Teams Administrator)' },
    { id: 'powerplatform', label: 'Power Platform', def: false, role: 'Power Platform Administrator with a "Power Apps for Office 365" license', special: true },
    { id: 'powerbi', label: 'Power BI', def: false, role: 'Fabric Administrator (or Global Administrator) with a Power BI or Fabric license', special: true },
  ];
  const DEFAULT_PRODUCTS = PRODUCTS.filter((p) => p.def).map((p) => p.id);
  const SP_PRODUCTS = ['aad', 'exo', 'sharepoint', 'teams', 'securitysuite', 'powerplatform'];

  const SP_PERMS = {
    aad: 'Microsoft Graph application permissions: Directory.Read.All, Policy.Read.All, PrivilegedAccess.Read.AzureADGroup, PrivilegedEligibilitySchedule.Read.AzureADGroup, RoleManagement.Read.Directory, RoleManagementPolicy.Read.AzureADGroup, User.Read.All',
    securitysuite: 'Office 365 Exchange Online: Exchange.ManageAsApp, plus the Global Reader role',
    exo: 'Office 365 Exchange Online: Exchange.ManageAsApp, plus the Global Reader role',
    sharepoint: 'SharePoint: Sites.FullControl.All (ScubaGear never writes; this is the minimum the API allows)',
    teams: 'Global Reader role',
    powerplatform: 'One-time registration with Power Platform (automated setup handles this)',
    powerbi: 'No API permission. Add the service principal to a security group and enable the Power BI admin tenant setting.',
  };

  const CAP_POLICIES = ['MS.AAD.1.1v1', 'MS.AAD.2.1v1', 'MS.AAD.2.3v1', 'MS.AAD.3.1v1', 'MS.AAD.3.2v2', 'MS.AAD.3.6v1', 'MS.AAD.3.7v1', 'MS.AAD.3.8v1', 'MS.AAD.3.9v1'];

  /* ---------- block constructors ---------- */
  const P = (text) => ({ type: 'p', text });
  const NOTE = (kind, text) => ({ type: 'note', kind, text });
  const CODE = (text, caption) => ({ type: 'code', text, caption });
  const LIST = (items) => ({ type: 'list', items });
  const STEPS = (items) => ({ type: 'steps', items });
  const TABLE = (head, rows) => ({ type: 'table', head, rows });
  const LINKS = (items) => ({ type: 'links', items });
  const DETAILS = (summary, blocks) => ({ type: 'details', summary, blocks });
  const SECTION = (id, title, blocks, extra) => Object.assign({ id, title, blocks }, extra || {});

  /* ---------- state -> normalized context ---------- */
  function ctx(s) {
    const text = (key, kind) => validate(kind, s[key]);
    const products = Array.isArray(s.products) && s.products.length ? s.products.filter((p) => PRODUCTS.some((x) => x.id === p)) : DEFAULT_PRODUCTS;
    const appId = text('appId', 'guid');
    const thumb = text('thumb', 'thumb');
    const tenant = text('tenant', 'domain');
    const out = text('outPath', 'path');
    const cfg = text('configPath', 'path');
    const users = validate('guidlist', s.bgUsers);
    const groups = validate('guidlist', s.bgGroups);
    const tenantLabel = tenant.value ? tenant.value.split('.')[0] : '';
    return {
      s,
      env: ENVIRONMENTS.some((e) => e.value === s.env) ? s.env : 'commercial',
      products,
      productArg: products.length === PRODUCTS.length ? '*' : products.join(', '),
      has: (id) => products.includes(id),
      auth: s.auth === 'sp' ? 'sp' : 'interactive',
      appId: appId.value, thumb: thumb.value, tenant: tenant.value, tenantLabel,
      appIdPh: appId.value || '<app-id>',
      thumbPh: thumb.value || '<certificate-thumbprint>',
      tenantPh: tenant.value || 'contoso.onmicrosoft.com',
      outPath: out.value,
      configPath: cfg.value,
      orgName: global.U.clean(s.orgName), orgUnit: global.U.clean(s.orgUnit),
      users: users.value, groups: groups.value,
      dark: !!s.dark, quiet: !!s.quiet, transcript: !!s.transcript,
    };
  }

  function invokeParts(c, opts) {
    const o = opts || {};
    const parts = [];
    if (o.configPath) parts.push('-ConfigFilePath ' + psq(o.configPath));
    if (!o.configOnly) {
      parts.push('-ProductNames ' + c.productArg);
      if (c.env !== 'commercial' && c.auth !== 'sp') parts.push('-M365Environment ' + c.env);
      if (c.auth === 'sp') {
        parts.push('-AppID ' + c.appIdPh, '-CertificateThumbprint ' + c.thumbPh, '-Organization ' + c.tenantPh);
      }
      if (c.outPath && !o.configPath) parts.push('-OutPath ' + psq(c.outPath));
    }
    if (c.dark) parts.push('-DarkMode');
    if (c.quiet) parts.push('-Quiet');
    if (c.transcript) parts.push('-Transcript');
    return parts;
  }

  /* ---------- config YAML ---------- */
  function buildYaml(c) {
    const L = [];
    L.push('Description: |');
    L.push('  ScubaGear configuration generated by the ScubaGear Wizard. Review every value before use.');
    L.push('  Reference: docs/configuration/configuration.md in the ScubaGear repository.');
    L.push('');
    L.push('ProductNames:');
    c.products.forEach((p) => L.push('  - ' + p));
    L.push('M365Environment: ' + c.env);
    L.push('');
    L.push('# Documents the owner of this tenant (needed for BOD submissions).');
    L.push('OrgName: ' + yq(c.orgName || 'Department of Example') + (c.orgName ? '' : '  # TODO: replace'));
    L.push('OrgUnitName: ' + yq(c.orgUnit || 'Subdepartment of Example') + (c.orgUnit ? '' : '  # TODO: replace'));
    if (c.auth === 'sp') {
      L.push('');
      L.push('# Non-interactive (service principal) authentication.');
      L.push('Organization: ' + yq(c.tenantPh) + (c.tenant ? '' : '  # TODO: initial domain, e.g. contoso.onmicrosoft.com'));
      L.push('AppID: ' + yq(c.appIdPh) + (c.appId ? '' : '  # TODO'));
      L.push('CertificateThumbprint: ' + yq(c.thumbPh) + (c.thumb ? '' : '  # TODO'));
    }
    if (c.outPath) {
      L.push('');
      L.push('OutPath: ' + yq(c.outPath));
    }
    L.push('');
    L.push('# Omit only with approval through your risk management process.');
    L.push('OmitPolicy: {}');
    L.push('');
    L.push('# Annotate failing controls with plans of action (Comment, RemediationDate, IncorrectResult).');
    L.push('AnnotatePolicy: {}');

    const list = (indent, name, values) => {
      if (!values.length) return;
      L.push(indent + name + ':');
      values.forEach((v) => L.push(indent + '  - ' + yq(v)));
    };
    if (c.has('aad') && (c.users.length || c.groups.length)) {
      L.push('');
      L.push('# Exclusions must be approved by your risk acceptance process.');
      L.push('Aad:');
      CAP_POLICIES.forEach((id, i) => {
        if (i === 0) {
          L.push('  ' + id + ': &OrgCapExclusions');
          L.push('    CapExclusions:');
          list('      ', 'Users', c.users);
          list('      ', 'Groups', c.groups);
        } else {
          L.push('  ' + id + ': *OrgCapExclusions');
        }
      });
      L.push('  MS.AAD.7.4v1:');
      L.push('    RoleExclusions:');
      list('      ', 'Users', c.users);
      list('      ', 'Groups', c.groups);
    }
    return L.join('\n') + '\n';
  }

  /* ---------- shared section builders ---------- */
  const ps51Note = () => NOTE('warn', 'ScubaGear runs on **Windows** in **Windows PowerShell 5.1**, not PowerShell 7. Open "Windows PowerShell" from the Start menu (not "PowerShell 7"). Some Microsoft libraries ScubaGear depends on only work in 5.1.');

  function installBlocks(c) {
    const how = c.s.install || 'psgallery';
    if (how === 'installed') {
      return [
        P('You already have ScubaGear. Check that it is current:'),
        CODE('Test-ScubaGearVersion', 'Status "Up to Date" means no action. "Update Available" or "Needs attention" means run the next command.'),
        CODE('Update-ScubaGear', 'Use `Update-ScubaGear -Source GitHub` if PSGallery is blocked on your network.'),
        NOTE('info', 'If you originally downloaded ScubaGear from GitHub, import it in every new session: `Import-Module .\\PowerShell\\ScubaGear` (no trailing backslash).'),
        LINKS([link('installation/update.md', 'Update instructions')]),
      ];
    }
    if (how === 'github') {
      return [
        STEPS([
          'Open the [releases page](https://github.com/cisagov/ScubaGear/releases) and download the latest `ScubaGear-vX.Y.Z.zip` from **Assets**.',
          'Extract the zip to a folder of your choice.',
          'Allow local scripts to run if your execution policy blocks them (see below).',
          'In Windows PowerShell 5.1, go to the extracted folder and import the module.',
        ]),
        CODE('Get-ExecutionPolicy', 'Windows clients default to Restricted, which blocks ScubaGear.'),
        CODE("Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine", 'Run as Administrator. If your organization forbids changing this, ScubaGear cannot run.'),
        CODE('Import-Module .\\PowerShell\\ScubaGear', 'Repeat in every new PowerShell session. Do not add a trailing backslash.'),
        DETAILS('Getting a "mark of the web" warning?', [
          P('Files downloaded from the internet carry a marker. Only after vetting the files, unblock them from inside the ScubaGear folder:'),
          CODE('Get-ChildItem *.ps1 -Recurse | Unblock-File\nGet-ChildItem *.psm1 -Recurse | Unblock-File\nGet-ChildItem *.psd1 -Recurse | Unblock-File'),
        ]),
        LINKS([link('installation/github.md', 'Download from GitHub')]),
      ];
    }
    return [
      P('Installing from the PowerShell Gallery is the recommended method and survives new sessions without extra import steps.'),
      CODE('Install-Module -Name ScubaGear', 'A specific version: add `-RequiredVersion 2.0.0`.'),
      LINKS([link('installation/psgallery.md', 'Install from PSGallery')]),
    ];
  }

  function depsBlocks() {
    return [
      P('This installs the required PowerShell modules and the Open Policy Agent (OPA) engine into a `.scubagear` folder in your home directory.'),
      CODE('Install-ScubaDependencies'),
      CODE('Get-ScubaGearDependencyStatus', 'Optional: verify what is installed.'),
      NOTE('info', '`Install-ScubaDependencies` was previously named `Initialize-SCuBA`; the old name still works as an alias.'),
      LINKS([link('prerequisites/dependencies.md', 'Dependencies')]),
    ];
  }

  function networkBlocks(c) {
    const parts = [];
    if (c.products.length !== PRODUCTS.length) parts.push('-ProductNames ' + c.products.join(', '));
    parts.push('-M365Environment ' + c.env);
    if (c.tenantLabel) parts.push('-Domain ' + c.tenantLabel);
    parts.push('-Format Markdown', "-OutFile .\\scubagear-hosts.md");
    return [
      P('ScubaGear talks to Microsoft services over HTTPS (port 443). Generate the host list for your products and send it to your network team for an allow-list.'),
      CODE(cmd('Get-ScubaGearEndpointRest', parts), 'Writes a table you can paste into a ticket. Re-run after each ScubaGear update.'),
      NOTE('info', 'SharePoint admin hosts include your tenant name. Without `-Domain` they appear as `<tenant>-admin.sharepoint.com`.'),
      DETAILS('Proxy errors?', [
        P('Make PowerShell use your Windows credentials for the proxy:'),
        CODE('$Wcl=New-Object System.Net.WebClient\n$Wcl.Proxy.Credentials=[System.Net.CredentialCache]::DefaultNetworkCredentials'),
        P('Set the proxy itself in your default browser (for example Edge).'),
      ]),
      LINKS([link('prerequisites/networkaccess.md', 'Network access'), link('troubleshooting/proxy.md', 'Proxy troubleshooting')]),
    ];
  }

  function interactivePermBlocks(c) {
    const rows = PRODUCTS.filter((p) => c.has(p.id)).map((p) => [p.label, p.role]);
    const b = [
      P('You sign in with your own account. It needs at least these roles for the products you selected (Global Administrator always works):'),
      TABLE(['Product', 'Minimum role'], rows),
    ];
    if (c.has('aad') || c.has('sharepoint')) {
      b.push(NOTE('info', 'Entra ID uses the Microsoft Graph PowerShell app. On first run an administrator must **consent** to its read permissions (Directory.Read.All, Policy.Read.All, PrivilegedAccess.Read.AzureADGroup, PrivilegedEligibilitySchedule.Read.AzureADGroup, RoleManagement.Read.Directory, RoleManagementPolicy.Read.AzureADGroup, User.Read.All). ScubaGear prompts for this.'));
    }
    if (c.has('powerbi')) b.push(NOTE('warn', 'Power BI: the account needs a Power BI or Microsoft Fabric license and must have signed in to the Power BI portal at least once before running ScubaGear.'));
    if (c.has('powerplatform')) {
      b.push(NOTE('warn', 'Power Platform: use the correct environment value (`' + c.env + '`) or ScubaGear reports a region mismatch error.'));
    }
    b.push(LINKS([link('prerequisites/interactive.md', 'Interactive permissions'), link('troubleshooting/power.md', 'Power Platform troubleshooting')]));
    return b;
  }

  function spProducts(c) {
    return c.products.filter((p) => SP_PRODUCTS.includes(p));
  }
  const psList = (arr) => arr.map((p) => "'" + p + "'").join(', ');

  function spCreateBlocks(c, name) {
    const prods = spProducts(c);
    const b = [
      NOTE('info', 'You need the **Application Administrator** role (or Global Admin, Security Admin, or Cloud App Admin). The commands ask for Graph consent when run.'),
      P('1. Preview what will be created (no changes made):'),
      CODE(cmd('New-ScubaGearServicePrincipal', ['-M365Environment ' + c.env, '-ProductNames ' + psList(prods), '-WhatIf'])),
      P('2. Create it. Keep the output; it holds the **AppID** and **CertThumbprint** you need to run ScubaGear:'),
      CODE('$NewSP = ' + cmd('New-ScubaGearServicePrincipal', ['-M365Environment ' + c.env, '-ProductNames ' + psList(prods), '-ServicePrincipalName ' + psq(name || 'ScubaGear-Production')])),
      P('3. Verify:'),
      CODE(cmd('Get-ScubaGearAppPermission', ['-AppID $NewSP.AppID', '-M365Environment $NewSP.M365Environment', '-ProductNames $NewSP.ProductNames']), 'Expected: "No action needed - service principal is correctly configured."'),
    ];
    if (c.has('powerbi')) {
      b.push(NOTE('warn', '**Power BI** is not covered by the automated setup. Create a plain security group containing the service principal, then in the Power BI admin portal enable **Tenant settings > Admin API settings > Service principals can access read-only admin APIs** for that group. Do **not** add `Tenant.Read.All` to the app registration.'));
    }
    if (c.has('teams')) {
      b.push(NOTE('warn', '**Teams limitation:** with certificate authentication ScubaGear cannot read the org-wide app settings (MS.TEAMS.5.1v2, 5.2v2, 5.3v2). Run Teams interactively when you need those policies.'));
    }
    if (c.env === 'gcchigh' || c.env === 'dod') {
      b.push(NOTE('info', 'GCC High and DoD have extra non-interactive details; see the non-interactive permissions page.'));
    }
    b.push(LINKS([link('prerequisites/serviceprincipal-workflows.md', 'Service principal workflows'), link('prerequisites/noninteractive.md', 'Non-interactive permissions'), link('prerequisites/serviceprincipal-troubleshooting.md', 'Service principal troubleshooting')]));
    return b;
  }

  function spExistingBlocks(c) {
    const prods = spProducts(c);
    return [
      P('Audit the service principal you already have:'),
      CODE('$Audit = ' + cmd('Get-ScubaGearAppPermission', ['-AppID ' + c.appIdPh, '-M365Environment ' + c.env, '-ProductNames ' + psList(prods)])),
      CODE('$Audit\n$Audit.FixPermissionIssues', 'Look at MissingPermissions, ExtraPermissions, MissingRoles, and Status. FixPermissionIssues is the exact remediation command.'),
      P('If problems are listed, preview, then apply the fix (requires Privileged Role Administrator or similar):'),
      CODE('$Audit | Set-ScubaGearAppPermission -WhatIf\n$Audit | Set-ScubaGearAppPermission'),
      CODE(cmd('Get-ScubaGearAppCert', ['-AppID ' + c.appIdPh, '-M365Environment ' + c.env]), 'Check certificate expiry. Rotation steps are in the "service principal stopped working" quick answer.'),
      c.has('powerbi') ? NOTE('warn', 'Power BI needs the admin tenant setting and a security group instead of API permissions.') : null,
      c.has('teams') ? NOTE('warn', 'Teams org-wide app settings (MS.TEAMS.5.1v2, 5.2v2, 5.3v2) cannot be read with certificate auth; run Teams interactively for those.') : null,
      LINKS([link('prerequisites/serviceprincipal-workflows.md', 'Service principal workflows')]),
    ].filter(Boolean);
  }

  function configBlocks(c) {
    const how = c.s.configTool || 'app';
    const yaml = buildYaml(c);
    const b = [];
    if (how === 'app') {
      b.push(P('The Configuration UI is a Windows desktop app that writes the YAML for you.'));
      b.push(CODE(c.has('aad') ? 'Start-SCuBAConfigApp -Online' + (c.env !== 'commercial' ? ' -M365Environment ' + c.env : '') : 'Start-SCuBAConfigApp', c.has('aad') ? '`-Online` signs in to Microsoft Graph so you can search for users and groups to exclude.' : null));
      b.push(STEPS(['Enter your organization name and unit.', 'Choose your products.', 'Add exclusions, annotations, and omissions as needed.', 'Review the preview, save the file, and optionally run ScubaGear from the app.']));
      b.push(P('Want a head start? Save the generated file below and import it:'));
      b.push(CODE('Start-SCuBAConfigApp -ConfigFilePath ' + psq('.\\scuba-config.yaml')));
      b.push(LINKS([link('configuration/scubaconfigapp.md', 'Configuration UI guide'), link('configuration/scubaconfigapp-walkthrough.md', 'UI walkthrough')]));
    } else if (how === 'cli') {
      b.push(P('`New-SCuBAConfig` generates a template and works on any OS:'));
      const parts = [];
      if (c.tenant) parts.push('-Organization ' + psq(c.tenant));
      parts.push('-ProductNames "' + c.products.join(',') + '"');
      b.push(CODE(cmd('New-SCuBAConfig', parts), 'Creates SampleConfig.yaml. Edit it by hand.'));
      b.push(LINKS([link('configuration/configuration.md', 'Configuration file reference')]));
    } else {
      b.push(P('Save the generated file below as `scuba-config.yaml` and edit it by hand.'));
      b.push(LINKS([link('configuration/configuration.md', 'Configuration file reference')]));
    }
    b.push({ type: 'file', filename: 'scuba-config.yaml', text: yaml, caption: 'Generated starting configuration' });
    if (c.users.length || c.groups.length) {
      b.push(NOTE('warn', 'The break-glass users and groups you entered are excluded from every Conditional Access policy check listed (and the role-exclusion check MS.AAD.7.4v1). Exclusions weaken the assessment, so only keep them if your risk acceptance process approves them.'));
    } else if (c.has('aad')) {
      b.push(NOTE('info', 'No exclusions were added. If Entra ID Conditional Access checks fail only because of emergency-access accounts, the Config Analyzer later in this plan will tell you exactly which exclusions to add.'));
    }
    b.push(NOTE('info', '`OmitPolicy` and `AnnotatePolicy` are left empty on purpose. Omissions and annotations need your organization\'s approval and justification.'));
    return b;
  }

  function runBlocks(c) {
    const cfgName = '.\\scuba-config.yaml';
    const b = [];
    if (c.auth === 'sp') b.push(NOTE('info', 'Authentication is read from the config file (Organization, AppID, CertificateThumbprint). Make sure those TODO values are filled in.'));
    b.push(P('Run the assessment using your configuration file:'));
    b.push(CODE(cmd('Invoke-SCuBA', invokeParts(c, { configPath: cfgName, configOnly: true })), c.auth === 'interactive' ? 'Sign in when prompted; you may get one prompt per product.' : 'No sign-in prompt appears; the certificate authenticates.'));
    b.push(DETAILS('Prefer a quick first look without a config file?', [
      CODE(cmd('Invoke-SCuBA', invokeParts(c)), 'Passing no `-ProductNames` assesses everything except Power Platform and Power BI.'),
    ]));
    b.push(P('Results land in a timestamped folder:'));
    b.push(TABLE(['File', 'What it is'], [
      ['BaselineReports.html', 'Human-readable summary. Opens in your browser unless you used `-Quiet`.'],
      ['ScubaResults_<uuid>.json', 'Everything in one file. Feed this to the Config Analyzer and to Invoke-SCuBADiff.'],
      ['ScubaResults.csv', 'Per-control results for spreadsheets.'],
      ['ActionPlan.csv', 'Failing SHALL controls with columns for reasons and remediation dates.'],
    ]));
    b.push(LINKS([link('execution/execution.md', 'Execution'), link('configuration/parameters.md', 'All parameters'), link('execution/reports.md', 'Reports')]));
    return b;
  }

  function resultsBlocks(c) {
    return [
      P('Failures are normal on a first run. Work through them in this loop until the report matches reality:'),
      STEPS([
        'Open `BaselineReports.html` and review each failing control. Use the viewer below if you are unsure what a policy checks.',
        'Fix real gaps in your tenant.',
        'For failures caused by justified exclusions (for example break-glass accounts), run the Config Analyzer and load your results file.',
        'Keep only the exclusions you can justify, paste them into your config, and run ScubaGear again.',
        'Annotate remaining failures with a plan of action, or omit controls only with approval.',
      ]),
      CODE('Show-SCuBABaselinePolicyViewer', 'Windows UI: look up what any baseline policy requires.'),
      CODE('Start-SCuBAConfigAnalyzer', 'In the window click **Load Results JSON...** and choose `ScubaResults_<uuid>.json`. Turn on "Configurable only" to see failures a config file can resolve. Offline mode uses ScubaGear\'s own verdict; the live "Connect & Scan Tenant" mode is only predictive.'),
      NOTE('warn', 'The analyzer never changes your tenant and never writes omissions. It proposes exclusions; you decide which are justified.'),
      c.has('aad') ? P('After editing the config, re-run:') : P('Re-run after changes:'),
      CODE(cmd('Invoke-SCuBA', invokeParts(c, { configPath: '.\\scuba-config.yaml', configOnly: true }))),
      LINKS([link('configuration/scubaconfiganalyzer.md', 'Config Analyzer'), link('configuration/configuration.md#annotate-policies', 'Annotate policies'), link('configuration/configuration.md#omit-policies', 'Omit policies')]),
    ];
  }

  function iterateBlocks(c) {
    return [
      P('**Compare two runs.** After a second assessment, see what improved or regressed. It works offline on two results files:'),
      CODE(cmd('Invoke-SCuBADiff', ["-BeforePath 'C:\\Runs\\Q1\\ScubaResults_<uuid>.json'", "-AfterPath 'C:\\Runs\\Q2\\ScubaResults_<uuid>.json'", "-OutPath '.\\diff'"]), 'Creates DiffReport.html, DiffResults.json, and DiffResults.csv. Put the earlier run first.'),
      P('**Iterate without rescanning.** `Invoke-SCuBACached` re-analyzes data that was already exported, with no sign-in. It is aimed at development and testing, and its documented parameter list does not include `-ConfigFilePath`; confirm with `Get-Help Invoke-SCuBACached -Full` before relying on it for config changes.'),
      CODE("Invoke-SCuBACached -ProductNames " + c.productArg + " -ExportProvider $false -OutPath 'C:\\ScubaResults\\<run-folder>'"),
      LINKS([link('execution/diff.md', 'Invoke-SCuBADiff'), link('execution/scubacached.md', 'Invoke-SCuBACached')]),
    ];
  }

  function troubleBlocks() {
    return [
      DETAILS('Sign-in keeps looping (Entra ID)', [CODE('Disconnect-MgGraph'), P('Then run ScubaGear again.')]),
      DETAILS('"Key not valid for use in specified state."', [CODE('Remove-Item -Recurse $HOME\\.graph'), P('Then run ScubaGear again.')]),
      DETAILS('"Could not load file or assembly Microsoft.Graph.Authentication"', [CODE('Install-ScubaDependencies'), P('If it persists, run `Reset-ScubaGearDependencies` (it asks before removing modules).')]),
      DETAILS('Connected to the wrong tenant', [CODE('Disconnect-SCuBATenant'), P('Tip: add `-DisconnectOnExit` when testing multiple tenants.')]),
      DETAILS('Install looks broken', [CODE('Get-ScubaGearDependencyStatus\nReset-ScubaGearDependencies')]),
      DETAILS('Need to report a problem', [
        CODE('Invoke-SCuBA -ConfigFilePath .\\scuba-config.yaml -Transcript', 'Transcripts are not redacted; do not share them publicly.'),
        CODE("Get-ScubaDebugLogReport -DebugLogPath 'C:\\ScubaResults\\<run-folder>\\DebugLogs\\ScubaGear-DebugLog-<timestamp>.log'", 'Creates a Markdown report suitable for a GitHub issue.'),
        LINKS([link('troubleshooting/debuglogging.md', 'Debug logging'), { label: 'ScubaGear issues', url: 'https://github.com/cisagov/ScubaGear/issues' }]),
      ]),
    ];
  }

  /* ---------- guided plan ---------- */
  function buildGuided(s) {
    const c = ctx(s);
    const sections = [];
    const failed = s.ranBefore === 'failed';

    sections.push(SECTION('prepare', 'Prepare your computer', [
      ps51Note(),
      CODE('$PSVersionTable.PSVersion', 'Major should be 5 and Minor 1.'),
      s.install === 'github' ? P('Because you are downloading from GitHub, the execution policy steps in the next section matter.') : null,
    ].filter(Boolean)));

    sections.push(SECTION('install', 'Install ScubaGear', installBlocks(c)));
    sections.push(SECTION('deps', 'Install the dependencies', depsBlocks()));
    if (s.network) sections.push(SECTION('network', 'Allow network access', networkBlocks(c)));

    if (c.auth === 'sp') {
      sections.push(SECTION('perms', 'Create or verify the service principal', s.spHave === 'existing' ? spExistingBlocks(c) : spCreateBlocks(c, '')));
    } else {
      sections.push(SECTION('perms', 'Check your permissions', interactivePermBlocks(c)));
    }

    sections.push(SECTION('config', 'Create your configuration file', configBlocks(c)));
    sections.push(SECTION('run', 'Run ScubaGear', runBlocks(c)));
    sections.push(SECTION('results', 'Review results and fix failures', resultsBlocks(c), failed ? { highlight: 'You said you already have failing results. Start here.' } : null));
    sections.push(SECTION('iterate', 'Track progress over time', iterateBlocks(c)));
    sections.push(SECTION('trouble', 'If something goes wrong', troubleBlocks()));

    return {
      title: 'Your ScubaGear plan',
      intro: 'Environment: ' + ENVIRONMENTS.find((e) => e.value === c.env).label + '. Products: ' + c.productArg + '. Sign-in: ' + (c.auth === 'sp' ? 'service principal (unattended)' : 'interactive') + '.',
      sections, files: [{ filename: 'scuba-config.yaml', text: buildYaml(c) }],
      jumpTo: failed ? 'results' : null,
    };
  }

  /* ---------- shared field definitions ---------- */
  const F = {
    env: { key: 'env', type: 'cards', label: 'Which Microsoft 365 environment is the tenant in?', options: ENVIRONMENTS, def: 'commercial' },
    products: { key: 'products', type: 'multi', label: 'Which products do you need to assess?', help: 'Power Platform and Power BI have extra licensing and one-time setup, so they are off by default.', options: PRODUCTS.map((p) => ({ value: p.id, label: p.label + ' (' + p.id + ')', desc: p.special ? 'Needs extra setup' : '' })), def: DEFAULT_PRODUCTS },
    auth: { key: 'auth', type: 'cards', label: 'How will ScubaGear sign in?', def: 'interactive', options: [
      { value: 'interactive', label: 'Interactive (me, now)', desc: 'You sign in with an admin account. Best for manual runs.' },
      { value: 'sp', label: 'Unattended (service principal)', desc: 'A certificate-based app identity for scheduled jobs and pipelines.' },
    ] },
    spHave: { key: 'spHave', type: 'cards', label: 'Do you already have a service principal for ScubaGear?', def: 'new', when: (s) => s.auth === 'sp', options: [
      { value: 'new', label: 'No, create one', desc: 'New-ScubaGearServicePrincipal builds it with the right permissions.' },
      { value: 'existing', label: 'Yes, audit it', desc: 'Check and fix permissions and certificate.' },
    ] },
    appId: { key: 'appId', type: 'text', kind: 'guid', label: 'Application (client) ID', help: 'Optional. Leave blank to get a placeholder.', placeholder: '00000000-0000-0000-0000-000000000000', when: (s) => s.auth === 'sp' },
    thumb: { key: 'thumb', type: 'text', kind: 'thumb', label: 'Certificate thumbprint', help: 'Optional. 40 hex characters.', placeholder: 'FEDCBA9876543210FEDCBA9876543210FEDCBA98', when: (s) => s.auth === 'sp' },
    tenant: { key: 'tenant', type: 'text', kind: 'domain', label: 'Tenant initial domain', help: 'Of the form contoso.onmicrosoft.com. Required for unattended runs.', placeholder: 'contoso.onmicrosoft.com' },
    outPath: { key: 'outPath', type: 'text', kind: 'path', label: 'Output folder (optional)', help: 'Defaults to a timestamped folder in the current directory.', placeholder: 'C:\\ScubaResults' },
    dark: { key: 'dark', type: 'toggle', label: 'Dark mode HTML report' },
    quiet: { key: 'quiet', type: 'toggle', label: 'Do not open the report in a browser (-Quiet)' },
    transcript: { key: 'transcript', type: 'toggle', label: 'Save a PowerShell transcript for troubleshooting (-Transcript)', help: 'Transcripts are not redacted.' },
  };

  const GUIDED = [
    { id: 'tenant', title: 'Your tenant', intro: 'ScubaGear needs to know which cloud your Microsoft 365 tenant lives in.', fields: [F.env, F.tenant] },
    { id: 'products', title: 'Products', intro: 'Pick the baselines your agency needs to assess.', fields: [F.products] },
    { id: 'machine', title: 'This computer', intro: 'Tell us how ScubaGear gets onto your machine.', fields: [
      { key: 'install', type: 'cards', label: 'ScubaGear installation', def: 'psgallery', options: [
        { value: 'psgallery', label: 'Install from PowerShell Gallery', desc: 'Recommended.' },
        { value: 'github', label: 'Download from GitHub', desc: 'Use when PSGallery is not allowed.' },
        { value: 'installed', label: 'It is already installed', desc: 'Just check for updates.' },
      ] },
      { key: 'network', type: 'toggle', label: 'Outbound traffic is restricted by a firewall, proxy, or gateway', help: 'You will get a list of hosts to allow.' },
    ] },
    { id: 'signin', title: 'Sign-in method', intro: 'Choose between a person signing in and an unattended service principal.', fields: [F.auth, F.spHave, F.appId, F.thumb] },
    { id: 'org', title: 'Organization', intro: 'These are written to the config file and used in BOD submissions.', fields: [
      { key: 'orgName', type: 'text', kind: 'any', label: 'Organization name', placeholder: 'Department of Example' },
      { key: 'orgUnit', type: 'text', kind: 'any', label: 'Organizational unit name', placeholder: 'Subdepartment of Example' },
    ] },
    { id: 'config', title: 'Configuration', intro: 'Most failures come from legitimate exclusions such as break-glass accounts. You can add them now or let the Config Analyzer find them after your first run.', fields: [
      { key: 'configTool', type: 'cards', label: 'How do you want to build the config file?', def: 'app', options: [
        { value: 'app', label: 'Configuration UI', desc: 'Windows desktop app (Start-SCuBAConfigApp).' },
        { value: 'cli', label: 'Command line', desc: 'New-SCuBAConfig, works on any OS.' },
        { value: 'file', label: 'Edit YAML by hand', desc: 'Start from the generated file.' },
      ] },
      { key: 'bgUsers', type: 'text', kind: 'guidlist', multiline: true, label: 'Break-glass / emergency access user IDs (optional)', help: 'Entra object IDs, separated by commas. Excluded from Conditional Access checks.', placeholder: '12345678-1234-1234-1234-123456789012', when: (s) => (s.products || DEFAULT_PRODUCTS).includes('aad') },
      { key: 'bgGroups', type: 'text', kind: 'guidlist', multiline: true, label: 'Break-glass group IDs (optional)', placeholder: '11111111-1111-1111-1111-111111111111', when: (s) => (s.products || DEFAULT_PRODUCTS).includes('aad') },
      { key: 'ranBefore', type: 'cards', label: 'Where are you in the process?', def: 'no', options: [
        { value: 'no', label: 'This is my first run', desc: '' },
        { value: 'failed', label: 'I ran ScubaGear and have failures', desc: 'Jump straight to fixing them.' },
      ] },
    ] },
    { id: 'output', title: 'Run options', intro: 'Optional switches for the run.', fields: [F.outPath, F.dark, F.quiet, F.transcript] },
  ];

  /* ---------- quick-answer recipes ---------- */
  const WIN = 'Windows only';
  const RECIPES = [
    {
      id: 'config', title: 'I need a config', cmd: 'Start-SCuBAConfigApp / New-SCuBAConfig',
      fields: [F.products, F.env, F.tenant,
        { key: 'orgName', type: 'text', kind: 'any', label: 'Organization name', placeholder: 'Department of Example' },
        { key: 'orgUnit', type: 'text', kind: 'any', label: 'Organizational unit name', placeholder: 'Subdepartment of Example' }],
      build(s) {
        const c = ctx(s);
        return { title: 'Create a ScubaGear config', sections: [
          SECTION('ui', 'Option 1: Configuration UI (' + WIN + ')', [
            CODE(c.has('aad') ? 'Start-SCuBAConfigApp -Online' + (c.env !== 'commercial' ? ' -M365Environment ' + c.env : '') : 'Start-SCuBAConfigApp'),
            LINKS([link('configuration/scubaconfigapp.md', 'Configuration UI guide')]),
          ]),
          SECTION('cli', 'Option 2: Command line (any OS)', [
            CODE(cmd('New-SCuBAConfig', [c.tenant ? '-Organization ' + psq(c.tenant) : null, '-ProductNames "' + c.products.join(',') + '"'])),
            LINKS([link('configuration/configuration.md#generate-configuration-files', 'Generate configuration files')]),
          ]),
          SECTION('yaml', 'Or start from a generated file', [{ type: 'file', filename: 'scuba-config.yaml', text: buildYaml(c), caption: 'Generated starting configuration' }]),
        ], files: [{ filename: 'scuba-config.yaml', text: buildYaml(c) }] };
      },
    },
    {
      id: 'policy', title: 'What does this policy even check?', cmd: 'Show-SCuBABaselinePolicyViewer', fields: [],
      build() {
        return { title: 'Look up a baseline policy', sections: [SECTION('viewer', 'Open the policy viewer (' + WIN + ')', [
          CODE('Show-SCuBABaselinePolicyViewer', 'Search by control ID, such as MS.AAD.1.1v1.'),
          P('Use it before adding exclusions or omissions so you know exactly what a control requires.'),
          LINKS([link('configuration/scubaconfigapp-walkthrough.md', 'UI walkthrough'), { label: 'Baseline documents', url: 'https://github.com/cisagov/ScubaGear/tree/main/baselines' }]),
        ])] };
      },
    },
    {
      id: 'run', title: 'Run the assessment', cmd: 'Invoke-SCuBA',
      fields: [F.products, F.env, F.auth, F.appId, F.thumb, F.tenant,
        { key: 'configPath', type: 'text', kind: 'path', label: 'Config file (optional)', placeholder: '.\\scuba-config.yaml' }, F.outPath, F.dark, F.quiet],
      build(s) {
        const c = ctx(s);
        const b = [ps51Note(), CODE(cmd('Invoke-SCuBA', invokeParts(c, { configPath: c.configPath })))];
        if (c.has('powerplatform') || c.has('powerbi')) b.push(NOTE('warn', 'Power Platform and Power BI have special licensing and first-time setup. See the permissions page.'));
        b.push(LINKS([link('execution/execution.md', 'Execution'), link('configuration/parameters.md', 'Parameters')]));
        return { title: 'Run Invoke-SCuBA', sections: [SECTION('run', 'Command', b), SECTION('perm', 'Required access', c.auth === 'sp' ? spCreateBlocks(c, '') : interactivePermBlocks(c))] };
      },
    },
    {
      id: 'cached', title: "I'm iterating and don't want to re-scan", cmd: 'Invoke-SCuBACached',
      fields: [F.products,
        { key: 'cachedMode', type: 'cards', label: 'What do you want to do?', def: 'cached', options: [
          { value: 'cached', label: 'Re-analyze saved data', desc: 'No sign-in; uses the previous export.' },
          { value: 'full', label: 'Export fresh data first', desc: 'Like Invoke-SCuBA but writes directly into the folder.' },
        ] },
        { key: 'outPath', type: 'text', kind: 'path', label: 'Folder with the previous export', placeholder: 'C:\\ScubaResults\\M365BaselineConformance_2025_09_22_10_19_24' }, F.dark, F.quiet],
      build(s) {
        const c = ctx(s);
        const parts = ['-ProductNames ' + c.productArg];
        if (s.cachedMode !== 'full') parts.push('-ExportProvider $false');
        parts.push('-OutPath ' + psq(c.outPath || 'C:\\ScubaResults\\<run-folder>'));
        if (c.dark) parts.push('-DarkMode');
        if (c.quiet) parts.push('-Quiet');
        return { title: 'Re-run without re-scanning', sections: [
          SECTION('cmd', 'Command', [CODE(cmd('Invoke-SCuBACached', parts)), NOTE('warn', 'This command is designed for development and testing. Its documented parameters do not include `-ConfigFilePath`, so use `Invoke-SCuBA` when you need to apply config changes. Run `Get-Help Invoke-SCuBACached -Full` to confirm for your version.')]),
          SECTION('req', 'Requirements', [LIST(['The folder must contain the provider export JSON from an earlier run.', 'Product names must match what was exported.', 'Do not reuse exports across ScubaGear versions.']), LINKS([link('execution/scubacached.md', 'Invoke-SCuBACached')])]),
        ] };
      },
    },
    {
      id: 'analyzer', title: 'It failed - what exclusion do I need?', cmd: 'Start-SCuBAConfigAnalyzer', fields: [],
      build(s) {
        const c = ctx(s);
        return { title: 'Triage failures with the Config Analyzer', sections: [
          SECTION('steps', 'Steps (' + WIN + ')', [
            STEPS(['Run ScubaGear first and note the output folder.', 'Open the analyzer.', 'Click **Load Results JSON...** and pick `ScubaResults_<uuid>.json`.', 'Turn on **Issues only** and **Configurable only**.', 'For each finding, read the root cause and recommendation. Copy the generated YAML for exclusions you can justify.', 'Paste into your config and re-run `Invoke-SCuBA`.']),
            CODE('Start-SCuBAConfigAnalyzer'),
            NOTE('info', 'Prefer offline mode: it uses ScubaGear\'s own verdicts. **Connect & Scan Tenant** is a predictive pre-check; confirm with a real run.'),
            NOTE('warn', 'It only suggests exclusions and allow-lists. It never changes your tenant, never writes omissions, and never creates Conditional Access policies.'),
            LINKS([link('configuration/scubaconfiganalyzer.md', 'Config Analyzer')]),
          ]),
        ] };
      },
    },
    {
      id: 'diff', title: 'What changed since last time?', cmd: 'Invoke-SCuBADiff',
      fields: [
        { key: 'before', type: 'text', kind: 'path', label: 'Earlier ScubaResults JSON', placeholder: 'C:\\Runs\\Q1\\ScubaResults_abc123.json' },
        { key: 'after', type: 'text', kind: 'path', label: 'Later ScubaResults JSON', placeholder: 'C:\\Runs\\Q2\\ScubaResults_def456.json' },
        F.outPath, F.dark],
      build(s) {
        const c = ctx(s);
        const before = validate('path', s.before).value || 'C:\\Runs\\Q1\\ScubaResults_<uuid>.json';
        const after = validate('path', s.after).value || 'C:\\Runs\\Q2\\ScubaResults_<uuid>.json';
        const parts = ['-BeforePath ' + psq(before), '-AfterPath ' + psq(after)];
        if (c.outPath) parts.push('-OutPath ' + psq(c.outPath));
        if (c.dark) parts.push('-DarkMode');
        return { title: 'Compare two runs', sections: [
          SECTION('cmd', 'Command', [CODE(cmd('Invoke-SCuBADiff', parts), 'Offline: no sign-in and no tenant contact.'), NOTE('info', 'Order matters. A swapped pair produces a valid but inverted report, and ScubaGear warns if the timestamps look reversed.')]),
          SECTION('out', 'Output', [LIST(['`DiffReport.html`: color-coded, self-contained report', '`DiffResults.json`: machine-readable delta', '`DiffResults.csv`: one row per policy']), LINKS([link('execution/diff.md', 'Invoke-SCuBADiff')])]),
        ] };
      },
    },
    {
      id: 'firstsetup', title: 'First-time setup', cmd: 'Install-ScubaDependencies',
      fields: [
        { key: 'install', type: 'cards', label: 'How do you want to install ScubaGear?', def: 'psgallery', options: [
          { value: 'psgallery', label: 'PowerShell Gallery', desc: 'Recommended.' }, { value: 'github', label: 'GitHub download', desc: '' }] },
        F.env, F.products, F.tenant,
        { key: 'network', type: 'toggle', label: 'My network restricts outbound traffic' }],
      build(s) {
        const c = ctx(s);
        const sections = [SECTION('ps', 'Start in the right shell', [ps51Note()]), SECTION('install', 'Install ScubaGear', installBlocks(c)), SECTION('deps', 'Install dependencies', depsBlocks())];
        if (s.network) sections.push(SECTION('net', 'Allow network access', networkBlocks(c)));
        return { title: 'First-time setup', sections };
      },
    },
    {
      id: 'broken', title: 'Something is off with my install', cmd: 'Get-ScubaGearDependencyStatus', fields: [],
      build() {
        return { title: 'Diagnose and repair your install', sections: [
          SECTION('check', 'Check what is installed', [CODE('Get-ScubaGearDependencyStatus')]),
          SECTION('fix', 'Pick a repair', [
            P('**Gentle:** install or update missing modules without removing anything.'),
            CODE('Install-ScubaDependencies -Scope CurrentUser'),
            P('**Clean slate:** remove and reinstall all ScubaGear dependencies (prompts before removing). Avoid if other tools share those modules.'),
            CODE('Reset-ScubaGearDependencies'),
            NOTE('info', 'Close every PowerShell window first so no module is loaded. Use Windows PowerShell 5.1.'),
            LINKS([link('installation/update.md#2-dependency-updating', 'Dependency updating'), link('troubleshooting/graph.md', 'Microsoft Graph errors')]),
          ]),
        ] };
      },
    },
    {
      id: 'version', title: 'Is there a newer version?', cmd: 'Test-ScubaGearVersion', fields: [
        { key: 'updSource', type: 'cards', label: 'Where do you update from?', def: 'psgallery', options: [{ value: 'psgallery', label: 'PowerShell Gallery', desc: 'Recommended.' }, { value: 'github', label: 'GitHub release', desc: 'Use when PSGallery is blocked.' }] }],
      build(s) {
        const gh = s.updSource === 'github';
        return { title: 'Check for and apply updates', sections: [
          SECTION('check', 'Check', [CODE(gh ? 'Test-ScubaGearVersion -CheckGitHub' : 'Test-ScubaGearVersion'), TABLE(['Status', 'Action'], [['Up to Date', 'Nothing to do'], ['Update Available', 'Run Update-ScubaGear'], ['Needs attention (multiple versions)', 'Run Update-ScubaGear to clean up']])]),
          SECTION('update', 'Update', [CODE(gh ? 'Update-ScubaGear -Source GitHub' : 'Update-ScubaGear'), CODE('Test-ScubaGearVersion', 'Verify.'), NOTE('info', 'Close other PowerShell windows that may have ScubaGear loaded. After a major update also run `Install-ScubaDependencies`.'), LINKS([link('installation/update.md', 'Update instructions')])]),
        ] };
      },
    },
    {
      id: 'unattended', title: 'I want to run ScubaGear unattended', cmd: 'New-ScubaGearServicePrincipal',
      fields: [F.env, F.products, F.tenant, { key: 'spName', type: 'text', kind: 'any', label: 'Service principal name (optional)', placeholder: 'ScubaGear-Production' }, F.dark],
      build(s) {
        const c = ctx(s);
        c.auth = 'sp';
        const cfg = buildYaml(c);
        return { title: 'Set up unattended ScubaGear', sections: [
          SECTION('sp', 'Create the service principal', spCreateBlocks(c, validate('any', s.spName).value)),
          SECTION('run', 'Run it', [CODE(cmd('Invoke-SCuBA', ['-AppID $NewSP.AppID', '-CertificateThumbprint $NewSP.CertThumbprint', '-ProductNames ' + c.productArg, '-Organization ' + c.tenantPh])), NOTE('info', 'The certificate lives in `Cert:\\CurrentUser\\My`, so schedule the job as the same user on the same machine. Default certificate lifetime is 6 months, so plan rotation.')]),
          SECTION('yaml', 'Optional: use a config file', [{ type: 'file', filename: 'scuba-config.yaml', text: cfg, caption: 'Includes your Organization, AppID, and thumbprint placeholders' }]),
        ], files: [{ filename: 'scuba-config.yaml', text: cfg }] };
      },
    },
    {
      id: 'spbroken', title: 'My service principal stopped working', cmd: 'Get-ScubaGearAppPermission / Get-ScubaGearAppCert',
      fields: [F.env, F.products,
        { key: 'appId', type: 'text', kind: 'guid', label: 'Application (client) ID', placeholder: '00000000-0000-0000-0000-000000000000' }],
      build(s) {
        const c = ctx(s);
        const app = c.appIdPh;
        return { title: 'Troubleshoot a service principal', sections: [
          SECTION('perm', '1. Audit permissions', spExistingBlocks(c)),
          SECTION('cert', '2. Check the certificate and rotate if needed', [
            P('Expired or expiring certificates are the most common cause.'),
            CODE('$CertStatus = ' + cmd('Get-ScubaGearAppCert', ['-AppID ' + app, '-M365Environment ' + c.env]) + '\n$CertStatus.CertificatesSummary', 'Look for HasExpiredCerts or HasExpiringSoon.'),
            P('Add a new certificate before removing the old one:'),
            CODE('$NewCert = ' + cmd('New-ScubaGearAppCert', ['-AppID ' + app, '-M365Environment ' + c.env, '-CertValidityMonths 6'])),
            CODE(cmd('Invoke-SCuBA', ['-AppID ' + app, '-CertificateThumbprint $NewCert.Thumbprint', '-ProductNames aad', '-Organization ' + c.tenantPh]), 'Test with the new certificate.'),
            CODE(cmd('Remove-ScubaGearAppCert', ['-AppID ' + app, '-M365Environment ' + c.env, "-CertThumbprint '<old-thumbprint>'"]), 'Remove the old one once verified.'),
            LINKS([link('prerequisites/serviceprincipal-workflows.md#certificate-rotation-workflow', 'Certificate rotation'), link('prerequisites/serviceprincipal-troubleshooting.md', 'Service principal troubleshooting')]),
          ]),
        ] };
      },
    },
    {
      id: 'report', title: 'I need to report a problem', cmd: 'Get-ScubaDebugLogReport',
      fields: [{ key: 'logPath', type: 'text', kind: 'path', label: 'Path to the debug log (optional)', placeholder: 'C:\\ScubaResults\\M365BaselineConformance_...\\DebugLogs\\ScubaGear-DebugLog-....log' },
        { key: 'fromCached', type: 'toggle', label: 'The log came from Invoke-SCuBACached' }],
      build(s) {
        const log = validate('path', s.logPath).value || 'C:\\ScubaResults\\<run-folder>\\DebugLogs\\ScubaGear-DebugLog-<timestamp>.log';
        const parts = ['-DebugLogPath ' + psq(log), "-OutputPath '.\\ScubaGearDebugReport.md'"];
        if (s.fromCached) parts.push('-FromScubaCached');
        return { title: 'Generate a debug report', sections: [
          SECTION('report', 'Generate the report', [
            P('Every run writes a log to the `DebugLogs` subfolder of the output folder. Turn it into a shareable Markdown report:'),
            CODE(cmd('Get-ScubaDebugLogReport', parts), 'Add `-PassThru` to also print it on screen.'),
            NOTE('warn', 'Review the report for tenant-specific details before posting publicly. Transcripts (`-Transcript`) are not redacted and are not parsed by this command.'),
            LINKS([link('troubleshooting/debuglogging.md#generating-a-debug-report', 'Debug logging'), { label: 'Open a ScubaGear issue', url: 'https://github.com/cisagov/ScubaGear/issues' }]),
          ]),
        ] };
      },
    },
  ];

  global.SG = { ENVIRONMENTS, PRODUCTS, DEFAULT_PRODUCTS, GUIDED, RECIPES, buildGuided, ctx, F };
})(window);
