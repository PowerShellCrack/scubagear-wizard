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
      connect: s.where === 'connect',
      connectMode: s.connectMode === 'cisa' ? 'cisa' : 'selfhost',
      subId: validate('guid', s.subId).value,
      rgName: validate('rgname', s.rgName).value,
      email: validate('email', s.email).value,
      schedule: ['Hour', 'Day', 'Week', 'Month'].includes(s.schedule) ? s.schedule : 'Week',
      multiTenant: !!s.multiTenant,
      federal: !!s.federal,
      auth: s.auth === 'sp' && s.where !== 'connect' ? 'sp' : 'interactive',
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
    if (c.outPath && !c.connect) {
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
    const cfgFile = c.connect ? connectFileName(c) : 'scuba-config.yaml';
    const b = [];
    if (how === 'app') {
      b.push(P('The Configuration UI is a Windows desktop app that writes the YAML for you.'));
      b.push(CODE(c.has('aad') ? 'Start-SCuBAConfigApp -Online' + (c.env !== 'commercial' ? ' -M365Environment ' + c.env : '') : 'Start-SCuBAConfigApp', c.has('aad') ? '`-Online` signs in to Microsoft Graph so you can search for users and groups to exclude.' : null));
      b.push(STEPS(['Enter your organization name and unit.', 'Choose your products.', 'Add exclusions, annotations, and omissions as needed.', 'Review the preview, save the file, and optionally run ScubaGear from the app.']));
      b.push(P('Want a head start? Save the generated file below and import it:'));
      b.push(CODE('Start-SCuBAConfigApp -ConfigFilePath ' + psq('.\\' + cfgFile)));
      b.push(LINKS([link('configuration/scubaconfigapp.md', 'Configuration UI guide'), link('configuration/scubaconfigapp-walkthrough.md', 'UI walkthrough')]));
    } else if (how === 'cli') {
      b.push(P('`New-SCuBAConfig` generates a template and works on any OS:'));
      const parts = [];
      if (c.tenant) parts.push('-Organization ' + psq(c.tenant));
      parts.push('-ProductNames "' + c.products.join(',') + '"');
      b.push(CODE(cmd('New-SCuBAConfig', parts), 'Creates SampleConfig.yaml. Edit it by hand.'));
      b.push(LINKS([link('configuration/configuration.md', 'Configuration file reference')]));
    } else {
      b.push(P('Save the generated file below as `' + cfgFile + '` and edit it by hand.'));
      b.push(LINKS([link('configuration/configuration.md', 'Configuration file reference')]));
    }
    b.push({ type: 'file', filename: cfgFile, text: yaml, caption: 'Generated starting configuration' });
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

  /* ---------- ScubaConnect ---------- */
  const CONNECT_REPO = 'https://github.com/cisagov/ScubaConnect';
  const connectFileName = (c) => (c.tenant || 'contoso.onmicrosoft.com') + '.yaml';
  const connectLink = (path, label) => ({ label, url: CONNECT_REPO + '/blob/main/' + path });
  const SCHEDULE_WORDS = { Hour: 'every hour', Day: 'every day', Week: 'every week', Month: 'every month' };

  function connectEnvNote(c) {
    return c.env === 'dod' ? NOTE('warn', 'ScubaConnect (GearConnect) supports Commercial, GCC, and GCC High. DoD is not listed, so choose another option or confirm with CISA.') : null;
  }

  function connectToolsSection(c) {
    return SECTION('tools', 'Install ScubaGear for the config tools', [
      P('ScubaConnect does the scanning, but the tools that build your YAML config (`Start-SCuBAConfigApp`, `New-SCuBAConfig`, `Start-SCuBAConfigAnalyzer`) live in the ScubaGear module. Install it on a Windows computer.'),
      ps51Note(),
      ...installBlocks(c),
      CODE('Install-ScubaDependencies'),
    ]);
  }

  function connectConfigSection(c) {
    const blocks = configBlocks(c);
    blocks.push(NOTE(c.tenant ? 'info' : 'warn', 'ScubaConnect looks for a file named `<tenant_fqdn>.yaml`, for example `contoso.onmicrosoft.com.yaml`. ' + (c.tenant ? 'The generated file is already named `' + connectFileName(c) + '`.' : 'Rename your file to match your tenant\'s initial domain, or enter it in step 1 to get the right name.')));
    return SECTION('config', 'Create your configuration file', blocks);
  }

  function installAppSection(c, appIdPh, intro) {
    return SECTION('installapp', 'Install the ScubaConnect app in your tenant', [
      connectEnvNote(c),
      P(intro),
      NOTE('info', 'Run it as a **Global Administrator**, or as a Privileged Role Administrator who is also a Fabric Administrator.'),
      STEPS(['Download the signed `Install-GearConnect.ps1` from the [ScubaConnect releases page](https://github.com/cisagov/ScubaConnect/releases).', 'Open PowerShell and run it from the folder where you saved it.']),
      CODE(cmd('.\\Install-GearConnect.ps1', ['-AppId ' + appIdPh, '-M365Environment ' + (c.env === 'dod' ? 'gcc' : c.env)]), 'The script defaults to `gcc` if you omit `-M365Environment`. Valid values are `commercial`, `gcc`, and `gcchigh`.'),
      P('The script will:'),
      LIST(['Install the PowerShell dependencies it needs on the local machine.', 'Open a consent page so you can add the ScubaConnect application with the permissions ScubaGear needs.', 'Grant the application the Global Reader role.', 'Register the application as a Power Apps admin.', 'Create a security group for the service principal and guide you to grant that group read-only Power BI access.']),
      LINKS([connectLink('m365/README.adoc#onboard', 'Onboarding a tenant')]),
    ].filter(Boolean));
  }

  function localTestBlocks(c) {
    return [
      P('**Optional: test your config locally.** You can run ScubaGear yourself with the same YAML to preview results before ScubaConnect scans, and to use the Config Analyzer on a results file.'),
      CODE('Invoke-SCuBA -ConfigFilePath ' + psq('.\\' + connectFileName(c)), 'Interactive sign-in. Results appear in a timestamped folder.'),
      CODE('Start-SCuBAConfigAnalyzer', 'Click **Load Results JSON...** and choose a `ScubaResults_<uuid>.json` file. Keep only the exclusions you can justify, add them to your YAML, and upload the YAML again.'),
      LINKS([link('configuration/scubaconfiganalyzer.md', 'Config Analyzer')]),
    ];
  }

  function failedHighlight(s) {
    return s.ranBefore === 'failed' ? { highlight: 'You said you already have failing results. Start here.' } : null;
  }

  function cyberscopeSection(mode) {
    const collect = mode === 'selfhost'
      ? 'Download the most recent `ScubaResults.json` from your output storage container.'
      : 'Locate `ScubaResults_<uuid>.json` (and `ActionPlan.csv`) from your most recent assessment in its output folder.';
    return SECTION('cyberscope', 'Submit to CyberScope every quarter (Federal agencies)', [
      NOTE('warn', 'As a Federal agency you must also **submit your results to CyberScope at the end of every calendar quarter**. This applies whether you run the ScubaGear module on a Windows device or self-hosted ScubaConnect.'),
      TABLE(['Quarter', 'Months covered'], [['Q1', 'January - March'], ['Q2', 'April - June'], ['Q3', 'July - September'], ['Q4', 'October - December']]),
      STEPS([collect, 'Sign in to the [CISA Registration Portal](https://access.cisa.gov/) and open CyberScope.', 'Submit the results for the quarter.', 'Record any remediation plans for failing controls (annotations in your config, or `ActionPlan.csv`).']),
      NOTE('info', 'Confirm the exact submission format and due dates with CISA at scuba@cisa.dhs.gov, since they can change.'),
    ]);
  }

  function buildCisa(c) {
    return [
      SECTION('about', 'How CISA-hosted ScubaConnect works', [
        NOTE('warn', 'CISA-hosted ScubaConnect is still being defined. Confirm the current onboarding steps, your application ID, and timelines with CISA at scuba@cisa.dhs.gov before you begin.'),
        P('CISA runs ScubaConnect from its own tenant using a multi-tenant service principal. It connects to your tenant, scans it on a schedule, and uploads the results to its hosted database, CyberScope. You do not deploy any infrastructure. You do three things:'),
        STEPS(['Build your ScubaGear YAML config.', 'Register with CyberScope and upload the YAML to onboard.', 'Install the ScubaConnect app in your tenant with `Install-GearConnect.ps1`.']),
      ]),
      connectToolsSection(c),
      connectConfigSection(c),
      SECTION('register', 'Register with CyberScope and upload your config', [
        STEPS(['Go to the [CISA Registration Portal](https://access.cisa.gov/) and register for CyberScope.', 'Upload your YAML file to onboard your tenant.', 'Get the **application ID** for the ScubaConnect app. CISA provides it along with the install script.']),
        NOTE('info', 'Whenever you change your config, upload the new YAML again so the next scan uses it.'),
      ]),
      installAppSection(c, '<cisa-provided-app-id>', 'This registers CISA\'s multi-tenant ScubaConnect application in your tenant so CISA\'s scans can read your configuration.'),
      SECTION('results', 'Review results and keep improving', [
        P('Scheduled scans upload results to CyberScope. How results and reports are returned to your team, such as an emailed HTML report, is still being defined. Ask CISA how you will receive them.'),
        c.federal ? NOTE('info', 'Because CISA uploads your results to CyberScope for you, the quarterly manual submission needed for the module and self-hosted options may not apply. Confirm with CISA during onboarding.') : null,
        ...localTestBlocks(c),
      ].filter(Boolean), failedHighlight(c.s)),
      SECTION('offboard', 'Offboarding', [
        P('A tenant admin can stop ScubaConnect at any time by deleting the application: Azure portal > Enterprise applications > ScubaConnect > Properties > Delete. Ask CISA to remove your config from the service as well.'),
        LINKS([connectLink('m365/README.adoc#_frequently_asked_questions', 'ScubaConnect FAQ')]),
      ]),
    ];
  }

  function tfProvider(c) {
    const env = c.env === 'gcchigh' ? 'usgovernment' : 'public';
    return [
      'terraform {',
      '  required_version = ">= 1.1.0"',
      '}',
      '',
      'provider "azurerm" {',
      '  subscription_id = "' + (c.subId || 'TODO-your-subscription-uuid') + '"',
      '  environment     = "' + env + '"',
      '  features {}',
      '}',
      '',
      'provider "azuread" {',
      '  environment = "' + env + '"',
      '}',
      '',
    ].join('\n');
  }

  function tfMain(c) {
    const L = [
      'module "scuba_connect" {',
      '  # TODO: replace vX.Y.Z with the latest tag from https://github.com/cisagov/ScubaConnect/releases',
      '  source = "github.com/cisagov/ScubaConnect.git//m365/terraform?ref=vX.Y.Z"',
      '',
      '  contact_emails      = ["' + (c.email || 'TODO@example.com') + '"]',
      '  resource_group_name = "' + (c.rgName || 'TODO-myresourcegroup') + '"',
      '  schedule_interval   = "' + c.schedule + '"',
      '  tenants_dir_path    = "./tenants"',
    ];
    if (c.multiTenant) L.push('  app_multi_tenant    = true');
    L.push('}', '');
    return L.join('\n');
  }

  function buildSelfHost(c) {
    const cfg = connectFileName(c);
    return [
      SECTION('about', 'How self-hosted ScubaConnect works', [
        P('You deploy ScubaConnect into your own Azure subscription with Terraform. An Azure container runs ScubaGear on a schedule, reads your YAML config from blob storage, and writes `ScubaResults.json` back to blob storage. You still build and maintain the YAML.'),
        STEPS(['Install Terraform and the Azure CLI, then deploy with Terraform.', 'Install the ScubaConnect app in each tenant you want to scan.', 'Upload one YAML config per tenant.', 'Review the results and refine the config.']),
        c.env === 'gcchigh' ? NOTE('info', 'GCC High is separate from Commercial and GCC. Deploy ScubaConnect in the same environment as the tenants you scan.') : null,
        connectEnvNote(c),
      ].filter(Boolean)),
      SECTION('prereq', 'Install the prerequisites', [
        LIST(['[Terraform](https://developer.hashicorp.com/terraform/install)', '[Azure CLI](https://learn.microsoft.com/en-us/cli/azure/get-started-with-azure-cli), then sign in', 'An account with administrator privileges for the Azure subscription']),
        CODE('az login\naz account show', 'Copy the subscription ID from `az account show`.'),
      ]),
      connectToolsSection(c),
      connectConfigSection(c),
      SECTION('deploy', 'Deploy with Terraform', [
        STEPS(['Copy the `m365/example_deployment` folder from the [ScubaConnect repo](https://github.com/cisagov/ScubaConnect/tree/main/m365/example_deployment) to a working folder. This is your deployment directory.', 'Replace its `provider.tf` and `main.tf` with the files below, filling in every `TODO`.', 'Run the commands.']),
        { type: 'file', filename: 'provider.tf', text: tfProvider(c), caption: 'Set `subscription_id` from `az account show`. GCC High uses `usgovernment` in both provider blocks.' },
        { type: 'file', filename: 'main.tf', text: tfMain(c), caption: c.multiTenant ? '`app_multi_tenant = true` lets other tenants install the app.' : 'The app is limited to this tenant. Set `app_multi_tenant = true` to scan other tenants too.' },
        CODE('terraform init\nterraform apply', 'Review the plan and type `yes`. The `ref` in `main.tf` should be a release tag.'),
        CODE('terraform output app_id\nterraform output input_storage_container_url', 'You need `app_id` to install the app in your tenants.'),
        NOTE('info', 'For production, also configure the remote state backend (`backend "azurerm"`) in `provider.tf`.'),
        LINKS([connectLink('m365/README.adoc#deploy', 'Deploying with Terraform'), connectLink('m365/README.adoc#terraform-variables', 'All Terraform variables')]),
      ]),
      installAppSection(c, '<app_id-from-terraform>', 'Run this even if you scan only your own tenant. For more tenants, each tenant\'s admin repeats it.'),
      SECTION('upload', 'Upload your config', [
        P('Name the file `<tenant_fqdn>.yaml` (yours: `' + cfg + '`). Then do either of these:'),
        LIST(['Put it in the `tenants/` folder of your deployment directory and run `terraform apply` again.', 'Upload it directly to the container shown by `terraform output input_storage_container_url`.']),
        NOTE('info', 'Use one file per tenant. Uploading a changed file is how you change what ScubaConnect checks.'),
      ]),
      SECTION('run', 'Schedule, run, and monitor', [
        P('Scans run ' + SCHEDULE_WORDS[c.schedule] + ', starting from when you deployed. To check or change the schedule, open Azure Portal > `scubaconnect-runner-automation-x` > Shared Resources > Schedules.'),
        P('For an on-demand scan, start the `scubaconnect-adhoc-container` container instance in the Azure portal and watch **Settings > Containers > Logs**.'),
        P('Logs go to the `scubaconnect-monitor-loganalytics` workspace, which has saved searches "scubaconnect Container Logs (7d)" and "scubaconnect Last Run Output". An email goes to your contact address if a run fails.'),
        LINKS([connectLink('m365/README.adoc#_logging_alerts', 'Logging and alerts')]),
      ]),
      SECTION('results', 'Review results and fix failures', [
        P('Results are written to the output storage container (`terraform output output_storage_container_urls`). Download a `ScubaResults.json` and work the same loop as a local run:'),
        STEPS(['Open the results in the Config Analyzer (**Load Results JSON...**).', 'Keep only the exclusions you can justify and add them to your YAML.', 'Upload the updated YAML, then wait for the next scan or start the ad-hoc container.']),
        CODE('Start-SCuBAConfigAnalyzer'),
        CODE(cmd('Invoke-SCuBADiff', ["-BeforePath '.\\before\\ScubaResults.json'", "-AfterPath '.\\after\\ScubaResults.json'"]), 'Compare two scans to see what changed.'),
        DETAILS('Optional: Power BI report', [
          P('ScubaConnect includes a Power BI report for history across scans and tenants. Automated refresh works in Commercial and GCC High; GCC uses manual updates with Power BI Desktop.'),
          LINKS([connectLink('m365/powerbi/README.md', 'Power BI integration steps')]),
        ]),
        LINKS([link('configuration/scubaconfiganalyzer.md', 'Config Analyzer'), link('execution/diff.md', 'Invoke-SCuBADiff')]),
      ], failedHighlight(c.s)),
      ...(c.federal ? [cyberscopeSection('selfhost')] : []),
      SECTION('maintain', 'Maintenance', [
        LIST(['The certificate ScubaConnect uses expires after one year. An email arrives 60 days before; running `terraform apply` renews it.', 'The container image updates itself, so no action is needed for new ScubaGear versions.', 'For new Terraform releases, change the `ref` in `main.tf`, read the release notes, then run:']),
        CODE('terraform init -upgrade\nterraform apply'),
        P('To offboard a tenant, delete the app under Enterprise applications > ScubaConnect, then delete its YAML from the input container.'),
        LINKS([connectLink('m365/README.adoc#_maintenance', 'Maintenance'), connectLink('m365/README.adoc#_frequently_asked_questions', 'FAQ')]),
      ]),
    ];
  }

  function buildConnect(s) {
    const c = ctx(s);
    const cisa = c.connectMode === 'cisa';
    return {
      title: 'Your ScubaConnect plan',
      intro: 'Mode: ' + (cisa ? 'CISA-hosted (CyberScope)' : 'self-hosted in your Azure subscription') + '. Environment: ' + ENVIRONMENTS.find((e) => e.value === c.env).label + '. Products: ' + c.productArg + '.',
      sections: cisa ? buildCisa(c) : buildSelfHost(c),
      files: [{ filename: connectFileName(c), text: buildYaml(c) }],
      jumpTo: s.ranBefore === 'failed' ? 'results' : null,
    };
  }

  /* ---------- guided plan ---------- */
  function buildGuided(s) {
    if (s.where === 'connect') return buildConnect(s);
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
    if (c.federal) sections.push(cyberscopeSection('local'));
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
    tenant: { key: 'tenant', type: 'text', kind: 'domain', label: 'Tenant initial domain', help: 'Of the form contoso.onmicrosoft.com. Required for unattended runs, and ScubaConnect uses it to name your config file.', placeholder: 'contoso.onmicrosoft.com' },
    outPath: { key: 'outPath', type: 'text', kind: 'path', label: 'Output folder (optional)', help: 'Defaults to a timestamped folder in the current directory.', placeholder: 'C:\\ScubaResults' },
    dark: { key: 'dark', type: 'toggle', label: 'Dark mode HTML report' },
    quiet: { key: 'quiet', type: 'toggle', label: 'Do not open the report in a browser (-Quiet)' },
    transcript: { key: 'transcript', type: 'toggle', label: 'Save a PowerShell transcript for troubleshooting (-Transcript)', help: 'Transcripts are not redacted.' },
  };

  const GUIDED = [
    { id: 'tenant', title: 'Your tenant', intro: 'ScubaGear needs to know which cloud your Microsoft 365 tenant lives in.', fields: [F.env, F.tenant] },
    { id: 'products', title: 'Products', intro: 'Pick the baselines your agency needs to assess.', fields: [F.products] },
    { id: 'machine', title: 'How will you run ScubaGear?', intro: 'Run it yourself on a Windows computer, or let ScubaConnect automate the scans. Either way you need a YAML config, and the plan walks you through it.', fields: [
      { key: 'where', type: 'cards', label: 'Where will ScubaGear run?', def: 'local', options: [
        { value: 'local', label: 'On my own Windows computer', desc: 'You install ScubaGear and run Invoke-SCuBA when you want results.' },
        { value: 'connect', label: 'Automated with ScubaConnect', desc: 'A container runs ScubaGear for you on a schedule. You still build and maintain the YAML config.' },
      ] },
      { key: 'install', type: 'cards', label: 'How do you get ScubaGear?', def: 'psgallery', when: (s) => s.where !== 'connect', options: [
        { value: 'psgallery', label: 'Install from PowerShell Gallery', desc: 'Recommended.' },
        { value: 'github', label: 'Download from GitHub', desc: 'Use when PSGallery is not allowed.' },
        { value: 'installed', label: 'It is already installed', desc: 'Just check for updates.' },
      ] },
      { key: 'network', type: 'toggle', label: 'Outbound traffic is restricted by a firewall, proxy, or gateway', help: 'You will get a list of hosts to allow.', when: (s) => s.where !== 'connect' },
      { key: 'connectMode', type: 'cards', label: 'Who hosts ScubaConnect?', def: 'selfhost', when: (s) => s.where === 'connect', options: [
        { value: 'cisa', label: 'CISA-hosted (CyberScope)', desc: 'CISA scans your tenant with its multi-tenant app. You submit your YAML to Cyberscope and install the app. This process is still being defined.' },
        { value: 'selfhost', label: 'Self-hosted in my Azure subscription', desc: 'You deploy ScubaConnect with Terraform. Scans run on your schedule and results stay in your Azure storage.' },
      ] },
    ] },
    { id: 'connect', title: 'ScubaConnect settings', intro: 'These fill in the Terraform files for your Azure deployment. Everything is optional; blanks become TODO placeholders.', when: (s) => s.where === 'connect' && s.connectMode !== 'cisa', fields: [
      { key: 'subId', type: 'text', kind: 'guid', label: 'Azure subscription ID', help: 'Run az account show to find it.', placeholder: '00000000-0000-0000-0000-000000000000' },
      { key: 'rgName', type: 'text', kind: 'rgname', label: 'Resource group name (will be created)', placeholder: 'rg-scubaconnect' },
      { key: 'email', type: 'text', kind: 'email', label: 'Contact email for failures and certificate expiry', placeholder: 'secops@agency.gov' },
      { key: 'schedule', type: 'cards', label: 'How often should ScubaConnect scan?', def: 'Week', options: [
        { value: 'Hour', label: 'Hourly', desc: '' }, { value: 'Day', label: 'Daily', desc: '' }, { value: 'Week', label: 'Weekly', desc: 'Default.' }, { value: 'Month', label: 'Monthly', desc: '' },
      ] },
      { key: 'multiTenant', type: 'toggle', label: 'Also scan other tenants (multi-tenant app)', help: 'Leave off to scan only the tenant you deploy in.' },
    ] },
    { id: 'signin', title: 'Sign-in method', intro: 'Choose between a person signing in and an unattended service principal.', when: (s) => s.where !== 'connect', fields: [F.auth, F.spHave, F.appId, F.thumb] },
    { id: 'org', title: 'Organization', intro: 'These are written to the config file and used in BOD submissions.', fields: [
      { key: 'orgName', type: 'text', kind: 'any', label: 'Organization name', placeholder: 'Department of Example' },
      { key: 'orgUnit', type: 'text', kind: 'any', label: 'Organizational unit name', placeholder: 'Subdepartment of Example' },
      { key: 'federal', type: 'toggle', label: 'We are a Federal Civilian Executive Branch agency (BOD 25-01)', help: 'Adds a step for the quarterly CyberScope submission.' },
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
    { id: 'output', title: 'Run options', intro: 'Optional switches for the run.', when: (s) => s.where !== 'connect', fields: [F.outPath, F.dark, F.quiet, F.transcript] },
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
      id: 'connect', title: 'I want ScubaConnect to run scans for me', cmd: 'Install-GearConnect.ps1',
      fields: [
        { key: 'connectMode', type: 'cards', label: 'Who hosts ScubaConnect?', def: 'selfhost', options: [
          { value: 'cisa', label: 'CISA-hosted (CyberScope)', desc: 'Still being defined.' },
          { value: 'selfhost', label: 'Self-hosted in my Azure subscription', desc: 'Terraform deployment.' }] },
        F.env, F.products, F.tenant,
        { key: 'subId', type: 'text', kind: 'guid', label: 'Azure subscription ID', placeholder: '00000000-0000-0000-0000-000000000000', when: (s) => s.connectMode !== 'cisa' },
        { key: 'rgName', type: 'text', kind: 'rgname', label: 'Resource group name', placeholder: 'rg-scubaconnect', when: (s) => s.connectMode !== 'cisa' },
        { key: 'email', type: 'text', kind: 'email', label: 'Contact email', placeholder: 'secops@agency.gov', when: (s) => s.connectMode !== 'cisa' },
        { key: 'multiTenant', type: 'toggle', label: 'Also scan other tenants (multi-tenant app)', when: (s) => s.connectMode !== 'cisa' }],
      build(s) {
        const plan = buildConnect(Object.assign({}, s, { where: 'connect' }));
        return { title: plan.title, sections: plan.sections };
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
