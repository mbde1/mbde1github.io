<!DOCTYPE html>
<html lang="de">
<head>
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="default">
<meta name="apple-mobile-web-app-title" content="Nachtzuschlag">
<meta name="theme-color" content="#ffd234">
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Nachtzuschlag-Rechner Â· Ennoo Hannover</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@latest/dist/tabler-icons.min.css">
<script src="https://cdn.sheetjs.com/xlsx-0.20.2/package/dist/xlsx.full.min.js"></script>
<style>
.field input,.field select,.time-24h-input{ ... font-size:16px; ... }
:root{
  --bg:#f0f2f5;
  --surface:#ffffff;
  --surface2:#f8fafc;
  --border:#e2e8f0;
  --border-light:#cbd5e1;
  --text:#1e293b;
  --text-muted:#64748b;
  --accent:#d97706;
  --gold1:#ffd234;
  --gold2:#ffe680;
  --dark:#1a3356;
  --green:#16a34a;
  --blue:#2563eb;
  --red:#dc2626;
  --purple:#7c3aed;
  --night:#0284c7;
  --dst:#d97706;
}
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:'Inter',sans-serif;background:var(--bg);color:var(--text);min-height:100vh;font-size:15px;}
.app{max-width:1040px;margin:0 auto;padding:20px 16px 60px;}

/* â”€â”€ TOPBAR â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ */
.topbar{background:linear-gradient(100deg,#ffd234 0%,#ffe680 100%);border-radius:14px;padding:16px 20px;margin-bottom:18px;display:flex;align-items:center;justify-content:space-between;gap:12px;box-shadow:0 2px 8px rgba(255,210,52,0.3);}
.tb-left{display:flex;align-items:center;gap:12px;}
.tb-logo{width:44px;height:44px;border-radius:10px;overflow:hidden;flex-shrink:0;background:#2a2a4a;display:flex;align-items:center;justify-content:center;}
.tb-logo img{width:100%;height:100%;object-fit:cover;display:block;}
.tb-title{font-size:19px;font-weight:700;color:#1a1a1a;letter-spacing:-.3px;}
.tb-sub{font-size:12px;color:rgba(0,0,0,0.55);margin-top:2px;}
.tb-badge{font-size:11px;font-weight:700;letter-spacing:.5px;padding:5px 12px;border-radius:20px;background:rgba(26,51,86,0.9);color:#ffe680;white-space:nowrap;}

/* â”€â”€ SECTION CARDS â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ */
.sec{background:var(--surface);border:2px solid var(--border);border-radius:14px;margin-bottom:14px;overflow:hidden;}
.sec-head{display:flex;align-items:center;gap:10px;padding:13px 16px;background:var(--surface2);border-bottom:2px solid var(--border);}
.sec-icon{width:32px;height:32px;border-radius:8px;display:flex;align-items:center;justify-content:center;flex-shrink:0;}
.sec-icon svg{width:18px;height:18px;fill:none;stroke-width:2;stroke-linecap:round;stroke-linejoin:round;}
.ic-blue{background:#dbeafe;}   .ic-blue svg{stroke:#1d4ed8;}
.ic-purple{background:#ede9fe;} .ic-purple svg{stroke:#7c3aed;}
.ic-orange{background:#ffedd5;} .ic-orange svg{stroke:#ea580c;}
.ic-green{background:#dcfce7;}  .ic-green svg{stroke:#16a34a;}
.ic-teal{background:#ccfbf1;}   .ic-teal svg{stroke:#0d9488;}
.ic-red{background:#fee2e2;}    .ic-red svg{stroke:#dc2626;}
.sec-title{font-size:14px;font-weight:600;color:var(--text);flex:1;}
.sec-body{padding:16px;}

/* â”€â”€ SETTINGS â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ */
.settings-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;}
@media(max-width:720px){.settings-grid{grid-template-columns:1fr 1fr;}}
.field{display:flex;flex-direction:column;gap:6px;}
.field label{font-size:11px;font-weight:600;color:var(--text-muted);text-transform:uppercase;letter-spacing:.04em;}
.field input,.field select,.time-24h-input{width:100%;padding:9px 11px;border:2px solid var(--border-light);border-radius:8px;font-size:14px;color:var(--text);background:var(--surface2);font-family:'Inter',sans-serif;outline:none;transition:border-color .15s,box-shadow .15s;}
.field input:focus,.field select:focus,.time-24h-input:focus{border-color:var(--gold1);background:#fff;box-shadow:0 0 0 3px rgba(255,210,52,.25);}

/* â”€â”€ CALLOUTS (Regeln / DST) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ */
.info-box{background:#eff6ff;border:2px solid #bfdbfe;border-radius:12px;padding:14px 16px;margin-bottom:14px;font-size:12.5px;color:#1e40af;line-height:1.7;}
.info-box strong{color:#1e3a8a;}
.dst-banner{display:none;background:#fffbeb;border:2px solid #fcd34d;border-radius:12px;padding:14px 16px;margin-bottom:14px;font-size:12.5px;color:#92400e;line-height:1.7;}
.dst-banner.visible{display:block;}
.dst-banner strong{color:#78350f;}

/* â”€â”€ IMPORT â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ */
.import-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px;}
@media(max-width:640px){.import-grid{grid-template-columns:1fr;}}
.import-block h3{font-size:12px;font-weight:600;color:var(--text-muted);margin-bottom:10px;display:flex;align-items:center;gap:6px;}
.file-input-wrapper{position:relative;display:block;overflow:hidden;}
.file-input-wrapper input[type="file"]{position:absolute;left:-9999px;}
.file-input-label{display:flex;align-items:center;gap:8px;justify-content:center;padding:14px 16px;background:var(--surface2);border:2px dashed var(--border-light);border-radius:9px;cursor:pointer;transition:all .2s;font-size:13px;color:var(--text-muted);font-weight:500;}
.file-input-label:hover{border-color:var(--gold1);color:var(--accent);background:#fffbeb;}
textarea{width:100%;background:var(--surface2);border:2px solid var(--border-light);border-radius:8px;padding:10px 12px;color:var(--text);font-family:'Inter',sans-serif;font-size:13px;resize:vertical;outline:none;line-height:1.5;transition:border-color .15s,box-shadow .15s;}
textarea:focus{border-color:var(--gold1);background:#fff;box-shadow:0 0 0 3px rgba(255,210,52,.25);}
textarea::placeholder{color:#94a3b8;}
.btn-import{margin-top:10px;background:linear-gradient(100deg,#ffd234 0%,#ffe680 100%);padding:10px 14px;border-radius:9px;border:none;width:100%;font-family:'Inter',sans-serif;font-size:13px;font-weight:600;color:#1a1a1a;cursor:pointer;transition:opacity .15s;display:flex;align-items:center;justify-content:center;gap:7px;}
.btn-import:hover{opacity:.88;}

/* â”€â”€ MANUAL ADD â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ */
.form-row{display:grid;grid-template-columns:1.5fr 1fr 1fr auto;gap:10px;align-items:end;}
@media(max-width:640px){.form-row{grid-template-columns:1fr 1fr;}.form-row .add-btn{grid-column:span 2;}}
.add-btn{background:linear-gradient(100deg,#ffd234 0%,#ffe680 100%);border:none;border-radius:9px;padding:10px 20px;color:#1a1a1a;font-family:'Inter',sans-serif;font-size:14px;font-weight:600;cursor:pointer;transition:opacity .15s,transform .1s;white-space:nowrap;}
.add-btn:hover{opacity:.88;transform:translateY(-1px);}

/* â”€â”€ ENTRIES TABLE â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ */
.sec-head .clear-all-btn{background:#fff;border:2px solid var(--border);border-radius:8px;padding:6px 12px;color:var(--red);font-family:'Inter',sans-serif;font-size:12px;font-weight:600;cursor:pointer;transition:all .15s;display:flex;align-items:center;gap:5px;}
.sec-head .clear-all-btn:hover{background:#fef2f2;border-color:var(--red);}
.entries-container{overflow-x:auto;}
.entries-table{width:100%;border-collapse:collapse;min-width:860px;}
.entries-table thead th{background:var(--surface2);padding:11px 14px;text-align:left;font-size:11px;font-weight:600;color:var(--text-muted);text-transform:uppercase;letter-spacing:.04em;border-bottom:2px solid var(--border);}
.entries-table tbody tr{border-bottom:1.5px solid #f1f5f9;transition:background .1s;}
.entries-table tbody tr:hover{background:var(--surface2);}
.entries-table td{padding:11px 14px;font-size:13.5px;}
.tag{display:inline-flex;align-items:center;gap:4px;font-size:10px;font-weight:600;padding:3px 9px;border-radius:20px;}
.tag-night{background:#e0f2fe;color:#0369a1;border:1px solid #bae6fd;}
.tag-wknight{background:#f3e8ff;color:#7c3aed;border:1px solid #e9d5ff;}
.tag-dst{background:#fef3c7;color:#b45309;border:1px solid #fde68a;font-size:9px;}
.del-btn{background:#fff;border:2px solid var(--border);border-radius:7px;color:var(--text-muted);cursor:pointer;padding:5px 9px;font-size:13px;transition:all .15s;}
.del-btn:hover{border-color:var(--red);color:var(--red);background:#fef2f2;}
.empty-state{text-align:center;padding:36px 20px;color:#94a3b8;font-size:13.5px;}

/* â”€â”€ RESULTS â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ */
.results{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;}
@media(max-width:640px){.results{grid-template-columns:1fr;}}
.result-card{background:var(--surface2);border:2px solid var(--border);border-radius:12px;padding:16px 18px;position:relative;overflow:hidden;}
.result-card::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;}
.result-card.night::before{background:var(--night);}
.result-card.combined::before{background:var(--purple);}
.result-card.total::before{background:var(--accent);}
.result-label{font-size:11px;font-weight:600;color:var(--text-muted);text-transform:uppercase;letter-spacing:.04em;margin-bottom:10px;}
.result-value{font-size:30px;font-weight:700;line-height:1;margin-bottom:6px;color:var(--text);}
.result-value.night{color:var(--night);}
.result-value.combined{color:var(--purple);}
.result-value.total{color:var(--accent);}
.result-sub{font-size:12px;color:var(--text-muted);}

/* â”€â”€ BREAKDOWN â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ */
.breakdown-table{width:100%;border-collapse:collapse;}
.breakdown-table th{padding:11px 14px;text-align:left;font-size:11px;font-weight:600;color:var(--text-muted);text-transform:uppercase;letter-spacing:.04em;border-bottom:2px solid var(--border);}
.breakdown-table td{padding:12px 14px;font-size:13.5px;border-bottom:1.5px solid #f1f5f9;color:var(--text-muted);}
.breakdown-table tr:last-child td{border-bottom:none;}
.num{text-align:right;}
.highlight{color:var(--text)!important;font-weight:600;}

/* â”€â”€ ACTIONS â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ */
.actions{display:flex;gap:10px;flex-wrap:wrap;}
.btn{display:inline-flex;align-items:center;gap:8px;padding:11px 18px;border-radius:9px;font-size:14px;font-weight:600;cursor:pointer;transition:all .15s;font-family:'Inter',sans-serif;border:none;}
.btn svg{width:16px;height:16px;fill:none;stroke-width:2;stroke-linecap:round;stroke-linejoin:round;}
.btn-primary{background:linear-gradient(100deg,#ffd234 0%,#ffe680 100%);color:#1a1a1a;}
.btn-primary svg{stroke:#1a1a1a;}
.btn-primary:hover{opacity:.88;transform:translateY(-1px);}
.btn-outline{background:#fff;border:2px solid var(--border);color:var(--text-muted);}
.btn-outline svg{stroke:var(--text-muted);}
.btn-outline:hover{border-color:var(--border-light);color:var(--text);background:var(--surface2);}

/* â”€â”€ TOAST â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€ */
.toast{position:fixed;bottom:24px;right:24px;background:#fff;border:2px solid var(--green);color:var(--green);padding:11px 18px;border-radius:9px;font-size:13px;font-weight:600;z-index:999;transform:translateY(80px);opacity:0;transition:all .25s ease;box-shadow:0 4px 14px rgba(0,0,0,.12);}
.toast.show{transform:translateY(0);opacity:1;}
</style>
</head>
<body>
<div class="app">

  <!-- TOPBAR -->
  <div class="topbar">
    <div class="tb-left">
      <div class="tb-logo">
        <img id="logo-img" src="data:image/webp;base64,UklGRn4BAABXRUJQVlA4IHIBAAAwDQCdASpgAGAAPpFCm0ilpCKhLVZpILASCWIAzQorfx2qhRn+DplmcA8+Dz5nA4CSY8jGiR9ZnRzUJnJ/AqHinag0KgHqKrBDNKWDOBlZH/qCJ+LmgWmlbLbxWLtF1CJ/wm2hP6MJe1xo7vVxwEQVuAAA/vPCZwHsD1+DxoQOSnKGCzHadoIpRS3s1q6meEixBNZghQRrsDhnDRL+o2pLWxwnI4x1t2QqPxnOkcREGwJBmiVFRA6UY+ghYr7FkSOfofyQvIdepJuuekYK+/zQGsWyN4+vS13wGbfvIo7aaGm+/iK635fnQy+hyQwCdgvikUcQofifnxUcGqZKYOL4SPRvOhebPKV5cYZVEXWLjdY9tP8U6xAzvII1i9QdwOaBuz4LYYRNEta3YFpBfyMmZd7Ln1lqGk0G9B/0/IRg+kGTlDHlmN9jBmTYkQEADQ1uuhJUIG9c93uRrdM9mmRwQdfrUMm9uhmICyAyrGJWgAAA" alt="Logo">
      </div>
      <div>
        <div class="tb-title">Nachtzuschlag-Rechner</div>
        <div class="tb-sub">Ennoo Hannover GmbH &nbsp;Â·&nbsp; Zeitumstellung automatisch (Europa/Berlin)</div>
      </div>
    </div>

  </div>

  <!-- 1 Â· EINSTELLUNGEN -->
  <div class="sec">
    <div class="sec-head">
      <div class="sec-icon ic-teal"><svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg></div>
      <span class="sec-title">1 Â· Einstellungen â€” Nachtfenster (23:00â€“06:00) &amp; ZuschlÃ¤ge</span>
    </div>
    <div class="sec-body">
      <div class="settings-grid">
        <div class="field"><label>Grundlohn (â‚¬/Std)</label><input type="number" id="baseWage" value="14.00" min="0" step="0.01"/></div>
        <div class="field"><label>Nachtzuschlag (Moâ€“Do, Soâ€“Mi)</label><input type="number" id="nightBonus" value="1.00" min="0" step="0.01"/></div>
        <div class="field"><label>WE-Nachtzuschlag (Fr &amp; Sa Nacht)</label><input type="number" id="weekendNightBonus" value="2.00" min="0" step="0.01"/></div>
        <div class="field"><label>Mitarbeiter Name</label><input type="text" id="employeeName" value="Hassan Azzam"/></div>
      </div>
    </div>
  </div>

  <!-- INFO / REGELN -->
  <div class="info-box">
    <strong>Wichtige Regeln:</strong><br>
    â€¢ <strong>Wochenend-Nacht:</strong> Fr 23:00â€“Sa 06:00 und Sa 23:00â€“So 06:00 (hÃ¶herer Satz).<br>
    â€¢ <strong>Zeitumstellung:</strong> Schichten werden in echter UTC-Zeit berechnet. Bei Sommerzeit (MÃ¤rz) hat eine Nacht 23h â€“ die fehlende Stunde wird nicht bezahlt. Bei Winterzeit (Oktober) hat eine Nacht 25h â€“ die extra Stunde wird bezahlt.<br>
    â€¢ Berechnung minutengenau, DST-bewusst (Europa/Berlin).
  </div>

  <div id="dstBanner" class="dst-banner"></div>

  <!-- 2 Â· IMPORT -->
  <div class="sec">
    <div class="sec-head">
      <div class="sec-icon ic-blue"><svg viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg></div>
      <span class="sec-title">2 Â· Daten importieren</span>
    </div>
    <div class="sec-body">
      <div class="import-grid">
        <div class="import-block">
          <h3><i class="ti ti-file-spreadsheet"></i> Excel/XLS Datei hochladen</h3>
          <div class="file-input-wrapper">
            <input type="file" id="fileInput" accept=".xlsx,.xls"/>
            <label for="fileInput" class="file-input-label"><i class="ti ti-upload"></i> Datei wÃ¤hlen oder hierher ziehen</label>
          </div>
        </div>
        <div class="import-block">
          <h3><i class="ti ti-clipboard"></i> Daten einfÃ¼gen (alternativ)</h3>
          <textarea id="pasteArea" rows="3" placeholder="Kopiere aus Excel (Spalten: Datum | Start | Ende)&#10;z.B.: 01.05.2026  22:00  06:00"></textarea>
          <button class="btn-import" onclick="importFromPaste()"><i class="ti ti-table-import"></i> Importieren</button>
        </div>
      </div>
    </div>
  </div>

  <!-- 3 Â· MANUELL HINZUFÃœGEN -->
  <div class="sec">
    <div class="sec-head">
      <div class="sec-icon ic-green"><svg viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg></div>
      <span class="sec-title">3 Â· Schicht manuell hinzufÃ¼gen</span>
    </div>
    <div class="sec-body">
      <div class="form-row">
        <div class="field"><label>Datum (Schichtstart)</label><input type="date" id="inputDate"/></div>
        <div class="field"><label>Von (Uhrzeit)</label><input type="text" id="inputStart" class="time-24h-input" value="22:00" placeholder="HH:MM"/></div>
        <div class="field"><label>Bis (Uhrzeit)</label><input type="text" id="inputEnd" class="time-24h-input" value="06:00" placeholder="HH:MM"/></div>
        <button class="add-btn" onclick="addEntry()">+ HinzufÃ¼gen</button>
      </div>
    </div>
  </div>

  <!-- 4 Â· EINGETRAGENE SCHICHTEN -->
  <div class="sec">
    <div class="sec-head">
      <div class="sec-icon ic-purple"><svg viewBox="0 0 24 24"><line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/></svg></div>
      <span class="sec-title">4 Â· Eingetragene Schichten</span>
      <button class="clear-all-btn" onclick="clearAllEntries()"><i class="ti ti-trash"></i> Alle lÃ¶schen</button>
    </div>
    <div class="sec-body" style="padding:0">
      <div class="entries-container">
        <table class="entries-table">
          <thead><tr><th>Datum</th><th>Von â€“ Bis</th><th>Uhr-Std.</th><th>Echte Std. (UTC)</th><th>Normal Nacht</th><th>WE-Nacht</th><th>Typ</th><th></th></tr></thead>
          <tbody id="entriesBody"><tr><td colspan="8"><div class="empty-state">Keine Schichten eingetragen.</div></td></tr></tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- 5 Â· ERGEBNIS -->
  <div class="sec">
    <div class="sec-head">
      <div class="sec-icon ic-orange"><svg viewBox="0 0 24 24"><rect x="4" y="2" width="16" height="20" rx="2"/><line x1="8" y1="6" x2="16" y2="6"/><line x1="8" y1="11" x2="8.01" y2="11"/><line x1="12" y1="11" x2="12.01" y2="11"/><line x1="16" y1="11" x2="16.01" y2="11"/><line x1="8" y1="15" x2="8.01" y2="15"/><line x1="12" y1="15" x2="12.01" y2="15"/><line x1="16" y1="15" x2="16.01" y2="15"/></svg></div>
      <span class="sec-title">5 Â· Ergebnis</span>
    </div>
    <div class="sec-body">
      <div class="results">
        <div class="result-card night"><div class="result-label">Normale Nachtstunden</div><div class="result-value night" id="resNight">0h 00m</div><div class="result-sub" id="resNightBonus">0,00 â‚¬ Zuschlag</div></div>
        <div class="result-card combined"><div class="result-label">WE-Nachtstunden (Fr &amp; Sa Nacht)</div><div class="result-value combined" id="resWkNight">0h 00m</div><div class="result-sub" id="resWkNightBonus">0,00 â‚¬ Zuschlag</div></div>
        <div class="result-card total"><div class="result-label">Gesamtzuschlag</div><div class="result-value total" id="resTotalBonus">0,00 â‚¬</div><div class="result-sub">zzgl. Grundlohn</div></div>
      </div>
    </div>
  </div>

  <!-- BREAKDOWN -->
  <div class="sec" id="breakdownSection" style="display:none">
    <div class="sec-head">
      <div class="sec-icon ic-red"><svg viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg></div>
      <span class="sec-title">DetailaufschlÃ¼sselung (Lohn + ZuschlÃ¤ge)</span>
    </div>
    <div class="sec-body" style="padding:0">
      <table class="breakdown-table"><thead><tr><th>Kategorie</th><th class="num">Stunden</th><th class="num">Faktor (â‚¬)</th><th class="num">Gesamt (â‚¬)</th></tr></thead><tbody id="breakdownBody"></tbody></table>
    </div>
  </div>

  <!-- AKTIONEN -->
  <div class="sec">
    <div class="sec-head">
      <div class="sec-icon ic-blue"><svg viewBox="0 0 24 24"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/></svg></div>
      <span class="sec-title">Aktionen</span>
    </div>
    <div class="sec-body">
      <div class="actions">
        <button class="btn btn-primary" onclick="exportXLS()"><svg viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="8" y1="13" x2="16" y2="13"/><line x1="8" y1="17" x2="16" y2="17"/></svg> XLS exportieren</button>
        <button class="btn btn-outline" onclick="printReport()"><svg viewBox="0 0 24 24"><polyline points="6 9 6 2 18 2 18 9"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8"/></svg> Bericht drucken</button>
      </div>
    </div>
  </div>

</div>
<div class="toast" id="toast"></div>

<script>
// â”€â”€ DST-AWARE TIME ENGINE â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Europa/Berlin DST rules (covers all years):
// Sommerzeit: letzter Sonntag MÃ¤rz 02:00 -> 03:00 (+1h, Uhr vor)
// Winterzeit: letzter Sonntag Oktober 03:00 -> 02:00 (-1h, Uhr zurÃ¼ck)

function getLastSunday(year, month) {
  // month: 0-based JS month (2=MÃ¤rz, 9=Oktober)
  const d = new Date(year, month + 1, 0); // last day of month
  d.setDate(d.getDate() - d.getDay()); // go back to Sunday
  return d;
}

function getBerlinOffsetMinutes(utcMs) {
  // Returns UTC offset in minutes for Europe/Berlin at a given UTC timestamp
  // Base offset: CET = +60min, CEST = +120min
  // DST start: last Sunday March 01:00 UTC (= 02:00 CET)
  // DST end:   last Sunday October 01:00 UTC (= 03:00 CEST)
  const d = new Date(utcMs);
  const year = d.getUTCFullYear();

  const dstStart = getLastSunday(year, 2); // last Sunday in March
  const dstEnd = getLastSunday(year, 9);   // last Sunday in October

  // DST starts at 01:00 UTC on that Sunday
  const dstStartUTC = Date.UTC(year, 2, dstStart.getDate(), 1, 0, 0);
  // DST ends at 01:00 UTC on that Sunday
  const dstEndUTC = Date.UTC(year, 9, dstEnd.getDate(), 1, 0, 0);

  if (utcMs >= dstStartUTC && utcMs < dstEndUTC) {
    return 120; // CEST = UTC+2
  }
  return 60; // CET = UTC+1
}

function localToUTC(dateStr, timeStr) {
  // dateStr: 'YYYY-MM-DD', timeStr: 'HH:MM'
  // Returns UTC milliseconds, handling DST correctly for Europe/Berlin
  const [y, mo, dd] = dateStr.split('-').map(Number);
  const [h, m] = timeStr.split(':').map(Number);

  // First estimate: assume CET (+60min)
  const utcEstimate = Date.UTC(y, mo - 1, dd, h, m) - 60 * 60000;
  const offset = getBerlinOffsetMinutes(utcEstimate);
  // Recalculate with correct offset
  return Date.UTC(y, mo - 1, dd, h, m) - offset * 60000;
}

function dstInfoForShift(startDateStr, startTimeStr, endDateStr, endTimeStr) {
  // Returns info about DST transitions within this shift
  const startUTC = localToUTC(startDateStr, startTimeStr);
  let endUTC = localToUTC(endDateStr, endTimeStr);
  if (endUTC <= startUTC) endUTC += 24 * 3600000; // next day

  const startOffset = getBerlinOffsetMinutes(startUTC);
  const endOffset = getBerlinOffsetMinutes(endUTC - 1);

  if (startOffset !== endOffset) {
    const diff = endOffset - startOffset; // +60 = Winterzeit (extra hour), -60 = Sommerzeit (lost hour)
    return { hasDST: true, diff, type: diff > 0 ? 'winter' : 'summer' };
  }
  return { hasDST: false, diff: 0 };
}

function calcShiftDurationMinutes(startDateStr, startTimeStr, endDateStr, endTimeStr) {
  // Returns REAL minutes (UTC-based), accounting for DST
  let startUTC = localToUTC(startDateStr, startTimeStr);
  let endUTC = localToUTC(endDateStr, endTimeStr);

  // Handle overnight: if end is before or equal to start, add a day
  if (endUTC <= startUTC) endUTC += 24 * 3600000;

  return Math.round((endUTC - startUTC) / 60000);
}

// â”€â”€ NIGHT HOUR LOGIC â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Night = 23:00â€“06:00 local Berlin time (DST-aware)
// We step through each minute in UTC and check local hour

function getBerlinLocalHourMinute(utcMs) {
  const offset = getBerlinOffsetMinutes(utcMs);
  const localMs = utcMs + offset * 60000;
  const d = new Date(localMs);
  return { h: d.getUTCHours(), m: d.getUTCMinutes(), weekday: d.getUTCDay() };
}

function isNightMinute(utcMs) {
  const { h } = getBerlinLocalHourMinute(utcMs);
  return h >= 23 || h < 6;
}

function isWeekendNightMinuteUTC(utcMs) {
  // Weekend night: the night of Fr->Sa (Fr 23:00 â€“ Sa 06:00) or Sa->So (Sa 23:00 â€“ So 06:00)
  if (!isNightMinute(utcMs)) return false;
  const { h, weekday } = getBerlinLocalHourMinute(utcMs);
  // For after-midnight hours (00:00â€“05:59), the "night start day" is the previous day
  const nightStartDay = h < 6 ? (weekday + 6) % 7 : weekday; // shift back
  // nightStartDay: 4=Freitag (starts Fr->Sa night), 5=Samstag (starts Sa->So night)
  return nightStartDay === 5 || nightStartDay === 6;
}

function calcShift(entry) {
  let startUTC = localToUTC(entry.date, entry.start);
  let endUTC = localToUTC(entry.date.replace(/(\d{4}-\d{2}-)(\d{2})/, (_, p, d) => p + d), entry.end);

  // We need end date - compute it properly
  // If the file has explicit end date, use it; otherwise infer
  const endDate = entry.endDate || entry.date;
  endUTC = localToUTC(endDate, entry.end);
  if (endUTC <= startUTC) endUTC += 24 * 3600000;

  const totalMinutes = Math.round((endUTC - startUTC) / 60000);
  let normalNightMin = 0, weekendNightMin = 0;

  // Step through each minute
  for (let i = 0; i < totalMinutes; i++) {
    const tUTC = startUTC + i * 60000;
    if (!isNightMinute(tUTC)) continue;
    if (isWeekendNightMinuteUTC(tUTC)) weekendNightMin++;
    else normalNightMin++;
  }

  const dst = dstInfoForShift(entry.date, entry.start, endDate, entry.end);
  return { totalMinutes, normalNightMin, weekendNightMin, dst };
}

// â”€â”€ STATE â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
let entries = [];
let nextId = 1;

function pad(n) { return String(n).padStart(2, '0'); }
function fmtMin(m) { return `${Math.floor(m/60)}h ${pad(m%60)}m`; }
function fmtEuro(v) { return v.toFixed(2).replace('.', ',') + ' â‚¬'; }
function timeToMin(t) { const [h,m] = t.split(':').map(Number); return h*60+m; }
function getSettings() {
  return {
    baseWage: parseFloat(document.getElementById('baseWage').value)||0,
    nightBonusEUR: parseFloat(document.getElementById('nightBonus').value)||0,
    weekendNightBonusEUR: parseFloat(document.getElementById('weekendNightBonus').value)||0
  };
}

function inferEndDate(entry) {
  if (entry.endDate) return entry.endDate;
  // If end time <= start time, end is next day
  const sm = timeToMin(entry.start), em = timeToMin(entry.end);
  if (em <= sm) {
    const [y,mo,dd] = entry.date.split('-').map(Number);
    const next = new Date(y, mo-1, dd+1);
    return `${next.getFullYear()}-${pad(next.getMonth()+1)}-${pad(next.getDate())}`;
  }
  return entry.date;
}

// â”€â”€ RENDER â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function renderTable() {
  const tbody = document.getElementById('entriesBody');
  if (!entries.length) {
    tbody.innerHTML = '<tr><td colspan="8"><div class="empty-state">Keine Schichten eingetragen.</div></td></tr>';
    return;
  }
  const sorted = [...entries].sort((a,b) => a.date.localeCompare(b.date) || a.start.localeCompare(b.start));
  const dstRows = [];
  tbody.innerHTML = sorted.map(e => {
    const endDate = inferEndDate(e);
    const c = calcShift({...e, endDate});
    const [y,mo,dd] = e.date.split('-').map(Number);
    const dt = new Date(y, mo-1, dd);
    const dayName = dt.toLocaleDateString('de-DE', {weekday:'short'});
    const dateDisplay = `${pad(dd)}.${pad(mo)}.${y} (${dayName})`;

    // Clock hours (naive)
    const sm = timeToMin(e.start), em = timeToMin(e.end);
    const clockMin = em > sm ? em - sm : em + 1440 - sm;

    let tags = '';
    if (c.weekendNightMin > 0) tags += '<span class="tag tag-wknight">WE-Nacht</span> ';
    else if (c.normalNightMin > 0) tags += '<span class="tag tag-night">Normal Nacht</span> ';
    else tags += '<span style="color:var(--text-muted);font-size:11px">Tag</span>';

    let dstTag = '';
    if (c.dst.hasDST) {
      if (c.dst.type === 'summer') {
        dstTag = '<span class="tag tag-dst"><i class="ti ti-clock-minus" style="font-size:9px"></i> âˆ’1h Sommerzeit</span>';
        dstRows.push({date: dateDisplay, type: 'summer'});
      } else {
        dstTag = '<span class="tag tag-dst"><i class="ti ti-clock-plus" style="font-size:9px"></i> +1h Winterzeit</span>';
        dstRows.push({date: dateDisplay, type: 'winter'});
      }
    }

    const realDiff = c.totalMinutes - clockMin;
    const realColor = realDiff < 0 ? 'var(--red)' : realDiff > 0 ? 'var(--green)' : 'var(--text)';

    return `<tr>
      <td>${dateDisplay}</td>
      <td>${e.start} â€“ ${e.end}</td>
      <td style="color:var(--text-muted)">${fmtMin(clockMin)}</td>
      <td style="color:${realColor};font-weight:${realDiff!==0?'600':'400'}">${fmtMin(c.totalMinutes)}${realDiff!==0?` <span style="font-size:10px">(${realDiff>0?'+':''}${realDiff/60}h DST)</span>`:''}</td>
      <td style="color:var(--night)">${c.normalNightMin>0?fmtMin(c.normalNightMin):'â€“'}</td>
      <td style="color:var(--purple)">${c.weekendNightMin>0?fmtMin(c.weekendNightMin):'â€“'}</td>
      <td>${tags}${dstTag}</td>
      <td><button class="del-btn" onclick="deleteEntry(${e.id})" aria-label="LÃ¶schen"><i class="ti ti-trash" aria-hidden="true"></i></button></td>
    </tr>`;
  }).join('');

  // Update DST banner
  const banner = document.getElementById('dstBanner');
  if (dstRows.length) {
    banner.className = 'dst-banner visible';
    banner.innerHTML = '<strong><i class="ti ti-clock-exclamation" style="font-size:13px;vertical-align:-2px"></i> Zeitumstellung erkannt!</strong><br>' +
      dstRows.map(r => r.type === 'summer'
        ? `â€¢ <strong>${r.date}:</strong> Sommerzeit (Uhr vor) â€“ diese Schicht hat 1 echte Stunde weniger als die Uhr zeigt.`
        : `â€¢ <strong>${r.date}:</strong> Winterzeit (Uhr zurÃ¼ck) â€“ diese Schicht hat 1 echte Stunde mehr als die Uhr zeigt.`
      ).join('<br>');
  } else {
    banner.className = 'dst-banner';
  }
}

function renderResults() {
  const s = getSettings();
  let totalNormal=0, totalWeekend=0, totalReal=0;
  entries.forEach(e => {
    const c = calcShift({...e, endDate: inferEndDate(e)});
    totalReal += c.totalMinutes;
    totalNormal += c.normalNightMin;
    totalWeekend += c.weekendNightMin;
  });
  document.getElementById('resNight').textContent = fmtMin(totalNormal);
  document.getElementById('resWkNight').textContent = fmtMin(totalWeekend);
  const nb = (totalNormal/60)*s.nightBonusEUR;
  const wb = (totalWeekend/60)*s.weekendNightBonusEUR;
  document.getElementById('resNightBonus').textContent = fmtEuro(nb)+' Zuschlag';
  document.getElementById('resWkNightBonus').textContent = fmtEuro(wb)+' Zuschlag';
  document.getElementById('resTotalBonus').textContent = fmtEuro(nb+wb);

  const bd = document.getElementById('breakdownSection');
  if (!entries.length) { bd.style.display='none'; return; }
  bd.style.display='';
  const tH=totalReal/60, nH=totalNormal/60, wH=totalWeekend/60;
  const basePay=tH*s.baseWage, nPay=nH*s.nightBonusEUR, wPay=wH*s.weekendNightBonusEUR;
  document.getElementById('breakdownBody').innerHTML=`
    <tr><td class="highlight">Alle Arbeitsstunden (real, UTC)</td><td class="num">${tH.toFixed(2)}</td><td class="num">${fmtEuro(s.baseWage)}</td><td class="num highlight">${fmtEuro(basePay)}</td></tr>
    <tr><td class="highlight">Normale Nachtstunden</td><td class="num">${nH.toFixed(2)}</td><td class="num">${fmtEuro(s.nightBonusEUR)}</td><td class="num" style="color:var(--accent)">${fmtEuro(nPay)}</td></tr>
    <tr><td class="highlight">WE-Nachtstunden (Fr &amp; Sa Nacht)</td><td class="num">${wH.toFixed(2)}</td><td class="num">${fmtEuro(s.weekendNightBonusEUR)}</td><td class="num" style="color:var(--accent)">${fmtEuro(wPay)}</td></tr>
    <tr style="border-top:2px solid var(--border);font-weight:700"><td class="highlight">GESAMT (Lohn + ZuschlÃ¤ge)</td><td class="num">${tH.toFixed(2)}</td><td class="num">â€“</td><td class="num highlight" style="font-size:16px">${fmtEuro(basePay+nPay+wPay)}</td></tr>
  `;
}

// â”€â”€ CRUD â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function addEntry() {
  const date = document.getElementById('inputDate').value;
  const start = document.getElementById('inputStart').value.trim();
  const end = document.getElementById('inputEnd').value.trim();
  if (!date||!start||!end) return showToast('Alle Felder fÃ¼llen!', true);
  if (!/^([0-1]?\d|2[0-3]):[0-5]\d$/.test(start)||!/^([0-1]?\d|2[0-3]):[0-5]\d$/.test(end))
    return showToast('UngÃ¼ltiges Zeitformat (HH:MM)', true);
  entries.push({id:nextId++, date, start, end});
  renderTable(); renderResults(); showToast('Schicht hinzugefÃ¼gt');
}

function deleteEntry(id) { entries=entries.filter(e=>e.id!==id); renderTable(); renderResults(); showToast('GelÃ¶scht'); }
function clearAllEntries() {
  if (!entries.length) return showToast('Keine Schichten vorhanden', true);
  if (confirm('Wirklich ALLE Schichten lÃ¶schen?')) { entries=[]; renderTable(); renderResults(); showToast('Alle gelÃ¶scht'); }
}

function showToast(msg, err=false) {
  const t=document.getElementById('toast');
  t.textContent=msg; t.style.borderColor=err?'var(--red)':'var(--green)'; t.style.color=err?'var(--red)':'var(--green)';
  t.classList.add('show'); setTimeout(()=>t.classList.remove('show'),2500);
}

// â”€â”€ IMPORT â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function parseDate(s) {
  if (!s) return null;
  const str = String(s).trim();
  let m = str.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (m) return `${m[3]}-${pad(m[2])}-${pad(m[1])}`;
  m = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return `${m[1]}-${pad(m[2])}-${pad(m[3])}`;
  return null;
}
function parseTime(s) {
  if (!s) return null;
  const str = String(s).trim();
  if (/^\d{1,2}:\d{2}$/.test(str)) { const [h,m]=str.split(':'); return `${pad(parseInt(h))}:${pad(parseInt(m))}`; }
  return null;
}

function importFromPaste() {
  const text = document.getElementById('pasteArea').value.trim();
  if (!text) return showToast('Bitte Daten einfÃ¼gen', true);
  let imported = 0;
  text.split('\n').forEach(line => {
    const parts = line.split(/\t|,/).map(s=>s.trim());
    if (parts.length < 3) return;
    const date=parseDate(parts[0]), start=parseTime(parts[1]), end=parseTime(parts[2]);
    if (date&&start&&end) { entries.push({id:nextId++,date,start,end}); imported++; }
  });
  if (imported) { renderTable(); renderResults(); showToast(`${imported} Schicht(en) importiert`); document.getElementById('pasteArea').value=''; }
  else showToast('Keine gÃ¼ltigen Daten', true);
}

document.getElementById('fileInput').addEventListener('change', function(e) {
  const file = e.target.files[0]; if (!file) return;
  const reader = new FileReader();
  reader.onload = function(ev) {
    try {
      const data = new Uint8Array(ev.target.result);
      const wb = XLSX.read(data, {type:'array'});
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, {header:1, defval:''});
      let imported=0, headerRow=-1;
      for (let i=0;i<rows.length;i++) {
        if (rows[i]&&rows[i][0]&&String(rows[i][0]).includes('Schicht Beginn')) { headerRow=i; break; }
      }
      if (headerRow===-1) { showToast('Keine Kopfzeile gefunden', true); return; }
      const headers = rows[headerRow];
      let sDateC=-1, sTimeC=-1, eDateC=-1, eTimeC=-1;
      headers.forEach((h,idx)=>{
        const c=String(h||'').toLowerCase();
        if (c.includes('schicht beginn')&&!c.includes('time')) sDateC=idx;
        if (c.includes('schicht beginn')&&c.includes('time')) sTimeC=idx;
        if (c.includes('schicht ende')&&!c.includes('time')) eDateC=idx;
        if (c.includes('schicht ende')&&c.includes('time')) eTimeC=idx;
      });
      if (sDateC===-1||sTimeC===-1||eDateC===-1||eTimeC===-1) { showToast('Spalten nicht gefunden', true); return; }
      for (let i=headerRow+1;i<rows.length;i++) {
        const row=rows[i];
        if (!row||!row.length) continue;
        const sd=parseDate(row[sDateC]), st=parseTime(row[sTimeC]);
        const ed=parseDate(row[eDateC]), et=parseTime(row[eTimeC]);
        if (sd&&st&&et) {
          entries.push({id:nextId++, date:sd, start:st, end:et, endDate:ed||undefined});
          imported++;
        }
      }
      if (imported) { renderTable(); renderResults(); showToast(`${imported} Schichten importiert`); }
      else showToast('Keine gÃ¼ltigen Schichten', true);
    } catch(err) { showToast('Fehler: '+err.message, true); }
    document.getElementById('fileInput').value='';
  };
  reader.readAsArrayBuffer(file);
});

// â”€â”€ EXPORT â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function exportXLS() {
  if (!entries.length) return showToast('Keine Daten', true);
  const s=getSettings(), name=document.getElementById('employeeName').value||'Mitarbeiter';
  const data=[['Mitarbeiter',name],['Export',new Date().toLocaleDateString('de-DE')],[],
    ['Datum','Wochentag','Start','Ende','Uhr-Std','Echte Std (DST)','DST-Hinweis','Normal-Nacht Min','WE-Nacht Min','Normal-Zuschlag â‚¬','WE-Zuschlag â‚¬','Gesamt-Zuschlag â‚¬']];
  [...entries].sort((a,b)=>a.date.localeCompare(b.date)).forEach(e => {
    const endDate=inferEndDate(e);
    const c=calcShift({...e,endDate});
    const [y,m,d]=e.date.split('-').map(Number), dt=new Date(y,m-1,d);
    const sm=timeToMin(e.start),em=timeToMin(e.end), clockMin=em>sm?em-sm:em+1440-sm;
    const nB=(c.normalNightMin/60)*s.nightBonusEUR, wB=(c.weekendNightMin/60)*s.weekendNightBonusEUR;
    const dstNote=c.dst.hasDST?(c.dst.type==='summer'?'-1h Sommerzeit':'+1h Winterzeit'):'';
    data.push([`${pad(d)}.${pad(m)}.${y}`,dt.toLocaleDateString('de-DE',{weekday:'long'}),e.start,e.end,
      (clockMin/60).toFixed(2),(c.totalMinutes/60).toFixed(2),dstNote,
      c.normalNightMin,c.weekendNightMin,nB.toFixed(2),wB.toFixed(2),(nB+wB).toFixed(2)]);
  });
  const ws=XLSX.utils.aoa_to_sheet(data);
  const wbOut=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wbOut,ws,'Nachtzuschlag_Report');
  XLSX.writeFile(wbOut,`Nachtzuschlag_${name.replace(/[^a-zA-Z0-9]/g,'_')}_${new Date().toISOString().slice(0,10)}.xlsx`);
  showToast('XLS exportiert');
}
function printReport() { if(entries.length) window.print(); else showToast('Keine Daten',true); }

// â”€â”€ INIT â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
document.getElementById('inputDate').valueAsDate = new Date();
['baseWage','nightBonus','weekendNightBonus'].forEach(id =>
  document.getElementById(id).addEventListener('input', ()=>{ renderResults(); renderTable(); })
);
renderTable(); renderResults();
</script>
</body>
</html>
