/**
 * Crée un menu personnalisé dans Google Sheets à l'ouverture du fichier.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🚗 Outils Candidatures')
    .addItem('Calculer les temps de trajet', 'ouvrirDialogueAdresse')
    .addSeparator()
    .addItem('🔑 Configurer la clé API Google Maps', 'configurerCleApi')
    .addToUi();
}

/**
 * Demande la clé API Google Maps et la stocke dans les propriétés du script.
 * La clé n'est jamais écrite en dur dans le code ni visible dans la feuille.
 */
function configurerCleApi() {
  var ui = SpreadsheetApp.getUi();
  var cleActuelle = PropertiesService.getScriptProperties().getProperty("GOOGLE_MAPS_API_KEY");

  var message = cleActuelle
    ? "Une clé est déjà enregistrée (se termine par …" + cleActuelle.slice(-4) + ").\n\nColle une nouvelle clé pour la remplacer, ou laisse vide pour annuler."
    : "Colle ta clé API Google Maps (Directions API activée).\n\nExemple : AIzaSy...";

  var reponse = ui.prompt("🔑 Clé API Google Maps", message, ui.ButtonSet.OK_CANCEL);
  if (reponse.getSelectedButton() !== ui.Button.OK) return;

  var cle = reponse.getResponseText().trim();
  if (!cle) {
    ui.alert("Aucune clé saisie, rien n'a été modifié.");
    return;
  }

  PropertiesService.getScriptProperties().setProperty("GOOGLE_MAPS_API_KEY", cle);
  ui.alert("✅ Clé enregistrée !\n\nTu peux maintenant lancer le calcul des temps de trajet.");
}

/**
 * Récupère la clé API stockée, ou null si absente.
 */
function getCleApi() {
  return PropertiesService.getScriptProperties().getProperty("GOOGLE_MAPS_API_KEY");
}

/**
 * Génère et affiche la fenêtre Pop-up avec la liste déroulante des adresses.
 */
function ouvrirDialogueAdresse() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!getCleApi()) {
    SpreadsheetApp.getUi().alert(
      "Aucune clé API Google Maps n'est configurée.\n\n" +
      "Va dans le menu « 🚗 Outils Candidatures » → « 🔑 Configurer la clé API Google Maps » " +
      "avant de lancer le calcul."
    );
    return;
  }

  var configSheet = ss.getSheetByName("Adresses de départ");

  if (!configSheet) {
    configSheet = ss.insertSheet("Adresses de départ");
    configSheet.appendRow(["Nom du lieu (Ex: Maison, Appart...)", "Adresse exacte"]);
    configSheet.appendRow(["Maison principale", "16 Rue Jules Vallès, 54490 Piennes"]);
    configSheet.getRange("A1:B1").setFontWeight("bold");

    SpreadsheetApp.getUi().alert(
      "Un onglet 'Adresses de départ' vient d'être créé !\n\n" +
      "Ajoute tes adresses dans ce tableau, puis relance le script."
    );
    return;
  }

  var data = configSheet.getDataRange().getValues();
  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert("L'onglet 'Adresses de départ' est vide. Ajoute au moins une adresse sous la ligne d'entête.");
    return;
  }

  var feuilleCible = ss.getActiveSheet();
  if (feuilleCible.getName() === "Adresses de départ") {
    var toutesFeuilles = ss.getSheets();
    for (var s = 0; s < toutesFeuilles.length; s++) {
      if (toutesFeuilles[s].getName() !== "Adresses de départ") {
        feuilleCible = toutesFeuilles[s];
        break;
      }
    }
  }
  PropertiesService.getDocumentProperties().setProperty("feuilleCibleNom", feuilleCible.getName());

  var optionsHtml = "";
  for (var i = 1; i < data.length; i++) {
    var nomLieu = data[i][0];
    var adresseExacte = data[i][1];
    if (nomLieu && adresseExacte) {
      var adressePropre = adresseExacte.toString().replace(/"/g, '&quot;');
      var nomPropre = nomLieu.toString().replace(/</g, '&lt;').replace(/>/g, '&gt;');
      optionsHtml += '<option value="' + adressePropre + '">' + nomPropre + '</option>';
    }
  }

  // En-têtes de l'onglet cible -> listes déroulantes pour choisir les colonnes.
  var headersCible = feuilleCible.getLastColumn() > 0
    ? feuilleCible.getRange(1, 1, 1, feuilleCible.getLastColumn()).getValues()[0]
    : [];
  var colAdresseHtml = optionsColonnes(headersCible, "Adresse");
  var colResultatHtml = optionsColonnes(headersCible, "Trajet");

  var htmlOutput = HtmlService.createHtmlOutput(getHtmlTemplate(optionsHtml, colAdresseHtml, colResultatHtml))
    .setWidth(460)
    .setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '🚗 Calculateur d\'itinéraires');
}

/**
 * Construit les <option> d'une liste de colonnes à partir des en-têtes.
 * value = numéro de colonne (1-based). L'en-tête égal à `nomParDefaut`
 * (ex: "Adresse" / "Trajet") est présélectionné s'il existe.
 */
function optionsColonnes(headers, nomParDefaut) {
  var html = "";
  for (var c = 0; c < headers.length; c++) {
    var nomCol = headers[c] ? headers[c].toString() : "";
    var libelle = nomCol || ("Colonne " + colToLetter(c + 1));
    var libellePropre = libelle.replace(/</g, "&lt;").replace(/>/g, "&gt;");
    var sel = (nomCol && nomCol === nomParDefaut) ? " selected" : "";
    html += '<option value="' + (c + 1) + '"' + sel + '>' + libellePropre + '</option>';
  }
  return html;
}

/**
 * Convertit un numéro de colonne (1-based) en lettre(s) de colonne (A, B, ..., AA).
 */
function colToLetter(col) {
  var lettre = "";
  while (col > 0) {
    var reste = (col - 1) % 26;
    lettre = String.fromCharCode(65 + reste) + lettre;
    col = Math.floor((col - 1) / 26);
  }
  return lettre;
}

/**
 * Démarre un calcul EN ARRIÈRE-PLAN.
 * Prépare la liste des lignes à traiter, enregistre l'état du "job" dans les
 * propriétés du document, puis crée un déclencheur temporel qui fera tourner
 * le calcul côté serveur, indépendamment de la fenêtre. On peut donc fermer
 * la fenêtre : le calcul continue.
 * Renvoie { total } (nombre de lignes à traiter).
 */
function demarrerCalculFond(opts) {
  if (!getCleApi()) {
    throw new Error("Aucune clé API configurée. Menu → 🔑 Configurer la clé API.");
  }

  opts = opts || {};
  var startAddress = opts.startAddress;
  var nomLieu = opts.nomLieu;
  var mode = opts.mode || "driving";                 // driving | walking | transit
  var transitMode = opts.transitMode || "";          // train | bus (si mode = transit)
  var idxAdresse = parseInt(opts.colAdresse, 10);    // 1-based
  var idxTemps = parseInt(opts.colResultat, 10);     // 1-based

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var nomFeuille = PropertiesService.getDocumentProperties().getProperty("feuilleCibleNom");
  var sheet = nomFeuille ? ss.getSheetByName(nomFeuille) : ss.getActiveSheet();
  if (!sheet) sheet = ss.getActiveSheet();

  if (!idxAdresse || !idxTemps) {
    throw new Error("Colonnes non sélectionnées. Choisis la colonne de l'adresse et celle du résultat.");
  }
  if (idxAdresse === idxTemps) {
    throw new Error("La colonne de l'adresse et celle du résultat doivent être différentes.");
  }

  var data = sheet.getDataRange().getValues();

  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var endAddress = data[i][idxAdresse - 1];
    var t = data[i][idxTemps - 1] ? data[i][idxTemps - 1].toString() : "";
    if (endAddress && (!t || t.trim() === "" || t.indexOf("Erreur") !== -1 || t.indexOf("Introuvable") !== -1)) {
      rows.push(i + 1);
    }
  }

  // On repart proprement : on stoppe un éventuel job précédent.
  supprimerTriggersFond();
  PropertiesService.getDocumentProperties().deleteProperty("jobArret");

  var job = {
    actif: rows.length > 0,
    startAddress: startAddress,
    nomLieu: nomLieu,
    mode: mode,
    transitMode: transitMode,
    sheetName: sheet.getName(),
    idxAdresse: idxAdresse,
    idxTemps: idxTemps,
    rows: rows,
    i: 0,
    total: rows.length,
    ok: 0, introuvable: 0, erreur: 0, quotaCourt: 0,
    statut: rows.length > 0 ? "en_cours" : "termine",
    detail: "",
    derniereAdresse: "",
    recents: [],
    maj: Date.now()
  };
  ecrireJob(job);

  if (rows.length > 0) {
    // Démarrage quasi immédiat via un déclencheur temporel "one-shot".
    ScriptApp.newTrigger("executerLotFond").timeBased().after(500).create();
  }

  return { total: rows.length };
}

/**
 * Handler du déclencheur : traite un lot de lignes côté serveur.
 * S'exécute par tranches de ~5 min (limite Apps Script) et se re-planifie
 * lui-même tant qu'il reste du travail.
 */
function executerLotFond() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return; // une autre exécution est déjà en cours

  try {
    var job = lireJob();
    if (!job || !job.actif) { supprimerTriggersFond(); return; }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(job.sheetName);
    if (!sheet) {
      job.actif = false; job.statut = "erreur"; job.detail = "Feuille introuvable.";
      ecrireJob(job); supprimerTriggersFond(); return;
    }

    var debut = Date.now();
    var LIMITE_MS = 300000; // 5 min, sous la limite d'exécution d'Apps Script
    var props = PropertiesService.getDocumentProperties();

    while (job.i < job.rows.length && job.statut === "en_cours") {
      if (Date.now() - debut > LIMITE_MS) break;

      // Arrêt demandé depuis la fenêtre ?
      if (props.getProperty("jobArret") === "1") {
        job.actif = false; job.statut = "arrete";
        break;
      }

      var row = job.rows[job.i];
      var address = sheet.getRange(row, job.idxAdresse).getValue();
      address = address ? address.toString() : "";
      job.derniereAdresse = address;

      if (!address) { job.i++; continue; }

      var r = traiterRowFond(sheet, row, job.idxTemps, job.startAddress, address, job.mode, job.transitMode);

      if (r.statut === "quota_jour") {
        job.statut = "quota_jour"; job.detail = r.detail; job.actif = false;
        break; // on n'incrémente pas i : la ligne pourra être reprise plus tard
      }
      if (r.statut === "config_erreur") {
        job.statut = "config_erreur"; job.detail = r.detail; job.actif = false;
        break;
      }

      if (r.statut === "ok") job.ok++;
      else if (r.statut === "introuvable") job.introuvable++;
      else if (r.statut === "quota_court") job.quotaCourt++;
      else job.erreur++;

      job.recents.unshift({ address: address, label: r.label, cls: r.cls });
      if (job.recents.length > 15) job.recents.pop();

      job.i++;
      job.maj = Date.now();
      ecrireJob(job);               // avancement visible en direct par la fenêtre
      SpreadsheetApp.flush();       // rend l'avancement visible dans la feuille
      Utilities.sleep(150);         // reste courtois avec l'API
    }

    if (job.i >= job.rows.length && job.statut === "en_cours") {
      job.statut = "termine"; job.actif = false;
    }

    ecrireJob(job);
    supprimerTriggersFond();

    if (job.actif) {
      // Encore du travail (tranche de 5 min atteinte) -> on reprogramme.
      ScriptApp.newTrigger("executerLotFond").timeBased().after(1000).create();
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * Traite une ligne côté serveur et écrit le résultat dans la feuille.
 * Renvoie { statut, label, cls, detail } — label/cls servent au journal affiché.
 */
function traiterRowFond(sheet, row, idxTemps, startAddress, address, mode, transitMode) {
  var res = appelerDirectionsApi(startAddress, address, mode, transitMode);
  if (res.status === "OVER_QUERY_LIMIT") {
    Utilities.sleep(2000);
    res = appelerDirectionsApi(startAddress, address, mode, transitMode);
  }

  if (res.status === "OK") {
    var minutes = Math.round(res.dureeSecondes / 60);
    sheet.getRange(row, idxTemps).setValue(minutes);
    return { statut: "ok", label: minutes + " min", cls: "ok" };
  }
  if (res.status === "ZERO_RESULTS" || res.status === "NOT_FOUND") {
    sheet.getRange(row, idxTemps).setValue("Introuvable");
    return { statut: "introuvable", label: "Introuvable", cls: "warn" };
  }
  if (res.status === "OVER_DAILY_LIMIT") {
    return { statut: "quota_jour", detail: res.detail, label: "Quota", cls: "err" };
  }
  if (res.status === "REQUEST_DENIED") {
    return { statut: "config_erreur", detail: res.detail, label: "Clé API", cls: "err" };
  }
  if (res.status === "OVER_QUERY_LIMIT") {
    // limite momentanée persistante -> on n'écrit rien, on avancera quand même
    return { statut: "quota_court", detail: res.detail, label: "Ignoré (ralenti)", cls: "warn" };
  }
  sheet.getRange(row, idxTemps).setValue("Erreur adresse");
  return { statut: "erreur", detail: res.detail, label: "Erreur", cls: "err" };
}

/**
 * Appelle la Directions API de Google Maps avec la clé configurée.
 * Renvoie { status, dureeSecondes, detail }.
 * status reprend les statuts officiels de l'API :
 *   OK, ZERO_RESULTS, NOT_FOUND, OVER_QUERY_LIMIT, OVER_DAILY_LIMIT,
 *   REQUEST_DENIED, INVALID_REQUEST, ERREUR_RESEAU
 */
function appelerDirectionsApi(origin, destination, mode, transitMode) {
  var cle = getCleApi();
  if (!cle) {
    return { status: "REQUEST_DENIED", detail: "Aucune clé API configurée." };
  }

  mode = mode || "driving"; // driving | walking | transit

  var url = "https://maps.googleapis.com/maps/api/directions/json"
    + "?origin=" + encodeURIComponent(origin)
    + "&destination=" + encodeURIComponent(destination)
    + "&mode=" + encodeURIComponent(mode)
    + "&region=fr"
    + "&language=fr"
    + "&key=" + encodeURIComponent(cle);

  // Pour les transports en commun, on précise train ou bus.
  if (mode === "transit" && transitMode) {
    url += "&transit_mode=" + encodeURIComponent(transitMode);
  }

  try {
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var json = JSON.parse(response.getContentText());
    var status = json.status;

    if (status === "OK" && json.routes && json.routes.length > 0) {
      return { status: "OK", dureeSecondes: json.routes[0].legs[0].duration.value };
    }

    // error_message donne le détail exact renvoyé par Google (clé, facturation, etc.)
    return { status: status, detail: json.error_message || status };
  } catch (e) {
    return { status: "ERREUR_RESEAU", detail: e.toString() };
  }
}

/**
 * Renvoie l'état courant du job de fond, pour que la fenêtre affiche
 * la progression (et puisse la retrouver après une réouverture).
 */
function getEtatProgression() {
  var job = lireJob();
  if (!job) return { existe: false };
  return {
    existe: true,
    actif: job.actif,
    statut: job.statut,
    total: job.total,
    traites: job.i,
    ok: job.ok,
    introuvable: job.introuvable,
    erreur: job.erreur,
    quotaCourt: job.quotaCourt,
    derniereAdresse: job.derniereAdresse,
    nomLieu: job.nomLieu,
    detail: job.detail,
    recents: job.recents || []
  };
}

/**
 * Demande l'arrêt du job de fond (bouton « Arrêter » de la fenêtre).
 * Le lot en cours verra le drapeau et s'arrêtera proprement.
 */
function arreterCalculFond() {
  PropertiesService.getDocumentProperties().setProperty("jobArret", "1");
  var job = lireJob();
  if (job) { job.actif = false; job.statut = "arrete"; ecrireJob(job); }
  supprimerTriggersFond();
  return true;
}

/* ---- Helpers état / déclencheurs ---- */

function lireJob() {
  var s = PropertiesService.getDocumentProperties().getProperty("jobFond");
  return s ? JSON.parse(s) : null;
}

function ecrireJob(job) {
  PropertiesService.getDocumentProperties().setProperty("jobFond", JSON.stringify(job));
}

function supprimerTriggersFond() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "executerLotFond") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * Template HTML du panel.
 */
function getHtmlTemplate(optionsHtml, colAdresseHtml, colResultatHtml) {
  return `
    <html>
      <head>
        <style>
          * { box-sizing: border-box; }
          body {
            font-family: 'Google Sans', 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            padding: 0;
            margin: 0;
            color: #202124;
            background: #ffffff;
          }
          .header {
            padding: 18px 20px 14px 20px;
            border-bottom: 1px solid #e8eaed;
          }
          .header h1 {
            font-size: 16px;
            margin: 0;
            font-weight: 600;
            display: flex;
            align-items: center;
            gap: 8px;
          }
          .header p {
            font-size: 12.5px;
            color: #5f6368;
            margin: 6px 0 0 0;
          }
          .body {
            padding: 18px 20px;
          }
          label {
            font-size: 13px;
            color: #202124;
            font-weight: 600;
            display: block;
            margin-bottom: 8px;
          }
          select {
            width: 100%;
            padding: 10px 12px;
            border-radius: 8px;
            border: 1px solid #dadce0;
            font-size: 14px;
            background-color: white;
            margin-bottom: 4px;
            outline: none;
            transition: border-color 0.15s ease;
          }
          select:focus { border-color: #1a73e8; box-shadow: 0 0 0 3px rgba(26,115,232,0.15); }

          .cols-row {
            display: flex;
            gap: 10px;
            margin-top: 14px;
          }
          .cols-col { flex: 1; min-width: 0; }
          .cols-col label { margin-bottom: 8px; }

          .btn-container {
            display: flex;
            justify-content: flex-end;
            gap: 8px;
            margin-top: 18px;
          }
          button {
            padding: 9px 18px;
            border: none;
            border-radius: 8px;
            cursor: pointer;
            font-size: 13.5px;
            font-weight: 500;
            transition: background-color 0.15s ease, transform 0.05s ease;
          }
          button:active { transform: scale(0.97); }
          .btn-primary { background-color: #1a73e8; color: white; }
          .btn-primary:hover { background-color: #1557b0; }
          .btn-primary:disabled { background-color: #c8d7f5; cursor: not-allowed; }
          .btn-secondary { background-color: #f1f3f4; color: #3c4043; }
          .btn-secondary:hover { background-color: #e8eaed; }

          #setupView, #progressView, #doneView { display: none; }
          #setupView.active, #progressView.active, #doneView.active { display: block; }

          .progress-track {
            width: 100%;
            height: 8px;
            background: #e8eaed;
            border-radius: 6px;
            overflow: hidden;
            margin: 14px 0 10px 0;
          }
          .progress-fill {
            height: 100%;
            width: 0%;
            background: linear-gradient(90deg, #1a73e8, #4285f4);
            border-radius: 6px;
            transition: width 0.25s ease;
          }
          .progress-stats {
            display: flex;
            justify-content: space-between;
            font-size: 12.5px;
            color: #5f6368;
            margin-bottom: 14px;
          }
          .progress-stats b { color: #202124; }

          .current-line {
            font-size: 12.5px;
            color: #3c4043;
            background: #f8f9fa;
            border: 1px solid #e8eaed;
            border-radius: 8px;
            padding: 10px 12px;
            margin-bottom: 12px;
            min-height: 16px;
            display: flex;
            align-items: center;
            gap: 8px;
          }
          .dot {
            width: 7px; height: 7px; border-radius: 50%;
            background: #1a73e8;
            flex-shrink: 0;
            animation: pulse 1s infinite ease-in-out;
          }
          @keyframes pulse {
            0%, 100% { opacity: 0.3; }
            50% { opacity: 1; }
          }

          .log {
            max-height: 130px;
            overflow-y: auto;
            border: 1px solid #e8eaed;
            border-radius: 8px;
            font-size: 12px;
            background: #fbfbfb;
          }
          .log-entry {
            padding: 6px 10px;
            border-bottom: 1px solid #f1f3f4;
            display: flex;
            justify-content: space-between;
            gap: 8px;
          }
          .log-entry:last-child { border-bottom: none; }
          .log-entry .addr { color: #3c4043; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
          .log-entry .val { flex-shrink: 0; font-weight: 600; }
          .val.ok { color: #188038; }
          .val.warn { color: #b06000; }
          .val.err { color: #d93025; }

          .summary-icon {
            font-size: 36px;
            text-align: center;
            margin-bottom: 6px;
          }
          .summary-title {
            text-align: center;
            font-size: 15px;
            font-weight: 600;
            margin-bottom: 4px;
          }
          .summary-sub {
            text-align: center;
            font-size: 12.5px;
            color: #5f6368;
            margin-bottom: 16px;
          }
          .summary-grid {
            display: flex;
            gap: 10px;
            margin-bottom: 6px;
          }
          .summary-card {
            flex: 1;
            border: 1px solid #e8eaed;
            border-radius: 8px;
            padding: 10px;
            text-align: center;
          }
          .summary-card .num { font-size: 18px; font-weight: 700; }
          .summary-card .label { font-size: 11px; color: #5f6368; margin-top: 2px; }

          .error-banner {
            font-size: 12.5px;
            color: #d93025;
            background: #fce8e6;
            border-radius: 8px;
            padding: 10px 12px;
            margin-top: 10px;
            display: none;
          }
          .bg-hint {
            font-size: 11.5px;
            color: #5f6368;
            background: #e8f0fe;
            border-radius: 8px;
            padding: 8px 10px;
            margin: 10px 0 0 0;
            line-height: 1.35;
          }
        </style>
      </head>
      <body>

        <div class="header">
          <p id="headerSub">Choisis un point de départ pour lancer le calcul des temps de trajet.</p>
        </div>

        <div class="body">

          <div id="setupView" class="active">
            <label for="adresse">Point de départ</label>
            <select id="adresse">${optionsHtml}</select>

            <label for="mode" style="margin-top:14px;">Mode de transport</label>
            <select id="mode">
              <option value="driving" selected>🚗 Voiture</option>
              <option value="walking">🚶 À pied</option>
              <option value="transit|train">🚆 Train</option>
              <option value="transit|bus">🚌 Bus</option>
            </select>

            <div class="cols-row">
              <div class="cols-col">
                <label for="colAdresse">Colonne de l'adresse</label>
                <select id="colAdresse">${colAdresseHtml}</select>
              </div>
              <div class="cols-col">
                <label for="colResultat">Colonne du résultat</label>
                <select id="colResultat">${colResultatHtml}</select>
              </div>
            </div>

            <div class="btn-container">
              <button class="btn-secondary" onclick="google.script.host.close()">Annuler</button>
              <button class="btn-primary" id="btnValider" onclick="demarrer()">Lancer le calcul</button>
            </div>
          </div>

          <div id="progressView">
            <div class="progress-track">
              <div class="progress-fill" id="progressFill"></div>
            </div>
            <div class="progress-stats">
              <span><b id="statDone">0</b> / <span id="statTotal">0</span> lignes traitées</span>
              <span id="statPercent">0%</span>
            </div>
            <div class="current-line">
              <span class="dot"></span>
              <span id="currentAddr">Préparation…</span>
            </div>
            <div class="log" id="logBox"></div>
            <div class="error-banner" id="errorBanner"></div>
            <p class="bg-hint">Le calcul tourne en arrière-plan : tu peux fermer cette fenêtre, il continuera tout seul. Rouvre le menu pour revoir la progression.</p>
            <div class="btn-container">
              <button class="btn-secondary" onclick="google.script.host.close()">Fermer (continue en fond)</button>
              <button class="btn-secondary" id="btnStop" onclick="arreter()">Arrêter</button>
            </div>
          </div>

          <div id="doneView">
            <div class="summary-icon" id="summaryIcon">✅</div>
            <div class="summary-title" id="summaryTitle">Calcul terminé</div>
            <div class="summary-sub" id="summarySub"></div>
            <div class="summary-grid">
              <div class="summary-card">
                <div class="num" id="cardOk" style="color:#188038;">0</div>
                <div class="label">Calculés</div>
              </div>
              <div class="summary-card">
                <div class="num" id="cardWarn" style="color:#b06000;">0</div>
                <div class="label">Introuvables</div>
              </div>
              <div class="summary-card">
                <div class="num" id="cardErr" style="color:#d93025;">0</div>
                <div class="label">Erreurs</div>
              </div>
            </div>
            <div class="btn-container">
              <button class="btn-primary" onclick="google.script.host.close()">Fermer</button>
            </div>
          </div>

        </div>

        <script>
          var nomLieu = "";
          var pollTimer = null;

          function showView(id) {
            ["setupView", "progressView", "doneView"].forEach(function(v) {
              document.getElementById(v).classList.toggle("active", v === id);
            });
          }

          // À l'ouverture : si un calcul tourne déjà en fond, on affiche direct sa progression.
          window.addEventListener("load", function() {
            google.script.run.withSuccessHandler(function(etat) {
              if (etat && etat.existe && etat.actif) {
                nomLieu = etat.nomLieu || "";
                document.getElementById("headerSub").textContent = "Depuis : " + nomLieu;
                showView("progressView");
                appliquerEtat(etat);
                google.script.run.executerLotFond(); // relance si rien ne tourne (verrou anti-doublon)
                demarrerPolling();
              }
            }).getEtatProgression();
          });

          function demarrer() {
            var select = document.getElementById("adresse");
            var startAddress = select.value;
            nomLieu = select.options[select.selectedIndex].text;

            // Mode : la valeur est "driving", "walking", "transit|train" ou "transit|bus".
            var modeSelect = document.getElementById("mode");
            var modeParts = modeSelect.value.split("|");
            var mode = modeParts[0];
            var transitMode = modeParts[1] || "";

            var colAdresse = document.getElementById("colAdresse").value;
            var colResultat = document.getElementById("colResultat").value;

            if (colAdresse === colResultat) {
              document.getElementById("headerSub").textContent =
                "Choisis deux colonnes différentes pour l'adresse et le résultat.";
              return;
            }

            document.getElementById("btnValider").disabled = true;
            document.getElementById("headerSub").textContent = "Préparation du calcul…";

            google.script.run
              .withSuccessHandler(function(result) {
                if (result.total === 0) {
                  document.getElementById("headerSub").textContent = "Rien à calculer";
                  showView("doneView");
                  document.getElementById("summaryIcon").textContent = "👍";
                  document.getElementById("summaryTitle").textContent = "Tout est déjà calculé";
                  document.getElementById("summarySub").textContent = "Aucune ligne ne nécessitait de calcul.";
                  return;
                }
                document.getElementById("statTotal").textContent = result.total;
                document.getElementById("headerSub").textContent = "Depuis : " + nomLieu;
                showView("progressView");
                google.script.run.executerLotFond(); // démarre tout de suite
                demarrerPolling();
              })
              .withFailureHandler(function(error) {
                document.getElementById("btnValider").disabled = false;
                document.getElementById("headerSub").textContent = "Erreur : " + (error && error.message ? error.message : error);
              })
              .demarrerCalculFond({
                startAddress: startAddress,
                nomLieu: nomLieu,
                mode: mode,
                transitMode: transitMode,
                colAdresse: colAdresse,
                colResultat: colResultat
              });
          }

          function demarrerPolling() {
            if (pollTimer) clearInterval(pollTimer);
            poll();
            pollTimer = setInterval(poll, 1500);
          }
          function arreterPolling() {
            if (pollTimer) { clearInterval(pollTimer); pollTimer = null; }
          }

          function poll() {
            google.script.run
              .withSuccessHandler(function(etat) {
                if (!etat || !etat.existe) return;
                appliquerEtat(etat);
                if (etat.statut !== "en_cours") {
                  arreterPolling();
                  afficherResume(etat);
                }
              })
              .withFailureHandler(function() { /* réseau : on retentera au prochain tick */ })
              .getEtatProgression();
          }

          // Met à jour barre / stats / journal à partir de l'état serveur.
          function appliquerEtat(etat) {
            document.getElementById("statTotal").textContent = etat.total;
            var pct = etat.total ? Math.round((etat.traites / etat.total) * 100) : 0;
            document.getElementById("progressFill").style.width = pct + "%";
            document.getElementById("statDone").textContent = etat.traites;
            document.getElementById("statPercent").textContent = pct + "%";
            document.getElementById("currentAddr").textContent =
              etat.statut === "en_cours"
                ? (etat.derniereAdresse ? "Calcul vers : " + etat.derniereAdresse : "Préparation…")
                : "Terminé";

            var logBox = document.getElementById("logBox");
            logBox.innerHTML = "";
            (etat.recents || []).forEach(function(r) {
              var entry = document.createElement("div");
              entry.className = "log-entry";
              var addrSpan = document.createElement("span");
              addrSpan.className = "addr";
              addrSpan.textContent = r.address;
              var valSpan = document.createElement("span");
              valSpan.className = "val " + r.cls;
              valSpan.textContent = r.label;
              entry.appendChild(addrSpan);
              entry.appendChild(valSpan);
              logBox.appendChild(entry);
            });
          }

          function afficherResume(etat) {
            document.getElementById("cardOk").textContent = etat.ok;
            document.getElementById("cardWarn").textContent = etat.introuvable + etat.quotaCourt;
            document.getElementById("cardErr").textContent = etat.erreur;

            if (etat.statut === "quota_jour") {
              document.getElementById("summaryIcon").textContent = "⏸️";
              document.getElementById("summaryTitle").textContent = "Quota atteint";
              document.getElementById("summarySub").textContent = etat.ok + " trajet(s) calculé(s) avant l'arrêt.";
            } else if (etat.statut === "config_erreur") {
              document.getElementById("summaryIcon").textContent = "🔑";
              document.getElementById("summaryTitle").textContent = "Clé API à vérifier";
              document.getElementById("summarySub").textContent =
                etat.detail || "Google a refusé la clé (clé, activation de la Directions API, ou facturation).";
            } else if (etat.statut === "arrete") {
              document.getElementById("summaryIcon").textContent = "🛑";
              document.getElementById("summaryTitle").textContent = "Calcul interrompu";
              document.getElementById("summarySub").textContent = etat.ok + " trajet(s) calculé(s) avant l'arrêt.";
            } else {
              document.getElementById("summaryIcon").textContent = "✅";
              document.getElementById("summaryTitle").textContent = "Calcul terminé";
              document.getElementById("summarySub").textContent = "Depuis : " + (etat.nomLieu || nomLieu);
            }
            showView("doneView");
          }

          function arreter() {
            var btn = document.getElementById("btnStop");
            btn.disabled = true;
            btn.textContent = "Arrêt…";
            google.script.run.arreterCalculFond();
            // Le prochain sondage verra statut = "arrete" et basculera sur le résumé.
          }
        </script>
      </body>
    </html>
  `;
}
