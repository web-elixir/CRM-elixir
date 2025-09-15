function normKey(s) {
  return String(s || '')
    .toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]/g, '');
}

// --------- DÉTECTION DES COLONNES (nouveau schéma + anciens alias) ----------
function getSheetCtx_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getDataRange();
  const values = range.getValues();
  const display = range.getDisplayValues();
  const headers = values[0] || [];
  const keys = headers.map(normKey);

  const findIdx = (aliases) => {
    for (let a of aliases) {
      const i = keys.indexOf(a);
      if (i >= 0) return i;
    }
    return -1;
  };

  const idx = {
    societe:      findIdx(['societe']),
    magazine:     findIdx(['magazine', 'magazinebvvougb']),
    dateRappel:   findIdx(['daterappel', 'datederappel', 'datederapprel', 'datederapppl', 'relance']),
    commentaires: findIdx(['commentaires', 'notes']),
    telephone:    findIdx(['telephone']),
    mail:         findIdx(['mail', 'email']),
    nom:          findIdx(['nom']),
    prenom:       findIdx(['prenom']),
    poste:        findIdx(['poste']),
    adresse:      findIdx(['adresse']),
    // pour compat dans doGet (facultatif si tu veux garder les if de secours)
    relance:      findIdx(['relance']),
    notes:        findIdx(['notes']),
    statut:       findIdx(['statut']),
  };

  return { sheet, values, display, headers, keys, idx };
}

function formatColsIfPossible_(sheet, idx) {
  const last = sheet.getLastRow();
  if (last <= 1) return; // seulement entêtes
  if (idx.telephone >= 0) sheet.getRange(2, idx.telephone + 1, last - 1, 1).setNumberFormat('@');
  if (idx.dateRappel >= 0) sheet.getRange(2, idx.dateRappel + 1, last - 1, 1).setNumberFormat('yyyy-mm-dd');
}

function toISO_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const s = String(v || '').trim();
  return /^(\d{4})-(\d{2})-(\d{2})$/.test(s) ? s : '';
}

// Retourne la valeur à écrire pour un header NORMALISÉ donné, à partir d'un objet lc normalisé (synonymes gérés)
function fromLcForHeader_(headerNorm, lc) {
  // map des synonymes entrants -> champs cibles
  // lc contient déjà des clés normalisées du payload (ex: 'societe', 'email'->'mail', etc.)
  const pick = (...names) => {
    for (const n of names) {
      if (lc[n] != null && lc[n] !== '') return lc[n];
    }
    return '';
  };

  switch (headerNorm) {
    case 'societe':                return pick('societe');
    case 'magazine':
    case 'magazinebvvougb':        return pick('magazine');
    case 'daterappel':
    case 'datederappel':
    case 'datederapprel':
    case 'datederapppl':
    case 'relance':                return toISO_(pick('daterappel','datederappel','relance','date')); // robustesse
    case 'commentaires':
    case 'notes':                  return pick('commentaires','notes');
    case 'telephone':              return String(pick('telephone'));
    case 'mail':
    case 'email':                  return pick('mail','email');
    case 'nom':                    return pick('nom');
    case 'prenom':                 return pick('prenom');
    case 'poste':                  return pick('poste');
    case 'adresse':                return pick('adresse');
    // champs non utilisés par le front mais qu'on évite d'écraser si présents
    case 'statut':                 return pick('statut');
    default:                       return ''; // colonnes inconnues -> vide
  }
}

/** ---------- GET ---------- */
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'list';
  if (action !== 'list') {
    return ContentService.createTextOutput(JSON.stringify({ error: 'Action GET inconnue' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const { values, display, headers, idx } = getSheetCtx_();
  const out = [];

  // helper lecture SÛR dans UNE LIGNE (et pas dans la matrice)
  const at = (rowArr, j) => (j >= 0 ? (rowArr[j] ?? '') : '');

  for (let i = 1; i < values.length; i++) {
    const rowV = values[i]  || []; // valeurs brutes
    const rowD = display[i] || []; // valeurs affichées (formatées)

    const o = { gsId: i };

    // 🔧 Utiliser rowV/rowD, pas "values"
    o.societe      = at(rowV, idx.societe);
    o.magazine     = at(rowV, idx.magazine);
    o.dateRappel   = toISO_(at(rowV, idx.dateRappel));
    o.commentaires = at(rowV, idx.commentaires);

    // téléphone: affichage pour conserver zéros/espaces
    o.telephone    = at(rowD, idx.telephone);

    // mail: brut
    o.mail         = at(rowV, idx.mail);

    o.nom          = at(rowV, idx.nom);
    o.prenom       = at(rowV, idx.prenom);
    o.poste        = at(rowV, idx.poste);
    o.adresse      = at(rowV, idx.adresse);

    // Compat ascendante (si colonnes anciennes existent encore)
    if (!o.dateRappel && idx.relance >= 0)      o.dateRappel   = toISO_(at(rowV, idx.relance));
    if (!o.commentaires && idx.notes >= 0)      o.commentaires = at(rowV, idx.notes);

    out.push(o);
  }

  return ContentService.createTextOutput(JSON.stringify({ rows: out }))
    .setMimeType(ContentService.MimeType.JSON);
}

/** ---------- POST ---------- */
function doPost(e) {
  const { sheet, headers, idx } = getSheetCtx_();

  // action + payload (form-urlencoded OU JSON)
  let action = (e && e.parameter && e.parameter.action) || '';
  let body = {};
  if (e && e.postData && e.postData.contents) {
    try { body = JSON.parse(e.postData.contents); } catch (_) {}
    if (!action && body.action) action = body.action;
  }

  let incomingRows = [];
  if (e && e.parameter && e.parameter.rows) {
    try { incomingRows = JSON.parse(e.parameter.rows); } catch (_) {}
  } else if (body.rows) {
    incomingRows = body.rows;
  }

  try {
    if (action === 'bulkUpsert') {
      const results = [];

      incomingRows.forEach((contact) => {
        // normaliser toutes les clés entrantes
        const lc = {};
        Object.keys(contact || {}).forEach(k => lc[normKey(k)] = contact[k]);
        const gsId = Number(lc.gsid) || null;

        // mise à jour
        if (gsId && gsId > 0 && (gsId + 1) <= sheet.getLastRow()) {
          const realRow = gsId + 1;
          headers.forEach((h, col) => {
            const headerNorm = normKey(h);
            const val = fromLcForHeader_(headerNorm, lc);

            // Formats spéciaux
            if (col === idx.telephone && idx.telephone >= 0) {
              sheet.getRange(realRow, col + 1).setNumberFormat('@').setValue(String(val || ''));
            } else if (col === idx.dateRappel && idx.dateRappel >= 0) {
              sheet.getRange(realRow, col + 1).setNumberFormat('yyyy-mm-dd').setValue(toISO_(val) || '');
            } else {
              sheet.getRange(realRow, col + 1).setValue(val);
            }
          });
          results.push({ gsId });
        } else {
          // création
          const newRow = headers.map((h, col) => {
            const headerNorm = normKey(h);
            let val = fromLcForHeader_(headerNorm, lc);
            if (col === idx.dateRappel && idx.dateRappel >= 0) val = toISO_(val);
            return val || '';
          });
          sheet.appendRow(newRow);
          const newRealRow = sheet.getLastRow();
          if (idx.telephone >= 0) sheet.getRange(newRealRow, idx.telephone + 1).setNumberFormat('@');
          if (idx.dateRappel >= 0) sheet.getRange(newRealRow, idx.dateRappel + 1).setNumberFormat('yyyy-mm-dd');
          results.push({ gsId: newRealRow - 1 });
        }
      });

      formatColsIfPossible_(sheet, idx);

      return ContentService.createTextOutput(JSON.stringify({ results }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'delete') {
      const idParam = (e && e.parameter && e.parameter.id) || (body && body.id);
      const gsId = Number(idParam);
      if (!gsId || gsId < 1) {
        return ContentService.createTextOutput(JSON.stringify({ error: 'ID invalide' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      const realRow = gsId + 1;
      const last = sheet.getLastRow();
      if (realRow < 2 || realRow > last) {
        return ContentService.createTextOutput(JSON.stringify({ error: 'ID hors limites' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      sheet.deleteRow(realRow);
      return ContentService.createTextOutput(JSON.stringify({ result: 'success' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ error: 'Action POST inconnue' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
