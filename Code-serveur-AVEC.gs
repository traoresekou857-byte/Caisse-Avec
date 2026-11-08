/* Serveur de la base globale – Outils de gestion des groupements AVEC
   © 2026 Sékou Traoré. À coller dans : Extensions > Apps Script. */
var KEY = 'avec-ea807024887f4b2e';
// >>> Écrivez votre mot de passe superviseur entre les apostrophes ci-dessous <<<
var ADMIN_PWD = 'ECRIVEZ-ICI-VOTRE-MOT-DE-PASSE';

var H_G = ['ID','Groupement','Région','Cercle','Commune','Village','Animateur','Envoyé par','Téléphone','Structure',
           'Membres actifs','Épargne totale','Fonds social','Prêts à récupérer','Prêts en retard','Argent en caisse',
           'Nb réunions','Dernière réunion','Créé le','Dernière mise à jour'];
var H_M = ['ID groupement','Groupement','Village','Membre','Téléphone','Parts','Statut','Épargne','Doit'];
var H_D = ['ID','Données (JSON, ne pas modifier)'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var d = JSON.parse(e.postData.contents);
    if (d.key !== KEY) return out({ok: false, err: 'cle'});
    if (d.action === 'delete') {
      removeRows(sh('Groupements', H_G), d.id);
      removeRows(sh('Membres', H_M), d.id);
      removeRows(sh('Donnees', H_D), d.id);
      return out({ok: true});
    }
    var g = d.group, s = d.summary || {}, f = d.from || {};
    upsert(sh('Groupements', H_G), g.id, [g.id, g.nom, g.region || '', g.cercle || '', g.commune || '', g.village || '',
      g.animateur || '', f.nom || '', f.tel || '', f.structure || '', s.membres || 0, s.epargne || 0, s.social || 0,
      s.encours || 0, s.retards || 0, s.caisse || 0, s.reunions || 0, s.derniere || '', g.cree || '', new Date()]);

    var ms = sh('Membres', H_M);
    removeRows(ms, g.id);
    var rows = (d.membres || []).map(function (m) {
      return [g.id, g.nom, g.village || '', m.nom, m.tel || '', m.parts || 1, m.statut || '', m.epargne || 0, m.doit || 0];
    });
    if (rows.length) ms.getRange(ms.getLastRow() + 1, 1, rows.length, H_M.length).setValues(rows);

    var json = JSON.stringify(g), chunks = [];
    for (var i = 0; i < json.length; i += 45000) chunks.push(json.substr(i, 45000));
    var ds = sh('Donnees', H_D);
    removeRows(ds, g.id);
    var row = [g.id].concat(chunks);
    ds.getRange(ds.getLastRow() + 1, 1, 1, row.length).setNumberFormat('@').setValues([row]);
    return out({ok: true});
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  if (!e || !e.parameter || e.parameter.key !== KEY) return out({ok: false, err: 'cle'});
  if (ADMIN_PWD.length < 6 || ADMIN_PWD === 'ECRIVEZ-ICI-VOTRE-MOT-DE-PASSE' || e.parameter.pwd !== ADMIN_PWD) {
    Utilities.sleep(1500);
    return out({ok: false, err: 'mdp'});
  }
  var v = sh('Donnees', H_D).getDataRange().getValues(), list = [];
  for (var r = 1; r < v.length; r++) {
    var j = v[r].slice(1).join('');
    if (j) { try { list.push(JSON.parse(j)); } catch (x) {} }
  }
  var gs = sh('Groupements', H_G).getDataRange().getValues(), from = {};
  for (r = 1; r < gs.length; r++) from[gs[r][0]] = {nom: gs[r][7], tel: gs[r][8], structure: gs[r][9], maj: gs[r][19]};
  return out({ok: true, groups: list, from: from});
}

function sh(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); s.appendRow(headers); s.setFrozenRows(1); }
  return s;
}
function findRow(s, id) {
  var n = s.getLastRow(); if (n < 2) return 0;
  var ids = s.getRange(2, 1, n - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}
function upsert(s, id, row) {
  var r = findRow(s, id);
  if (r) s.getRange(r, 1, 1, row.length).setValues([row]); else s.appendRow(row);
}
function removeRows(s, id) {
  var n = s.getLastRow(); if (n < 2) return;
  var ids = s.getRange(2, 1, n - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === String(id)) s.deleteRow(i + 2);
}
function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
