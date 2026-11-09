/* Serveur de la base globale – Outils de gestion des groupements AVEC
   © 2026 Sékou Traoré – Tous droits réservés. */
var KEY = 'avec-ea807024887f4b2e';
// >>> Écrivez votre mot de passe PRINCIPAL (accès à tous les projets) entre les apostrophes <<<
var ADMIN_PWD = 'ECRIVEZ-ICI-VOTRE-MOT-DE-PASSE';

var H_G = ['ID','Groupement','Région','Cercle','Commune','Village','Animateur','Envoyé par','Téléphone','Structure',
           'Membres actifs','Épargne totale','Fonds social','Prêts à récupérer','Prêts en retard','Argent en caisse',
           'Nb réunions','Dernière réunion','Créé le','Dernière mise à jour',
           'Code projet','Projet / Organisation','Femmes','Hommes','Filles','Garçons','Latitude','Longitude','Carte','Pays'];
var H_M = ['ID groupement','Groupement','Village','Membre','Téléphone','Parts','Statut','Épargne','Doit','Catégorie','Code projet'];
var H_D = ['ID','Données (JSON, ne pas modifier)'];
var H_P = ['Code projet','Projet / Organisation','Mot de passe superviseur du projet','Actif (OUI/NON)','Date de fin (facultatif)','Contact'];

function projets() {
  var v = sh('Projets', H_P).getDataRange().getValues(), list = {}, now = new Date();
  for (var r = 1; r < v.length; r++) {
    var code = String(v[r][0] || '').trim().toUpperCase();
    if (!code) continue;
    var actif = /^(oui|o|yes|y|1|vrai|true)$/i.test(String(v[r][3]).trim());
    var fin = v[r][4];
    if (!(fin instanceof Date) && fin) {
      var m = String(fin).trim().match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
      fin = m ? new Date(+m[3], +m[2] - 1, +m[1]) : null;
    }
    if (fin instanceof Date && fin.getTime() < now.getTime() - 86400000) actif = false;
    list[code] = {code: code, nom: String(v[r][1] || code), pwd: String(v[r][2] || ''), actif: actif};
  }
  return list;
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var d = JSON.parse(e.postData.contents);
    if (d.key !== KEY) return out({ok: false, err: 'cle'});
    var code = String(d.projet || '').trim().toUpperCase(), p = projets()[code];
    if (!p || !p.actif) return out({ok: false, err: 'projet'});
    var gsh = sh('Groupements', H_G), gid = d.action === 'delete' ? d.id : (d.group && d.group.id), ex = findRow(gsh, gid);
    if (ex) {
      var owner = String(gsh.getRange(ex, 21, 1, 1).getValues()[0][0] || '').toUpperCase();
      if (owner && owner !== code) return out({ok: false, err: 'autre-projet'});
    }
    if (d.action === 'delete') {
      removeRows(sh('Groupements', H_G), d.id);
      removeRows(sh('Membres', H_M), d.id);
      removeRows(sh('Donnees', H_D), d.id);
      return out({ok: true});
    }
    var g = d.group, s = d.summary || {}, f = d.from || {};
    g.projet = code;
    upsert(sh('Groupements', H_G), g.id, [g.id, g.nom, g.region || '', g.cercle || '', g.commune || '', g.village || '',
      g.animateur || '', f.nom || '', f.tel || '', f.structure || '', s.membres || 0, s.epargne || 0, s.social || 0,
      s.encours || 0, s.retards || 0, s.caisse || 0, s.reunions || 0, s.derniere || '', g.cree || '', new Date(),
      code, p.nom, s.femmes || 0, s.hommes || 0, s.filles || 0, s.garcons || 0,
      g.lat || '', g.lng || '', (g.lat && g.lng) ? 'https://maps.google.com/?q=' + g.lat + ',' + g.lng : '', g.pays || '']);
    var ms = sh('Membres', H_M);
    removeRows(ms, g.id);
    var rows = (d.membres || []).map(function (m) {
      return [g.id, g.nom, g.village || '', m.nom, m.tel || '', m.parts || 1, m.statut || '', m.epargne || 0, m.doit || 0, m.categorie || '', code];
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
  var q = (e && e.parameter) || {};
  if (q.key !== KEY) return out({ok: false, err: 'cle'});
  var all = projets();
  if (q.action === 'check') {
    var p = all[String(q.code || '').trim().toUpperCase()];
    if (!p || !p.actif) { Utilities.sleep(1000); return out({ok: false, err: 'code'}); }
    return out({ok: true, code: p.code, nom: p.nom});
  }
  var role = '', only = '';
  if (ADMIN_PWD.length >= 6 && ADMIN_PWD !== 'ECRIVEZ-ICI-VOTRE-MOT-DE-PASSE' && q.pwd === ADMIN_PWD) role = 'admin';
  else {
    for (var c in all) if (all[c].pwd.length >= 6 && q.pwd === all[c].pwd) { role = 'projet'; only = c; break; }
  }
  if (!role) { Utilities.sleep(1500); return out({ok: false, err: 'mdp'}); }
  var v = sh('Donnees', H_D).getDataRange().getValues(), list = [];
  for (var r = 1; r < v.length; r++) {
    var j = v[r].slice(1).join('');
    if (!j) continue;
    try { var g = JSON.parse(j); if (!only || String(g.projet || '').toUpperCase() === only) list.push(g); } catch (x) {}
  }
  var gs = sh('Groupements', H_G).getDataRange().getValues(), from = {};
  for (r = 1; r < gs.length; r++) from[gs[r][0]] = {nom: gs[r][7], tel: gs[r][8], structure: gs[r][9], maj: gs[r][19], projetNom: gs[r][21] || ''};
  var noms = {};
  for (c in all) if (role === 'admin' || c === only) noms[c] = all[c].nom;
  return out({ok: true, role: role, projet: only, projets: noms, groups: list, from: from});
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
