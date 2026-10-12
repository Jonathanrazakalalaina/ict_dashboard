/*************************************************************
 * UNFPA ICT DASHBOARD - IT SUPPORT PLATFORM
 * Backend Google Apps Script
 * -----------------------------------------------------------
 * Base de données : Google Sheets (auto-provisionnée)
 * Frontend        : Vue 3 + Vuetify 3 (voir Index.html)
 *************************************************************/

const APP_NAME   = 'UNFPA ICT Dashboard';
const TZ         = Session.getScriptTimeZone() || 'Africa/Nairobi';
const UPLOAD_FOLDER_ID = '12OdRLM5Mft4iHAc14RpvErzNnkNo5ge7';

const SHEETS = {
  TICKETS   : 'Tickets',
  ASSETS    : 'Assets',
  USERS     : 'Users',
  KB        : 'KnowledgeBase',
  COMMENTS  : 'Comments',
  AUDIT     : 'AuditLog',
  SETTINGS  : 'Settings',
  TRANSFERS : 'AssetTransfers'
};

const HEADERS = {
  TICKETS : ['ID','Titre','Description','Categorie','Priorite','Statut','Demandeur','Email','Departement','Bureau','Assigne','DateCreation','DateMiseAJour','DateResolution','SLA_Echeance','SLA_Statut','PiecesJointes'],
  ASSETS  : ['ID','Type','Marque','Modele','NumeroSerie','TagInventaire','Statut','UtilisePar','Departement','Bureau','DateAchat','FinGarantie','AssigneA','Cout','Notes','NomDocument','LienDocument','AssetNumber','Categorie','Etat','Book','Regions'],
  USERS   : ['ID','Nom','Email','Departement','Poste','Bureau','Telephone','Role','Statut','DateAjout','Regions'],
  KB      : ['ID','Titre','Categorie','Contenu','Auteur','DateCreation','DateMiseAJour','Vues','TagsMots'],
  COMMENTS: ['ID','TicketID','Auteur','Message','Date','Interne'],
  AUDIT   : ['ID','Date','Utilisateur','Action','Cible','Details'],
  SETTINGS: ['Cle','Valeur'],
  TRANSFERS: ['ID','AssetID','AssetLabel','Emprunteur','EmprunteurEmail','Departement','Bureau','DateEmprunt','DateRetourPrevue','DateRetourEffective','Statut','EffectuePar','Notes']
};

const PRIORITIES = ['Basse','Moyenne','Haute','Critique'];
const SLA_HOURS  = { 'Critique': 4, 'Haute': 8, 'Moyenne': 24, 'Basse': 72 };
const STATUSES   = ['Nouveau','Ouvert','En cours','En attente','Résolu','Fermé'];
const CATEGORIES = ['Matériel','Logiciel','Réseau','Compte/Accès','Appuis technique','Imprimante','Téléphonie','Codage','Sécurité','Acquisition/Achat IT', 'Suppliers', 'Autre'];
const ASSET_TYPES = ['Ordinateur portable','Ordinateur de bureau', 'Flybox', 'Visioconférence', 'Imprimante','Écran','Serveur','Routeur/Switch','Smartphone','Tablette','Licence logicielle', 'Appareil photo', 'Téléphone satellite','Radio VHF','Vidéo projecteur', 'Domino', 'Autre'];
// const ASSET_STATES = ['Très bonne état (TBE)', 'Bonne état (BE)', 'Mauvaise état (ME)', 'Très mauvaise état (TME)'];
const ASSET_CATEGORIES = ['Buildings', 'Communications & IT Equipement', 'Furniture and Fixtures', 'Machinery and Equipement', 'Vehicles'];
const REGIONS = ['Antananarivo', 'Ambovombe', 'Diégo', 'Majunga', 'Manakara', 'Ranomafana', 'Tamatave', 'Tuléar']
const ROLES      = ['Admin','ICT','Asset focal point','Utilisateur'];

/*************************************************************
 * WEB APP ENTRY POINT
 *************************************************************/
function doGet(e) {
  const tpl = HtmlService.createTemplateFromFile('Index');
  return tpl.evaluate()
    .setTitle(APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/*************************************************************
 * DATABASE BOOTSTRAP
 *************************************************************/
let _cachedSS = null;

function getSS() {
  if (_cachedSS) return _cachedSS;

  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SPREADSHEET_ID');
  let ss;
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (err) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create(APP_NAME + ' - Database');
    props.setProperty('SPREADSHEET_ID', ss.getId());
  }
  ensureSheets(ss);
  _cachedSS = ss; // AJOUT
  return ss;
}

function seedDefaultSettings(ss) {
  const sheet = ss.getSheetByName(SHEETS.SETTINGS);

  // "true"/"false" en booléens natifs) — elle ne doit JAMAIS empêcher une
  // lecture de fonctionner si la feuille est protégée pour l'utilisateur.
  try {
    sheet.getRange('B2:B1000').setNumberFormat('@');
  } catch (err) {
    Logger.log('seedDefaultSettings: formatage ignoré (feuille protégée) : ' + err);
  }

  if (sheet.getLastRow() < 2) {
    // MODIFIÉ : idem, protège l'amorçage initial des réglages par défaut
    try {
      sheet.appendRow(['ORG_NAME','UNFPA - Fonds des Nations Unies pour la Population']);
      sheet.appendRow(['SUPPORT_EMAIL', Session.getEffectiveUser().getEmail()]);
      sheet.appendRow(['NOTIFY_ON_CREATE','true']);
      sheet.appendRow(['NOTIFY_ON_ASSIGN','true']);
      sheet.appendRow(['NOTIFY_ON_SLA_BREACH','true']);
    } catch (err) {
      Logger.log('seedDefaultSettings: amorçage Settings ignoré (feuille protégée) : ' + err);
    }
  }

  

  const usersSheet = ss.getSheetByName(SHEETS.USERS);
  if (usersSheet.getLastRow() < 2) {
    try {
      const email = Session.getEffectiveUser().getEmail() || 'admin@unfpa.org';
      usersSheet.appendRow([genId('USR'), 'Administrateur ICT', email, 'ICT', 'ICT Manager', 'Siège', '', 'Admin', 'Actif', new Date()]);
    } catch (err) {
      Logger.log('seedDefaultSettings: amorçage Users ignoré (feuille protégée) : ' + err);
    }
  }
}

function ensureSheets(ss) {
  Object.keys(SHEETS).forEach(function(key) {
    const name = SHEETS[key];
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      try {
        sheet = ss.insertSheet(name);
        sheet.appendRow(HEADERS[key]);
        sheet.setFrozenRows(1);
        sheet.getRange(1,1,1,HEADERS[key].length).setFontWeight('bold').setBackground('#0b6cb3').setFontColor('#ffffff');
      } catch (err) {
        Logger.log('ensureSheets: création "' + name + '" impossible (protection ?) : ' + err); // AJOUT
      }
    }
  });
  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('Feuille 1');
  if (def && ss.getSheets().length > 1) { try { ss.deleteSheet(def); } catch(e){} }

  seedDefaultSettings(ss);
}

/*************************************************************
 * USER PERMISSIONS
 *************************************************************/
/*************************************************************
 * CONTRÔLE D'ACCÈS PAR RÔLE
 * -----------------------------------------------------------
 * Reflète UNIQUEMENT les restrictions déjà existantes côté frontend :
 * - MANAGE_USERS   : onglet "Utilisateurs" (Admin uniquement)
 * - MANAGE_LISTS   : éditeur de listes déroulantes (Admin uniquement)
 * - SWITCH_DASHBOARD_USER : bascule du dashboard vers un autre agent (Admin/ICT)
 * Rien d'autre n'était restreint auparavant ; rien d'autre ne l'est ici.
 *************************************************************/
const PERMISSIONS = {
  MANAGE_USERS: ['Admin'],
  MANAGE_LISTS: ['Admin'],
  SWITCH_DASHBOARD_USER: ['Admin', 'ICT'],
  ACCESS_TICKETS: ['Admin', 'ICT']
};

function currentUserHasPermission(perm) {
  const user = getCurrentUser();
  const allowed = PERMISSIONS[perm] || [];
  return allowed.indexOf(user.Role) !== -1;
}

function assertPermission(perm) {
  if (!currentUserHasPermission(perm)) {
    throw new Error('Accès refusé : action réservée à un rôle supérieur.');
  }
}

// Endpoint léger pour le frontend : évite un aller-retour juste pour savoir
// si l'utilisateur peut voir tel bouton/onglet.
function getMyPermissions() {
  const user = getCurrentUser();
  const result = {};
  Object.keys(PERMISSIONS).forEach(function(p) {
    result[p] = PERMISSIONS[p].indexOf(user.Role) !== -1;
  });
  return result;
}

function sanitizeAssignedToEmail(assignedToEmail) {
  if (!assignedToEmail) return assignedToEmail;
  const me = getCurrentUser();
  if (assignedToEmail === me.Email) return assignedToEmail; // consulter ses propres stats : toujours permis
  if (!currentUserHasPermission('SWITCH_DASHBOARD_USER')) return me.Email; // repli silencieux sur soi-même
  return assignedToEmail;
}


/*************************************************************
 * GENERIC SHEET HELPERS
 *************************************************************/
function sheetToObjects(sheetName) {
  const sheet = getSS().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).filter(function(r){ return r.join('') !== ''; }).map(function(row) {
    const obj = {};
    headers.forEach(function(h, i) {
      let v = row[i];
      if (v instanceof Date) v = v.toISOString();
      obj[h] = v;
    });
    return obj;
  });
}

function findRowIndexById(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1; // 1-based row number
  }
  return -1;
}

function genId(prefix) {
  return prefix + '-' + Utilities.formatDate(new Date(), TZ, 'yyMMdd') + '-' + Math.floor(1000 + Math.random()*9000);
}

function nowISO() { return new Date().toISOString(); }

function getCurrentUser() {
  const email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  const users = sheetToObjects(SHEETS.USERS);
  let user = users.find(function(u){ return u.Email === email; });
  if (!user) {
    user = { ID: genId('USR'), Nom: email.split('@')[0], Email: email, Departement:'', Poste:'', Bureau:'', Telephone:'', Role: users.length ? 'Utilisateur' : 'Admin', Statut:'Actif' };
  }
  return user;
}

function logAudit(action, cible, details) {
  const sheet = getSS().getSheetByName(SHEETS.AUDIT);
  const user = getCurrentUser();
  sheet.appendRow([genId('LOG'), new Date(), user.Nom + ' (' + user.Email + ')', action, cible, details || '']);
}

/*************************************************************
 * FILE UPLOAD - DRAG & DROP
 *************************************************************/
function uploadFileToDrive(fileObj) {
  try {
    if (!fileObj || !fileObj.base64Data || !fileObj.fileName) {
      throw new Error("Données de fichier invalides.");
    }
    const allowedExtensions = ['pdf', 'doc', 'docx', 'xls', 'xlsx'];
    const extension = fileObj.fileName.split('.').pop().toLowerCase();
    
    if (!allowedExtensions.includes(extension)) throw new Error("Format non autorisé côté serveur.");

    const decodedBytes = Utilities.base64Decode(fileObj.base64Data);
    const MAX_SIZE_BYTES = 10 * 1024 * 1024; // 10 Mo
    if (decodedBytes.length > MAX_SIZE_BYTES) throw new Error("Fichier trop volumineux (5 Mo max).");

    const blob = Utilities.newBlob(decodedBytes, fileObj.mimeType, fileObj.fileName);
    const folder = DriveApp.getFolderById(UPLOAD_FOLDER_ID);
    const file = folder.createFile(blob);

    file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);

    return { success: true, id: file.getId(), url: file.getUrl(), name: file.getName() };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/*************************************************************
 * TICKETS
 *************************************************************/
function getAllTickets() {
  assertPermission('ACCESS_TICKETS');
  const tickets = sheetToObjects(SHEETS.TICKETS);
  const comments = sheetToObjects(SHEETS.COMMENTS);
  tickets.forEach(function(t) {
    t.SLA_Statut = computeSlaStatus(t);
    t.NbCommentaires = comments.filter(function(c){ return c.TicketID === t.ID; }).length;
  });

  let alltickets = tickets.sort(function(a,b){ return new Date(b.DateCreation) - new Date(a.DateCreation); })
  // console.log(alltickets);
  return alltickets;
}

function computeSlaStatus(ticket) {
  if (['Résolu','Fermé'].indexOf(ticket.Statut) !== -1) return 'Respecté';
  if (!ticket.SLA_Echeance) return 'N/A';
  const due = new Date(ticket.SLA_Echeance);
  return (new Date() > due) ? 'Dépassé' : 'Dans les temps';
}

function getSlaHours(priority) {
  const raw = getSetting('SLA_HOURS_MAP');
  let map = SLA_HOURS;
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      if (parsed && typeof parsed === 'object') map = parsed;
    } catch (e) {}
  }
  return map[priority] || 24;
}

function createTicket(data) {
  const sheet = getSS().getSheetByName(SHEETS.TICKETS);
  const id = genId('TCK');
  const now = new Date();
  const slaHours = getSlaHours(data.Priorite);
  const slaDue = new Date(now.getTime() + slaHours * 3600 * 1000);
  const user = getCurrentUser();

  sheet.appendRow([
    id, data.Titre, data.Description, data.Categorie, data.Priorite, 'Nouveau',
    data.Demandeur || user.Nom, data.Email || user.Email, data.Departement || user.Departement,
    data.Bureau || user.Bureau, data.Assigne || '', now, now, '', slaDue, 'Dans les temps',
    data.PiecesJointes || ''
  ]);

  logAudit('Création ticket', id, data.Titre);
  notify('CREATE', { ID:id, Titre:data.Titre, Email:data.Email || user.Email, Priorite:data.Priorite });
  return { success:true, id:id };
}

function updateTicket(id, data) {
  const sheet = getSS().getSheetByName(SHEETS.TICKETS);
  const row = findRowIndexById(sheet, id);
  if (row === -1) return { success:false, error:'Ticket introuvable' };

  const headers = HEADERS.TICKETS;
  const current = sheet.getRange(row,1,1,headers.length).getValues()[0];
  const rowObj = {};
  headers.forEach(function(h,i){ rowObj[h]=current[i]; });

  Object.keys(data).forEach(function(k) {
    if (headers.indexOf(k) !== -1) rowObj[k] = data[k];
  });
  rowObj.DateMiseAJour = new Date();
  if (data.Statut === 'Résolu' || data.Statut === 'Fermé') {
    rowObj.DateResolution = new Date();
  }
  if (data.Priorite && data.Priorite !== current[headers.indexOf('Priorite')]) {
    const hrs = getSlaHours(data.Priorite);
    rowObj.SLA_Echeance = new Date(new Date(rowObj.DateCreation).getTime() + hrs*3600*1000);
  }

  const newRow = headers.map(function(h){ return rowObj[h]; });
  sheet.getRange(row,1,1,headers.length).setValues([newRow]);

  logAudit('Mise à jour ticket', id, JSON.stringify(data));
  if (data.Assigne) notify('ASSIGN', { ID:id, Titre:rowObj.Titre, AssigneEmail:data.Assigne });
  // data.NotifyEmail is a UI-only flag (not a sheet column) letting the technician
  // decide, via a front-end confirmation, whether the requester gets an email.
  if (data.Statut) notify('STATUS', { ID:id, Titre:rowObj.Titre, Email:rowObj.Email, Statut:data.Statut, NotifyEmail:data.NotifyEmail, NotifyCC:data.NotifyCC }); // MODIFIÉ : + NotifyCC
  return { success:true };
}

function deleteTicket(id) {
  const sheet = getSS().getSheetByName(SHEETS.TICKETS);
  const row = findRowIndexById(sheet, id);
  if (row === -1) return { success:false };
  sheet.deleteRow(row);
  logAudit('Suppression ticket', id, '');
  return { success:true };
}

/*************************************************************
 * COMMENTS
 *************************************************************/
function getTicketComments(ticketId) {
  return sheetToObjects(SHEETS.COMMENTS)
    .filter(function(c){ return c.TicketID === ticketId; })
    .sort(function(a,b){ return new Date(a.Date) - new Date(b.Date); });
}

function addComment(ticketId, message, interne) {
  const sheet = getSS().getSheetByName(SHEETS.COMMENTS);
  const user = getCurrentUser();
  const id = genId('CMT');
  sheet.appendRow([id, ticketId, user.Nom, message, new Date(), interne ? 'Oui':'Non']);
  logAudit('Commentaire ajouté', ticketId, message.substring(0,80));
  return { success:true, id:id };
}

/*************************************************************
 * NOTIFICATIONS (EMAIL)
 *************************************************************/
function getSetting(key) {
  const rows = sheetToObjects(SHEETS.SETTINGS);
  const found = rows.find(function(r){ return r.Cle === key; });
  return found ? found.Valeur : null;
}

// Google Sheets silently converts text like "true"/"false" typed (or written via
// setValue/appendRow) into a native boolean TRUE/FALSE. That means getSetting()
// can return either the JS boolean `true` or the string 'true' depending on how
// the cell was populated. This helper normalizes both cases so notification
// toggles never silently fail because of a strict string comparison.
function isSettingTrue(key) {
  const v = getSetting(key);
  if (typeof v === 'boolean') return v;
  if (v === null || v === undefined) return false;
  return String(v).trim().toLowerCase() === 'true';
}

function notify(type, payload) {
  try {
    if (type === 'CREATE' && isSettingTrue('NOTIFY_ON_CREATE')) {
      const support = getSetting('SUPPORT_EMAIL');
      if (support) {
        MailApp.sendEmail(support, '[UNFPA ICT] Nouveau ticket ' + payload.ID,
          'Un nouveau ticket a été créé.\n\nTitre: ' + payload.Titre + '\nPriorité: ' + payload.Priorite + '\n\nMerci de le traiter dans le dashboard.',
          {
            name: 'Unité ICT UNFPA Madagascar'
          }
        );
      }
    }
    if (type === 'ASSIGN' && isSettingTrue('NOTIFY_ON_ASSIGN') && payload.AssigneEmail) {
      MailApp.sendEmail(payload.AssigneEmail, '[UNFPA ICT] Ticket assigné ' + payload.ID,
        'Le ticket "' + payload.Titre + '" (' + payload.ID + ') vous a été assigné.',
        {
          name: 'Unité ICT UNFPA Madagascar'
        }
      );
    }
    if (type === 'STATUS' && payload.Email && payload.NotifyEmail !== false) {
      const options = { name: 'Unité ICT UNFPA Madagascar' };
      if (payload.NotifyCC) options.cc = payload.NotifyCC;
      MailApp.sendEmail(payload.Email, '[UNFPA ICT] Mise à jour ticket ' + payload.ID,
        'Le statut de votre ticket "' + payload.Titre + '" est maintenant : ' + payload.Statut,
        options
      );
    }
    if (type === 'LOAN' && payload.Email && payload.NotifyEmail !== false) {
      const dueLine = payload.DateRetourPrevue ? ('\nDate de retour prévue : ' + payload.DateRetourPrevue) : '';
      const notes = payload.Notes ? ('\n\nRemarque : ' + payload.Notes) : '';
      const fixedCC = 'razakalalaina@unfpa.org';
      const cc = payload.NotifyCC ? (fixedCC + ',' + payload.NotifyCC) : fixedCC;
      MailApp.sendEmail(payload.Email, '[UNFPA ICT] Actif confié : ' + payload.AssetLabel,
        'Bonjour ' + payload.Nom + ',\n\nL\'actif "' + payload.AssetLabel + '" vous a été confié.' + dueLine + notes +
        '\n\nMerci d\'en prendre soin et de le restituer à l\'équipe ICT le moment venu.',
        {
          name: 'Unité ICT UNFPA Madagascar',
          cc: cc
        }
      );
    }
  } catch (err) {
    Logger.log('Notify error: ' + err);
  }
}

/*************************************************************
 * SLA BREACH CHECK (trigger-friendly)
 *************************************************************/
function checkSlaBreaches() {
  const tickets = getAllTickets();
  const support = getSetting('SUPPORT_EMAIL');
  if (!isSettingTrue('NOTIFY_ON_SLA_BREACH') || !support) return;
  const breached = tickets.filter(function(t){
    return t.SLA_Statut === 'Dépassé' && ['Résolu','Fermé'].indexOf(t.Statut) === -1;
  });
  if (breached.length) {
    const list = breached.map(function(t){ return '- ' + t.ID + ' : ' + t.Titre + ' (' + t.Priorite + ')'; }).join('\n');
    MailApp.sendEmail(support, '[UNFPA ICT] ' + breached.length + ' ticket(s) en dépassement SLA', list,
      {
        name: 'Unité ICT UNFPA Madagascar'
      }
    );
  }
}

/*************************************************************
 * ASSETS / INVENTORY
 *************************************************************/
function getAllAssets() { return sheetToObjects(SHEETS.ASSETS); }

function _createAssetCore(data) {
  const sheet = getSS().getSheetByName(SHEETS.ASSETS);
  const id = genId('AST');
  sheet.appendRow([
    id, data.Type, data.Marque, data.Modele, data.NumeroSerie, data.TagInventaire || '',
    data.Statut || 'En service', data.UtilisePar || '', data.Departement || '', data.Bureau || '',
    data.DateAchat || '', data.FinGarantie || '', data.AssigneA || '', data.Cout || '', data.Notes || '', data.NomDocument || '', data.LienDocument || '', data.AssetNumber || '', data.Categorie || '', data.Etat || '', data.Book || '', data.Regions || ''
  ]);
  logAudit('Création actif', id, data.Type + ' ' + data.Modele);
  return { success:true, id:id };
}

function _updateAssetCore(id, data) {
  const sheet = getSS().getSheetByName(SHEETS.ASSETS);
  const row = findRowIndexById(sheet, id);
  if (row === -1) return { success:false };
  const headers = HEADERS.ASSETS;
  const current = sheet.getRange(row,1,1,headers.length).getValues()[0];
  const rowObj = {};
  headers.forEach(function(h,i){ rowObj[h]=current[i]; });
  Object.keys(data).forEach(function(k){ if (headers.indexOf(k) !== -1) rowObj[k]=data[k]; });
  sheet.getRange(row,1,1,headers.length).setValues([headers.map(function(h){ return rowObj[h]; })]);
  logAudit('Mise à jour actif', id, '');
  return { success:true };
}

function createAsset(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const all = getAllAssets();
    _assertUniqueFields(data, null, all); // MODIFIÉ : remplace les deux appels précédents
    if (_isExtra(data)) data.TagInventaire = _normTag(data.TagInventaire);
    const result = _createAssetCore(data);
    return result;
  } finally { lock.releaseLock(); }
}

function updateAsset(id, data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const all = getAllAssets();
    const current = all.find(function(a) { return a.ID === id; }) || {};
    const candidate = Object.assign({}, current, data);
    _assertUniqueFields(candidate, id, all);
    if (_isExtra(candidate) && data.TagInventaire !== undefined) data.TagInventaire = _normTag(data.TagInventaire);
    const result = _updateAssetCore(id, data);
    return result;
  } finally { lock.releaseLock(); }
}

function deleteAsset(id) {
  const sheet = getSS().getSheetByName(SHEETS.ASSETS);
  const row = findRowIndexById(sheet, id);
  if (row === -1) return { success:false };
  sheet.deleteRow(row);
  logAudit('Suppression actif', id, '');
  return { success:true };
}

/*************************************************************
 * RÈGLE TAG "EXTRA" : format XXX000 + unicité
 *************************************************************/
const EXTRA_ASSET_NUMBER = 'EXTRA';
const EXTRA_TAG_REGEX = /^[A-Z]{3}[0-9]{3}$/;

function _normTag(v) { return String(v === null || v === undefined ? '' : v).trim().toUpperCase(); }
function _isExtra(a) { return _normTag(a && a.AssetNumber) === EXTRA_ASSET_NUMBER; }

// Centralise tous les contrôles d'unicité/format avant écriture d'un actif.
function _assertUniqueFields(candidate, excludeId, allAssets) {
  // Tag EXTRA : format XXX000 + unicité
  if (_isExtra(candidate)) {
    const tag = _normTag(candidate.TagInventaire);
    if (!tag) throw new Error("Tag d'inventaire obligatoire pour un actif EXTRA (format XXX000).");
    if (!EXTRA_TAG_REGEX.test(tag)) throw new Error('Format de tag invalide : 3 lettres + 3 chiffres attendus (ex. TBL005).');
    const dupTag = allAssets.find(function(a) { return a.ID !== excludeId && _normTag(a.TagInventaire) === tag; });
    if (dupTag) throw new Error('Le tag ' + tag + ' est déjà utilisé (actif ' + dupTag.ID + ').');
  }

  // AssetNumber : unique, sauf s'il vaut "EXTRA" (valeur partagée par design)
  const num = String(candidate.AssetNumber || '').trim();
  if (num && num.toUpperCase() !== 'EXTRA') {
    const dupNum = allAssets.find(function(a) { return a.ID !== excludeId && String(a.AssetNumber || '').trim().toUpperCase() === num.toUpperCase(); });
    if (dupNum) throw new Error('Le numéro d\'actif ' + num + ' est déjà utilisé (actif ' + dupNum.ID + ').');
  }

  // NumeroSerie : unique dès qu'il est renseigné
  const serial = String(candidate.NumeroSerie || '').trim();
  if (serial) {
    const dupSerial = allAssets.find(function(a) { return a.ID !== excludeId && String(a.NumeroSerie || '').trim().toUpperCase() === serial.toUpperCase(); });
    if (dupSerial) throw new Error('Le numéro de série ' + serial + ' est déjà utilisé (actif ' + dupSerial.ID + ').');
  }
}

// Dernier numéro et prochain numéro disponible par préfixe.
function _buildExtraTagSummary(assets) {
  const byPrefix = {};
  const invalid = [];
  assets.forEach(function(a) {
    if (!_isExtra(a)) return;
    const tag = _normTag(a.TagInventaire);
    if (!EXTRA_TAG_REGEX.test(tag)) { invalid.push({ ID: a.ID, tag: a.TagInventaire || '' }); return; }
    const prefix = tag.substring(0, 3);
    const num = parseInt(tag.substring(3), 10);
    if (!byPrefix[prefix]) byPrefix[prefix] = { prefix: prefix, count: 0, last: 0, tags: [], marques: [] };
    byPrefix[prefix].count++;
    byPrefix[prefix].tags.push(tag);
    if (num > byPrefix[prefix].last) byPrefix[prefix].last = num;

    const marque = String(a.Marque || '').trim();
    if (marque && !byPrefix[prefix].marques.some(function(m) { return m.toLowerCase() === marque.toLowerCase(); })) {
      byPrefix[prefix].marques.push(marque);
    }
  });
  const prefixes = Object.keys(byPrefix).sort().map(function(p) {
    const e = byPrefix[p];
    e.tags.sort();
    e.marques.sort();
    e.lastTag = p + ('00' + e.last).slice(-3);
    e.next = e.last < 999 ? p + ('00' + (e.last + 1)).slice(-3) : null;
    return e;
  });
  return { prefixes: prefixes, invalid: invalid };
}

/*************************************************************
 * ASSETS / DASHBOARD ANALYTIQUE
 *************************************************************/
// Agrégation générique par valeur de colonne — ne suppose aucune liste
// figée, fonctionne même si de nouvelles valeurs sont ajoutées plus tard
// (comportement identique à renderAgentAllChart/renderDeptChart côté client).
function _groupAssetsBy(assets, field) {
  const counts = {};
  assets.forEach(function(a) {
    const raw = a[field];
    const key = (raw === undefined || raw === null || String(raw).trim() === '') ? 'Non renseigné' : String(raw).trim();
    counts[key] = (counts[key] || 0) + 1;
  });
  return counts;
}

function getAssetDashboardStats() {
  const assets = getAllAssets();

  const byType     = _groupAssetsBy(assets, 'Type');
  const byStatus    = _groupAssetsBy(assets, 'Statut');
  const byBook      = _groupAssetsBy(assets, 'Book');
  const byRegion    = _groupAssetsBy(assets, 'Regions');
  const byCategorie = _groupAssetsBy(assets, 'Categorie'); // disponible si besoin futur
  const byEtat      = _groupAssetsBy(assets, 'Etat');       // disponible si besoin futur

  const byRegionBook = {};
  assets.forEach(function(a) {
    const region = (a.Regions === undefined || a.Regions === null || String(a.Regions).trim() === '') ? 'Non renseigné' : String(a.Regions).trim();
    const book = (a.Book === undefined || a.Book === null || String(a.Book).trim() === '') ? 'Non renseigné' : String(a.Book).trim();
    if (!byRegionBook[region]) byRegionBook[region] = {};
    byRegionBook[region][book] = (byRegionBook[region][book] || 0) + 1;
  });

  let totalCost = 0, capitalizedCost = 0, expensedCost = 0;
  let capitalizedCount = 0, expensedCount = 0;
  assets.forEach(function(a) {
    const cost = parseFloat(a.Cout) || 0;
    totalCost += cost;
    if (a.Book === 'Capitalized') { capitalizedCost += cost; capitalizedCount++; }
    else if (a.Book === 'Expensed') { expensedCost += cost; expensedCount++; }
  });

  return {
    totalAssets: assets.length,
    totalCost: totalCost,
    capitalizedCost: capitalizedCost,
    expensedCost: expensedCost,
    capitalizedCount: capitalizedCount,
    expensedCount: expensedCount,
    byType: byType,
    byStatus: byStatus,
    byBook: byBook,
    byRegion: byRegion,
    byRegionBook: byRegionBook,
    extraTags: _buildExtraTagSummary(assets),
    byCategorie: byCategorie,
    byEtat: byEtat
  };
}

/*************************************************************
 * ASSET TRANSFERS / EMPRUNTS
 * -----------------------------------------------------------
 * Tracks the loan/transfer history of an asset over time (who
 * had it and when). This is genuinely new data — the Assets
 * sheet only ever stores the CURRENT holder, not the history —
 * so a dedicated sheet is warranted. Everything else is reused:
 *   - the asset is picked from the existing Assets list,
 *   - the borrower is picked from the existing Users directory
 *     (same data source as the "Demandeur"/"Assigné à" pickers),
 *   - every event is logged through the SAME shared logAudit()
 *     / AuditLog sheet already used for tickets, assets, users
 *     and KB articles — no parallel audit system.
 *************************************************************/
function getAllTransfers() {
  return sheetToObjects(SHEETS.TRANSFERS).sort(function(a,b){ return new Date(b.DateEmprunt) - new Date(a.DateEmprunt); });
}

function createTransfer(data) {
  if (!data.AssetID || !data.Emprunteur) return { success:false, error:'Actif et emprunteur requis' };

  const assetsSheet = getSS().getSheetByName(SHEETS.ASSETS);
  const assetRow = findRowIndexById(assetsSheet, data.AssetID);
  if (assetRow === -1) return { success:false, error:'Actif introuvable' };
  const aHeaders = HEADERS.ASSETS;
  const aCurrent = assetsSheet.getRange(assetRow,1,1,aHeaders.length).getValues()[0];
  const assetObj = {}; aHeaders.forEach(function(h,i){ assetObj[h]=aCurrent[i]; });
  const assetLabel = [assetObj.Marque, assetObj.Modele].filter(Boolean).join(' ') + (assetObj.NumeroSerie ? ' (' + assetObj.NumeroSerie + ')' : '');
  const previousHolder = assetObj.UtilisePar || '';

  const sheet = getSS().getSheetByName(SHEETS.TRANSFERS);
  const id = genId('TRF');
  const actor = getCurrentUser();
  sheet.appendRow([
    id, data.AssetID, assetLabel, data.Emprunteur, data.EmprunteurEmail || '',
    data.Departement || '', data.Bureau || '', new Date(), data.DateRetourPrevue || '',
    '', 'En cours', actor.Nom, data.Notes || ''
  ]);

  // MODIFIÉ : assetObj.AssigneA → assetObj.UtilisePar
  assetObj.UtilisePar = data.Emprunteur;
  if (data.Departement) assetObj.Departement = data.Departement;
  if (data.Bureau) assetObj.Bureau = data.Bureau;
  assetObj.Statut = 'En service';
  assetsSheet.getRange(assetRow,1,1,aHeaders.length).setValues([aHeaders.map(function(h){ return assetObj[h]; })]);

  const who = previousHolder && previousHolder !== data.Emprunteur ? (previousHolder + ' \u2192 ' + data.Emprunteur) : data.Emprunteur;
  logAudit('Emprunt actif', assetLabel + ' (' + data.AssetID + ')', 'Confié à ' + who);
  notify('LOAN', { Email: data.EmprunteurEmail, Nom: data.Emprunteur, AssetLabel: assetLabel, DateRetourPrevue: data.DateRetourPrevue, Notes: data.Notes, NotifyEmail: data.NotifyEmail, NotifyCC: data.NotifyCC });
  return { success:true, id:id };
}

// Marks an active loan as returned and frees up the asset (Statut -> "En stock").

function returnTransfer(id, notes) {
  const sheet = getSS().getSheetByName(SHEETS.TRANSFERS);
  const row = findRowIndexById(sheet, id);
  if (row === -1) return { success:false, error:'Emprunt introuvable' };
  const headers = HEADERS.TRANSFERS;
  const current = sheet.getRange(row,1,1,headers.length).getValues()[0];
  const obj = {}; headers.forEach(function(h,i){ obj[h]=current[i]; });
  if (obj.Statut === 'Retourné') return { success:true };

  obj.DateRetourEffective = new Date();
  obj.Statut = 'Retourné';
  if (notes) obj.Notes = (obj.Notes ? obj.Notes + ' | ' : '') + notes;
  sheet.getRange(row,1,1,headers.length).setValues([headers.map(function(h){ return obj[h]; })]);

  const assetsSheet = getSS().getSheetByName(SHEETS.ASSETS);
  const assetRow = findRowIndexById(assetsSheet, obj.AssetID);
  if (assetRow !== -1) {
    const aHeaders = HEADERS.ASSETS;
    const aCurrent = assetsSheet.getRange(assetRow,1,1,aHeaders.length).getValues()[0];
    const aObj = {}; aHeaders.forEach(function(h,i){ aObj[h]=aCurrent[i]; });
    // MODIFIÉ : aObj.AssigneA → aObj.UtilisePar (des deux côtés de la comparaison et de l'affectation)
    if (aObj.UtilisePar === obj.Emprunteur) {
      aObj.UtilisePar = '';
      aObj.Statut = 'En stock';
      assetsSheet.getRange(assetRow,1,1,aHeaders.length).setValues([aHeaders.map(function(h){ return aObj[h]; })]);
    }
  }

  logAudit('Retour actif', obj.AssetLabel + ' (' + obj.AssetID + ')', 'Rendu par ' + obj.Emprunteur);
  return { success:true };
}

/*************************************************************
 * ASSETS / EXPORT
 *************************************************************/

const ASSET_EXPORT_COLUMNS = {
  ID:            { label: 'ID',            get: function(a){ return a.ID || ''; } },
  AssetNumber:   { label: 'Asset number',            get: function(a){ return a.AssetNumber || ''; } },
  Categorie:     { label: 'Major category',get: function(a){ return a.Categorie || ''; } },
  Type:          { label: 'Type',          get: function(a){ return a.Type || ''; } },
  MarqueModele:  { label: 'Marque/Modèle', get: function(a){ return (a.Marque || '') + ' ' + (a.Modele || ''); } },
  NumeroSerie:   { label: 'N° Série',      get: function(a){ return a.NumeroSerie || ''; } },
  UtilisePar:    { label: 'Utilisé par',   get: function(a){ return a.UtilisePar || ''; } },
  Bureau:        { label: 'Bureau',        get: function(a){ return a.Bureau || ''; } },
  Regions:       { label: 'Régions',        get: function(a){ return a.Regions || ''; } },
  Departement:   { label: 'Département',  get: function(a){ return a.Departement || ''; } },
  AssigneA:      { label: 'Assigné à',     get: function(a){ return a.AssigneA || ''; } },
  DateAchat:     { label: "Date d'achat",  get: function(a){ return a.DateAchat || ''; } },
  FinGarantie:   { label: 'Fin de garantie', get: function(a){ return a.FinGarantie || ''; } },
  TagInventaire: { label: "Tag number", get: function(a){ return a.TagInventaire || ''; } },
  Cout:          { label: 'Coût',          get: function(a){ return a.Cout || ''; } },
  Etat:          { label: 'Etat',          get: function(a){ return a.Etat || ''; } },
  Statut:        { label: 'Statut',        get: function(a){ return a.Statut || ''; } },
  Book:          { label: 'Book',        get: function(a){ return a.Book || ''; } },
  Notes:         { label: 'Notes',         get: function(a){ return a.Notes || ''; } }
};

/**
 * Exporte les actifs au format Excel
 * @param {Array} assets Liste des actifs à exporter
 * @return {String} URL de téléchargement du fichier Excel
 */

function getAssetExportColumns() {
  return Object.keys(ASSET_EXPORT_COLUMNS).map(function(key) {
    return { key: key, label: ASSET_EXPORT_COLUMNS[key].label };
  });
}

/**
 * Exporte les actifs au format Excel
 * @param {Array} assets Liste des actifs à exporter
 * @param {Array} columns Clés de colonnes sélectionnées (cf. ASSET_EXPORT_COLUMNS).
 *                        Si omis/vide, utilise un jeu de colonnes par défaut (comportement historique).
 * @return {String} URL de téléchargement du fichier Excel
 */
function exportAssetsToExcel(assets, columns) {
  if (!assets || !Array.isArray(assets)) {
    throw new Error('Aucune donnée à exporter.');
  }

  // MODIFIÉ : colonnes dynamiques au lieu d'un tableau d'en-têtes en dur.
  // Repli sur le jeu de colonnes historique si rien n'est fourni.
  const selectedKeys = (columns && columns.length)
    ? columns.filter(function(k){ return ASSET_EXPORT_COLUMNS[k]; })
    : ['ID','Type','MarqueModele','NumeroSerie','Statut','AssigneA','FinGarantie','Notes'];

  if (!selectedKeys.length) {
    throw new Error('Aucune colonne valide sélectionnée.');
  }

  const spreadsheet = SpreadsheetApp.create(
    'Export_Assets_' + Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyyMMdd_HHmmss'
    )
  );

  const sheet = spreadsheet.getSheets()[0];
  sheet.setName('Assets');

  const headers = selectedKeys.map(function(k){ return ASSET_EXPORT_COLUMNS[k].label; });

  sheet
    .getRange(1, 1, 1, headers.length)
    .setValues([headers]);

  const rows = assets.map(function(a) {
    return selectedKeys.map(function(k){ return ASSET_EXPORT_COLUMNS[k].get(a); });
  });

  if (rows.length > 0) {
    sheet
      .getRange(2, 1, rows.length, headers.length)
      .setValues(rows);
  }

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);

  SpreadsheetApp.flush();

  const exportUrl = `https://docs.google.com/spreadsheets/d/${spreadsheet.getId()}/export?format=xlsx`;
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    DriveApp.getFileById(spreadsheet.getId()).setTrashed(true);
    throw new Error('Impossible de générer le fichier Excel.');
  }

  const fileName = 'ICTAssets_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.xlsx';
  const excelFile = DriveApp.createFile(response.getBlob().setName(fileName));
  DriveApp.getFileById(spreadsheet.getId()).setTrashed(true);

  return excelFile.getDownloadUrl();
}

/*************************************************************
 * USERS / STAFF DIRECTORY
 *************************************************************/

function _getAllUsersRaw() {
  return sheetToObjects(SHEETS.USERS);
}

function getUserDirectoryLite() {
  return _getAllUsersRaw().map(function(u) {
    return { ID: u.ID, Nom: u.Nom, Email: u.Email, Departement: u.Departement, Bureau: u.Bureau, Role: u.Role, Regions: u.Regions };
  });
}

function getAllUsers() {
  assertPermission('MANAGE_USERS');
  return _getAllUsersRaw();
}

function createUser(data) {
  assertPermission('MANAGE_USERS'); // AJOUT
  const sheet = getSS().getSheetByName(SHEETS.USERS);
  const id = genId('USR');
  sheet.appendRow([id, data.Nom, data.Email, data.Departement || '', data.Poste || '', data.Bureau || '',
    data.Telephone || '', data.Role || 'Utilisateur', data.Statut || 'Actif', new Date(), data.Regions]);
  logAudit('Création utilisateur', id, data.Nom);
  return { success:true, id:id };
}

function updateUser(id, data) {
  assertPermission('MANAGE_USERS'); // AJOUT
  const sheet = getSS().getSheetByName(SHEETS.USERS);
  const row = findRowIndexById(sheet, id);
  if (row === -1) return { success:false };
  const headers = HEADERS.USERS;
  const current = sheet.getRange(row,1,1,headers.length).getValues()[0];
  const rowObj = {};
  headers.forEach(function(h,i){ rowObj[h]=current[i]; });
  Object.keys(data).forEach(function(k){ if (headers.indexOf(k) !== -1) rowObj[k]=data[k]; });
  sheet.getRange(row,1,1,headers.length).setValues([headers.map(function(h){ return rowObj[h]; })]);
  logAudit('Mise à jour utilisateur', id, '');
  return { success:true };
}

function deleteUser(id) {
  assertPermission('MANAGE_USERS'); // AJOUT
  const sheet = getSS().getSheetByName(SHEETS.USERS);
  const row = findRowIndexById(sheet, id);
  if (row === -1) return { success:false };
  sheet.deleteRow(row);
  logAudit('Suppression utilisateur', id, '');
  return { success:true };
}

function getAssignableUsers() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const idxNom = headers.indexOf('Nom');
  const idxRole = headers.indexOf('Role');

  const eligibleRoles = ['Admin', 'ICT'];

  return data.slice(1)
    .filter(row => eligibleRoles.includes(row[idxRole]))
    .map(row => ({ name: row[idxNom], role: row[idxRole] }));
}

/*************************************************************
 * KNOWLEDGE BASE
 *************************************************************/
function getAllArticles() { return sheetToObjects(SHEETS.KB); }

function createArticle(data) {
  const sheet = getSS().getSheetByName(SHEETS.KB);
  const id = genId('KB');
  const user = getCurrentUser();
  sheet.appendRow([id, data.Titre, data.Categorie, data.Contenu, user.Nom, new Date(), new Date(), 0, data.TagsMots || '']);
  logAudit('Création article KB', id, data.Titre);
  return { success:true, id:id };
}

function updateArticle(id, data) {
  const sheet = getSS().getSheetByName(SHEETS.KB);
  const row = findRowIndexById(sheet, id);
  if (row === -1) return { success:false };
  const headers = HEADERS.KB;
  const current = sheet.getRange(row,1,1,headers.length).getValues()[0];
  const rowObj = {};
  headers.forEach(function(h,i){ rowObj[h]=current[i]; });
  Object.keys(data).forEach(function(k){ if (headers.indexOf(k) !== -1) rowObj[k]=data[k]; });
  rowObj.DateMiseAJour = new Date();
  sheet.getRange(row,1,1,headers.length).setValues([headers.map(function(h){ return rowObj[h]; })]);
  return { success:true };
}

function deleteArticle(id) {
  const sheet = getSS().getSheetByName(SHEETS.KB);
  const row = findRowIndexById(sheet, id);
  if (row === -1) return { success:false };
  sheet.deleteRow(row);
  return { success:true };
}

function incrementArticleViews(id) {
  const sheet = getSS().getSheetByName(SHEETS.KB);
  const row = findRowIndexById(sheet, id);
  if (row === -1) return;
  const col = HEADERS.KB.indexOf('Vues') + 1;
  const cur = sheet.getRange(row, col).getValue() || 0;
  sheet.getRange(row, col).setValue(cur + 1);
}

/*************************************************************
 * AUDIT LOG (read)
 *************************************************************/
function getAuditLog() {
  return sheetToObjects(SHEETS.AUDIT).sort(function(a,b){ return new Date(b.Date) - new Date(a.Date); }).slice(0,300);
}

/*************************************************************
 * SETTINGS
 *************************************************************/
function getAllSettings() { return sheetToObjects(SHEETS.SETTINGS); }

function updateSetting(key, value) {
  const sheet = getSS().getSheetByName(SHEETS.SETTINGS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      const cell = sheet.getRange(i+1, 2);
      cell.setNumberFormat('@'); // keep as plain text, prevent Sheets auto-boolean conversion
      cell.setValue(value);
      return { success:true };
    }
  }
  sheet.appendRow([key, value]);
  return { success:true };
}

/*************************************************************
 * DASHBOARD / ANALYTICS
 *************************************************************/
function getDashboardStats(assignedToEmail) {
  assertPermission('ACCESS_TICKETS');
  assignedToEmail = sanitizeAssignedToEmail(assignedToEmail);
  const tickets = getAllTickets();
  const assets  = getAllAssets();
  const users   = _getAllUsersRaw();

  const openStatuses = ['Nouveau','Ouvert','En cours','En attente'];
  const today = new Date(); today.setHours(0,0,0,0);

  const scopedTickets = assignedToEmail
    ? tickets.filter(function(t){ return t.Assigne === assignedToEmail; })
    : tickets;

  const byStatus = {};
  getEditableList('STATUSES').forEach(function(s){ byStatus[s] = 0; });
  scopedTickets.forEach(function(t){ if (byStatus[t.Statut] !== undefined) byStatus[t.Statut]++; });

  const byPriority = {};
  getEditableList('PRIORITIES').forEach(function(p){ byPriority[p] = 0; });
  scopedTickets.forEach(function(t){ if (byPriority[t.Priorite] !== undefined) byPriority[t.Priorite]++; });

  const byCategory = {};
  getEditableList('CATEGORIES').forEach(function(c){ byCategory[c] = 0; });
  scopedTickets.forEach(function(t){ if (byCategory[t.Categorie] !== undefined) byCategory[t.Categorie]++; });


  const agents = {};
  tickets.forEach(function(t){
    if (t.Assigne && openStatuses.indexOf(t.Statut) !== -1) {
      agents[t.Assigne] = (agents[t.Assigne]||0) + 1;
    }
  });

  return {
    totalTickets: scopedTickets.length,
    openTickets: scopedTickets.filter(function(t){ return openStatuses.indexOf(t.Statut) !== -1; }).length,
    resolvedThisMonth: scopedTickets.filter(function(t){
      if (!t.DateResolution) return false;
      const d = new Date(t.DateResolution);
      const now = new Date();
      return d.getFullYear() === now.getFullYear() && d.getMonth() === now.getMonth();
    }).length,
    resolvedTotal: scopedTickets.filter(function(t){ return t.Statut === 'Résolu'; }).length,
    slaBreaches: scopedTickets.filter(function(t){ return t.SLA_Statut === 'Dépassé'; }).length,
    totalAssets: assets.length,
    totalUsers: users.length,
    unassigned: tickets.filter(function(t){ return !t.Assigne && openStatuses.indexOf(t.Statut) !== -1; }).length,
    byStatus: byStatus,
    byPriority: byPriority,
    byCategory: byCategory,
    agentWorkload: agents
  };
}

/**
 * Endpoint dédié au graphique de tendance, appelable indépendamment du reste
 * des stats du dashboard (évite de recalculer byStatus/byPriority/byCategory
 * à chaque changement de période).
 */
function getTicketTrend(assignedToEmail, trendPeriod) {
  assertPermission('ACCESS_TICKETS');
  assignedToEmail = sanitizeAssignedToEmail(assignedToEmail);
  const tickets = getAllTickets();
  const scopedTickets = assignedToEmail
    ? tickets.filter(function(t){ return t.Assigne === assignedToEmail; })
    : tickets;

  const today = new Date(); today.setHours(0,0,0,0);
  return buildTrend(scopedTickets, trendPeriod || 'month', today);
}

/**
 * Calcule la série "trend" (tickets créés dans le temps) selon 3 modes :
 * - '14d'   : un point par jour, sur les 14 derniers jours (labels 'dd/MM')
 * - 'month' : un point par jour, sur le mois calendaire en cours (labels 'dd') — comportement historique
 * - 'max'   : un point par jour, du 17 août (année en cours) à aujourd'hui (labels 'dd/MM')
 */
function buildTrend(tickets, period, today) {
  const countsByDate = {};
  tickets.forEach(function(t) {
    if (!t.DateCreation) return;
    const key = Utilities.formatDate(new Date(t.DateCreation), TZ, 'dd/MM/yyyy');
    countsByDate[key] = (countsByDate[key] || 0) + 1;
  });

  if (period === '14d') {
    const trend = [];
    for (let i = 13; i >= 0; i--) {
      const d = new Date(today);
      d.setDate(d.getDate() - i);
      const fullKey = Utilities.formatDate(d, TZ, 'dd/MM/yyyy');
      const label = Utilities.formatDate(d, TZ, 'dd/MM');
      trend.push({ day: label, count: countsByDate[fullKey] || 0 });
    }
    return { trend: trend, trendMonthLabel: '14 derniers jours' };
  }

  if (period === 'max') {
    const startDate = new Date(today.getFullYear(), 7, 17); // mois 7 = août (0-indexé)
    // Si "aujourd'hui" est avant le 17 août (ex: en début d'année), on part du
    // 17 août de l'année précédente pour garder une plage cohérente.
    if (startDate > today) startDate.setFullYear(startDate.getFullYear() - 1);

    const trend = [];
    const cursor = new Date(startDate);
    while (cursor <= today) {
      const fullKey = Utilities.formatDate(cursor, TZ, 'dd/MM/yyyy');
      const label = Utilities.formatDate(cursor, TZ, 'dd/MM');
      trend.push({ day: label, count: countsByDate[fullKey] || 0 });
      cursor.setDate(cursor.getDate() + 1);
    }
    return { trend: trend, trendMonthLabel: 'Tous les temps' };
  }

  // Mode par défaut : 'month' — comportement historique, inchangé
  const daysInMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();
  const trend = [];
  for (let d = 1; d <= daysInMonth; d++) {
    const dayDate = new Date(today.getFullYear(), today.getMonth(), d);
    const fullKey = Utilities.formatDate(dayDate, TZ, 'dd/MM/yyyy');
    const dayLabel = Utilities.formatDate(dayDate, TZ, 'dd');
    trend.push({ day: dayLabel, count: countsByDate[fullKey] || 0 });
  }
  const monthLabel = Utilities.formatDate(today, TZ, 'MMMM yyyy');
  const trendMonthLabel = monthLabel.charAt(0).toUpperCase() + monthLabel.slice(1);
  return { trend: trend, trendMonthLabel: trendMonthLabel };
}

/*************************************************************
 * META (constants for frontend)
 *************************************************************/
function getMeta() {
  return {
    priorities: getEditableList('PRIORITIES'),
    statuses: getEditableList('STATUSES'),
    categories: getEditableList('CATEGORIES'),
    assetTypes: getEditableList('ASSET_TYPES'),
    assetCategories: getEditableList('ASSET_CATEGORIES'),
    regions: getEditableList('REGIONS'),
    roles: getEditableList('ROLES'),
    currentUser: getCurrentUser(),
    orgName: getSetting('ORG_NAME') || 'UNFPA'
  };
}

/*************************************************************
 * LISTES DÉROULANTES ÉDITABLES
 * -----------------------------------------------------------
 * CATEGORIES, ASSET_TYPES, PRIORITIES, STATUSES, ROLES restent
 * déclarées en haut du fichier comme valeurs PAR DÉFAUT. La valeur
 * réellement utilisée par l'application est stockée en JSON dans la
 * feuille Settings et peut être modifiée depuis l'interface (Admin
 * uniquement) sans toucher au code.
 *************************************************************/
const EDITABLE_LISTS = {
  CATEGORIES:  { settingKey: 'LIST_CATEGORIES',  defaults: CATEGORIES,  label: 'Catégories de tickets' },
  ASSET_TYPES: { settingKey: 'LIST_ASSET_TYPES', defaults: ASSET_TYPES, label: "Types d'actifs" },
  ASSET_CATEGORIES: { settingKey: 'LIST_ASSET_CATEGORIES', defaults: ASSET_CATEGORIES, label: "Catégories d'actifs" },
  REGIONS: { settingKey: 'LIST_REGIONS', defaults: REGIONS, label: "Régions" },
  PRIORITIES:  { settingKey: 'LIST_PRIORITIES',  defaults: PRIORITIES,  label: 'Priorités' },
  STATUSES:    { settingKey: 'LIST_STATUSES',    defaults: STATUSES,    label: 'Statuts de ticket' },
  ROLES:       { settingKey: 'LIST_ROLES',       defaults: ROLES,       label: 'Rôles utilisateur' }
};

function getEditableList(listKey) {
  const def = EDITABLE_LISTS[listKey];
  if (!def) return [];
  const raw = getSetting(def.settingKey);
  if (!raw) return def.defaults.slice();
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) && parsed.length ? parsed : def.defaults.slice();
  } catch (e) {
    return def.defaults.slice();
  }
}

// Renvoie toutes les listes éditables + leur libellé, pour l'écran Paramètres.
function getAllEditableLists() {
  assertPermission('MANAGE_LISTS');
  const result = {};
  Object.keys(EDITABLE_LISTS).forEach(function(key) {
    result[key] = { label: EDITABLE_LISTS[key].label, items: getEditableList(key) };
  });
  return result;
}

// Sauvegarde une liste. Réservé aux Admin. Rejette les listes vides ou dupliquées.
function updateEditableList(listKey, items) {
  assertPermission('MANAGE_LISTS');
  const currentUser = getCurrentUser();
  if (currentUser.Role !== 'Admin') {
    return { success:false, error:'Seul un administrateur peut modifier les listes.' };
  }

  const def = EDITABLE_LISTS[listKey];
  if (!def) return { success:false, error:'Liste inconnue' };
  if (!Array.isArray(items) || items.length === 0) {
    return { success:false, error:'La liste ne peut pas être vide' };
  }

  const cleaned = [];
  items.forEach(function(v) {
    const s = String(v || '').trim();
    if (s && cleaned.indexOf(s) === -1) cleaned.push(s);
  });
  if (!cleaned.length) return { success:false, error:'Aucune valeur valide' };

  updateSetting(def.settingKey, JSON.stringify(cleaned));
  logAudit('Modification liste', def.label, JSON.stringify(cleaned));
  return { success:true, items: cleaned };
}

/*************************************************************
 * GLOBAL SEARCH
 *************************************************************/
function globalSearch(query) {
  query = (query || '').toLowerCase();
  if (!query) return { tickets:[], assets:[], users:[], articles:[] };
  const tickets = sheetToObjects(SHEETS.TICKETS).filter(function(t){
    return (t.Titre + t.Description + t.ID).toLowerCase().indexOf(query) !== -1;
  }).slice(0,10);
  const assets = sheetToObjects(SHEETS.ASSETS).filter(function(a){
    return (a.Marque + a.Modele + a.NumeroSerie + a.ID).toLowerCase().indexOf(query) !== -1;
  }).slice(0,10);
  const users = sheetToObjects(SHEETS.USERS).filter(function(u){
    return (u.Nom + u.Email).toLowerCase().indexOf(query) !== -1;
  }).slice(0,10);
  const articles = sheetToObjects(SHEETS.KB).filter(function(a){
    return (a.Titre + a.Contenu + a.TagsMots).toLowerCase().indexOf(query) !== -1;
  }).slice(0,10);
  return { tickets:tickets, assets:assets, users:users, articles:articles };
}

/*************************************************************
 * OPTIONAL: create a time-driven trigger for SLA checks
 * Run this once manually from the Apps Script editor.
 *************************************************************/
function setupSlaTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'checkSlaBreaches') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('checkSlaBreaches').timeBased().everyHours(1).create();
}