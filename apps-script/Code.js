// Galeria ślubna – backend na Google Apps Script.
// Przechowuje zdjęcia w Google Drive właściciela skryptu w trzech folderach.
// Goście mogą wysyłać TYLKO do folderu "Od gości". Reszta jest tylko do odczytu.

var CATS = { gosc: 'Od gości', foto: 'Od fotografa', film: 'Film' };
var MAX_BYTES = 40 * 1024 * 1024; // limit pojedynczego pliku (Apps Script przyjmuje ~50 MB POST)
var CACHE_SEC = 45;

// Uruchom RAZ ręcznie w edytorze (Run > setup), żeby nadać uprawnienia i utworzyć foldery.
function setup() {
  var props = PropertiesService.getScriptProperties();
  var rootId = props.getProperty('ROOT_ID');
  var root = rootId ? DriveApp.getFolderById(rootId) : DriveApp.createFolder('Ślub Kacper i Zuzanna – zdjęcia');
  props.setProperty('ROOT_ID', root.getId());
  root.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  Object.keys(CATS).forEach(function (k) {
    var id = props.getProperty('CAT_' + k);
    var f = id ? DriveApp.getFolderById(id) : root.createFolder(CATS[k]);
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    props.setProperty('CAT_' + k, f.getId());
  });
  Logger.log('Folder główny: ' + root.getUrl());
  return root.getUrl();
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function folderFor_(cat) {
  var id = PropertiesService.getScriptProperties().getProperty('CAT_' + cat);
  if (!id) throw new Error('Brak folderu dla kategorii ' + cat + ' – uruchom setup()');
  return DriveApp.getFolderById(id);
}

function listFiles_(cat) {
  var folder = folderFor_(cat);
  var it = folder.getFiles();
  var files = [];
  while (it.hasNext()) {
    var f = it.next();
    var mime = f.getMimeType() || '';
    if (!/^image\/|^video\//.test(mime)) continue;
    files.push({
      id: f.getId(),
      name: f.getName(),
      mime: mime,
      created: f.getDateCreated().getTime(),
      by: f.getDescription() || '',
      size: f.getSize()
    });
  }
  files.sort(function (a, b) { return b.created - a.created; });
  return { ok: true, cat: cat, folderUrl: folder.getUrl(), files: files };
}

function doGet(e) {
  var cat = (e && e.parameter && e.parameter.cat) || '';
  if (!CATS[cat]) return json_({ ok: false, error: 'Nieznana kategoria' });
  var cache = CacheService.getScriptCache();
  var hit = cache.get('list_' + cat);
  if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
  var out = JSON.stringify(listFiles_(cat));
  if (out.length < 90000) cache.put('list_' + cat, out, CACHE_SEC);
  return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.cat !== 'gosc') return json_({ ok: false, error: 'Goście mogą dodawać tylko do "Od gości"' });
    var mime = String(body.mime || '');
    if (!/^image\/|^video\//.test(mime)) return json_({ ok: false, error: 'Tylko zdjęcia i filmy' });
    var bytes = Utilities.base64Decode(body.data || '');
    if (!bytes.length) return json_({ ok: false, error: 'Pusty plik' });
    if (bytes.length > MAX_BYTES) return json_({ ok: false, error: 'Plik za duży (max 40 MB)' });
    var name = String(body.name || 'zdjecie').replace(/[\\\/:*?"<>|]/g, '_').slice(0, 120);
    var file = folderFor_('gosc').createFile(Utilities.newBlob(bytes, mime, name));
    if (body.by) file.setDescription(String(body.by).slice(0, 80));
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    CacheService.getScriptCache().remove('list_gosc');
    return json_({ ok: true, id: file.getId() });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}
