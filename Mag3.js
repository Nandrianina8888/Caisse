//const admin = ["nandrianina8888@gmail.com","gaitan.ravonindratsiry@kaloes.com"];
const ceClasseur = SpreadsheetApp.getActive();
const feuille = ceClasseur.getSheetByName("Mag3");
const command = ceClasseur.getSheetByName("Command");
var caissiers = command.getRange("A6:A").getValues().flat().filter((c)=>(c));
var magasiniers = command.getRange("B6:B").getValues().flat().filter((c)=>(c));
//const caissiers = ["maherysetraandriarilala@gmail.com","sylviaarijaona@gmail.com"];
//const magasinier = ["ravonivoharison@gmail.com"];
const unauthorized = caissiers.concat(magasiniers);
const dateSet = ['B2', 'K2', 'T2', 'B61', 'K61', 'T61', 'B120', 'K120', 'T120'];
const inputZone = ['B5:G54', 'K5:P54', 'T5:Y54', 'B64:G113', 'K64:P113', 'T64:Y113', 'B123:G172', 'K123:P172', 'T123:Y172'];
const checkZone = ['I5:I54', 'R5:R54', 'AA5:AA54', 'I64:I113', 'R64:R113', 'AA64:AA113', 'I123:I172', 'R123:R172', 'AA123:AA172'];
const notChangeZone = ['A1:AA4', 'A55:AA63', 'A114:AA122', 'A173:AA175', 'A:A', 'H:H', 'J:J', 'Q:Q', 'S:S', 'Z:Z'];

function onEditLockCellRange(e) {
    var cetteFeuille = e.range.getSheet();
    var protectIndex = e.range.getA1Notation();
    var isChecked = cetteFeuille.getRange(protectIndex).getValue();
    const src = protectIndex.match(/([I|R|AA])([0-9]+)/);
    if (src != null && src.length == 3) {
	    const col = src[1];
	    const row = Number(src[2]);
	    var rangeLock = null;
	    var tmp = null;
	    switch (col) {
	      case "I": 
          rangeLock = `B${row}:${col}${row}`;
          break;
	      case "R": 
          rangeLock = `K${row}:${col}${row}`;
          break;
	      case "AA": 
          rangeLock = `T${row}:${col}${row}`;
          break;
	      } 
	    if (isChecked &&
          ((row > 4 && row < 55) ||
	        (row > 63 && row < 114) ||
	        (row > 122 && row < 173))) {
	      var protection = cetteFeuille.getRange(rangeLock).protect().setDescription("Validé pour Payement, validé et coché: " + protectIndex);
        protection.removeEditors(unauthorized);
        if (protection.canDomainEdit()) {
          protection.setDomainEdit(false);
        }
	    }
    }
}
function lockValid(){
  for (n=0;n<checkZone.length;n++) {
    var r = feuille.getRange(checkZone[n]);
    r.protect().removeEditors(magasiniers);
  }
}

function lockNotChangeZone() {
  for (n=0;n<notChangeZone.length;n++) {
    var r = feuille.getRange(notChangeZone[n]);
    r.protect().removeEditors(unauthorized);
  }
}

function unclockAll() {
  var protections = feuille.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < protections.length; i++) {
    var protection = protections[i];
    if (protection.canEdit()) {
      protection.remove();
    }
  };
  for (n=0;n<checkZone.length;n++) {
    var r = feuille.getRange(checkZone[n]);
    r.setValue(false).insertCheckboxes();
  }
}

function clearAll () {
  for (n=0;n<inputZone.length;n++) {
    var r = feuille.getRange(inputZone[n]);
    r.setValue("");
  };
  setPermission();
  Browser.msgBox("Madio");
}

function setDate() {
  var androany = new Date().toLocaleDateString('fr');
  for (n=0;n<dateSet.length;n++){
    feuille.getRange(dateSet[n]).setValue(androany);
  }
}

function onOpen(e) {
  setDate();
}

function setPermission() {
  unclockAll();
  lockNotChangeZone();
  lockValid();
}
