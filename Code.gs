/**
 * FITANTANANA NY TSY FAHATONGAVANA — Divizionan'ny Visa
 * Backend Google Apps Script (miasa miaraka amin'ny index.html)
 *
 * INONA ITY FICHIER ITY :
 * Apetraho ao amin'ny Google Apps Script mifamatotra amin'ny Google
 * Sheet iray ity code ity, dia "publier" ho "Web app". Ny website
 * (index.html) dia hifandray amin'ny adiresy ("URL") havoakan'io
 * fametrahana io mba hitehirizana sy haka ny fangatahana rehetra.
 *
 * Jereo ny fichier TOROLALANA.txt momba ny dingana fanaovana azy.
 */

const SHEET_NAME = "Demandes";
const HEADERS = ["id","agentId","agentNom","agentMatricule","type","debut","fin","motif","statut","creeLe"];

function getSheet(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if(!sheet){
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if(sheet.getLastRow() === 0){
    sheet.appendRow(HEADERS);
  }
  return sheet;
}

function jsonOut(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* GET : mamerina ny lisitry ny fangatahana rehetra, endrika JSON */
function doGet(e){
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const rows = data
    .filter(row => row[0] !== "" && row[0] !== null)
    .map(row => {
      const obj = {};
      headers.forEach((h,i)=> obj[h] = row[i]);
      return obj;
    });
  return jsonOut(rows);
}

/* POST : mandraikitra fangatahana vaovao, manova sata, na mamafa */
function doPost(e){
  let body;
  try{
    body = JSON.parse(e.postData.contents);
  }catch(err){
    return jsonOut({success:false, error:"Angona diso endrika"});
  }

  const sheet = getSheet();
  const action = body.action;

  if(action === 'add'){
    const d = body.demande || {};
    sheet.appendRow([
      d.id, d.agentId, d.agentNom || "", d.agentMatricule || "",
      d.type, d.debut, d.fin, d.motif || "", d.statut, d.creeLe
    ]);
    return jsonOut({success:true});
  }

  if(action === 'updateStatus'){
    const data = sheet.getDataRange().getValues();
    for(let i=1; i<data.length; i++){
      if(String(data[i][0]) === String(body.id)){
        sheet.getRange(i+1, 9).setValue(body.statut); // andalana 9 = "statut"
        break;
      }
    }
    return jsonOut({success:true});
  }

  if(action === 'delete'){
    const data = sheet.getDataRange().getValues();
    for(let i=1; i<data.length; i++){
      if(String(data[i][0]) === String(body.id)){
        sheet.deleteRow(i+1);
        break;
      }
    }
    return jsonOut({success:true});
  }

  return jsonOut({success:false, error:"Tsy fantatra ny hetsika angatahina"});
}
