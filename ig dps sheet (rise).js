// IG DPS Sheet (Rise Edition) by taters

// set up the dropdown preset menus
function onOpen()
{
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Presets')
       .addSubMenu(ui.createMenu('Glaives')
          .addItem('Demo', 'demoGlaive'))
       .addSubMenu(ui.createMenu('Hitzones')
          .addItem('Arzuros\' Ass', 'arzAss')
          .addItem('Great Izuchi\'s Head', 'izuchiHead'))
      .addToUi();
};

function showMessageBox(title, str) 
{
  var ui = SpreadsheetApp.getUi();
  Browser.msgBox(title, str, ui.ButtonSet.OK);
}

var dataSheet = 0;
var rawHZVCell = 'J2';
var eleHZVCell = 'J3';

function setHZV(raw, ele)
{
  var ss = SpreadsheetApp.getActiveSpreadsheet() // active spreadsheet
  var sheet = ss.getSheets()[dataSheet] // data sheet
  sheet.getRange(rawHZVCell).setValue(raw); // raw hzv
  sheet.getRange(eleHZVCell).setValue(ele); // ele hzv  
  
  showMessageBox('Hitzones Set', 'Hitzones set to: ' + raw + ' / ' + ele);
  Logger.log('Hitzones set to %s / %s', raw, ele);
};

var extCell = 'C1';
var enrCell = 'C2';
var sharpCell = 'C3';
var rawCell = 'C4';
var eleCell = 'C5';
var critCell = 'C6';
var affCell = 'C7';

function setStats(id, ext, enbr, sharp, raw, ele, crit, affinity)
{
  var ss = SpreadsheetApp.getActiveSpreadsheet() // active spreadsheet
  var sheet = ss.getSheets()[dataSheet] // data sheet
  sheet.getRange(extCell).setValue(ext); 
  sheet.getRange(enrCell).setValue(enbr); 
  sheet.getRange(sharpCell).setValue(sharp); 
  sheet.getRange(rawCell).setValue(raw); 
  sheet.getRange(eleCell).setValue(ele); 
  sheet.getRange(critCell).setValue(crit); 
  sheet.getRange(affCell).setValue(affinity); 
  
  showMessageBox('Stats Set', 'Stats set to: ' + id);
  Logger.log('STats set to %s', id);  
};

// HZV
function arzAss(){ setHZV('66%', '0%'); }
function izuchiHead(){ setHZV('80%', '0%'); }

// GLAIVE
function demoGlaive(){ setStats('Demo Glaive', 1.15, 1.0, 1.05, 140, 0, 1.25, 0); }