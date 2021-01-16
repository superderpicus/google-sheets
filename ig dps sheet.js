// IG DPS Sheet preset script
// by taters

//////////////////////////////////////////////////////////////////////////////////
//    Setup the Menu
//////////////////////////////////////////////////////////////////////////////////
function onOpen() 
{
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Presets')
       .addSubMenu(ui.createMenu('Glaives')
          .addItem('Fresh', 'clear')
          .addItem('Safi Meta', 'safiMeta')
          .addItem('Brachy Meta', 'brachyMeta')
          .addItem('Safi MAX DPS', 'safiMax')
          .addItem('No Tenderize Safi', 'noTendy')
          .addItem('Kjarr Para', 'kPara')
          .addItem('Kjarr Water', 'kjarrWater')
          .addItem('Safi Dragonvein Element', 'safiEle')
          .addItem('Safi Peak Element', 'safiPeakEle')
          .addItem('Ala Peak Element', 'alaPeak')
          .addItem('Ala TCE', 'alaTCE')
          .addItem('Ala SafiBrachy', 'alaSafiBrachy')
          .addItem('Fatalis', 'fatalis'))
      .addSubMenu(ui.createMenu('Kinsects')
          .addItem('Folicath 3 Forz', 'folicath')
          .addItem('Vezirstag 3 Forz', 'vezir')
          .addItem('Valorwing 3 Medis', 'valorwing')
          .addItem('Gleambeetle 3 Velox', 'gleambeetle')
          .addItem('Nexus Dragon Soul', 'dragonsoul'))
      .addSubMenu(ui.createMenu('Hitzones')
          .addItem('Training Pole', 'trainingPole')
          .addItem('Velk Head (Broken)', 'velk')
          .addItem('Teostra Head', 'teoHead')
          .addItem('Kulve Head (Released)', 'ktHead')
          .addItem('Rajang Head', 'raj')
          .addItem('Rathalos Head', 'RathalosHead')
          .addItem('Ruiner Head', 'RuinerHead')
          .addItem('Ruiner Head (Spikes)', 'RuinerWhite')
          .addItem('Savage Deviljho Chest', 'JhoChest')
          .addItem('Seething Bazel Tail', 'bazelTail')
          .addItem('Stygian Zinogre Head', 'stygHead')
          .addItem('Tigrex Head', 'tigrexHead')
          .addItem('RBrachy Head (No Slime)', 'rbrachHead')
          .addItem('RBrachy Hand (No Slime)', 'rbrachHand')
          .addItem('Alatreon Head', 'alaHead')
          .addItem('Alatreon Arm', 'alaArm')
          .addItem('Fatalis Head', 'fatHead')
          .addItem('Fatalis Chest P3', 'fatChest'))
       .addSeparator()
       .addItem('Reset Attack Stats', 'reset')
      .addToUi();
}

//////////////////////////////////////////////////////////////////////////////////
//    Quick Functions
//////////////////////////////////////////////////////////////////////////////////
function trainingPole(){ setHZV('80%', '30%'); }
function velk(){ setHZV('65%', '20%'); }
function raj(){ setHZV('60%', '30%'); }
function rbrachHead(){ setHZV('60%', '25%'); }
function rbrachHand(){ setHZV('55%', '25%'); }    
function RathalosHead(){ setHZV('60%', '30%'); }    
function RuinerHead(){ setHZV('45%', '15%'); } 
function RuinerWhiteHead(){ setHZV('75%', '15%'); }
function JhoChest(){ setHZV('60%', '30%'); }
function bazelTail(){ setHZV('65%', '25%'); }
function teoHead(){ setHZV('55%', '30%'); }
function ktHead(){ setHZV('90%', '20%'); }
function stygHead(){ setHZV('55%', '15%'); }
function tigrexHead(){ setHZV('65%', '25%'); }
function alaHead(){ setHZV('88%', '14%'); }
function alaArm(){ setHZV('70%', '22%'); }
function fatHead(){ setHZV('76%', '25%'); }
function fatChest(){ setHZV('72%', '15%'); }

// kinsect callback functions
function folicath() { setKinsect('14', '20' ); }
function vezir() { setKinsect('19', '20' ); }
function gleambeetle() { setKinsect('16', '20' ); }
function valorwing() { setKinsect('4', '18' ); }
function dragonsoul() { setKinsect('16', '16' ); }

// standard variables
var EFRCalcsRange = ['C10:C21', 'F10:F15', 'I10:I14', 'L10:L14', 'C25:C28', 'F25:F26', 'I25', 'L25:L29', 'C37:C38', 'C44', 'C45'];
var glaive1Ranges = ['C7:C18', 'F7:F12', 'I7:I11', 'L7:L11', 'C22:C25', 'F22:F23', 'I22', 'L22:L26', 'F29:F30', 'F36', 'F37'];
var glaive2Ranges = ['O7:O18', 'R7:R12', 'U7:U11', 'X7:X11', 'O22:O25', 'R22:R23', 'U22', 'X22:X26', 'U29:U30', 'U36', 'U37'];
var blankValues = [ [[0],[0],[0],[0],[0],[0],[0],[0],[0],[0],[0],[0]], [[0], [0], [1], [1], [1], [1]], [[1], [1], [1], [1], [1]], [[1], [1], [1], [1], [1]], [[0], [0], [0], [0]], [[0], ['FALSE']], 1, [[1], [1], [1], [1], [1]], [[0], [0]],  '100%', 'FALSE'];
var SafiMetaValues = [ [[15],[10],[10],[7],[9],[6],[12],[0],[20],[18],[28],[0]], [[1], [1], [1], [1], [1], [1]], [[1.39], [1.4], [1], [1], [1]], [[1.15], [1], [1], [1], [1]], [[0], [0], [0], [0]], [[1], ['FALSE']], 1, [[1], [1], [1], [1], [1]], [[310], [0]],  '100%', 'FALSE'];
var ClearValues = [ [[0],[0],[0],[0],[0],[0],[0],[0],[0],[0],[0],[0]], [[1], [1], [1], [1], [1], [1]], [[1], [1], [1], [1], [1]], [[1], [1], [1], [1], [1]], [[0], [0], [0], [0]], [[1], ['FALSE']], 1, [[1], [1], [1], [1], [1]], [[0], [0]], '0%', 'FALSE' ];
var KParaValues = [ [[15],[10],[10],[7],[9],[6],[0],[0],[20],[21],[28],[0]], [[1], [1], [1], [1], [1], [1]], [[1.32], [1.4], [1], [1], [1]], [[1.15], [1], [1], [1], [1]], [[0], [0], [0], [0]], [[1], ['FALSE']], 1, [[1], [1], [1], [1], [1]], [[312], [0]],  '100%', 'FALSE'];
var BrachyMetaValues = [ [[15],[10],[10],[7],[9],[6],[12],[0],[20],[21],[28],[0]], [[1], [1], [1], [1], [1], [1]], [[1.39], [1.4], [1], [1], [1]], [[1.15], [1], [1], [1], [1]], [[0], [0], [0], [0]], [[1], ['FALSE']], 1, [[1], [1], [1], [1], [1]], [[305], [0]],  '100%', 'FALSE'];
var KWaterValues = [ [[15],[10],[10],[7],[9],[6],[0],[0],[20],[12],[28],[0]], [[1], [1], [1], [1], [1], [1]], [[1.32], [1.4], [1], [1], [1]], [[1.15], [1], [1], [1], [1]], [[0], [0], [0], [0]], [[1], ['FALSE']], 1, [[1], [1], [1], [1.35], [1.15]], [[293], [53]],  '100%', 'FALSE'];
var SafiEleValues = [ [[15],[10],[10],[7],[9],[6],[0],[0],[0],[21],[28],[25]], [[1], [1], [1], [1], [1], [1]], [[1.39], [1.4], [1], [1], [1]], [[1.15], [1], [1], [1], [1]], [[8], [0], ['10 +20%'], [0]], [[1], ['FALSE']], 1, [[1], [1], [1], [1], [1.25]], [[300], [18]],  '100%', 'FALSE'];
var SafiPeakEleValues = [ [[15],[10],[10],[7],[9],[6],[0],[0],[20],[21],[28],[0]], [[1], [1], [1], [1], [1], [1]], [[1.39], [1.4], [1], [1], [1]], [[1.15], [1], [1], [1], [1]], [[0], [0], [10], [0]], [[1], ['FALSE']], 1, [[1], [1], [1], [1], [1.25]], [[300], [15]],  '100%', 'FALSE'];
var NoTendyValues = [ [[15],[10],[10],[7],[9],[6],[0],[0],[20],[15],[28],[0]], [[1], [1], [1], [1], [1], [1]], [[1.39], [1.4], [1], [1], [1]], [[1.15], [1], [1], [1], [1]], [[0], [0], [0], [0]], [[1], ['FALSE']], 1, [[1], [1], [1], [1], [1]], [[305], [0]],  '100%', 'FALSE'];
var SafiMaxValues = [ [[15],[10],[10],[7],[9],[6],[12],[0],[20],[21],[28],[0]], [[1], [1], [1], [1], [1], [1]], [[1.39], [1.4], [1], [1], [1]], [[1.15], [1], [1], [1], [1]], [[0], [0], [0], [0]], [[1], ['FALSE']], 1, [[1], [1], [1], [1], [1]], [[310], [0]],  '100%', 'FALSE'];
var AlaKaiserBrachy = [ [[15],[10],[10],[7],[9],[6],[0],[0],[20],[12],[28],[0]], [[1], [1], [1], [1], [1], [1]], [[1.39], [1.4], [1], [1], [1]], [[1.15], [1], [1], [1], [1]], [[0], [0], [10], [0]], [[1], ['FALSE']], 1, [[1], [1], [1], [1], [1.25]], [[280], [54]],  '100%', 'FALSE'];
var AlaTCE = [ [[15],[10],[10],[7],[9],[6],[0],[0],[0],[0],[20],[0]], [[1], [1], [1], [1], [1], [1]], [[1.39], [1.4], [1], [1], [1]], [[1.15], [1], [1], [1], [1]], [[0], [0], ['10 +20%'], [0]], [[1], ['FALSE']], 1, [[1], [1], [1], [1.55], [1.25]], [[280], [54]],  '100%', 'FALSE'];
var alaSafi = [ [[15],[10],[10],[7],[9],[6],[0],[0],[0],[18],[28],[25]], [[1], [1], [1], [1], [1], [1]], [[1.39], [1.4], [1], [1], [1]], [[1.15], [1], [1], [1], [1]], [[8], [0], ['10 +20%'], [0]], [[1], ['FALSE']], 1, [[1], [1], [1], [1], [1.25]], [[290], [57]],  '100%', 'FALSE'];
var fatValues = [ [[15],[10],[10],[7],[9],[6],[18],[0],[20],[21],[28],[0]], [[1], [1], [1], [1], [1], [1]], [[1.39], [1.4], [1], [1], [1]], [[1.15], [1], [1], [1], [1]], [[8], [0], [0], [0]], [[1], ['FALSE']], 1, [[1], [1], [1], [1], [1]], [[355], [12]], '95%', 'FALSE'];


// sheets ( indexed starting from 0 )
var AttackSheet = 0;
var EFRCalcsSheet = 1;
var comboSheet = 2;
var kinsectSheet = 3;
var compareSheet = 4;

//////////////////////////////////////////////////////////////////////////////////
//    Large Functions
//////////////////////////////////////////////////////////////////////////////////
// message box function for ease of use
function showMessageBox(title, str) 
{
  var ui = SpreadsheetApp.getUi();
  Browser.msgBox(title, str, ui.ButtonSet.OK);
}

// set hitzone values function
// also shows a message box
function setHZV(raw, ele)
{
  var ss = SpreadsheetApp.getActiveSpreadsheet() // active spreadsheet (IG DPS Calcs)
  
  var sheet = ss.getSheets()[0] // Attack Stats sheet
  sheet.getRange('J5').setValue(raw); // raw hzv
  sheet.getRange('J6').setValue(ele); // ele hzv
  
  var sheet2 = ss.getSheets()[1] // EFR Calcs sheet
  sheet.getRange('C39').setValue(raw); // raw hzv
  sheet.getRange('C40').setValue(ele); // ele hzv

  var sheet3 = ss.getSheets()[3] // kinsect compare sheet
  sheet.getRange('I5').setValue(raw); // raw hzv
  sheet.getRange('I6').setValue(ele); // ele hzv

  var sheet4 = ss.getSheets()[4] // Compare set sheet
  sheet.getRange('F31').setValue(raw); // raw hzv
  sheet.getRange('F32').setValue(ele); // ele hzv

  var sheet5 = ss.getSheets()[4] // Compare set sheet
  sheet.getRange('U31').setValue(raw); // raw hzv
  sheet.getRange('U32').setValue(ele); // ele hzv
  
  showMessageBox('Hitzones Set', 'Hitzones set to: ' + raw + ' / ' + ele);
  Logger.log('Hitzones set to %s / %s', raw, ele);
}

// set kinsect power function
// also shows a message box
function setKinsect(raw, ele)
{
  var ss = SpreadsheetApp.getActiveSpreadsheet() // active spreadsheet (IG DPS Calcs)
  
  var sheet = ss.getSheets()[0] // Attack Stats sheet
  sheet.getRange('J7').setValue(raw); // raw level
  sheet.getRange('J8').setValue(ele); // ele level
  
  showMessageBox('Kinsect Power Set', 'Kinsect Power set to: ' + raw + ' / ' + ele);
  Logger.log('Kinsect Power set to %s / %s', raw, ele);
}

// reset callback function
function reset()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // active spreadsheet (IG DPS Calcs)  
  var sheet = ss.getSheets()[EFRCalcsSheet]; // EFR Calcs sheet
  
  setValues(EFRCalcsSheet, EFRCalcsRange, blankValues)
  showMessageBox('Reset Attack Stats', 'MVs, status, and element mods all reset');
}

// on editing a cell :
// designed to handle the compare sets sheet
function onEdit(e)
{
  var ss = SpreadsheetApp.getActiveSpreadsheet() // active spreadsheet (IG DPS Calcs)
  var activeSheet = ss.getSheetName(); // active sheet name as string
  var range = e.range; // range is the cells that were edited
  var row = range.getRow(); // row = numbers
  var column = range.getColumn(); // column = letters
  var value = range.getValue(); // text of the cell edited

  if (ss.getSheetName() == "Compare Sets")
  {
    // glaive 1 box is: B3 AKA 2, 3
    var glaive = 0;
    if (row == 3 && column == 2)    
      glaive = 1;    
    else if(row == 3 && column == 14)    
      glaive = 2;
    
    if (glaive == 0)
      return;
    
    var _range = glaive == 1 ? glaive1Ranges : glaive2Ranges;
    if (value != null && value != "None")
    {
      switch(value)
      {
      case "Fresh":
        setValues(CompareSheet, _range, ClearValues);
        break;
      case "Safi Meta":
        setValues(CompareSheet, _range, SafiMetaValues);
        break;
      case "Brachy Meta":
        setValues(CompareSheet, _range, BrachyMetaValues);
        break;
      case "Safi MAX":
        setValues(CompareSheet, _range, SafiMaxValues);
        break;
      case "No Tendy Safi":
        setValues(CompareSheet, _range, NoTendyValues);
        break;
      case "Kjarr Para":
        setValues(CompareSheet, _range, KParaValues);
        break;
      case "Kjarr Water":
        setValues(CompareSheet, _range, KWaterValues);
        break;
      case "Safi Dragonvein Ele":
        setValues(CompareSheet, _range, SafiEleValues);
        break;
      case "Safi Peak Ele":
        setValues(CompareSheet, _range, SafiPeakEleValues);
        break;
      }
    }
  }
}

// EFR calc callback functions
function safiMeta()
{
  setValues(EFRCalcsSheet, EFRCalcsRange, SafiMetaValues);
  showMessageBox('EFR Settings Set', 'EFR Settings set to: Safi Meta');
  Logger.log('EFR Settings set to: Safi Meta');
}

function fatalis()
{
  setValues(EFRCalcsSheet, EFRCalcsRange, fatValues);
  showMessageBox('EFR Settings Set', 'EFR Settings set to: Fatalis');
  Logger.log('EFR Settings set to: Fatalis');
}

function kPara()
{
  setValues(EFRCalcsSheet, EFRCalcsRange, KParaValues); 
  showMessageBox('EFR Settings Set', 'EFR Settings set to: Kjarr Para');
  Logger.log('EFR Settings set to: Kjarr Para');
}

function alaPeak()
{
  setValues(EFRCalcsSheet, EFRCalcsRange, AlaKaiserBrachy); 
  showMessageBox('EFR Settings Set', 'EFR Settings set to: Ala Peak Ele');
  Logger.log('EFR Settings set to: Ala Peak Ele');
}

function alaTCE()
{
  setValues(EFRCalcsSheet, EFRCalcsRange, AlaTCE); 
  showMessageBox('EFR Settings Set', 'EFR Settings set to: Ala TCE');
  Logger.log('EFR Settings set to: Ala Peak Ele');
}

function alaSafiBrachy()
{
  setValues(EFRCalcsSheet, EFRCalcsRange, alaSafi); 
  showMessageBox('EFR Settings Set', 'EFR Settings set to: Ala Safi Brachy');
  Logger.log('EFR Settings set to: Ala Peak Ele');
}



function brachyMeta()
{
  setValues(EFRCalcsSheet, EFRCalcsRange, BrachyMetaValues);
  showMessageBox('EFR Settings Set', 'EFR Settings set to: Brachy Meta');
  Logger.log('EFR Settings set to: Brachy Meta');
}

function kjarrWater()
{
  setValues(EFRCalcsSheet, EFRCalcsRange, KWaterValues);
  showMessageBox('EFR Settings Set', 'EFR Settings set to: Kjarr Water');
  Logger.log('EFR Settings set to: Kjarr Water');
}

function safiEle()
{
  setValues(EFRCalcsSheet, EFRCalcsRange, SafiEleValues);
  showMessageBox('EFR Settings Set', 'EFR Settings set to: Safi Ele');
  Logger.log('EFR Settings set to: Safi Ele');
}

function safiPeakEle()
{
  setValues(EFRCalcsSheet, EFRCalcsRange, SafiPeakEleValues);   
  showMessageBox('EFR Settings Set', 'EFR Settings set to: Safi Peak Ele');
  Logger.log('EFR Settings set to: Safi Peak Ele');
}

function noTendy()
{
  setValues(EFRCalcsSheet, EFRCalcsRange, NoTendyValues);
  showMessageBox('EFR Settings Set', 'EFR Settings set to: Safi No Tendy');
  Logger.log('EFR Settings set to: Safi No Tendy');
}

function safiMax()
{
  setValues(EFRCalcsSheet, EFRCalcsRange, SafiMaxValues);
  showMessageBox('EFR Settings Set', 'EFR Settings set to: Safi Max DPS');
  Logger.log('EFR Settings set to: Safi Max DPS');
}

function clear()
{  
  setValues(EFRCalcsSheet, EFRCalcsRange, ClearValues);
  showMessageBox('EFR Settings Set', 'EFR Settings set to: Fresh');
  Logger.log('EFR Settings set to: Fresh');
}


// main function to set values
// Params:
//   sheet : sheet array index (0 to max)
//   ranges : array of ranges each module is in
//   values : array of values to assign to ranges
function setValues(sheet, ranges, values)
{
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // active spreadsheet (IG DPS Calcs)  
  var sheet = ss.getSheets()[sheet]; // EFR Calcs sheet
  
  // static raw buffs
  var range = sheet.getRange(ranges[0]);
  range.setValues(values[0]);
  
  // base raw multi
  range = sheet.getRange(ranges[1]);
  range.setValues(values[1]);
  
  // damage multi
  range = sheet.getRange(ranges[2]);
  range.setValues(values[2]);
  
  // total raw multi
  range = sheet.getRange(ranges[3]);
  range.setValues(values[3]);
  
  // static ele buffs
  range = sheet.getRange(ranges[4]);
  range.setValues(values[4]);
  
  // ele base buffs
  range = sheet.getRange(ranges[5]);
  range.setValues(values[5]);
  
  // precap ele mods
  range = sheet.getRange(ranges[6]);
  range.setValue(values[6]);
  
  // postcap ele mods
  range = sheet.getRange(ranges[7]);
  range.setValues(values[7]);
  
  // raw
  range = sheet.getRange(ranges[8]);
  range.setValues(values[8]);
  
  // affinity
  range = sheet.getRange(ranges[9]);
  range.setValue(values[9]);
  
  // tenderized
  range = sheet.getRange(ranges[10]);
  range.setValue(values[10]);
  
}