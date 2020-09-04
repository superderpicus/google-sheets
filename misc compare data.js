// misc compare data.js
// by taters

// This script is created to automate some calculations,
// it is meant to estimate the damage of various
// sets vs certain matchups when ran.

// loops is an array of different combo loops,
// currently only tornado loop is supported
var loops =
[  
    // hit count, mv, ele mod, status mod
    [6, [12, 11, 18, 22, 24, 42], [0.8, 0.8, 0.8, 0.8, 0.8, 0.8], [0.8, 0.8, 0.8, 0.8, 0.8, 0.8] ], // tornado loop
    [6, [28, 42,  7, 36, 24, 42], [1.0, 1.0, 0.0, 1.0, 0.8, 0.8], [1.0, 1.0, 0.0, 1.0, 0.8, 0.8] ] // descending loop
];

// index values for certain loops
var tornado = 0;
var dt = 1;

// function to generate random colour for new tab (who knows lol)
function random_hex_color_code()
{
  var n = (Math.random() * 0xfffff * 1000000).toString(16);
  return '#' + n.slice(0, 6);
};

// calculate various damage function
function calc(weap, loop, hp, rawHZV, eleHZV, questMultiplier1, questMultiplier2, firstProc, inc, _cap, procDmg, agiUptime, agiMulti, outputRange, main)
{
    var sheet = main; // original input sheet
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // 'misc compare data' sheet
    var newSheet = activeSpreadsheet.getSheetByName('simulation_output'); // output sheet
    var info = sheet.getRange(weap).getValues(); // get weapon data from main sheet
    var fillColour = sheet.getRange(weap).getBackgroundObject() // get background colour from the current set's range

    // load values from the info array, 
    // which was loaded with efr, efe, etc from 
    // original input sheet.
    var name = info[0][0];
    var efr = info[0][1];
    var efe = info[0][2];
    var efs = info[0][3];

    // load weapon options into variables
    // that we will be using later in the calcs
    var _efr = efr / 100;
    var _efe = efe;
    var _efs = efs;
    var blast = false;
     if (efs > 0)
        blast = true; // set flag that weapon is using blast

    // create variables to be used in the loop
    var blastValue = 0; // current blast value    
    var procs = 0; // blast proc counter
    var threshold = firstProc * questMultiplier1; // threshold is blast required for a proc
    var increase = inc * questMultiplier2; // increase is the additional amount needed for proc on top of previous threshold
    var cap = _cap * questMultiplier2; // cap for threshold, if it goes higher always set it to the cap
    var monsterHP = hp; // set current hp for monster to the max hp loaded from options
    var hit = 0; // which hit in the combo we are currently on
    var hitCount = loops[loop][0] // load max hit count for loop
    var mvs = loops[loop][1]; // load mv for loop
    var eleMods = loops[loop][2]; // load ele mod for loop
    var statusMods = loops[loop][3]; // load status mod for loop
    var hits = 0; // total hit counter for the matchup provided
    var blastRaw = 0; // blast raw damage 
    var blastDamage = 0; // blast portion of the damage
    var eleRaw = 0; // element raw damage portion
    var eleDamage = 0; // element actual element portion
    var _agiMulti = ((agiMulti - 1) * agiUptime) + 1; // average engraged multiplier

    // while monster is alive,  deal damage based 
    // on the next hit
    while (monsterHP > 0)
    {        
        // if the blast flag was set
        // to true, apply blast calcs
        if (blast)
        {
            var a = (mvs[hit] * _efr * rawHZV * _agiMulti); // raw damage calculation for blast            
            monsterHP -= a; // deduct the hp calculated
            blastRaw += a; // add to the blast raw portion variable
            blastValue += (_efs * statusMods[hit]); // add to blast threshold
            // check if blast is procced
            if (blastValue >= threshold)
            {
                // blast is procced, reset current value and add blast damage
                blastValue = 0; // reset blast to 0
                procs++; // increment procs
                blastDamage += procDmg; // increment blast damage by proc damage
                monsterHP -= procDmg; // lower monster hp based on proc damage
                threshold += increase; // increase threshold
                // check if threshold is over the cap. if it is, 
                // set it to the cap instead
                if (threshold > cap)                
                    threshold = cap;                
            }
        }
        else // element calculation, since blast was not set
        {
            var a = (mvs[hit] * _efr * rawHZV * _agiMulti); // raw damage calculation for element
            var b = (eleMods[hit] * _efe * eleHZV * _agiMulti); // element damage calculation for element
            monsterHP = monsterHP - a; // raw dmg deducted
            monsterHP = monsterHP - b; // ele dmg deducted
            eleRaw = eleRaw + a; // add to the raw portion variable
            eleDamage = eleDamage + b; // add to the element portion variable
        }
        // total hit counter increase by 1
        hits++;

        // increment current hit.
        // if hit is too high,
        // set to 0
        hit++;
        if (hit >= hitCount)
            hit = 0;
    }
   
    // write output to new sheet
    if (blast) // blast
        newSheet.getRange(outputRange).setValues([[name, blastRaw, blastDamage, hits, ((loop == 0) ? 'Tornado' : 'DT')]]);    
    else // element          
        newSheet.getRange(outputRange).setValues([[name, eleRaw, eleDamage, hits, ((loop == 0) ? 'Tornado' : 'DT')]]);
    newSheet.getRange(outputRange).setNumberFormats([["0.#", "0.#", "0.#", "0.#", "0.#"]]) // format the numbers for 1 decimal place max
    newSheet.getRange(outputRange).setBackgroundObject(fillColour); // set background colour to the colour loaded earlier
}

// function to create new sheet because lazy
function createNewSheet(name)
{
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // 'misc compare data' sheet
    var newSheet = activeSpreadsheet.getSheetByName(name); // check if the sheet exists already
    if (newSheet != null) // if it exists, delete it 
        activeSpreadsheet.deleteSheet(newSheet);
    newSheet = activeSpreadsheet.insertSheet(); // create new spreadsheet
    newSheet.setName(name); // rename the new spreadsheet
    newSheet.setTabColor(random_hex_color_code()); // generate random hex colour for the tab and set it
    return newSheet; // return sheet object
}

// main entry function
function calcBlast()
{
    // google sheet stuff
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // 'misc compare data' sheet
    var sheet = SpreadsheetApp.getActiveSheet(); // current input sheet
    // load the settings cells into array
    var monsterOptions = sheet.getRange('P13:P23').getValues();
    // load monster options into variables
    var hp = monsterOptions[0];
    var rawHZV = monsterOptions[1];
    var eleHZV = monsterOptions[2];
    var questMultiplier1 = monsterOptions[3];
    var questMultiplier2 = monsterOptions[4];
    var firstProc = monsterOptions[5];
    var inc = monsterOptions[6];
    var _cap = monsterOptions[7];
    var procDmg = parseInt(monsterOptions[8], 10);
    var agiUptime = monsterOptions[9];
    var agiMulti = monsterOptions[10];
    // options loaded, create new sheet and set up the headers 
    var newSheet = createNewSheet('simulation_output');
    newSheet.getRange('B1:F1').setValues([['Set', 'Raw Damage', 'Ele/Blast Damage', 'Hits to Kill', "Loop"]]);
    newSheet.getRange('B1:F1').setFontWeights([['bold', 'bold', 'bold', 'bold', 'bold']]);
    newSheet.getRange('B1:F1').setHorizontalAlignments([['center', 'center', 'center', 'center', 'center' ]])
    newSheet.getRange('A1:A5').setValues([[hp + ' HP'], ['raw: ' + rawHZV], ['ele: ' + eleHZV], ['agi: ' + agiUptime], ['mult: ' + agiMulti]]);
    // headers are now set. time to iterate the main sets and calculate their damage
    var first = 13; // the first set starts at 'A13'
    var outputFirst = 2; // the first output cell starts at 'B2'
    // loop through all 24 sets for the main tornado loop section
    for (x = 0; x <= 23; x++)
    {
        calc(('A' + first) + ':' + ('D' + first), tornado, hp, rawHZV, eleHZV, questMultiplier1, questMultiplier2, firstProc, inc, _cap, procDmg, agiUptime, agiMulti, (('B' + outputFirst) + ':' + ('F' + outputFirst)), sheet)   
        first++; // increment to get next set
        outputFirst++; // increment to ouput to 1 line below
    }
    /* DT
    first = 13;
    outputFirst++;
    for (x = 0; x <= 23; x++)
    {
        calc(('H' + first) + ':' + ('M' + first), dt, hp, rawHZV, eleHZV, questMultiplier1, questMultiplier2, firstProc, inc, _cap, procDmg, agiUptime, agiMulti, (('B' + outputFirst) + ':' + ('F' + outputFirst)), sheet)   
        first++;
        outputFirst++;
    }
    */

    // filter will let you sort ascending or descending based on columns
    var filter = newSheet.getRange('B1:F25').createFilter();
    // this loop will set columns to the widest size they need to be,
    // while also remaining readable
    for (var x = 1; x < 25; x++)
    {
        newSheet.autoResizeColumn(x); 
    }

}

// onOpen function:
// create dropdown menu when loaded
function onOpen() 
{
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Presets')
       .addSubMenu(ui.createMenu('Non-Tenderized')
          .addItem('will', 'clear')
          .addItem('add', 'safiMeta')
          .addItem('later', 'brachyMeta')
          .addItem('but', 'safiMax')
          .addItem('im', 'noTendy')
          .addItem('lazy', 'kPara'))
      .addSubMenu(ui.createMenu('Tenderized')
          .addItem('Folicath 3 Forz', 'folicath')
          .addItem('Vezirstag 3 Forz', 'vezir')
          .addItem('Valorwing 3 Medis', 'valorwing')
          .addItem('Gleambeetle 3 Velox', 'gleambeetle')
          .addItem('Nexus Dragon Soul', 'dragonsoul'))
       .addSeparator()
       .addItem('Reset Values', 'reset')
      .addToUi();
}

var totalMV = 1.29;
var totalEle = 4.8;
var animationTime = 4.22; 

var dtTotalMV = 1.79;
var dtAnimationTime = 6.07;

var sets =
[
    ['Safi Shatterspear (100% Coal)',   1002.19, 0, 7.33, 13],
    ['Safi Shatterspear (75% Coal)',    996.35, 0, 7.33, 14],
    ['Safi Shatterspear (50% Coal)',    988.57, 0, 7.33, 15],
    ['Safi Shatterspear (25% Coal)',    982.73, 0, 7.33, 16],
    ['Safi Shatterspear (0% Coal)',     978.84, 0, 7.33, 17],
    ['Lightbreak Press (100% Coal)',    990.51, 0, 7.33, 18],
    ['Lightbreak Press (75% Coal)',     984.68, 0, 7.33, 19],
    ['Lightbreak Press (50% Coal)',     978.84, 0, 7.33, 20],
    ['Lightbreak Press (25% Coal)',     971.05, 0, 7.33, 21],
    ['Lightbreak Press (0% Coal)',      965.22, 0, 7.33, 22],
    ['Alatreon\'s Embrace',             969.11, 85, 0, 23],
    ['Safi Raw Ele (100% Coal)',        1002.19, 26.25, 0, 24],
    ['Safi Raw Ele (75% Coal)',         996.35, 25.31, 0, 25],
    ['Safi Raw Ele (50% Coal)',         988.57, 24.38, 0, 26],
    ['Safi Raw Ele (25% Coal)',         982.73, 23.44, 0, 27],
    ['Safi Raw Ele (0% Coal)',          974.95, 22.5, 0, 28],
    ['KPara',                           911.06, 0, 0, 29],
    ['KWater',                          872.26, 97.81, 0, 30],
    ['Safihemoth',                      945.76, 50, 0, 31],
    ['Brakulve (100% Coal)',            980.78, 35, 0, 32],
    ['Brakulve (75% Coal)',             973, 34.06, 0, 33],
    ['Brakulve (50% Coal)',             967.16, 33.13, 0, 34],
    ['Brakulve (25% Coal)',             959.38, 32.19, 0, 35],
    ['Brakulve (0% Coal)',              953.54, 31.25, 0, 36],
]

function setValue(index, loop)
{

    var sheet = SpreadsheetApp.getActiveSheet(); // current input sheet
    if (loop == 0)
    {
        var values = sets[index];
        sheet.getRange('A' + values[4] + ':F' + values[4]).setValues([[values[0], values[1], values[2], values[3], '=((' + totalMV + ') * B' + values[4] + ' * ($B$3/100) + (' + totalEle + '*(C' + values[4] + '/100)*$B$4))', '=E' + values[4] + '/' + animationTime]]);
    }
    else
    {
        var values = sets[index];
        sheet.getRange('H' + values[4] + ':M' + values[4]).setValues([[values[0], values[1], values[2], values[3], '=((' + dtTotalMV + ') * I' + values[4] + ' * ($B$3/100) + (' + totalEle + '*(J' + values[4] + '/100)*$B$4)) + ($B$8*Data!$C$12)', '=L' + values[4] + '/' + dtAnimationTime]]);
    }
}

function reset()
{

    var sheet = SpreadsheetApp.getActiveSheet(); // current input sheet
    for (var x = 0; x < sets.length; x++)
    {
        setValue(x, 0);
        setValue(x, 1);
    }

    sheet.getRange('B3:B8').setValues([[80], [30], [19], [20], ['Regular'], [3]]);
}