// misc compare data.js
// by taters

// This script is created to automate some calculations,
// it is meant to estimate the damage of various
// sets vs certain matchups when ran.

// array of monster data for easier automation
var monsters =
[
    // ['monster name', hp, enrage, raw hzv, fire, water, thunder, ice, dragon, ailment mod1, ailment mod 2, initial, buildup, cap, damage]
    // A  0 to 18
    ['Acidic Glavenus', 19136, 1, 57, 20, 10, 15, 5, 15, 1.4, 2.2, 70, 30, 670, 100],
    ['Alatreon (Head)', 52500, 1.2, 88, 14, 10, 10, 14, 1, 2, 4, 150, 200, 2150, 100],
    ['Alatreon (Arm)',  52500, 1.2, 70, 22, 16, 16, 22, 2, 2, 4, 150, 200, 2150, 100],
    ['Anjanath',        18239, 1, 70, 0, 25, 10, 15, 5, 1.3, 2, 120, 90, 1920, 100],
    ['Azure Rathalos',  20800, 1.1, 70, 0, 15, 15, 20, 30, 1.4, 2.2, 120, 60, 720, 100],
    ['Banbaro (Horn)',  16240, 1.1, 73, 20, 0, 10, 0, 25, 1.25, 1.9, 40, 20, 440, 100],
    ['Banbaro (Leg)',   16240, 1.1, 66, 5, 0, 0, 0, 10, 1.25, 1.9, 40, 20, 440, 100],
    ['Barioth',         19200, 1.05, 73, 20, 10, 15, 0, 15, 1.35, 2.1, 70, 30, 670, 100],
    ['Barroth',         14210, 1, 73, 30, 0, 0, 10, 5, 1.25, 1.9, 40, 20, 440, 100],
    ['Bazelgeuse',      14210, 1.1, 85, 0, 15, 25, 20, 20, 1.25, 1.9, 40, 20, 440, 40],
    ['Behemoth',        35000, 1, 62, 5, 10, 5, 5, 15, 1.7, 1, 70, 100, 3000, 40],
    ['Extremoth',       51800, 1, 62, 5, 10, 5, 5, 15, 1.7, 1, 70, 100, 3000, 40],
    ['Beotodus',        12296, 1.05, 77, 30, 10, 15, 0, 0, 1.25, 1.9, 40, 20, 440, 100],
    ['Black Diablos (Chest)',22080, 1.05, 68, 0, 30, 10, 35, 0, 1.4, 2.2,  70, 30, 670, 100],
    ['Black Diablos (Head)',22080, 1.05, 57, 0, 25, 10, 30, 0, 1.4, 2.2,  70, 30, 670, 100],
    ['Blackveil',       27520, 1.05, 81, 20, 5, 10, 15, 25, 1.45, 2.3, 70, 30, 670, 100],
    ['Brachydios',      27520, 0.95, 66, 0, 15, 10, 20, 5, 1.35, 2.1, 120, 90, 1920, 100],
    ['Brute Tigrex',    21000, 1.1, 58, 0, 20, 18, 12, 5, 1.5, 2.4, 70, 30, 670, 100],
    ['Coral Pukei',     18300, 1.1, 70, 10, 0, 15, 30, 10, 1.3, 2, 70, 30, 670, 100],

    // B 19 to 40
    ['Deviljho',        8888,  1.1, 81, 20, 20, 25, 15, 25, 1.2, 1.2, 70, 30, 670, 40],
    ['Diablos (Chest)',     20800, 1.1, 73, 0, 30, 10, 35, 30, 1.35, 2.1, 70, 30, 670, 100],
    ['Diablos (Head)',      20800, 1.1, 58, 0, 20, 10, 25, 20, 1.35, 2.1, 70, 30, 670, 100],
    ['Dodogama (Rock 1)',   17100, 1.1, 70, 0, 20, 30, 20, 10, 1.25, 1.9, 120, 90, 1920, 100],
    ['Ebony Odogaron',      19200, 1, 73, 20, 25, 20, 20, 5, 1.4, 2.2, 70, 30, 670, 100],
    ['Frostfang Barioth',   22200, 1.05, 73, 20, 10, 15, 0, 15, 1.6, 3, 70, 30, 670, 100],
    ['Fulgur Anjanath (Uncharged)', 19136, 1, 70, 10, 15, 0, 25, 5, 1.4, 2.2, 120, 90, 1920, 100],
    ['Fulgur Anjanath (Charged)', 19136, 1, 66, 10, 25, 0, 15, 5, 1.4, 2.2, 120, 90, 1920, 100],
    ['Furious Jang (Head)', 21090, 1, 55, 0, 5, 0, 20, 0, 1.6, 3, 70, 30, 670, 100],
    ['Furious Jang (Arm)',  21090, 1, 53, 0, 5, 0, 10, 0, 1.6, 3, 70, 30, 670, 100],
    ['Glavenus',            19200, 1, 62, 0, 25, 10, 15, 15, 1.35, 2.1, 70, 30, 670, 100],
    ['Gold Rathian',        21000, 1.2, 71, 0, 20, 30, 10, 0, 1.5, 2.4, 120, 60, 720, 100],
    ['Great Girros',        13920, 1.1, 73, 15, 30, 0, 20, 10, 1.25, 1.9, 70, 30, 670, 100],
    ['Great Jagras',        12760, 1.1, 73, 30, 0, 20, 15, 10, 1.25, 1.9, 40, 20, 440, 100],
    ['Jyuratodus',          16240, 0.9, 68, 5, 0, 10, 5, 5, 1.25, 1.9, 120, 90, 1920, 100],
    ['Kirin',               21440, 1.1, 71, 25, 20, 0, 20, 5, 1.45, 2.3, 70, 30, 670, 100],
    ['Kulu-Ya-Ku',          15370, 1, 76, 15, 20, 15, 15, 15, 1.25, 1.9, 70, 30, 670, 100],
    ['Kulve (Molten)(Head/Arm)', 54000, 1.2, 85, 10, 15, 20, 0, 10, 2, 4, 120, 120, 1920, 200],
    ['Kulve (Released)(Head)', 54000, 1.2, 92, 0, 10, 0, 20, 15, 2, 4, 120, 120, 1920, 200],
    ['Kulve (Released)(Arm)', 54000, 1.2, 77, 0, 10, 0, 20, 15, 2, 4, 120, 120, 1920, 200],
    ['Kushala',         25472, 1.1, 66, 10, 5, 20, 0, 10, 1.45, 2.3, 40, 20, 440, 100],
    ['Lavasioth (Broken Head)',  25600, 0.9, 56, 5, 35, 15, 15, 5, 1.35, 3.1, 120, 90, 1920, 100],

    // C 41 to 55
    ['Legiana',         19200, 1.1, 69, 20, 15, 25, 0, 10, 1.35, 2.1, 70, 30, 670, 100],
    ['Lunastra (Head)',  30100, 1.2, 52, 0, 7, 7, 30, 25, 1.5, 2.4, 70, 30, 670, 100],
    ['Namielle (Hydrated)',25600, 0.9, 60, 40, 0, 0, 15, 30, 1.45, 2.3, 40, 20, 440, 100],
    ['Namielle (Dehydrated)',25600, 0.9, 75, 5, 0, 0, 5, 10, 1.45, 2.3, 40, 20, 440, 100],
    ['AT Namielle (Hydrated)',36000, 0.9, 60, 40, 0, 0, 15, 30, 2.5, 4.5, 40, 20, 440, 100],
    ['AT Namielle (Dehydrated)',36000, 0.9, 75, 5, 0, 0, 5, 10, 2.5, 4.5, 40, 20, 440, 100],
    ['Nargacuga', 20160, 1.1, 73, 20, 0, 30, 10, 10, 1.35, 2.1, 70, 30, 670, 100],
    ['Nergigante', 7110, 1, 58, 5, 10, 15, 15, 5, 1.3, 1.3, 70, 30, 670, 40],
    ['Nightshade Paolumu', 17995, 1, 66, 15, 25, 10, 10, 0, 1.3, 2, 70, 30, 670, 100],
    ['Odogaron', 19200, 1.1, 70, 10, 15, 20, 30, 0, 1.35, 2.1, 70, 30, 670, 100],
    ['Paolumu', 17995, 1, 66, 25, 0, 15, 10, 10, 1.3, 2, 40, 20, 440, 100],
    ['Pink Rathian', 18300, 1.1, 70, 0, 15, 20, 15, 30, 1.3, 2, 120, 60, 720, 100],
    ['Pukei-Pukei', 16240, 1.1, 70, 15, 0, 30, 20, 10, 1.25, 1.9, 70, 30, 670, 100],
    ['Radobaan', 18300, 1, 70, 20, 20, 20, 20, 30, 1.3, 2, 40, 20, 440, 100],
    ['Raging Brachy', 18300, 1.05, 70, 0, 10, 10, 25, 5, 1.6, 3, 120, 90, 1920, 100],

    // 56 to 67
    ['Rajang (Head)', 18900, 1, 60, 0, 5, 0, 30, 0, 1.5, 2.4, 70, 30, 670, 100],
    ['Rajang (Arm)', 18900, 1, 57, 0, 5, 0, 15, 0, 1.5, 2.4, 70, 30, 670, 100],
    ['Rathalos', 20800, 1.1, 70, 0, 15, 20, 15, 30, 1.35, 2.1, 120, 60, 720, 100],
    ['Rathian', 18300, 1.1, 70, 0, 15, 20, 15, 30, 1.3, 2, 120, 60, 720, 100],
    ['Ruiner', 23680, 1, 58, 5, 10, 10, 15, 15, 1.5, 2.3, 40, 20, 440, 100],
    ['Ruiner (Ode)', 27380, 1, 58, 5, 10, 10, 15, 15, 1.6, 3, 40, 20, 440, 100],
    ['Safi\'jiiva', 120000, 1.1, 45, 20, 20, 20, 20, 20, 1, 2.5, 250, 60, 800, 200],
    ['Savage Jho', 23040, 1.2, 75, 25, 20, 30, 15, 25, 1.45, 2.3, 70, 30, 670, 100],
    ['Scarred Garuga', 21120, 1.1, 70, 0, 30, 0, 10, 15, 1.5, 2.3, 40, 20, 440, 100],
    ['Seething Bazel (Leg)', 25600, 1.1, 51, 0, 5, 10, 15, 5, 1.45, 2.3, 120, 90, 1920, 23],
    ['Shara Ishvalda (Rock)', 37100, 1.05, 62, 5, 15, 0, 20, 5, 1.5, 2.4, 70, 30, 670, 200],
    ['Shara Ishvalda (Released)', 37100, 1.05, 77, 10, 20, 0, 25, 5, 1.5, 2.4, 70, 30, 670, 200],

    // D 68 to 84
    ['Shrieking Legiana', 19200, 1.1, 73, 15, 10, 15, 0, 10, 1.4, 2.2, 70, 30, 670, 100],
    ['Silver Rathalos', 21000, 1.2, 71, 0, 30, 20, 10, 0, 1.5, 2.4, 120, 60, 720, 100],
    ['Stygian Zinogre (Charged)', 21000, 1, 58, 20, 25, 30, 20, 0, 1.5, 2.4, 40, 20, 440, 100],
    ['Stygian Zinogre (Uncharged)', 21000, 1, 66, 10, 13, 15, 10, 0, 1.5, 2.4, 40, 20, 440, 100],
    ['Teostra', 27520, 1.2, 66, 0, 30, 10, 25, 10, 1.45, 2.3, 120, 90, 1920, 100],
    ['Tigrex', 19200, 1.1, 62, 0, 15, 20, 10, 5, 1.35, 2.1, 70, 30, 670, 100],
    ['Tobi-Kadachi', 16820, 1.1, 73, 20, 30, 0, 20, 15, 1.25, 1.9, 70, 30, 670, 100],
    ['Tzitzi-Ya-Ku', 15370, 1, 69, 15, 15, 20, 20, 15, 1.25, 1.9, 70, 30, 670, 100],
    ['Uragaan', 27520, 1.2, 71, 0, 20, 5, 15, 20, 1.35, 2.1, 40, 20, 440, 100],
    ['Vaal Hazak', 9494, 1.05, 77, 25, 5, 10, 15, 25, 1.3, 1.3, 120, 90, 1920, 40],
    ['Velkhana', 23040, 1.2, 73, 20, 0, 15, 0, 20, 1.45, 2.3, 40, 20, 440, 100],
    ['Viper Tobi', 17690, 1.1, 73, 10, 0, 25, 20, 10, 1.3, 2, 70, 30, 670, 100],
    ['Xeno', 15150, 0.95, 62, 20, 10, 10, 10, 20, 1.3, 1.3, 120, 90, 1920, 40],
    ['Yian Garuga', 22050, 1.1, 70, 0, 30, 0, 10, 15, 1.5, 2.4, 40, 20, 440, 100],
    ['Zinogre (Charged)', 23450, 1, 73, 10, 20, 0, 25, 10, 1.5, 2.4, 70, 30, 670, 100],
    ['Zinogre (Uncharged)', 23450, 1, 58, 5, 10, 0, 15, 5, 1.5, 2.4, 70, 30, 670, 100],

    // fatalis
    ['Fatalis (Head)', 66000, 1.1, 76, 17, 5, 5, 12, 25, 3, 5, 150, 200, 3000, 100],
    ['Fatalis (Chest)', 66000, 1.1, 53, 13, 5, 5, 10, 15, 3, 5, 150, 200, 3000, 100],
    ['Fatalis (Chest)(Critical)', 66000, 1.1, 72, 13, 5, 5, 10, 15, 3, 5, 150, 200, 3000, 100],
];

var monsterStats =
[
    'Name',
    'HP',
    'AgiMulti',
    'RawHZV',
    'FireHZV',
    'WaterHZV',
    'ThunderHZV',
    'IceHZV',
    'DragonHZV',
    'Qmulti1',
    'Qmulti2',
    'First',
    'Increase',
    'Cap',
    'Damage',
    'AgiUptime'
]

// element variables
var NoEle = 0;
var Fire = 1;
var Water = 2;
var Thunder = 3;
var Ice = 4;
var Dragon = 5;
var AllEle = 6;
var Blast = 7;

var agiUptime;

// array of sets for easier automation
var sets =
[
    // ['name', efr, efe, efs, offset]
    ['Safi Shatterspear (100% Coal)',   1002.19, 0, 7.33, 13, Blast],
    //['Safi Shatterspear (75% Coal)',    996.35, 0, 7.33, 14, Blast],
    ['Safi Shatterspear (50% Coal)',    988.57, 0, 7.33, 15, Blast],
    //['Safi Shatterspear (25% Coal)',    982.73, 0, 7.33, 16, Blast],
    ['Safi Shatterspear (0% Coal)',     978.84, 0, 7.33, 17, Blast],
    ['Lightbreak Press (100% Coal)',    990.51, 0, 7.33, 18, Blast],
   // ['Lightbreak Press (75% Coal)',     984.68, 0, 7.33, 19, Blast],
    ['Lightbreak Press (50% Coal)',     978.84, 0, 7.33, 20, Blast],
    //['Lightbreak Press (25% Coal)',     971.05, 0, 7.33, 21, Blast],
    ['Lightbreak Press (0% Coal)',      965.22, 0, 7.33, 22, Blast],
    ['Alatreon\'s Embrace',             969.11, 85, 0, 23, Dragon],
    ['Safi Raw Ele (100% Coal)',        1002.19, 26.25, 0, 24, AllEle],
    //['Safi Raw Ele (75% Coal)',         996.35, 25.31, 0, 25, AllEle],
    ['Safi Raw Ele (50% Coal)',         988.57, 24.38, 0, 26, AllEle],
    //['Safi Raw Ele (25% Coal)',         982.73, 23.44, 0, 27, AllEle],
    ['Safi Raw Ele (0% Coal)',          974.95, 22.5, 0, 28, AllEle],
    ['KPara',                           911.06, 0, 0, 29, NoEle],
    ['KWater',                          872.26, 97.81, 0, 30, Water],
    ['Safihemoth',                      945.76, 50, 0, 31, AllEle],
    //['Brakulve (100% Coal)',            980.78, 35, 0, 32, AllEle],
   //['Brakulve (75% Coal)',             973, 34.06, 0, 33, AllEle],
    //['Brakulve (50% Coal)',             967.16, 33.13, 0, 34, AllEle],
    //['Brakulve (25% Coal)',             959.38, 32.19, 0, 35, AllEle],
    //['Brakulve (0% Coal)',              953.54, 31.25, 0, 36, AllEle],
    ['Fatalis (100% Coal)',                 1117, 30, 0, 0, Dragon],
    ['Fatalis (50% Coal)',                 1097.54, 24.38, 0, 0, Dragon],
    ['Fatalis (0% Coal)',                 1076.14, 18.75, 0, 0, Dragon],
    //['Fatalis mt',                      1103.38, 22.5, 0, 0, Dragon],
]

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

// delete all sheets but the calculation ones
function wipe()
{
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // 'misc compare data' sheet
    var sheets = activeSpreadsheet.getSheets(); // sheets array
    for (var i = 0; i < sheets.length; i++)
    {
        if (sheets[i].getName() != 'Main' && sheets[i].getName() != 'Data') // check that the sheet isn't necessary
        {
            activeSpreadsheet.deleteSheet(sheets[i]);
        }   
    }
};

// calculate various damage function
function calc(weap, loop, monsterOptions, outputRange, main)
{
    // load monster options into variables
    var name = monsterOptions[0];
    var hp = monsterOptions[1];
    var enrage = monsterOptions[2];
    var rawHZV = monsterOptions[3] / 100;
    var fireHZV = monsterOptions[4];
    var waterHZV = monsterOptions[5];
    var thunderHZV = monsterOptions[6];
    var iceHZV = monsterOptions[7];
    var dragonHZV = monsterOptions[8];
    var qMulti1 = monsterOptions[9];
    var qMulti2 = monsterOptions[10];
    var initial = monsterOptions[11];
    var buildup = monsterOptions[12];
    var _cap = monsterOptions[13];
    var damage = monsterOptions[14];

    var sheet = main; // original input sheet
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // 'misc compare data' sheet
    var newSheet = activeSpreadsheet.getSheetByName(name); // output sheet

    // load values from the info array, 
    // which was loaded with efr, efe, etc from 
    // original input sheet.
    var name = weap[0];
    var efr = weap[1];
    var efe = weap[2];
    var efs = weap[3];
    var element = weap[5];
    var eleHZV = 0;

    // load weapon options into variables
    // that we will be using later in the calcs
    var _efr = efr / 100;
    var _efe = efe;
    var _efs = efs;
    var blast = false;
     if (element == Blast)
        blast = true; // set flag that weapon is using blast

    switch(element)
    {
        case Fire:
            eleHZV = fireHZV;
            break;
        case Water:
            eleHZV = waterHZV;
            break;
        case Thunder:
            eleHZV = thunderHZV;
            break;
        case Ice:
            eleHZV = iceHZV;
            break;
        case Dragon:
            eleHZV = dragonHZV;
            break;
        case AllEle:
            eleHZV = Math.max(fireHZV, waterHZV, thunderHZV, iceHZV, dragonHZV);
            break;
        default:
        eleHZV = 0;
    };

    eleHZV = eleHZV / 100;

    // create variables to be used in the loop
    var blastValue = 0; // current blast value    
    var procs = 0; // blast proc counter
    var threshold = initial * qMulti1; // threshold is blast required for a proc
    var increase = buildup * qMulti2; // increase is the additional amount needed for proc on top of previous threshold
    var cap = _cap * qMulti2; // cap for threshold, if it goes higher always set it to the cap
    var monsterHP = hp; // set current hp for monster to the max hp loaded from options
    var _agiMulti = ((enrage - 1) * agiUptime) + 1; // average engraged multiplier
    var procDmg = damage * 3; // blast damage

    var loop = tornado;
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

    // while monster is alive,  deal damage based 
    // on the next hit
    while (monsterHP > 0)
    {        
        // if the blast flag was set
        // to true, apply blast calcs
        //console.log({message: 'dmg calc', initialData: [mvs[hit], _efr , rawHZV , _agiMulti]});
        if (blast)
        {
            var a = (mvs[hit] * _efr * rawHZV * _agiMulti); // raw damage calculation for blast            
            console.log(mvs[hit], _efr, rawHZV, _agiMulti, a);
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
    //newSheet.getRange(outputRange).setBackgroundObject(fillColour); // set background colour to the colour loaded earlier
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

function automate(start, end)
{
    // empty sheet
    //wipe();

    var sheet = SpreadsheetApp.getActiveSheet(); // current input sheet
    agiUptime = sheet.getRange('D2').getValue(); // 70%

    // iterate the monsters array, and calculate :poggies:
    for (var i = start; i <= end; i++)
    {
    // ['monster name', hp, enrage, raw hzv, fire, water, thunder, ice, dragon, ailment mod1, ailment mod 2, initial, buildup, cap, damage]
        calcAndOutput(monsters[i]);
    }
}

function simA() { automate(0, 18); }
function simB() { automate(19, 40); }
function simC() { automate(41, 55); }
function simD() { automate(56, 67); }
function simE() { automate(68, 83); }
function simF() { automate(84, 86); }

function calcAndOutput(monsterOptions)
{

    var name = monsterOptions[0];
    var hp = monsterOptions[1];
    var enrage = monsterOptions[2];
    var rawHZV = monsterOptions[3];
    var fireHZV = monsterOptions[4];
    var waterHZV = monsterOptions[5];
    var thunderHZV = monsterOptions[6];
    var iceHZV = monsterOptions[7];
    var dragonHZV = monsterOptions[8];
    var qMulti1 = monsterOptions[9];
    var qMulti2 = monsterOptions[10];
    var initial = monsterOptions[11];
    var buildup = monsterOptions[12];
    var cap = monsterOptions[13];
    var damage = monsterOptions[14];

    var outputColumn = 'D';
    var outputColumn2 = 'H';

     // google sheet stuff
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet(); // 'misc compare data' sheet
    var sheet = SpreadsheetApp.getActiveSheet(); // current input sheet
    var newSheet = createNewSheet(name);

    newSheet.getRange(outputColumn + '1:' + outputColumn2 + '1').setValues([['Set', 'Raw Damage', 'Ele/Blast Damage', 'Hits to Kill', "Loop"]]);
    newSheet.getRange(outputColumn + '1:' + outputColumn2 + '1').setFontWeights([['bold', 'bold', 'bold', 'bold', 'bold']]);
    newSheet.getRange(outputColumn + '1:' + outputColumn2 + '1').setHorizontalAlignments([['center', 'center', 'center', 'center', 'center' ]])
    //newSheet.getRange('A1:A' + monsterOptions.length).setValues(monsterOptions);
    // hmmm this isn't working for some reason....

    monsterOptions.push(agiUptime);
    for (var b = 0; b < monsterOptions.length; b++)
    {
        newSheet.getRange('A' + (b + 1)).setValue(monsterStats[b] + ':');
        newSheet.getRange('A' + (b + 1)).setFontWeight('bold');
        newSheet.getRange('B' + (b + 1)).setValue(monsterOptions[b]);
    }
 
    // headers are now set. time to iterate the main sets and calculate their damage
    var outputFirst = 2; // the first output cell starts at 'B2'
    // loop through all 24 sets for the main tornado loop section
    for (x = 0; x < sets.length; x++)
    {
        calc(sets[x], tornado, monsterOptions, ((outputColumn + outputFirst) + ':' + (outputColumn2 + outputFirst)), sheet)   
        outputFirst++; // increment to ouput to 1 line below
    }

    // filter will let you sort ascending or descending based on columns
    var filter = newSheet.getRange(outputColumn + '1:' + outputColumn2 + (sets.length + 1)).createFilter();
    // this loop will set columns to the widest size they need to be,
    // while also remaining readable
    for (var x = 1; x < 25; x++)
    {
        newSheet.autoResizeColumn(x); 
    }
}