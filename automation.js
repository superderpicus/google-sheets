function run()
{
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = SpreadsheetApp.getActiveSheet();
    var letters = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];

    // check if sheet is correct or not, otherwise something went wrong
    if (sheet.getName() != 'Automation')
    	return;    

    // variables
    var values = sheet.getRange('C2:C12').getValues(); 
    var EFR = values[0] / 100;
    var EFE = values[1];
    var combo = values[9];
    var mode = "dps";
    var columnWidth = 65;
    var tornadoLoopTime = 4.22; // seconds
    var rawHZVTable = [];
    var eleHZVTable = [];

    // check if new sheet already exists, delete it if it does
    var name = sheet.getRange('C10').getValue();
    var newSheet = activeSpreadsheet.getSheetByName(name);
    if (newSheet != null)
    {
        activeSpreadsheet.deleteSheet(newSheet);
    }

    newSheet = activeSpreadsheet.insertSheet(); // create new spreadsheet
    newSheet.setName(name); // set the name of the tab
    newSheet.setTabColor(random_hex_color_code()); // generate random hex colour for the tab

    newSheet.getRange('A1:A4').setValues([["EFR:"], ["EFE:"], ["COMBO:"],["DPS:"]])
    newSheet.getRange('A1:A4').setFontWeight("bold");

    newSheet.getRange('D1').setValue('VS Safi:');
    newSheet.getRange('E1').setValue("=1 - (G7/'Safis Shatterspear'!G7)");
    newSheet.getRange('E1').setNumberFormat('##.#%')

    newSheet.getRange('B1:B4').setValues([[EFR * 100], [EFE], [combo],[mode]])
    newSheet.getRange('B1:B3').setHorizontalAlignment("right");
    newSheet.getRange('B4').setHorizontalAlignment("center")
    newSheet.getRange('B4').insertCheckboxes();
    
    // RAW HZV HEADERS
    var rawHZVCount;  
    if (values[4] == 0 || EFR == 0)
    {
    	rawHZVCount = 1;
    	var range = newSheet.getRange('C6');
        range.setValue(values[2] + '%');
        range.setFontWeight("bold");
        range.setFontSize(12);
        rawHZVTable.push(values[2])
    }
    else
    {
    	var s = parseInt(values[2], 10); // start raw hzv
    	var e = parseInt(values[3], 10); // end raw hzv
    	var i = parseInt(values[4], 10); // raw decrement

    	var count = 0;
    	for (var x = s; x >= e; x -= i)
    	{
    		count = count + 1;
    		var range =	newSheet.getRange(letters[count + 1] + '6');
    		range.setValue(x + '%');
    		//range.setFontWeight("bold");

            //range.setHorizontalAlignment("center");
    		//range.setFontSize(12);
    		rawHZVTable.push(x);
    	}																	
    	rawHZVCount = count;

        var rawHeaderRange = newSheet.getRange(letters[0] + '6:' + letters[count + 1] + '6')
        rawHeaderRange.setHorizontalAlignment("center");
        rawHeaderRange.setFontSize(12);
        rawHeaderRange.setFontWeight("bold");

    }

    // ELE HZV HEADERS
    var eleHZVCount;  
    if (values[7] == 0 || EFE == 0)
    {
    	eleHZVCount = 1;   

        var range = newSheet.getRange('B7');
        range.setValue('0%');
        range.setFontWeight("bold");
        range.setFontSize(12);
        eleHZVTable.push(0);
    }
    else
    {
    	var s = parseInt(values[5], 10); // start raw hzv
    	var e = parseInt(values[6], 10); // end raw hzv
    	var i = parseInt(values[7], 10); // raw decrement

    	var count = 0;
    	for (var x = s; x >= e; x -= i)
    	{
    		count = count + 1;
    		var range =	newSheet.getRange('B' + (count + 6));
    		range.setValue(x + '%');

            //range.setHorizontalAlignment("center");
    		//range.setFontWeight("bold");
    		//range.setFontSize(12);
    		eleHZVTable.push(x);
    	}
    	eleHZVCount = count;

        var eleHeaderRange = newSheet.getRange('B7:B' + count + 6)
        eleHeaderRange.setHorizontalAlignment("right");
        eleHeaderRange.setFontSize(12);
        eleHeaderRange.setFontWeight("bold");

    }

    // headers are set, time to actually do stuff
    for (var x = 0; x < rawHZVCount; x++)
    {
    	for (var y = 0; y < eleHZVCount; y++)
    	{
    		var a = rawHZVTable[x] / 100;
    		var b = eleHZVTable[y] / 100;

    		var val = calc(EFR, EFE, a, b, mode);
    		var range = newSheet.getRange(letters[x + 2] + (y + 7))
    		range.setFontSize(12);
    		range.setNumberFormat("0.##");

            range.setHorizontalAlignment("center");
    		//range.setValue(val);
            range.setValue('=IF($B$4 = TRUE,' + val + ', ' + val * tornadoLoopTime + ')')
    	}
    }
    newSheet.setColumnWidths(3, 24, columnWidth);
}
var x = false;
function calc(efr, efe, rawHZV, eleHZV, mode)
{
	// calc the dmg ! oh boy !
	var time = 0;
	var mv = 1;
	var eleMod = 2;

	var tornado =   [2.17, 66, 1.6]
	var widesweep = [1.3,  40, 1.6]
	var thrust =    [0.75, 23, 1.6]

	var tornado_dps =   (Math.round(efr * tornado[mv] *   rawHZV) + (efe * eleHZV * tornado[eleMod]));
	var widesweep_dps = (Math.round(efr * widesweep[mv] * rawHZV) + (efe * eleHZV * widesweep[eleMod]));
	var thrust_dps =    (Math.round(efr * thrust[mv] *    rawHZV) + (efe * eleHZV * thrust[eleMod]));

	if (mode == 'dps')
	{
		var loop_dps = (tornado_dps + widesweep_dps + thrust_dps) / (tornado[time] + widesweep[time] + thrust[time]);
		return loop_dps;
	}
	else if(mode == 'dmg')
	{
		var dmg = (tornado_dps + widesweep_dps + thrust_dps);
		return dmg;
	}
}

function del()
{
    var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = SpreadsheetApp.getActiveSheet();
    var name = sheet.getRange('C10').getValue();
    var newSheet = activeSpreadsheet.getSheetByName(name);
    if (newSheet != null)
    {
        activeSpreadsheet.deleteSheet(newSheet);
    }
}

function random_hex_color_code()
{
  var n = (Math.random() * 0xfffff * 1000000).toString(16);
  return '#' + n.slice(0, 6);
};
