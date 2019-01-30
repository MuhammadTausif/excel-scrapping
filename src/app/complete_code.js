// This is externall library you must have to install via npm.
var XLSX = require('xlsx');

// The code
var fs = require('fs');
var json_data = []

// data1.xlsx is the input data. You can take it from input.
var workbook = XLSX.readFile('data (4).xlsx', {
    cellStyles: true
});


var sheet_name_list = workbook.SheetNames;

var sheet2arr = function (sheet) {
    var result = [];
    var row;
    var rowNum;
    var colNum;
    var range = XLSX.utils.decode_range(sheet['!ref']);
    var companiesNameArray = [];
    var blueTotalTestFlag = false;
    var blueTotalRow = 0;
    var blueTotalEntry;
    var firstCellFlage = false;
    var firstCellReminder = 0;

    for (rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
        row = [];
        var blueColTestFlag = false;
        var temp_colAFound;
        for (colNum = range.s.c; colNum <= range.e.c; colNum++) {
            var nextCell = sheet[
                XLSX.utils.encode_cell({
                    r: rowNum,
                    c: colNum
                })
            ];

            if (typeof nextCell === 'undefined') {
                row.push(void 0);
            } else {

                // // Checing the first company
                // if(rowNum == 8 && firstCellFlage == false){
                //     firstCellFlage = true;
                //     firstCellReminder = colNum % 4;
                // }

                // Writing names
                if (colNum == 0) {
                    temp_colAFound = nextCell.w;
                    var patrn = /\b2.[0-9]+ /;
                    if (typeof nextCell.s.fgColor != "undefined") {
                        if (nextCell.s.fgColor.rgb == '93CDDD') {
                            // if (patrn.test(temp_colAFound)) {
                            blueColTestFlag = true;
                            companiesNameArrayForIteration = Array.from(companiesNameArray);
                        }
                    }

                    var patrn3 = /\b2 C/;
                    if (typeof nextCell.s.fgColor != "undefined") {
                        if (nextCell.s.fgColor.rgb == '93CDDD') {

                            // if (patrn3.test(temp_colAFound)) {
                            blueTotalTestFlag = true;
                            blueTotalRow = rowNum + 1;
                            blueTotalEntry = temp_colAFound;
                            companiesNameArrayForIteration = Array.from(companiesNameArray);
                        }
                    }
                }
                if (blueTotalTestFlag) {
                    if (rowNum == blueTotalRow) {
                        // companiesNameArrayForIteration = Array.from(companiesNameArray);
                        //93CDDD Blue
                        //FFC000 Yellow 
                        if (typeof nextCell.s.fgColor != "undefined") {
                            if (nextCell.s.fgColor.rgb == 'FFC000') {
                                // if (colNum % 4 == firstCellReminder + 1 && colNum != 0) {
                                var comp = companiesNameArrayForIteration.shift();
                                var obj_names = {
                                    name: blueTotalEntry,
                                    price: nextCell.w
                                }
                                var index = 0;
                                var temp_index = 0;
                                json_data.forEach(element => {
                                    if (element.Name == comp) {
                                        temp_index = index;
                                    }
                                    index = index + 1;
                                });

                                json_data[temp_index]['Data'].push(obj_names);
                            }
                        }
                    }
                }
                if (blueColTestFlag) {
                    if (colNum % 4 == (firstCellReminder + 1) && colNum != 0) {
                        var comp = companiesNameArrayForIteration.shift();
                        var obj_names = {
                            name: temp_colAFound,
                            price: nextCell.w
                        }
                        var index = 0;
                        var temp_index = 0;
                        json_data.forEach(element => {
                            if (element.Name == comp) {
                                temp_index = index;
                            }
                            index = index + 1;
                        });

                        json_data[temp_index]['Data'].push(obj_names);
                    }
                }

                // Name writing
                if (rowNum == 8) {
                    //93CDDD Blue
                    //FFC000 Yellow 
                    if (nextCell.s.fgColor.rgb == '93CDDD') {
                        // if (colNum % 4 == firstCellReminder) {
                        companiesNameArray.push(nextCell.w);
                        var temp_comp_name = {
                            "Name": nextCell.w,
                            "Data": []
                        }
                        json_data.push(temp_comp_name);
                    }
                }

                row.push(nextCell.w);
            }
        }
        result.push(row);
    }
    return result;
};

sheet2arr(workbook.Sheets[sheet_name_list[0]]);

var first_sheet = workbook.Sheets[sheet_name_list[0]];
var desired_cell = first_sheet['A2'];
var color = desired_cell.s.fgColor.rgb;
// var color = abc

fs.writeFile('complete_json.json', JSON.stringify(json_data, null, 100));