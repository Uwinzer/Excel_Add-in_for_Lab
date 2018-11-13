/* 
 * Excel operating library
*/

"use strict";

(function () {
    // Office.initialize = function (reason) {
    //     $(document).ready(function () {
    //     });
    // };

    // register name space
    var rGenerator = window.rGenerator || {};
    var ExcelOP = rGenerator.ExcelOP || {};

    ExcelOP.Filter = function (sheet, regex, data_headers) {
        // @description: filter useful data to a new sheet
        // @sheet {string} sheet that filter applys for
        // @regex {Regex} regular expression for filtring
        // @data_headers {array} data header names
        Excel.run(function (context) {
            const USED_RANGE = context.workbook.worksheets.getItem(sheet).getUsedRange();
            USED_RANGE.load('values');
            return context.sync()
                .then(function () {
                    // convert to JSON
                    var data = {};
                    // initialize
                    for (var i in data_headers) {
                        data[data_headers[i]] = [];
                    }
                    // match string and valid it to see if it is in the list
                    for (var i = 0; i < USED_RANGE.values.length; i++) {
                        if (regex.test(USED_RANGE.values[i][0]) && (data_headers.indexOf(USED_RANGE.values[i][0].match(regex)[0]) > -1)) {
                            data[(USED_RANGE.values[i][0].match(regex))[0]].push(USED_RANGE.values[i][3]);
                        }
                    }
                    return context.sync()
                        .then(function () {
                            console.log(data);
                        });
                });
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    };

    rGenerator.ExcelOP = ExcelOP;
    window.rGenerator = rGenerator;
})();