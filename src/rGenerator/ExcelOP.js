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

    /*
     * filter useful data to a new sheet
     * NOTE: FUNCTIONS MUST BE "async" TYPE WHERE THIS FILTER IS USED
     * @param {string} sheet: sheet that filter applys for
     * @param {Regex} regex: regular expression for filtring
     * @param {array} data_headers: data header names
     * @return {Promise} resolved values with keys in data_headers
    */
    ExcelOP.Filter = function (sheet, regex, data_headers) {
        var filtered_data = {};
        // initialize
        for (var i in data_headers) {
            filtered_data[data_headers[i]] = [];
        }
        return Excel.run(function (context) {
            const USED_RANGE = context.workbook.worksheets.getItem(sheet).getUsedRange();
            USED_RANGE.load('values');
            return context.sync()
                .then(function () {
                    // convert to JSON
                    // match string and valid it to see if it is in the list
                    for (var i = 0; i < USED_RANGE.values.length; i++) {
                        if (regex.test(USED_RANGE.values[i][0]) && (data_headers.indexOf(USED_RANGE.values[i][0].match(regex)[0]) > -1)) {
                            filtered_data[USED_RANGE.values[i][0].match(regex)[0]].push(USED_RANGE.values[i][3]);
                        }
                    }
                    return context.sync()
                        .then(function () {
                            // console.log(filtered_data);
                        });
                });
        }).then(function () {
            return new Promise(function (resolve, reject) {
                resolve(filtered_data);
            });
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    };

    /*
     * format data from Filter()
     * @param {JSON} filtered_data: returned value from Filter()
     * @param {number} axis_scale: scale axis (row) for ploting
     * @param {number} axis_rpm: rpm axis for (column) ploting
     * @return {JSON} in Plotly.js supported format

    */
    ExcelOP.Formater = function (filtered_data, axis_scale, axis_rpm) {
        var formated_data = {};
        var temp = [];
        for (var name in filtered_data) {
            formated_data[name] = [];
            // scan row by row
            for (var i = 0; i < axis_scale; i++) {
                temp = [];
                for (var j = 0; j < axis_rpm; j++) {
                    temp.push(filtered_data[name][i + j * axis_scale]);
                }
                formated_data[name].push(temp);
            }
        }
        return formated_data;
    };

    rGenerator.ExcelOP = ExcelOP;
    window.rGenerator = rGenerator;
})();