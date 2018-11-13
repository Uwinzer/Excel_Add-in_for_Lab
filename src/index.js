(function () {
    "use strict";
    Office.initialize = function (reason) {
            $(document).ready(function () {
                $('#Filter').click(Filter);
                $('#NEW1').click(new1);
                $('#NEW2').click(RGNameSpace.rgDialog.openDialog);
            });
    };

    function Filter () {
        Excel.run(function (context) {
            // filter useful data to a new sheet
            const USED_RANGE = context.workbook.worksheets.getItem("Results").getUsedRange();
            // const NEW_SHEET = context.workbook.worksheets.add("New");
            const REGEX = /([a-zA-Z ]{1,})(?=\_[0-9]{1,})/g;
            const DATA_THAT_I_NEED = [
                "Copper Loss",
                "Rotational Loss",
                "Inverter Loss",
                "Motor Loss",
                "System Loss",
                "Total Loss",
                "Calculated System Efficiency",
                "Calculated Motor Efficiency",
                "Calculated Inverter Efficiency"
            ];
            USED_RANGE.load('values');
            return context.sync()
                .then(function () {
                    // convert to JSON
                    var data = {};
                    // initialize
                    for (var i in DATA_THAT_I_NEED) {
                        data[DATA_THAT_I_NEED[i]] = [];
                    }
                    for (var i = 0; i < USED_RANGE.values.length; i++) {
                        // match string and valid it to see if it is in the list
                        if (REGEX.test(USED_RANGE.values[i][0]) && (DATA_THAT_I_NEED.indexOf(USED_RANGE.values[i][0].match(REGEX)[0]) > -1)) {
                            data[(USED_RANGE.values[i][0].match(REGEX))[0]].push(USED_RANGE.values[i][3]);
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
    }

    function new1 () {
        Excel.run(function (context) {
            return context.sync().then(function () {
                
            });
        }).catch(function (error) {
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function new2 () {
        Excel.run(function (context) {

        }).catch(function (error) {
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

})();