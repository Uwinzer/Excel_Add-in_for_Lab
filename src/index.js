(function() {
    Office.initialize = function (reason) {
            $(document).ready(function () {
                $('#Filter').click(Filter);
                $('#NEW1').click(new1);
                $('#NEW2').click(new2);
            });
    };

    function Filter() {
        Excel.run(function (context) {
            // get a temp sheet to process data
            const used_range = context.workbook.worksheets.getItem("Results").getUsedRange();
            // const new_sheet = context.workbook.worksheets.add("New");

            used_range.load('values');
            return context.sync()
                .then(function() {
                    var regex = /([a-zA-Z ]{1,})(?=\_[0-9]{1,})/g;
                    var myjson = [];
                    for (var i = 0; i < used_range.values.length; i++) {
                        if (regex.test(used_range.values[i][0])) {
                            var json = {};
                            json[used_range.values[i][0].match(regex)] = used_range.values[i][3];
                            myjson.push(json);
                        }
                    }
                    console.log(myjson);
                });
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
     }

    function new1() {
        Excel.run(function (context) {
            
        })
        .catch(function (error) {
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function new2() {
        Excel.run(function (context) {

        })
        .catch(function (error) {
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

})();



