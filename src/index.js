/* 
 * Home page
*/

"use strict";

(function () {
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#Filter').click(Filter);
            $('#Open_a_Dialog').click(openDialog);
            $('#test').click(test);
        });
    };

    var filtered_data = {};
    var formated_data = {};
    function Filter () {
        var sheet = "Results";
        var regex = /([a-zA-Z ]{1,})(?=\_[0-9]{1,})/g;
        var data_that_i_need = [
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
        // use promise to handle async operation
        rGenerator.ExcelOP.Filter(sheet, data_that_i_need, regex).then(function (data) {
            filtered_data = data;
            formated_data = formated_data = rGenerator.ExcelOP.Formater(filtered_data, 8, 2);
            localStorage.setItem("formated_data", JSON.stringify(formated_data));
            console.log(filtered_data);
            console.log(formated_data);
        });
    }

// --------------test code----------------------------------
    var dialog;
    function openDialog () {
        Office.context.ui.displayDialogAsync(window.location.origin + "/Dialog.html",
        { height: 60, width: 80 }, 
        dialogCallback);
        // console.log(window.location.origin);
    }
    
    function dialogCallback (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, messageHandler);
        console.log("called callback");
        // console.log(result);
    }

    function messageHandler (arg) {
        // var z = data["Copper Loss"];

        // var data_z1 = {z: z, type: 'surface'};
        
        // Plotly.newPlot('tester', [data_z1]);
        console.log(arg.message);
    }

    async function test () {
    }

})();