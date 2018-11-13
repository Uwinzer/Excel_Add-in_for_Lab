/* 
 * Home page
*/

"use strict";

(function () {
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#Filter').click(Filter);
            $('#NEW1').click(new1);
            $('#NEW2').click(openDialog);
        });
    };

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
        window.rGenerator.ExcelOP.Filter(sheet, regex, data_that_i_need);
        return;
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

    var dialog;
    function openDialog () {
        Office.context.ui.displayDialogAsync(window.location.origin + "/Dialog.html",
        { height: 50, width: 50 }, 
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
        console.log(arg.message);
    }
})();