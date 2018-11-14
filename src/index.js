/* 
 * Home page
*/

"use strict";

(function () {
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#Filter').click(Filter);
            $('#NEW1').click(compute);
            $('#NEW2').click(openDialog);
        });
    };

    async function Filter () {
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
        var data = {};
        data = await rGenerator.ExcelOP.Filter(sheet, regex, data_that_i_need);
        
        // rGenerator.ExcelOP.Filter(sheet, regex, data_that_i_need).then(function (arg) {return data = arg;});
        var fdata = rGenerator.ExcelOP.Formater(data, 8, 2);
        window.fdata = fdata;

        console.log(data);
        console.log(fdata);
        return;
    }

// --------------test code----------------------------------
    var dialog;
    function openDialog () {
        localStorage.setItem("fuckingdata", JSON.stringify(fdata));
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

    function getSomething() {
        var r = 0;
        var data = [];
        return new Promise(function(resolve) {
            data.push(1);
            data.push(666);
            data.push("fuck");
            setTimeout(function() {
                r = 2;
                resolve(r);
            }, 2000);
        }).then(function () {
            return new Promise(function (resolve, reject) {
                resolve(data);
            });
        });
    }
    
    async function compute() {
        // var x = await getSomething();
        // console.log(x * 2);
        console.log(await getSomething());
    }

})();