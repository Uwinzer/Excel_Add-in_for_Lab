/* 
 * Dialog page
*/

"use strict";

(function () {
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#Show').click(show);
        });
    };
    // register name space
    // var Dialog = window.rGenerator.Dialog || {};
    // window.rGenerator.Dialog = Dialog;

    function show () {
        var data = JSON.parse(localStorage.getItem("fuckingdata"));
        var z = data["Copper Loss"];

        var data_z1 = {z: z, type: 'surface'};
        
        Plotly.newPlot('tester', [data_z1]);
        console.log("fuck");
        console.log(data);
        console.log(z);

        Office.context.ui.messageParent("dialog button clicked");
    };
    
})();
