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
        var formated_data = JSON.parse(localStorage.getItem("formated_data"));
        var z_arg = {};
        var z_array = [];
        var count = 0;
        for (var i in formated_data) {
            if (count == 0) {
                z_arg = {z: formated_data[i], type: 'surface'};
                count++;
            }
            else {
                z_arg = {z: formated_data[i], showscale: false, type: 'surface'};
                count++;
            }
            z_array.push(z_arg);
        }
        
        Plotly.newPlot('tester', z_array);
        console.log(count);

        Office.context.ui.messageParent("dialog button clicked");
    };
    
})();
