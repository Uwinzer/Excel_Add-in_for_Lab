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
        var tester = $('#tester').get(0);
        Plotly.plot( tester, [{
        x: [1, 2, 3, 4, 5],
        y: [1, 2, 4, 8, 16] }], {
        margin: { t: 0 } } );
        Office.context.ui.messageParent("dialog message sent");
    };
    
})();
