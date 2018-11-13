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
        Office.context.ui.messageParent("message sent");
    };
    
})();
