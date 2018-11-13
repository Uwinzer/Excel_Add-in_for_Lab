(function () {
    // register name space
    var RGNameSpace = window.RGNameSpace || {};
    RGNameSpace.rgDialog = {};

    RGNameSpace.rgDialog.openDialog = function () {
        Office.context.ui.displayDialogAsync(window.location.origin + "/Dialog.html",
        { height: 50, width: 50 }, dialogCallback);
        console.log(window.location.origin);
    };
    
    function dialogCallback (result) {
        console.log("called callback");
        // console.log(result);
    }
    window.RGNameSpace = RGNameSpace;
})();
