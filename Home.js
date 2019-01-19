
(function () {
    "use strict";

    var messageBanner;

    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();



            // TODO1: Assign event handler for insert-image button.
            // TODO4: Assign event handler for insert-text button.
            // TODO6: Assign event handler for get-slide-metadata button.
            // TODO8: Assign event handlers for the four navigation buttons.
        });
    };

   
    
   

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();