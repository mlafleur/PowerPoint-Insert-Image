/// <reference path="/Scripts/FabricUI/MessageBanner.js" />

(function () {
    "use strict";

    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();

            $('#insert-image').click(insertImage);
            $('#copy-image').click(copyImage);
        });
    };

    // Inserts an image into the current slide
    function insertImage() {
        var base64Image = getBase64Image();

        // Insert base64 encoded image
        Office.context.document.setSelectedDataAsync(base64Image,
            {
                coercionType: "image", 
                imageLeft: 0, 
                imageTop: 0
            },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {

                    // If we reach this point than we ran into an error 
                    // inserting the image. Here we will fall back to using
                    // the local clip-board. 
                    copyImage();                    
                }
                showNotification("Complete", "The image has been inserted");
            });
    }

    // Takes the image and adds it to the local clipboard
    function copyImage() {
        var image = document.createElement("img");
        image.id = 'clipboardImage';
        $(image).css("display", "none");
        $("body").prepend(image);
        $(image).attr("src", '../../Images/Bing_logo_(2016).png');
        $(image).css("width", '100%');
        $(image).css("height", 'auto');
        image.contentEditable = "true";
        var controlRange;
        if (document.body.createControlRange) {
            controlRange = document.body.createControlRange();
            controlRange.addElement(image);
            controlRange.execCommand("Copy");
            showNotification("Copied", "The image has been copied to your clipboard");
        }
        image.contentEditable = "false";
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }

    // Helper function to simulate a base64 encoded image (would normally be done server-side)
    function getBase64Image() {
        return "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";
    }
})();