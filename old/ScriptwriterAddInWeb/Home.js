'use strict';

(function () {

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            // The document is ready
            // Use this to check whether the API is supported in the Word client.
            if (Office.context.requirements.isSetSupported('WordApi', '1.1')) {
                // Do something that is only available via the new APIs
                $('#loadCharacterNames').click(loadCharacterNames);
                //$('#emerson').click(insertEmersonQuoteAtSelection);
                //$('#checkhov').click(insertChekhovQuoteAtTheBeginning);
                //$('#proverb').click(insertChineseProverbAtTheEnd);
                $('#supportedVersion').html('This code is using Word 2016 or later.');
            }
            else {
                // Just letting you know that this code will not work with your version of Word.
                $('#supportedVersion').html('This code requires Word 2016 or later.');
            }
        });
    });

   //async function setStyleForSelection(selection, styleName) {
   //     Word.run(function (context) {
   //         // Create a proxy object for the document.
   //         var thisDocument = context.document;
   //         selection.style = styleName;
   //         thisDocument.body.load(['style']);
   //         await context.sync();
   //         //thisDocument.body.style = "";
   //     });
   // }

    function loadCharacterNames() {
        Word.run(function (context) {
             //Create a proxy object for the document.
            thisDocument = context.document;
            context.sync();
            thisDocument.body.load(['template']);
            let p = load(propertyNames);
            let t = load("template");
             //Queue a command to get the current selection.
             //Create a proxy range object for the selection.
            var range = thisDocument.getSelection();

             //Queue a command to replace the selected text.
            range.insertText('Loading Character Names List.\n', Word.InsertLocation.replace);

             //Synchronize the document state by executing the queued commands,
             //and return a promise to indicate task completion.
            return context.sync().then(function () {
                console.log('Loading Character Names.');
            });
        })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
    }
 });