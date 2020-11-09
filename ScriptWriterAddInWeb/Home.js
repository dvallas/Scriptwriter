"use strict";

(function () {
    var messageBanner;
    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it
            //var element = document.querySelector('.MessageBanner');
            //messageBanner = new components.MessageBanner(element);
            //messageBanner.hideBanner();

            // If not using Word 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
                $("#template-description").text("This sample displays the selected text.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selected text");
                return;
            }
            //Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
            //Office.context.document.settings.saveAsync();

            //loadSampleData();
            //setTemplate();
            //setStyles();
            //setDataPreloads();
            //setMacros();

            // #region Click Events;
            // Add a click event handler for each button.
            $('#btnListCharNames').click(btnListCharNames);
            $('#selectName').change(selectNameChanged);
            $('#btnSlugline').click(btnSlugline);
            $('#btnAction').click(btnAction);
            $('#btnName').click(btnName);
            $('#btnDirection').click(btnDirection);
            $('#btnDialog').click(btnDialog);
            $('#btnCutTo').click(btnCutTo);
            $('#btnDissolveTo').click(btnDissolveTo);
            $('#btn2ndSlug').click(btn2ndSlug);
            $('#btnNotes').click(btnNotes);
            $('#btnParaphrase').click(btnParaphrase);
            $('#btnScene').click(btnScene);
            $('#btnNoteToDo').click(btnNoteToDo);

            $('#btnUpToTop').click(btnUpToTop_click);
            $('#btnWrite').click(btnWrite_click);
            $('#btnAnalyze').click(btnAnalyze_click);

            $('#btnFlow').click(btnFlow_click);
            $('#btnDialogReport').click(btnDialogReport_click);
            $('#btnGroupings').click(btnGroupings_click);
            $('#btnFlow').mouseover(btnFlow_mouseover);
            $('#btnDialogReport').mouseover(btnDialogReport_mouseover);
            $('#btnGroupings').mouseover(btnGroupings_mouseover);


            // #endregion
        });
    }

    function selectNameChanged() {
        return;
        //$('#listCharNames').show();
        //showNotification($('#selectName').val());
        getScenesWIthCharacter($('#selectName').val(), function (sceneList) {
            if (sceneList) {
                ($("#displayDiv").html(sceneList));
            }
        });
    }

    // #region Buttons

    function btnListCharNames() {
        listCharacterNames(function (nameList) {
            ($('#selectName').html(nameList));
        });

        ($('#selectName').show());
        //($('#btnListCharNames').hide());
    }

    function btnSlugline() {
        Word.run(function (context) {
            // Get the selection point and change the style to the current button's value
            context.document.getSelection().style = "Heading 1";
            showNotification("", "Set to 'Slugline'");
            return context.sync();
        })
            .catch(errorHandler);
    }

    function btnAction() {
        Word.run(function (context) {

            // Get the current paragraph, adjust the Font and Paragraph attributes, and sync it back.  
            // Send a notification ot the Notification area.

            var p = context.document.getSelection().paragraphs.getFirstOrNullObject();
            p.load();
            if (p == null || p == undefined || context.document.getSelection().paragraphs.length < 1) {
                showNotification("", "No paragraph selected");
                return context.sync();
            }
            context.sync().then(function () {
                p.load();
                p.font.set(
                    {
                        name: "Courier",
                        size: 11,
                        color: "#000000"
                    });

                p.set({
                    lineSpacing: 6,
                    leftIndent: 0.25,
                    rightIndent: -0.25,
                    spaceAfter: 8,
                    spaceBefore: 8
                })
                //context.document.getSelection().getRange().insertText("\r\n", "End");
            }).then(context.sync);

            showNotification("", "Set to 'Action'")
            //context.document.getSelection().originalRange.insertText("", "End");
            return context.sync();
        })
            .catch(errorHandler);
    }

    function btnName() {
        Word.run(function (context) {

            // Get the current paragraph, adjust the Font and Paragraph attributes, and sync it back.  
            // Send a notification ot the Notification area.

            var px = context.document.getSelection().paragraphs;
            context.load(px, "items");
            return context.sync().then(function () {
                var p = px.items[0];
                p.load("text, lineSpacing, leftIndent, spaceBefore, font/size, font/name, font/color");
                p.insertText(p.text.toUpperCase(), Word.InsertLocation.replace);
                p.font.set({
                    name: "Courier",
                    size: 11,
                    color: "#000000",
                })
                p.set({
                    lineSpacing: 12,
                    leftIndent: 180,
                    spaceBefore: 12,
                });
                context.sync();
                showNotification("", "Set to 'Character Name'");
            }).catch(errorHandler);
        });
    }

    function btnDirection() {
        Word.run(function (context) {
            // Queue a command to get the current selection and then
            // create a proxy range object with the results.
            var range = context.document.getSelection();

            // This variable will keep the search results for the longest word.
            var searchResults;

            // Queue a command to load the range selection result.
            context.load(range, 'text');

            // Synchronize the document state by executing the queued commands
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    // Get the longest word from the selection.
                    var words = range.text.split(/\s+/);
                    var longestWord = words.reduce(function (word1, word2) { return word1.length > word2.length ? word1 : word2; });

                    // Queue a search command.
                    searchResults = range.search(longestWord, { matchCase: true, matchWholeWord: true });

                    // Queue a commmand to load the font property of the results.
                    context.load(searchResults, 'font');
                })
                .then(context.sync)
                .then(function () {
                    // Queue a command to highlight the search results.
                    searchResults.items[0].font.highlightColor = '#FFFF00'; // Yellow
                    searchResults.items[0].font.bold = true;
                })
                .then(context.sync);
        })
            .catch(errorHandler);
    }

    function btnDialog() {
        Word.run(function (context) {
            // Queue a command to get the current selection and then
            // create a proxy range object with the results.
            var range = context.document.getSelection();

            // This variable will keep the search results for the longest word.
            var searchResults;

            // Queue a command to load the range selection result.
            context.load(range, 'text');

            // Synchronize the document state by executing the queued commands
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    // Get the longest word from the selection.
                    var words = range.text.split(/\s+/);
                    var longestWord = words.reduce(function (word1, word2) { return word1.length > word2.length ? word1 : word2; });

                    // Queue a search command.
                    searchResults = range.search(longestWord, { matchCase: true, matchWholeWord: true });

                    // Queue a commmand to load the font property of the results.
                    context.load(searchResults, 'font');
                })
                .then(context.sync)
                .then(function () {
                    // Queue a command to highlight the search results.
                    searchResults.items[0].font.highlightColor = '#FFFF00'; // Yellow
                    searchResults.items[0].font.bold = true;
                })
                .then(context.sync);
        })
            .catch(errorHandler);
    }

    function btnCutTo() {
        Word.run(function (context) {
            // Queue a command to get the current selection and then
            // create a proxy range object with the results.
            var range = context.document.getSelection();

            // This variable will keep the search results for the longest word.
            var searchResults;

            // Queue a command to load the range selection result.
            context.load(range, 'text');

            // Synchronize the document state by executing the queued commands
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    // Get the longest word from the selection.
                    var words = range.text.split(/\s+/);
                    var longestWord = words.reduce(function (word1, word2) { return word1.length > word2.length ? word1 : word2; });

                    // Queue a search command.
                    searchResults = range.search(longestWord, { matchCase: true, matchWholeWord: true });

                    // Queue a commmand to load the font property of the results.
                    context.load(searchResults, 'font');
                })
                .then(context.sync)
                .then(function () {
                    // Queue a command to highlight the search results.
                    searchResults.items[0].font.highlightColor = '#FFFF00'; // Yellow
                    searchResults.items[0].font.bold = true;
                })
                .then(context.sync);
        })
            .catch(errorHandler);
    }

    function btnDissolveTo() {
        Word.run(function (context) {
            // Queue a command to get the current selection and then
            // create a proxy range object with the results.
            var range = context.document.getSelection();

            // This variable will keep the search results for the longest word.
            var searchResults;

            // Queue a command to load the range selection result.
            context.load(range, 'text');

            // Synchronize the document state by executing the queued commands
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    // Get the longest word from the selection.
                    var words = range.text.split(/\s+/);
                    var longestWord = words.reduce(function (word1, word2) { return word1.length > word2.length ? word1 : word2; });

                    // Queue a search command.
                    searchResults = range.search(longestWord, { matchCase: true, matchWholeWord: true });

                    // Queue a commmand to load the font property of the results.
                    context.load(searchResults, 'font');
                })
                .then(context.sync)
                .then(function () {
                    // Queue a command to highlight the search results.
                    searchResults.items[0].font.highlightColor = '#FFFF00'; // Yellow
                    searchResults.items[0].font.bold = true;
                })
                .then(context.sync);
        })
            .catch(errorHandler);
    }

    function btn2ndSlug() {
        Word.run(function (context) {
            // Queue a command to get the current selection and then
            // create a proxy range object with the results.
            var range = context.document.getSelection();

            // This variable will keep the search results for the longest word.
            var searchResults;

            // Queue a command to load the range selection result.
            context.load(range, 'text');

            // Synchronize the document state by executing the queued commands
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    // Get the longest word from the selection.
                    var words = range.text.split(/\s+/);
                    var longestWord = words.reduce(function (word1, word2) { return word1.length > word2.length ? word1 : word2; });

                    // Queue a search command.
                    searchResults = range.search(longestWord, { matchCase: true, matchWholeWord: true });

                    // Queue a commmand to load the font property of the results.
                    context.load(searchResults, 'font');
                })
                .then(context.sync)
                .then(function () {
                    // Queue a command to highlight the search results.
                    searchResults.items[0].font.highlightColor = '#FFFF00'; // Yellow
                    searchResults.items[0].font.bold = true;
                })
                .then(context.sync);
        })
            .catch(errorHandler);
    }

    function btnNotes() {
        Word.run(function (context) {
            // Queue a command to get the current selection and then
            // create a proxy range object with the results.
            var range = context.document.getSelection();

            // This variable will keep the search results for the longest word.
            var searchResults;

            // Queue a command to load the range selection result.
            context.load(range, 'text');

            // Synchronize the document state by executing the queued commands
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    // Get the longest word from the selection.
                    var words = range.text.split(/\s+/);
                    var longestWord = words.reduce(function (word1, word2) { return word1.length > word2.length ? word1 : word2; });

                    // Queue a search command.
                    searchResults = range.search(longestWord, { matchCase: true, matchWholeWord: true });

                    // Queue a commmand to load the font property of the results.
                    context.load(searchResults, 'font');
                })
                .then(context.sync)
                .then(function () {
                    // Queue a command to highlight the search results.
                    searchResults.items[0].font.highlightColor = '#FFFF00'; // Yellow
                    searchResults.items[0].font.bold = true;
                })
                .then(context.sync);
        })
            .catch(errorHandler);
    }

    function btnParaphrase() {
        Word.run(function (context) {
            // Queue a command to get the current selection and then
            // create a proxy range object with the results.
            var range = context.document.getSelection();

            // This variable will keep the search results for the longest word.
            var searchResults;

            // Queue a command to load the range selection result.
            context.load(range, 'text');

            // Synchronize the document state by executing the queued commands
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    // Get the longest word from the selection.
                    var words = range.text.split(/\s+/);
                    var longestWord = words.reduce(function (word1, word2) { return word1.length > word2.length ? word1 : word2; });

                    // Queue a search command.
                    searchResults = range.search(longestWord, { matchCase: true, matchWholeWord: true });

                    // Queue a commmand to load the font property of the results.
                    context.load(searchResults, 'font');
                })
                .then(context.sync)
                .then(function () {
                    // Queue a command to highlight the search results.
                    searchResults.items[0].font.highlightColor = '#FFFF00'; // Yellow
                    searchResults.items[0].font.bold = true;
                })
                .then(context.sync);
        })
            .catch(errorHandler);
    }

    function btnScene() {
        Word.run(function (context) {
            // Queue a command to get the current selection and then
            // create a proxy range object with the results.
            var range = context.document.getSelection();

            // This variable will keep the search results for the longest word.
            var searchResults;

            // Queue a command to load the range selection result.
            context.load(range, 'text');

            // Synchronize the document state by executing the queued commands
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    // Get the longest word from the selection.
                    var words = range.text.split(/\s+/);
                    var longestWord = words.reduce(function (word1, word2) { return word1.length > word2.length ? word1 : word2; });

                    // Queue a search command.
                    searchResults = range.search(longestWord, { matchCase: true, matchWholeWord: true });

                    // Queue a commmand to load the font property of the results.
                    context.load(searchResults, 'font');
                })
                .then(context.sync)
                .then(function () {
                    // Queue a command to highlight the search results.
                    searchResults.items[0].font.highlightColor = '#FFFF00'; // Yellow
                    searchResults.items[0].font.bold = true;
                })
                .then(context.sync);
        })
            .catch(errorHandler);
    }

    function btnNoteToDo() {
        ($('#Analyze').hide());
        ($('#Write').hide());
        ($('#topTabs').show());
    }

    // #endregion

    // #region Tabs

    function btnUpToTop_click() {
        ($("#divHeaderMessage").hide());
       ($("#displayDiv").html(""));
        ($('#Analyze').hide());
        ($('#Write').hide());
        ($('#Tokyo').hide());
        ($('#topTabs').show());
    }

    function btnWrite_click() {
        ($("#displayDiv").html(""));
        ($('#Analyze').hide());
        ($('#Write').show());
    }

    function btnAnalyze_click() {
       ($("#divHeaderMessage").html("Analysis Tools"));
        ($("#divHeaderMessage").show());
       ($('#topTabs').hide());
        ($('#Write').hide());
        ($('#Analyze').show());
        ($('#Tokyo').show());
        $('#selectName').focus();
        listCharacterNames(function (nameList) {
            ($('#selectName').html(nameList));
        });
    }

    function btnDialogReport_click() {
        ($("#divUserMessage").html("All of character(s) speeches grouped together"));
        ($("#displayDiv").html(""));
        //showNotification("");
        getCharacterDialog($('#selectName').val(), function (sceneList) {
            if (sceneList) {
                ($("#displayDiv").html(sceneList));
            }
        });
    }

    function btnDialogReport_mouseover() {
        //showNotification("View story structure as a whole");
        ($("#divUserMessage").html("All of character(s) speeches grouped together"));
   }

    function btnGroupings_mouseover() {
        ($("#divUserMessage").html("Groupings of selected characters throughout the story"));
    }

    function btnGroupings_click() {
       ($("#divUserMessage").html("Groupings of selected characters throughout the story"));
       ($("#displayDiv").html(""));
        //showNotification("");

        getCharactersInScenes($('#selectName').val(), function (sceneList) {
            var output;

            for (let i = 0; i < sceneList.length; i++) {
                if (Array.isArray(sceneList[i])) {
                    sceneList[i] = sceneList[i].join(", ");

                }
                if (Array.isArray(sceneList[i + 1])) {
                    sceneList[i + 1] = sceneList[i + 1].join(", ");
                }
                output +=  sceneList[i + i] + "<br />" + sceneList[i];
                i++
            }
            if (sceneList) {
                ($("#displayDiv").html(output));
            }
        })
    }

    function btnFlow_mouseover() {
        //showNotification("Character(s) in scenes as they flow through the story");
        //ms - font - s ms - fontColor - white
        
            ($("#divUserMessage").html("Character(s) in scenes as they flow through the story"));

    }

    function btnFlow_click() {
        ($("#divUserMessage").html("Character(s) in scenes as they flow through the story"));
       ($("#displayDiv").html(""));
        //showNotification("");
        getScenesByCharacter($('#selectName').val(), function (sceneList) {
            if (sceneList) {
                ($("#displayDiv").html(sceneList));
            }
        });
    }

    // #endregion

    // #region Helpers

    function setTemplate() {
        Word.run(function (context) {
            var a = context.application.createDocument(this, "./MovieTemplate.txt", DocumentType.Base64);
            //a.load();
            //a.open();
            context.sync();

            // showNotification("", Word.DocumentProperties["template"]);
            var b = Word.DocumentProperties(a);

            //context.document.load(a);
            //const newdoc1 = context.application.createDocument().open('./MovieTemplate.txt');
            //var itworked = context.application.createDocument('./MovieTemplate.txt').load().open();
            //context.document.load(newdoc);
            //context.sync();
            //newdoc.open();
            //const tmpl = context.document.properties.template;

            //Word.http.get('./MovieTemplate.txt').subscribe(response => {
            //    Word.run(async context => {
            //        const myNewDoc = context.application.createDocument(response);
            //        context.load(myNewDoc);
            //        await context.sync();
            //        myNewDoc.open();
            //        await context.sync();
            //    });
            return context.sync();
        })
            .catch(errorHandler);
    };

    function applyStyle(para, stylename) {
        Word.run(function (context) {

            para.style = stylename;
            return context.sync();
        })
            .catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
    }

    function loadSampleData() {
        // Run a batch operation against the Word object model.
        Word.run(function (context) {
            // Create a proxy object for the document body.
            var body = context.document.body;

            // Queue a commmand to clear the contents of the body.
            body.clear();
            // Queue a command to insert text into the end of the Word document body.
            body.insertText(
                "This is a sample text inserted in the document",
                Word.InsertLocation.end);

            // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
            return context.sync();
        })
            .catch(errorHandler);
    }

    function displaySelectedText() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error:', result.error.message);
                }
            });
    }

    //$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
    function errorHandler(error) {
        // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
        showNotification("Error:", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }

    function arrayContainsArray(superset, subset) {
        if (0 === subset.length) {
            return false;
        }
        return subset.every(function (value) {
            return (superset.indexOf(value) >= 0);
        });
    }


    // #endregion

    function listCharacterNames(callback) {
        Word.run(function (context) {
            var out = "";
            var charNameList;
            var paragraph;
            var paras = context.document.body.paragraphs;
            context.load(paras, 'text, style');
            return context.sync()
                .then(function () {
                    for (let i = 0; i < paras.items.length; i++) {
                        paragraph = paras.items[i];
                        if (paragraph.style === "sCharacter Name" && paragraph.text.length > 0)
                            charNameList += "," + paras.items[i].text.toUpperCase();
                    }
                    context.sync()
                        .then(function () {
                            out = sortByFrequency(charNameList.split(",").filter(Boolean));
                            out.filter(name => name != 'undefined' && name != "");
                            for (var k = 0; k < out.length; k++) {
                                out[k] = "<option>" + out[k] + "</option>";
                            }
                            //delete out[0];
                            //out.splice(0, 0, "<option><option>");
                            callback(out);
                        });
                })
                .catch(function (error) {
                    showNotification('Error: ' + JSON.stringify(error));
                    if (error instanceof OfficeExtension.Error) {
                        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                    }
                });
        });
    }

    function getScenesByCharacter(namesToFind, callback) {
        Word.run(function (context) {
            //var charSummaryMap = ['<image src="~../../images/paragraph.jpg"> ', ''];
            var paragraph;
            var summ;
            var charsFoundInScene = [];
            var paras = context.document.body.paragraphs;
            context.load(paras, 'text, style');
            return context.sync()
                .then(function () {
                    var charSummaryMap = [];
                    for (var i = 0; i < paras.items.length; i++) {
                        paragraph = paras.items[i];
                        if (paragraph.style === "Heading 1,Act Break")
                            charSummaryMap.push("<b>" + paragraph.text + "</b><br />" + "<hr />");
                        if (paragraph.style === "Heading 2,Summary") {
                            summ = paragraph.text + "<br />";
                            let j = ++i;
                            paragraph = paras.items[j];
                            while (j < paras.items.length && paragraph.style != "Heading 2,Summary") {
                                paragraph = paras.items[j];
                                if (paragraph.style === "sCharacter Name"
                                    && namesToFind.includes(paragraph.text.toUpperCase())) {
                                    charsFoundInScene.push(paragraph.text.toUpperCase());
                                }
                                j++;
                            }
                            if (arrayContainsArray(namesToFind, charsFoundInScene) && !charSummaryMap.includes(summ)) {
                                charSummaryMap.push(summ);
                                charsFoundInScene = [];
                            }
                        }
                    } // end for
                    callback(charSummaryMap.join("<br>"));
                    context.sync();
                })
        })
            .catch(function (error) {
                showNotification('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            })
    }

    function getCharactersInScenes(namesToFind, callback) {
        Word.run(function (context) {
            // Show all the characters grouped together, for every scene
            var paragraph;
            var summ;
            var paras = context.document.body.paragraphs;
            context.load(paras, 'text, style');
            return context.sync()
                .then(function () {
                    var summ;
                    var charSummaryMap = [];
                    var charsFoundInScene = [];
                    for (var i = 0; i < paras.items.length; i++) {
                        paragraph = paras.items[i];
                        if (paragraph.style === "Heading 1,Act Break" && paragraph.text != undefined)
                            charSummaryMap.push("<b>" + paragraph.text + "</b><br />" + "<hr />");
                        if (paragraph.style === "Heading 2,Summary"
                            && paragraph.text != undefined
                            && !charSummaryMap.includes(paragraph.text)) {
                            summ = paragraph.text + "<br />";
                            i++;
                            paragraph = paras.items[i];
                            let j = i;
                            while (j < paras.items.length && paragraph.style != "Heading 2,Summary") {
                                if (paragraph.style === "sCharacter Name"
                                    && paragraph.text != undefined
                                    && !charsFoundInScene.includes(paragraph.text.toUpperCase())) {
                                    charsFoundInScene.push(paragraph.text.toUpperCase());
                                }
                                j++;
                                paragraph = paras.items[j];
                            } // end while
                            charSummaryMap.push(charsFoundInScene ? charsFoundInScene : "");
                            charSummaryMap.push(summ);

                            charsFoundInScene = [];
                            i++;
                            paragraph = paras.items[i];
                        }
                    } // end for

                    callback(charSummaryMap);
                    context.sync();
                })
        })
            .catch(function (error) {
                showNotification('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            })
    }

    function getCharacterDialog(nameToMap, callback) {
        Word.run(function (context) {
            var charSummaryMap = ['', ''];
            var paragraph, summ, isInScene = false;
            var paras = context.document.body.paragraphs;
            context.load(paras, 'text, style');
            return context.sync()
                .then(function () {
                    for (var i = 0; i < paras.items.length; i++) {
                        paragraph = paras.items[i];
                        //grab the Act, put it in the output
                        if (paragraph.style === "Heading 1,Act Break")
                            charSummaryMap.push("<b>" + paragraph.text + "</b>");

                        if (paragraph.style === "Heading 2,Summary") {
                            summ = paragraph.text;
                            i++
                            paragraph = paras.items[i];
                            //grabbed the scene summary, now check if the character is in that scene
                            //if nameToMap is found in the scene, store the scene summary in output array
                            while (i < paras.items.length && paragraph.style != "Heading 2,Summary") {
                                paragraph = paras.items[i];
                                if (paragraph.style === "Heading 1,Act Break") {
                                    charSummaryMap.push("<b>" + paragraph.text + "</b>");
                                }
                                if (paragraph.style === "sCharacter Name" && !isInScene) {
                                    if (paragraph.text.trim().toUpperCase() === nameToMap.trim()) {
                                        if (summ.trim().length > 0) {
                                            charSummaryMap.push('', summ);
                                            isInScene = true;
                                        }
                                    }
                                }
                                i++;
                            }
                            isInScene = false;
                        }
                    }
                    callback(charSummaryMap.join("<br>"));
                    charSummaryMap = ['', ''];
                    context.sync();
                })
                .catch(function (error) {
                    showNotification('Error: ' + JSON.stringify(error));
                    if (error instanceof OfficeExtension.Error) {
                        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                    }
                });
        })
    }

    function sortByFrequency(arr) {
        let counter = arr.reduce(
            (counter, key) => {
                counter[key] = 1 + counter[key] || 1;
                return counter
            }, {});
        //console.log(counter);
        // {"apples": 1, "oranges": 4, "bananas": 2}

        // sort counter by values (compare position 1 entries)
        // the result is an array
        let sorted_counter = Object.entries(counter).sort((a, b) => b[1] - a[1]);
        //showNotification(sorted_counter);
        // [["oranges", 4], ["bananas", 2], ["apples", 1]]

        // show only keys of the sorted array
        return (sorted_counter.map(x => x[0]));
    }

})();