"use strict";

(function () {
    var messageBanner;
    var whichReport;
    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();

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
            // Add event handlers for each button.
            $('#btnListCharNames').click(btnListCharNames);

            //$('#selectName').change(selectNameChanged);
            $('#btnRunReport').click(selectNameChanged);

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
            $('#dropDownAnalyze').click(dropDownAnalyze_click);

            $('#btnDialogReport').mouseover(btnDialogReport_mouseover);
            $('#btnDialogReport').click(btnDialogReport_click);
            $('#btnGroupings').click(btnGroupings_click);
            $('#btnGroupings').mouseover(btnGroupings_mouseover);
            $('#btnFlow').mouseover(btnFlow_mouseover);
            $('#btnFlow').click(btnFlow_click);

            $('#btnHome').click(btnHome_click);
            $('#dropDownAnalyze').mouseover(dropDownAnalyze_mouseover);
            $('#dropDownAnalyze').mouseout(dropDownAnalyze_mouseout);

            //$('#hamburger').mouseover(MenuActiveToggle("hamburger"));
            //$('#hamburger').click(hamburger_click);

            ($('#TopNav').show());
            //($('#divTopMessage').text("this is display text"));

            // #endregion
        });
    }

    function selectNameChanged() {
        ($("#Tokyo").hide());
        ($("#divTopMessage").html(""));
        ($("#divUserMessage").html(""));
        ($('#Write').hide());
        ($("#displayDiv").html(""));


        if (whichReport && whichReport === "flow") {
            ($("#divTopMessage").html("Scene appearances of " + $('#selectName').val().join(" + ")));
            ($("#divUserMessage").html("Character(s) in scenes as they flow through the story"));
           getSceneFlowByCharacter($('#selectName').val(), function (sceneList) {
                if (sceneList) {
                    ($("#displayDiv").html(sceneList));
                }
            });
        } else if (whichReport === "groupings") {
            ($("#divTopMessage").html("Scene Groupings for " + $('#selectName').val().join(" + ")));
            ($("#divUserMessage").html("Groupings of selected characters throughout the story"));
            getCharacterGroupingsInScenes($('#selectName').val(), function (sceneList) {
                let output = [];
                let names = "";
                let summary = "";
                for (let i = 0; i < sceneList.length; i++) {
                    summary = sceneList[0, i];
                    names = sceneList[1, i];
                    if (!summary === undefined && !summary === names) {
                        output.push("<span>" + summary + "<br />" + names + "</span><br />");
                    } else {
                        output.push("<span>" + names + "</span ><br />");
                    }
                }
                if (sceneList) {
                    ($("#displayDiv").html(output));
                }
            });
        } else if (whichReport === "dialog") {
            ($("#divTopMessage").html("All Speeches From " + $('#selectName').val().join(" + ")));
            ($("#divUserMessage").html("All of character(s) speeches grouped together"));
            getCharacterDialog($('#selectName').val(), function (dialogList) {
                if (dialogList) {
                    ($("#displayDiv").html(dialogList));
                }
            });
        }
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
        ($('#TopNav').show());
    }

    function btnWrite_click() {
        ($("#divTopMessage").html("Formatting"));
        ($("#divUserMessage").html("Manually assign formatting to paragraphs"));
        MenuActiveToggle("btnWrite");
        ($("#displayDiv").html(""));
        ($('#Analyze').hide());
        ($('#Write').show());
    }

    function dropDownAnalyze_mouseover() {
        //($("#displayDiv").html(""));
        //($('#Analyze').hide());
        ($('#analyzeMenu').show());
    }

    function dropDownAnalyze_mouseout() {
        //($("#displayDiv").html(""));
        //($('#Analyze').hide());
        ($('#analyzeMenu').hide());
    }

    function dropDownAnalyze_click() {
        //($("#divHeaderMessage").html("Analysis Tools"));
        //($("#divHeaderMessage").show());
        //($('#TopNav').hide());
        ($('#Write').hide());
    }

    function btnDialogReport_click() {
        whichReport = "dialog";
        listCharacterNames(function (nameList) {
            ($('#selectName').html(nameList));
        });
        ($('#Tokyo').show());
        $('#selectName').focus();
    }

    function btnDialogReport_mouseover() {
        //showNotification("View story structure as a whole");
        ($("#divUserMessage").html("All of character(s) speeches grouped together"));
    }

    function btnGroupings_mouseover() {
        //($("#divUserMessage").html("Groupings of selected characters throughout the story"));
    }

    function btnGroupings_click() {
        whichReport = "groupings";
        listCharacterNames(function (nameList) {
            ($('#selectName').html(nameList));
        });
        ($('#Tokyo').show());
        $('#selectName').focus();


    }

    function btnFlow_mouseover() {
        //showNotification("Character(s) in scenes as they flow through the story");
        //ms - font - s ms - fontColor - white
        //listCharacterNames(function (nameList) {
        //    ($('#selectName').html(nameList));
        //});

        //($("#Tokyo").show());
        //($('#selectName').show());


    }

    function btnFlow_click() {
        whichReport = "flow";
        listCharacterNames(function (nameList) {
            ($('#selectName').html(nameList));
        });
        ($('#Tokyo').show());
        $('#selectName').focus();
    }

    function btnHome_click() {
        MenuActiveToggle("btnHome");
        ($("#divUserMessage").text("Written By"));
        ($("#divTopMessage").text(""));
        ($("#displayDiv").hide());
        ($('#Write').hide());
        ($('#Tokyo').hide());

        //($("#divTopMessage").text("hello"));

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
        $("#notification-body").text(content);
        $("#notification-header").text(header);
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

    function MenuActiveToggle(element) {
        var x = document.getElementById(element);
        if (x.style.class === "") {
            x.style.class = "Active";
        } else {
            x.style.class = "";
        }
    }

    /* Toggle between adding and removing the "responsive" class to topnav when the user clicks on the icon */
    function hamburger_click() {
        var x = document.getElementById("hamburger");
        if (x.className === "topnav") {
            x.className += " responsive";
        } else {
            x.className = "topnav";
        }
    }


    // #endregion

    // #region Logic

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

    function getSceneFlowByCharacter(namesToFind, callback) {
        Word.run(function (context) {
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

    function getCharacterGroupingsInScenes(namesToFind, callback) {
        Word.run(function (context) {
            // Show all the characters grouped together, for every scene
            var paragraph;
            var paras = context.document.body.paragraphs;
            context.load(paras, 'text, style');
            return context.sync()
                .then(function () {
                    var summ;
                    var charSummaryMap = []
                    var charsFoundInScene = [];
                    let i = 0;
                    while (i < paras.items.length) {
                        paragraph = paras.items[i];
                        if (paragraph.style === "Heading 1,Act Break") {
                            charSummaryMap.push(["<b>" + paragraph.text + "</b><br /><hr />"]);
                        }
                        if (paragraph.style === "Heading 2,Summary") {
                            summ = paragraph.text + "<br />";
                            i++;
                            // get the characters in the scene
                            while (i < paras.items.length) {
                                //need to get all the names in each scene, then check if any of the names 
                                //is in the namesToFind list.  Either discard all, or add all
                                paragraph = paras.items[i];
                                //check the rest of the names in this scene
                                if (paragraph.style === "sCharacter Name" && !charsFoundInScene.includes(paragraph.text.toUpperCase())) {
                                    charsFoundInScene.push(paragraph.text.toUpperCase());
                                }
                                if (paragraph.style === "Heading 1,Act Break" && !charSummaryMap.includes(paragraph.text)) {
                                    charSummaryMap.push(["<b>" + paragraph.text + "</b><br /><hr />"]);
                                }
                                if (paragraph.style === "Heading 2,Summary") {
                                    break;
                                }
                                i++;
                                paragraph = paras.items[i];
                            }// end inner while

                        }// end if == Heading 2, Summary
                        //push the scene summary and list of names to the collector array if appropriate
                        if (charsFoundInScene.length > 0 && namesToFind.some(ai => charsFoundInScene.includes(ai))) {
                            // if (!charSummaryMap.includes(summ)) {
                            charSummaryMap.push([summ, charsFoundInScene]);
                            // }
                        }
                        //summ = "";
                        charsFoundInScene = [];
                        i++;
                    } //end outer while
                    callback(charSummaryMap);
                    context.sync();
                }) // end .then
                .catch(function (error) {
                    showNotification('Error: ' + JSON.stringify(error));
                    if (error instanceof OfficeExtension.Error) {
                        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                    }
                }) // end catch
        }) // end Word.run
    }// end function

    function getCharacterDialog(nameToMap, callback) {
        Word.run(function (context) {
            var charDialogList = [];
            var paragraph, charName;
            var paras = context.document.body.paragraphs;
            context.load(paras, 'text, style, font');
            return context.sync()
                .then(function () {
                    for (var i = 0; i < paras.items.length; i++) {
                        paragraph = paras.items[i];
                        //grab the Act, put it in the output
                        if (paragraph.style === "Heading 1,Act Break") {
                            charDialogList.push("<br /><b>" + paragraph.text + "</b><br />");
                        }
                        // grab selected characters' dialog per scene (demarcated by sSlugline)
                        if (paragraph.style === "sCharacter Name" && nameToMap.includes(paragraph.text.toUpperCase())) {
                            while (i < paras.items.length && paragraph.style != "sSlugline") {
                                if (paragraph.style === "Heading 1,Act Break") {
                                    charDialogList.push("<br /><b>" + paragraph.text + "</b><br />");
                                }
                                if (paragraph.style === "sCharacter Name" && nameToMap.includes(paragraph.text.toUpperCase())) {
                                    charName = paragraph.text;
                                    if (i < paras.items.length) {
                                        i++;
                                        paragraph = paras.items[i];
                                    }
                                    if (paragraph.style === "sDialog") {
                                        let f = paragraph.font;
                                        if (f.strikeThrough) {
                                            charDialogList.push(charName.toUpperCase() + "<br / ><strike>" + paragraph.text + "</strike><br />");
                                        }
                                        else {
                                            charDialogList.push(charName.toUpperCase() + "<br />" + paragraph.text + "<br /></br />");
                                        }
                                    }
                                }
                                if (i < paras.items.length) {
                                    i++;
                                    paragraph = paras.items[i];
                                }
                            }
                        }
                    }
                    callback(charDialogList);
                    charDialogList = [];
                    context.sync();
                })
                .catch(function (error) {
                    //showNotification('Error: ' + error.content.join(", "));
                    showNotification('Error: ' + JSON.stringify(error));
                    if (error instanceof OfficeExtension.Error) {
                        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                    }
                });
        })
    }

    // #endregion

})();