
    // Get the document mode and the URL of the active project.
    function showDocumentProperties() {
        var output = String.format(
            'The document mode is {0}.<br/>The URL of the active project is {1}.',
            Office.context.document.mode,
            Office.context.document.url);
        $('#divTopMessage').html(output);
    }

    function highlightDiv(e) {
        //$( "#test" ).html("background-color:grey")
        $(document).on('mouseover', 'div', function (e) {
            $("#test").html((e.target).getAttribute('id'));

        });
    }

    function printAbsoluteMousePos(e, isMid) {
        cursorX = e.pageX;
        cursorY = e.pageY;
        $("#test").html("");
        var elem = $('#' + (e.target).getAttribute('id'));


        let midpoint = $(window).height() / 2;

        let out = (cursorY <= midpoint ? "Above" : "Below");
        isMid(out);
    }

    function printLine(e) {
        printAbsoluteMousePos(e, function (isMid) {

            if (isMid === "Above") {
                createBar(e, "up");
            } else {
                createBar(e, "down")
            }

        });
        /*    $("#test").html("pageX: " + cursorX + ", pageY: " + cursorY + "  max chars: " + $("#matrixRight").innerHeight());*/
    }

    //var x = document.createElement("CANVAS");
    //document.getElementById("container").appendChild(x);

    function createBar(element, direction) {

        var f = element.target.id.substring(0, 1) === "c" ? document.getElementById(element.target.id) : document.getElementById("c" + element.target.id);

        var ctx;
        try {
            ctx = f.getContext("2d");
        } catch(error) {
            console.log("failed to get context for " + element.target.id);
            return false;
        }
        //var mid = $(window).height() / 2;


        let factor = 1 / $(window).height();
        let mid = 50;
        let cY = Math.ceil(cursorY * factor * 100);
        let cX = Math.ceil(cursorX * factor * 100);
        console.log("mid: " + mid + " canvas height: " + f.height + " canvas width: " + f.width);
        console.log("Factor: " + factor + " Mid:" + mid + " Y:" + cY + " X:" + cX);

        ctx.beginPath();
        console.log("Direction: " + direction + " CursorX " + cursorX + " CursorY " + cursorY + " ID " + element.target.id);
        if (direction === "down") {
            //x, y, width, height
            ctx.rect(0, mid, 60, cY - mid);
        } else {
            // up
            //x, y, width, height
            ctx.rect(0, cY, 60, mid - cY);
        }

        ctx.stroke();
        ctx.fillStyle = "red";
        ctx.fill();

        /*    var ctx = document.getElementById("c4").getContext("2d");
            ctx.fillStyle = "red";
            ctx.fill();*/
    }
