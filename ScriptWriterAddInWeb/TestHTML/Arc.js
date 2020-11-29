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

    let midpoint = 400;

    let out = (cursorY >= midpoint ? "Above" : "Below");

    isMid(out);
}

function printLine(e) {
    printAbsoluteMousePos(e, function (isMid) {
        //var elem = $('#' + (e.target).getAttribute('id'));
        let mid = isMid;

        if (isMid === "Above") {
            //let y = e.target;
            createBar(e, "up");
        } else {
            createBar(e, "down")
                //createBar(e, "up")
        }
    });
    $("#test").html("pageX: " + cursorX + ", pageY: " + cursorY + "  max chars: " + $("#matrixRight").innerHeight());
}


function printMousePos(e) {
    cursorX = e.pageX;
    cursorY = e.pageY;
    var elem = $('#' + (e.target).getAttribute('id'));


    let char = "x";
    elem.html(char.repeat(cursorX));
    //.html("pageX: " + cursorX + ",pageY: " + cursorY);
}


function createBar(element, direction) {
    // var j = "#c" + element.target.id;
    //var b=$j[0];
    var f = element.target.id.substring(0, 1) === "c" ? document.getElementById(element.target.id) : document.getElementById("c" + element.target.id)
    console.log("CursorX " + cursorX + "CursorY " + cursorY + " ID " + element.target.id);

    //var b = $(j).get(0);
    var ctx = f.getContext("2d");
    ctx.beginPath();
    if (direction === "down") {
        //x, y, width, height

        console.log("X=" + '0' + " Y=" + cursorY + " Width=" + '20' +
            " Height=" + (400 - cursorY));

        ctx.rect(0, cursorY, 20, document.getElementById(element.target.id).height - cursorY);
    } else {
        //start, top, end, height
        ctx.rect(0, 0, 20, cursorY);
        //ctx.rect(cursorX, 0, 5, cursorY);
    }
    //ctx.rect(0, 0, 5, cursorY);
    ctx.stroke();


    ctx.fill();
}


function printMouseTrail(e) {
    cursorX = e.pageX;
    cursorY = e.pageY;

    var id = "('#" + (e.target) + "')";
    for (i = 200; i > 20; i--) {
        setCursorPosition(i, id);
    }
}

//SET CURSOR POSITION
function setCursorPosition(index, element) {
    this.each(function (index, elem) {
        if (elem.setSelectionRange) {
            elem.setSelectionRange(pos, pos);
        } else if (elem.createTextRange) {
            var range = elem.createTextRange();
            range.collapse(true);
            range.moveEnd('character', pos);
            range.moveStart('character', pos);
            range.select();
        }
    });
    return this;
}
