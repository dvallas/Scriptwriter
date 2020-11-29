function highlightDiv(e) {
    //$( "#test" ).html("background-color:grey")
    $(document).on('mouseover', 'div', function (e) {
        $("#test").html((e.target).getAttribute('id'));
    });
}


function printMousePos(e) {

    cursorX = e.at;
    cursorY = e.pageY;
    $("#" + (e.target).getAttribute('id')).html("")
    $("#" + (e.target).getAttribute('id')).html("x".repeat(cursorY / 8))
    $("#test").html("pageX: " + cursorX + ",pageY: " + cursorY);
}