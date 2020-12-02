﻿function highlightDiv(e) {
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

        let mid = isMid;

        if (isMid === "Above") {
            createBar(e, "up");
        } else {
            createBar(e, "down")
        }

    });
    /*    $("#test").html("pageX: " + cursorX + ", pageY: " + cursorY + "  max chars: " + $("#matrixRight").innerHeight());*/
}

function createBar(element, direction) {

    var f = element.target.id.substring(0, 1) === "c" ? document.getElementById(element.target.id) : document.getElementById("c" + element.target.id)

    var ctx;
    try {
        ctx = f.getContext("2d");
    } catch {
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





/*
/******************************************************************************************
 *
 *  EASY CURVE FUNCITON
 *
 *******************************************************************************************/

function drawCurve() {
    return false;
    var points = [
                [10, 10],
                [40, 30],
                [100, 10],
                [200, 100],
                [200, 50],
                [250, 120]
            ];
    let test = points[0][0];
    var ctx = document.getElementById("cnv").getContext("2d");
    ctx.moveTo(points[0][0], points[0][1]);

    ctx.beginPath();
    for (i = 1; i < points.length - 2; i++) {
        var xc = (points[i][0] + points[i + 1][0]) / 2;
        var yc = (points[i][1] + points[i + 1][1]) / 2;
        ctx.quadraticCurveTo(points[i][0], points[i][1], xc, yc);
    }
    // curve through the last two points
    ctx.quadraticCurveTo(points[i][0], points[i][1], points[i + 1][0], points[i + 1][1]);
    ctx.stroke();
    ctx.fill();

}

/*

*****************************************************************************************
 *
 *  COMPLEX CURVE FUNCITON
 *
 *******************************************************************************************

function runCurve() {
    var ctx = document.getElementById("cnv").getContext("2d");
    var myPoints = [10, 10, 40, 30, 100, 10, 200, 100, 200, 50, 250, 120]; //minimum two points
    var tension = 1;

    if (CanvasRenderingContext2D != 'undefined') {
        CanvasRenderingContext2D.prototype.drawCurve =
            function (pts, tension, isClosed, numOfSegments, showPoints) {
                drawCurve(this, pts, tension, isClosed, numOfSegments, showPoints)
            }
    }

    drawCurve(ctx, myPoints); //default tension=0.5
    drawCurve(ctx, myPoints, tension);
}

function drawCurve(ctx, ptsa, tension, isClosed, numOfSegments, showPoints) {

    ctx.beginPath();

    drawLines(ctx, getCurvePoints(ptsa, tension, isClosed, numOfSegments));

    if (showPoints) {
        ctx.beginPath();
        for (var i = 0; i < ptsa.length - 1; i += 2)
            ctx.rect(ptsa[i] - 2, ptsa[i + 1] - 2, 4, 4);
    }

    ctx.stroke();
}

function drawLines(ctx, pts) {
    ctx.moveTo(pts[0], pts[1]);
    for (i = 2; i < pts.length - 1; i += 2) ctx.lineTo(pts[i], pts[i + 1]);
}

function getCurvePoints(pts, tension, isClosed, numOfSegments) {

    // use input value if provided, or use a default value	 
    tension = (typeof tension != 'undefined') ? tension : 0.5;
    isClosed = isClosed ? isClosed : false;
    numOfSegments = numOfSegments ? numOfSegments : 16;

    var _pts = [],
        res = [], // clone array
        x, y, // our x,y coords
        t1x, t2x, t1y, t2y, // tension vectors
        c1, c2, c3, c4, // cardinal points
        st, t, i; // steps based on num. of segments

    // clone array so we don't change the original
    //
    _pts = pts; //.slice(0);

    // The algorithm require a previous and next point to the actual point array.
    // Check if we will draw closed or open curve.
    // If closed, copy end points to beginning and first points to end
    // If open, duplicate first points to befinning, end points to end
    if (isClosed) {
        _pts.unshift(pts[pts.length - 1]);
        _pts.unshift(pts[pts.length - 2]);
        _pts.unshift(pts[pts.length - 1]);
        _pts.unshift(pts[pts.length - 2]);
        _pts.push(pts[0]);
        _pts.push(pts[1]);
    } else {
        _pts.unshift(pts[1]); //copy 1. point and insert at beginning
        _pts.unshift(pts[0]);
        _pts.push(pts[pts.length - 2]); //copy last point and append
        _pts.push(pts[pts.length - 1]);
    }

    // ok, lets start..

    // 1. loop goes through point array
    // 2. loop goes through each segment between the 2 pts + 1e point before and after
    for (i = 2; i < (_pts.length - 4); i += 2) {
        for (t = 0; t <= numOfSegments; t++) {

            // calc tension vectors
            t1x = (_pts[i + 2] - _pts[i - 2]) * tension;
            t2x = (_pts[i + 4] - _pts[i]) * tension;

            t1y = (_pts[i + 3] - _pts[i - 1]) * tension;
            t2y = (_pts[i + 5] - _pts[i + 1]) * tension;

            // calc step
            st = t / numOfSegments;

            // calc cardinals
            c1 = 2 * Math.pow(st, 3) - 3 * Math.pow(st, 2) + 1;
            c2 = -(2 * Math.pow(st, 3)) + 3 * Math.pow(st, 2);
            c3 = Math.pow(st, 3) - 2 * Math.pow(st, 2) + st;
            c4 = Math.pow(st, 3) - Math.pow(st, 2);

            // calc x and y cords with common control vectors
            x = c1 * _pts[i] + c2 * _pts[i + 2] + c3 * t1x + c4 * t2x;
            y = c1 * _pts[i + 1] + c2 * _pts[i + 3] + c3 * t1y + c4 * t2y;

            //store points in array
            res.push(x);
            res.push(y);

        }
    }
    return res;
}
*/
