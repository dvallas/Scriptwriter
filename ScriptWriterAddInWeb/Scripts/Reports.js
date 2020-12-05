    function getSummaries(act, callback) {
        Word.run(function (context) {
            var paras = context.document.body.paragraphs;
            var paragraph;
            var charSummaries = [];
            context.load(paras, 'text, style');
            return context.sync()
                .then(function () {
                    for (var i = 0; i < paras.items.length; i++) {
                        paragraph = paras.items[i];
                        //grab the Act, put it in the output
                        if (paragraph.style === "Act Break") {
                            charSummaries.push("<br /><b>" + paragraph.text + "</b><br />");
                        }
                        // grab selected scene summary
                        if (paragraph.style === "Summary") {
                            charSummaries.push(paragraph.text.substring(0, 100));
                        }
                    }
                    callback(charSummaries);
                    context.sync();
                })
                .catch(function (error) {
                    //showNotification('Error: ' + error.content.join(", "));
                    //showNotification('Error: ' + JSON.stringify(error));
                    ($("#divUserMessage").html(error.message));
                    if (error instanceof OfficeExtension.Error) {
                        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                    }
                });
        })

    }


/******************************************************************************************
 *
 *  EASY CURVE FUNCITON
 *
 *******************************************************************************************/

function easyCurve() {

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

}
