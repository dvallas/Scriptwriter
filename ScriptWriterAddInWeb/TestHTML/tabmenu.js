/* Toggle between adding and removing the "responsive" class to topnav when the user clicks on the icon */
//function myFunction() {
//  var x = document.getElementById("myTopnav");
//  if (x.className === "topnav") {
//    x.className += " responsive";
//  } else {
//    x.className = "topnav";
//  }
//}



function myFunction() {
  var x = document.getElementById("matrix");
  if (x.style.display === "block") {
    x.style.display = "none";
  } else {
    x.style.display = "block";
  }
}

function highlightDiv(e){
    //$( "#test" ).html("background-color:grey")
    $(document).on('mouseover', 'div', function(e) {
    $( "#test" ).html((e.target).getAttribute('id'));
});
}


function printMousePos(e){

      cursorX = e.at;
      cursorY= e.pageY;
    $( "#" + (e.target).getAttribute('id')).html("")
    $( "#" + (e.target).getAttribute('id')).html("x".repeat(cursorY/10))
    $( "#test" ).html( "pageX: " + cursorX +",pageY: " + cursorY );
}