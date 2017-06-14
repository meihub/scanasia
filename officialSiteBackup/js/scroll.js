//We wrap all the code in an object so that it doesn't interfere with any other code
var scroller = {
  init:   function() {

    //collect the variables
    scroller.docH = document.getElementById("media_content").offsetHeight;
    scroller.contH = document.getElementById("content_scroll").offsetHeight;
    scroller.scrollAreaH = document.getElementById("scrollArea").offsetHeight;
    scroller.scrollH = 105;
      
    //calculate height of scroller and resize the scroller div
    //(however, we make sure that it isn't to small for long pages)
    // scroller.scrollH = (scroller.contH * scroller.scrollAreaH) / scroller.docH;
    //if(scroller.scrollH < 15) scroller.scrollH = 15;
    //document.getElementById("scroller").style.height = Math.round(scroller.scrollH) + "px";
    
    //what is the effective scroll distance once the scoller's height has been taken into account
    scroller.scrollDist = Math.round(scroller.scrollAreaH-scroller.scrollH);
    
    //make the scroller div draggable
    Drag.init(document.getElementById("scroller"),null,0,0,-1,scroller.scrollDist);
    
    //add ondrag function
    document.getElementById("scroller").onDrag = function (x,y) {
      if ($("#media_content").height() > $("#content_scroll").height()){
      	 var scrollY = parseInt(document.getElementById("scroller").style.top);
      var docY = 0 - (scrollY * (scroller.docH - scroller.contH) / scroller.scrollDist);
      	document.getElementById("media_content").style.top = docY + "px";
      }
    }
  }
}
