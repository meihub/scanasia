//Default, hide all tier two menu
$(".tierTwoMenu").hide();

//Listen in on html class, tierOneMenu.  If clicked then display headers from html data-menu-name
//If tierOneMenu is clicked then assign its data-menu-name to tierTwoMenu variable
// # sign means ID attribute
// . sign means Class attribute
$(".tierOneMenu").click(function() {
    var tierTwoMenu = $(this).attr("data-menu-name");

    //Now we know which headers to show or hide
    //Turn var tierTwoMenu into an ID to grab what we need from HTML
    $(".tierTwoMenu").hide();
    $("#" + tierTwoMenu).show();

});

//Test if scanasia.js is being called by, console.log("aidjfaosdijfpaidsf");
//Default, hide all content section from About Us page
$(".contentBody").hide();

//Listen in on html class, nextArrow.  If clicked then display content from contentBody class
//If nextArrow is clicked then assign its data-menu-name to variable
$(".nextButton").click(function(e) {
    //Prevent default action of original click event therefore, page will not jump and stays still after click
    e.preventDefault();
    var dataMenuName = $(this).attr("data-menu-name");
    console.log("#" + dataMenuName);
    //Now we know which headers to show or hide
    //Turn var contentBody into an ID to grab what we need from HTML
    // # sign means ID attribute
    // . sign means Class attribute
    $(".contentBody").hide();
    console.log($("#" + dataMenuName));
    $("#" + dataMenuName).show();

});