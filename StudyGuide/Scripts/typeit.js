//var myString = "Place your string data here, and as much as you like. The animation is c"+
//                "oming from Javascript! Place your string data here, and as much as you like. The animati"+
//                "on is coming from Javascript! Place your string data here, and as much as you like. The"+
//                "animation is coming from Javascript! Place your string data here, and as much as you li"+
//                "ke. The animation is coming from Javascript! Place your string data here, and as much"+
//                "as you like. The animation is coming from Javascript! ";
var mytextdiv = document.getElementById("chapterstring");
var myString = mytextdiv.innerText;
mytextdiv.setAttribute("style","display:none")
var myArray = myString.split("");
var loopTimer;
// sleep time expects milliseconds

function frameLooper() {
    if (myArray.length > 0) {


        var charToType = myArray.shift();
        if (charToType === '.') {
            document.getElementById("myTypingText").innerHTML += charToType;
            clearTimeout(loopTimer);
            var sleepForABit = setTimeout('subotage()', 700);
            return false;
        }
        // Usage!
        // Do something after the sleep!
        document.getElementById("myTypingText").innerHTML += charToType;

    } else {

        clearTimeout(loopTimer);
        return false;
    }
    loopTimer = setTimeout('frameLooper()', 150);
}
frameLooper();


function subotage() {

    loopTimer = setTimeout('frameLooper()', 150);
}
