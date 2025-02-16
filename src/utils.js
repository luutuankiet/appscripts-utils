var logMessages = []; // Store log messages globally
var ui = SlidesApp.getUi();


var colorNameToHex = {
    "red": "#FF0000",
    "green": "#00FF00",
    "blue": "#0000FF",
    "yellow": "#FFFF00",
    "black": "#000000",
    "white": "#FFFFFF",
    "gray": "#808080",
    "purple": "#800080",
    "orange": "#FFA500",
    "pink": "#FFC0CB"
    // Add more colors as needed
};

// Function to log messages dynamically
function logMessage(message) {
    logMessages.push(message);
    showLogsInSidebar(); // Refresh logs in sidebar
}

// Display logs in the sidebar
function showLogsInSidebar() {
    var htmlContent = "<h3>Debug Logs</h3><pre>" + logMessages.join("\n") + "</pre>";
    var htmlOutput = HtmlService.createHtmlOutput(htmlContent).setTitle("Debug Logs");
    SlidesApp.getUi().showSidebar(htmlOutput);
}



// Get currently selected slides
function getSelectedSlides() {
    var selection = SlidesApp.getActivePresentation().getSelection();
    var selectedSlides = selection.getPageRange().getPages();
    logMessage("Selected slides count: " + selectedSlides.length);
    return selectedSlides;
}
