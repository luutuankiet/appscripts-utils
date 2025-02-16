var logMessages = []; // Store log messages globally
var ui = SlidesApp.getUi();

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