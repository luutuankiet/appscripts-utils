var logMessages = []; // Store log messages globally

var ui = SlidesApp.getUi();
function onOpen() {
    ui.createMenu("Utility for imported slides")
        .addItem("Bulk apply style to selected slides","applyStyleFromReferenceSlide")
        .addItem("Bulk apply fonts to selected slides", "changeFontOnSelectedSlides")
        .addItem("Bulk update header text color & size to selected slides", "changeUppermostTextStyle")
        .addItem("Show Debug Logs", "showLogsInSidebar")
        .addToUi();
    logMessage("Menu created with options to apply fonts and update text styles.");
}

// Function to log messages dynamically
function logMessage(message) {
    logMessages.push(message);
    // logMessages.push(new Date().toLocaleTimeString() + " - " + message);
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

// Change font for all text boxes on selected slides
function changeFontOnSelectedSlides() {
    var slides = getSelectedSlides();
    if (slides.length === 0) {
        SlidesApp.getUi().alert("No slides selected. Please select slides first.");
        logMessage("No slides selected for font change.");
        return;
    }

    var fontInput = ui.prompt("Enter font name (e.g., Arial, Google Sans):", Browser.Buttons.OK_CANCEL);
    if (fontInput === "cancel" || fontInput.getResponseText().trim() === "") {
        logMessage("Font change operation cancelled or invalid input.");
        return;
    }

    logMessage("Changing font to: " + fontInput.getResponseText().trim());
    slides.forEach(slide => {
        var shapes = slide.getShapes();
        shapes.forEach(shape => {
                try {
                    if (shape.getShapeType() === SlidesApp.ShapeType.TEXT_BOX) {
                        var text = shape.getText();
                        text.getTextStyle().setFontFamily(fontInput.getResponseText().trim());
                    }
                } catch (err) {
                    logMessage("ERROR: " + err.message);
                }
            });  
        logMessage("Applied font to slide: " + slide.getObjectId());
    });

    SlidesApp.getUi().alert("Font updated to '" + fontInput.getResponseText() + "' on selected slides.");
}

// Mapping of color names to HEX codes
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

// Change the uppermost text element on selected slides with user-defined color & font size
function changeUppermostTextStyle() {
    var slides = getSelectedSlides();
    if (slides.length === 0) {
        SlidesApp.getUi().alert("No slides selected. Please select slides first.");
        logMessage("No slides selected for text style change.");
        return;
    }

    var colorInput = ui.prompt("Enter text color (e.g., red, blue, green):", Browser.Buttons.OK_CANCEL);
    var colorName = colorInput.getResponseText().trim().toLowerCase();
    var hexColor = colorNameToHex[colorName];

    if (!hexColor) {
        SlidesApp.getUi().alert("Invalid color name. Please enter a valid color name.");
        logMessage("Invalid color name entered: " + colorName);
        return;
    }
    logMessage("Color selected: " + colorName + " (HEX: " + hexColor + ")");

    var sizeInput = ui.prompt("Enter font size (e.g., 24):", Browser.Buttons.OK_CANCEL);
    var fontSize = parseInt(sizeInput.getResponseText(), 10);
    if (isNaN(fontSize)) {
        logMessage("Invalid font size entered.");
        return;
    }
    logMessage("Font size selected: " + fontSize);

    slides.forEach(slide => {
        var shapes = slide.getPageElements();
        if (shapes.length > 0) {
            logMessage('Processing shapes on slide.');
            // Sort shapes by Y position (topmost first)
            shapes.sort((a, b) => a.getTop() - b.getTop());
            
            // Get the topmost shape
            var topShape = shapes[0];
            logMessage('Topmost shape identified.');

            if (topShape.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
                var textRange = topShape.asShape().getText();
                textRange.getTextStyle().setForegroundColor(hexColor);
                textRange.getTextStyle().setFontSize(fontSize);
                logMessage('Text style updated for topmost shape.');
            }
        }
    });

    SlidesApp.getUi().alert("Updated uppermost text to color " + colorName + " and font size " + fontSize + ".");
}


function applyStyleFromReferenceSlide() {
    var ui = SlidesApp.getUi();
    var slides = SlidesApp.getActivePresentation().getSlides();
    var slideCount = slides.length;

    // Prompt user for a reference slide number
    var slideNumberInput = ui.prompt("Enter reference slide number (1 to " + slideCount + "):", Browser.Buttons.OK_CANCEL);
    var slideNumber = parseInt(slideNumberInput.getResponseText(), 10);

    if (isNaN(slideNumber) || slideNumber < 1 || slideNumber > slideCount) {
        ui.alert("Invalid slide number. Please enter a number between 1 and " + slideCount + ".");
        logMessage("Invalid slide number entered: " + slideNumber);
        return;
    }

    logMessage("Reference slide number: " + slideNumber);

    // Get the reference slide
    var referenceSlide = slides[slideNumber - 1];
    var referenceFont = null;
    var referenceColor = null;
    var referenceFontSize = null;
    var referencePositionTop = null;
    var referencePositionLeft = null;

    // Step 1: Parse font from the first available text box
    var shapes = referenceSlide.getShapes();
    for (var i = 0; i < shapes.length; i++) {
        var text = shapes[i].getText();
        if (text && text.asString().trim() !== "") {
            referenceFont = text.getTextStyle().getFontFamily();
            logMessage("Parsed font from reference slide: " + referenceFont);
            break;
        }
    }

    // Step 2: Parse the uppermost text's color, font size, and position
    var pageElements = referenceSlide.getPageElements();
    var foundValidText = false;

    if (pageElements.length > 0) {
        pageElements.sort((a, b) => a.getTop() - b.getTop());

        for (var i = 0; i < pageElements.length; i++) {
            var element = pageElements[i];
            logMessage("Checking element type: " + element.getPageElementType());
            if (element.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
                var textRange = element.asShape().getText();
                logMessage("Checking text content: " + textRange.asString().trim());
                if (textRange && textRange.asString().trim() !== "") {
                    referenceColor = textRange.getTextStyle().getForegroundColor();
                    referenceFontSize = textRange.getTextStyle().getFontSize();
                    referencePositionTop = element.getTop();
                    referencePositionLeft = element.getLeft();
                    logMessage("Parsed uppermost text color: " + referenceColor);
                    logMessage("Parsed uppermost text font size: " + referenceFontSize);
                    logMessage("Parsed uppermost text position: " + referencePositionTop + ", " + referencePositionLeft);
                    foundValidText = true;
                    break;
                }
            }
        }
    }

    if (!foundValidText) {
        ui.alert("No valid text found in the reference slide. Please ensure the slide contains text.");
        logMessage("No valid text found in the reference slide. Script terminated.");
        return;
    }

    // Step 3: Clone non-text shapes from the reference slide to selected slides
    var selectedSlides = getSelectedSlides();
    if (selectedSlides.length === 0) {
        ui.alert("No slides selected. Please select slides first.");
        logMessage("No slides selected for style update.");
        return;
    }

    var nonTextShapes = referenceSlide.getPageElements().filter(element => {
        return element.getPageElementType() !== SlidesApp.PageElementType.SHAPE || 
               (element.asShape().getText() && element.asShape().getText().asString().trim() === "");
    });

    selectedSlides.forEach(slide => {
        nonTextShapes.forEach(shape => {
            slide.insertPageElement(shape).sendToBack();
            logMessage("Cloned non-text shape to selected slide.");
        });
    });

    // Step 4: Set all text of selected slides to the reference font
    selectedSlides.forEach(slide => {
        var slideShapes = slide.getShapes();
        slideShapes.forEach(shape => {
            var text = shape.getText();
            if (text && text.asString().trim() !== "" && referenceFont) {
                text.getTextStyle().setFontFamily(referenceFont);
                logMessage("Updated font for shape on slide.");
            }
        });

        // Step 5: Set uppermost text shape of each selected slide to the reference font size, color, and position
        var slideElements = slide.getPageElements();
        slideElements.sort((a, b) => a.getTop() - b.getTop());

        for (var i = 0; i < slideElements.length; i++) {
            var topSlideElement = slideElements[i];
            logMessage("Checking element type");
            if (topSlideElement.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
                var topTextRange = topSlideElement.asShape().getText();
                if (topTextRange && topTextRange.asString().trim() !== "") {
                    if (referenceColor) {
                        topTextRange.getTextStyle().setForegroundColor(referenceColor);
                        logMessage("Updated uppermost text color for slide: " + topTextRange.asString().trim());
                    }
                    if (referenceFontSize) {
                        topTextRange.getTextStyle().setFontSize(referenceFontSize);
                        logMessage("Updated uppermost text font size for slide.");
                    }
                    if (referencePositionTop !== null && referencePositionLeft !== null) {
                        topSlideElement.setTop(referencePositionTop);
                        topSlideElement.setLeft(referencePositionLeft);
                        logMessage("Updated uppermost text position for slide.");
                    }
                    logMessage("Executing break");
                    break; // Exit loop after updating the first text element
                }
            }
        }
    });

    ui.alert("Updated selected slides with styles from reference slide " + slideNumber + ".");
    logMessage("Updated selected slides with styles from reference slide " + slideNumber + ".");
}
