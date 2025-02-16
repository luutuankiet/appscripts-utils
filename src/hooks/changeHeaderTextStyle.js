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
        logMessage("Processing slide: " + slide.getObjectId());
        var selectedSlideElements = slide.getPageElements();
        var textShapes = selectedSlideElements.map(element => new TextShape(element)).filter(shape => shape.hasText);

        textShapes.forEach(shape => {
            if (shape.isUppermost(selectedSlideElements)) {
                shape.setTextStyle({
                    foregroundColor: hexColor,
                    fontSize: fontSize
                });
                logMessage('Text style updated for topmost shape.');
            } else {
                logMessage('Skipping shape that is not the topmost.' + shape.textRange.asString());
            }
            
        });
    });

    SlidesApp.getUi().alert("Updated uppermost text to color " + colorName + " and font size " + fontSize + ".");
}
