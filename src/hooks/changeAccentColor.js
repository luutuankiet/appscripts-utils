function changeAccentColor() {

    var colorInput = ui.prompt("Enter the target text color to update (e.g., black):", Browser.Buttons.OK_CANCEL);
    var colorName = colorInput.getResponseText().trim().toLowerCase();
    var hexColor = colorNameToHex[colorName];

    if (!hexColor) {
        ui.alert("Invalid color name. Please enter a valid color name.");
        return;
    }

    var slides = getSelectedSlides();
    if (slides.length === 0) {
        ui.alert("No slides selected. Please select slides first.");
        return;
    }

    slides.forEach(slide => {
        var slideElements = slide.getPageElements();
        var contentTextShapes = slideElements.map(shape => new TextShape(shape)).filter(shape => shape.hasText);
        contentTextShapes.forEach(shape => {
            if (!shape.isUppermost(slideElements) && !shape.isLink) {
                shape.setTextStyle({
                    foregroundColor: hexColor
                });
            }
        logMessage("Updated color for slide: " + slide.getObjectId());
        });
    });
    ui.alert("Updated text color to " + colorName + " on selected slides.");
}
