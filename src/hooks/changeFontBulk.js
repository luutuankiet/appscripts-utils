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
        var slideElements = slide.getPageElements();
        var textShapes = slideElements.map(element => new TextShape(element)).filter(shape => shape.hasText);

        textShapes.forEach(shape => {
            shape.setTextStyle({ fontFamily: fontInput.getResponseText().trim() });
            logMessage("Applied font to slide: " + slide.getObjectId());
        });
    });

    SlidesApp.getUi().alert("Font updated to '" + fontInput.getResponseText() + "' on selected slides.");
}
