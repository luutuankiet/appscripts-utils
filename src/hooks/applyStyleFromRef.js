function applyStyleFromReferenceSlide() {
    var slides = SlidesApp.getActivePresentation().getSlides();
    var slideCount = slides.length;

    var slideNumberInput = ui.prompt("Enter reference slide number (1 to " + slideCount + "):", Browser.Buttons.OK_CANCEL);
    var slideNumber = parseInt(slideNumberInput.getResponseText(), 10);

    if (isNaN(slideNumber) || slideNumber < 1 || slideNumber > slideCount) {
        ui.alert("Invalid slide number. Please enter a number between 1 and " + slideCount + ".");
        logMessage("Invalid slide number entered: " + slideNumber);
        return;
    }

    logMessage("Reference slide number: " + slideNumber);

    var referenceSlide = slides[slideNumber - 1];
    var referenceTextShape = null;

    var refSlidePageElements = referenceSlide.getPageElements();
    var textShapes = refSlidePageElements.map(element => new TextShape(element)).filter(shape => shape.hasText);

    if (textShapes.length > 0) {
        referenceTextShape = textShapes[0];
        logMessage("Parsed font from reference slide: " + referenceTextShape.getTextStyle().getFontFamily());
    } else {
        ui.alert("No valid text found in the reference slide. Please ensure the slide contains text.");
        logMessage("No valid text found in the reference slide. Script terminated.");
        return;
    }

    var selectedSlides = getSelectedSlides();
    if (selectedSlides.length === 0) {
        ui.alert("No slides selected. Please select slides first.");
        logMessage("No slides selected for style update.");
        return;
    }

    selectedSlides.forEach(slide => {
        var slideElements = slide.getPageElements();
        var slideTextShapes = slideElements.map(element => new TextShape(element)).filter(shape => shape.hasText);


        // process for text based shapes
        slideTextShapes.forEach(shape => {
            if (shape.isUppermost(slideElements)) {
                shape.setTextStyle({
                    fontFamily: referenceTextShape.getTextStyle().getFontFamily(),
                    foregroundColor: referenceTextShape.getTextStyle().getForegroundColor(),
                    fontSize: referenceTextShape.getTextStyle().getFontSize()
                });
                shape.setPosition(referenceTextShape.top, referenceTextShape.left);
            } else {
                shape.setTextStyle({
                    fontFamily: referenceTextShape.getTextStyle().getFontFamily(),
                });
            }
        });

        // process for non-text based shapes
        var nonTextElem = refSlidePageElements.map(element => new TextShape(element)).filter(element => !element.isTextBox);
        nonTextElem.forEach(elem => {
            logMessage("targetting " + elem.element.getPageElementType() + elem.element.getObjectId());
            slide.insertPageElement(elem.element).sendToBack();
            logMessage("Cloned non-text shape to selected slide.");
        });
    });
    ui.alert("Updated selected slides with styles from reference slide " + slideNumber + ".");
    logMessage("Updated selected slides with styles from reference slide " + slideNumber + ".");
}