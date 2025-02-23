function applyTOC() {
    var slides = SlidesApp.getActivePresentation().getSlides();
    var slideCount = slides.length;


    // DEBUG MODE
    var slideNumberInput = ui.prompt("Enter reference slide number (1 to " + slideCount + "):", Browser.Buttons.OK_CANCEL);
    var slideNumber = parseInt(slideNumberInput.getResponseText(), 10);

    if (isNaN(slideNumber) || slideNumber < 1 || slideNumber > slideCount) {
        ui.alert("Invalid slide number. Please enter a number between 1 and " + slideCount + ".");
        logMessage("Invalid slide number entered: " + slideNumber);
        return;
    }

    logMessage("Reference slide number: " + slideNumber);

    var referenceSlide = slides[slideNumber - 1];
    // END DEBUG MODE
    
    // referenceSlide = slides[8 - 1]

    var refSlidePageElements = referenceSlide.getPageElements();


    var selectedSlides = getSelectedSlides();
    if (selectedSlides.length === 0) {
        ui.alert("No slides selected. Please select slides first.");
        logMessage("No slides selected for style update.");
        return;
    }

    selectedSlides.forEach(slide => {
            
        // step 1 identify the border box. By default the leftmost area is the first elem.
        var selectedSlideElements = slide.getPageElements();
        var borderBox = selectedSlideElements
            .map(element => new TextShape(element))
            .filter(shape => shape.isTextBox === false)[0];
        
        // grab the border position from its left and width. Appscript slides dont have getRight or getBottom
        var borderBounrdary = borderBox.element.getWidth() - borderBox.left;


        // step 2 identify text inside the border box
        var contentElements = selectedSlideElements
            .map(element => new TextShape(element))
            .filter(shape => shape.left <= borderBounrdary);
        
        
        // step 3 remove em
        contentElements.forEach(element => element.element.remove());

        refSlidePageElements.forEach(elem => {
            slide.insertPageElement(elem);
        });
        logMessage("Cloned shape to slide", slide.getObjectId())
    });
    ui.alert("Updated selected slides with styles from reference slide " + slideNumber + ".");
    logMessage("Updated selected slides with styles from reference slide " + slideNumber + ".");

}