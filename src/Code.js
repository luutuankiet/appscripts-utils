// Import functions from other files
// import { onOpen } from 'utils.js';
// import { changeFontOnSelectedSlides } from './hooks/changeFont.js';
// import { changeUppermostTextStyle } from './hooks/changeTextStyle.js';
// import { applyStyleFromReferenceSlide } from './hooks/applyStyle.js';

// Set up the menu in Google Slides
function onOpen() {
    ui.createMenu("Utility for imported slides")
        .addItem("Bulk apply style to selected slides", "applyStyleFromReferenceSlide")
        .addItem("Bulk apply fonts to selected slides", "changeFontOnSelectedSlides")
        .addItem("Bulk update header text color & size to selected slides", "changeUppermostTextStyle")
        .addItem("Show Debug Logs", "showLogsInSidebar")
        .addToUi();
}