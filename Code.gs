// Developed by Dhaval in 2021 (dhavalsoneji.com)


var CONTENTS_PAGE_ID = 'CONTENTS_PAGE_ID';
let presentation = SlidesApp.getActivePresentation();
let slides = presentation.getSlides();
const PAGE_HEIGHT = presentation.getPageHeight();
const PAGE_WIDTH = presentation.getPageWidth();
const BOX_HEIGHT = 30;
const BOX_WIDTH = ( PAGE_WIDTH - 100 ) / 2;
const BACK_IMAGE_URL = "https://aux.iconspalace.com/uploads/135210286.png"

// https://stackoverflow.com/a/1418059
if(typeof(String.prototype.trim) === "undefined")
{
    String.prototype.trim = function() 
    {
        return String(this).replace(/^\s+|\s+$/g, '');
    };
}


/**
 * Runs when the add-on is installed.
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall(e) {
  onOpen();
}

/**
 * Trigger for opening a presentation.
 * @param {object} e The onOpen event.
 */
function onOpen(e) {
  SlidesApp.getUi().createAddonMenu()
      .addItem('Generate Contents Page', 'create')
      .addItem('Clear Contents Page and Back Buttons', 'clear')
      .addToUi();
}

/**
 * Gets the slide number of the contents page
 * @returns {number} Slide Number
 */
function getContentsPageNumber(){
  for (let i = 0; i < slides.length; ++i) {
    let slide = slides[i];
    let notes = slide.getNotesPage().getSpeakerNotesShape().getText().asString()
    // console.log(notes);
    if ( notes.startsWith("CONTENTS_PAGE") ) {
      return i
    }
  }
  throw("Error: 'CONTENTS_PAGE' note not found")
}

/**
 * Main Script to create the contents pages and back buttons
 */
function create() {

  START = getContentsPageNumber();
  clear(START);

  /**
   * Make a textbox
   * @param {String} text text to be displayed
   * @param {number} x coordinate
   * @param {number} y coordinate
   * @param {number} the start slide number 
   */
  function makeBox(text, x, y, link=START) {
    const contents_page = slides[START];
    let box = contents_page.insertTextBox(text, x, y, BOX_WIDTH, BOX_HEIGHT);
    box.getText().getTextStyle().setForegroundColor('#ffffff');
    box.setLinkSlide(link);
  }

  // find half point
  var half = ( ( ( slides.length - START ) / 2) >> 0) + START ;

  // generate first half of contents page
  for (let i = START + 1; i < half; ++i) {
    let slide = slides[i];
    let notes = slide.getNotesPage().getSpeakerNotesShape().getText().asString().split('\n')[0];
    if (notes.startsWith("Title: ")) {
      title = notes.trim().split('Title: ')[1];
      makeBox(title, 40, 30 + (i-START)*25, i);
    }
  }
  
  // generate second half of contents page
  for (let i = half; i < slides.length; ++i) {
    let slide = slides[i];
    let notes = slide.getNotesPage().getSpeakerNotesShape().getText().asString().split('\n')[0];
    if (notes.startsWith("Title: ")) {
      title = notes.trim().split('Title: ')[1];
      makeBox(title, PAGE_WIDTH/2 + 40, 30 + (i - half)*25, i);
    }
  }
  
  // generate back buttons
  for (let i = START + 1; i < slides.length; ++i) {
    let slide = slides[i];
    let back = slide.insertImage(BACK_IMAGE_URL,0,0,15,15);
    back.setLinkSlide(START);
  }

}
/**
 * Function to reset everything
 * Removes all contents page links
 * Removes all back buttons
 */
function clear(START = getContentsPageNumber()) {

  // Removes all contents page links
  const contents_page = slides[START];
  const elements = contents_page.getPageElements();
  for (let i = 0; i < elements.length; ++i) {
    let el = elements[i];
    if (el.getPageElementType() === SlidesApp.PageElementType.SHAPE && el.asShape().getLink() ) {
      el.remove();
    }
  }

  // Removes all back buttons
  for (let i = START + 1; i < slides.length; ++i) {
    let slide = slides[i];
    const elements = slide.getPageElements();
    for (let i = 0; i < elements.length; ++i) {
      let el = elements[i];
      if (el.getPageElementType() === SlidesApp.PageElementType.IMAGE && el.asImage().getSourceUrl() === BACK_IMAGE_URL ) {
        el.remove();
      }
    }
  }
}












































// if (barWidth > 0) {
//   var bar = slides[i].insertShape(SlidesApp.ShapeType.RECTANGLE, x, y,
//                                   barWidth, BAR_HEIGHT);
//   bar.getBorder().setTransparent();
//   bar.setLinkUrl(BAR_ID);
// }
// console.log(barWidth);





































