/**
 * This is the "Morning Announcements" script by Tom Hinkle
 * https://www.tomhinkle.net/proj/morning-announcements/
 * Original Script lives here: 
 * https://script.google.com/u/0/home/projects/1h4sY5Fv9-mDYMss4qk68go7wkBMQoEvhl-rKdykPrVetcbsNwYekRzQB/edit
 * 
 * Versions and Notes:
 * 1.1 - Handle copied slides
 *     -> It turns out asking teachers not to copy slides is a BIG ask --
 *        people really like making copies and dragging and dropping to 
 *        move slides. This version implements a system for tracking 
 *        whether a slide has been copied by appending the ID of each
 *        slide to its note sheet. If we detect an ID that does not 
 *        match the slide, we assume it's been copied and treat the slide
 *        as fresh.
 *        
 * 
 * 1.0 - Basic implementation that has been in place for a year as of
 *       January 2023.
 *     Features:
 *     - Slides expire after time
 *     - Slides are marked "new"
 *     - We manage badges.
 *     - We use "Notes" as a settings page for slides.
 */
/** * @OnlyCurrentDoc */
function onOpen() {
  let ui = SlidesApp.getUi().createMenu(
    'Morning Announcements'
  ).addItem(
    'Update','updateAllSlides'
  ).addItem(
    'Add QR Code','addQR'
  ).
  addSeparator()
  .addItem('Set up timer', 'setupAutomation')
  .addItem('Remove timer', 'removeAutomation')  
  .addToUi()
}

function setupAutomation() {
  var ui = SlidesApp.getUi();
  // Initial confirmation prompt
  var initialResponse = ui.prompt(
    'Confirmation Required',
    'Only one user should set up a daily timer to update the slides. If you\'re sure no one else already has a timer running, you should go ahead and set this up; otherwise, you probably shouldn\'t! Type "yes, I\'m sure" to continue.',
    ui.ButtonSet.OK_CANCEL
  );

  // Check if the user confirmed
  if (initialResponse.getSelectedButton() == ui.Button.OK && initialResponse.getResponseText().toLowerCase() == "yes, i'm sure") {
    // Prompt for the hour
    var hourResponse = ui.prompt('Setup Timer', 'Enter the hour (0-24) for the daily timer:', ui.ButtonSet.OK_CANCEL);

    // Process the user's response for the hour
    if (hourResponse.getSelectedButton() == ui.Button.OK) {
      var hour = parseInt(hourResponse.getResponseText());
      if (!isNaN(hour) && hour >= 0 && hour < 24) {
        // Create a new daily trigger at the specified hour
        ScriptApp.newTrigger('updateAllSlides')
          .timeBased()
          .everyDays(1)
          .atHour(hour)
          .create();
        ui.alert(`Timer set for ${hour}:00 daily.`);
      } else {
        ui.alert('Invalid hour. Please enter a number between 0 and 23.');
      }
    }
  } else if (initialResponse.getSelectedButton() == ui.Button.OK) {
    // User typed something other than "yes, I'm sure"
    ui.alert('Confirmation not recognized. Timer setup cancelled.');
  }
}


function removeAutomation() {
  // Get all script triggers
  var allTriggers = ScriptApp.getProjectTriggers();

  for (var i = 0; i < allTriggers.length; i++) {
    // Check if the trigger is for the 'updateAllSlides' function
    if (allTriggers[i].getHandlerFunction() === 'updateAllSlides') {
      // Delete the trigger
      ScriptApp.deleteTrigger(allTriggers[i]);
    }
  }

  // Optionally, notify the user that the timers have been removed
  SlidesApp.getUi().alert('All timers for "updateAllSlides" have been removed.');
}



function addBadge (slide,text, color='#0033a0') {

  let box = slide.insertTextBox(text,475,0,400,75);
  
  box.setRotation(45);
  box.getFill().setSolidFill(color);
  let textObj = box.getText()
  let style = textObj.getTextStyle()
  style.setFontFamily('Cantarell')   
  style.setFontSize(18) 
  style.setForegroundColor('#fefefe');
  box.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE)
  textObj.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER)
  console.log('Made box',box);
  return box.getObjectId();
}

function handleNewBadge (slide, fields) {
  let now = new Date();
  let removeBadge = true;
  if (fields.Highlight && fields.Highlight.getTime() > now.getTime()) {
    removeBadge = false;
    console.log('DO NOT remove badge')
  } else {
    console.log('Remove badge')
  }
  if (!fields.badge && !removeBadge) {
    console.log('Add badge!');
    fields.badge = addBadge(slide,'new')
  }
  if (removeBadge) {
    let theBadge = slide.getShapes().find((shape)=>shape.getObjectId()==fields.badge);
    if (theBadge) {
      console.log('Badge exists!');
      theBadge.remove();
      console.log('Removed badge');
      delete fields.badge;
    }
  }
}

function handleExpiration (slide, fields) {
  let now = new Date();
  if (fields.Expires && !fields.expired) {
    if (fields.Expires.getTime() < now.getTime()) {
      console.log('Expired!');  
      fields.expired = addBadge(slide,'expired','#7f7f7f');
      fields.text.push('This slide has expired. If you want to bring it back to life, delete the contents of these notes and drag it back into the slideshow')
      slide.move(SlidesApp.getActivePresentation().getSlides().length)
    }
  }
}

function updateAllSlides () {
  let now = new Date();
  let nowSeconds = now.getTime();
  let slides = SlidesApp.getActivePresentation().getSlides();
  slides.forEach(
    (slide)=>{      
      let pageId = slide.getObjectId();
      let notes = slide.getNotesPage().getSpeakerNotesShape().getText();
      let currentValue = notes.asString();
      let fields = parseNotesText(currentValue);
      if (fields.id && fields.id != pageId) {
        console.log('A copy: resetting fields!!')
        fields = {}; // reset fields
      }
      fields.id = pageId; // Include pageId in fields!
      if (fields.Permanent || fields.expired) {
        // Skip permanent slides!
        return;
      }
      updateFields(fields);
      handleNewBadge(slide,fields);
      handleExpiration(slide,fields);
      let newText = createNotesText(fields);
      notes.setText(newText);
    }
  )
}

function createNotesText (fields) {
  let text = ''  
  for (let key in fields) {
    if (key != 'text') {
      let value = fields[key];
      if (value instanceof Date) {
        value = value.toLocaleString({weekday:'short',year:'numeric',month:'numeric',day:'numeric'})
      }
      text = `${text}\n${key}:${value}`
    }
  }
  if (fields.text) {
    text = `${text}${fields.text.join("\n")}\n\n`
  }
  return text;
}

function updateFields (fields) {

  if (!fields['Created']) {
    fields['Created'] = new Date();
  }
  if (!fields['Expires']) {
    fields['Expires'] = getExpirationDate(fields);
  }  
  if (!fields['Highlight']) {
    let startDate = fields['Created'];    
    let weekday = startDate.getDay();
    let delta = 1; // default delta is 2 days -- stop highlighting two days after we add it.
    if (weekday == 5) {
      delta += 2;
    }
    if (weekday == 6) {
      dela += 1; // Saturday, add one
    }
    fields['Highlight'] = new Date(startDate.getFullYear(),startDate.getMonth(),startDate.getDate()+delta,startDate.getHours(),startDate.getMinutes());
  }
}

function getExpirationDate (fields) {
  let now = new Date();
  return new Date(now.getFullYear(),now.getMonth(),now.getDate()+7,now.getHours(),now.getMinutes());
}

function parseNotesText (text) {
  let lines = text.split(/\n/);
  let json = {text:[]};
  for (let line of lines) {
    fields = line.split(':');    
    if (fields.length>=2) {
      let field = fields.shift();
      let value = fields.join(':');
      json[field] = value;
    } else {
      if (line) {
        json.text.push(line);
      }
    }
  }
  for (let dateField of ['Created','Expires','Highlight']) {
    if (json[dateField]) {
      json[dateField] = new Date(json[dateField]);
    }
  }
  console.log('Parsed',json)
  return json;
}
