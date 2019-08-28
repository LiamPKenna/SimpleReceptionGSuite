const moment = Moment.load();

const GLOBAL = {
  //id of the Google form
  formId : "1tyT79FW71cgHpb8IFnsGbWtj6rIfA2cNqHyI32PR11s",
  //id of the calendar
  calendarId : "stumptowncoffee.com_hgocsc81bh3gfamj2qn94bhjj4@group.calendar.google.com",
  //mapping of form sections for later use
  formMap : {
    eventTitle : "Name of guest",
    startTime : "Date and Time expected",
    hereFor : "Here to see",
    contactBy : "Upon arrival, contact me by",
    otherInfo : "Other information (Where you will be, additional requests)",
    email : "Your email (If you would like a calendar reminder)",
  },
}

function onFormSubmit() {
  const eventObject = getFormResponse();
  const event = createCalendarEvent(eventObject);
}

function getFormResponse() {
  // Get form using id stored in the GLOBAL variable object
  const form = FormApp.openById(GLOBAL.formId),
      //Get all responses from form and return as an array
      responses = form.getResponses(),
      //find the length of the responses array
      length = responses.length,
      //find the index of the most recent form response (0 index = -1)
      lastResponse = responses[length-1],
      //get an array of responses to every question item which received an answer
      itemResponses = lastResponse.getItemResponses(),
      //create empty object to store data from the last form response for later use
      eventObject = {};
  //Loop through each item response in the item response array
  for (const i = 0; i < itemResponses.length; i += 1) {
    //Get the title of the form item being iterated on
    const thisItem = itemResponses[i].getItem().getTitle(),
        //get the submitted response
        thisResponse = itemResponses[i].getResponse();
    //based on the form section title, map the response into eventObject variable
    //use the GLOBAL variable formMap sub object to match section titles to property keys in the event object
    switch (thisItem) {
      case GLOBAL.formMap.eventTitle:
        eventObject.title = thisResponse;
        break;
      case GLOBAL.formMap.startTime:
        eventObject.startTime = thisResponse;
        break;
      case GLOBAL.formMap.hereFor:
        eventObject.hereFor = thisResponse;
        break;
      case GLOBAL.formMap.contactBy:
        eventObject.contactBy = thisResponse;
        break;
      case GLOBAL.formMap.otherInfo:
        eventObject.otherInfo = thisResponse;
        break;
      case GLOBAL.formMap.email:
        eventObject.email = thisResponse;
        break;
    }
  }
  return eventObject;
}

function createCalendarEvent(eventObject) {
  //Get a calendar object by opening the calendar using id stored in the GLOBAL variable object, then populate with values
  const calendar = CalendarApp.getCalendarById(GLOBAL.calendarId),
      title = eventObject.title + " for " + eventObject.hereFor + " by " + eventObject.contactBy,
      startTime = moment(eventObject.startTime).toDate(),
      addEndTime = moment(startTime).add(30, 'minutes'),
      endTime = moment(addEndTime).toDate();
  //an options object containing the description and location for the event
  const options = {
    description : "Here to see " + eventObject.hereFor + ".  Contact by " + eventObject.contactBy + ".  Other information: " + eventObject.otherInfo,
    guests : eventObject.email,
    location: "Reception",
  };
  try {
    //create a calendar event with info stored earlier
    const event = calendar.createEvent(title, startTime,
                                     endTime, options)
    } catch (e) {
      //delete the guest property from the options variable, as an invalid email address with cause this method to throw an error.
      delete options.guests
      //create the event without including the guest
      const event = calendar.createEvent(title, startTime,
                                       endTime, options)
      }
    //add a popup reminder 5 min before the event (change number for more/less time, place // before next line to disable)
    event.addPopupReminder(5)
  return event;
}
