function onMessageComposeHandler(event) {
  //setSubject(event);
}
function onAppointmentComposeHandler(event) {
  setSubject(event);
}
function setSubject(event, d="Message & Appointment Compose Handler Event set this!") {
  Office.context.mailbox.item.subject.setAsync(
    d,
    {
      "asyncContext": eventr
    },
    function (asyncResult) {
      // Handle success or error.
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("Failed to set subject: " + JSON.stringify(asyncResult.error));
      }

      // Call event.completed() after all work is done.
      asyncResult.asyncContext.completed();
    });
}
function setBody(event, d) {
	Office.context.mailbox.item.setSelectedDataAsync(d, function(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    console.log("Selected text has been updated successfully.");
  } else {
    console.error(asyncResult.error);
  }
});
	
	/*
 Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html,function (result) {
        var newHtml = result.value.replace("</body>", "<br/ >"+d+"</body>")

        Office.context.mailbox.item.body.setAsync(newHtml, { coercionType: Office.CoercionType.Html });
    }); */
}
function onAppointmentSendHandler(event) {
  //setSubject(event, "OnAppointmentSendHandler"); //promptOnSend(event);
  let message = 'You must use our add-in to save/send appointment ${Office.context.mailbox.item.itemType}';
  console.log(message);
  sendEvent.completed({ allowEvent: false, errorMessage: message });
  return;
}
function onMessageSendHandler(event) {
  /*Office.context.mailbox.item.body.getAsync(
    "text",
    { asyncContext: event },
    getBodyCallback
  );*/
}

function getBodyCallback(asyncResult){
  let event = asyncResult.asyncContext;
  let body = "";
  if (asyncResult.status !== Office.AsyncResultStatus.Failed && asyncResult.value !== undefined) {
    body = asyncResult.value;
  } else {
    let message = "Failed to get body text";
    console.error(message);
    event.completed({ allowEvent: false, errorMessage: message });
    return;
  }

  let matches = hasMatches(body);
  if (matches) {
    Office.context.mailbox.item.getAttachmentsAsync(
      { asyncContext: event },
      getAttachmentsCallback);
  } else {
    event.completed({ allowEvent: true });
  }
}

function hasMatches(body) {
  if (body == null || body == "") {
    return false;
  }

  const arrayOfTerms = ["send", "picture", "document", "attachment"];
  for (let index = 0; index < arrayOfTerms.length; index++) {
    const term = arrayOfTerms[index].trim();
    const regex = RegExp(term, 'i');
    if (regex.test(body)) {
      return true;
    }
  }

  return false;
}

function getAttachmentsCallback(asyncResult) {
  let event = asyncResult.asyncContext;
  if (asyncResult.value.length > 0) {
    for (let i = 0; i < asyncResult.value.length; i++) {
      if (asyncResult.value[i].isInline == false) {
        event.completed({ allowEvent: true });
        return;
      }
    }

    event.completed({ allowEvent: false, errorMessage: "Looks like you forgot to include an attachment?" });
  } else {
    event.completed({ allowEvent: false, errorMessage: "Looks like you're forgetting to include an attachment?" });
  }
}

function onAppointmentAttendeesChangedHandler(event){

    let message = 'You must use our add-in to save/send appointment ${Office.context.mailbox.item.itemType}';
  console.log(message);
  sendEvent.completed({ allowEvent: false, errorMessage: message });
  return;
}

function onAppointmentTimeChangedHandler(event){
	setBody(event, "Appointment time Changed!")
}

function onAppointmentRecurrenceChangedHandler(event){
	setBody(event, "Recurrence Changed!")
}

// 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
//Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
//Office.actions.associate("OnMessageSend", onMessageSendHandler);
Office.actions.associate("OnAppointmentSend", onAppointmentSendHandler);
Office.actions.associate("OnAppointmentAttendeesChanged", onAppointmentAttendeesChangedHandler);
Office.actions.associate("OnAppointmentTimeChanged", onAppointmentTimeChangedHandler);
Office.actions.associate("OnAppointmentRecurrenceChanged", onAppointmentRecurrenceChangedHandler);