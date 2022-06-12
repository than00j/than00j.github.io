(function () {
  "use strict";

  var messageBanner;
  var item;
	
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    //document.getElementById("app-body").style.display = "flex";
    //document.getElementById("run").onclick = run;
	        var element = document.querySelector('.MessageBanner');
      messageBanner = new components.MessageBanner(element);
      messageBanner.hideBanner();
	  showNotification("Testing", "This notification was triggered from Office On ready!");
  }
});

  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function(reason) {
	  item = Office.context.mailbox.item;
    $(document).ready(function () {
      var element = document.querySelector('.MessageBanner');
      messageBanner = new components.MessageBanner(element);
      messageBanner.hideBanner();
	    showNotification("Testing", "LoadProps & getStarttime function will be triggered!");
      loadProps();
	    getStartTime();
      
	    showNotification("Testing", "All functions triggered!");
    });
  };
  
  function getStartTime() {
    item.start.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
				document.getElementById('message').innerText += asyncResult.error.message;
		    showNotification("Error", asyncResult.error.message); //showNotification(header, content)
            }
            else {
                // Successfully got the start time, display it, first in UTC and 
                // then convert the Date object to local time and display that.
                write ('The start time in UTC is: ' + asyncResult.value.toString());
                write ('The start time in local time is: ' + asyncResult.value.toLocaleString());
				document.getElementById('message').innerText += asyncResult.value.toLocaleString();
		    showNotification("Showing start time:", asyncResult.value.toLocaleString()); //showNotification(header, content)
            }
        });
}

  // Take an array of AttachmentDetails objects and build a list of attachment names, separated by a line-break.
  function buildAttachmentsString(attachments) {
    if (attachments && attachments.length > 0) {
      var returnString = "";
      
      for (var i = 0; i < attachments.length; i++) {
        if (i > 0) {
          returnString = returnString + "<br/>";
        }
        returnString = returnString + attachments[i].name;
      }

      return returnString;
    }

    return "None";
  }

  // Format an EmailAddressDetails object as
  // GivenName Surname <emailaddress>
  function buildEmailAddressString(address) {
    return address.displayName + " &lt;" + address.emailAddress + "&gt;";
  }

  // Take an array of EmailAddressDetails objects and
  // build a list of formatted strings, separated by a line-break
  function buildEmailAddressesString(addresses) {
    if (addresses && addresses.length > 0) {
      var returnString = "";

      for (var i = 0; i < addresses.length; i++) {
        if (i > 0) {
          returnString = returnString + "<br/>";
        }
        returnString = returnString + buildEmailAddressString(addresses[i]);
      }

      return returnString;
    }

    return "None";
  }

  // Load properties from the Item base object, then load the
  // message-specific properties.
  function loadProps() {
    //var item = Office.context.mailbox.item.organizer;

    document.getElementById('dateTimeCreated').text = "test"; //item.requiredAttendees);
    $('#dateTimeModified').text(item.start.toLocaleString());
    $('#itemClass').text(item.itemClass);
    $('#itemId').text(item.itemId);
    $('#itemType').text(item.itemType);

    /*$('#message-props').show();

    $('#attachments').html(buildAttachmentsString(item.attachments));
    $('#cc').html(buildEmailAddressesString(item.cc));
    $('#conversationId').text(item.conversationId);
    $('#from').html(buildEmailAddressString(item.organizer.emailAddress ));
    $('#internetMessageId').text(item.internetMessageId);
    $('#normalizedSubject').text(item.normalizedSubject);
    $('#sender').html(buildEmailAddressString(item.sender));
    $('#subject').text(item.subject);
    $('#to').html(buildEmailAddressesString(item.requiredAttendees)); */
  }

  // Helper function for displaying notifications
  function showNotification(header, content) {
	  document.querySelector(".MessageBanner").style.backgroundColor = "red";
    $("#notificationHeader").text(header);
    $("#notificationBody").text(content);
    messageBanner.showBanner();
    messageBanner.toggleExpansion();
  }
  function write(message){
    document.getElementById('message').innerText += message; 
}
})();
