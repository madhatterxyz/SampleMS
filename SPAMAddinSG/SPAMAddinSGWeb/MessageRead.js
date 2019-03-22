//import * as Msal from 'msal';


(function () {
  "use strict";

  var messageBanner;

  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function (reason) {
      $(document).ready(function () {
        
      console.log("Office init");
      var element = document.querySelector('.ms-MessageBanner');
      messageBanner = new fabric.MessageBanner(element);
      messageBanner.hideBanner();
      loadProps();

          
          
    });
  };

    function dialogCallback(asyncResult) {
        if (asyncResult.status == "failed") {

            // In addition to general system errors, there are 3 specific errors for 
            // displayDialogAsync that you can handle individually.
            switch (asyncResult.error.code) {
                case 12004:
                    showNotification("Domain is not trusted");
                    break;
                case 12005:
                    showNotification("HTTPS is required");
                    break;
                case 12007:
                    showNotification("A dialog is already opened.");
                    break;
                default:
                    showNotification(asyncResult.error.message);
                    break;
            }
        }
        else {
            dialog = asyncResult.value;
            /*Messages are sent by developers programatically from the dialog using office.context.ui.messageParent(...)*/
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, messageHandler);

            /*Events are sent by the platform in response to user actions or errors. For example, the dialog is closed via the 'x' button*/
            dialog.addEventHandler(Office.EventType.DialogEventReceived, eventHandler);
        }
    }

    function messageHandler(arg) {
        dialog.close();
        showNotification(arg.message);
    }


    function eventHandler(arg) {

        // In addition to general system errors, there are 2 specific errors 
        // and one event that you can handle individually.
        switch (arg.error) {
            case 12002:
                showNotification("Cannot load URL, no such page or bad URL syntax.");
                break;
            case 12003:
                showNotification("HTTPS is required.");
                break;
            case 12006:
                // The dialog was closed, typically because the user the pressed X button.
                showNotification("Dialog closed by user");
                break;
            default:
                showNotification("Undefined error in dialog window");
                break;
        }
    }

    function openDialog() {
        Office.context.ui.displayDialogAsync(window.location.origin + "/AlertTaskPane.html",
            { height: 50, width: 50 }, dialogCallback);
    }

    function openDialogAsIframe() {
        //IMPORTANT: IFrame mode only works in Online (Web) clients. Desktop clients (Windows, IOS, Mac) always display as a pop-up inside of Office apps. 
        Office.context.ui.displayDialogAsync(window.location.origin + "/Dialog.html",
            { height: 50, width: 50, displayInIframe: true }, dialogCallback);
    }

    
 function getHeader() {

        var soapToGetItemData = getItemDataSoap();

        Office.context.mailbox.makeEwsRequestAsync(soapToGetItemData, soapToGetItemDataCallbackForDelete);
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
    var item = Office.context.mailbox.item;

    $('#dateTimeCreated').text(item.dateTimeCreated.toLocaleString());
    $('#dateTimeModified').text(item.dateTimeModified.toLocaleString());
    $('#itemClass').text(item.itemClass);
    $('#itemId').text(item.itemId);
    $('#itemType').text(item.itemType);

    $('#message-props').show();

    $('#attachments').html(buildAttachmentsString(item.attachments));
    $('#cc').html(buildEmailAddressesString(item.cc));
    $('#conversationId').text(item.conversationId);
    $('#from').html(buildEmailAddressString(item.from));
    $('#internetMessageId').text(item.internetMessageId);
    $('#normalizedSubject').text(item.normalizedSubject);
    $('#sender').html(buildEmailAddressString(item.sender));
    $('#subject').text(item.subject);
    $('#to').html(buildEmailAddressesString(item.to));
    
     



  }

  function callback(asyncResult) {
        var dictionary = asyncResult.value;
        var header1_value = dictionary["header1"];
    }

  // Helper function for displaying notifications
  function showNotification(header, content) {
    $("#notificationHeader").text(header);
    $("#notificationBody").text(content);
    messageBanner.showBanner();
    messageBanner.toggleExpansion();
  }
})();