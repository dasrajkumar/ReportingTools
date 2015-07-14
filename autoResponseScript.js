//@author: Adonis Settouf (aka Scriptator)
//@mail: asettouf@lexmark.com

//Spreadsheet to write the numbers
var sps = SpreadsheetApp.getActiveSheet();
//mailbox mailaddress
var alias = Session.getActiveUser().getEmail();
//Generic auto response
var autoResponseGen = "Thank you for contacting the Lexmark Solution Support. This is an automatic response. Your email has been received and will be processed promptly.";
//French auto response
var autoResponseFr = "Nous vous remercions d'avoir contacté l'assistance aux solutions Lexmark. Ceci est une réponse automatique. Votre courriel a bien été reçu  et sera traité dans les meilleurs délais.";
//auto responseGerman
var autoResponseDe = "Vielen Dank, dass Sie den Lexmark Solution Support kontaktiert haben. Dies ist eine automatische Antwort. Ihre E-Mail ist angekommen und wird zeitnah bearbeitet.";
//adding mail addresses we don't want to send an automatic reply
var forbiddenList = ["lfmhelp@lexmark.com","apps-scripts-notifications@google.com", "followme@ringdale.com", "support@pharos.com"
                    ,"Identity.Manager@lexmark.com", "Postmaster@lexmark.com", "noreply@wetransfer.com"];


//watch out when testing, possible spam!
function anAnswerMail(){
  //!!!!Some (basic) computations might be necessary to have a correct number
  var threadsToCheck = GmailApp.getInboxThreads(0, 11);
  var i = 0;
  var autoResponse = "";
  var sender = "";
  for(i; i < threadsToCheck.length;i++){
    thread = threadsToCheck[i];
    Logger.log(thread.getFirstMessageSubject());
    if (isSendAutoResponse_(thread)){
      msg = thread.getMessages()[0];
      languageMessage = stripGetFrom_(msg.getTo()).split("@")[1].split(".")[1];
      sender = msg.getFrom();
      recipient = msg.getTo()
      //Logger.log(languageMessage);
      //recipient = "asettouf@lexmark.com";
      //recipient = msg.getTo();
      if (languageMessage.indexOf("de") != -1 || languageMessage.indexOf("at") != -1){
        autoResponse = autoResponseDe;
      } else if(languageMessage.indexOf("fr") != -1 || languageMessage.indexOf("eu") != -1 || languageMessage.indexOf("be") != -1){
        autoResponse = autoResponseFr;
      } else if(languageMessage.indexOf("ch") != -1){
        autoResponse = autoResponseFr + "<br/><br/>" + autoResponseDe;
      } else if(languageMessage.indexOf("com") != -1){
        autoResponse = autoResponseGen + "<br/><br/>" + autoResponseFr + "<br/><br/>" + autoResponseDe;
      } else {
        autoResponse = autoResponseGen;
      }
      if (isToSend_(sender) && isToSend_(recipient)){
        sendGmailTemplate_(thread, autoResponse);
        //fingAlban_(thread);
      }
    }
  }
}

//Loop to see if sender is in forbidden list
function isToSend_(sender){
  var i = 0;
  for (i; i < forbiddenList.length; i++){
    if (sender.indexOf(forbiddenList[i]) != -1){
      return false;
    }
  }
  return true;
}

//In your face A.!
function fingAlban_(thread){
  if (thread.isUnread()) {
    var alban = GmailApp.getUserLabelByName("Alban");
    thread.addLabel(alban);
  } 
}


function isSendAutoResponse_(thread){
  var i = 0
  var msgs = thread.getMessages();
  if (msgs === null) {
    return false;
  } else {
    for (i; i< msgs.length; i++){
      msg = msgs[i];
      if (stripGetFrom_(msg.getFrom()).indexOf(alias) != -1){
        return false;
      }
    }
    return true;
  }
   
}

//Strip a sender from unnecessary signs
function stripGetFrom_(sender){
  return  (sender.indexOf("<") != -1) ? sender.split("<")[1].split(">")[0]: sender;
}
   
//Note of the author: thanks to stackoverflow and user Mogsdad for the following two functions that basically add
//the signature to the email, see http://stackoverflow.com/questions/18493808/gmail-sending-emails-from-spreadsheet-how-to-add-signature-with-image
/**
 * Insert the given email body text into an email template, and send
 * it to the indicated recipient. The template is a draft message with
 * the subject "TEMPLATE"; if the template message is not found, an
 * exception will be thrown. The template must contain text indicating
 * where email content should be placed: {BODY}.
 *
 * @param {String} recipient  Email address to send message to.
 * @param {String} subject    Subject line for email.
 * @param {String} body       Email content, may be plain text or HTML.
 * @param {Object} options    (optional) Options as supported by GmailApp.
 *
 * @returns        GmailApp   the Gmail service, useful for chaining
 */
function sendGmailTemplate_(thread, body, options) {
  options = options || {};  // default is no options
  var drafts = GmailApp.getDraftMessages();
  var found = false;
  for (var i=0; i<drafts.length && !found; i++) {
    if (drafts[i].getSubject() == "TEMPLATE") {
      found = true;
      var template = drafts[i];
    }
  }
  if (!found) throw new Error( "TEMPLATE not found in drafts folder" );

  // Generate htmlBody from template, with provided text body
  var imgUpdates = updateInlineImages_(template);
  options.htmlBody = imgUpdates.templateBody.replace('{BODY}', body);
  options.attachments = imgUpdates.attachments;
  options.inlineImages = imgUpdates.inlineImages;
  return thread.reply(body, options);
}


/**
 * This function was adapted from YetAnotherMailMerge by Romain Vaillard.
 * Given a template email message, identify any attachments that are used
 * as inline images in the message, and move them from the attachments list
 * to the inlineImages list, updating the body of the message accordingly.
 *
 * @param   {GmailMessage} template  Message to use as template
 * @returns {Object}                 An object containing the updated 
 *                                   templateBody, attachments and inlineImages.
 */
function updateInlineImages_(template) {
  //////////////////////////////////////////////////////////////////////////////
  // Get inline images and make sure they stay as inline images
  //////////////////////////////////////////////////////////////////////////////
  var templateBody = template.getBody();
  var rawContent = template.getRawContent();
  var attachments = template.getAttachments();

  var regMessageId = new RegExp(template.getId(), "g");
  if (templateBody.match(regMessageId) != null) {
    var inlineImages = {};
    var nbrOfImg = templateBody.match(regMessageId).length;
    var imgVars = templateBody.match(/<img[^>]+>/g);
    var imgToReplace = [];
    if(imgVars != null){
      for (var i = 0; i < imgVars.length; i++) {
        if (imgVars[i].search(regMessageId) != -1) {
          var id = imgVars[i].match(/realattid=([^&]+)&/);
          if (id != null) {
            var temp = rawContent.split(id[1])[1];
            temp = temp.substr(temp.lastIndexOf('Content-Type'));
            var imgTitle = temp.match(/name="([^"]+)"/);
            if (imgTitle != null) imgToReplace.push([imgTitle[1], imgVars[i], id[1]]);
          }
        }
      }
    }
    for (var i = 0; i < imgToReplace.length; i++) {
      for (var j = 0; j < attachments.length; j++) {
        if(attachments[j].getName() == imgToReplace[i][0]) {
          inlineImages[imgToReplace[i][2]] = attachments[j].copyBlob();
          attachments.splice(j, 1);
          var newImg = imgToReplace[i][1].replace(/src="[^\"]+\"/, "src=\"cid:" + imgToReplace[i][2] + "\"");
          templateBody = templateBody.replace(imgToReplace[i][1], newImg);
        }
      }
    }
  }
  var updatedTemplate = {
    templateBody: templateBody,
    attachments: attachments,
    inlineImages: inlineImages
  }
  return updatedTemplate;
}