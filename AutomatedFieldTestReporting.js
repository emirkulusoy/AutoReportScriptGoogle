////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Authors     : Emir K Ulusoy
// Email       : emir.kursad.ulusoy@gmail.com 
// Information : This JavaScript creates Automated Field Test Reporting
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// Constants and hard-coded variables
const TEMPLATE_ID = "please update based on your template file id"; // Create your tempate docx file on google drive and get the ID of the file
const FOLDER_NAME = 'please update based on your directory selection'; // directory where you would like to move your files
const EMAIL_NAME = "Automated Field Test Reporting Project"; // you can set any email sender name 
const MAX_IMAGE_WIDTH = 640; // optimizing size of pictures
const SITENAME_REF_DOCX =  'XXXXXX' // temporary ref key word for site name at the tempate docx
const DATE_REF_DOCX = 'YYYYYY' // temporary ref key word for date at the tempate docx
const SOW_REF_DOCX = 'ZZZZZZ' // temporary ref key word for SoW at the tempate docx
const HEADER1 = ["Timestamp", "Email Address", "Site Name ", "Is it a pre-log-in or post-log-out activity?  ", "Type of Activity  ", "Scope of Work Details  ", "Activity Order Number", "Service-effecting or not Service-effecting", "Site Photo", "Equipment Photo"];


/**
 * The main function triggers on the "form submission" step. You should go to Triggers (script.google.com) and add a new trigger.
 *    Edit Trigger for AutomatedFieldTestReporting
 *    Choose which function to run: onFormSubmit
 *    Which runs at deployment: Head
 *    Select event source: Form spreadsheet
 *    Select event type: On form submit
 *    Failure notification settings: notify me immediately
 * This function manages the primary process of creating a report, emailing it, and saving it to Google Drive.
 * @param {Object} forminputs - The form submission event object. It collects required values via form.
 */
function onFormSubmit(forminputs) {
  var values = forminputs.namedValues; 
  var path1 = generatePath(values);
  var DateStamp = Utilities.formatDate(new Date(), "GMT-8", "MM/dd/yyyy");
  
  var folderid = getFolderId(FOLDER_NAME);
  var subfolderid = createSubFolder(folderid, path1);
  
  var doc01id = copyTemplate(subfolderid, path1);
  var doc01 = DocumentApp.openById(doc01id);
  var body01 = doc01.getBody();
  
  replacePlaceholders(body01, values, DateStamp);
  var htmlBody1 = '<ul>';
  htmlBody1 = appendDataToDocument(body01, values, subfolderid,htmlBody1);

  finalizeDocument(doc01, subfolderid, path1, values["Email Address"], htmlBody1, values); // you can use any email distro list
}

/**
 * Gets the ID of a Google Drive folder by name.
 * @param {string} folderName - The name of the folder.
 * @return {string} The ID of the folder.
 */
function getFolderId(folderName) {
  return DriveApp.getFoldersByName(folderName).next().getId();
}

/**
 * Creates a subfolder within a given parent folder and returns its ID.
 * This subfolder structure helps to easily manage files in the future.
 * @param {string} parentFolderId - The ID of the parent folder.
 * @param {string} subFolderName - The name of the subfolder to create.
 * @return {string} The ID of the subfolder.
 */
function createSubFolder(parentFolderId, subFolderName) {
  var parentFolder = DriveApp.getFolderById(parentFolderId);
  parentFolder.createFolder(subFolderName);
  return DriveApp.getFoldersByName(subFolderName).next().getId();
}

/**
 * Copies a template file into a specified folder and returns the ID of the new file.
 * We have a docx template file, so we can easily manage a standardized format for pdf. 
 * Not to overwrite on it, we are creating a copy.
 * @param {string} subFolderId - The ID of the subfolder where the new file will be created.
 * @param {string} newName - The name of the new file.
 * @return {string} The ID of the copied file.
 */
function copyTemplate(subFolderId, newName) {
  return DriveApp.getFileById(TEMPLATE_ID).makeCopy(newName, DriveApp.getFolderById(subFolderId)).getId();
}

/**
 * Generates a path string based on form submission values.
 * Even though this part is hard-coded, we can say any field activity will have these key words.
 * @param {Object} values - The form submission values.
 * @return {string} The generated path string.
 */
function generatePath(values) {
  return `${values["Site Name"]}_${values["Is it a pre-log-in or post-log-out activity?"]}_${values["Type of Activity"]}_${values["Timestamp"]}`;
}

/**
 * Replaces placeholders in the document body with actual values from the form submission.
 * Even though this part is hard-coded based on template, we can say any field activity will have these key words.
 * @param {Object} body - The document body.
 * @param {Object} values - The form submission values.
 * @param {string} dateStamp - The current date stamp.
 */
function replacePlaceholders(body, values, dateStamp) {
  body.replaceText(SITENAME_REF_DOCX, values["Site Name"]);
  body.replaceText(DATE_REF_DOCX, dateStamp);
  body.replaceText(SOW_REF_DOCX, `${values["Is it a pre-log-in or post-log-out activity?"]}_${values["Type of Activity"]}`);
}

/**
 * Adding pictures and text to pdf file and email body(only text).
 * @param {Object} body - The document body.
 * @param {Object} values - The form submission values.
 * @param {string} subfolderId - The ID of the subfolder where photos will be moved.
 */
function appendDataToDocument(body, values, subfolderId,htmlBody1) {
  HEADER1.forEach(function (key) {
    var data = values[key];
    if (data) {
      if (key.toLowerCase().includes("photo")) {
        appendPhotos(body, data, key, subfolderId);
      } else {
        body.appendParagraph(`${key}: ${data}`).setHeading(DocumentApp.ParagraphHeading.HEADING2);
        htmlBody1 += '<li>' + `${key}: ${data}` + '</li>';
      }
    }
  });
  return htmlBody1
}

/**
 * Appends photos to the pdf file
 * @param {Object} body - The document body.
 * @param {string} data - The photo URLs.
 * @param {string} label - The label for the photos.
 * @param {string} subfolderId - The ID of the subfolder where photos will be moved.
 */
function appendPhotos(body, data, label, subfolderId) {
  var fileUrls = data.split(",");
  fileUrls.forEach(function (fileUrl, index) {
    body.appendParagraph(`${label} [${index}]`).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    var fileId = extractFileId(fileUrl);
    //Logger.log(fileId, subfolderId, `${label}_${index}`)
    moveFiles(fileId, subfolderId, `${label}_${index}`);
    var blob = DriveApp.getFileById(fileId).getBlob();
    var inlineImage = body.appendImage(blob);
    resizeImage(inlineImage);
  });
}

/**
 * Extracts the file ID from a Google Drive file URL.
 * When the form is submitted, all pictures are uploaded to Google Drive. 
 * Google Forms return the full link to the photos. 
 * So we parse it to get the ID.
 * @param {string} fileUrl - The URL of the file.
 * @return {string} The extracted file ID.
 */
function extractFileId(fileUrl) {
  var index = fileUrl.indexOf('=') + 1;
  return fileUrl.substring(index);
}

/**
 * Moves a file to a specified folder and renames it.
 * Just organizing the folder structure to make it more efficient.
 * @param {string} sourceFileId - The ID of the source file.
 * @param {string} targetFolderId - The ID of the target folder.
 * @param {string} targetName - The new name for the file.
 */
function moveFiles(sourceFileId, targetFolderId, targetName) {
  var file = DriveApp.getFileById(sourceFileId);
  var targetFolder = DriveApp.getFolderById(targetFolderId);
  file.moveTo(targetFolder);//.getParents().next().removeFile(file);
  file.setName(targetName);//DriveApp.getFolderById(targetFolderId).addFile(file).setName(targetName);
}

/**
 * Resizes an image to fit within a maximum width of 640 pixels.
 * Just a make-up session to make sure the pictures are located well on the page.
 * @param {Object} image - The image to resize.
 */
function resizeImage(image) {
  var width = image.getWidth();
  var height = image.getHeight();
  var ratio = width / height;
  if (width > MAX_IMAGE_WIDTH) {
    if (width > height) {
      image.setWidth(MAX_IMAGE_WIDTH).setHeight(MAX_IMAGE_WIDTH / ratio);
    } else {
      image.setHeight(MAX_IMAGE_WIDTH).setWidth(MAX_IMAGE_WIDTH * ratio);
    }
  }
}

/**
 * Creates the HTML body content for the email.
 * This function simply adds all the text content from the form submission to the email body. 
 * This step improves any email search based on key words.
 * @param {Object} values - The form submission values.
 * @return {string} The generated HTML body content.
 */
function createHtmlBody(htmlBody1) {
  // Implement the function to generate HTML body
  //htmlBody1 += '<li>' + label + ": " +  data + '</li>';
  return '<ul></ul>'; // Placeholder implementation
}

/**
 * Finalizes the document by saving it, creating a PDF, and sending an email with the PDF attached.
 * @param {Object} doc - The Google Document object.
 * @param {string} subfolderId - The ID of the subfolder where the PDF will be saved.
 * @param {string} path - The path used for naming the PDF.
 * @param {string} recipient - The email address of the recipient.
 * @param {string} htmlBody - The HTML body content for the email.
 */
function finalizeDocument(doc, subfolderId, path, recipient, htmlBody, values) {
  doc.saveAndClose();
  var pdf = doc.getBlob().setName(`${path}.pdf`);
  var folder = DriveApp.getFolderById(subfolderId);
  folder.createFile(pdf);
  
  MailApp.sendEmail({
    name: EMAIL_NAME,
    to: recipient, // you can use any email distro list
    subject: `${values["Site Name"]} Site Visit for ${values["Is it a pre-log-in or post-log-out activity?"]}-${values["Type of Activity"]} on ${values["Timestamp"]}`,
    htmlBody: htmlBody,
    attachments: [pdf]
  });
}
