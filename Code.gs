/**
 * @OnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which it is used, and not
 * all of the user's files. The authorization request message presented
 * to users will reflect this limited scope.
 */

/**
 * Callback for rendering the main card.
 * @param {Object} e The event object.
 * @return {CardService.Card} The card to show to the user.
 */
function onGmailContext(e) {
  return createExtractorCard();
}

/**
 * Creates a card that allows the user to start the email extraction process.
 * @return {CardService.Card} The card to show to the user.
 */
function createExtractorCard() {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Mail Extractor'))
    .addSection(
      CardService.newCardSection()
        .addWidget(CardService.newTextInput()
          .setFieldName('labelName')
          .setTitle('Enter Gmail Label')
        )
        .addWidget(CardService.newTextButton()
          .setText('Extract Emails')
          .setOnClickAction(CardService.newAction().setFunctionName('extractEmails'))
        )
    )
    .build();
}

/**
 * Extracts emails from the specified label and saves them to a new Google Sheet.
 * @param {Object} e The event object.
 */
function extractEmails(e) {
  // Set a maximum number of threads to process to avoid script timeouts (30-second limit).
  const MAX_THREADS_TO_PROCESS = 500;
  let notificationText = '';

  const labelName = e.formInput.labelName;
  if (!labelName) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Please enter a label name.'))
      .build();
  }

  try {
    const label = GmailApp.getUserLabelByName(labelName);

    if (!label) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText(`Error: The label "${labelName}" was not found.`))
        .build();
    }

    const threads = label.getThreads();

    if (threads.length === 0) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText(`The label "${labelName}" contains no emails to extract.`))
        .build();
    }

    // Warn the user if the thread count exceeds the processing limit.
    if (threads.length > MAX_THREADS_TO_PROCESS) {
      notificationText = `Warning: This label has ${threads.length} threads, but only the first ${MAX_THREADS_TO_PROCESS} will be processed to prevent timeouts.\n\n`;
    }

    const threadsToProcess = threads.slice(0, MAX_THREADS_TO_PROCESS);
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const sheet = SpreadsheetApp.create(`Emails from ${labelName} ${timestamp}`);
    const sheetData = [['datum', 'afzender', 'ontvanger', 'titel', 'tekstinhoud', 'attachments']];

    threadsToProcess.forEach(thread => {
      const messages = thread.getMessages();
      messages.forEach(message => {
        // Ensure data exists before pushing, preventing errors from malformed emails.
        const attachments = message.getAttachments().map(att => att.getName()).join(', ');
        sheetData.push([
          message.getDate() || new Date(),
          message.getFrom() || '',
          message.getTo() || '',
          message.getSubject() || '(No Subject)',
          message.getPlainBody() || '',
          attachments || ''
        ]);
      });
    });

    sheet.getRange(1, 1, sheetData.length, sheetData[0].length).setValues(sheetData);

    const sheetUrl = sheet.getUrl();
    const openLinkAction = CardService.newOpenLink().setUrl(sheetUrl);

    notificationText += `Extracted ${sheetData.length - 1} emails successfully!`;

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText(notificationText))
      .setOpenLink(openLinkAction)
      .build();

  } catch (error) {
    console.error(`Error in extractEmails for label "${labelName}": ${error.stack}`);
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText(`An unexpected error occurred. Please check the script logs for details.`))
      .build();
  }
}