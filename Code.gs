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
  const labelName = e.formInput.labelName;
  if (!labelName) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Please enter a label name.'))
      .build();
  }

  try {
    const threads = GmailApp.getUserLabelByName(labelName).getThreads();
    const sheet = SpreadsheetApp.create(`Emails from ${labelName}`);
    const sheetData = [['datum', 'afzender', 'ontvanger', 'titel', 'tekstinhoud', 'attachments']];

    threads.forEach(thread => {
      const messages = thread.getMessages();
      messages.forEach(message => {
        const attachments = message.getAttachments().map(att => att.getName()).join(', ');
        sheetData.push([
          message.getDate(),
          message.getFrom(),
          message.getTo(),
          message.getSubject(),
          message.getPlainBody(),
          attachments
        ]);
      });
    });

    sheet.getRange(1, 1, sheetData.length, sheetData[0].length).setValues(sheetData);

    const sheetUrl = sheet.getUrl();
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText(`Emails extracted successfully! [View Sheet](${sheetUrl})`))
      .build();

  } catch (error) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText(`Error: ${error.toString()}`))
      .build();
  }
}