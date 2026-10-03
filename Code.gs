const CONTACT_EMAIL_RECIPIENTS = [
  'pandeaniket57@gmail.com',
  'atulpaikrao1114@gmail.com'
];

function notifyContactEmails(event) {
  if (!event || !event.response) {
    throw new Error('Run this function from a Google Form submission trigger.');
  }

  const response = event.response;
  const answers = {};

  response.getItemResponses().forEach(function(itemResponse) {
    answers[itemResponse.getItem().getTitle()] = itemResponse.getResponse();
  });

  let collectedEmail = '';
  if (typeof response.getRespondentEmail === 'function') {
    collectedEmail = response.getRespondentEmail() || '';
  }

  const name = toText(answers['Your name']) || 'Not provided';
  const email = toText(answers['Your email address']) ||
    toText(answers['Email']) ||
    toText(collectedEmail);
  const message = toText(answers['Message']) || 'No message provided';
  const subjectName = name.replace(/[\r\n]+/g, ' ').slice(0, 100);
  const subject = 'DLIGHT FOODS contact message from ' + subjectName;
  const body = [
    'New message submitted through the DLIGHT FOODS contact form.',
    '',
    'Name: ' + name,
    'Email: ' + (email || 'Not provided'),
    '',
    'Message:',
    message,
    '',
    'Submitted: ' + response.getTimestamp()
  ].join('\n');

  const options = {};
  if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    options.replyTo = email;
  }

  MailApp.sendEmail(
    CONTACT_EMAIL_RECIPIENTS.join(','),
    subject,
    body,
    options
  );
}

function installContactEmailTrigger() {
  const form = FormApp.getActiveForm();
  if (!form) {
    throw new Error('Open this script from the Google Form, then run this function.');
  }

  const alreadyInstalled = ScriptApp.getProjectTriggers().some(function(trigger) {
    return trigger.getHandlerFunction() === 'notifyContactEmails' &&
      trigger.getEventType() === ScriptApp.EventType.ON_FORM_SUBMIT;
  });

  if (!alreadyInstalled) {
    ScriptApp.newTrigger('notifyContactEmails')
      .forForm(form)
      .onFormSubmit()
      .create();
  }
}

function toText(value) {
  if (Array.isArray(value)) {
    return value.join(', ').trim();
  }
  return value === null || value === undefined ? '' : String(value).trim();
}