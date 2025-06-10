import { debugLog } from '../utils/debug.js';

export function loadEmailData() {
  const emailDataDiv = document.getElementById('emailData');

  if (!window.Office || !Office.context.mailbox.item) {
    emailDataDiv.textContent = 'Kein E-Mail-Objekt verfügbar';
    return;
  }

  const item = Office.context.mailbox.item;

  item.subject.getAsync((res) => {
    if (res.status === Office.AsyncResultStatus.Succeeded) {
      emailDataDiv.textContent = 'Betreff: ' + res.value;
    } else {
      debugLog('Fehler beim Laden des Betreffs');
    }
  });

  item.body.getAsync(Office.CoercionType.Text, (res) => {
    if (res.status === Office.AsyncResultStatus.Succeeded) {
      debugLog('E-Mail-Text geladen');
    } else {
      debugLog('Fehler beim Laden des E-Mail-Texts');
    }
  });
}
