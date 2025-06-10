import { debugLog } from '../utils/debug.js';

export function copyEmail() {
  if (!window.Office || !Office.context.mailbox.item) {
    debugLog('Kein Office-Kontext verfügbar');
    return;
  }

  const ticket = document.getElementById('ticket').value;
  const folder = document.getElementById('folderSelect').value;
  updateSubject(ticket);
  debugLog(`Kopie an Ordner: ${folder}`);
}

export function updateSubject(ticket) {
  const item = Office.context.mailbox.item;
  const newSubject = `[Ticket: ${ticket}] ${item.subject}`;
  item.subject.setAsync(newSubject, (res) => {
    if (res.status === Office.AsyncResultStatus.Succeeded) {
      debugLog('Betreff aktualisiert');
    } else {
      debugLog('Fehler beim Aktualisieren des Betreffs');
    }
  });
}
