import { debugLog } from '../utils/debug.js';

export function notify(type, text, autoHide = false) {
  const div = document.createElement('div');
  div.className = type;
  div.innerHTML = text;
  const nm = document.getElementById('notifications');
  nm.innerHTML = '';
  nm.appendChild(div);
  debugLog(`Notification (${type}): ${text}`);

  if (autoHide) {
    setTimeout(() => div.remove(), 5000);
  }
}
