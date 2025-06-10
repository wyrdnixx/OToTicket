import { debugLog } from '../utils/debug.js';

export function updateStatus(message, showSpinner = false) {
  const statusDiv = document.getElementById('status');
  const spinner = showSpinner ? '<div class="spinner"></div>' : '';
  statusDiv.innerHTML = spinner + message;
  debugLog('Status: ' + message);
}
