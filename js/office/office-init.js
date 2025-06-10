import { fallbackInitialization } from './folder-loader.js';
import { debugLog } from '../utils/debug.js';

export function initializeAddin() {
  if (window.Office && Office.context && Office.context.mailbox) {
    Office.onReady().then(handleOfficeReady).catch(() => {
      debugLog('Office initialization failed, falling back.');
      fallbackInitialization();
    });
  } else {
    debugLog('Office not detected, using fallback.');
    fallbackInitialization();
  }
}

function handleOfficeReady(info) {
  if (info.host === Office.HostType.Outlook) {
    debugLog('Office initialized');
    if (Office.context.mailbox.item) {
      debugLog('Item: ' + Office.context.mailbox.item.itemId);
    }
  }
}
