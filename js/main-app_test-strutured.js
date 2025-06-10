import { initializeAddin } from './office/office-init.js';
import { toggleDebug } from './utils/debug.js';
import { validateTicket } from './ticket/ticket-validation.js';
import { debounceSearch } from './ticket/ticket-search.js';
import { copyEmail } from './office/copy-email.js';

window.addEventListener('DOMContentLoaded', () => {
  initializeAddin();

  document.getElementById('debugMode').addEventListener('change', toggleDebug);
  document.getElementById('ticket').addEventListener('input', validateTicket);
  document.getElementById('searchInput').addEventListener('input', debounceSearch);
  document.getElementById('copyBtn').addEventListener('click', copyEmail);
});
