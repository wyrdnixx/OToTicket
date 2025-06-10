import { debugLog } from '../utils/debug.js';

export function fallbackInitialization() {
  debugLog('Fallback initialisiert');
  addFallbackFolders();
}

export function addFallbackFolders() {
  const folderSelect = document.getElementById('folderSelect');
  const fallbackFolders = ['Support', 'Vertrieb', 'Marketing'];

  fallbackFolders.forEach(folder => {
    const option = document.createElement('option');
    option.value = folder;
    option.textContent = folder;
    folderSelect.appendChild(option);
  });
  debugLog('Fallback-Ordner hinzugefügt');
}
