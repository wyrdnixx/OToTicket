export let debugMode = false;

export function toggleDebug() {
  debugMode = document.getElementById('debugMode').checked;
  const debugOutput = document.getElementById('debugOutput');
  debugOutput.style.display = debugMode ? 'block' : 'none';
  if (!debugMode) debugOutput.innerHTML = '';
}

export function debugLog(message) {
  if (debugMode) {
    const debugOutput = document.getElementById('debugOutput');
    const timestamp = new Date().toLocaleTimeString();
    debugOutput.innerHTML += `<div><strong>${timestamp}:</strong> ${message}</div>`;
    debugOutput.scrollTop = debugOutput.scrollHeight;
  }
  console.log('OToTicket: ' + message);
}
