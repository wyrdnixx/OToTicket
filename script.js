let emailData = {};
let debugMode = false;
let officeReady = false;
let initTimeout;
let searchTimeout;
let selectedFolder = null;

// Target folder configuration
const TARGET_FOLDER_CONFIG = {
  mailbox: "jojo@ulewu.de", // E-Mail-Adresse des Zielpostfachs
  folderId: "AAMkADA3YmY0ZmY4LWJkMmMtNGMzZi1iZTJhLWI5NDRkNDVjMzMxNwAuAAAAAADnhTqEDCWlQqgSFCKBQnUUAQBjqVUJLaU+RJUA55He6pa2AAAAkE80AAA=", // ID des Zielordners
  changeKey: "AQAAABYAAADyLxnmgs5iRK2L2NHC4VQYAANnYp/x" // ChangeKey des Zielordners
};

/*
jojo@ulewu.de / test_OTRS
ID: AAMkADA3YmY0ZmY4LWJkMmMtNGMzZi1iZTJhLWI5NDRkNDVjMzMxNwAuAAAAAADnhTqEDCWlQqgSFCKBQnUUAQBjqVUJLaU+RJUA55He6pa2AAAAkE80AAA=
ChangeKey: AQAAABYAAADyLxnmgs5iRK2L2NHC4VQYAANnYp/x
Parent ID: root
*/




// Enhanced status updates with spinner control
function updateStatus(message, showSpinner = false) {
  const statusDiv = document.getElementById('status');
  const spinner = showSpinner ? '<div class="spinner"></div>' : '';
  statusDiv.innerHTML = spinner + message;
  debugLog('Status: ' + message);
}

function validateTicket() {
  const ticketInput = document.getElementById('ticket');
  const validation = document.getElementById('ticketValidation');
  const value = ticketInput.value;
  
  if (value === '') {
    validation.textContent = '';
    validation.className = 'input-validation';
  } else if (/^[0-9]{16}$/.test(value)) {
    validation.textContent = '✓ Gültige Ticket-Nummer';
    validation.className = 'input-validation valid';
  } else if (/^[0-9]{1,16}$/.test(value)) {
    validation.textContent = `Noch ${16 - value.length} Ziffer(n) benötigt`;
    validation.className = 'input-validation invalid';
  } else {
    validation.textContent = '✗ Nur 16 Ziffern erlaubt';
    validation.className = 'input-validation invalid';
  }
}

function debounceSearch() {
  clearTimeout(searchTimeout);
  const query = document.getElementById('searchInput').value.trim();
  
  if (query.length >= 2) {
    searchTimeout = setTimeout(() => fetchTickets(query), 300);
  } else if (query.length === 0) {
    document.getElementById('ticketList').style.display = 'none';
    document.getElementById('selectedTicket').style.display = 'none';
  }
}

function toggleDebug() {
  debugMode = document.getElementById('debugMode').checked;
  const debugOutput = document.getElementById('debugOutput');
  debugOutput.style.display = debugMode ? 'block' : 'none';
  if (!debugMode) debugOutput.innerHTML = '';
}

function debugLog(message) {
  if (debugMode) {
    const debugOutput = document.getElementById('debugOutput');
    const timestamp = new Date().toLocaleTimeString();
    debugOutput.innerHTML += `<div><strong>${timestamp}:</strong> ${message}</div>`;
    debugOutput.scrollTop = debugOutput.scrollHeight;
  }
  console.log('OToTicket: ' + message);
}

function notify(type, text, autoHide = false) {
  const div = document.createElement('div');
  div.className = type;
  div.innerHTML = text;
  const nm = document.getElementById('notifications');
  nm.innerHTML = '';
  nm.appendChild(div);
  debugLog(`Notification (${type}): ${text}`);
  
  if (autoHide) {
    setTimeout(() => {
      if (div.parentNode) {
        div.parentNode.removeChild(div);
      }
    }, 5000);
  }
}

function escapeXml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// Enhanced initialization with better error handling
function initializeAddin() {
  debugLog('Initializing addin...');
  updateStatus('Initialisiere Add-in...', true);

  initTimeout = setTimeout(function() {
    if (!officeReady) {
      debugLog('Office.js timeout - trying fallback initialization');
      updateStatus('Office.js Timeout - versuche Fallback...', true);
      fallbackInitialization();
    }
  }, 10000);

  if (typeof Office !== 'undefined') {
    debugLog('Office object found, calling Office.onReady');
    Office.onReady(function(info) {
      handleOfficeReady(info);
    });
  } else {
    debugLog('Office object not found, waiting...');
    setTimeout(function() {
      if (typeof Office !== 'undefined') {
        Office.onReady(function(info) {
          handleOfficeReady(info);
        });
      } else {
        fallbackInitialization();
      }
    }, 2000);
  }
}

function handleOfficeReady(info) {
  clearTimeout(initTimeout);
  officeReady = true;
  
  debugLog('Office.onReady called successfully');
  debugLog('Host: ' + (info ? info.host : 'unknown'));
  debugLog('Platform: ' + (info ? info.platform : 'unknown'));
  
  updateStatus('Office.js geladen - lade E-Mail Daten...', true);

  if (!Office.context || !Office.context.mailbox) {
    debugLog('No mailbox context available');
    updateStatus('❌ Fehler: Kein Mailbox-Kontext verfügbar');
    notify('error', 'Add-in nicht im korrekten E-Mail-Kontext gestartet');
    return;
  }

  document.getElementById('copyBtn').disabled = false;
  loadEmailData();
}

function fallbackInitialization() {
  debugLog('Fallback initialization');
  updateStatus('Fallback-Initialisierung...', true);
  
  if (typeof Office === 'undefined' || 
      !Office.context || 
      !Office.context.mailbox) {
    updateStatus('❌ Fehler: Office.js nicht verfügbar');
    notify('error', 'Office.js konnte nicht geladen werden. Bitte Add-in neu starten.');
    return;
  }

  officeReady = true;
  document.getElementById('copyBtn').disabled = false;
  loadEmailData();
}

function loadEmailData() {
  debugLog('Loading email data...');
  
  try {
    const item = Office.context.mailbox.item;
    if (!item) {
      debugLog('No mail item available');
      updateStatus('❌ Fehler: Keine E-Mail ausgewählt');
      return;
    }

    emailData.subject = item.subject || '';
    emailData.body = '';
    emailData.to = item.to || [];
    emailData.sender = item.sender || {};
    
    debugLog('Email subject: ' + emailData.subject);
    debugLog('Email sender: ' + emailData.sender.emailAddress);

    item.body.getAsync('html', function(res) {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        emailData.body = res.value;
        debugLog('Email body loaded successfully');
      } else {
        debugLog('Error loading email body: ' + (res.error ? res.error.message : 'Unknown error'));
      }
    });

    updateStatus('📧 E-Mail Daten geladen - bereit', true);
    
  } catch (error) {
    debugLog('Error in loadEmailData: ' + error.message);
    updateStatus('❌ Fehler beim Laden der E-Mail Daten');
    notify('error', 'Fehler beim Laden der E-Mail Daten: ' + error.message);
  }
}

function getItemDetails(itemId) {
  debugLog('Getting item details for ID: ' + itemId);
  
  const getItemSoap = `<?xml version="1.0" encoding="utf-8"?>
    <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
                   xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                   xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
      <soap:Header>
        <t:RequestServerVersion Version="Exchange2013"/>
      </soap:Header>
      <soap:Body>
        <m:GetItem>
          <m:ItemShape>
            <t:BaseShape>Default</t:BaseShape>
            <t:AdditionalProperties>
              <t:FieldURI FieldURI="item:ItemId"/>
            </t:AdditionalProperties>
          </m:ItemShape>
          <m:ItemIds>
            <t:ItemId Id="${escapeXml(itemId)}"/>
          </m:ItemIds>
        </m:GetItem>
      </soap:Body>
    </soap:Envelope>`;

  return new Promise((resolve, reject) => {
    Office.context.mailbox.makeEwsRequestAsync(getItemSoap, function(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        debugLog('GetItem response: ' + result.value);
        const responseXml = new DOMParser().parseFromString(result.value, 'text/xml');
        
        // Check for SOAP fault
        const fault = responseXml.getElementsByTagNameNS("http://schemas.xmlsoap.org/soap/envelope/", "Fault");
        if (fault.length > 0) {
          const faultString = fault[0].getElementsByTagNameNS("http://schemas.xmlsoap.org/soap/envelope/", "faultstring")[0];
          const errorMsg = faultString ? faultString.textContent : 'Unknown SOAP fault';
          debugLog('SOAP Fault: ' + errorMsg);
          reject(new Error(errorMsg));
          return;
        }

        const item = responseXml.getElementsByTagName('t:ItemId')[0];
        if (item) {
          const id = item.getAttribute('Id');
          const changeKey = item.getAttribute('ChangeKey');
          debugLog('Got item ID: ' + id + ', ChangeKey: ' + changeKey);
          resolve({ id, changeKey });
        } else {
          reject(new Error('No item found in response'));
        }
      } else {
        const errorMsg = result.error ? result.error.message : 'Unknown error';
        debugLog('GetItem error: ' + errorMsg);
        reject(new Error(errorMsg));
      }
    });
  });
}

// Modify the copyEmail function to use the target folder configuration
async function copyEmail() {
  debugLog('Copy email function called');
  
  if (!officeReady) {
    notify('error', 'Add-in noch nicht bereit');
    return;
  }

  const ticket = document.getElementById('ticket').value.trim();

  if (!/^[0-9]{16}$/.test(ticket)) {
    notify('error', '❌ Ticket muss genau 16 Ziffern haben');
    document.getElementById('ticket').focus();
    return;
  }

  // Validate email item
  const item = Office.context.mailbox.item;
  if (!item || !item.itemId) {
    notify('error', '❌ Keine E-Mail zum Kopieren ausgewählt');
    return;
  }

  // Disable button during operation
  const copyBtn = document.getElementById('copyBtn');
  copyBtn.disabled = true;
  copyBtn.textContent = '⏳ Kopiere...';

  notify('info', '📋 E‑Mail wird kopiert…');
  debugLog('Starting email copy process...');
  debugLog('Source item ID: ' + item.itemId);

  try {
    const targetFolder = `<t:FolderId Id="${escapeXml(TARGET_FOLDER_CONFIG.folderId)}" ChangeKey="${escapeXml(TARGET_FOLDER_CONFIG.changeKey)}">
      <t:Mailbox>
        <t:EmailAddress>${escapeXml(TARGET_FOLDER_CONFIG.mailbox)}</t:EmailAddress>
        <t:RoutingType>SMTP</t:RoutingType>
      </t:Mailbox>
    </t:FolderId>`;

    const copySoap = `<?xml version="1.0" encoding="utf-8"?>
      <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
                     xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                     xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
        <soap:Header>
          <t:RequestServerVersion Version="Exchange2013"/>
        </soap:Header>
        <soap:Body>
          <m:CopyItem>
            <m:ToFolderId>
              ${targetFolder}
            </m:ToFolderId>
            <m:ItemIds>
              <t:ItemId Id="${escapeXml(item.itemId)}"/>
            </m:ItemIds>
          </m:CopyItem>
        </soap:Body>
      </soap:Envelope>`;

    debugLog('Sending copy request: ' + copySoap);

    Office.context.mailbox.makeEwsRequestAsync(copySoap, function(result) {
      copyBtn.disabled = false;
      copyBtn.textContent = '📋 E‑Mail kopieren';

      if (result.status === Office.AsyncResultStatus.Succeeded) {
        debugLog('Email copy successful');
        debugLog('Copy response: ' + result.value);
        
        const responseXml = new DOMParser().parseFromString(result.value, 'text/xml');
        const copiedItem = responseXml.getElementsByTagName('t:ItemId')[0];
        
        if (copiedItem) {    
          const newItemId = copiedItem.getAttribute('Id');
          const newChangeKey = copiedItem.getAttribute('ChangeKey');
          const newSubject = `[MCB#${ticket}] ${emailData.subject}`;
          updateSubject(newItemId, newChangeKey, newSubject);
        } else {
          notify('success', '✅ E-Mail wurde erfolgreich kopiert');
        }
        
      } else {
        const errorMsg = result.error ? result.error.message : 'Unbekannter Fehler';
        notify('error', '❌ Fehler beim Kopieren: ' + errorMsg);
        debugLog('Copy error: ' + errorMsg);
      }
    });
    
  } catch (error) {
    copyBtn.disabled = false;
    copyBtn.textContent = '📋 E‑Mail kopieren';
    debugLog('Error in copyEmail: ' + error.message);
    notify('error', '❌ Fehler beim Kopieren: ' + error.message);
  }
}

function updateSubject(itemId, changeKey, newSubject) {
  debugLog('Updating subject for item: ' + itemId);
  debugLog('New subject: ' + newSubject);

  const updateSoap = `<?xml version="1.0" encoding="utf-8"?>
    <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
                   xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                   xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
      <soap:Header>
        <t:RequestServerVersion Version="Exchange2013" />
      </soap:Header>
      <soap:Body>
        <m:UpdateItem MessageDisposition="SaveOnly" ConflictResolution="AutoResolve">
          <m:ItemChanges>
            <t:ItemChange>
              <t:ItemId Id="${escapeXml(itemId)}" ChangeKey="${escapeXml(changeKey)}"/>
              <t:Updates>
                <t:SetItemField>
                  <t:FieldURI FieldURI="item:Subject" />
                  <t:Message>
                    <t:Subject>${escapeXml(newSubject)}</t:Subject>
                  </t:Message>
                </t:SetItemField>
              </t:Updates>
            </t:ItemChange>
          </m:ItemChanges>
        </m:UpdateItem>
      </soap:Body>
    </soap:Envelope>`;

  debugLog('Sending update subject request: ' + updateSoap);

  Office.context.mailbox.makeEwsRequestAsync(updateSoap, result => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      debugLog('Subject update successful');
      debugLog('Update response: ' + result.value);
      notify('success', '✅ E-Mail wurde erfolgreich kopiert und Betreff angepasst.', true);
    } else {
      const errorMsg = result.error ? result.error.message : 'Unbekannter Fehler';
      debugLog('Subject update error: ' + errorMsg);
      debugLog('Error response: ' + result.value);
      notify('error', '❌ Fehler beim Aktualisieren des Betreffs: ' + errorMsg);
    }
  });
}

function closeDialog() {
  debugLog('Close dialog called');
  
  try {
    if (typeof Office !== 'undefined' && 
        Office.context && 
        Office.context.ui && 
        Office.context.ui.closeContainer) {
      Office.context.ui.closeContainer();
    } else if (window.close) {
      window.close();
    } else {
      document.body.style.display = 'none';
    }
  } catch (error) {
    debugLog('Error closing dialog: ' + error.message);
    if (window.close) window.close();
  }
}

// Enhanced ticket search with better UX
async function fetchTickets(query = null) {
  const button = document.getElementById("ticketSearchBtn");
  const ticketList = document.getElementById("ticketList");
  const selectedDiv = document.getElementById("selectedTicket");
  const searchInput = document.getElementById("searchInput");

  const searchQuery = query || searchInput.value.trim();
  
  if (!searchQuery) {
    notify('warning', '⚠️ Bitte Suchbegriff eingeben');
    return;
  }

  button.disabled = true;
  button.textContent = "🔄 Suche...";

  try {          
    console.log("Searching for " + emailData.sender.emailAddress); 
    const response = await fetch(`http://localhost:8080/api/tickets/suggestions?q=${encodeURIComponent(searchQuery)}&mail=${encodeURIComponent(emailData.sender.emailAddress)}`);
    
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }
    
    const tickets = await response.json();

    ticketList.innerHTML = "";
    selectedDiv.style.display = "none";

    if (tickets.length === 0) {
      ticketList.innerHTML = '<div class="ticket-item" style="text-align: center; color: #6c757d;">Keine Tickets gefunden</div>';
      ticketList.style.display = "block";
      return;
    }

    tickets.forEach(ticket => {
      const div = document.createElement("div");
      div.className = "ticket-item";
      div.innerHTML = `
        <div style="font-weight: 600; color: #0078d4;">#${ticket.tn}</div>
        <div style="margin: 4px 0;">${ticket.title}</div>
        <div style="font-size: 12px; color: #6c757d;">${ticket.name}</div>
      `;
      div.dataset.tn = ticket.tn;

      div.addEventListener("click", () => {
        // Remove previous selection
        document.querySelectorAll(".ticket-item").forEach(el => el.classList.remove("selected"));
        div.classList.add("selected");
        
        // Update ticket input
        document.getElementById("ticket").value = ticket.tn;
        validateTicket();
        
        // Show selection info
        document.getElementById("selectedTicketInfo").textContent = `#${ticket.tn} - ${ticket.title}`;
        selectedDiv.style.display = "block";
        
        notify('success', `✅ Ticket #${ticket.tn} ausgewählt`, true);
      });

      ticketList.appendChild(div);
    });

    ticketList.style.display = "block";
    notify('info', `🔍 ${tickets.length} Ticket(s) gefunden`, true);
    
  } catch (err) {
    ticketList.innerHTML = '<div class="ticket-item" style="color: #dc3545; text-align: center;">❌ Fehler beim Laden der Tickets</div>';
    ticketList.style.display = "block";
    notify('error', '❌ Fehler beim Laden der Tickets: ' + err.message);
    console.error('Ticket search error:', err);
  } finally {
    button.disabled = false;
    button.textContent = "🔍 Suchen";
  }
}

// Initialize when DOM is ready
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initializeAddin);
} else {
  initializeAddin();
}

// Additional fallback after page load
window.addEventListener('load', function() {
  if (!officeReady) {
    debugLog('Window load event - Office still not ready, trying again...');
    setTimeout(function() {
      if (!officeReady) {
        fallbackInitialization();
      }
    }, 1000);
  }
});

// Keyboard shortcuts
document.addEventListener('keydown', function(e) {
  // Ctrl/Cmd + Enter to copy email
  if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
    e.preventDefault();
    if (!document.getElementById('copyBtn').disabled) {
      copyEmail();
    }
  }
  
  // Escape to close
  if (e.key === 'Escape') {
    closeDialog();
  }
  
  // Enter in search to search
  if (e.key === 'Enter' && e.target.id === 'searchInput') {
    e.preventDefault();
    fetchTickets();
  }
});

// Auto-focus ticket input when ready
function focusTicketInput() {
  if (officeReady) {
    document.getElementById('ticket').focus();
  }
}

// Call focus after initialization
setTimeout(focusTicketInput, 1000);

function openConfigDialog() {
  const modal = document.getElementById('configModal');
  modal.style.display = 'block';
  document.getElementById('configEmail').value = TARGET_FOLDER_CONFIG.mailbox;
}

function closeConfigDialog() {
  const modal = document.getElementById('configModal');
  modal.style.display = 'none';
}

function fetchFolderInfo() {
  const email = document.getElementById('configEmail').value.trim();
  if (!email) {
    notify('error', '❌ Bitte E-Mail-Adresse eingeben');
    return;
  }

  const folderList = document.getElementById('folderList');
  folderList.innerHTML = '<div class="loading">Lade Ordnerinformationen...</div>';
  debugLog('Starte Ordnerabfrage für: ' + email);

  const findFolderSoap =  `<?xml version="1.0" encoding="utf-8"?>
  <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
                 xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                 xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
    <soap:Header>
      <t:RequestServerVersion Version="Exchange2013" />
      <t:MailboxCulture>de-DE</t:MailboxCulture>
      <t:TimeZoneContext>
        <t:TimeZoneDefinition Id="W. Europe Standard Time" />
      </t:TimeZoneContext>
    </soap:Header>
    <soap:Body>
      <m:FindFolder Traversal="Deep">
        <m:FolderShape>
          <t:BaseShape>Default</t:BaseShape>
        </m:FolderShape>
        <m:ParentFolderIds>
          <t:DistinguishedFolderId Id="msgfolderroot">
            <t:Mailbox>
              <t:EmailAddress>${escapeXml(email)}</t:EmailAddress>
              <t:RoutingType>SMTP</t:RoutingType>
            </t:Mailbox>
          </t:DistinguishedFolderId>
        </m:ParentFolderIds>
      </m:FindFolder>
    </soap:Body>
  </soap:Envelope>`;

  debugLog('SOAP Request: ' + findFolderSoap);

  Office.context.mailbox.makeEwsRequestAsync(findFolderSoap, function(result) {
    debugLog('EWS Response Status: ' + result.status);
    
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      debugLog('EWS Response: ' + result.value);
      
      const responseXml = new DOMParser().parseFromString(result.value, 'text/xml');
      
      // Check for SOAP fault
      const fault = responseXml.getElementsByTagNameNS("http://schemas.xmlsoap.org/soap/envelope/", "Fault");
      if (fault.length > 0) {
        const faultString = fault[0].getElementsByTagNameNS("http://schemas.xmlsoap.org/soap/envelope/", "faultstring")[0];
        const errorMsg = faultString ? faultString.textContent : 'Unknown SOAP fault';
        debugLog('SOAP Fault: ' + errorMsg);
        folderList.innerHTML = `<div class="error">❌ SOAP Fehler: ${errorMsg}</div>`;
        notify('error', '❌ SOAP Fehler: ' + errorMsg);
        return;
      }

      const folders = responseXml.getElementsByTagName('t:Folder');
      debugLog('Gefundene Ordner: ' + folders.length);
      
      let html = '<div class="folder-list">';
      for (let folder of folders) {
        const folderId = folder.getElementsByTagName('t:FolderId')[0];
        const displayName = folder.getElementsByTagName('t:DisplayName')[0];
        const parentFolderId = folder.getElementsByTagName('t:ParentFolderId')[0];
        
        if (folderId && displayName) {
          const id = folderId.getAttribute('Id');
          const changeKey = folderId.getAttribute('ChangeKey');
          const name = displayName.textContent;
          const parentId = parentFolderId ? parentFolderId.getAttribute('Id') : 'root';
          
          debugLog(`Ordner gefunden: ${name} (ID: ${id}, ChangeKey: ${changeKey}, Parent: ${parentId})`);
          
          html += `
            <div class="folder-item">
              <div class="folder-name">📁 ${name}</div>
              <div class="folder-details">
                <div><strong>ID:</strong> <code>${id}</code></div>
                <div><strong>ChangeKey:</strong> <code>${changeKey}</code></div>
                <div><strong>Parent ID:</strong> <code>${parentId}</code></div>
              </div>
            </div>
          `;
        } else {
          debugLog('Ordner ohne ID oder DisplayName gefunden: ' + folder.outerHTML);
        }
      }
      html += '</div>';
      
      if (folders.length === 0) {
        html = '<div class="error">Keine Ordner gefunden</div>';
        debugLog('Keine Ordner in der Antwort gefunden');
      }
      
      folderList.innerHTML = html;
    } else {
      const errorMsg = result.error ? result.error.message : 'Unbekannter Fehler';
      debugLog('EWS Fehler: ' + errorMsg);
      debugLog('Fehler Details: ' + JSON.stringify(result.error));
      folderList.innerHTML = `<div class="error">❌ Fehler beim Abrufen der Ordner: ${errorMsg}</div>`;
      notify('error', '❌ Fehler beim Abrufen der Ordner: ' + errorMsg);
    }
  });
}

// Close modal when clicking outside
window.onclick = function(event) {
  const modal = document.getElementById('configModal');
  if (event.target === modal) {
    closeConfigDialog();
  }
}