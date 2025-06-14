let emailData = {};
let debugMode = false;
let officeReady = false;
let initTimeout;
let searchTimeout;
let selectedFolder = null;

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
  loadFolders();
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
  loadFolders();
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

    updateStatus('📧 E-Mail Daten geladen - lade Ordner...', true);
    
  } catch (error) {
    debugLog('Error in loadEmailData: ' + error.message);
    updateStatus('❌ Fehler beim Laden der E-Mail Daten');
    notify('error', 'Fehler beim Laden der E-Mail Daten: ' + error.message);
  }
}

function loadFolders() {
  debugLog('Starting to load folders...');
  
  try {
    const currentMailbox = Office.context.mailbox.userProfile.emailAddress;
    debugLog('Current mailbox: ' + currentMailbox);

    // First, get the current user's folders
    const currentMailboxRequest = `<?xml version="1.0" encoding="utf-8"?>
      <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
                     xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                     xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
        <soap:Header>
          <t:RequestServerVersion Version="Exchange2013" />
        </soap:Header>
        <soap:Body>
          <m:FindFolder Traversal="Deep">
            <m:FolderShape>
              <t:BaseShape>Default</t:BaseShape>
            </m:FolderShape>
            <m:ParentFolderIds>
              <t:DistinguishedFolderId Id="msgfolderroot" />
            </m:ParentFolderIds>
          </m:FindFolder>
        </soap:Body>
      </soap:Envelope>`;

    // Get other mailboxes using GetMailTips
    const mailTipsRequest = `<?xml version="1.0" encoding="utf-8"?>
      <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
                     xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                     xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
        <soap:Header>
          <t:RequestServerVersion Version="Exchange2013" />
        </soap:Header>
        <soap:Body>
          <m:GetMailTips>
            <m:RequestedMailTips>MailboxFullStatus</m:RequestedMailTips>
            <m:Recipients>
              <t:Mailbox>
                <t:EmailAddress>${escapeXml(currentMailbox)}</t:EmailAddress>
              </t:Mailbox>
            </m:Recipients>
          </m:GetMailTips>
        </soap:Body>
      </soap:Envelope>`;

    debugLog('Sending mail tips request...');
    
    Office.context.mailbox.makeEwsRequestAsync(mailTipsRequest, function(mailTipsResult) {
      if (mailTipsResult.status === Office.AsyncResultStatus.Succeeded) {
        debugLog('Mail tips request successful');
        
        // Get the list of mailboxes from the user's Outlook profile
        const otherMailboxes = [];
        try {
          // Try to get other mailboxes from the current item's context
          const item = Office.context.mailbox.item;
          if (item && item.to) {
            item.to.forEach(recipient => {
              if (recipient.emailAddress && recipient.emailAddress !== currentMailbox) {
                otherMailboxes.push(recipient.emailAddress);
              }
            });
          }
        } catch (error) {
          debugLog('Error getting other mailboxes: ' + error.message);
        }

        // Add some common mailboxes that might be accessible
        const commonMailboxes = [
          'otrs@ulewu.de',  // Replace with actual common mailboxes
          'ulewu@example.com'
        ];
        
        otherMailboxes.push(...commonMailboxes);

        debugLog('Found other mailboxes: ' + otherMailboxes.join(', '));

        // Create folder requests for each mailbox
        const folderRequests = [
          { mailbox: currentMailbox, request: currentMailboxRequest }
        ];

        otherMailboxes.forEach(mailbox => {
          const request = `<?xml version="1.0" encoding="utf-8"?>
            <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
                           xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                           xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
              <soap:Header>
                <t:RequestServerVersion Version="Exchange2013" />
              </soap:Header>
              <soap:Body>
                <m:FindFolder Traversal="Deep">
                  <m:FolderShape>
                    <t:BaseShape>Default</t:BaseShape>
                  </m:FolderShape>
                  <m:ParentFolderIds>
                    <t:DistinguishedFolderId Id="msgfolderroot">
                      <t:Mailbox>
                        <t:EmailAddress>${escapeXml(mailbox)}</t:EmailAddress>
                      </t:Mailbox>
                    </t:DistinguishedFolderId>
                  </m:ParentFolderIds>
                </m:FindFolder>
              </soap:Body>
            </soap:Envelope>`;
          folderRequests.push({ mailbox, request });
        });

        // Process each mailbox's folders
        let processedCount = 0;
        const allFolders = new Map();

        folderRequests.forEach(({ mailbox, request }) => {
          Office.context.mailbox.makeEwsRequestAsync(request, function(result) {
            processedCount++;
            
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              const xmlDoc = new DOMParser().parseFromString(result.value, "text/xml");
              const folders = xmlDoc.getElementsByTagNameNS("http://schemas.microsoft.com/exchange/services/2006/types", "Folder");
              
              debugLog(`Found ${folders.length} folders in ${mailbox}`);
              
              const mailboxFolders = [];
              for (let i = 0; i < folders.length; i++) {
                const folder = folders[i];
                const displayNameElem = folder.getElementsByTagNameNS("http://schemas.microsoft.com/exchange/services/2006/types", "DisplayName")[0];
                const folderIdElem = folder.getElementsByTagNameNS("http://schemas.microsoft.com/exchange/services/2006/types", "FolderId")[0];

                if (!displayNameElem || !folderIdElem) continue;

                const displayName = displayNameElem.textContent;
                const folderId = folderIdElem.getAttribute("Id");
                const changeKey = folderIdElem.getAttribute("ChangeKey") || "";

                if (displayName.startsWith("~") || displayName === "Conversation Action Settings") {
                  continue;
                }

                mailboxFolders.push({
                  displayName,
                  value: folderId + ";" + changeKey + ";" + mailbox,
                  mailbox
                });
              }

              if (mailboxFolders.length > 0) {
                allFolders.set(mailbox, mailboxFolders);
              }
            }

            // When all requests are processed, update the UI
            if (processedCount === folderRequests.length) {
              updateFolderSelect(allFolders, currentMailbox);
            }
          });
        });
      } else {
        debugLog('Mail tips request failed: ' + (mailTipsResult.error ? mailTipsResult.error.message : 'Unknown error'));
        // Fallback to just loading current mailbox folders
        loadCurrentMailboxFolders();
      }
    });
    
  } catch (error) {
    debugLog('Error in loadFolders: ' + error.message);
    updateStatus('❌ Fehler beim Laden der Ordner');
    notify('error', 'Fehler beim Laden der Ordner: ' + error.message);
    loadCurrentMailboxFolders();
  }
}

function loadCurrentMailboxFolders() {
  const currentMailbox = Office.context.mailbox.userProfile.emailAddress;
  const request = `<?xml version="1.0" encoding="utf-8"?>
    <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
                   xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
                   xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
      <soap:Header>
        <t:RequestServerVersion Version="Exchange2013" />
      </soap:Header>
      <soap:Body>
        <m:FindFolder Traversal="Deep">
          <m:FolderShape>
            <t:BaseShape>Default</t:BaseShape>
          </m:FolderShape>
          <m:ParentFolderIds>
            <t:DistinguishedFolderId Id="msgfolderroot" />
          </m:ParentFolderIds>
        </m:FindFolder>
      </soap:Body>
    </soap:Envelope>`;

  Office.context.mailbox.makeEwsRequestAsync(request, function(result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const allFolders = new Map();
      const xmlDoc = new DOMParser().parseFromString(result.value, "text/xml");
      const folders = xmlDoc.getElementsByTagNameNS("http://schemas.microsoft.com/exchange/services/2006/types", "Folder");
      
      const mailboxFolders = [];
      for (let i = 0; i < folders.length; i++) {
        const folder = folders[i];
        const displayNameElem = folder.getElementsByTagNameNS("http://schemas.microsoft.com/exchange/services/2006/types", "DisplayName")[0];
        const folderIdElem = folder.getElementsByTagNameNS("http://schemas.microsoft.com/exchange/services/2006/types", "FolderId")[0];

        if (!displayNameElem || !folderIdElem) continue;

        const displayName = displayNameElem.textContent;
        const folderId = folderIdElem.getAttribute("Id");
        const changeKey = folderIdElem.getAttribute("ChangeKey") || "";

        if (displayName.startsWith("~") || displayName === "Conversation Action Settings") {
          continue;
        }

        mailboxFolders.push({
          displayName,
          value: folderId + ";" + changeKey + ";" + currentMailbox,
          mailbox: currentMailbox
        });
      }

      allFolders.set(currentMailbox, mailboxFolders);
      updateFolderSelect(allFolders, currentMailbox);
    } else {
      addFallbackFolders();
    }
  });
}

function updateFolderSelect(allFolders, currentMailbox) {
  const folderSelect = document.getElementById("folderSelect");
  folderSelect.innerHTML = '<option value="">📁 Bitte wählen</option>';

  // Sort mailboxes (current mailbox first)
  const sortedMailboxes = Array.from(allFolders.keys()).sort((a, b) => {
    if (a === currentMailbox) return -1;
    if (b === currentMailbox) return 1;
    return a.localeCompare(b);
  });

  let totalFolders = 0;

  // Add folders grouped by mailbox
  sortedMailboxes.forEach(mailbox => {
    const mailboxFolders = allFolders.get(mailbox);
    if (!mailboxFolders || mailboxFolders.length === 0) return;

    const isCurrentMailbox = mailbox === currentMailbox;
    const mailboxName = isCurrentMailbox ? "Meine Mailbox" : mailbox.split('@')[0];
    
    // Add mailbox group header
    const optgroup = document.createElement("optgroup");
    optgroup.label = `📧 ${mailboxName}`;
    folderSelect.appendChild(optgroup);

    // Sort folders alphabetically
    mailboxFolders.sort((a, b) => a.displayName.localeCompare(b.displayName));

    // Add folders for this mailbox
    mailboxFolders.forEach(folder => {
      const option = document.createElement("option");
      option.value = folder.value;
      option.textContent = folder.displayName;
      optgroup.appendChild(option);
      totalFolders++;
    });
  });

  const savedFolder = Office.context.roamingSettings.get("lastSelectedFolder");
  if (savedFolder) {
    folderSelect.value = savedFolder;
  }

  updateStatus(`✅ ${totalFolders} Ordner geladen - bereit`);
  notify('success', `${totalFolders} Ordner erfolgreich geladen`, true);
}

function addFallbackFolders() {
  debugLog('Adding fallback folders...');
  const folderSelect = document.getElementById("folderSelect");
  folderSelect.innerHTML = '<option value="">📁 Bitte wählen</option>';
  
  const fallbackFolders = [
    { name: "📥 Posteingang", id: "inbox;" },
    { name: "📤 Gesendete Objekte", id: "sentitems;" },
    { name: "📝 Entwürfe", id: "drafts;" }
  ];
  
  fallbackFolders.forEach(folder => {
    const option = document.createElement("option");
    option.value = folder.id;
    option.textContent = folder.name;
    folderSelect.appendChild(option);
  });
  
  updateStatus('⚠️ Fallback-Ordner geladen - bereit');
  notify('warning', 'Standard-Ordner geladen (Ordner-Laden fehlgeschlagen)', true);
}

function showFolderDialog() {
  document.getElementById('folderDialog').style.display = 'flex';
  document.getElementById('mailboxInput').focus();
}

function hideFolderDialog() {
  document.getElementById('folderDialog').style.display = 'none';
}

function clearFolderSelection() {
  selectedFolder = null;
  document.getElementById('selectedFolderName').textContent = 'Kein Ordner ausgewählt';
  document.getElementById('copyBtn').disabled = true;
}

function loadMailboxFolders() {
  const mailboxInput = document.getElementById('mailboxInput');
  const mailbox = mailboxInput.value.trim();
  
  if (!mailbox) {
    notify('error', '❌ Bitte Mailbox-Adresse eingeben');
    return;
  }

  const folderList = document.getElementById('folderList');
  folderList.innerHTML = '<div class="folder-item" style="text-align: center; color: #6c757d;">Lade Ordner...</div>';

  debugLog('Loading folders for mailbox: ' + mailbox);

  // Directly try to access the target mailbox folders
  /* const request = `<?xml version="1.0" encoding="utf-8"?>
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
            <t:DistinguishedFolderId Id="Posteingang">
              <t:Mailbox>
                <t:EmailAddress>${escapeXml(mailbox)}</t:EmailAddress>
                <t:RoutingType>SMTP</t:RoutingType>
              </t:Mailbox>
            </t:DistinguishedFolderId>
          </m:ParentFolderIds>
        </m:FindFolder>
      </soap:Body>
    </soap:Envelope>`; */

  const request = `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
  <soap:Body>
    <FindFolder Traversal="Deep" xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
      <FolderShape>
        <t:BaseShape>Default</t:BaseShape>
      </FolderShape>
      <ParentFolderIds>
        <t:DistinguishedFolderId Id="msgfolderroot"/>
          <t:Mailbox>
                <t:EmailAddress>${escapeXml(mailbox)}</t:EmailAddress>
                <t:RoutingType>SMTP</t:RoutingType>
              </t:Mailbox>
      </ParentFolderIds>
    </FindFolder>
  </soap:Body>
</soap:Envelope>`;

  debugLog('Sending FindFolder request: ' + request);

  Office.context.mailbox.makeEwsRequestAsync(request, function(result) {
    debugLog('FindFolder Response status: ' + result.status);
    
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      debugLog('FindFolder Response: ' + result.value);
      
      try {
        const xmlDoc = new DOMParser().parseFromString(result.value, "text/xml");
        
        // Check for SOAP fault
        const fault = xmlDoc.getElementsByTagNameNS("http://schemas.xmlsoap.org/soap/envelope/", "Fault");
        if (fault.length > 0) {
          const faultString = fault[0].getElementsByTagNameNS("http://schemas.xmlsoap.org/soap/envelope/", "faultstring")[0];
          const errorMsg = faultString ? faultString.textContent : 'Unknown SOAP fault';
          debugLog('SOAP Fault: ' + errorMsg);
          folderList.innerHTML = '<div class="folder-item" style="text-align: center; color: #dc3545;">❌ Fehler: ' + errorMsg + '</div>';
          notify('error', '❌ Fehler beim Laden der Ordner: ' + errorMsg);
          return;
        }

        const folders = xmlDoc.getElementsByTagNameNS("http://schemas.microsoft.com/exchange/services/2006/types", "Folder");
        debugLog('Found ' + folders.length + ' folders');
        
        folderList.innerHTML = '';
        
        if (folders.length === 0) {
          folderList.innerHTML = '<div class="folder-item" style="text-align: center; color: #6c757d;">Keine Ordner gefunden</div>';
          return;
        }

        for (let i = 0; i < folders.length; i++) {
          const folder = folders[i];
          const displayNameElem = folder.getElementsByTagNameNS("http://schemas.microsoft.com/exchange/services/2006/types", "DisplayName")[0];
          const folderIdElem = folder.getElementsByTagNameNS("http://schemas.microsoft.com/exchange/services/2006/types", "FolderId")[0];

          if (!displayNameElem || !folderIdElem) {
            debugLog('Skipping folder - missing displayName or folderId');
            continue;
          }

          const displayName = displayNameElem.textContent;
          const folderId = folderIdElem.getAttribute("Id");
          const changeKey = folderIdElem.getAttribute("ChangeKey") || "";

          debugLog('Processing folder: ' + displayName + ' (ID: ' + folderId + ')');

          if (displayName.startsWith("~") || displayName === "Conversation Action Settings") {
            debugLog('Skipping system folder: ' + displayName);
            continue;
          }

          const div = document.createElement("div");
          div.className = "folder-item";
          div.textContent = displayName;
          div.onclick = function() {
            // Remove previous selection
            document.querySelectorAll(".folder-item").forEach(el => el.classList.remove("selected"));
            div.classList.add("selected");
            
            // Store selected folder
            selectedFolder = {
              displayName,
              value: folderId + ";" + changeKey + ";" + mailbox,
              mailbox
            };
            
            // Update UI
            document.getElementById('selectedFolderName').textContent = `${displayName} (${mailbox})`;
            document.getElementById('copyBtn').disabled = false;
            
            // Close dialog
            hideFolderDialog();
            
            notify('success', `✅ Ordner "${displayName}" ausgewählt`, true);
          };

          folderList.appendChild(div);
        }
      } catch (error) {
        debugLog('Error parsing XML response: ' + error.message);
        folderList.innerHTML = '<div class="folder-item" style="text-align: center; color: #dc3545;">❌ Fehler beim Verarbeiten der Antwort</div>';
        notify('error', '❌ Fehler beim Verarbeiten der Antwort: ' + error.message);
      }
    } else {
      const errorMsg = result.error ? result.error.message : 'Unbekannter Fehler';
      debugLog('FindFolder Request failed: ' + errorMsg);
      folderList.innerHTML = '<div class="folder-item" style="text-align: center; color: #dc3545;">❌ Fehler beim Laden der Ordner</div>';
      notify('error', '❌ Fehler beim Laden der Ordner: ' + errorMsg);
    }
  });
}

function copyEmail() {
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
  if (!selectedFolder) {
    notify('error', '❌ Bitte Zielordner auswählen');
    return;
  }

  // Disable button during operation
  const copyBtn = document.getElementById('copyBtn');
  copyBtn.disabled = true;
  copyBtn.textContent = '⏳ Kopiere...';

  notify('info', '📋 E‑Mail wird kopiert…');
  debugLog('Starting email copy process...');

  try {
    const [folderId, changeKey, mailbox] = selectedFolder.value.split(";");
    const currentMailbox = Office.context.mailbox.userProfile.emailAddress;
    
    let targetFolder;
    if (mailbox === currentMailbox) {
      if (folderId === 'inbox') {
        targetFolder = '<t:DistinguishedFolderId Id="inbox" />';
      } else if (folderId === 'sentitems') {
        targetFolder = '<t:DistinguishedFolderId Id="sentitems" />';
      } else if (folderId === 'drafts') {
        targetFolder = '<t:DistinguishedFolderId Id="drafts" />';
      } else {
        targetFolder = `<t:FolderId Id="${escapeXml(folderId)}" ChangeKey="${escapeXml(changeKey)}"/>`;
      }
    } else {
      // For other mailboxes, we need to specify the mailbox
      targetFolder = `<t:FolderId Id="${escapeXml(folderId)}" ChangeKey="${escapeXml(changeKey)}">
        <t:Mailbox>
          <t:EmailAddress>${escapeXml(mailbox)}</t:EmailAddress>
        </t:Mailbox>
      </t:FolderId>`;
    }

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
              <t:ItemId Id="${escapeXml(Office.context.mailbox.item.itemId)}"/>
            </m:ItemIds>
          </m:CopyItem>
        </soap:Body>
      </soap:Envelope>`;

    Office.context.mailbox.makeEwsRequestAsync(copySoap, function(result) {
      copyBtn.disabled = false;
      copyBtn.textContent = '📋 E‑Mail kopieren';

      if (result.status === Office.AsyncResultStatus.Succeeded) {
        debugLog('Email copy successful');
        
        const responseXml = new DOMParser().parseFromString(result.value, 'text/xml');
        const copiedItem = responseXml.getElementsByTagName('t:ItemId')[0];
        
        if (copiedItem) {    
          const newItemId = copiedItem.getAttribute('Id');
          const newChangeKey = copiedItem.getAttribute('ChangeKey');
          const newSubject = `[MCB#${ticket}] ${emailData.subject}`;
          updateSubject(newItemId, newChangeKey, newSubject, mailbox);
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
  const updateSoap = `
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

  Office.context.mailbox.makeEwsRequestAsync(updateSoap, result => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      notify('success', '✅ E-Mail wurde erfolgreich kopiert und Betreff angepasst.', true);
    } else {
      notify('error', '❌ Fehler beim Aktualisieren des Betreffs: ' + (result.error ? result.error.message : 'Unbekannt'));
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