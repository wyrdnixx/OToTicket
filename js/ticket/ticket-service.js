export async function fetchTickets(query = null) {
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
        //const response = await fetch(`http://localhost:8080/api/tickets/suggestions?q=${encodeURIComponent(searchQuery)}`);
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

//export { fetchTickets };