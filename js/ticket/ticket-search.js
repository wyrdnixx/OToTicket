import { fetchTickets } from './ticket-service.js';

let searchTimeout;

export function debounceSearch() {
  clearTimeout(searchTimeout);
  const query = document.getElementById('searchInput').value.trim();

  if (query.length >= 2) {
    searchTimeout = setTimeout(() => fetchTickets(query), 300);
  } else if (query.length === 0) {
    document.getElementById('ticketList').style.display = 'none';
    document.getElementById('selectedTicket').style.display = 'none';
  }
}

window.fetchTickets = fetchTickets;