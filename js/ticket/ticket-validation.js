export function validateTicket() {
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
