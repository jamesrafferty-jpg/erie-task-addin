// ===== EDIT THESE TWO =====
const HUB_URL = 'https://script.google.com/macros/s/AKfycbzstRgKA8pIdVyA2BWmnqjXVDZj8N_0g_5HUeJ6BCkUzWJq79wQujoqI_DlpkuMRWuAbg/exec'; // your web-app /exec URL
const SECRET  = '486486465v45dsf4a134Q24342RWE_))_s_d)__a4a3';                // must match CFG.SECRET in TaskHub.gs
// ==========================

let ownerName = '';

Office.onReady(() => {
  const item = Office.context.mailbox.item;
  // Prefill the task from the email subject (read mode gives a plain string).
  document.getElementById('task').value = item.subject || '';
  ownerName = (Office.context.mailbox.userProfile.displayName || '').trim();

  loadEmployees();
  document.getElementById('add').addEventListener('click', submit);
});

// Pull the active employee list from the hub to fill both dropdowns.
async function loadEmployees() {
  const by = document.getElementById('createdBy');
  const to = document.getElementById('assignedTo');
  let names = [];
  try {
    const res = await fetch(HUB_URL + '?action=employees'); // simple GET, no preflight
    const data = await res.json();
    if (data && data.ok) names = data.employees || [];
  } catch (e) { /* fall back below */ }

  if (!names.length) names = ownerName ? [ownerName] : ['James Rafferty'];

  for (const sel of [by, to]) {
    sel.innerHTML = names.map(n => '<option>' + n + '</option>').join('');
    if (ownerName && names.indexOf(ownerName) >= 0) sel.value = ownerName;
  }
  document.getElementById('add').disabled = false;
}

async function submit() {
  const btn = document.getElementById('add');
  const msg = document.getElementById('msg');
  btn.disabled = true; msg.textContent = 'Adding…'; msg.className = '';

  const payload = {
    secret: SECRET,
    task: document.getElementById('task').value,
    job: document.getElementById('job').value || 'Misc',
    createdBy: document.getElementById('createdBy').value,
    assignedTo: document.getElementById('assignedTo').value,
    due: document.getElementById('due').value || ''
  };

  try {
    // text/plain keeps this a "simple" request so Apps Script doesn't need to answer a CORS preflight.
    const res = await fetch(HUB_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload)
    });
    const data = await res.json();
    if (data.ok) {
      msg.textContent = 'Added task ' + data.taskId + ' (row ' + data.row + ')';
      msg.className = 'ok';
    } else {
      msg.textContent = 'Error: ' + (data.error || 'unknown');
      msg.className = 'err';
    }
  } catch (e) {
    msg.textContent = 'Error: ' + e.message;
    msg.className = 'err';
  }
  btn.disabled = false;
}
