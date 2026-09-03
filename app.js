let entries = JSON.parse(localStorage.getItem('pacsData') || '[]');
let serial = entries.length > 0
  ? Math.max(...entries.map(e => e.sr)) + 1
  : 1;


// =========================
// SERIAL NUMBER
// =========================

function updateSrField() {
  document.getElementById('sr').value = serial;
}


// =========================
// LOCAL STORAGE
// =========================

function saveToStorage() {
  localStorage.setItem('pacsData', JSON.stringify(entries));
}


// =========================
// TOAST MESSAGE
// =========================

function showToast(msg, color = '#1e7e34') {

  const t = document.getElementById('toast');

  t.textContent = msg;
  t.style.background = color;
  t.style.display = 'block';

  setTimeout(() => {
    t.style.display = 'none';
  }, 2500);
}


// =========================
// ADD ENTRY
// =========================

function addEntry() {

  const naam = document.getElementById('naam').value.trim();
  const pita = document.getElementById('pita').value.trim();

  if (!naam || !pita) {

    showToast(
      '⚠️ सदस्य का नाम और पिता/पति का नाम जरूरी है!',
      '#c0392b'
    );

    return;
  }

  const entry = {

    sr: serial,

    naam: naam,

    pita: pita,

    pata: document.getElementById('pata').value.trim(),

    sadasy: document.getElementById('sadasy').value,

    aadhar: document.getElementById('aadhar').value.trim(),

    dob: document.getElementById('dob').value,

    aayu: document.getElementById('aayu').value,

    ling: document.getElementById('ling').value,

    vivah: document.getElementById('vivah').value,

    dharm: document.getElementById('dharm').value,

    jaati: document.getElementById('jaati').value.trim(),

    dccb: document.getElementById('dccb').value.trim(),

    email: document.getElementById('email').value.trim(),

    mobile: document.getElementById('mobile').value.trim(),

    nominee: document.getElementById('nominee').value.trim()
  };


  entries.push(entry);

  serial++;

  saveToStorage();

  renderTable();

  clearForm();

  updateSrField();

  showToast('✅ एंट्री सफलतापूर्वक जोड़ी गई!');
}


// =========================
// DELETE ENTRY
// =========================

function deleteEntry(sr) {

  entries = entries.filter(e => e.sr !== sr);

  saveToStorage();

  renderTable();

  showToast(
    '🗑️ एंट्री हटाई गई',
    '#e67e22'
  );
}


// =========================
// CLEAR FORM
// =========================

function clearForm() {

  [
    'naam',
    'pita',
    'pata',
    'aadhar',
    'dob',
    'aayu',
    'jaati',
    'dccb',
    'email',
    'mobile',
    'nominee'
  ].forEach(id => {

    document.getElementById(id).value = '';

  });


  [
    'sadasy',
    'ling',
    'vivah',
    'dharm'
  ].forEach(id => {

    document.getElementById(id).value = '';

  });
}


// =========================
// CLEAR ALL
// =========================

function clearAll() {

  if (!confirm('क्या आप सभी एंट्री हटाना चाहते हैं?')) {
    return;
  }

  entries = [];

  serial = 1;

  saveToStorage();

  renderTable();

  updateSrField();

  showToast(
    '🗑️ सब हटा दिया गया',
    '#c0392b'
  );
}


// =========================
// RENDER TABLE
// =========================

function renderTable() {

  const tbody = document.getElementById('tableBody');

  document.getElementById('count').textContent =
    entries.length;


  if (entries.length === 0) {

    tbody.innerHTML = `
      <tr>
        <td colspan="17"
            style="padding:20px;color:#999;">
          कोई एंट्री नहीं है
        </td>
      </tr>
    `;

    return;
  }


  tbody.innerHTML = entries.map(e => `

    <tr>

      <td>${e.sr}</td>

      <td>${e.naam}</td>

      <td>${e.pita}</td>

      <td>${e.pata}</td>

      <td>${e.sadasy}</td>

      <td>${e.aadhar}</td>

      <td>${e.dob}</td>

      <td>${e.aayu}</td>

      <td>${e.ling}</td>

      <td>${e.vivah}</td>

      <td>${e.dharm}</td>

      <td>${e.jaati}</td>

      <td>${e.dccb}</td>

      <td>${e.email}</td>

      <td>${e.mobile}</td>

      <td>${e.nominee}</td>

      <td>
        <button
          class="btn-del"
          onclick="deleteEntry(${e.sr})">
          हटाएं
        </button>
      </td>

    </tr>

  `).join('');
}


// =========================
// EXPORT TO EXCEL
// =========================

function exportExcel() {

  if (entries.length === 0) {

    showToast(
      '⚠️ पहले कुछ एंट्री जोड़ें!',
      '#c0392b'
    );

    return;
  }


  const headers = [

    'क्र0 सं0',
    'सदस्य का नाम',
    'पिता/पति का नाम',
    'पता',
    'सदस्य/सहसदस्य',
    'आधार कार्ड संख्या',
    'जन्म तिथि',
    'आयु',
    'लिंग',
    'वैवाहिक स्थिति',
    'धर्म',
    'जाति',
    'DCCB बचत खाता संख्या',
    'ईमेल आईडी',
    'मोबाईल नम्बर',
    'नॉमिनी'

  ];


  const rows = entries.map(e => [

    e.sr,
    e.naam,
    e.pita,
    e.pata,
    e.sadasy,
    e.aadhar,
    e.dob,
    e.aayu,
    e.ling,
    e.vivah,
    e.dharm,
    e.jaati,
    e.dccb,
    e.email,
    e.mobile,
    e.nominee

  ]);


  const wb = XLSX.utils.book_new();


  const titleRows = [

    ['PACS COMPUTERIZATION PROJECT'],

    ['Membership डेटा एंट्री टेम्पलेट'],

    [
      'जिला – रोहतास',
      '',
      '',
      '',
      'प्रखण्ड – नोखा',
      '',
      '',
      '',
      'पैक्स– नोख नगर पंचायत'
    ],

    headers,

    ...rows

  ];


  const ws = XLSX.utils.aoa_to_sheet(titleRows);


  ws['!cols'] = [

    { wch: 6 },
    { wch: 20 },
    { wch: 20 },
    { wch: 20 },
    { wch: 14 },
    { wch: 16 },
    { wch: 12 },
    { wch: 6 },
    { wch: 8 },
    { wch: 14 },
    { wch: 10 },
    { wch: 12 },
    { wch: 18 },
    { wch: 22 },
    { wch: 14 },
    { wch: 18 }

  ];


  XLSX.utils.book_append_sheet(
    wb,
    ws,
    'Membership Data'
  );


  XLSX.writeFile(
    wb,
    'PACS_Membership_Data.xlsx'
  );


  showToast(
    '📥 Excel फाइल डाउनलोड हो गई!'
  );
}


// =====================================================
// EXPORT TO PDF
// Hindi Support - Browser Print Method
// =====================================================

function exportPDF() {

  if (entries.length === 0) {

    showToast(
      '⚠️ पहले कुछ एंट्री जोड़ें!',
      '#c0392b'
    );

    return;
  }


  // =========================
  // TABLE HEADERS
  // =========================

  const headers = [

    'क्र0 सं0',
    'सदस्य का नाम',
    'पिता/पति का नाम',
    'पता',
    'सदस्य/सहसदस्य',
    'आधार संख्या',
    'जन्म तिथि',
    'आयु',
    'लिंग',
    'वैवाहिक स्थिति',
    'धर्म',
    'जाति',
    'DCCB खाता',
    'ईमेल',
    'मोबाईल',
    'नॉमिनी'

  ];


  // =========================
  // CREATE TABLE ROWS
  // =========================

  const tableRows = entries.map(e => `

    <tr>

      <td>${e.sr || ''}</td>

      <td>${e.naam || ''}</td>

      <td>${e.pita || ''}</td>

      <td>${e.pata || ''}</td>

      <td>${e.sadasy || ''}</td>

      <td>${e.aadhar || ''}</td>

      <td>${e.dob || ''}</td>

      <td>${e.aayu || ''}</td>

      <td>${e.ling || ''}</td>

      <td>${e.vivah || ''}</td>

      <td>${e.dharm || ''}</td>

      <td>${e.jaati || ''}</td>

      <td>${e.dccb || ''}</td>

      <td>${e.email || ''}</td>

      <td>${e.mobile || ''}</td>

      <td>${e.nominee || ''}</td>

    </tr>

  `).join('');


  // =========================
  // CREATE HEADER
  // =========================

  const tableHeader = headers.map(header => `

    <th>${header}</th>

  `).join('');


  // =========================
  // PRINT WINDOW
  // =========================

  const printWindow = window.open('', '_blank');


  if (!printWindow) {

    showToast(
      '⚠️ Popup blocked है। Browser में popup allow करें।',
      '#c0392b'
    );

    return;
  }


  // =========================
  // PRINT HTML
  // =========================

  printWindow.document.write(`

<!DOCTYPE html>

<html lang="hi">

<head>

<meta charset="UTF-8">

<meta name="viewport"
      content="width=device-width, initial-scale=1.0">

<title>PACS Membership Data</title>


<style>

  * {
    box-sizing: border-box;
  }


  body {

    font-family:

      "Noto Sans Devanagari",
      "Mangal",
      "Arial Unicode MS",
      Arial,
      sans-serif;

    margin: 0;

    padding: 8mm;

    color: #000;

  }


  .title {

    text-align: center;

    font-size: 18px;

    font-weight: bold;

    margin-bottom: 3px;

  }


  .subtitle {

    text-align: center;

    font-size: 11px;

    margin-bottom: 10px;

  }


  .details {

    display: flex;

    justify-content: space-between;

    font-size: 10px;

    margin-bottom: 7px;

  }


  table {

    width: 100%;

    border-collapse: collapse;

    table-layout: fixed;

    font-size: 6.5px;

  }


  th {

    background: #1a3c6e;

    color: white;

    border: 1px solid #000;

    padding: 4px 2px;

    text-align: center;

    font-weight: bold;

  }


  td {

    border: 1px solid #555;

    padding: 3px 2px;

    text-align: center;

    vertical-align: middle;

    word-wrap: break-word;

    overflow-wrap: break-word;

  }


  tr:nth-child(even) {

    background: #f5f5f5;

  }


  th:nth-child(1),
  td:nth-child(1) {
    width: 4%;
  }


  th:nth-child(2),
  td:nth-child(2) {
    width: 9%;
  }


  th:nth-child(3),
  td:nth-child(3) {
    width: 9%;
  }


  th:nth-child(4),
  td:nth-child(4) {
    width: 9%;
  }


  th:nth-child(5),
  td:nth-child(5) {
    width: 7%;
  }


  th:nth-child(6),
  td:nth-child(6) {
    width: 8%;
  }


  th:nth-child(7),
  td:nth-child(7) {
    width: 7%;
  }


  th:nth-child(8),
  td:nth-child(8) {
    width: 4%;
  }


  th:nth-child(9),
  td:nth-child(9) {
    width: 5%;
  }


  th:nth-child(10),
  td:nth-child(10) {
    width: 8%;
  }


  th:nth-child(11),
  td:nth-child(11) {
    width: 5%;
  }


  th:nth-child(12),
  td:nth-child(12) {
    width: 6%;
  }


  th:nth-child(13),
  td:nth-child(13) {
    width: 8%;
  }


  th:nth-child(14),
  td:nth-child(14) {
    width: 10%;
  }


  th:nth-child(15),
  td:nth-child(15) {
    width: 7%;
  }


  th:nth-child(16),
  td:nth-child(16) {
    width: 8%;
  }


  @page {

    size: A4 landscape;

    margin: 6mm;

  }


  @media print {

    body {

      padding: 0;

    }

    thead {

      display: table-header-group;

    }

    tr {

      page-break-inside: avoid;

    }

  }

</style>

</head>


<body>


  <div class="title">

    PACS COMPUTERIZATION PROJECT

  </div>


  <div class="subtitle">

    Membership डेटा एंट्री

  </div>


  <div class="details">

    <span>
      जिला – रोहतास
    </span>

    <span>
      प्रखण्ड – नोखा
    </span>

    <span>
      पैक्स – नोख नगर पंचायत
    </span>

  </div>


  <table>

    <thead>

      <tr>

        ${tableHeader}

      </tr>

    </thead>


    <tbody>

      ${tableRows}

    </tbody>

  </table>


</body>

</html>

  `);


  printWindow.document.close();


  // =========================
  // PRINT
  // =========================

  printWindow.onload = function () {

    setTimeout(() => {

      printWindow.focus();

      printWindow.print();

    }, 500);

  };


  showToast(
    '📄 PDF के लिए Print window खुल रही है...'
  );
}


// =========================
// INIT
// =========================

updateSrField();

renderTable();