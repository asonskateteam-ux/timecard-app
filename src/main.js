import './style.css';
import ExcelJS from 'exceljs';

let entries =
  JSON.parse(localStorage.getItem('timecardEntries')) || [];

const DB_NAME = 'CooperTimecardDatabase';
const DB_VERSION = 2;

const SUBMITTED_STORE = 'submittedTimecards';
const TEMPLATE_STORE = 'templateStore';

const modal = document.querySelector('#modal');
const lunchModal = document.querySelector('#lunchModal');

const entriesContainer = document.querySelector('#entries');
const totalHoursElement = document.querySelector('#totalHours');
const calculatedHoursElement =
  document.querySelector('#calculatedHours');

const roInput = document.querySelector('#ro');
const jcInput = document.querySelector('#jc');
const descriptionInput = document.querySelector('#description');
const startTimeInput = document.querySelector('#startTime');
const finishTimeInput = document.querySelector('#finishTime');

const timecardPage = document.querySelector('#timecardPage');
const submittedPage = document.querySelector('#submittedPage');

const timecardTab = document.querySelector('#timecardTab');
const submittedTab = document.querySelector('#submittedTab');

const submittedTimecardsContainer =
  document.querySelector('#submittedTimecards');

const templateStatus =
  document.querySelector('#templateStatus');

const selectTemplateButton =
  document.querySelector('#selectTemplateButton');

const templateFileInput =
  document.querySelector('#templateFileInput');

document.querySelector('#currentDate').textContent =
  new Date().toLocaleDateString('en-AU', {
    weekday: 'long',
    day: 'numeric',
    month: 'long',
    year: 'numeric',
  });

// ==================================================
// NAVIGATION
// ==================================================

timecardTab.addEventListener('click', () => {
  showPage('timecard');
});

submittedTab.addEventListener('click', async () => {
  showPage('submitted');
  await renderSubmittedTimecards();
});

function showPage(page) {
  timecardPage.classList.remove('active-page');
  submittedPage.classList.remove('active-page');

  timecardTab.classList.remove('active-nav');
  submittedTab.classList.remove('active-nav');

  if (page === 'timecard') {
    timecardPage.classList.add('active-page');
    timecardTab.classList.add('active-nav');
  }

  if (page === 'submitted') {
    submittedPage.classList.add('active-page');
    submittedTab.classList.add('active-nav');
  }
}

// ==================================================
// TEMPLATE SELECTION
// ==================================================

selectTemplateButton.addEventListener('click', () => {
  templateFileInput.click();
});

templateFileInput.addEventListener('change', async () => {
  const file = templateFileInput.files[0];

  if (!file) return;

  if (!file.name.toLowerCase().endsWith('.xlsx')) {
    alert('Please select an Excel .xlsx file.');
    templateFileInput.value = '';
    return;
  }

  try {
    const arrayBuffer = await file.arrayBuffer();

    // Test that ExcelJS can actually read the workbook
    const testWorkbook = new ExcelJS.Workbook();
    await testWorkbook.xlsx.load(arrayBuffer);

    await saveTemplate(file.name, file);

    templateStatus.textContent = `Selected: ${file.name}`;
    templateStatus.classList.add('template-ready');

    selectTemplateButton.textContent = 'Change Template';

    alert('Template saved on this device.');
  } catch (error) {
    console.error(error);

    alert(
      'That file could not be read as an Excel template.'
    );
  }

  templateFileInput.value = '';
});

// ==================================================
// WORK ENTRY BUTTONS
// ==================================================

document
  .querySelector('#addEntryButton')
  .addEventListener('click', () => {
    clearForm();
    modal.classList.remove('hidden');
  });

document
  .querySelector('#cancelButton')
  .addEventListener('click', () => {
    modal.classList.add('hidden');
  });

document
  .querySelector('#saveButton')
  .addEventListener('click', saveEntry);

startTimeInput.addEventListener(
  'input',
  updateCalculatedHours
);

finishTimeInput.addEventListener(
  'input',
  updateCalculatedHours
);

// ==================================================
// LUNCH BUTTONS
// ==================================================

document
  .querySelector('#lunchButton')
  .addEventListener('click', () => {
    lunchModal.classList.remove('hidden');
  });

document
  .querySelector('#cancelLunchButton')
  .addEventListener('click', () => {
    lunchModal.classList.add('hidden');
  });

document
  .querySelector('#lunch30Button')
  .addEventListener('click', () => {
    addLunch(0.5);
  });

document
  .querySelector('#lunch60Button')
  .addEventListener('click', () => {
    addLunch(1);
  });

// ==================================================
// FINISH TIMECARD
// ==================================================

document
  .querySelector('#finishButton')
  .addEventListener('click', async () => {
    await generateExcelTimecard();
  });

// ==================================================
// TIME CALCULATION
// ==================================================

function calculateHours(start, finish) {
  if (!start || !finish) return 0;

  const [startHour, startMinute] =
    start.split(':').map(Number);

  const [finishHour, finishMinute] =
    finish.split(':').map(Number);

  const startMinutes =
    startHour * 60 + startMinute;

  const finishMinutes =
    finishHour * 60 + finishMinute;

  let difference =
    finishMinutes - startMinutes;

  if (difference < 0) {
    difference += 24 * 60;
  }

  return difference / 60;
}

function updateCalculatedHours() {
  const hours = calculateHours(
    startTimeInput.value,
    finishTimeInput.value
  );

  calculatedHoursElement.textContent =
    `${hours.toFixed(2)} hrs`;
}

// ==================================================
// SAVE WORK ENTRY
// ==================================================

function saveEntry() {
  const ro = roInput.value.trim();
  const jc = jcInput.value.trim();
  const description = descriptionInput.value.trim();

  const start = startTimeInput.value;
  const finish = finishTimeInput.value;

  if (!ro) {
    alert('Enter an RO number.');
    return;
  }

  if (!start || !finish) {
    alert('Enter a start and finish time.');
    return;
  }

  const hours =
    calculateHours(start, finish);

  const entry = {
    id: Date.now(),
    type: 'work',
    ro,
    jc,
    description,
    start,
    finish,
    hours,
  };

  entries.push(entry);

  saveEntries();
  renderEntries();

  modal.classList.add('hidden');
}

// ==================================================
// ADD LUNCH
// ==================================================

function addLunch(lunchHours) {
  const entry = {
    id: Date.now(),
    type: 'lunch',
    description: 'Unpaid Lunch',
    lunchHours,
    hours: 0,
  };

  entries.push(entry);

  saveEntries();
  renderEntries();

  lunchModal.classList.add('hidden');
}

// ==================================================
// DELETE CURRENT ENTRY
// ==================================================

function deleteEntry(id) {
  entries =
    entries.filter(
      (entry) => entry.id !== id
    );

  saveEntries();
  renderEntries();
}

// ==================================================
// LOCAL STORAGE
// ==================================================

function saveEntries() {
  localStorage.setItem(
    'timecardEntries',
    JSON.stringify(entries)
  );
}

// ==================================================
// RENDER CURRENT ENTRIES
// ==================================================

function renderEntries() {
  entriesContainer.innerHTML = '';

  if (entries.length === 0) {
    entriesContainer.innerHTML = `
      <div class="empty">
        No work entered yet.
      </div>
    `;
  }

  entries.forEach((entry) => {
    const card =
      document.createElement('div');

    card.className =
      entry.type === 'lunch'
        ? 'entry-card lunch-entry'
        : 'entry-card';

    if (entry.type === 'lunch') {
      card.innerHTML = `
        <div class="entry-top">
          <div>
            <strong>Lunch</strong>
          </div>

          <strong>
            ${entry.lunchHours.toFixed(1)} hr unpaid
          </strong>
        </div>

        <div class="description">
          Unpaid lunch break
        </div>

        <div class="entry-bottom">
          <span>
            Not included in daily total
          </span>

          <button class="delete-button">
            Delete
          </button>
        </div>
      `;
    } else {
      card.innerHTML = `
        <div class="entry-top">
          <div>
            <strong>
              RO ${entry.ro}
            </strong>

            ${
              entry.jc
                ? `<span class="jc">
                    J/C ${entry.jc}
                   </span>`
                : ''
            }
          </div>

          <strong>
            ${entry.hours.toFixed(2)} hrs
          </strong>
        </div>

        <div class="description">
          ${entry.description || 'No description'}
        </div>

        <div class="entry-bottom">
          <span>
            ${entry.start} – ${entry.finish}
          </span>

          <button class="delete-button">
            Delete
          </button>
        </div>
      `;
    }

    card
      .querySelector('.delete-button')
      .addEventListener('click', () => {
        deleteEntry(entry.id);
      });

    entriesContainer.appendChild(card);
  });

  const total =
    entries.reduce(
      (sum, entry) =>
        sum + (entry.hours || 0),
      0
    );

  totalHoursElement.textContent =
    `${total.toFixed(2)} hrs`;
}

// ==================================================
// CLEAR FORM
// ==================================================

function clearForm() {
  roInput.value = '';
  jcInput.value = '';
  descriptionInput.value = '';

  startTimeInput.value = '';
  finishTimeInput.value = '';

  calculatedHoursElement.textContent =
    '0.00 hrs';
}

// ==================================================
// EXCEL TIME CONVERSION
// ==================================================

function convertTimeToExcel(timeString) {
  if (!timeString) return '';

  const [hours, minutes] =
    timeString.split(':').map(Number);

  return (
    (hours * 60 + minutes) /
    (24 * 60)
  );
}

// ==================================================
// DATABASE
// ==================================================

function openDatabase() {
  return new Promise((resolve, reject) => {
    const request =
      indexedDB.open(
        DB_NAME,
        DB_VERSION
      );

    request.onupgradeneeded = (event) => {
      const db = event.target.result;

      if (
        !db.objectStoreNames.contains(
          SUBMITTED_STORE
        )
      ) {
        db.createObjectStore(
          SUBMITTED_STORE,
          {
            keyPath: 'id',
          }
        );
      }

      if (
        !db.objectStoreNames.contains(
          TEMPLATE_STORE
        )
      ) {
        db.createObjectStore(
          TEMPLATE_STORE,
          {
            keyPath: 'id',
          }
        );
      }
    };

    request.onsuccess = () => {
      resolve(request.result);
    };

    request.onerror = () => {
      reject(request.error);
    };
  });
}

// ==================================================
// SAVE TEMPLATE
// ==================================================

async function saveTemplate(filename, blob) {
  const db =
    await openDatabase();

  return new Promise((resolve, reject) => {
    const transaction =
      db.transaction(
        TEMPLATE_STORE,
        'readwrite'
      );

    const store =
      transaction.objectStore(
        TEMPLATE_STORE
      );

    store.put({
      id: 'mainTemplate',
      filename,
      blob,
      savedAt:
        new Date().toISOString(),
    });

    transaction.oncomplete = () => {
      db.close();
      resolve();
    };

    transaction.onerror = () => {
      db.close();
      reject(transaction.error);
    };
  });
}

// ==================================================
// GET TEMPLATE
// ==================================================

async function getTemplate() {
  const db =
    await openDatabase();

  return new Promise((resolve, reject) => {
    const transaction =
      db.transaction(
        TEMPLATE_STORE,
        'readonly'
      );

    const store =
      transaction.objectStore(
        TEMPLATE_STORE
      );

    const request =
      store.get('mainTemplate');

    request.onsuccess = () => {
      db.close();
      resolve(request.result || null);
    };

    request.onerror = () => {
      db.close();
      reject(request.error);
    };
  });
}

// ==================================================
// TEMPLATE STATUS
// ==================================================

async function updateTemplateStatus() {
  try {
    const template =
      await getTemplate();

    if (!template) {
      templateStatus.textContent =
        'No template selected';

      selectTemplateButton.textContent =
        'Select Template';

      return;
    }

    templateStatus.textContent =
      `Selected: ${template.filename}`;

    templateStatus.classList.add(
      'template-ready'
    );

    selectTemplateButton.textContent =
      'Change Template';
  } catch (error) {
    console.error(error);
  }
}

// ==================================================
// SAVE SUBMITTED TIMECARD
// ==================================================

async function saveSubmittedTimecard(
  filename,
  blob,
  totalHours
) {
  const db =
    await openDatabase();

  return new Promise((resolve, reject) => {
    const transaction =
      db.transaction(
        SUBMITTED_STORE,
        'readwrite'
      );

    const store =
      transaction.objectStore(
        SUBMITTED_STORE
      );

    store.put({
      id: Date.now(),
      filename,
      blob,
      totalHours,
      submittedAt:
        new Date().toISOString(),
    });

    transaction.oncomplete = () => {
      db.close();
      resolve();
    };

    transaction.onerror = () => {
      db.close();
      reject(transaction.error);
    };
  });
}

// ==================================================
// GET SUBMITTED
// ==================================================

async function getSubmittedTimecards() {
  const db =
    await openDatabase();

  return new Promise((resolve, reject) => {
    const transaction =
      db.transaction(
        SUBMITTED_STORE,
        'readonly'
      );

    const store =
      transaction.objectStore(
        SUBMITTED_STORE
      );

    const request =
      store.getAll();

    request.onsuccess = () => {
      const results =
        request.result || [];

      results.sort(
        (a, b) =>
          new Date(b.submittedAt) -
          new Date(a.submittedAt)
      );

      db.close();

      resolve(results);
    };

    request.onerror = () => {
      db.close();
      reject(request.error);
    };
  });
}

// ==================================================
// DELETE SUBMITTED
// ==================================================

async function deleteSubmittedTimecard(id) {
  const db =
    await openDatabase();

  return new Promise((resolve, reject) => {
    const transaction =
      db.transaction(
        SUBMITTED_STORE,
        'readwrite'
      );

    const store =
      transaction.objectStore(
        SUBMITTED_STORE
      );

    store.delete(id);

    transaction.oncomplete = () => {
      db.close();
      resolve();
    };

    transaction.onerror = () => {
      db.close();
      reject(transaction.error);
    };
  });
}

// ==================================================
// RENDER SUBMITTED PAGE
// ==================================================

async function renderSubmittedTimecards() {
  submittedTimecardsContainer.innerHTML = `
    <div class="empty">
      Loading...
    </div>
  `;

  try {
    const timecards =
      await getSubmittedTimecards();

    submittedTimecardsContainer.innerHTML =
      '';

    if (timecards.length === 0) {
      submittedTimecardsContainer.innerHTML = `
        <div class="empty">
          No submitted timecards yet.
        </div>
      `;

      return;
    }

    timecards.forEach((timecard) => {
      const card =
        document.createElement('div');

      card.className =
        'submitted-card';

      const submittedDate =
        new Date(
          timecard.submittedAt
        );

      const displayDate =
        submittedDate.toLocaleDateString(
          'en-AU',
          {
            weekday: 'short',
            day: 'numeric',
            month: 'short',
            year: 'numeric',
          }
        );

      const displayTime =
        submittedDate.toLocaleTimeString(
          'en-AU',
          {
            hour: 'numeric',
            minute: '2-digit',
          }
        );

      card.innerHTML = `
        <div class="submitted-info">
          <strong>
            ${displayDate}
          </strong>

          <span>
            ${timecard.totalHours.toFixed(2)}
            paid hrs
          </span>

          <span class="submitted-time">
            Submitted ${displayTime}
          </span>
        </div>

        <div class="submitted-actions">
          <button class="download-button">
            Download Excel
          </button>

          <button class="archive-delete-button">
            Delete
          </button>
        </div>
      `;

      card
        .querySelector(
          '.download-button'
        )
        .addEventListener(
          'click',
          () => {
            downloadBlob(
              timecard.blob,
              timecard.filename
            );
          }
        );

      card
        .querySelector(
          '.archive-delete-button'
        )
        .addEventListener(
          'click',
          async () => {
            const confirmed =
              confirm(
                'Delete this submitted timecard from the app?'
              );

            if (!confirmed) return;

            await deleteSubmittedTimecard(
              timecard.id
            );

            await renderSubmittedTimecards();
          }
        );

      submittedTimecardsContainer.appendChild(
        card
      );
    });
  } catch (error) {
    console.error(error);

    submittedTimecardsContainer.innerHTML = `
      <div class="empty">
        Could not load submitted timecards.
      </div>
    `;
  }
}

// ==================================================
// DOWNLOAD
// ==================================================

function downloadBlob(blob, filename) {
  const url =
    URL.createObjectURL(blob);

  const link =
    document.createElement('a');

  link.href = url;
  link.download = filename;

  document.body.appendChild(link);

  link.click();

  document.body.removeChild(link);

  setTimeout(() => {
    URL.revokeObjectURL(url);
  }, 1000);
}

// ==================================================
// GENERATE EXCEL
// ==================================================

async function generateExcelTimecard() {
  if (entries.length === 0) {
    alert(
      'There are no entries to export.'
    );
    return;
  }

  if (entries.length > 11) {
    alert(
      'The company timecard only has 11 entry spaces.'
    );
    return;
  }

  try {
    const template =
      await getTemplate();

    if (!template) {
      alert(
        'Please select your Excel template first.'
      );

      return;
    }

    const templateData =
      await template.blob.arrayBuffer();

    const workbook =
      new ExcelJS.Workbook();

    await workbook.xlsx.load(
      templateData
    );

    const worksheet =
      workbook.worksheets[0];

    const today =
      new Date();

    worksheet.getCell('A3').value =
      '2918CL';

    worksheet.getCell('B3').value =
      'Cooper Lane';

    worksheet.getCell('E3').value =
      today;

    worksheet.getCell('E3').numFmt =
      'dd/mm/yyyy';

    const entryRows = [
      6,
      8,
      10,
      12,
      14,
      16,
      18,
      20,
      22,
      24,
      26,
    ];

    entryRows.forEach((row) => {
      worksheet.getCell(`A${row}`).value = '';
      worksheet.getCell(`B${row}`).value = '';
      worksheet.getCell(`C${row}`).value = '';
      worksheet.getCell(`D${row}`).value = '';

      worksheet.getCell(`F${row}`).value = '';
      worksheet.getCell(`F${row + 1}`).value = '';
    });

    entries.forEach(
      (entry, index) => {
        const row =
          entryRows[index];

        if (
          entry.type === 'lunch'
        ) {
          worksheet
            .getCell(`C${row}`)
            .value =
              `UNPAID LUNCH - ${entry.lunchHours} HR`;

          worksheet
            .getCell(`D${row}`)
            .value = 0;

          worksheet
            .getCell(`D${row}`)
            .numFmt = '0.00';
        } else {
          worksheet
            .getCell(`A${row}`)
            .value = entry.ro;

          worksheet
            .getCell(`B${row}`)
            .value = entry.jc;

          worksheet
            .getCell(`C${row}`)
            .value =
              entry.description;

          worksheet
            .getCell(`D${row}`)
            .value =
              entry.hours;

          worksheet
            .getCell(`D${row}`)
            .numFmt = '0.00';

          worksheet
            .getCell(`F${row}`)
            .value =
              convertTimeToExcel(
                entry.start
              );

          worksheet
            .getCell(`F${row}`)
            .numFmt = 'h:mm';

          worksheet
            .getCell(
              `F${row + 1}`
            )
            .value =
              convertTimeToExcel(
                entry.finish
              );

          worksheet
            .getCell(
              `F${row + 1}`
            )
            .numFmt = 'h:mm';
        }
      }
    );

    worksheet
      .getCell('D29')
      .value = {
        formula:
          'D6+D8+D10+D12+D14+D16+D18+D20+D22+D24+D26',
      };

    worksheet
      .getCell('D29')
      .numFmt = '0.00';

    const buffer =
      await workbook.xlsx.writeBuffer();

    const blob =
      new Blob(
        [buffer],
        {
          type:
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        }
      );

    const year =
      today.getFullYear();

    const month =
      String(
        today.getMonth() + 1
      ).padStart(2, '0');

    const day =
      String(
        today.getDate()
      ).padStart(2, '0');

    const filename =
      `Timecard_Cooper_Lane_${year}-${month}-${day}.xlsx`;

    const totalHours =
      entries.reduce(
        (sum, entry) =>
          sum + (entry.hours || 0),
        0
      );

    await saveSubmittedTimecard(
      filename,
      blob,
      totalHours
    );

    downloadBlob(
      blob,
      filename
    );

    alert(
      'Timecard created and saved under Submitted.'
    );
  } catch (error) {
    console.error(error);

    alert(
      'There was a problem creating the Excel file.\n\n' +
      error.message
    );
  }
}

// ==================================================
// START APP
// ==================================================

renderEntries();
updateTemplateStatus();