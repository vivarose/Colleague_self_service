let busyBlocks = [];

function toggleFilterPanel() {
  const filterBody = document.getElementById('filterBody');
  const caret = document.getElementById('panelCaret');
  const toggleText = document.getElementById('toggleText');
  const headerBtn = document.getElementById('headerToggleBtn');

  if (filterBody.classList.contains('hidden')) {
    filterBody.classList.remove('hidden');
    caret.textContent = '▲';
    toggleText.textContent = 'Hide Filters';
    headerBtn.textContent = '▲ Hide';
  } else {
    filterBody.classList.add('hidden');
    caret.textContent = '▼';
    toggleText.textContent = 'Show Filters';
    headerBtn.textContent = '▼ Show';
  }
}

function parseTimeStringToMinutes(timeStr) {
  if (!timeStr) return null;
  const parts = timeStr.split(':');
  if (parts.length < 2) return null;
  return parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
}

function formatMinutesTo12H(minutes) {
  let h = Math.floor(minutes / 60);
  let m = minutes % 60;
  const ampm = h >= 12 ? 'PM' : 'AM';
  h = h % 12;
  if (h === 0) h = 12;
  const mStr = m < 10 ? '0' + m : m;
  return `${h}:${mStr} ${ampm}`;
}

function addBusyBlock() {
  const dayCheckboxes = document.querySelectorAll('.busy-day:checked');
  const selectedDays = Array.from(dayCheckboxes).map(cb => cb.value);

  const startStr = document.getElementById('busyStartTime').value;
  const endStr = document.getElementById('busyEndTime').value;

  const startMin = parseTimeStringToMinutes(startStr);
  const endMin = parseTimeStringToMinutes(endStr);

  if (selectedDays.length === 0) {
    alert('Please select at least one day (M, T, W, R, F).');
    return;
  }
  if (startMin === null || endMin === null || startMin >= endMin) {
    alert('Please enter a valid start and end time (start must be before end).');
    return;
  }

  const dayLabel = selectedDays.join('');
  const label = `${dayLabel} ${formatMinutesTo12H(startMin)} – ${formatMinutesTo12H(endMin)}`;

  busyBlocks.push({
    days: selectedDays,
    start: startMin,
    end: endMin,
    label: label
  });

  renderBusyChips();
  filterCourses();
}

function removeBusyBlock(index) {
  busyBlocks.splice(index, 1);
  renderBusyChips();
  filterCourses();
}

function renderBusyChips() {
  const container = document.getElementById('busyList');
  container.innerHTML = '';
  busyBlocks.forEach((block, idx) => {
    const chip = document.createElement('div');
    chip.className = 'busy-chip';
    chip.innerHTML = `🚫 ${block.label} <button class="remove-btn" onclick="removeBusyBlock(${idx})" title="Remove filter">&times;</button>`;
    container.appendChild(chip);
  });
}

function filterCourses() {
  const query = document.getElementById('searchInput').value.toLowerCase().trim();
  const selectedYear = document.getElementById('classYearFilter').value;
  const hideBefore = parseInt(document.getElementById('timeBeforeSelect').value, 10);
  const hideAfter = parseInt(document.getElementById('timeAfterSelect').value, 10);
  const wiOnly = document.getElementById('wiOnly').checked;
  const noPrereqOnly = document.getElementById('noPrereqOnly').checked;

  const cards = document.querySelectorAll('.course-card');
  let visibleCount = 0;

  cards.forEach(card => {
    let show = true;

    // 1. Search Query Filter
    if (query !== '') {
      const cardText = card.innerText.toLowerCase();
      if (!cardText.includes(query)) {
        show = false;
      }
    }

    // 2. Class Year Eligibility
    if (show && selectedYear !== 'all') {
      const yearStatus = card.getAttribute(`data-${selectedYear}`);
      if (yearStatus === 'no') {
        show = false;
      }
    }

    // 3. Hide Starts Before Boundary
    if (show && !isNaN(hideBefore)) {
      const startMin = parseInt(card.getAttribute('data-start-min'), 10);
      if (!isNaN(startMin) && startMin < hideBefore) {
        show = false;
      }
    }

    // 4. Hide Ends After Boundary
    if (show && !isNaN(hideAfter)) {
      const endMin = parseInt(card.getAttribute('data-end-min'), 10);
      if (!isNaN(endMin) && endMin > hideAfter) {
        show = false;
      }
    }

    // 5. Writing Intensive Check
    if (show && wiOnly) {
      if (card.getAttribute('data-wi') !== 'true') {
        show = false;
      }
    }

    // 6. Prerequisites Check
    if (show && noPrereqOnly) {
      if (card.getAttribute('data-has-prereq') === 'true') {
        show = false;
      }
    }

    // 7. Busy Time Block / Schedule Conflict Exclusion
    if (show && busyBlocks.length > 0) {
      const slotsRaw = card.getAttribute('data-slots');
      if (slotsRaw) {
        try {
          const courseSlots = JSON.parse(slotsRaw);
          for (let b of busyBlocks) {
            for (let cSlot of courseSlots) {
              const cDays = cSlot.days.split('');
              const hasDayOverlap = cDays.some(d => b.days.includes(d));

              if (hasDayOverlap) {
                if (cSlot.start < b.end && cSlot.end > b.start) {
                  show = false;
                  break;
                }
              }
            }
            if (!show) break;
          }
        } catch (e) {
          console.error('Error parsing course slots JSON:', e);
        }
      }
    }

    // Toggle Visibility
    if (show) {
      card.classList.remove('hidden');
      visibleCount++;
    } else {
      card.classList.add('hidden');
    }
  });

  // Update Counter Badge
  const counterBadge = document.getElementById('courseCounter');
  if (counterBadge) {
    counterBadge.textContent = `Showing ${visibleCount} of ${cards.length} courses`;
  }
}

function resetFilters() {
  document.getElementById('searchInput').value = '';
  document.getElementById('classYearFilter').value = 'all';
  document.getElementById('timeBeforeSelect').value = '';
  document.getElementById('timeAfterSelect').value = '';
  document.getElementById('wiOnly').checked = false;
  document.getElementById('noPrereqOnly').checked = false;

  document.querySelectorAll('.busy-day').forEach(cb => cb.checked = false);
  busyBlocks = [];
  renderBusyChips();

  filterCourses();
}

// Attach Event Listeners on DOM Load
window.addEventListener('DOMContentLoaded', () => {
  document.getElementById('searchInput').addEventListener('input', filterCourses);
  document.getElementById('classYearFilter').addEventListener('change', filterCourses);
  document.getElementById('timeBeforeSelect').addEventListener('change', filterCourses);
  document.getElementById('timeAfterSelect').addEventListener('change', filterCourses);
  document.getElementById('wiOnly').addEventListener('change', filterCourses);
  document.getElementById('noPrereqOnly').addEventListener('change', filterCourses);

  filterCourses();
});