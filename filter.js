// Hamilton Course Explorer Filtering Logic

let busyBlocks = [];

async function updateVisitCount() {
  const counterElement = document.getElementById('visitCounter');
  if (!counterElement) return;

  try {
    // Replace 'hamilton-course-explorer' and 'page-views' with unique strings for your project
    const namespace = 'hamilton-course-explorer';
    const key = 'page-views';
    
    const response = await fetch(`https://api.counterapi.dev/v1/${namespace}/${key}/up`);
    const data = await response.json();
    
    if (data && data.count !== undefined) {
      counterElement.textContent = `Visits: ${data.count.toLocaleString()}`;
    } else {
      counterElement.textContent = `Visits: 1`;
    }
  } catch (err) {
    console.error('Visit counter error:', err);
    counterElement.style.display = 'none'; // Gracefully hide if offline or blocked
  }
}

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

function hideCourseCard(btn) {
  const card = btn.closest('.course-card');
  card.setAttribute('data-user-hidden', 'true');
  filterCourses();
}

function shuffleCourses() {
  const container = document.getElementById('courseContainer');
  if (!container) return;

  // Get all course cards currently in the DOM
  const cards = Array.from(container.children);

  // Fisher-Yates shuffle algorithm
  for (let i = cards.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [cards[i], cards[j]] = [cards[j], cards[i]];
  }

  // Re-append nodes in the new random order
  cards.forEach(card => container.appendChild(card));
}

function filterCourses() {
  const query = document.getElementById('searchInput').value.toLowerCase().trim();
  const selectedYear = document.getElementById('classYearFilter').value;
  const hideBefore = parseInt(document.getElementById('timeBeforeSelect').value, 10);
  const hideAfter = parseInt(document.getElementById('timeAfterSelect').value, 10);
  const wiOnly = document.getElementById('wiOnly').checked;
  const noPrereqOnly = document.getElementById('noPrereqOnly').checked;

  // Safe fetch for department dropdown value
  const deptEl = document.getElementById('deptFilter');
  const selectedDept = deptEl ? deptEl.value : 'all';

  const cards = document.querySelectorAll('.course-card');
  let visibleCount = 0;

  cards.forEach(card => {
    let show = true;


	// Check manual hide status first
    if (card.getAttribute('data-user-hidden') === 'true') {
      show = false;
    }

    // Filter by department
    if (show && selectedDept !== 'all') {
      const cardDept = card.getAttribute('data-subject');
      if (cardDept !== selectedDept) {
        show = false;
      }
    }
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

  const counterEl = document.getElementById('courseCounter');
  if (counterEl) {
    counterEl.textContent = `Showing ${visibleCount} of ${cards.length} courses`;
  }
}

function resetFilters() {
  // Reset text and dropdown inputs
  document.getElementById('deptFilter').value = 'all';
  document.getElementById('searchInput').value = '';
  document.getElementById('classYearFilter').value = 'all';
  document.getElementById('timeBeforeSelect').value = '';
  document.getElementById('timeAfterSelect').value = '';

  // Reset checkboxes
  document.getElementById('wiOnly').checked = false;
  document.getElementById('noPrereqOnly').checked = false;

  // Clear schedule conflict selections
  document.querySelectorAll('.busy-day').forEach(cb => cb.checked = false);
  busyBlocks = [];
  renderBusyChips();

  // Re-run filter engine to unhide all cards and update counter
  filterCourses();
}

function getVisibleBookletCards() {
  const cards = Array.from(document.querySelectorAll('#courseContainer > .course-card:not(.hidden)'));
  const collator = new Intl.Collator('en', { numeric: true, sensitivity: 'base' });
  return cards.sort((first, second) =>
    collator.compare(first.dataset.subject, second.dataset.subject) ||
    collator.compare(first.querySelector('h3').textContent, second.querySelector('h3').textContent)
  );
}

function getCurrentCatalogSnapshot() {
  const element = document.getElementById('catalogSnapshot');
  if (!element) throw new Error('The current catalog snapshot is missing from this page.');
  return JSON.parse(element.textContent);
}

function validateBookletBaseline(baseline) {
  if (baseline?.format !== 'hamilton-course-booklet-baseline' || baseline.format_version !== 1) {
    throw new Error('This is not a supported Hamilton booklet baseline file.');
  }
  const snapshot = baseline.snapshot;
  if (!snapshot || snapshot.schema_version !== 1 || !Array.isArray(snapshot.records)) {
    throw new Error('The baseline does not contain a supported catalog snapshot.');
  }
  if (baseline.snapshot_id !== snapshot.snapshot_id || baseline.term !== snapshot.term ||
      typeof snapshot.snapshot_id !== 'string' || !/^[a-f0-9]{64}$/.test(snapshot.snapshot_id) ||
      typeof snapshot.term !== 'string' || !snapshot.term ||
      typeof baseline.created_at !== 'string' || !Number.isFinite(Date.parse(baseline.created_at)) ||
      !Array.isArray(baseline.selected_record_ids) || !baseline.selected_record_ids.length ||
      new Set(baseline.selected_record_ids).size !== baseline.selected_record_ids.length) {
    throw new Error('The baseline metadata does not match its catalog snapshot.');
  }
  const recordIds = new Set();
  for (const record of snapshot.records) {
    if (!record || typeof record.id !== 'string' || !Array.isArray(record.variants) ||
        record.variants.length === 0 || recordIds.has(record.id) ||
        record.variants.some(variant => !variant || typeof variant !== 'object' ||
          Array.isArray(variant) || Object.values(variant).some(value => typeof value !== 'string'))) {
      throw new Error('The baseline catalog contains an invalid or duplicate section record.');
    }
    recordIds.add(record.id);
  }
  if (baseline.selected_record_ids.some(id => typeof id !== 'string' || !recordIds.has(id))) {
    throw new Error('The baseline selection refers to a section missing from its catalog snapshot.');
  }
  return baseline;
}

function compareCatalogSnapshots(baseline, current, selectedIds) {
  if (baseline.term !== current.term || baseline.schema_version !== current.schema_version) {
    throw new Error('The baseline and current catalog must use the same term and schema.');
  }
  const baselineById = new Map(baseline.records.map(record => [record.id, record]));
  const currentById = new Map(current.records.map(record => [record.id, record]));
  const result = { changed: [], removed: [], added: [], ambiguous: [], unchanged: 0 };

  for (const id of selectedIds) {
    const previous = baselineById.get(id);
    if (!previous) throw new Error(`Selected section ${id} is missing from the baseline snapshot.`);
    const latest = currentById.get(id);
    if (previous.variants.length !== 1 || (latest && latest.variants.length !== 1)) {
      result.ambiguous.push({
        id,
        baseline_count: previous.variants.length,
        current_count: latest?.variants.length || 0,
        record: latest || previous,
      });
    } else if (!latest) {
      result.removed.push({ id, record: previous });
    } else if ([...new Set([
      ...Object.keys(previous.variants[0]), ...Object.keys(latest.variants[0])
    ])].some(field => previous.variants[0][field] !== latest.variants[0][field])) {
      result.changed.push({ id, before: previous, current: latest });
    } else {
      result.unchanged++;
    }
  }

  for (const record of current.records) {
    if (baselineById.has(record.id)) continue;
    if (record.variants.length === 1) {
      result.added.push(record);
    } else {
      result.ambiguous.push({
        id: record.id,
        baseline_count: 0,
        current_count: record.variants.length,
        record,
      });
    }
  }
  return result;
}

function setCorrigendumStatus(message, isError = false) {
  const status = document.getElementById('corrigendumStatus');
  if (!status) return;
  status.textContent = message;
  status.setAttribute('aria-live', isError ? 'assertive' : 'polite');
}

function saveBookletBaseline() {
  try {
    const cards = getVisibleBookletCards();
    if (!cards.length) throw new Error('Select at least one visible course before saving a baseline.');
    const snapshot = getCurrentCatalogSnapshot();
    const metadata = JSON.parse(document.getElementById('catalogMetadata').textContent);
    const selectedIds = [];
    for (const card of cards) {
      const ids = JSON.parse(card.dataset.recordIds || '[]');
      if (!ids.length) throw new Error('A selected course is missing its catalog section identity.');
      selectedIds.push(...ids);
    }
    const baseline = {
      format: 'hamilton-course-booklet-baseline',
      format_version: 1,
      created_at: new Date().toISOString(),
      term: snapshot.term,
      term_label: snapshot.term_label,
      snapshot_id: snapshot.snapshot_id,
      selected_course_count: cards.length,
      selected_record_ids: [...new Set(selectedIds)],
      snapshot,
    };
    validateBookletBaseline(baseline);
    const blob = new Blob([JSON.stringify(baseline, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `${snapshot.term}-booklet-baseline.json`;
    link.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    setCorrigendumStatus(
      `Saved ${cards.length} selected course${cards.length === 1 ? '' : 's'} from snapshot ${metadata.snapshot_id}. Keep this JSON file to compare a future catalog.`
    );
  } catch (error) {
    setCorrigendumStatus(`Could not save booklet baseline: ${error.message}`, true);
    console.error('Could not save booklet baseline:', error);
  }
}

function recordLabel(record) {
  const row = record.variants[0] || {};
  const subject = row['Subject Code'] || '';
  const number = row['Course Number+'] || row['Course Number'] || '';
  const section = row.Section ? `-${row.Section}` : '';
  const title = row['Course Name'] || row.TITLE || '';
  return `${subject} ${number}${section}${title ? `: ${title}` : ''}`.trim();
}

const CORRIGENDUM_FIELD_LABELS = {
  MAXCAP: 'Seat cap', 'Sec Capacity': 'Section capacity', 'Fr Cap': 'First-year cap',
  'So Cap': 'Sophomore cap', 'Jr Cap': 'Junior cap', 'Sr Cap': 'Senior cap',
  'Course Name': 'Name', TITLE: 'Title', Active: 'Active', Distribution: 'Distribution',
  Prerequisites: 'Prerequisites', Corequisites: 'Corequisites', COMMENTS: 'Comments',
  CREDIT: 'Credit', TEACHINGMETHOD: 'Teaching method', RESTRICTED: 'Restricted',
  'Writing-intensive': 'Writing-intensive',
};
const CORRIGENDUM_IGNORED_FIELDS = new Set(['Course Description']);
const CORRIGENDUM_LONG_FIELDS = new Set(['Notes', 'Prerequisite Narrative',
  'Corequisite Narrative', 'Credit Hours Narrative', 'Cross Listed Courses Narrative']);

function corrigendumFieldLabel(field) {
  if (CORRIGENDUM_FIELD_LABELS[field]) return CORRIGENDUM_FIELD_LABELS[field];
  let match = /^DAY(\d)$/.exec(field);
  if (match) return `Meeting ${match[1]} days`;
  match = /^TIME(\d)-(start|end)$/.exec(field);
  if (match) return `Meeting ${match[1]} ${match[2]}`;
  match = /^ROOM(\d)$/.exec(field);
  if (match) return `Meeting ${match[1]} room`;
  match = /^INSTRUCTOR(\d)$/.exec(field);
  if (match) return `Instructor ${match[1]}`;
  return field;
}

function corrigendumValue(value) {
  const text = String(value ?? '').trim();
  if (!text) return '(blank)';
  return text.length > 60 ? `${text.slice(0, 57)}...` : text;
}

function describeCorrigendumChange(change) {
  const oldRow = change.before.variants[0];
  const newRow = change.current.variants[0];
  return [...new Set([...Object.keys(oldRow), ...Object.keys(newRow)])]
    .sort()
    .filter(field => oldRow[field] !== newRow[field] && !CORRIGENDUM_IGNORED_FIELDS.has(field))
    .map(field => CORRIGENDUM_LONG_FIELDS.has(field)
      ? `${corrigendumFieldLabel(field)} updated`
      : `${corrigendumFieldLabel(field)}: ${corrigendumValue(oldRow[field])} \u2192 ${corrigendumValue(newRow[field])}`)
    .join('; ');
}

function plainDescription(html) {
  const text = String(html ?? '').replace(/<[^>]*>/g, ' ')
    .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'")
    .replace(/\s+/g, ' ').trim();
  return text;
}

function appendCorrigendumList(container, heading, items) {
  if (!items.length) return;
  const section = document.createElement('section');
  section.className = 'corrigendum-section';
  const title = document.createElement('h2');
  title.textContent = `${heading} (${items.length})`;
  section.appendChild(title);
  const list = document.createElement('ul');
  items.forEach(entry => {
    const { text, detail } = typeof entry === 'string' ? { text: entry } : entry;
    const item = document.createElement('li');
    item.textContent = text;
    if (detail) {
      const description = document.createElement('div');
      description.className = 'corrigendum-description';
      description.textContent = detail;
      item.appendChild(description);
    }
    list.appendChild(item);
  });
  section.appendChild(list);
  container.appendChild(section);
}

function prepareCorrigendumBooklet() {
  if (!loadedCorrigendum) return;
  const booklet = document.getElementById('courseBooklet');
  if (!booklet) return;
  const { baseline, current, changes } = loadedCorrigendum;
  const heading = document.createElement('h1');
  heading.textContent = `Hamilton Course Corrigendum: ${current.term_label || baseline.term_label}`;
  const provenance = document.createElement('p');
  provenance.className = 'booklet-provenance';
  provenance.textContent = `Baseline snapshot: ${baseline.snapshot_id} (${baseline.created_at}); current snapshot: ${current.snapshot_id} (${current.generated_at}).`;
  const changedEntries = changes.changed
    .map(change => ({ change, summary: describeCorrigendumChange(change) }))
    .filter(entry => entry.summary);
  const summary = document.createElement('p');
  summary.textContent = `${changes.removed.length} removed from your selection, ${changes.added.length} new catalog section${changes.added.length === 1 ? '' : 's'}, ${changedEntries.length} changed, ${changes.ambiguous.length} needing manual review, ${changes.unchanged} unchanged.`;
  const disclaimer = document.createElement('p');
  disclaimer.className = 'booklet-disclaimer';
  disclaimer.textContent = 'Unofficial catalog comparison. Verify all changes in Colleague Self-Service.';
  booklet.replaceChildren(heading, provenance, summary, disclaimer);

  appendCorrigendumList(booklet, 'Removed: no longer in the catalog',
    changes.removed.map(change => recordLabel(change.record)));
  appendCorrigendumList(booklet, 'Added: new catalog sections', changes.added.map(record => ({
    text: recordLabel(record),
    detail: plainDescription(record.variants[0]['Course Description']),
  })));
  appendCorrigendumList(booklet, 'Changed', changedEntries.map(({ change, summary }) =>
    `${recordLabel(change.current)} \u2014 ${summary}`));
  appendCorrigendumList(booklet, 'Needs manual review: duplicate section identities',
    changes.ambiguous.map(issue =>
      `${recordLabel(issue.record) || issue.id} \u2014 ${issue.baseline_count} baseline variant(s), ${issue.current_count} current variant(s)`));
  if (!changedEntries.length && !changes.removed.length && !changes.added.length && !changes.ambiguous.length) {
    const noChanges = document.createElement('p');
    noChanges.textContent = 'No changes were found for the saved selection.';
    booklet.appendChild(noChanges);
  }
}

let loadedCorrigendum = null;
let comparisonRequest = 0;

async function loadBookletBaseline(input) {
  const file = input.files?.[0];
  if (!file) return;
  const request = ++comparisonRequest;
  loadedCorrigendum = null;
  document.getElementById('printCorrigendumButton').hidden = true;
  setCorrigendumStatus('Comparing baseline...');
  try {
    const text = await file.text();
    if (request !== comparisonRequest) return;
    const baseline = validateBookletBaseline(JSON.parse(text));
    const current = getCurrentCatalogSnapshot();
    if (baseline.term !== current.term) {
      throw new Error(`This baseline is for ${baseline.term}; the current catalog is ${current.term}.`);
    }
    const changes = compareCatalogSnapshots(baseline.snapshot, current, baseline.selected_record_ids);
    loadedCorrigendum = { baseline, current, changes };
    prepareCorrigendumBooklet();
    document.getElementById('printCorrigendumButton').hidden = false;
    setCorrigendumStatus(
      `Comparison complete: ${changes.changed.length} changed, ${changes.removed.length} removed, ${changes.added.length} new catalog sections, ${changes.ambiguous.length} ambiguous, ${changes.unchanged} unchanged.`
    );
  } catch (error) {
    if (request !== comparisonRequest) return;
    loadedCorrigendum = null;
    document.getElementById('printCorrigendumButton').hidden = true;
    setCorrigendumStatus(`Could not compare baseline: ${error.message}`, true);
    console.error('Could not compare booklet baseline:', error);
  } finally {
    if (request === comparisonRequest) input.value = '';
  }
}

function prepareCourseBooklet() {
  const booklet = document.getElementById('courseBooklet');
  if (!booklet) return;

  const cards = getVisibleBookletCards();

  const metadataElement = document.getElementById('catalogMetadata');
  const metadata = metadataElement ? JSON.parse(metadataElement.textContent) : null;
  const preparedAt = new Date();
  const heading = document.createElement('h1');
  heading.textContent = `Hamilton Course Booklet: ${metadata?.term_label || 'Term unavailable'}`;
  const provenance = document.createElement('p');
  provenance.className = 'booklet-provenance';
  const dateLabel = preparedAt.toLocaleString('en-US', { dateStyle: 'medium', timeStyle: 'short' });
  provenance.textContent = `Prepared: ${dateLabel} (${Intl.DateTimeFormat().resolvedOptions().timeZone}); ${preparedAt.toISOString()}`;
  if (metadata) {
    provenance.textContent += ` | Catalog built: ${metadata.generated_at} | Snapshot: ${metadata.snapshot_id}`;
    provenance.textContent += metadata.source_updated_at ? ` | Source updated: ${metadata.source_updated_at}` : ' | Source update time unavailable';
  } else {
    provenance.textContent += ' | Catalog provenance unavailable';
  }
  const summary = document.createElement('p');
  summary.textContent = `${cards.length} course${cards.length === 1 ? '' : 's'} selected`;
  const filters = [];
  const query = document.getElementById('searchInput').value.trim();
  if (query) filters.push(`Keyword: ${query}`);
  for (const [identifier, label] of [
    ['deptFilter', 'Department'], ['classYearFilter', 'Class year'],
    ['timeBeforeSelect', 'Starts at or after'], ['timeAfterSelect', 'Ends at or before']
  ]) {
    const select = document.getElementById(identifier);
    if (select.value && select.value !== 'all') {
      filters.push(`${label}: ${select.selectedOptions[0].textContent}`);
    }
  }
  if (document.getElementById('wiOnly').checked) filters.push('Writing intensive only');
  if (document.getElementById('noPrereqOnly').checked) filters.push('No prerequisites only');
  busyBlocks.forEach(block => filters.push(`Unavailable: ${block.label}`));
  const hiddenCount = document.querySelectorAll('#courseContainer > [data-user-hidden="true"]').length;
  if (hiddenCount) filters.push(`Manually hidden: ${hiddenCount}`);
  const filterSummary = document.createElement('p');
  filterSummary.textContent = filters.length ? filters.join('; ') : 'All courses; no filters applied.';
  const disclaimer = document.createElement('p');
  disclaimer.className = 'booklet-disclaimer';
  disclaimer.textContent = 'Unofficial planning aid. Filtering is imperfect and course details may change. Verify official listings and complete registration through Colleague Self-Service.';
  booklet.replaceChildren(heading, provenance, summary, filterSummary, disclaimer);
  if (metadata?.identity_issues.length) {
    const identityWarning = document.createElement('p');
    identityWarning.textContent = `Catalog identity warning: ${metadata.identity_issues.length} ambiguous section key(s). These identities require manual review in a corrigendum.`;
    booklet.appendChild(identityWarning);
  }

  if (!cards.length) {
    const empty = document.createElement('p');
    empty.textContent = 'No courses match the current selection.';
    booklet.appendChild(empty);
  }

  let department;
  let section;
  cards.forEach(card => {
    if (card.dataset.subject !== department) {
      department = card.dataset.subject;
      section = document.createElement('section');
      const departmentHeading = document.createElement('h2');
      const option = Array.from(document.getElementById('deptFilter').options)
        .find(item => item.value === department);
      departmentHeading.textContent = option ? option.textContent : department;
      section.appendChild(departmentHeading);
      booklet.appendChild(section);
    }
    const entry = card.cloneNode(true);
    entry.className = 'booklet-course';
    entry.querySelectorAll('button').forEach(button => button.remove());
    [entry, ...entry.querySelectorAll('[style]')].forEach(element => element.removeAttribute('style'));
    entry.querySelectorAll('b, strong').forEach(label => {
      if (label.previousSibling?.nodeType === Node.TEXT_NODE) {
        label.previousSibling.textContent = label.previousSibling.textContent.replace(/[\p{Extended_Pictographic}\uFE0F]/gu, '');
      }
      if (label.firstChild?.nodeType === Node.TEXT_NODE) {
        label.firstChild.textContent = label.firstChild.textContent.replace(/^\p{Extended_Pictographic}\uFE0F?\s*/u, '');
      }
    });
    const descriptionLabel = entry.children[2]?.firstElementChild;
    if (descriptionLabel?.textContent.trim() === '\u{1F4D6} Description') {
      descriptionLabel.textContent = 'Description';
      descriptionLabel.className = 'booklet-description-label';
    }
    section.appendChild(entry);
  });
}

function printCourseBooklet() {
  printMode = 'booklet';
  prepareCourseBooklet();
  window.print();
}

function printCorrigendum() {
  if (!loadedCorrigendum) {
    setCorrigendumStatus('Load a valid baseline before printing a corrigendum.', true);
    return;
  }
  printMode = 'corrigendum';
  prepareCorrigendumBooklet();
  window.print();
}

let printMode = 'booklet';

async function requestDailyRefresh() {
  const status = document.getElementById('refreshStatus');
  if (!status || location.protocol === 'file:') return;
  try {
    const configResponse = await fetch(new URL('refresh-config.json', location.href), { cache: 'no-store' });
    if (!configResponse.ok) throw new Error(`Refresh configuration HTTP ${configResponse.status}`);
    const config = await configResponse.json();
    if (!config.service_url) {
      status.textContent = 'Daily refresh is not configured; showing the published catalog.';
      return;
    }
    const service = new URL(config.service_url);
    if (service.protocol !== 'https:') throw new Error('Refresh service must use HTTPS.');
    let path = '/refresh';
    for (let attempt = 0; attempt < 120; attempt++) {
      const response = await fetch(new URL(path, service), {
        method: path === '/refresh' ? 'POST' : 'GET', cache: 'no-store'
      });
      if (!response.ok) throw new Error(`Refresh service HTTP ${response.status}`);
      const result = await response.json();
      if (result.status === 'succeeded') {
        const metadata = JSON.parse(document.getElementById('catalogMetadata').textContent);
        status.textContent = `Catalog refreshed for ${result.date} (fixed EST). `;
        if (metadata.refresh_date !== result.date) {
          const link = document.createElement('a');
          const updated = new URL(location.href);
          updated.searchParams.set('refresh', Date.now().toString());
          link.href = updated.href;
          link.textContent = 'Reload to use the refreshed catalog (resets filters).';
          status.appendChild(link);
        }
        return;
      }
      if (result.status === 'failed') {
        throw new Error(result.message || 'The refresh workflow failed; a later visit may retry.');
      }
      if (!['pending', 'not_requested'].includes(result.status)) {
        throw new Error('Refresh service returned an unknown status.');
      }
      status.textContent = 'Catalog refresh queued or running; showing the last published data.';
      path = result.status === 'not_requested' ? '/refresh' : '/status';
      await new Promise(resolve => setTimeout(resolve, 15_000));
    }
    throw new Error('Refresh is still pending. Reopen the app later to check its status.');
  } catch (error) {
    status.textContent = `Daily refresh unavailable: ${error.message} Showing the last published data.`;
    console.error('Daily catalog refresh:', error);
  }
}

if (typeof window !== 'undefined') {
  window.addEventListener('beforeprint', () => {
    if (printMode === 'corrigendum') prepareCorrigendumBooklet();
    else prepareCourseBooklet();
  });
  window.addEventListener('afterprint', () => {
    document.getElementById('courseBooklet')?.replaceChildren();
    printMode = 'booklet';
  });

  window.addEventListener('DOMContentLoaded', () => {
    requestDailyRefresh();
    updateVisitCount();
    document.getElementById('searchInput')?.addEventListener('input', filterCourses);
    document.getElementById('classYearFilter')?.addEventListener('change', filterCourses);
    document.getElementById('deptFilter')?.addEventListener('change', filterCourses);
    document.getElementById('timeBeforeSelect')?.addEventListener('change', filterCourses);
    document.getElementById('timeAfterSelect')?.addEventListener('change', filterCourses);
    document.getElementById('wiOnly')?.addEventListener('change', filterCourses);
    document.getElementById('noPrereqOnly')?.addEventListener('change', filterCourses);

    filterCourses();
  });
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = { compareCatalogSnapshots, validateBookletBaseline };
}