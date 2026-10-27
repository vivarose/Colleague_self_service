const assert = require('node:assert/strict');
const { readFileSync } = require('node:fs');
const { test } = require('node:test');
const vm = require('node:vm');
const { compareCatalogSnapshots, validateBookletBaseline } = require('../filter.js');

function baselineFile() {
  const record = {
    id: 'lecture',
    variants: [{ 'Subject Code': 'BIO', 'Course Number+': '101', Section: '1', TITLE: 'Biology' }]
  };
  return {
    format: 'hamilton-course-booklet-baseline',
    format_version: 1,
    created_at: '2026-10-27T00:00:00Z',
    term: '2027-spring',
    snapshot_id: 'a'.repeat(64),
    selected_record_ids: ['lecture'],
    snapshot: {
      term: '2027-spring', term_label: 'Spring 2027', schema_version: 1,
      snapshot_id: 'a'.repeat(64), records: [record]
    }
  };
}

test('baseline validation rejects malformed variants, metadata, and selections', () => {
  assert.ok(validateBookletBaseline(baselineFile()));
  const changes = [
    file => { file.snapshot.records[0].variants = [null]; },
    file => { file.snapshot.records[0].variants = [{ TITLE: 123 }]; },
    file => { file.snapshot.records.push(file.snapshot.records[0]); },
    file => { file.selected_record_ids = []; },
    file => { file.selected_record_ids = ['lecture', 'lecture']; },
    file => { file.selected_record_ids = ['missing']; },
    file => { file.created_at = 'not a date'; },
    file => { delete file.snapshot_id; delete file.snapshot.snapshot_id; },
    file => { file.term = '2027-fall'; }
  ];
  changes.forEach(change => {
    const file = baselineFile();
    change(file);
    assert.throws(() => validateBookletBaseline(file));
  });
});

test('field ordering does not create false changes; lab changes remain visible', () => {
  const baseline = baselineFile().snapshot;
  const current = structuredClone(baseline);
  current.records[0].variants[0] = Object.fromEntries(
    Object.entries(current.records[0].variants[0]).reverse()
  );
  assert.equal(compareCatalogSnapshots(baseline, current, ['lecture']).unchanged, 1);
  baseline.records.push({ id: 'lab', variants: [{ ROOM1: 'Science 100' }] });
  current.records.push({ id: 'lab', variants: [{ ROOM1: 'Science 200' }] });
  const changes = compareCatalogSnapshots(baseline, current, ['lecture', 'lab']);
  assert.equal(changes.unchanged, 1);
  assert.equal(changes.changed[0].id, 'lab');
});

function browserHarness(snapshot) {
  class Element {
    constructor() {
      this.children = [];
      this.textContent = '';
      this.value = '';
      this.checked = false;
      this.hidden = true;
    }
    appendChild(child) { this.children.push(child); }
    replaceChildren(...children) { this.children = children; }
    setAttribute() {}
  }
  const elements = new Map([
    'catalogSnapshot', 'catalogMetadata', 'courseBooklet', 'corrigendumStatus',
    'printCorrigendumButton', 'searchInput', 'deptFilter', 'classYearFilter',
    'timeBeforeSelect', 'timeAfterSelect', 'wiOnly', 'noPrereqOnly'
  ].map(id => [id, new Element()]));
  elements.get('catalogSnapshot').textContent = JSON.stringify(snapshot);
  elements.get('catalogMetadata').textContent = JSON.stringify({
    ...snapshot, identity_issues: []
  });
  const events = {};
  const printed = [];
  const downloads = [];
  const cards = [];
  const context = vm.createContext({
    document: {
      getElementById: id => elements.get(id),
      createElement: tag => {
        const element = new Element();
        if (tag === 'a') element.click = () => downloads.push(element.download);
        return element;
      },
      querySelectorAll: selector => selector.includes(':not(.hidden)') ? cards : []
    },
    window: {
      addEventListener: (event, handler) => { events[event] = handler; },
      print: () => {
        events.beforeprint();
        printed.push(elements.get('courseBooklet').children);
        events.afterprint();
      }
    },
    console: { error() {} },
    Blob,
    URL: {
      createObjectURL: blob => { downloads.push(blob); return 'blob:baseline'; },
      revokeObjectURL() {}
    },
    setTimeout: handler => handler(),
    Intl, Date
  });
  vm.runInContext(readFileSync(require.resolve('../filter.js'), 'utf8'), context);
  return { context, elements, events, printed, downloads, cards };
}

test('saving downloads a reloadable baseline with selected lecture and lab identities', async () => {
  const file = baselineFile();
  file.snapshot.records.push({ id: 'lab', variants: [{ TITLE: 'Biology Lab' }] });
  const { context, cards, downloads } = browserHarness(file.snapshot);
  cards.push({ dataset: { recordIds: JSON.stringify(['lecture', 'lab']) } });
  context.saveBookletBaseline();
  assert.equal(downloads[1], '2027-spring-booklet-baseline.json');
  const saved = validateBookletBaseline(JSON.parse(await downloads[0].text()));
  assert.deepEqual(saved.selected_record_ids, ['lecture', 'lab']);
  assert.equal(saved.selected_course_count, 1);
  assert.deepEqual(saved.snapshot, file.snapshot);
});

test('loading, printing, cleanup, and invalid-file recovery use the correct print mode', async () => {
  const file = baselineFile();
  const current = structuredClone(file.snapshot);
  current.records[0].variants[0].TITLE = 'Updated Biology';
  const { context, elements, events, printed } = browserHarness(current);
  const input = { files: [{ text: async () => JSON.stringify(file) }], value: 'baseline.json' };
  await context.loadBookletBaseline(input);
  assert.equal(input.value, '');
  assert.equal(elements.get('printCorrigendumButton').hidden, false);
  assert.match(elements.get('corrigendumStatus').textContent, /1 changed/);
  context.printCorrigendum();
  assert.match(printed[0][0].textContent, /Corrigendum/);
  const rows = printed[0][4].children[2].children;
  assert.equal(rows[1].children[1].textContent, 'Biology');
  assert.equal(rows[1].children[2].textContent, 'Updated Biology');
  assert.equal(elements.get('courseBooklet').children.length, 0);
  events.beforeprint();
  assert.match(elements.get('courseBooklet').children[0].textContent, /Course Booklet/);
  await context.loadBookletBaseline({ files: [{ text: async () => '{}' }], value: 'bad.json' });
  assert.equal(elements.get('printCorrigendumButton').hidden, true);
  assert.match(elements.get('corrigendumStatus').textContent, /Could not compare/);
  context.printCorrigendum();
  assert.equal(printed.length, 1);
});

test('a slower file read cannot replace a newer comparison', async () => {
  const file = baselineFile();
  const { context, elements } = browserHarness(file.snapshot);
  let finishFirst;
  const first = context.loadBookletBaseline({
    files: [{ text: () => new Promise(resolve => { finishFirst = resolve; }) }]
  });
  await context.loadBookletBaseline({
    files: [{ text: async () => '{}' }]
  });
  finishFirst(JSON.stringify(file));
  await first;
  assert.equal(elements.get('printCorrigendumButton').hidden, true);
  assert.match(elements.get('corrigendumStatus').textContent, /Could not compare/);
});
