const TRACKER_ID = '1IwfPam3I-o4cFu09R9Z6U86RK59LF96ZZjlEqArPAoM';
const LAB_ORIGIN = 'https://cyhendrix.github.io';
const LABS = {
  'Lesson 03: Drawing in Game Lab': {
    terms: ['Program','Bug','Debugging','rect()','ellipse()','fill()'],
    apply: [0,1,2,0], debug: [0,0,0], total: 15,
    orders: [[0,1,2,3],[0,1,2,3]]
  },
  'Lesson 01: Programming for a Purpose': {
    terms: ['User','Prototype','Define','Prepare','Try','Reflect'],
    apply: [1,0,0,0], debug: [1,0,0], total: 14
  },
  'Lesson 02: Plotting Shapes': {
    terms: ['Origin','x-coordinate','y-coordinate','Circle center','Rectangle corner','Drawing order'],
    apply: [0,1,2,1], debug: [0,2,1], total: 17,
    points: [[100,100],[250,150],[20,250],[330,70]]
  }
};

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({
    ok: true, service: 'Game Design Grade Submission', lessons: [1,2,3]
  })).setMimeType(ContentService.MimeType.JSON);
}

function gradeAttempt(p) {
  const config = LABS[p.lab], a = p.answers, checks = p.checks || {};
  if (!config) throw new Error('This lesson is not supported yet.');
  if (!a || !a.identify || !a.apply || !a.debug) throw new Error('Answers are missing.');
  let points = 0, buildPoints = 0;
  const earned = {identify: 0, apply: 0, debug: 0};
  config.terms.forEach(function(term, i) { if (a.identify[i] === term) { points++; earned.identify++; } });
  [ ['apply', config.apply], ['debug', config.debug] ].forEach(function(group) {
    group[1].forEach(function(answer, i) {
      if (a[group[0]][i] !== undefined && String(a[group[0]][i]) === String(answer)) { points++; earned[group[0]]++; }
    });
  });
  if (config.orders) {
    if (!a.build || typeof a.build !== 'object') throw new Error('Drawing code is missing.');
    config.orders.forEach(function(order, i) {
      const answer = a.build[i];
      if (!Array.isArray(answer) || answer.length !== 4 || new Set(answer).size !== 4 ||
          answer.some(function(v) { return ![0,1,2,3].includes(v); })) throw new Error('Invalid drawing code order.');
      if (answer.join('|') === order.join('|')) buildPoints++;
    });
  } else if (config.points) {
    if (!a.build || typeof a.build !== 'object') throw new Error('Grid answers are missing.');
    config.points.forEach(function(target, i) {
      const answer = a.build[i];
      if (answer && answer.x === target[0] && answer.y === target[1]) buildPoints++;
    });
  } else {
    const process = ['Define','Prepare','Try','Reflect'];
    if (!Array.isArray(a.order) || a.order.length !== 4 ||
        new Set(a.order).size !== 4 || a.order.some(function(v) { return !process.includes(v); }))
      throw new Error('Invalid process order.');
    if (a.order.join('|') === process.join('|')) buildPoints = 1;
  }
  const possible = {identify: config.terms.length, apply: config.apply.length,
    debug: config.debug.length, build: config.orders ? config.orders.length : config.points ? config.points.length : 1};
  earned.build = buildPoints;
  if (!Object.keys(possible).every(function(k) {
    return checks[k] === 2 || (checks[k] === 1 && earned[k] === possible[k]);
  })) throw new Error('Finish every activity before submitting.');
  points += buildPoints;
  return {points: points, total: config.total, score: Math.round(points / config.total * 100)};
}

function doPost(e) {
  let receipt = {ok: false, attemptId: '', message: 'Grade could not be recorded.'};
  try {
    const raw = e && e.parameter && e.parameter.payload;
    if (!raw || raw.length > 12000) throw new Error('Invalid submission.');
    const p = JSON.parse(raw);
    receipt.attemptId = String(p.attemptId || '');
    if (!/^[a-zA-Z0-9_-]{16,100}$/.test(receipt.attemptId)) throw new Error('Invalid attempt ID.');
    const name = String(p.name || '').trim().replace(/\s+/g, ' ');
    const period = String(p.period || '');
    if (name.length < 3 || name.length > 100 || /^[=+@-]/.test(name)) throw new Error('Check your student name.');
    if (!['1st Block','2nd Block','3rd Block','4th Block'].includes(period)) throw new Error('Choose a valid class period.');
    const grade = gradeAttempt(p);
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const sheet = SpreadsheetApp.openById(TRACKER_ID).getSheetByName('Form Responses 1');
      if (!sheet) throw new Error('The grade response sheet is missing.');
      const expected = ['Timestamp','Student Name','Class Period','Practice Lab','Score'];
      const headings = sheet.getRange(1,1,1,5).getDisplayValues()[0];
      if (headings.some(function(v,i) { return v.trim() !== expected[i]; }))
        throw new Error('The response sheet headings need to be checked.');
      const extra = sheet.getRange(1,6).getDisplayValue();
      if (extra && extra !== 'Attempt ID') throw new Error('Column F is already in use.');
      if (!extra) sheet.getRange(1,6).setValue('Attempt ID');
      const last = sheet.getLastRow();
      const found = last > 1 ? sheet.getRange(2,6,last-1,1)
        .createTextFinder(receipt.attemptId).matchEntireCell(true).findNext() : null;
      let rowNumber;
      if (found) {
        rowNumber = found.getRow();
      } else {
        rowNumber = last + 1;
        sheet.getRange(rowNumber,1,1,6).setValues([
          [new Date(),name,period,p.lab,grade.score,receipt.attemptId]
        ]);
        SpreadsheetApp.flush();
      }
      const stored = sheet.getRange(rowNumber,2,1,5).getValues()[0];
      if (stored[0] !== name || stored[1] !== period || stored[2] !== p.lab ||
          Number(stored[3]) !== grade.score || stored[4] !== receipt.attemptId)
        throw new Error('Unable to confirm this attempt. Check the name, period, and saved answers.');
      receipt = {ok: true, attemptId: receipt.attemptId, score: grade.score,
        points: grade.points, total: grade.total, message: 'Your grade was recorded successfully.'};
    } finally {
      lock.releaseLock();
    }
  } catch (error) {
    receipt.message = error.message || 'Please try submitting again.';
  }
  const safe = JSON.stringify(receipt).replace(/</g, '\\u003c');
  return HtmlService.createHtmlOutput(
    '<!doctype html><meta charset="utf-8"><p>Grade submission processed. Return to your practice lab.</p>' +
    '<script>window.top.postMessage(' + safe + ', ' + JSON.stringify(LAB_ORIGIN) + ');</script>'
  ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
