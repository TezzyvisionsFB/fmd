const SPREADSHEET_ID = "1-m3bKPF7af6vMKhYtaPc5kAnOjWXt9WcyFqe3-6nDjQ";
const TZ = "America/New_York";
const FLAT_SPIN_ENTRY_PRICE = 7;
const FLAT_SPIN_MAX = 100;
const HOUSE_PERCENT_DEFAULT = 15;
const SPADES_MAX_TEAMS = 16;

// -- Stripe -------------------------------------------------------------------
var STRIPE_SECRET_KEY = 'sk_test_51TO0EtFqQTLWDXzGC3bLnPmWbKuJifCT80UWWtUxv2u5TAJx3fFVX3BQ0b9UfoN8hzdmrvTgzHL6L1ZVyI4RhXHb00JiwIJFae';

// -- OneSignal Push -----------------------------------------------------------
var ONESIGNAL_APP_ID  = 'd6c9aeef-7069-4efd-8a41-af1eb0f31f12';
var ONESIGNAL_API_KEY = 'os_v2_app_23e2533qnfhp3csbv4plb4y7ck4zs3wmw5huy442y4fvu3fhyyhjbql5xb2tlgp6rrweqk5cbm35o3souj54o6mxctxxok6aj4tybuy'; // get from OneSignal dashboard > Settings > Keys & IDs

// -- Admin -------------------------------------------------------------------
var ADMIN_CASHAPPS = ['aheard193'];

function getSheet_(name) {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    Logger.log('Created missing sheet: ' + name);
  }
  return sheet;
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function textOutput_(text) {
  return ContentService
    .createTextOutput(String(text))
    .setMimeType(ContentService.MimeType.TEXT);
}

function todayEastern_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
}

function nowEasternIso_() {
  return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss");
}

function cleanString_(value) {
  return String(value == null ? '' : value).trim();
}

function cleanDate_(value) {
  if (!value) return '';
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, TZ, 'yyyy-MM-dd');
  }
  if (typeof value === 'number') {
    var d = new Date(Math.round((value - 25569) * 86400 * 1000));
    return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
  }
  var s = String(value).trim();
  if (s.indexOf('/') !== -1) {
    var parts = s.split('/');
    if (parts.length === 3) {
      var month = parts[0].padStart(2, '0');
      var day = parts[1].padStart(2, '0');
      var year = parts[2].length === 2 ? '20' + parts[2] : parts[2];
      return year + '-' + month + '-' + day;
    }
  }
  return s.substring(0, 10);
}

function parseMoney_(value) {
  const cleaned = cleanString_(value).replace(/[^0-9.\-]/g, '');
  const num = parseFloat(cleaned);
  return Number.isFinite(num) ? num : 0;
}

function getDataRows_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return values.filter(function (row) {
    return row.some(function (cell) { return cell !== ''; });
  });
}

// -- Flat Spin helpers --------------------------------------------------------

function getSpinAdminForDate_(drawDate) {
  const rows = getDataRows_(getSheet_('SpinAdmin'));
  const match = rows.find(function (row) {
    return cleanDate_(row[0]) === drawDate;
  });
  return {
    drawDate: drawDate,
    mode: cleanString_(match && match[1] ? match[1] : 'random').toLowerCase() || 'random',
    forcedNum1: Number(match && match[2] ? match[2] : ''),
    forcedNum2: Number(match && match[3] ? match[3] : ''),
    housePercent: Number(match && match[4] !== '' ? match[4] : HOUSE_PERCENT_DEFAULT),
    status: cleanString_(match && match[5] ? match[5] : 'open').toLowerCase() || 'open'
  };
}

function getFlatSpinEntriesForDate_(drawDate) {
  const rows = getDataRows_(getSheet_('FlatSpinEntries'));
  return rows
    .filter(function (row) { return cleanDate_(row[4]) === drawDate; })
    .map(function (row) {
      return {
        timestamp: row[0],
        name: cleanString_(row[1]),
        cashapp: cleanString_(row[2]),
        assignedNumber: Number(row[3]),
        drawDate: cleanString_(row[4]),
        paymentStatus: cleanString_(row[5] || 'Pending'),
        unit: cleanString_(row[6] || '')
      };
    })
    .filter(function (entry) { return !!entry.name; });
}

function getUsedFlatSpinNumbers_(drawDate) {
  return getFlatSpinEntriesForDate_(drawDate)
    .map(function (entry) { return Number(entry.assignedNumber); })
    .filter(function (n) { return Number.isFinite(n) && n >= 1 && n <= FLAT_SPIN_MAX; });
}

function getRandomAvailableNumber_(usedNumbers, max) {
  const used = new Set(usedNumbers.map(Number));
  const available = [];
  for (var i = 1; i <= max; i++) {
    if (!used.has(i)) available.push(i);
  }
  if (!available.length) return null;
  return available[Math.floor(Math.random() * available.length)];
}

function getLatestSpinResultForDate_(drawDate) {
  const rows = getDataRows_(getSheet_('FlatSpinResults'));
  const todayRows = rows.filter(function (row) { return cleanDate_(row[0]) === drawDate; });
  if (!todayRows.length) return null;
  const row = todayRows[todayRows.length - 1];
  return {
    drawDate: cleanString_(row[0]),
    winningNumber1: Number(row[1]),
    winningNumber2: Number(row[3]),
    payout1: Number(row[5] || 0),
    payout2: Number(row[6] || 0),
    unit1: cleanString_(row[11] || ''),
    unit2: cleanString_(row[12] || ''),
    totalPot: Number(row[7] || 0),
    modeUsed: cleanString_(row[9] || 'random'),
    createdAt: row[10]
  };
}

function calculateFlatSpinPayouts_(entries, rolloverAmount) {
  var totalEntries = entries.length;
  var paidEntries = entries.filter(function(e) { return e.paymentStatus.toLowerCase() === 'paid'; });
  var collectedTonight = paidEntries.length * FLAT_SPIN_ENTRY_PRICE;
  var rollover = Number(rolloverAmount) || 0;
  var totalPot = collectedTonight + rollover;
  var ADMIN_RATE = 0.15;
  var adminFee = Math.round(totalPot * ADMIN_RATE * 100) / 100;
  var playerPool = totalPot - adminFee;
  var payout1 = Math.round(playerPool * 0.80 * 100) / 100;
  var payout2 = Math.round(playerPool * 0.20 * 100) / 100;
  return {
    totalEntries: totalEntries,
    paidEntries: paidEntries.length,
    collectedTonight: collectedTonight,
    rolloverAmount: rollover,
    totalPot: totalPot,
    adminFee: adminFee,
    playerPool: playerPool,
    firstPayout: payout1,
    secondPayout: payout2,
    maxEntries: FLAT_SPIN_MAX,
    entryPrice: FLAT_SPIN_ENTRY_PRICE
  };
}

function getFlatSpinData_() {
  const drawDate = todayEastern_();
  const admin = getSpinAdminForDate_(drawDate);
  const entries = getFlatSpinEntriesForDate_(drawDate).sort(function (a, b) {
    return Number(a.assignedNumber) - Number(b.assignedNumber);
  });
  const rolloverInfo = getActiveRollover_();
  const payouts = calculateFlatSpinPayouts_(entries, rolloverInfo.totalRollover);
  const latestResult = getLatestSpinResultForDate_(drawDate);
  const historyRows = getDataRows_(getSheet_('FlatSpinHistory'));
  const history = historyRows.slice(-30).reverse().map(function (row) {
    return {
      date: cleanDate_(row[0]),
      place: cleanString_(row[1]),
      winningNumber: Number(row[2]),
      payout: Number(row[4] || 0),
      unit: cleanString_(row[5] || '')
    };
  });
  var paidEntries = entries.filter(function(e) { return e.paymentStatus.toLowerCase() === 'paid'; });
  var pendingEntries = entries.filter(function(e) { return e.paymentStatus.toLowerCase() !== 'paid'; });
  return jsonOutput_({
    ok: true,
    drawDate: drawDate,
    status: admin.status,
    entries: entries,
    totalEntries: payouts.totalEntries,
    paidCount: paidEntries.length,
    pendingCount: pendingEntries.length,
    collectedTonight: payouts.collectedTonight,
    rolloverAmount: payouts.rolloverAmount,
    rolloverNights: rolloverInfo.nights,
    hasRollover: rolloverInfo.hasRollover,
    totalPot: payouts.totalPot,
    adminFee: payouts.adminFee,
    playerPool: payouts.playerPool,
    firstPayout: payouts.firstPayout,
    secondPayout: payouts.secondPayout,
    latestResult: latestResult,
    result: latestResult,
    history: history
  });
}

function submitFlatSpinEntry_(p) {
  const drawDate = todayEastern_();
  const admin = getSpinAdminForDate_(drawDate);
  if (admin.status === 'locked') {
    return jsonOutput_({ ok: false, message: "Tonight's spin is locked." });
  }
  const name = cleanString_(p.name);
  const cashapp = cleanString_(p.cashapp);
  if (!name || !cashapp) {
    return jsonOutput_({ ok: false, message: 'Name and Cash App username are required.' });
  }
  const usedNumbers = getUsedFlatSpinNumbers_(drawDate);
  const assignedNumber = getRandomAvailableNumber_(usedNumbers, FLAT_SPIN_MAX);
  if (!assignedNumber) {
    return jsonOutput_({ ok: false, message: 'All 100 numbers are filled for tonight.' });
  }
  var unit = cleanString_(p.unit || '');
  var paymentStatus = p.paymentIntentId ? 'Paid' : 'Pending';
  if (p.paymentIntentId) markStripePaymentConfirmed_(p.paymentIntentId);
  getSheet_('FlatSpinEntries').appendRow([
    Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'),
    name, cashapp, assignedNumber, drawDate, paymentStatus, unit
  ]);
  return jsonOutput_({
    ok: true,
    assignedNumber: assignedNumber,
    drawDate: drawDate,
    message: 'Entry submitted! Your number is #' + assignedNumber + '. Good luck!'
  });
}

function drawFlatSpinWinners_(paidOnly) {
  const drawDate = todayEastern_();
  const admin = getSpinAdminForDate_(drawDate);
  const existing = getLatestSpinResultForDate_(drawDate);
  if (existing && admin.status === 'locked') {
    return jsonOutput_({ ok: true, reused: true, result: existing });
  }
  var allEntries = getFlatSpinEntriesForDate_(drawDate);
  if (!allEntries.length) {
    return jsonOutput_({ ok: false, message: 'No entries for today.' });
  }
  var entries = paidOnly
    ? allEntries.filter(function(e) { return e.paymentStatus.toLowerCase() === 'paid'; })
    : allEntries;
  if (!entries.length) {
    return jsonOutput_({ ok: false, message: 'No paid entries for today.' });
  }
  const assignedNumbers = entries
    .map(function (e) { return Number(e.assignedNumber); })
    .filter(function (n) { return Number.isFinite(n); });

  let num1, num2;
  if (admin.mode === 'fixed') {
    num1 = Number(admin.forcedNum1);
    num2 = Number(admin.forcedNum2);
    if (!assignedNumbers.includes(num1) || !assignedNumbers.includes(num2) || num1 === num2) {
      return jsonOutput_({ ok: false, message: 'Fixed winning numbers must match two different entered numbers.' });
    }
  } else {
    num1 = assignedNumbers[Math.floor(Math.random() * assignedNumbers.length)];
    const remaining = assignedNumbers.filter(function (n) { return n !== num1; });
    if (!remaining.length) {
      return jsonOutput_({ ok: false, message: 'Need at least 2 entries to draw 2 winners.' });
    }
    num2 = remaining[Math.floor(Math.random() * remaining.length)];
  }

  const winner1 = entries.find(function (e) { return Number(e.assignedNumber) === num1; });
  const winner2 = entries.find(function (e) { return Number(e.assignedNumber) === num2; });

  var rolloverInfo = getActiveRollover_();
  var payouts = calculateFlatSpinPayouts_(entries, rolloverInfo.totalRollover);
  var payout1 = payouts.firstPayout;
  var payout2 = payouts.secondPayout;
  var totalPot = payouts.totalPot;
  var adminFee = payouts.adminFee;

  if (rolloverInfo.hasRollover) markRolloverPaid_(drawDate);

  getSheet_('FlatSpinResults').appendRow([
    drawDate,
    num1, winner1 ? winner1.name : 'No Winner',
    num2, winner2 ? winner2.name : 'No Winner',
    payout1, payout2, totalPot, adminFee, admin.mode, new Date(),
    winner1 ? winner1.unit : '',
    winner2 ? winner2.unit : ''
  ]);
  getSheet_('FlatSpinHistory').appendRow([drawDate, '1st', num1, winner1 ? winner1.name : 'No Winner', payout1, new Date(), winner1 ? winner1.unit : '']);
  getSheet_('FlatSpinHistory').appendRow([drawDate, '2nd', num2, winner2 ? winner2.name : 'No Winner', payout2, new Date(), winner2 ? winner2.unit : '']);

  // Notify all members
  notifyDrawResult_(num1, '$' + payout1, num2, '$' + payout2);

  return jsonOutput_({
    ok: true,
    result: {
      drawDate: drawDate,
      winningNumber1: num1,
      unit1: winner1 ? winner1.unit : '',
      winningNumber2: num2,
      unit2: winner2 ? winner2.unit : '',
      payout1: payout1, payout2: payout2,
      totalPot: totalPot, adminFee: adminFee,
      rolloverAmount: rolloverInfo.totalRollover,
      modeUsed: admin.mode, createdAt: nowEasternIso_()
    }
  });
}

// -- Spades helpers -----------------------------------------------------------

function getSpadesTeams_() {
  const rows = getDataRows_(getSheet_('SpadesSignup'));
  const teams = rows.map(function (row) {
    return {
      timestamp: row[0],
      teamName: cleanString_(row[1]),
      unit: cleanString_(row[2]),
      player1Name: cleanString_(row[3]),
      player1Building: cleanString_(row[4]),
      player2Name: cleanString_(row[5]),
      player2Building: cleanString_(row[6]),
      cashapp: cleanString_(row[7]),
      buyin: cleanString_(row[8]),
      notes: cleanString_(row[9])
    };
  }).filter(function (t) { return !!t.teamName; });
  const bracket = loadBracket_();
  return jsonOutput_({
    ok: true,
    teams: teams,
    totalTeams: teams.length,
    teamsRemaining: Math.max(SPADES_MAX_TEAMS - teams.length, 0),
    maxTeams: SPADES_MAX_TEAMS,
    bracket: bracket
  });
}

function submitSpadesSignup_(p) {
  const sheet = getSheet_('SpadesSignup');
  const existing = getDataRows_(sheet).map(function (row) { return cleanString_(row[1]).toLowerCase(); });
  if (existing.length >= SPADES_MAX_TEAMS) {
    return jsonOutput_({ ok: false, message: 'The Spades tournament is full.' });
  }
  const teamName = cleanString_(p.teamname);
  if (!teamName) return jsonOutput_({ ok: false, message: 'Team name is required.' });
  if (existing.includes(teamName.toLowerCase())) {
    return jsonOutput_({ ok: false, message: 'That team name is already taken.' });
  }
  sheet.appendRow([
    new Date(), teamName,
    cleanString_(p.unit || ''),
    cleanString_(p.player1name),
    cleanString_(p.player1building),
    cleanString_(p.player2name),
    cleanString_(p.player2building),
    cleanString_(p.cashapp),
    cleanString_(p.buyin || '2 Flats Per Team'),
    cleanString_(p.notes)
  ]);
  return jsonOutput_({ ok: true, message: 'Team submitted successfully.' });
}

function saveBracket_(bracketObj) {
  const sheet = getSheet_('SpadesBracket');
  sheet.getRange('A1').setValue(JSON.stringify(bracketObj));
}

function loadBracket_() {
  try {
    const sheet = getSheet_('SpadesBracket');
    const val = sheet.getRange('A1').getValue();
    if (!val) return null;
    return JSON.parse(val);
  } catch (e) { return null; }
}

function generateBracket_() {
  const rows = getDataRows_(getSheet_('SpadesSignup'));
  const teams = rows.map(function (row) { return cleanString_(row[1]); }).filter(function (name) { return !!name; });
  if (teams.length < 2) {
    return jsonOutput_({ ok: false, error: 'Need at least 2 teams to generate a bracket.' });
  }
  for (var i = teams.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var tmp = teams[i]; teams[i] = teams[j]; teams[j] = tmp;
  }
  while (teams.length < 16) { teams.push('BYE'); }
  var r16 = [];
  for (var k = 0; k < 16; k += 2) {
    r16.push({ teamA: teams[k], teamB: teams[k + 1], winner: null });
  }
  var bracket = {
    generatedAt: nowEasternIso_(),
    r16: r16,
    qf: [
      { teamA: null, teamB: null, winner: null },
      { teamA: null, teamB: null, winner: null },
      { teamA: null, teamB: null, winner: null },
      { teamA: null, teamB: null, winner: null }
    ],
    sf: [
      { teamA: null, teamB: null, winner: null },
      { teamA: null, teamB: null, winner: null }
    ],
    final: { teamA: null, teamB: null, winner: null }
  };
  saveBracket_(bracket);
  return jsonOutput_({ ok: true, bracket: bracket });
}

var PROP_QUESTIONS = [
  { id: 'unit_winner', label: 'What unit will the tournament winner come from?', type: 'choice', options: ['Unit 1', 'Unit 2', 'Unit 3', 'Unit 4'] },
  { id: 'ten_books', label: 'Will any team score 10 books in a single game?', type: 'yesno' },
  { id: 'final_set', label: 'Will any team get set (black eye) in the championship?', type: 'yesno' },
  { id: 'game5', label: 'Will the championship go to game 5?', type: 'yesno' },
  { id: 'most_sets_unit', label: 'Which unit will have the most total sets?', type: 'choice', options: ['Unit 1', 'Unit 2', 'Unit 3', 'Unit 4'] },
  { id: 'sweep', label: 'Will any team win a match 3-0 (sweep)?', type: 'yesno' }
];

function getPropBets_() {
  var rows = getDataRows_(getSheet_('SpadesSideBets'));
  var bets = rows.map(function(row) {
    return {
      timestamp: row[0], bettorName: cleanString_(row[1]), cashapp: cleanString_(row[2]),
      propId: cleanString_(row[3]), propLabel: cleanString_(row[4]),
      answer: cleanString_(row[5]), betAmount: parseMoney_(row[6]),
      paymentStatus: cleanString_(row[7] || 'Pending')
    };
  }).filter(function(b) { return !!b.propId; });
  var summary = {};
  PROP_QUESTIONS.forEach(function(q) {
    summary[q.id] = { id: q.id, label: q.label, type: q.type, options: q.options || null, answers: {}, totalBets: 0, totalAmount: 0 };
  });
  bets.forEach(function(bet) {
    if (!summary[bet.propId]) return;
    var s = summary[bet.propId];
    s.totalBets += 1; s.totalAmount += bet.betAmount;
    if (!s.answers[bet.answer]) s.answers[bet.answer] = { answer: bet.answer, count: 0, totalAmount: 0 };
    s.answers[bet.answer].count += 1; s.answers[bet.answer].totalAmount += bet.betAmount;
  });
  var props = PROP_QUESTIONS.map(function(q) {
    var s = summary[q.id];
    var answersArr = Object.values(s.answers).sort(function(a,b) { return b.totalAmount - a.totalAmount; });
    return { id: q.id, label: q.label, type: q.type, options: q.options || null, totalBets: s.totalBets, totalAmount: s.totalAmount, answers: answersArr };
  });
  return jsonOutput_({ ok: true, props: props, totalBets: bets.length });
}

function submitPropBet_(p) {
  var bettorName = cleanString_(p.bettorname);
  var cashapp = cleanString_(p.cashapp);
  var propId = cleanString_(p.propid);
  var answer = cleanString_(p.answer);
  var betAmount = cleanString_(p.betamount);
  if (!bettorName || !cashapp || !propId || !answer || !betAmount) {
    return jsonOutput_({ ok: false, message: 'All prop bet fields are required.' });
  }
  var prop = PROP_QUESTIONS.filter(function(q) { return q.id === propId; })[0];
  if (!prop) return jsonOutput_({ ok: false, message: 'Invalid prop question.' });
  getSheet_('SpadesSideBets').appendRow([
    Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'),
    bettorName, cashapp, propId, prop.label, answer, betAmount, 'Pending'
  ]);
  return jsonOutput_({ ok: true, message: 'Prop bet submitted! Send payment via Cash App to lock it in.' });
}

function getSpadesMatches_() {
  const rows = getDataRows_(getSheet_('SpadesMatches'));
  const matches = rows.map(function (row) {
    return {
      matchTitle: cleanString_(row[0]), teamA: cleanString_(row[1]), teamB: cleanString_(row[2]),
      teamAWins: Number(row[3] || 0), teamBWins: Number(row[4] || 0),
      teamASets: Number(row[5] || 0), teamBSets: Number(row[6] || 0),
      status: cleanString_(row[7] || 'Posted')
    };
  }).filter(function (m) { return m.matchTitle && m.teamA && m.teamB; });
  return jsonOutput_({ ok: true, matches: matches });
}

function getZebraWinners_() {
  const rows = getDataRows_(getSheet_('ZebraWinners'));
  const winners = rows.slice(-20).reverse().map(function (row) {
    return { date: cleanDate_(row[0]), winner: cleanString_(row[1]), prize: cleanString_(row[2]), result: cleanString_(row[3]) };
  });
  return jsonOutput_({ ok: true, winners: winners });
}

// -- Quick Draw ---------------------------------------------------------------

function getQuickDrawEntries_() {
  var drawDate = todayEastern_();
  var rows = getDataRows_(getSheet_('QuickDrawEntries'));
  return rows
    .filter(function(row) { return cleanDate_(row[5]) === drawDate; })
    .map(function(row) {
      return { timestamp: row[0], name: cleanString_(row[1]), cashapp: cleanString_(row[2]), unit: cleanString_(row[3]), entryNumber: Number(row[4]), drawDate: cleanString_(row[5]) };
    })
    .filter(function(e) { return !!e.name; });
}

function getLatestQuickDrawResult_() {
  var drawDate = todayEastern_();
  try {
    var rows = getDataRows_(getSheet_('QuickDrawResults'));
    var todayRows = rows.filter(function(row) { return cleanDate_(row[0]) === drawDate; });
    if (!todayRows.length) return null;
    var row = todayRows[todayRows.length - 1];
    return {
      drawDate: cleanString_(row[0]),
      entryNumber1: Number(row[1]), unit1: cleanString_(row[3]), prize1: cleanString_(row[4]),
      entryNumber2: Number(row[5]), unit2: cleanString_(row[7]), prize2: cleanString_(row[8])
    };
  } catch(e) { return null; }
}

function getQuickDrawData_() {
  var drawDate = todayEastern_();
  var entries = getQuickDrawEntries_();
  var result = getLatestQuickDrawResult_();
  return jsonOutput_({ ok: true, drawDate: drawDate, entries: entries, totalEntries: entries.length, result: result });
}

function submitQuickDrawEntry_(p) {
  var drawDate = todayEastern_();
  var name = cleanString_(p.name);
  var cashapp = cleanString_(p.cashapp || '');
  var unit = cleanString_(p.unit || '');
  if (!name || !cashapp) return jsonOutput_({ ok: false, message: 'Name and Cash App username are required.' });
  var existing = getQuickDrawEntries_();
  var already = existing.filter(function(e) { return e.name.toLowerCase() === name.toLowerCase(); });
  if (already.length) return jsonOutput_({ ok: false, message: "You already have an entry in tonights Quick Draw." });
  var usedNumbers = existing.map(function(e) { return Number(e.entryNumber); });
  var assignedNumber = getRandomAvailableNumber_(usedNumbers, FLAT_SPIN_MAX);
  if (!assignedNumber) return jsonOutput_({ ok: false, message: 'All 100 Quick Draw spots are filled.' });
  if (p.paymentIntentId) markStripePaymentConfirmed_(p.paymentIntentId);
  getSheet_('QuickDrawEntries').appendRow([
    Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'),
    name, cashapp, unit, assignedNumber, drawDate
  ]);
  return jsonOutput_({ ok: true, entryNumber: assignedNumber, message: "Entry submitted! Your number is #" + assignedNumber + ". Good luck!" });
}

function drawQuickDrawWinner_() {
  var drawDate = todayEastern_();
  var existing = getLatestQuickDrawResult_();
  if (existing) return jsonOutput_({ ok: true, reused: true, result: existing });
  var entries = getQuickDrawEntries_();
  if (entries.length < 2) return jsonOutput_({ ok: false, message: 'Need at least 2 entries for Quick Draw.' });
  var idx1 = Math.floor(Math.random() * entries.length);
  var winner1 = entries[idx1];
  var remaining = entries.filter(function(e) { return e.entryNumber !== winner1.entryNumber; });
  var winner2 = remaining[Math.floor(Math.random() * remaining.length)];
  getSheet_('QuickDrawResults').appendRow([
    drawDate,
    winner1.entryNumber, winner1.name, winner1.unit, '$150',
    winner2.entryNumber, winner2.name, winner2.unit, '$50',
    Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss")
  ]);
  return jsonOutput_({
    ok: true,
    result: {
      drawDate: drawDate,
      entryNumber1: winner1.entryNumber, unit1: winner1.unit,
      entryNumber2: winner2.entryNumber, unit2: winner2.unit,
      prize1: '$150', prize2: '$50'
    }
  });
}

// -- Generic Raffle -----------------------------------------------------------

function getRaffleEntries_(sheetName, drawDate) {
  var rows = getDataRows_(getSheet_(sheetName));
  return rows
    .filter(function(row) { return cleanDate_(row[4]) === drawDate; })
    .map(function(row) {
      return { name: cleanString_(row[1]), cashapp: cleanString_(row[2]), assignedNumber: Number(row[3]), drawDate: cleanString_(row[4]), paymentStatus: cleanString_(row[5] || 'Pending'), unit: cleanString_(row[6] || '') };
    })
    .filter(function(e) { return !!e.name; });
}

function getUsedRaffleNumbers_(sheetName, drawDate) {
  return getRaffleEntries_(sheetName, drawDate)
    .map(function(e) { return Number(e.assignedNumber); })
    .filter(function(n) { return Number.isFinite(n) && n >= 1 && n <= FLAT_SPIN_MAX; });
}

function getLatestRaffleResult_(sheetName, drawDate) {
  try {
    var rows = getDataRows_(getSheet_(sheetName));
    var todayRows = rows.filter(function(row) { return cleanDate_(row[0]) === drawDate; });
    if (!todayRows.length) return null;
    var row = todayRows[todayRows.length - 1];
    return { drawDate: cleanString_(row[0]), winningNumber: Number(row[1]), unit: cleanString_(row[3] || ''), prize: cleanString_(row[4] || '') };
  } catch(e) { return null; }
}

function getRaffleData_(entrySheet, resultSheet, raffleName) {
  var drawDate = todayEastern_();
  var entries = getRaffleEntries_(entrySheet, drawDate);
  var result = getLatestRaffleResult_(resultSheet, drawDate);
  var paid = entries.filter(function(e) { return e.paymentStatus.toLowerCase() === 'paid'; });
  var displayMax = raffleName ? getRaffleDisplayMax_(raffleName) : null;
  return jsonOutput_({ ok: true, drawDate: drawDate, entries: entries, totalEntries: entries.length, paidCount: paid.length, displayMax: displayMax, result: result });
}

function submitRaffleEntry_(p, entrySheet) {
  var drawDate = todayEastern_();
  var name = cleanString_(p.name);
  var cashapp = cleanString_(p.cashapp);
  var unit = cleanString_(p.unit || '');
  if (!name || !cashapp) return jsonOutput_({ ok: false, message: 'Name and Cash App are required.' });
  var usedNumbers = getUsedRaffleNumbers_(entrySheet, drawDate);
  if (usedNumbers.length >= FLAT_SPIN_MAX) return jsonOutput_({ ok: false, message: 'All 100 spots are filled.' });
  var assignedNumber = getRandomAvailableNumber_(usedNumbers, FLAT_SPIN_MAX);
  var paymentStatus = p.paymentIntentId ? 'Paid' : 'Pending';
  if (p.paymentIntentId) markStripePaymentConfirmed_(p.paymentIntentId);
  getSheet_(entrySheet).appendRow([
    Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'),
    name, cashapp, assignedNumber, drawDate, paymentStatus, unit
  ]);
  return jsonOutput_({ ok: true, assignedNumber: assignedNumber, message: 'Entry submitted! You are #' + assignedNumber + '. Good luck!' });
}

function drawRaffleWinner_(entrySheet, resultSheet, prize, paidOnly, numWinners) {
  var drawDate = todayEastern_();
  numWinners = Number(numWinners) || 1;
  if (numWinners < 1) numWinners = 1;
  if (numWinners > 50) numWinners = 50;
  var allEntries = getRaffleEntries_(entrySheet, drawDate);
  var pool = paidOnly ? allEntries.filter(function(e) { return e.paymentStatus.toLowerCase() === 'paid'; }) : allEntries;
  if (!pool.length) return jsonOutput_({ ok: false, message: 'No entries found for today.' });
  if (pool.length < numWinners) {
    return jsonOutput_({ ok: false, message: 'Not enough entries to draw ' + numWinners + ' winners. Only ' + pool.length + ' available.' });
  }
  var shuffled = pool.slice();
  for (var i = shuffled.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var tmp = shuffled[i]; shuffled[i] = shuffled[j]; shuffled[j] = tmp;
  }
  var winners = shuffled.slice(0, numWinners);
  var sheet = getSheet_(resultSheet);
  var timestamp = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss");
  winners.forEach(function(w, idx) {
    sheet.appendRow([drawDate, w.assignedNumber, w.name, w.unit, prize, idx + 1, numWinners, timestamp]);
  });
  return jsonOutput_({
    ok: true,
    result: {
      drawDate: drawDate, numWinners: numWinners,
      winners: winners.map(function(w, idx) {
        return { place: idx + 1, assignedNumber: w.assignedNumber, unit: w.unit, prize: prize };
      })
    }
  });
}

function getLatestCommissaryResult_(drawDate) {
  try {
    var rows = getDataRows_(getSheet_('CommissaryRaffleResults'));
    var todayRows = rows.filter(function(row) { return cleanDate_(row[0]) === drawDate; });
    if (!todayRows.length) return null;
    var row = todayRows[todayRows.length - 1];
    return { winningNumber1: Number(row[1]), unit1: cleanString_(row[3]), winningNumber2: Number(row[5]), unit2: cleanString_(row[7]) };
  } catch(e) { return null; }
}

function drawCommissaryWinners_(paidOnly) {
  var drawDate = todayEastern_();
  var existing = getLatestCommissaryResult_(drawDate);
  if (existing) return jsonOutput_({ ok: true, reused: true, result: existing });
  var allEntries = getRaffleEntries_('CommissaryRaffle', drawDate);
  var entries = paidOnly ? allEntries.filter(function(e) { return e.paymentStatus.toLowerCase() === 'paid'; }) : allEntries;
  if (entries.length < 2) return jsonOutput_({ ok: false, message: 'Need at least 2 entries to draw 2 winners.' });
  var idx1 = Math.floor(Math.random() * entries.length);
  var winner1 = entries[idx1];
  var remaining = entries.filter(function(e) { return e.assignedNumber !== winner1.assignedNumber; });
  var winner2 = remaining[Math.floor(Math.random() * remaining.length)];
  getSheet_('CommissaryRaffleResults').appendRow([
    drawDate, winner1.assignedNumber, winner1.name, winner1.unit, '$150',
    winner2.assignedNumber, winner2.name, winner2.unit, '$50',
    Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss")
  ]);
  return jsonOutput_({ ok: true, result: { drawDate: drawDate, winningNumber1: winner1.assignedNumber, unit1: winner1.unit, winningNumber2: winner2.assignedNumber, unit2: winner2.unit } });
}

function getCommissaryRaffleData_() {
  var drawDate = todayEastern_();
  var entries = getRaffleEntries_('CommissaryRaffle', drawDate);
  var result = getLatestCommissaryResult_(drawDate);
  var paid = entries.filter(function(e) { return e.paymentStatus.toLowerCase() === 'paid'; });
  var displayMax = getRaffleDisplayMax_('commissary');
  return jsonOutput_({ ok: true, drawDate: drawDate, entries: entries, totalEntries: entries.length, paidCount: paid.length, displayMax: displayMax, result: result });
}

function submitCommissaryEntry_(p) {
  var drawDate = todayEastern_();
  var name = cleanString_(p.name);
  var cashapp = cleanString_(p.cashapp);
  var unit = cleanString_(p.unit || '');
  if (!name || !cashapp) return jsonOutput_({ ok: false, message: 'Name and Cash App are required.' });
  var existing = getRaffleEntries_('CommissaryRaffle', drawDate);
  var assignedNumber = existing.length + 1;
  var paymentStatus = p.paymentIntentId ? 'Paid' : 'Pending';
  if (p.paymentIntentId) markStripePaymentConfirmed_(p.paymentIntentId);
  getSheet_('CommissaryRaffle').appendRow([
    Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'),
    name, cashapp, assignedNumber, drawDate, paymentStatus, unit
  ]);
  return jsonOutput_({ ok: true, assignedNumber: assignedNumber, message: 'Entry submitted! You are entry #' + assignedNumber + '. Good luck!' });
}

function getRaffleDisplayMax_(raffleName) {
  try {
    var rows = getDataRows_(getSheet_('RaffleSettings'));
    var match = rows.find(function(row) { return cleanString_(row[0]).toLowerCase() === raffleName.toLowerCase(); });
    if (!match) return null;
    var val = Number(match[1]);
    return Number.isFinite(val) && val > 0 ? val : null;
  } catch(e) { return null; }
}

// -- Rollover -----------------------------------------------------------------

function getActiveRollover_() {
  try {
    var rows = getDataRows_(getSheet_('SpinRollover'));
    var active = rows.filter(function(row) { return cleanString_(row[2]).toLowerCase() === 'active'; });
    if (!active.length) return { hasRollover: false, totalRollover: 0, nights: 0 };
    var total = active.reduce(function(sum, row) { return sum + Number(row[1] || 0); }, 0);
    return { hasRollover: true, totalRollover: total, nights: active.length };
  } catch(e) { return { hasRollover: false, totalRollover: 0, nights: 0 }; }
}

function triggerRollover_() {
  var drawDate = todayEastern_();
  var entries = getFlatSpinEntriesForDate_(drawDate);
  var paidEntries = entries.filter(function(e) { return e.paymentStatus.toLowerCase() === 'paid'; });
  var tonightPot = paidEntries.length * FLAT_SPIN_ENTRY_PRICE;
  if (tonightPot <= 0) {
    return jsonOutput_({ ok: false, message: 'No paid entries tonight to roll over.' });
  }
  try {
    var existing = getDataRows_(getSheet_('SpinRollover'));
    var alreadyRolled = existing.filter(function(row) { return cleanDate_(row[0]) === drawDate; });
    if (alreadyRolled.length) {
      return jsonOutput_({ ok: false, message: "Tonight's pot has already been rolled over." });
    }
  } catch(e) {}
  getSheet_('SpinRollover').appendRow([drawDate, tonightPot, 'active', Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss")]);
  var rollover = getActiveRollover_();
  return jsonOutput_({ ok: true, message: "Tonight's pot of $" + tonightPot + " has been rolled over.", totalRollover: rollover.totalRollover, nights: rollover.nights });
}

function markRolloverPaid_(drawDate) {
  try {
    var sheet = getSheet_('SpinRollover');
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (cleanString_(rows[i][2]).toLowerCase() === 'active') {
        sheet.getRange(i + 1, 3).setValue('paid');
      }
    }
  } catch(e) { Logger.log('Rollover mark paid error: ' + e); }
}

// -- Members ------------------------------------------------------------------

function isAdmin_(cashapp) {
  var ca = cleanString_(cashapp).toLowerCase().replace('$','');
  for (var i = 0; i < ADMIN_CASHAPPS.length; i++) {
    if (ADMIN_CASHAPPS[i].toLowerCase() === ca) return true;
  }
  return false;
}

function getMember_(cashapp) {
  var rows = getDataRows_(getSheet_('Members'));
  cashapp = cleanString_(cashapp).toLowerCase().replace('@','').replace('$','');
  for (var i = 0; i < rows.length; i++) {
    var ca = cleanString_(rows[i][2]).toLowerCase().replace('@','').replace('$','');
    if (ca === cashapp) {
      return {
        name: cleanString_(rows[i][1]), cashapp: cleanString_(rows[i][2]),
        pin: cleanString_(rows[i][3]), referralCode: cleanString_(rows[i][4]),
        referredBy: cleanString_(rows[i][5]), status: cleanString_(rows[i][6] || 'active'),
        joinDate: cleanString_(rows[i][7] || ''), role: cleanString_(rows[i][8] || 'member')
      };
    }
  }
  return null;
}

function generateReferralCode_(name) {
  var base = name.replace(/[^a-zA-Z0-9]/g,'').substring(0,4).toUpperCase();
  var rand = Math.floor(Math.random() * 9000) + 1000;
  return base + rand;
}

function registerMember_(p) {
  try {
    var name    = cleanString_(p.name || '');
    var cashapp = cleanString_(p.cashapp || '').replace('$','');
    var pin     = cleanString_(p.pin || '');
    var refBy   = cleanString_(p.referredBy || '');
    if (!name || !cashapp || !pin) return jsonOutput_({ ok: false, message: 'Name, Cash App, and PIN are required.' });
    if (pin.length !== 4 || isNaN(Number(pin))) return jsonOutput_({ ok: false, message: 'PIN must be exactly 4 digits.' });
    var existing = getMember_(cashapp);
    if (existing) return jsonOutput_({ ok: false, message: 'An account with that Cash App already exists. Please log in.' });
    var referralCode = generateReferralCode_(name);
    var joinDate = todayEastern_();
    var role = isAdmin_(cashapp) ? 'admin' : 'member';
    var sheet = getSheet_('Members');
    sheet.appendRow([Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), name, '$' + cashapp, pin, referralCode, refBy, 'pending', joinDate, role]);
    return jsonOutput_({ ok: true, pending: true, message: 'Account created! Your account is pending admin approval. You will be notified once approved.' });
  } catch(err) {
    return jsonOutput_({ ok: false, message: 'Register error: ' + err.toString() });
  }
}

function loginMember_(p) {
  try {
    var cashapp = cleanString_(p.cashapp || '').replace('$','');
    var pin = cleanString_(p.pin || '');
    if (!cashapp || !pin) return jsonOutput_({ ok: false, message: 'Cash App and PIN are required.' });
    var member = getMember_(cashapp);
    if (!member) return jsonOutput_({ ok: false, message: 'No account found. Please register first.' });
    if (member.pin !== pin) return jsonOutput_({ ok: false, message: 'Incorrect PIN.' });
    if (member.status === 'pending') return jsonOutput_({ ok: false, pending: true, message: 'Your account is pending approval. You will be notified once approved.' });
    if (member.status === 'rejected') return jsonOutput_({ ok: false, message: 'Your account application was not approved. Contact admin.' });
    if (member.status === 'suspended') return jsonOutput_({ ok: false, message: 'Account suspended. Contact admin.' });
    if (member.status !== 'approved' && member.status !== 'active' && !isAdmin_(cashapp)) return jsonOutput_({ ok: false, pending: true, message: 'Your account is pending approval.' });
    var role = isAdmin_(cashapp) ? 'admin' : (member.role || 'member');
    return jsonOutput_({ ok: true, message: 'Welcome back, ' + member.name + '!', member: { name: member.name, cashapp: member.cashapp, referralCode: member.referralCode, referredBy: member.referredBy, role: role } });
  } catch(err) {
    return jsonOutput_({ ok: false, message: 'Login error: ' + err.toString() });
  }
}

function getMemberStats_(p) {
  var cashapp = cleanString_(p.cashapp || '');
  var member = getMember_(cashapp);
  if (!member) return jsonOutput_({ ok: false, message: 'Member not found.' });
  var rows = getDataRows_(getSheet_('Members'));
  var referrals = rows.filter(function(r) { return cleanString_(r[5]).toLowerCase() === member.referralCode.toLowerCase(); });
  return jsonOutput_({ ok: true, member: member, referralCount: referrals.length, weekEarnings: 0, totalEarnings: 0 });
}

function upgradeMemberRole_(cashapp, role) {
  var sheet = getSheet_('Members');
  var rows = sheet.getDataRange().getValues();
  var ca = cleanString_(cashapp).toLowerCase().replace('$','');
  for (var i = 1; i < rows.length; i++) {
    if (cleanString_(rows[i][2]).toLowerCase().replace('$','') === ca) {
      sheet.getRange(i + 1, 9).setValue(role);
      return true;
    }
  }
  return false;
}

function generateVIPCode_() {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  var code = 'VIP';
  for (var i = 0; i < 6; i++) code += chars[Math.floor(Math.random() * chars.length)];
  return code;
}

function createVIPCode_(p) {
  if (!isAdmin_(p.cashapp)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
  var code = generateVIPCode_();
  getSheet_('VIPCodes').appendRow([code, Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), '', '']);
  return jsonOutput_({ ok: true, code: code, message: 'VIP code created: ' + code });
}

function redeemVIPCode_(p) {
  var cashapp = cleanString_(p.cashapp || '');
  var code = cleanString_(p.code || '').toUpperCase();
  if (!cashapp || !code) return jsonOutput_({ ok: false, message: 'Cash App and code required.' });
  var sheet = getSheet_('VIPCodes');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (cleanString_(rows[i][0]).toUpperCase() === code) {
      if (cleanString_(rows[i][2])) return jsonOutput_({ ok: false, message: 'Code already used.' });
      sheet.getRange(i + 1, 3).setValue(cashapp);
      sheet.getRange(i + 1, 4).setValue(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'));
      upgradeMemberRole_(cashapp, 'vip');
      return jsonOutput_({ ok: true, message: 'VIP access granted! Welcome to the inner circle.' });
    }
  }
  return jsonOutput_({ ok: false, message: 'Invalid VIP code.' });
}

// -- Bets ---------------------------------------------------------------------

function submitBetWithWalletCheck_(p) {
  var cashapp = cleanString_(p.cashapp || '');
  var pick = cleanString_(p.pick || '');
  var amount = cleanString_(p.amount || '');
  var payout = cleanString_(p.payout || '');
  var gameDesc = cleanString_(p.gameDesc || '');
  if (!cashapp || !pick || !amount) return jsonOutput_({ ok: false, message: 'Required fields missing.' });
  var amtNum = parseFloat(amount.replace(/[^0-9.]/g, ''));
  var status = 'Pending';
  var autoApproved = false;
  if (cashapp && getWalletBalance_(cashapp) >= amtNum) {
    status = 'Approved';
    autoApproved = true;
    getSheet_('Wallet').appendRow([cashapp, 'bet', amtNum, 'approved', 'Bet: ' + pick, Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss')]);
  }
  getSheet_('BetLedger').appendRow([Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), cashapp, gameDesc, pick, amount, '-', status, payout, '']);
  var msg = autoApproved ? 'Bet placed and deducted from your wallet! Good luck.' : 'Bet submitted! Send ' + amount + ' via Cash App to $FreeMoneyDot to lock it in.';
  return jsonOutput_({ ok: true, message: msg, autoApproved: autoApproved });
}

function getBetHistory_(p) {
  var cashapp = cleanString_(p.cashapp || '');
  if (!cashapp) return jsonOutput_({ ok: false, message: 'Cash App required.' });
  var rows = getDataRows_(getSheet_('BetLedger'));
  var bets = rows.filter(function(r) {
    return cleanString_(r[1]).toLowerCase().replace('$','') === cashapp.toLowerCase().replace('$','');
  }).slice(-20).reverse().map(function(r) {
    return { timestamp: cleanString_(r[0]), game: cleanString_(r[2]), pick: cleanString_(r[3]), amount: cleanString_(r[4]), odds: cleanString_(r[5]), status: cleanString_(r[6]), payout: cleanString_(r[7] || '') };
  });
  return jsonOutput_({ ok: true, bets: bets });
}

function getAllBets_(p) {
  if (!isAdmin_(p.cashapp)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
  var rows = getDataRows_(getSheet_('BetLedger'));
  var bets = rows.map(function(r, i) {
    return { rowNum: i + 1, timestamp: cleanString_(r[0]), cashapp: cleanString_(r[1]), game: cleanString_(r[2]), pick: cleanString_(r[3]), amount: cleanString_(r[4]), odds: cleanString_(r[5]), status: cleanString_(r[6]), payout: cleanString_(r[7] || '') };
  }).filter(function(b) { return b.cashapp; });
  return jsonOutput_({ ok: true, bets: bets });
}

function gradeBet_(p) {
  if (!isAdmin_(p.adminCashapp)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
  var sheet = getSheet_('BetLedger');
  var rows = sheet.getDataRange().getValues();
  var rowNum = Number(p.rowNum);
  var outcome = cleanString_(p.outcome || '');
  if (!rowNum || !outcome) return jsonOutput_({ ok: false, message: 'Row and outcome required.' });
  sheet.getRange(rowNum, 7).setValue(outcome);
  sheet.getRange(rowNum, 9).setValue(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'));
  var cashapp = cleanString_(rows[rowNum-1][1]);
  var payout = cleanString_(rows[rowNum-1][7] || '');
  if (outcome === 'Lost') {
    var lossAmount = parseFloat(cleanString_(rows[rowNum-1][4]).replace(/[^0-9.]/g,'')) || 0;
    creditReferralCommission_(cashapp, lossAmount);
  }
  // Notify member
  notifyBetGraded_(cashapp, cleanString_(rows[rowNum-1][3]), outcome, payout);
    try {
    var _res = cleanString_(p.result||'').toLowerCase();
    var _ca  = cleanString_(p.cashapp||'');
    if (_ca) {
      if (_res === 'win')  sendPushToUser_(_ca, 'You won!', 'Your bet has been graded as a win. Check your wallet.', '');
      else if (_res === 'loss') sendPushToUser_(_ca, 'Bet result', 'Your bet has been graded. Better luck next time.', '');
    }
  } catch(_e){}

  return jsonOutput_({ ok: true, message: 'Bet graded: ' + outcome });
}

// -- Poker --------------------------------------------------------------------

function createPokerGame_(p) {
  if (!isAdmin_(p.cashapp)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
  var id = 'PKR' + Date.now();
  getSheet_('PokerGames').appendRow([id, cleanString_(p.title || 'VIP Poker Game'), cleanString_(p.date || ''), cleanString_(p.time || ''), Number(p.maxSeats || 9), Number(p.buyInMin || 50), Number(p.buyInMax || 500), 'open', cleanString_(p.cashapp)]);
  return jsonOutput_({ ok: true, gameId: id, message: 'Poker game created.' });
}

function getPokerGames_() {
  var rows = getDataRows_(getSheet_('PokerGames'));
  var seatRows = getDataRows_(getSheet_('PokerSeats'));
  var games = rows.filter(function(r) { return cleanString_(r[7]) === 'open'; }).map(function(r) {
    var gameId = cleanString_(r[0]);
    var seats = seatRows.filter(function(s) { return cleanString_(s[0]) === gameId; });
    return { id: gameId, title: cleanString_(r[1]), date: cleanString_(r[2]), time: cleanString_(r[3]), maxSeats: Number(r[4]), buyInMin: Number(r[5]), buyInMax: Number(r[6]), status: cleanString_(r[7]), seatsLeft: Number(r[4]) - seats.length, seats: seats.map(function(s) { return { cashapp: cleanString_(s[1]), name: cleanString_(s[2]), buyIn: cleanString_(s[3]), status: cleanString_(s[4]) }; }) };
  });
  return jsonOutput_({ ok: true, games: games });
}

function reservePokerSeat_(p) {
  var cashapp = cleanString_(p.cashapp || '');
  var gameId = cleanString_(p.gameId || '');
  var buyIn = cleanString_(p.buyIn || '');
  if (!cashapp || !gameId || !buyIn) return jsonOutput_({ ok: false, message: 'All fields required.' });
  var member = getMember_(cashapp);
  if (!member) return jsonOutput_({ ok: false, message: 'Please log in first.' });
  if (member.role !== 'vip' && member.role !== 'admin') return jsonOutput_({ ok: false, message: 'VIP access required to join poker games.' });
  var seats = getDataRows_(getSheet_('PokerSeats'));
  for (var i = 0; i < seats.length; i++) {
    if (cleanString_(seats[i][0]) === gameId && cleanString_(seats[i][1]).toLowerCase() === cashapp.toLowerCase()) return jsonOutput_({ ok: false, message: 'You already have a seat reserved.' });
  }
  getSheet_('PokerSeats').appendRow([gameId, member.cashapp, member.name, buyIn, 'Pending', '']);
  return jsonOutput_({ ok: true, message: 'Seat reserved! Send ' + buyIn + ' via Cash App to confirm your seat.' });
}

function confirmPokerSeat_(p) {
  if (!isAdmin_(p.adminCashapp)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
  var sheet = getSheet_('PokerSeats');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (cleanString_(rows[i][0]) === p.gameId && cleanString_(rows[i][1]).toLowerCase().replace('$','') === p.cashapp.toLowerCase().replace('$','')) {
      sheet.getRange(i + 1, 5).setValue('Confirmed');
      sheet.getRange(i + 1, 6).setValue(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'));
      return jsonOutput_({ ok: true, message: 'Seat confirmed.' });
    }
  }
  return jsonOutput_({ ok: false, message: 'Seat not found.' });
}

// -- Admin tools --------------------------------------------------------------

function getAllMembers_(p) {
  if (!isAdmin_(p.cashapp)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
  var rows = getDataRows_(getSheet_('Members'));
  var members = rows.map(function(r) {
    return { name: cleanString_(r[1]), cashapp: cleanString_(r[2]), role: cleanString_(r[8] || 'member'), status: cleanString_(r[6] || 'active'), joinDate: cleanString_(r[7] || ''), referralCode: cleanString_(r[4] || ''), referredBy: cleanString_(r[5] || '') };
  }).filter(function(m) { return m.cashapp; });
  return jsonOutput_({ ok: true, members: members });
}

function setMemberRole_(p) {
  if (!isAdmin_(p.adminCashapp)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
  upgradeMemberRole_(p.cashapp, p.role);
  return jsonOutput_({ ok: true, message: 'Role updated to ' + p.role });
}

function setMemberStatus_(p) {
  if (!isAdmin_(p.adminCashapp)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
  var sheet = getSheet_('Members');
  var rows = sheet.getDataRange().getValues();
  var target = cleanString_(p.cashapp).toLowerCase().replace('$','');
  for (var i = 1; i < rows.length; i++) {
    if (cleanString_(rows[i][2]).toLowerCase().replace('$','') === target) {
      sheet.getRange(i + 1, 7).setValue(p.status);
      return jsonOutput_({ ok: true, message: 'Status updated.' });
    }
  }
  return jsonOutput_({ ok: false, message: 'Member not found.' });
}

function markEntryPaid_(p) {
  if (!isAdmin_(p.adminCashapp)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
  getSheet_(cleanString_(p.sheet || 'FlatSpinEntries')).getRange(Number(p.rowNum) + 1, 6).setValue(cleanString_(p.status || 'Paid'));
  return jsonOutput_({ ok: true, message: 'Entry updated.' });
}

function removeEntry_(p) {
  if (!isAdmin_(p.adminCashapp)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
  getSheet_(cleanString_(p.sheet || 'FlatSpinEntries')).deleteRow(Number(p.rowNum) + 1);
  return jsonOutput_({ ok: true, message: 'Entry removed.' });
}

function updateSpadesMatch_(p) {
  if (!isAdmin_(p.adminCashapp)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
  var teamA = cleanString_(p.teamA || '');
  var teamB = cleanString_(p.teamB || '');
  var aWins = Number(p.teamAWins || 0);
  var bWins = Number(p.teamBWins || 0);
  if (!teamA || !teamB) return jsonOutput_({ ok: false, message: 'Both team names required.' });
  var sheet = getSheet_('SpadesMatches');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var a = cleanString_(rows[i][1]).toLowerCase();
    var b = cleanString_(rows[i][2]).toLowerCase();
    if ((a === teamA.toLowerCase() && b === teamB.toLowerCase()) || (a === teamB.toLowerCase() && b === teamA.toLowerCase())) {
      sheet.getRange(i + 1, 4).setValue(aWins);
      sheet.getRange(i + 1, 5).setValue(bWins);
      sheet.getRange(i + 1, 8).setValue((aWins >= 2 || bWins >= 2) ? 'Complete' : 'In Progress');
      return jsonOutput_({ ok: true, message: 'Match updated.' });
    }
  }
  sheet.appendRow([teamA + ' vs ' + teamB, teamA, teamB, aWins, bWins, 0, 0, 'In Progress']);
  return jsonOutput_({ ok: true, message: 'Match created and updated.' });
}

// -- Wallet -------------------------------------------------------------------

function getWalletBalance_(cashapp) {
  cashapp = cleanString_(cashapp).toLowerCase().replace('$','');
  var rows = getDataRows_(getSheet_('Wallet'));
  var balance = 0;
  rows.forEach(function(r) {
    var ca = cleanString_(r[0]).toLowerCase().replace('$','');
    if (ca === cashapp && cleanString_(r[3]) === 'approved') {
      var type = cleanString_(r[1]);
      var amt = parseFloat(cleanString_(r[2]).replace('$','')) || 0;
      if (type === 'deposit' || type === 'win' || type === 'credit') balance += amt;
      if (type === 'bet' || type === 'debit' || type === 'loss') balance -= amt;
    }
  });
  return Math.round(balance * 100) / 100;
}

function requestDeposit_(p) {
  var cashapp = cleanString_(p.cashapp || '');
  var amount = cleanString_(p.amount || '');
  var note = cleanString_(p.note || '');
  if (!cashapp || !amount) return jsonOutput_({ ok: false, message: 'Cash App and amount required.' });
  var member = getMember_(cashapp);
  if (!member) return jsonOutput_({ ok: false, message: 'Please log in first.' });
  getSheet_('Wallet').appendRow([member.cashapp, 'deposit', amount, 'pending', note, Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss')]);
  return jsonOutput_({ ok: true, message: 'Deposit request submitted! Admin will approve after payment is confirmed.' });
}

function approveDeposit_(p) {
  if (!isAdmin_(p.adminCashapp)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
  var sheet = getSheet_('Wallet');
  var rowNum = Number(p.rowNum);
  sheet.getRange(rowNum + 1, 4).setValue(p.status || 'approved');
    try {
    sendPushToUser_(cleanString_(p.cashapp||''), 'Deposit approved', 'Your deposit has been approved.', '');
  } catch(_e){}

  return jsonOutput_({ ok: true, message: 'Deposit ' + (p.status || 'approved') + '.' });
}

function adminCreditWallet_(p) {
  if (!isAdmin_(p.adminCashapp)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
  var cashapp = cleanString_(p.cashapp || '');
  var amount = cleanString_(p.amount || '');
  var type = cleanString_(p.type || 'credit');
  var note = cleanString_(p.note || 'Admin adjustment');
  if (!cashapp || !amount) return jsonOutput_({ ok: false, message: 'Cash App and amount required.' });
  getSheet_('Wallet').appendRow([cashapp, type, amount, 'approved', note, Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss')]);
  return jsonOutput_({ ok: true, message: 'Wallet updated.' });
}

function getWalletHistory_(p) {
  var cashapp = cleanString_(p.cashapp || '').toLowerCase().replace('$','');
  if (!cashapp) return jsonOutput_({ ok: false, message: 'Cash App required.' });
  var rows = getDataRows_(getSheet_('Wallet'));
  var txns = rows.filter(function(r) {
    return cleanString_(r[0]).toLowerCase().replace('$','') === cashapp;
  }).slice(-20).reverse().map(function(r) {
    return { cashapp: cleanString_(r[0]), type: cleanString_(r[1]), amount: cleanString_(r[2]), status: cleanString_(r[3]), note: cleanString_(r[4] || ''), timestamp: cleanString_(r[5] || '') };
  });
  var balance = getWalletBalance_(cashapp);
  return jsonOutput_({ ok: true, balance: balance, transactions: txns });
}

function getPendingDeposits_(p) {
  if (!isAdmin_(p.cashapp)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
  var rows = getDataRows_(getSheet_('Wallet'));
  var pending = rows.map(function(r, i) {
    return { rowNum: i+1, cashapp: cleanString_(r[0]), type: cleanString_(r[1]), amount: cleanString_(r[2]), status: cleanString_(r[3]), note: cleanString_(r[4] || ''), timestamp: cleanString_(r[5] || '') };
  }).filter(function(r) { return r.status === 'pending'; });
  return jsonOutput_({ ok: true, deposits: pending });
}

function getBetTicker_() {
  try { getSheet_('BetLedger'); } catch(e) { return jsonOutput_({ ok: true, bets: [] }); }
  var rows = getDataRows_(getSheet_('BetLedger'));
  var recent = rows.slice(-15).reverse().map(function(r) {
    return { cashapp: cleanString_(r[1]), pick: cleanString_(r[3]), amount: cleanString_(r[4]), timestamp: cleanString_(r[0] || '') };
  }).filter(function(b) { return b.cashapp && b.pick; });
  return jsonOutput_({ ok: true, bets: recent });
}

// -- Schedule -----------------------------------------------------------------

function getSchedule_() {
  var rows = getDataRows_(getSheet_('Schedule'));
  var events = rows.map(function(r) {
    return { id: cleanString_(r[0]), title: cleanString_(r[1]), type: cleanString_(r[2]), date: cleanString_(r[3]), time: cleanString_(r[4]), prize: cleanString_(r[5]), entryFee: cleanString_(r[6]), description: cleanString_(r[7] || ''), status: cleanString_(r[8] || 'active') };
  }).filter(function(e) { return e.title && e.status !== 'deleted'; });
  return jsonOutput_({ ok: true, events: events });
}

function createScheduleEvent_(p) {
  if (!isAdmin_(p.adminCashapp)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
  var id = 'EVT' + Date.now();
  getSheet_('Schedule').appendRow([id, cleanString_(p.title || ''), cleanString_(p.type || 'other'), cleanString_(p.date || ''), cleanString_(p.time || ''), cleanString_(p.prize || ''), cleanString_(p.entryFee || 'Free'), cleanString_(p.description || ''), 'active']);
  return jsonOutput_({ ok: true, id: id, message: 'Event created.' });
}

function deleteScheduleEvent_(p) {
  if (!isAdmin_(p.adminCashapp)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
  var sheet = getSheet_('Schedule');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (cleanString_(rows[i][0]) === cleanString_(p.id)) {
      sheet.getRange(i + 1, 9).setValue('deleted');
      return jsonOutput_({ ok: true, message: 'Event removed.' });
    }
  }
  return jsonOutput_({ ok: false, message: 'Event not found.' });
}

// -- VIP Waitlist -------------------------------------------------------------

function applyVIPWaitlist_(p) {
  var name = cleanString_(p.name || '');
  var cashapp = cleanString_(p.cashapp || '');
  var note = cleanString_(p.note || '');
  if (!name || !cashapp) return jsonOutput_({ ok: false, message: 'Name and Cash App required.' });
  getSheet_('VIPWaitlist').appendRow([Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), name, cashapp, note, 'pending']);
  return jsonOutput_({ ok: true, message: 'VIP application submitted! We will be in touch soon.' });
}

function getVIPWaitlist_(p) {
  if (!isAdmin_(p.cashapp)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
  var rows = getDataRows_(getSheet_('VIPWaitlist'));
  var list = rows.map(function(r, i) {
    return { rowNum: i+1, timestamp: cleanString_(r[0]), name: cleanString_(r[1]), cashapp: cleanString_(r[2]), note: cleanString_(r[3] || ''), status: cleanString_(r[4] || 'pending') };
  }).filter(function(r) { return r.name; });
  return jsonOutput_({ ok: true, applicants: list });
}

function approveVIPWaitlist_(p) {
  if (!isAdmin_(p.adminCashapp)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
  var sheet = getSheet_('VIPWaitlist');
  var rows = sheet.getDataRange().getValues();
  var rowNum = Number(p.rowNum);
  sheet.getRange(rowNum + 1, 5).setValue('approved');
  var code = generateVIPCode_();
  getSheet_('VIPCodes').appendRow([code, Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), '', '']);
  return jsonOutput_({ ok: true, message: 'Approved. VIP code: ' + code, code: code });
}

function rejectVIPWaitlist_(p) {
  if (!isAdmin_(p.adminCashapp)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
  var sheet = getSheet_('VIPWaitlist');
  var rows = sheet.getDataRange().getValues();
  var rowNum = Number(p.rowNum);
  sheet.getRange(rowNum + 1, 5).setValue('rejected');
  return jsonOutput_({ ok: true, message: 'Application rejected.' });
}

// -- Analytics ----------------------------------------------------------------

function getAnalytics_(p) {
  if (!isAdmin_(p.cashapp)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
  try {
    var members = getDataRows_(getSheet_('Members'));
    var bets = getDataRows_(getSheet_('BetLedger'));
    var entries = getDataRows_(getSheet_('FlatSpinEntries'));
    var wallet = getDataRows_(getSheet_('Wallet'));
    var qdRows = getDataRows_(getSheet_('QuickDrawEntries'));
    var totalMembers = members.length;
    var totalBets = bets.length;
    var pendingBets = bets.filter(function(b) { return cleanString_(b[6]) === 'Pending'; }).length;
    var wonBets = bets.filter(function(b) { return cleanString_(b[6]) === 'Won'; }).length;
    var lostBets = bets.filter(function(b) { return cleanString_(b[6]) === 'Lost'; }).length;
    var betVolume = bets.reduce(function(sum, b) { return sum + (parseFloat(cleanString_(b[4]).replace('$','')) || 0); }, 0);
    var totalDeposits = wallet.filter(function(w) { return cleanString_(w[1]) === 'deposit' && cleanString_(w[3]) === 'approved'; }).reduce(function(sum, w) { return sum + (parseFloat(cleanString_(w[2]).replace('$','')) || 0); }, 0);
    var pendingDeposits = wallet.filter(function(w) { return cleanString_(w[3]) === 'pending'; }).length;
    var vipMembers = members.filter(function(m) { return cleanString_(m[8] || '') === 'vip'; }).length;
    var spinEntries = entries.length;
    var paidEntries = entries.filter(function(e) { return cleanString_(e[5]).toLowerCase() === 'paid'; }).length;
    var betsByMember = {};
    bets.forEach(function(b) {
      var ca = cleanString_(b[1]);
      var amt = parseFloat(cleanString_(b[4]).replace('$','')) || 0;
      betsByMember[ca] = (betsByMember[ca] || 0) + amt;
    });
    var topBettors = Object.keys(betsByMember).map(function(ca) { return { cashapp: ca, volume: betsByMember[ca] }; }).sort(function(a,b) { return b.volume - a.volume; }).slice(0, 5);
    return jsonOutput_({ ok: true, totalMembers: totalMembers, vipMembers: vipMembers, totalBets: totalBets, pendingBets: pendingBets, wonBets: wonBets, lostBets: lostBets, betVolume: Math.round(betVolume * 100) / 100, totalDeposits: Math.round(totalDeposits * 100) / 100, pendingDeposits: pendingDeposits, spinEntries: spinEntries, paidSpinEntries: paidEntries, qdEntries: qdRows.length, topBettors: topBettors });
  } catch(err) {
    return jsonOutput_({ ok: false, message: 'Analytics error: ' + String(err) });
  }
}

// -- Misc forms ---------------------------------------------------------------

function submitSpinPropBet_(p) {
  var name = cleanString_(p.bettorname || '');
  var cashapp = cleanString_(p.cashapp || '');
  var answer = cleanString_(p.answer || '');
  var amount = parseFloat(cleanString_(p.betamount || '0'));
  if (!name || !cashapp || !answer) return jsonOutput_({ ok: false, message: 'All fields required.' });
  if (amount < 1 || amount > 50) return jsonOutput_({ ok: false, message: 'Amount must be between $1 and $50.' });
  getSheet_('SpadesSideBets').appendRow([Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), name, cashapp, 'spin_winner_unit', 'Which unit will the Flat Spin winner come from?', answer, '$' + amount, 'Pending']);
  return jsonOutput_({ ok: true, message: 'Prop bet submitted! Send $' + amount + ' via Cash App to lock it in.' });
}

function submitCookingEntry_(p) {
  var name = cleanString_(p.name || '');
  var cashapp = cleanString_(p.cashapp || '');
  var contest = cleanString_(p.contestName || '');
  var unit = cleanString_(p.unit || '');
  var dish = cleanString_(p.dishName || '');
  var payment = cleanString_(p.paymentMethod || '');
  var notes = cleanString_(p.notes || '');
  if (!name || !cashapp || !contest) return jsonOutput_({ ok: false, message: 'Name, Cash App, and contest required.' });
  getSheet_('CookingEntries').appendRow([Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), name, cashapp, contest, unit, dish, payment, notes, 'Pending']);
  return jsonOutput_({ ok: true, message: 'Entry submitted! Bring your 2 flats or send via $FreeMoneyDot to confirm.' });
}

function submitCookingProp_(p) {
  var name = cleanString_(p.name || '');
  var cashapp = cleanString_(p.cashapp || '');
  var contest = cleanString_(p.contest || '');
  var question = cleanString_(p.question || '');
  var answer = cleanString_(p.answer || '');
  var amount = cleanString_(p.amount || '');
  if (!name || !cashapp || !contest || !question || !answer) return jsonOutput_({ ok: false, message: 'All fields required.' });
  getSheet_('CookingProps').appendRow([Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), name, cashapp, contest, question, answer, amount, 'Pending']);
  return jsonOutput_({ ok: true, message: 'Prop bet submitted! Send ' + amount + ' via Cash App to lock it in.' });
}

function submitDominosSignup_(p) {
  var firstName = cleanString_(p.firstName || '');
  var lastName = cleanString_(p.lastName || '');
  var cashapp = cleanString_(p.cashapp || '');
  var unit = cleanString_(p.unit || '');
  var roomNum = cleanString_(p.roomNumber || '');
  var partner = cleanString_(p.partner || '');
  var notes = cleanString_(p.notes || '');
  var name = (firstName + ' ' + lastName).trim();
  if (!name || !cashapp) return jsonOutput_({ ok: false, message: 'Name and Cash App required.' });
  getSheet_('DominosSignups').appendRow([Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), name, cashapp, unit, partner, notes, roomNum]);
  return jsonOutput_({ ok: true, message: 'You are signed up for the Double Sixes Dominos Tournament! See you Saturday at 12PM at the Pavilion.' });
}

function submitDominosProp_(p) {
  var name = cleanString_(p.bettorname || '');
  var cashapp = cleanString_(p.cashapp || '');
  var propid = cleanString_(p.propid || '');
  var answer = cleanString_(p.answer || '');
  var amount = cleanString_(p.betamount || '');
  if (!name || !cashapp || !propid || !answer) return jsonOutput_({ ok: false, message: 'All fields required.' });
  var amt = parseFloat(amount);
  if (isNaN(amt) || amt < 1 || amt > 50) return jsonOutput_({ ok: false, message: 'Amount must be between $1 and $50.' });
  getSheet_('DominosProps').appendRow([Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), name, cashapp, propid, answer, '$' + amt, 'Pending']);
  return jsonOutput_({ ok: true, message: 'Prop bet submitted! Send $' + amt + ' via Cash App to lock it in.' });
}

function getDominosData_() {
  var rows = getDataRows_(getSheet_('DominosSignups'));
  var players = rows.map(function(r) {
    return { name: cleanString_(r[1]), cashapp: cleanString_(r[2]), unit: cleanString_(r[3]), partner: cleanString_(r[4] || ''), roomNumber: cleanString_(r[6] || '') };
  }).filter(function(p) { return p.name; });
  return jsonOutput_({ ok: true, players: players, totalPlayers: players.length });
}

// -- Broadcast ----------------------------------------------------------------

function createBroadcast_(p) {
  if (!isAdmin_(p.adminCashapp)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
  var msg = cleanString_(p.message || '');
  if (!msg) return jsonOutput_({ ok: false, message: 'Message required.' });
  var id = 'BC' + Date.now();
  var sheet = getSheet_('Broadcasts');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (cleanString_(rows[i][4]) === 'active') sheet.getRange(i+1, 5).setValue('inactive');
  }
  sheet.appendRow([id, msg, Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), cleanString_(p.adminCashapp), 'active']);
    try {
    var bmsg = cleanString_(p.message || p.text || '');
    if (bmsg) sendPushToAll_('Broadcast', bmsg, '');
  } catch(_e){}

  return jsonOutput_({ ok: true, message: 'Broadcast sent.' });
}

function clearBroadcast_(p) {
  if (!isAdmin_(p.adminCashapp)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
  var sheet = getSheet_('Broadcasts');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) { sheet.getRange(i+1, 5).setValue('inactive'); }
  return jsonOutput_({ ok: true, message: 'Broadcast cleared.' });
}

function getActiveBroadcast_() {
  var rows = getDataRows_(getSheet_('Broadcasts'));
  for (var i = rows.length - 1; i >= 0; i--) {
    if (cleanString_(rows[i][4]) === 'active') {
      return jsonOutput_({ ok: true, broadcast: { id: cleanString_(rows[i][0]), message: cleanString_(rows[i][1]), createdAt: cleanString_(rows[i][2]) } });
    }
  }
  return jsonOutput_({ ok: true, broadcast: null });
}

// -- Payouts ------------------------------------------------------------------

function recordPayout_(p) {
  if (!isAdmin_(p.adminCashapp)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
  var cashapp = cleanString_(p.cashapp || '');
  var amount = cleanString_(p.amount || '');
  var reason = cleanString_(p.reason || '');
  var method = cleanString_(p.method || 'Cash App');
  if (!cashapp || !amount || !reason) return jsonOutput_({ ok: false, message: 'All fields required.' });
  getSheet_('Payouts').appendRow([Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), cashapp, amount, reason, method, cleanString_(p.adminCashapp)]);
  getSheet_('Wallet').appendRow([cashapp, 'win', amount, 'approved', reason, Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss')]);
  return jsonOutput_({ ok: true, message: 'Payout recorded and wallet credited.' });
}

function getPayoutHistory_(p) {
  var cashapp = cleanString_(p.cashapp || '').toLowerCase().replace('$','');
  if (!cashapp) return jsonOutput_({ ok: false, message: 'Cash App required.' });
  var rows = getDataRows_(getSheet_('Payouts'));
  var history = rows.filter(function(r) { return cleanString_(r[1]).toLowerCase().replace('$','') === cashapp; }).reverse().slice(0, 30).map(function(r) { return { timestamp: cleanString_(r[0]), cashapp: cleanString_(r[1]), amount: cleanString_(r[2]), reason: cleanString_(r[3]), method: cleanString_(r[4]) }; });
  var total = rows.filter(function(r) { return cleanString_(r[1]).toLowerCase().replace('$','') === cashapp; }).reduce(function(sum, r) { return sum + (parseFloat(cleanString_(r[2]).replace(/[^0-9.]/g,'')) || 0); }, 0);
  return jsonOutput_({ ok: true, payouts: history, totalPaid: Math.round(total*100)/100 });
}

function getAllPayouts_(p) {
  if (!isAdmin_(p.cashapp)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
  var rows = getDataRows_(getSheet_('Payouts'));
  var payouts = rows.reverse().slice(0, 50).map(function(r) { return { timestamp: cleanString_(r[0]), cashapp: cleanString_(r[1]), amount: cleanString_(r[2]), reason: cleanString_(r[3]), method: cleanString_(r[4]) }; });
  return jsonOutput_({ ok: true, payouts: payouts });
}

// -- Referrals ----------------------------------------------------------------

function getReferralStats_(p) {
  var cashapp = cleanString_(p.cashapp || '');
  var member = getMember_(cashapp);
  if (!member) return jsonOutput_({ ok: false, message: 'Member not found.' });
  var rows = getDataRows_(getSheet_('Members'));
  var referrals = rows.filter(function(r) { return cleanString_(r[5]).toLowerCase() === member.referralCode.toLowerCase(); }).map(function(r) { return { name: cleanString_(r[1]), cashapp: cleanString_(r[2]), joinDate: cleanString_(r[7] || '') }; });
  return jsonOutput_({ ok: true, referralCode: member.referralCode, referrals: referrals, count: referrals.length });
}

function creditReferralCommission_(memberCashapp, lossAmount) {
  try {
    var member = getMember_(memberCashapp);
    if (!member || !member.referredBy) return;
    var rows = getDataRows_(getSheet_('Members'));
    var referrer = null;
    for (var i = 0; i < rows.length; i++) {
      if (cleanString_(rows[i][4]).toLowerCase() === member.referredBy.toLowerCase()) {
        referrer = { cashapp: cleanString_(rows[i][2]) };
        break;
      }
    }
    if (!referrer) return;
    var commission = Math.round(lossAmount * 0.20 * 100) / 100;
    if (commission < 0.01) return;
    getSheet_('Wallet').appendRow([referrer.cashapp, 'credit', '$' + commission, 'approved', 'Referral commission from ' + memberCashapp, Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss')]);
  } catch(err) { Logger.log('Referral commission error: ' + String(err)); }
}

function computeReferralCredits_() {
  var members = getDataRows_(getSheet_('Members'));
  var bets    = getDataRows_(getSheet_('BetLedger'));
  var wallet  = getDataRows_(getSheet_('Wallet'));
  var now = new Date();
  var weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);

  // Map: referralCode -> referrerCashapp
  var referralMap = {};
  members.forEach(function(m) {
    var code = cleanString_(m[4]);
    var ca   = cleanString_(m[2]);
    if (code) referralMap[code.toLowerCase()] = ca;
  });

  // Compute total weekly losses per referrer (sum across all their referees)
  // and a referral count.
  var ownedByReferrer = {}; // refCaLower -> { referrer, totalLosses, referralCount }
  members.forEach(function(m) {
    var memberCa = cleanString_(m[2]).toLowerCase().replace('$','');
    var refCode  = cleanString_(m[5] || '').toLowerCase();
    if (!refCode) return;
    var referrerCa = referralMap[refCode];
    if (!referrerCa) return;
    var losses = bets.filter(function(b) {
      var t = new Date(cleanString_(b[0]));
      return cleanString_(b[6]) === 'Lost'
          && t >= weekAgo
          && cleanString_(b[1]).toLowerCase().replace('$','') === memberCa;
    }).reduce(function(s, b) {
      return s + (parseFloat(cleanString_(b[4]).replace(/[^0-9.]/g,'')) || 0);
    }, 0);
    if (losses <= 0) return;
    var k = referrerCa.toLowerCase();
    if (!ownedByReferrer[k]) ownedByReferrer[k] = { referrer: referrerCa, totalLosses: 0, referralCount: 0 };
    ownedByReferrer[k].totalLosses += losses;
    ownedByReferrer[k].referralCount++;
  });

  // Compute total already-paid per referrer this week (any referral-commission
  // wallet row, regardless of source label - per-source or weekly-batch).
  // Wallet columns: 0=cashapp 1=type 2=amount 3=status 4=note 5=timestamp
  var alreadyPaidByReferrer = {};
  for (var i = 0; i < wallet.length; i++) {
    var note = cleanString_(wallet[i][4]);
    var ts   = new Date(cleanString_(wallet[i][5]));
    if (!note || isNaN(ts) || ts < weekAgo) continue;
    if (!/^Referral commission from /.test(note)) continue;
    var refCa = cleanString_(wallet[i][0]).toLowerCase();
    var amt   = parseFloat(cleanString_(wallet[i][2]).replace(/[^0-9.]/g,'')) || 0;
    alreadyPaidByReferrer[refCa] = (alreadyPaidByReferrer[refCa] || 0) + amt;
  }

  // Build the credits list: 20% of total weekly losses minus what was paid.
  var credits = [];
  Object.keys(ownedByReferrer).forEach(function(k) {
    var entry = ownedByReferrer[k];
    var gross = Math.round(entry.totalLosses * 0.20 * 100) / 100;
    var paid  = alreadyPaidByReferrer[k] || 0;
    var net   = Math.round((gross - paid) * 100) / 100;
    if (net < 0.01) return;
    credits.push({
      referrer:      entry.referrer,
      totalLosses:   Math.round(entry.totalLosses * 100) / 100,
      referralCount: entry.referralCount,
      grossOwed:     gross,
      alreadyPaid:   Math.round(paid * 100) / 100,
      amount:        net
    });
  });
  return credits;
}

function getReferralCommissions_(p) {
  if (!isAdmin_(p.cashapp)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
  try { var credits = computeReferralCredits_(); return jsonOutput_({ ok: true, credits: credits }); }
  catch(err) { return jsonOutput_({ ok: false, message: 'Error: ' + String(err) }); }
}

function applyReferralCredits_(p) {
  if (!isAdmin_(p.adminCashapp)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
  try {
    var credits = computeReferralCredits_();
    if (!credits.length) return jsonOutput_({ ok: true, count: 0, message: 'No commissions owed (everything already paid out).' });
    var wallet = getSheet_('Wallet');
    var ts = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
    var totalPaid = 0;
    credits.forEach(function(cr) {
      // Note format matches creditReferralCommission_ so the dedup picks it up.
      // We aggregate as one row per referrer with a generic source label so dedup keys are referrer|allmembers.
      // For best dedup, we use 'Referral commission from weekly-batch'.
      wallet.appendRow([cr.referrer, 'credit', '$' + cr.amount, 'approved',
        'Referral commission from weekly-batch (' + cr.referralCount + ' refs, $' + cr.totalLosses + ' losses)',
        ts]);
      totalPaid += cr.amount;
    });
    return jsonOutput_({ ok: true, count: credits.length, totalPaid: totalPaid,
      message: 'Applied ' + credits.length + ' commission(s) totalling $' + totalPaid.toFixed(2) });
  } catch(err) { return jsonOutput_({ ok: false, message: String(err) }); }
}

// -- Auto Grade ---------------------------------------------------------------

function autoGradeBetsFromAPI_() {
  var ODDS_KEY = '2655755ed6034eadcfbfd9fd5071c564';
  var SPORTS = ['basketball_nba', 'baseball_mlb', 'icehockey_nhl', 'basketball_wnba'];
  var sheet = getSheet_('BetLedger');
  var rows = sheet.getDataRange().getValues();
  var graded = 0;
  var errors = [];
  var pendingBets = [];
  for (var i = 1; i < rows.length; i++) {
    if (cleanString_(rows[i][6]) === 'Pending' || cleanString_(rows[i][6]) === 'Approved') {
      var betTime = new Date(cleanString_(rows[i][0]));
      var hoursSince = (new Date() - betTime) / 3600000;
      if (hoursSince >= 4) { pendingBets.push({ rowIndex: i, row: rows[i] }); }
    }
  }
  if (!pendingBets.length) return jsonOutput_({ ok: true, graded: 0, message: 'No bets ready to grade.' });
  var scoreMap = {};
  SPORTS.forEach(function(sport) {
    try {
      var url = 'https://api.the-odds-api.com/v4/sports/' + sport + '/scores/?apiKey=' + ODDS_KEY + '&daysFrom=2';
      var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (resp.getResponseCode() !== 200) return;
      var games = JSON.parse(resp.getContentText());
      games.forEach(function(g) {
        if (!g.completed || !g.scores) return;
        var homeScore = null, awayScore = null;
        g.scores.forEach(function(s) { if (s.name === g.home_team) homeScore = parseFloat(s.score); if (s.name === g.away_team) awayScore = parseFloat(s.score); });
        if (homeScore === null || awayScore === null) return;
        scoreMap[g.home_team] = { home: g.home_team, away: g.away_team, homeScore: homeScore, awayScore: awayScore };
        scoreMap[g.away_team] = { home: g.home_team, away: g.away_team, homeScore: homeScore, awayScore: awayScore };
      });
    } catch(e) { errors.push(sport + ': ' + e.toString()); }
  });
  pendingBets.forEach(function(b) {
    try {
      var pick = cleanString_(b.row[3]);
      var payout = cleanString_(b.row[7]);
      var cashapp = cleanString_(b.row[1]);
      var result = gradeSingleBet_(pick, scoreMap);
      if (result === null) return;
      sheet.getRange(b.rowIndex + 1, 7).setValue(result);
      sheet.getRange(b.rowIndex + 1, 9).setValue(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'));
      graded++;
      if (result === 'Won' && payout && cashapp) {
        getSheet_('Wallet').appendRow([cashapp, 'win', payout, 'approved', 'Auto-graded win: ' + pick, Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss')]);
      }
      if (result === 'Lost' && cashapp) {
        var lossAmt = parseFloat(cleanString_(b.row[4]).replace(/[^0-9.]/g,'')) || 0;
        if (lossAmt > 0) creditReferralCommission_(cashapp, lossAmt);
      }
      notifyBetGraded_(cashapp, pick, result, payout);
    } catch(e) { errors.push('Bet grade error: ' + e.toString()); }
  });
  return jsonOutput_({ ok: true, graded: graded, pending: pendingBets.length, errors: errors });
}

function gradeSingleBet_(pick, scoreMap) {
  var legs = pick.split(' + ');
  var results = [];
  for (var i = 0; i < legs.length; i++) {
    var r = gradeLeg_(legs[i].trim(), scoreMap);
    if (r === null) return null;
    results.push(r);
  }
  if (results.indexOf('Lost') >= 0) return 'Lost';
  if (results.indexOf('Push') >= 0) return 'Push';
  if (results.length > 0 && results.indexOf('Won') >= 0) return 'Won';
  return null;
}

function gradeLeg_(leg, scoreMap) {
  var isOver = leg.toUpperCase().indexOf('OVER') === 0;
  var isUnder = leg.toUpperCase().indexOf('UNDER') === 0;
  if (isOver || isUnder) { return null; }
  var tokens = leg.replace(/\([^)]+\)/g,'').trim().split(/\s+/);
  if (tokens.length < 2) return null;
  var abbr = tokens[0].toUpperCase();
  var spread = parseFloat(tokens[1]);
  if (isNaN(spread)) return null;
  var gameData = null;
  Object.keys(scoreMap).forEach(function(teamName) {
    var teamAbbr = teamName.split(' ').map(function(w){ return w[0]; }).join('').toUpperCase();
    var lastWord = teamName.split(' ').pop().substring(0,3).toUpperCase();
    if (teamAbbr === abbr || lastWord === abbr || teamName.toUpperCase().indexOf(abbr) === 0) {
      gameData = { team: teamName, data: scoreMap[teamName] };
    }
  });
  if (!gameData) return null;
  var d = gameData.data;
  var isHome = gameData.team === d.home;
  var teamScore = isHome ? d.homeScore : d.awayScore;
  var otherScore = isHome ? d.awayScore : d.homeScore;
  var margin = teamScore - otherScore + spread;
  if (margin > 0) return 'Won';
  if (margin < 0) return 'Lost';
  return 'Push';
}

// -- Daily Report -------------------------------------------------------------

function getDailyReport_(p) {
  if (!isAdmin_(p.cashapp)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
  try {
    var today = todayEastern_();
    var bets = getDataRows_(getSheet_('BetLedger'));
    var members = getDataRows_(getSheet_('Members'));
    var entries = getDataRows_(getSheet_('FlatSpinEntries'));
    var wallet = getDataRows_(getSheet_('Wallet'));
    var todayBets = bets.filter(function(r) { return cleanString_(r[0]).substring(0, 10) === today; });
    var wonToday = todayBets.filter(function(r) { return cleanString_(r[6]) === 'Won'; });
    var lostToday = todayBets.filter(function(r) { return cleanString_(r[6]) === 'Lost'; });
    var pendingBets = bets.filter(function(r) { return cleanString_(r[6]) === 'Pending'; });
    var needsGrading = bets.filter(function(r) { return cleanString_(r[6]) === 'Needs Grading'; });
    var totalBetAmt = todayBets.reduce(function(s, r) { return s + (parseFloat(cleanString_(r[4]).replace(/[^0-9.]/g,'')) || 0); }, 0);
    var payoutsOwed = wonToday.reduce(function(s, r) { return s + (parseFloat(cleanString_(r[7]).replace(/[^0-9.]/g,'')) || 0); }, 0);
    var todayEntries = entries.filter(function(r) { return cleanString_(r[0]).substring(0, 10) === today; });
    var paidEntries = todayEntries.filter(function(r) { return cleanString_(r[5]).toLowerCase() === 'paid'; });
    var pendingDeposits = wallet.filter(function(r) { return cleanString_(r[3]) === 'pending'; });
    var newMembers = members.filter(function(r) { return cleanString_(r[0]).substring(0, 10) === today; });
    return jsonOutput_({
      ok: true, date: today,
      bets: { total: todayBets.length, won: wonToday.length, lost: lostToday.length, pending: pendingBets.length, needsGrading: needsGrading.length, volume: Math.round(totalBetAmt * 100) / 100, payoutsOwed: Math.round(payoutsOwed * 100) / 100,
        winners: wonToday.map(function(r) { return { cashapp: cleanString_(r[1]), pick: cleanString_(r[3]), amount: cleanString_(r[4]), payout: cleanString_(r[7]) }; }),
        losers: lostToday.map(function(r) { return { cashapp: cleanString_(r[1]), pick: cleanString_(r[3]), amount: cleanString_(r[4]) }; }),
        gradingNeeded: needsGrading.map(function(r) { return { cashapp: cleanString_(r[1]), pick: cleanString_(r[3]), amount: cleanString_(r[4]), time: cleanString_(r[0]) }; })
      },
      spinEntries: { total: todayEntries.length, paid: paidEntries.length },
      pendingDeposits: pendingDeposits.length,
      newMembers: newMembers.length
    });
  } catch(err) { return jsonOutput_({ ok: false, message: 'Report error: ' + String(err) }); }
}

// -- Stripe -------------------------------------------------------------------
// Sheet: StripePayments | Timestamp | IntentID | CashApp | Amount | Description | Status

function createPaymentIntent_(p) {
  var amount = parseInt(p.amount || '700');
  var description = cleanString_(p.description || 'FreeMoneyDot Entry');
  var cashapp = cleanString_(p.cashapp || 'guest');
  if (amount < 100) return jsonOutput_({ ok: false, message: 'Invalid amount.' });
  try {
    var payload = 'amount=' + amount
      + '&currency=usd'
      + '&description=' + encodeURIComponent('FreeMoneyDot - ' + description)
      + '&metadata[cashapp]=' + encodeURIComponent(cashapp)
      + '&automatic_payment_methods[enabled]=true'
      + '&automatic_payment_methods[allow_redirects]=never';
    var response = UrlFetchApp.fetch('https://api.stripe.com/v1/payment_intents', {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + STRIPE_SECRET_KEY, 'Content-Type': 'application/x-www-form-urlencoded' },
      payload: payload,
      muteHttpExceptions: true
    });
    var code = response.getResponseCode();
    var result = JSON.parse(response.getContentText());
    if (code === 200 && result.client_secret) {
      getSheet_('StripePayments').appendRow([Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), result.id, cashapp, '$' + (amount / 100).toFixed(2), description, 'pending']);
      return jsonOutput_({ ok: true, clientSecret: result.client_secret, intentId: result.id });
    } else {
      Logger.log('Stripe error: ' + response.getContentText());
      return jsonOutput_({ ok: false, message: result.error ? result.error.message : 'Payment setup failed.' });
    }
  } catch(err) {
    Logger.log('Stripe createIntent error: ' + String(err));
    return jsonOutput_({ ok: false, message: 'Could not connect to payment processor.' });
  }
}

function markStripePaymentConfirmed_(intentId) {
  try {
    var sheet = getSheet_('StripePayments');
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (cleanString_(rows[i][1]) === intentId) {
        sheet.getRange(i + 1, 6).setValue('confirmed');
        return true;
      }
    }
  } catch(e) {}
  return false;
}

function getStripePayments_(p) {
  if (!isAdmin_(p.cashapp)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
  var rows = getDataRows_(getSheet_('StripePayments'));
  var payments = rows.slice(-50).reverse().map(function(r) {
    return { timestamp: cleanString_(r[0]), intentId: cleanString_(r[1]), cashapp: cleanString_(r[2]), amount: cleanString_(r[3]), description: cleanString_(r[4]), status: cleanString_(r[5]) };
  });
  var total = rows.filter(function(r) { return cleanString_(r[5]) === 'confirmed'; }).reduce(function(sum, r) { return sum + (parseFloat(cleanString_(r[3]).replace('$','')) || 0); }, 0);
  return jsonOutput_({ ok: true, payments: payments, totalCollected: total.toFixed(2) });
}

// -- OneSignal Push -----------------------------------------------------------

function broadcastPush_(p) {
  if (!isAdmin_(p.cashapp)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
  var title = cleanString_(p.title || 'FreeMoneyDot');
  var message = cleanString_(p.body || '');
  var url = cleanString_(p.url || 'https://freemoneydot.com/Dev/');
  if (!message) return jsonOutput_({ ok: false, message: 'Message required.' });
  return sendOneSignalPush_(title, message, url, null);
}

function sendPushToMember_(cashapp, title, message, url) {
  sendOneSignalPush_(title, message, url || 'https://freemoneydot.com/Dev/', cashapp);
}

function notifyDrawResult_(num1, payout1, num2, payout2) {
  sendOneSignalPush_('Flat Spin Flat Spin Results!', '1st: #' + num1 + ' wins ' + payout1 + '  |  2nd: #' + num2 + ' wins ' + payout2, 'https://freemoneydot.com/Dev/', null);
}

function notifyBetGraded_(cashapp, pick, result, payout) {
  var emoji = result === 'Won' ? '[WIN]' : result === 'Lost' ? '[X]' : '';
  var msg = pick + (result === 'Won' && payout ? ' - Payout: ' + payout : '');
  sendPushToMember_(cashapp, emoji + ' Bet ' + result + '!', msg, 'https://freemoneydot.com/Dev/');
}

function sendOneSignalPush_(title, message, url, targetCashapp) {
  var payload = {
    app_id: ONESIGNAL_APP_ID,
    headings: { en: title },
    contents: { en: message },
    url: url,
    chrome_web_icon: 'https://freemoneydot.com/Dev/icon-192v2.png'
  };
  if (targetCashapp) {
    payload.include_external_user_ids = [targetCashapp.toLowerCase().replace('$', '')];
  } else {
    payload.included_segments = ['All'];
  }
  try {
    var response = UrlFetchApp.fetch('https://onesignal.com/api/v1/notifications', {
      method: 'post',
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Key ' + ONESIGNAL_API_KEY },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    var code = response.getResponseCode();
    var result = JSON.parse(response.getContentText());
    if (code === 200 && result.id) {
      return jsonOutput_({ ok: true, sent: result.recipients || 0, id: result.id });
    } else {
      Logger.log('OneSignal error: ' + response.getContentText());
      return jsonOutput_({ ok: false, message: JSON.stringify(result.errors || result) });
    }
  } catch(err) {
    Logger.log('Push failed: ' + String(err));
    return jsonOutput_({ ok: false, message: String(err) });
  }
}

// Stubs so nothing breaks
function savePushSubscription_(p) { return jsonOutput_({ ok: true }); }
function subscribePush_(p)        { return jsonOutput_({ ok: true }); }
function unsubscribePush_(p)      { return jsonOutput_({ ok: true }); }

// -- Misc helpers -------------------------------------------------------------

function debugEntries_() {
  const sheet = getSheet_('FlatSpinEntries');
  const rows = getDataRows_(sheet);
  const today = todayEastern_();
  const debug = rows.map(function(row) {
    var raw = row[4];
    var cleaned = cleanDate_(raw);
    return { name: cleanString_(row[1]), rawDate: String(raw), rawType: typeof raw, cleanedDate: cleaned, todayDate: today, matches: cleaned === today };
  });
  return jsonOutput_({ ok: true, today: today, rows: debug });
}

function autoGradeTrigger() { autoGradeBetsFromAPI_(); }
function dailyReport() { autoGradeBetsFromAPI_(); }

// -- Router -------------------------------------------------------------------

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'ping') {
    return ContentService.createTextOutput('{"ok":true,"version":"V40"}').setMimeType(ContentService.MimeType.JSON);
  }
  try {
    const action = cleanString_(e && e.parameter ? e.parameter.action : '');
    if (action === 'flatspindata')        return getFlatSpinData_();
    if (action === 'drawflatspin')        return drawFlatSpinWinners_(cleanString_(e.parameter.paidOnly) === 'true');
    if (action === 'memberstats')         return getMemberStats_(e.parameter);
    if (action === 'bethistory')          return getBetHistory_(e.parameter);
    if (action === 'allbets')             return getAllBets_(e.parameter);
    if (action === 'pokergames')          return getPokerGames_();
    if (action === 'createvipcode')       return createVIPCode_(e.parameter);
    if (action === 'createpokergame')     return createPokerGame_(e.parameter);
    if (action === 'confirmpokerseat')    return confirmPokerSeat_(e.parameter);
    if (action === 'gradebet')            return gradeBet_(e.parameter);
    if (action === 'setmemberrole')       return setMemberRole_(e.parameter);
    if (action === 'setmemberstatus')     return setMemberStatus_(e.parameter);
    if (action === 'markentrypaid')       return markEntryPaid_(e.parameter);
    if (action === 'removeentry')         return removeEntry_(e.parameter);
    if (action === 'autograde')           return autoGradeBetsFromAPI_();
    if (action === 'triggerrollover')     return triggerRollover_();
    if (action === 'rolloverinfo')        { var ri = getActiveRollover_(); ri.ok = true; return jsonOutput_(ri); }
    if (action === 'broadcast')           return getActiveBroadcast_();
    if (action === 'createbroadcast')     return createBroadcast_(e.parameter);
    if (action === 'clearbroadcast')      return clearBroadcast_(e.parameter);
    if (action === 'payouthistory')       return getPayoutHistory_(e.parameter);
    if (action === 'allpayouts')          return getAllPayouts_(e.parameter);
    if (action === 'allpayoutsadmin')     return getAllPayouts_(e.parameter);
    if (action === 'recordpayout')        return recordPayout_(e.parameter);
    if (action === 'referralstats')       return getReferralStats_(e.parameter);
    if (action === 'referralcommissions') return getReferralCommissions_(e.parameter);
    if (action === 'applyreferralcredits') return applyReferralCredits_(e.parameter);
    if (action === 'walletticker')        return getBetTicker_();
    if (action === 'wallethistory')       return getWalletHistory_(e.parameter);
    if (action === 'pendingdeposits')     return getPendingDeposits_(e.parameter);
    if (action === 'approvedeposit')      return approveDeposit_(e.parameter);
    if (action === 'adminwalletcredit')   return adminCreditWallet_(e.parameter);
    if (action === 'schedule')            return getSchedule_();
    if (action === 'createevent')         return createScheduleEvent_(e.parameter);
    if (action === 'deleteevent')         return deleteScheduleEvent_(e.parameter);
    if (action === 'vipwaitlist')         return getVIPWaitlist_(e.parameter);
    if (action === 'approvevip')          return approveVIPWaitlist_(e.parameter);
    if (action === 'rejectvip')           return rejectVIPWaitlist_(e.parameter);
    if (action === 'analytics')           return getAnalytics_(e.parameter);
    if (action === 'dailyreport')         return getDailyReport_(e.parameter);
    if (action === 'allmembers')          return getAllMembers_(e.parameter);
    if (action === 'spadesmatches')       return getSpadesMatches_();
    if (action === 'spadesteams')         return getSpadesTeams_();
    if (action === 'spadesbets')          return getPropBets_();
    if (action === 'quickdrawdata')       return getQuickDrawData_();
    if (action === 'iphonedata')          return getRaffleData_('IphoneRaffle', 'IphoneRaffleResults', 'starry');
    if (action === 'commissarydata')      return getCommissaryRaffleData_();
    if (action === 'dominosdata')         return getDominosData_();
    if (action === 'broadcastpush')       return broadcastPush_(e.parameter);
    if (action === 'unsubscribepush')     return unsubscribePush_(e.parameter);
    if (action === 'stripepayments')      return getStripePayments_(e.parameter);
    if (action === 'debugentries')        return debugEntries_();
    if (action === 'zebrawinners')        return getZebraWinners_();
    if (action === 'betboard')            return getBetBoard_();
    if (action === 'dmthreads')           return getDMThreads_(e.parameter);
    if (action === 'dmthread')            return getDMThread_(e.parameter);
    if (action === 'adminentries')        return getAdminEntries_(e.parameter);
    if (action === 'adminbets')           return getAdminPendingBets_();
    if (action === 'adminmembers')        return getAllMembers_(e.parameter);
    if (action === 'admindeposits')       return getPendingDeposits_(e.parameter);
    if (action === 'eventvisibility')     return getEventVisibility_();
    if (action === 'seteventvisibility')  return setEventVisibility_(e.parameter);
    if (action === 'setbroadcast')        return createBroadcast_(e.parameter);
    if (action === 'markpaid')            return markEntryPaid_(e.parameter);
    if (action === 'setrole')             return setMemberRole_(e.parameter);
    if (action === 'setstatus')           return setMemberStatus_(e.parameter);
    if (action === 'generatevipcode')     return createVIPCode_(e.parameter);
    if (action === 'rejectdeposit')       return rejectDeposit_(e.parameter);
    if (action === 'referralcredits')     return getReferralCommissions_(e.parameter);
    if (action === 'drawquickdraw')       return drawQuickDrawWinners_(cleanString_(e.parameter.paidOnly) === 'true');
    if (action === 'drawcommissaryraffle') return drawCommissaryRaffleWinners_(cleanString_(e.parameter.paidOnly) === 'true');
    // --- Pager routes (no auth required for status read) ---
if (action === 'dlstandings')         return dlGetStandings_();
    if (action === 'dlschedule')          return dlGetSchedule_();
    if (action === 'dlresults')           return dlGetResults_();
    if (action === 'dlplayoffs')          return dlGetPlayoffs_();
        if (action === 'pagerstatus')    return pagerGetStatus_();
    if (action === 'pagertoggle')    return pagerToggle_(e.parameter);
    if (action === 'pageremergency') return pagerEmergency_(e.parameter);
    // --- Added GET routes ---
    if (action === 'drawiphoneraffle') return drawRaffleWinner_('IphoneRaffle', 'IphoneRaffleResults', '5-pack Starry', cleanString_(e.parameter.paidOnly) === 'true', 1);
    return jsonOutput_({ ok: false, message: 'Unknown action: ' + action });
  } catch (err) {
    return jsonOutput_({ ok: false, message: String(err) });
  }
}

function doPost(e) {
  try {
    const p = e && e.parameter ? e.parameter : {};
    const formType = cleanString_(p.formType || p.action || '').toLowerCase();
    if (formType === 'flatspin')             return submitFlatSpinEntry_(p);
    if (formType === 'spadessignup')         return submitSpadesSignup_(p);
    if (formType === 'spadessidebet')        return submitPropBet_(p);
    if (formType === 'propbet')              return submitPropBet_(p);
    if (formType === 'quickdraw')            return submitQuickDrawEntry_(p);
    if (formType === 'iphoneraffle')         return submitRaffleEntry_(p, 'IphoneRaffle');
    if (formType === 'commissaryraffle')     return submitCommissaryEntry_(p);
    if (formType === 'dominossignup')        return submitDominosSignup_(p);
    if (formType === 'dominosprop')          return submitDominosProp_(p);
    if (formType === 'cookingentry')         return submitCookingEntry_(p);
    if (formType === 'cookingprop')          return submitCookingProp_(p);
    if (formType === 'spinprop')             return submitSpinPropBet_(p);
    if (formType === 'submitbet')            return submitBetWithWalletCheck_(p);
    if (formType === 'reserveseat')          return reservePokerSeat_(p);
    if (formType === 'redeemvip')            return redeemVIPCode_(p);
    if (formType === 'register')             return registerMember_(p);
    if (formType === 'login')                return loginMember_(p);
    if (formType === 'depositrequest')       return requestDeposit_(p);
    if (formType === 'vipwaitlist')          return applyVIPWaitlist_(p);
    if (formType === 'createpaymentintent')  return createPaymentIntent_(p);
    if (formType === 'savepushsub')          return savePushSubscription_(p);
    if (formType === 'pushsubscribe')        return savePushSubscription_(p);
    if (formType === 'stripewallet')          return stripeWalletDeposit_(p);
    if (formType === 'subscribepush')        return subscribePush_(p);
    if (formType === 'betboard')             return postToBetBoard_(p);
    if (formType === 'boardreact')           return reactToBetBoardPost_(p);
    if (formType === 'sendmessage')          return sendDirectMessage_(p);
    if (formType === 'resetpin')             return adminResetMemberPin_(p);
    if (formType === 'walletadjust')         return adminWalletAdjust_(p);
    if (formType === 'recordpayout')         return recordPayout_(p);
    if (formType === 'applyreferralcredits') return applyReferralCredits_(p);
    if (formType === 'createpoker')          return createPokerGame_(p);
    if (formType === 'affiliate')            return submitAffiliateApp_(p);
    // --- Added routes (aliases + new features) ---
    if (formType === 'broadcast')            return createBroadcast_(p);
    if (formType === 'broadcastpush')        return broadcastPushFromPost_(p);
    if (formType === 'affiliateapply')       return submitAffiliateApp_(p);
    if (formType === 'reservepokerseat')     return reservePokerSeat_(p);
    if (formType === 'clearbroadcast')       return clearBroadcast_(p);
    if (formType === 'pagertoggle')          return pagerToggle_(p);
    if (formType === 'pageremergency')       return pagerEmergency_(p);
    if (formType === 'approvemember')       return approveRejectMember_(p);
    if (formType === 'dladdteam')          return dlAddTeam_(p);
    if (formType === 'dlresult')           return dlRecordResult_(p);
    if (formType === 'dlgenschedule')      return dlGenerateSchedule_(p);
    return jsonOutput_({ ok: false, message: 'Unknown form type. Got formType="' + (p.formType||'') + '" action="' + (p.action||'') + '" keys=' + Object.keys(p).join(',') });
  } catch (err) {
    return jsonOutput_({ ok: false, message: String(err) });
  }
}

// -- Stripe Wallet Deposit ----------------------------------------------------
// Called after Stripe payment succeeds on frontend
// Automatically credits wallet - no admin approval needed

function stripeWalletDeposit_(p) {
  var cashapp  = cleanString_(p.cashapp || '');
  var amount   = parseFloat(p.amount || '0');
  var intentId = cleanString_(p.intentId || '');

  if (!cashapp || !amount || !intentId) {
    return jsonOutput_({ ok: false, message: 'Missing required fields.' });
  }

  var member = getMember_(cashapp);
  if (!member) return jsonOutput_({ ok: false, message: 'Member not found.' });

  // Check this intent hasn't already been credited (prevent double credit)
  var walletRows = getDataRows_(getSheet_('Wallet'));
  for (var i = 0; i < walletRows.length; i++) {
    if (cleanString_(walletRows[i][4]) === intentId) {
      return jsonOutput_({ ok: false, message: 'This payment has already been credited.' });
    }
  }

  // Mark payment as confirmed in StripePayments sheet
  markStripePaymentConfirmed_(intentId);

  // Credit the wallet immediately - no admin approval needed
  var note = 'Stripe deposit - ' + intentId;
  getSheet_('Wallet').appendRow([
    member.cashapp,
    'deposit',
    amount,
    'approved',
    intentId,  // store intentId in note field for duplicate check
    Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss')
  ]);

  // Log to StripePayments too
  getSheet_('StripePayments').appendRow([
    Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'),
    intentId,
    member.cashapp,
    '$' + amount.toFixed(2),
    'Wallet deposit',
    'confirmed'
  ]);

  var newBalance = getWalletBalance_(cashapp);

  return jsonOutput_({
    ok: true,
    message: '$' + amount.toFixed(2) + ' added to your wallet!',
    newBalance: newBalance
  });
}

// ===================================================================
// v17 NEW FEATURES: Bet Board, DMs, Reset PIN, Affiliate
// ===================================================================

/** Bet Board: load all posts (newest first) */
function getBetBoard_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName('BetBoard');
  if (!sh) {
    sh = ss.insertSheet('BetBoard');
    sh.appendRow(['ID', 'Timestamp', 'Cashapp', 'Name', 'Pick', 'Game', 'Amount', 'Fire', 'Money', 'Lock', 'Reactors']);
  }
  var data = sh.getDataRange().getValues();
  var posts = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    posts.push({
      id: String(row[0]),
      timestamp: row[1] ? formatDate_(row[1]) : '',
      cashapp: row[2] || '',
      name: row[3] || '',
      pick: row[4] || '',
      game: row[5] || '',
      amount: row[6] || '',
      fire: Number(row[7]) || 0,
      money: Number(row[8]) || 0,
      lock: Number(row[9]) || 0
    });
  }
  posts.reverse();
  return jsonOutput_({ ok: true, posts: posts.slice(0, 50) });
}

function postToBetBoard_(p) {
  var cashapp = cleanString_(p.cashapp);
  var name = cleanString_(p.name);
  var pick = cleanString_(p.pick);
  if (!cashapp || !pick) return jsonOutput_({ ok: false, message: 'Cashapp and pick required.' });
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName('BetBoard');
  if (!sh) {
    sh = ss.insertSheet('BetBoard');
    sh.appendRow(['ID', 'Timestamp', 'Cashapp', 'Name', 'Pick', 'Game', 'Amount', 'Fire', 'Money', 'Lock', 'Reactors']);
  }
  var id = 'BB' + Date.now();
  sh.appendRow([id, new Date(), cashapp, name, pick, cleanString_(p.game), cleanString_(p.amount), 0, 0, 0, '']);
  return jsonOutput_({ ok: true, message: 'Pick posted!', id: id });
}

function reactToBetBoardPost_(p) {
  var postId = cleanString_(p.postId);
  var reaction = cleanString_(p.reaction).toLowerCase();
  var cashapp = cleanString_(p.cashapp);
  if (!postId || !reaction || !cashapp) return jsonOutput_({ ok: false, message: 'Missing fields.' });
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName('BetBoard');
  if (!sh) return jsonOutput_({ ok: false, message: 'No bet board.' });
  var data = sh.getDataRange().getValues();
  var colMap = { fire: 7, money: 8, lock: 9 };
  if (colMap[reaction] === undefined) return jsonOutput_({ ok: false, message: 'Invalid reaction.' });
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === postId) {
      // Dedupe: check reactors column
      var reactors = String(data[i][10] || '');
      var key = cashapp + ':' + reaction;
      if (reactors.indexOf(key) !== -1) {
        return jsonOutput_({ ok: true, message: 'Already reacted.' });
      }
      var current = Number(data[i][colMap[reaction]]) || 0;
      sh.getRange(i + 1, colMap[reaction] + 1).setValue(current + 1);
      sh.getRange(i + 1, 11).setValue(reactors + key + ',');
      return jsonOutput_({ ok: true });
    }
  }
  return jsonOutput_({ ok: false, message: 'Post not found.' });
}

/** DMs: thread list for a user */
function getDMThreads_(params) {
  var cashapp = cleanString_(params.cashapp);
  if (!cashapp) return jsonOutput_({ ok: false, message: 'No user.' });
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName('DirectMessages');
  if (!sh) {
    sh = ss.insertSheet('DirectMessages');
    sh.appendRow(['ID', 'Timestamp', 'FromCashapp', 'ToCashapp', 'ThreadID', 'Message', 'Read']);
  }
  var data = sh.getDataRange().getValues();
  var threads = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var from = row[2], to = row[3], threadId = row[4], msg = row[5];
    if (from !== cashapp && to !== cashapp) continue;
    var other = from === cashapp ? to : from;
    if (!threads[other]) {
      threads[other] = {
        threadId: threadId || [cashapp, other].sort().join('|'),
        otherCashapp: other,
        lastMessage: msg || '',
        timestamp: row[1] ? formatDate_(row[1]) : ''
      };
    } else {
      // newer timestamp wins
      threads[other].lastMessage = msg || '';
      threads[other].timestamp = row[1] ? formatDate_(row[1]) : '';
    }
  }
  var arr = [];
  Object.keys(threads).forEach(function(k) { arr.push(threads[k]); });
  return jsonOutput_({ ok: true, threads: arr });
}

function getDMThread_(params) {
  var me = cleanString_(params.cashapp);
  var other = cleanString_(params['with']);
  if (!me || !other) return jsonOutput_({ ok: false, message: 'Need both parties.' });
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName('DirectMessages');
  if (!sh) return jsonOutput_({ ok: true, messages: [] });
  var data = sh.getDataRange().getValues();
  var msgs = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var from = row[2], to = row[3];
    if ((from === me && to === other) || (from === other && to === me)) {
      msgs.push({
        from: from,
        to: to,
        message: row[5] || '',
        timestamp: row[1] ? formatDate_(row[1]) : ''
      });
    }
  }
  return jsonOutput_({ ok: true, messages: msgs });
}

function sendDirectMessage_(p) {
  var from = cleanString_(p.fromCashapp);
  var to = cleanString_(p.toCashapp);
  var message = cleanString_(p.message);
  if (!from || !to || !message) return jsonOutput_({ ok: false, message: 'Missing fields.' });
  if (from === to) return jsonOutput_({ ok: false, message: "Can't message yourself." });
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName('DirectMessages');
  if (!sh) {
    sh = ss.insertSheet('DirectMessages');
    sh.appendRow(['ID', 'Timestamp', 'FromCashapp', 'ToCashapp', 'ThreadID', 'Message', 'Read']);
  }
  var threadId = [from, to].sort().join('|');
  sh.appendRow(['DM' + Date.now(), new Date(), from, to, threadId, message, false]);
  // Optional: push notify recipient
  try { sendPushToMember_(to, 'New message from ' + from, message.substring(0, 60)); } catch (e) {}
    try {
    sendPushToUser_(cleanString_(p.to||''), 'New message from ' + cleanString_(p.from||''), String(p.message||'').substring(0,120), '');
  } catch(_e){}

  return jsonOutput_({ ok: true, message: 'Sent.' });
}

/** Admin Reset Member PIN */
function adminResetMemberPin_(p) {
  var adminCashapp = cleanString_(p.adminCashapp);
  var targetCashapp = cleanString_(p.cashapp || p.targetCashapp);
  var newPin = cleanString_(p.newPin);
  if (!isAdmin_(adminCashapp)) return jsonOutput_({ ok: false, message: 'Not authorized.' });
  if (!targetCashapp || !/^\d{4}$/.test(newPin)) return jsonOutput_({ ok: false, message: 'Need cashapp + 4-digit PIN.' });
  // Normalize: strip leading $ for comparison
  var normTarget = targetCashapp.replace('$', '').toLowerCase();
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName('Members');
  if (!sh) return jsonOutput_({ ok: false, message: 'No Members sheet.' });
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var cashCol = headers.indexOf('Cashapp');
  if (cashCol === -1) cashCol = 2; // fallback: column C
  var pinCol = headers.indexOf('Pin');
  if (pinCol === -1) pinCol = headers.indexOf('PinHash');
  if (pinCol === -1) pinCol = 3; // fallback: column D
  for (var i = 1; i < data.length; i++) {
    var rowCashapp = String(data[i][cashCol]).replace('$', '').toLowerCase();
    if (rowCashapp === normTarget) {
      sh.getRange(i + 1, pinCol + 1).setValue(newPin);
      return jsonOutput_({ ok: true, message: 'PIN reset for ' + targetCashapp });
    }
  }
  return jsonOutput_({ ok: false, message: 'Member not found.' });
}

/** Admin entries loader (dispatches by sheet name) */
function getAdminEntries_(params) {
  var sheet = cleanString_(params.sheet);
  if (!sheet) return jsonOutput_({ ok: false, message: 'No sheet.' });
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(sheet);
  if (!sh) return jsonOutput_({ ok: true, entries: [] });
  var data = sh.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h); });
  var entries = [];
  for (var i = 1; i < data.length; i++) {
    var e = {};
    for (var j = 0; j < headers.length; j++) {
      var k = headers[j].charAt(0).toLowerCase() + headers[j].slice(1);
      e[k] = data[i][j];
    }
    entries.push(e);
  }
  return jsonOutput_({ ok: true, entries: entries });
}

function getAdminPendingBets_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName('SportsBets');
  if (!sh) return jsonOutput_({ ok: true, bets: [] });
  var data = sh.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h); });
  var statusCol = headers.indexOf('Status');
  var bets = [];
  for (var i = 1; i < data.length; i++) {
    if (statusCol >= 0 && String(data[i][statusCol]).toLowerCase() !== 'pending') continue;
    var b = { id: data[i][0] };
    for (var j = 0; j < headers.length; j++) {
      var k = headers[j].charAt(0).toLowerCase() + headers[j].slice(1);
      b[k] = data[i][j];
    }
    bets.push(b);
  }
  return jsonOutput_({ ok: true, bets: bets });
}

/** Event visibility toggles stored in Settings sheet */
function getEventVisibility_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName('Settings');
  if (!sh) {
    sh = ss.insertSheet('Settings');
    sh.appendRow(['Key', 'Value']);
    sh.appendRow(['event_qd_visible', 'false']);
    sh.appendRow(['event_starry_visible', 'false']);
    sh.appendRow(['event_comm_visible', 'false']);
  }
  var data = sh.getDataRange().getValues();
  var out = { qd: false, starry: false, comm: false };
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === 'event_qd_visible') out.qd = String(data[i][1]) === 'true';
    if (data[i][0] === 'event_starry_visible') out.starry = String(data[i][1]) === 'true';
    if (data[i][0] === 'event_comm_visible') out.comm = String(data[i][1]) === 'true';
  }
  out.ok = true;
  return jsonOutput_(out);
}

function setEventVisibility_(params) {
  var key = cleanString_(params.key);
  var value = cleanString_(params.value);
  if (!key) return jsonOutput_({ ok: false, message: 'No key.' });
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName('Settings');
  if (!sh) {
    sh = ss.insertSheet('Settings');
    sh.appendRow(['Key', 'Value']);
  }
  var fullKey = 'event_' + key + '_visible';
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === fullKey) {
      sh.getRange(i + 1, 2).setValue(value);
      return jsonOutput_({ ok: true });
    }
  }
  sh.appendRow([fullKey, value]);
  return jsonOutput_({ ok: true });
}

/** Admin: manual wallet adjust */
function adminWalletAdjust_(p) {
  var adminCashapp = cleanString_(p.adminCashapp);
  if (!isAdmin_(adminCashapp)) return jsonOutput_({ ok: false, message: 'Not authorized.' });
  var cashapp = cleanString_(p.cashapp);
  var type = cleanString_(p.type).toLowerCase();
  var amount = parseFloat(p.amount) || 0;
  var note = cleanString_(p.note);
  if (!cashapp || amount <= 0) return jsonOutput_({ ok: false, message: 'Cashapp and amount required.' });
  // Normalize cashapp (add $ if missing)
  if (cashapp.charAt(0) !== '$') cashapp = '$' + cashapp;
  // For debit, store as negative; the existing wallet history calculation uses amount strings
  var displayAmount = (type === 'debit') ? '-$' + amount.toFixed(2) : '$' + amount.toFixed(2);
  var ledgerType = (type === 'debit') ? 'debit' : (type === 'win' ? 'win' : 'credit');
  getSheet_('Wallet').appendRow([
    cashapp,
    ledgerType,
    displayAmount,
    'approved',
    note || ('Admin ' + ledgerType),
    Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss')
  ]);
  // Return new balance
  var balResult = getWalletHistory_({ cashapp: cashapp });
  var balObj = {};
  try { balObj = JSON.parse(balResult.getContent()); } catch (e) {}
  return jsonOutput_({ ok: true, message: 'Wallet updated.', newBalance: balObj.balance || 0 });
}

function rejectDeposit_(params) {
  var adminCashapp = cleanString_(params.adminCashapp);
  if (!isAdmin_(adminCashapp)) return jsonOutput_({ ok: false, message: 'Not authorized.' });
  var id = cleanString_(params.id);
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName('WalletDeposits');
  if (!sh) return jsonOutput_({ ok: false, message: 'No deposits sheet.' });
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var idCol = headers.indexOf('ID');
  var statusCol = headers.indexOf('Status');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === id) {
      sh.getRange(i + 1, statusCol + 1).setValue('Rejected');
      return jsonOutput_({ ok: true });
    }
  }
  return jsonOutput_({ ok: false, message: 'Deposit not found.' });
}

/** Affiliate application (simple) */
function submitAffiliateApp_(p) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName('AffiliateApps');
  if (!sh) {
    sh = ss.insertSheet('AffiliateApps');
    sh.appendRow(['Timestamp', 'Name', 'Payment', 'Network', 'Notes', 'Status']);
  }
  sh.appendRow([new Date(), cleanString_(p.name), cleanString_(p.payment), cleanString_(p.network), cleanString_(p.notes), 'Pending']);
  return jsonOutput_({ ok: true, message: 'Application submitted!' });
}

/** Helper: format date consistently */
function formatDate_(d) {
  try {
    return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'MMM d, h:mm a');
  } catch (e) { return String(d); }
}

/** Helper: is user admin? - provided by existing code */

// ========= ADDED: OneSignal push + broadcast helpers =========
//
// Usage: set ONESIGNAL_APP_ID (already at top of this file) and
// ONESIGNAL_API_KEY (get from OneSignal dashboard -> Keys & IDs).
// Player IDs per member should be saved via the existing savePushSubscription_
// when the PWA registers a OneSignal user.
//
// Members sheet should have a "PushPlayerId" column. If not, the helpers
// degrade gracefully and log instead.

function fmd_getOneSignalKey_() {
  try {
    var key = (typeof ONESIGNAL_API_KEY !== 'undefined') ? ONESIGNAL_API_KEY : '';
    if (!key || key === 'YOUR_ONESIGNAL_REST_API_KEY') return '';
    return key;
  } catch (e) { return ''; }
}

function sendPushToPlayerIds_(playerIds, title, message, url) {
  var key = fmd_getOneSignalKey_();
  if (!key || !playerIds || !playerIds.length) return false;
  var payload = {
    app_id: ONESIGNAL_APP_ID,
    include_player_ids: playerIds,
    headings: { en: String(title || 'FreeMoneyDot') },
    contents: { en: String(message || '') }
  };
  if (url) payload.url = url;
  try {
    UrlFetchApp.fetch('https://onesignal.com/api/v1/notifications', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Key ' + key },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    return true;
  } catch (e) {
    Logger.log('OneSignal push failed: ' + e);
    return false;
  }
}

function sendPushToAll_(title, message, url) {
  var key = fmd_getOneSignalKey_();
  if (!key) return false;
  var payload = {
    app_id: ONESIGNAL_APP_ID,
    included_segments: ['All'],
    headings: { en: String(title || 'FreeMoneyDot') },
    contents: { en: String(message || '') }
  };
  if (url) payload.url = url;
  try {
    UrlFetchApp.fetch('https://onesignal.com/api/v1/notifications', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Key ' + key },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    return true;
  } catch (e) {
    Logger.log('OneSignal broadcast push failed: ' + e);
    return false;
  }
}

// Look up a member's OneSignal player ID by cashapp.
// Expects a 'PushPlayerId' column on Members sheet; returns '' if not found.
function getPushPlayerId_(cashapp) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sh = ss.getSheetByName('Members');
    if (!sh) return '';
    var data = sh.getDataRange().getValues();
    if (!data.length) return '';
    var headers = data[0].map(function(h){ return String(h).toLowerCase(); });
    var cashappCol = headers.indexOf('cashapp');
    var pidCol     = headers.indexOf('pushplayerid');
    if (cashappCol < 0 || pidCol < 0) return '';
    var target = String(cashapp).replace(/^\$+/, '').toLowerCase();
    for (var i = 1; i < data.length; i++) {
      var v = String(data[i][cashappCol] || '').replace(/^\$+/, '').toLowerCase();
      if (v === target) return String(data[i][pidCol] || '');
    }
  } catch (e) {
    Logger.log('getPushPlayerId_ err: ' + e);
  }
  return '';
}

function sendPushToUser_(cashapp, title, message, url) {
  var pid = getPushPlayerId_(cashapp);
  if (!pid) return false;
  return sendPushToPlayerIds_([pid], title, message, url);
}

// POST-compatible wrapper for broadcastpush formType
function broadcastPushFromPost_(p) {
  if (!isAdmin_(p.adminCashapp || p.cashapp)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
  var title   = cleanString_(p.title   || 'FreeMoneyDot');
  var message = cleanString_(p.message || p.body || '');
  if (!message) return jsonOutput_({ ok: false, message: 'Message required.' });
  var ok = sendPushToAll_(title, message);
  return jsonOutput_({ ok: !!ok, message: ok ? 'Broadcast push sent.' : 'Push not configured.' });
}
// ========= end added helpers =========

// ========= PAGER SYSTEM (Google Apps Script native - no PHP, no Telnyx) =========
//
// State stored in "PagerState" sheet:
//   col 0: key (rec, fence, unit_1, unit_2, unit_3, unit_4)
//   col 1: active (TRUE/FALSE)
//   col 2: updatedBy (cashapp)
//   col 3: updatedAt (timestamp)
//
// Alerts delivered via OneSignal push to all subscribed users.

var PAGER_KEYS   = ['rec','fence','unit_1','unit_2','unit_3','unit_4'];
var PAGER_LABELS = {rec:'REC',fence:'FENCE',unit_1:'Unit 1',unit_2:'Unit 2',unit_3:'Unit 3',unit_4:'Unit 4'};

function pagerGetSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName('PagerState');
  if (!sh) {
    sh = ss.insertSheet('PagerState');
    sh.appendRow(['Key','Active','UpdatedBy','UpdatedAt']);
    PAGER_KEYS.forEach(function(k) {
      sh.appendRow([k, false, '', '']);
    });
  }
  return sh;
}

function pagerReadState_() {
  var sh = pagerGetSheet_();
  var rows = sh.getDataRange().getValues();
  var state = { states: {}, updated_times: {}, updated_by: {}, last_updated: null };
  // Skip header row
  for (var i = 1; i < rows.length; i++) {
    var key = String(rows[i][0]);
    var active = rows[i][1] === true || rows[i][1] === 'TRUE' || rows[i][1] === 1;
    var updatedAt = rows[i][3] ? String(rows[i][3]) : null;
    state.states[key] = active;
    state.updated_times[key] = updatedAt;
    state.updated_by[key] = String(rows[i][2] || '');
    if (updatedAt && (!state.last_updated || updatedAt > state.last_updated)) {
      state.last_updated = updatedAt;
    }
  }
  return state;
}

function pagerWriteKey_(key, active, cashapp) {
  var sh = pagerGetSheet_();
  var rows = sh.getDataRange().getValues();
  var ts = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === key) {
      sh.getRange(i + 1, 2).setValue(active);
      sh.getRange(i + 1, 3).setValue(cashapp || '');
      sh.getRange(i + 1, 4).setValue(ts);
      return ts;
    }
  }
  // Key row not found - append it
  sh.appendRow([key, active, cashapp || '', ts]);
  return ts;
}

function pagerGetStatus_() {
  try {
    var state = pagerReadState_();
    return jsonOutput_({ ok: true, success: true, state: state });
  } catch(err) {
    return jsonOutput_({ ok: false, message: String(err) });
  }
}

function pagerToggle_(p) {
  var cashapp = cleanString_(p.cashapp || p.adminCashapp || '');
  if (!isAdmin_(cashapp)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
  var type = cleanString_(p.type || '');
  if (PAGER_KEYS.indexOf(type) < 0) return jsonOutput_({ ok: false, message: 'Invalid pager type.' });
  try {
    var state = pagerReadState_();
    var currentlyActive = !!(state.states[type]);
    var newActive = !currentlyActive;
    // Units are mutually exclusive - turning one on clears the others
    var isUnit = type.indexOf('unit_') === 0;
    if (isUnit && newActive) {
      ['unit_1','unit_2','unit_3','unit_4'].forEach(function(u) {
        if (u !== type) pagerWriteKey_(u, false, cashapp);
      });
    }
    pagerWriteKey_(type, newActive, cashapp);
    var label = PAGER_LABELS[type] || type;
    var statusWord = newActive ? 'ACTIVE' : 'CLEAR';
    var pushMsg = label + ' - ' + statusWord;
    try { sendPushToAll_('Pager Update', pushMsg, ''); } catch(e) {}
    var updatedState = pagerReadState_();
    return jsonOutput_({ ok: true, state: updatedState, message: pushMsg, active: newActive });
  } catch(err) {
    return jsonOutput_({ ok: false, message: String(err) });
  }
}

function pagerEmergency_(p) {
  var cashapp = cleanString_(p.cashapp || p.adminCashapp || '');
  if (!isAdmin_(cashapp)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
  var message = cleanString_(p.message || '');
  if (!message) return jsonOutput_({ ok: false, message: 'Message required.' });
  try {
    try { sendPushToAll_('EMERGENCY EMERGENCY', message, ''); } catch(e) {}
    // Log to a PagerLogs sheet for audit trail
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var logSh = ss.getSheetByName('PagerLogs') || ss.insertSheet('PagerLogs');
    if (logSh.getLastRow() === 0) logSh.appendRow(['Timestamp','Type','Message','SentBy']);
    logSh.appendRow([
      Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'),
      'emergency', message, cashapp
    ]);
    return jsonOutput_({ ok: true, message: 'Emergency alert sent to all via push.' });
  } catch(err) {
    return jsonOutput_({ ok: false, message: String(err) });
  }
}
// ========= end pager system =========

// ===== TEST FUNCTION - run this manually in Apps Script editor =====
function testOneSignalPush() {
  var payload = {
    app_id: ONESIGNAL_APP_ID,
    included_segments: ['All'],
    headings: { en: 'Test from FMD' },
    contents: { en: 'Push working!' }
  };
  var formats = [
    { label: 'Key prefix',    auth: 'Key '    + ONESIGNAL_API_KEY },
    { label: 'Basic prefix',  auth: 'Basic '  + ONESIGNAL_API_KEY },
    { label: 'Bearer prefix', auth: 'Bearer ' + ONESIGNAL_API_KEY },
    { label: 'Raw key',       auth: ONESIGNAL_API_KEY }
  ];
  formats.forEach(function(fmt) {
    try {
      var r = UrlFetchApp.fetch('https://onesignal.com/api/v1/notifications', {
        method: 'post',
        contentType: 'application/json',
        headers: { 'Authorization': fmt.auth },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      Logger.log(fmt.label + ' -> ' + r.getResponseCode() + ' | ' + r.getContentText().substring(0, 120));
    } catch(e) {
      Logger.log(fmt.label + ' -> ERROR: ' + e);
    }
  });
}
// ===== END TEST FUNCTION =====

// ===== ACCOUNT APPROVAL SYSTEM =====

function autoApproveExistingMembers() {
  // Run this once manually to approve all existing members
  var sh = getSheet_('Members');
  var data = sh.getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    var status = cleanString_(String(data[i][6]||''));
    if (status === 'active' || status === '' || status === 'Active') {
      data[i][6] = 'approved';
      count++;
    }
  }
  sh.getRange(1,1,data.length,data[0].length).setValues(data);
  Logger.log('Auto-approved ' + count + ' existing members.');
  return count;
}

function getPendingMembers_(p) {
  try {
    var adminCa = cleanString_(p.adminCashapp || '');
    if (!isAdmin_(adminCa)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
    var rows = getDataRows_(getSheet_('Members'));
    var pending = rows.filter(function(r){ return cleanString_(String(r[6]||'')).toLowerCase() === 'pending'; })
      .map(function(r){ return { name:cleanString_(r[1]), cashapp:cleanString_(r[2]), joinDate:cleanString_(r[7]||''), referredBy:cleanString_(r[5]||'') }; });
    return jsonOutput_({ ok: true, members: pending });
  } catch(e) { return jsonOutput_({ ok: false, message: String(e) }); }
}

function approveRejectMember_(p) {
  try {
    var adminCa  = cleanString_(p.adminCashapp || '');
    if (!isAdmin_(adminCa)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
    var cashapp  = cleanString_(p.cashapp || '').toLowerCase().replace('$','').replace('@','');
    var newStatus = cleanString_(p.status || ''); // 'approved' or 'rejected'
    if (!cashapp || (newStatus !== 'approved' && newStatus !== 'rejected' && newStatus !== 'suspended')) {
      return jsonOutput_({ ok: false, message: 'Invalid cashapp or status.' });
    }
    var sh = getSheet_('Members');
    var data = sh.getDataRange().getValues();
    var found = false;
    for (var i = 1; i < data.length; i++) {
      var ca = cleanString_(String(data[i][2]||'')).toLowerCase().replace('$','').replace('@','');
      if (ca === cashapp) {
        data[i][6] = newStatus;
        found = true;
        break;
      }
    }
    if (!found) return jsonOutput_({ ok: false, message: 'Member not found.' });
    sh.getRange(1,1,data.length,data[0].length).setValues(data);
    // Send push notification if approved
    if (newStatus === 'approved') {
      try { sendPushToAll_('[APPROVED] Account Approved', 'Your FreeMoneyDot account has been approved! Log in now.', ''); } catch(e) {}
    }
    return jsonOutput_({ ok: true, message: 'Account ' + newStatus + '.' });
  } catch(e) { return jsonOutput_({ ok: false, message: String(e) }); }
}

function getAllMembersAdmin_(p) {
  try {
    var adminCa = cleanString_(p.adminCashapp || '');
    if (!isAdmin_(adminCa)) return jsonOutput_({ ok: false, message: 'Unauthorized.' });
    var rows = getDataRows_(getSheet_('Members'));
    var members = rows.map(function(r){
      return { name:cleanString_(r[1]), cashapp:cleanString_(r[2]), status:cleanString_(String(r[6]||'active')), joinDate:cleanString_(r[7]||''), role:cleanString_(r[8]||'member'), referredBy:cleanString_(r[5]||'') };
    });
    return jsonOutput_({ ok: true, members: members });
  } catch(e) { return jsonOutput_({ ok: false, message: String(e) }); }
}
// ===== END ACCOUNT APPROVAL SYSTEM =====