// 納入台帳テキスト解析 メイン側 v2026.03.06-01

var VERSION_TEXT = '納入台帳テキスト解析 v2026.03.06-01';
var TAX_RATE = 0.08; // 消費税率（8%）

// コピー用保持
var lastHeaderText  = '';
var lastDetailsText = '';
var lastSummaryText = '';
var currentRows     = []; // {vendor,no,name,spec,unit,qty,price,amount,note}

// サマリ（原本値＋計算値）保持
var originalSummary = null;

// -------------------- 数値・文字列ユーティリティ --------------------

// 計算用パース（カンマ・¥・\ を除去）
function parseNumberForCalc(val) {
  if (val === null || val === undefined) return NaN;
  var s = String(val);
  s = s.replace(/[¥\\,]/g, '');
  s = s.replace(/^\s+|\s+$/g, '');
  if (!s) return NaN;
  var num = parseFloat(s);
  if (isNaN(num)) return NaN;
  return num;
}

// 3桁カンマ＋小数2桁
function formatAmount(val) {
  var num = parseFloat(val);
  if (isNaN(num)) return '';
  var fixed = num.toFixed(2);
  var parts = fixed.split('.');
  var intPart = parts[0];
  var decPart = parts[1];
  var re = /(\d+)(\d{3})/;
  while (re.test(intPart)) {
    intPart = intPart.replace(re, '$1' + ',' + '$2');
  }
  return intPart + '.' + decPart;
}

// 3桁カンマ・整数
function formatInt(val) {
  var num = parseInt(Math.floor(val), 10);
  if (isNaN(num)) return '';
  var s = String(num);
  var re = /(\d+)(\d{3})/;
  while (re.test(s)) {
    s = s.replace(re, '$1' + ',' + '$2');
  }
  return s;
}

// yyyy-mm-dd → yyyy/mm/dd
function formatDateForCopy(val) {
  if (!val) return '';
  var p = String(val).split('-');
  if (p.length === 3) {
    return p[0] + '/' + p[1] + '/' + p[2];
  }
  return val;
}

// 業者名行から「納地」「業者名」を分離
// 例: "滝川駐屯地株式会社くみあい食品" → site="滝川駐屯地", vendor="株式会社くみあい食品"
function splitVendorFull(full) {
  var s = (full || '').replace(/^\s+|\s+$/g, '');
  if (!s) {
    return { site: '', vendor: '' };
  }
  var forms = ['株式会社', '有限会社', '合名会社', '合資会社', '合同会社'];
  var i, form, idx;
  for (i = 0; i < forms.length; i++) {
    form = forms[i];
    idx = s.indexOf(form);
    if (idx > 0) {
      return {
        site: s.substring(0, idx),
        vendor: s.substring(idx)
      };
    }
  }
  return { site: '', vendor: s };
}

// 備考マーク操作
function addNoteMark(row, mark) {
  if (!row.note) row.note = '';
  if (row.note.indexOf(mark) === -1) {
    row.note += mark;
  }
}
function removeNoteMark(row, mark) {
  if (!row.note) return;
  var chars = row.note.split('');
  var filtered = [];
  var i;
  for (i = 0; i < chars.length; i++) {
    if (chars[i] !== mark) {
      filtered.push(chars[i]);
    }
  }
  row.note = filtered.join('');
}

// 行ごとの「計算OK（<）」判定
function applyCalcMark(row) {
  var q = parseNumberForCalc(row.qty);
  var p = parseNumberForCalc(row.price);
  var a = parseNumberForCalc(row.amount);
  if (!isNaN(q) && !isNaN(p) && !isNaN(a)) {
    var expected = q * p;
    if (Math.abs(a - expected) < 0.5) {
      addNoteMark(row, '<');
    } else {
      removeNoteMark(row, '<');
    }
  } else {
    removeNoteMark(row, '<');
  }
}

// 備考セルをDOM側で更新
function updateRowNoteCell(rowIndex) {
  var cell = document.querySelector('td[data-note-index="' + rowIndex + '"]');
  if (cell && currentRows[rowIndex]) {
    cell.textContent = currentRows[rowIndex].note || '';
  }
}

// 金額セルをDOM側で更新
function updateRowAmountCell(rowIndex) {
  var cell = document.querySelector('td[data-amount-index="' + rowIndex + '"]');
  if (cell && currentRows[rowIndex]) {
    cell.textContent = currentRows[rowIndex].amount || '';
  }
}

// -------------------- 明細TSV 再構築 --------------------

function rebuildDetailsText() {
  if (!currentRows || !currentRows.length) {
    lastDetailsText = '';
    return;
  }
  var linesOut = [];
  linesOut.push('No\t品名\t規格\t単位\t合計数量\t契約単価\t金額\t備考');
  var r, rr, lineOut;
  for (r = 0; r < currentRows.length; r++) {
    rr = currentRows[r];
    lineOut = [
      rr.no   || '',
      rr.name || '',
      rr.spec || '',
      rr.unit || '',
      rr.qty  || '',
      rr.price || '',
      rr.amount || '',
      rr.note || ''
    ].join('\t');
    linesOut.push(lineOut);
  }
  lastDetailsText = linesOut.join('\n');
}

// -------------------- サマリ表示＆再計算 --------------------

// サマリ表示の共通処理
function displaySummary(summary, isRecalc) {
  var sumBox = document.getElementById('sumBox');
  if (!summary || !(summary.base || summary.tax || summary.total || summary.calcBaseStr)) {
    sumBox.textContent = '';
    lastSummaryText = '';
    return;
  }

  var sText = '';

  // ① 金額の合計行（毎回最新の計算値を表示）
  if (summary.calcBaseStr) {
    sText += '金額の合計:' + summary.calcBaseStr;
    // 「表修正時の再計算」の場合、金額合計と課税対象額が一致していれば "<" を表示
    if (isRecalc && originalSummary && originalSummary.base) {
      var calcNum = parseNumberForCalc(summary.calcBaseStr);
      var origNum = parseNumberForCalc(originalSummary.base);
      if (!isNaN(calcNum) && !isNaN(origNum) && Math.abs(calcNum - origNum) < 0.5) {
        sText += '<';
      }
    } else if (summary.baseCheckMark) {
      // 「初回解析」の場合、パーサ判定の "<" を使用
      sText += summary.baseCheckMark;
    }
    sText += '\n';
  }

  // ② [最終ページ集計] ブロック（originalSummary があれば常に原本値を保持）
  if (originalSummary && (originalSummary.base || originalSummary.tax || originalSummary.total)) {
    // originalSummary の値を使用（初回解析時・表修正後も変不変更）
    sText += '[最終ページ集計]\n';
    if (originalSummary.base) {
      sText += '課税対象額: ' + originalSummary.base + '\n';
    }
    if (originalSummary.tax) {
      sText += '消費税: ' + originalSummary.tax + '\n';
    }
    if (originalSummary.total) {
      sText += '合計: ' + originalSummary.total + '\n';
    }
  } else {
    // originalSummary がない場合（初回解析で summary を使用）
    sText += '[最終ページ集計]\n';
    if (summary.base) {
      sText += '課税対象額: ' + summary.base + '\n';
    }
    if (summary.tax) {
      sText += '消費税: ' + summary.tax + '\n';
    }
    if (summary.total) {
      sText += '合計: ' + summary.total + '\n';
    }
  }

  sumBox.textContent = sText;
  lastSummaryText = sText;
}

// 金額合計から課税対象額・消費税・合計を再計算（再計算版）
function recalcTotalsFromRows() {
  if (!currentRows || !currentRows.length) {
    displaySummary({ base: '', tax: '', total: '', calcBaseStr: '' }, true);
    return;
  }
  var sum = 0;
  var i, a;
  for (i = 0; i < currentRows.length; i++) {
    a = parseNumberForCalc(currentRows[i].amount);
    if (!isNaN(a)) {
      sum += a;
    }
  }
  var baseNum = sum;
  var taxNum  = Math.floor(baseNum * TAX_RATE + 1e-6);
  var totalNum = baseNum + taxNum;

  var summaryNew = {
    base:  formatAmount(baseNum),
    tax:   formatInt(taxNum),
    total: formatInt(totalNum),
    calcBaseStr: formatAmount(baseNum),
    baseCheckMark: '' // 再計算時は "<" 判定は行わない（表示は displaySummary で行う）
  };
  displaySummary(summaryNew, true); // true: 表修正時の再計算フラグ
}

// 1行の数量・単価変更に伴う金額＆サマリ再計算
function recalcRowAndTotals(rowIndex) {
  var row = currentRows[rowIndex];
  if (!row) return;

  var q = parseNumberForCalc(row.qty);
  var p = parseNumberForCalc(row.price);

  if (!isNaN(q) && !isNaN(p)) {
    var newAmountVal = q * p;
    row.amount = formatAmount(newAmountVal);
    updateRowAmountCell(rowIndex);

    addNoteMark(row, '@'); // 数量 or 単価を変更した行
    applyCalcMark(row);    // "<" マークの再判定
  } else {
    removeNoteMark(row, '<');
  }
  updateRowNoteCell(rowIndex);
  rebuildDetailsText();
  recalcTotalsFromRows();
}

// -------------------- ヘッダー＋宛先用テキストの組み立て --------------------

function buildAllDataText() {
  var combined = VERSION_TEXT + '\n\n';

  // 日付
  var billDate = document.getElementById('billDate');
  var billDateVal = billDate ? (billDate.value || '') : '';
  if (billDateVal) {
    combined += '日付: ' + formatDateForCopy(billDateVal) + '\n\n';
  } else {
    combined += '日付:\n\n';
  }

  // 宛先・業者名
  var toText    = document.getElementById('toText');
  var vendorText = document.getElementById('vendorText');
  var toVal     = toText ? (toText.value || '') : '';
  var vendorVal = vendorText ? (vendorText.value || '') : '';

  combined += '宛先:\n' + (toVal ? toVal + '\n\n' : '\n\n');
  combined += '業者名:\n' + (vendorVal ? vendorVal + '\n\n' : '\n\n');

  // 納地・業者名ヘッダー
  if (lastHeaderText) {
    combined += lastHeaderText + '\n\n';
  }

  // 明細
  if (lastDetailsText) {
    combined += lastDetailsText + '\n\n';
  }

  // 最終ページ集計
  if (lastSummaryText) {
    combined += lastSummaryText + '\n';
  }

  return combined;
}

// -------------------- 描画 --------------------

function renderTable(rows, summary) {
  var resultDiv   = document.getElementById('result');
  var summaryDiv  = document.getElementById('summary');
  var headerBox   = document.getElementById('headerBox');
  var copyMsgSpan = document.getElementById('copyMsg');
  var vendorTA    = document.getElementById('vendorText');

  resultDiv.innerHTML   = '';
  summaryDiv.innerHTML  = '';
  headerBox.innerHTML   = 'ここに納地・業者名ヘッダーが表示されます。';
  document.getElementById('sumBox').innerHTML = '';
  copyMsgSpan.innerHTML = '';

  lastHeaderText  = '';
  lastDetailsText = '';
  lastSummaryText = '';
  currentRows     = [];
  originalSummary = summary || { base:'',tax:'',total:'',calcBaseStr:'',baseCheckMark:'' };

  if (!rows || !rows.length) {
    summaryDiv.innerHTML = '抽出された明細はありません。';
    return;
  }

  currentRows = rows; // グローバルに保持

  // ★ここから差し替え★
  // 行ごとの「金額 = 合計数量 × 契約単価」反映 ＋ "<"（計算OK）判定を初期付与
  var idx;
  for (idx = 0; idx < currentRows.length; idx++) {
    // 金額欄を「数量 × 単価」に統一
    var qInit = parseNumberForCalc(currentRows[idx].qty);
    var pInit = parseNumberForCalc(currentRows[idx].price);
    if (!isNaN(qInit) && !isNaN(pInit)) {
      currentRows[idx].amount = formatAmount(qInit * pInit);
    }
    // 計算OKなら備考に "<" を付与
    applyCalcMark(currentRows[idx]);
  }
  // ★ここまで差し替え★

  // 業者名一覧
  var vendorMap = {};
  var v;
  for (idx = 0; idx < rows.length; idx++) {
    v = rows[idx].vendor || '';
    if (v) {
      vendorMap[v] = true;
    }
  }
  var vendorList = [];
  for (var key in vendorMap) {
    if (vendorMap.hasOwnProperty(key)) {
      vendorList.push(key);
    }
  }

  if (vendorList.length === 0) {
    headerBox.textContent = '納地: （不明）\n業者名: （検出なし）';
    lastHeaderText = '納地: （不明）\n業者名: （検出なし）';
  } else if (vendorList.length === 1) {
    var fullVendor = vendorList[0];
    var sv = splitVendorFull(fullVendor);
    var hLines = [];
    if (sv.site) {
      hLines.push('納地: ' + sv.site);
    }
    hLines.push('業者名: ' + (sv.vendor || fullVendor));
    var headerText = hLines.join('\n');
    headerBox.textContent = headerText;
    lastHeaderText = headerText;

    // 業者名テキストボックスの1行目に業者名をセット（2行目以降は保持）
    if (vendorTA) {
      var cur = vendorTA.value || '';
      var lines = cur ? cur.split('\n') : ['','','','',''];
      if (lines.length < 5) {
        while (lines.length < 5) lines.push('');
      }
      lines[0] = (sv.vendor || fullVendor);
      vendorTA.value = lines.join('\n');
    }
  } else {
    var headerTextMulti = '業者名（' + vendorList.length + '件）:\n' + vendorList.join('\n');
    headerBox.textContent = headerTextMulti;
    lastHeaderText = headerTextMulti;
    // vendorText は自動設定しない
  }

  // 明細テーブル描画（品名・単位・数量・単価は編集可、金額は自動、備考付き）
  var table = document.createElement('table');
  var thead = document.createElement('thead');
  var trh   = document.createElement('tr');
  var headers = ['No', '品名', '規格', '単位', '合計数量', '契約単価', '金額', '備考'];

  for (idx = 0; idx < headers.length; idx++) {
    var th = document.createElement('th');
    th.appendChild(document.createTextNode(headers[idx]));
    trh.appendChild(th);
  }
  thead.appendChild(trh);
  table.appendChild(thead);

  var tbody = document.createElement('tbody');
  var r, row, tr;

  for (r = 0; r < rows.length; r++) {
    row = rows[r];
    tr  = document.createElement('tr');

    // No（固定表示）
    var tdNo = document.createElement('td');
    tdNo.appendChild(document.createTextNode(row.no || ''));
    tr.appendChild(tdNo);

    // 品名（編集可）
    var tdName = document.createElement('td');
    var nameInput = document.createElement('input');
    nameInput.type = 'text';
    nameInput.className = 'name-input';
    nameInput.value = row.name || '';
    nameInput.setAttribute('data-row-index', String(r));
    nameInput.oninput = function() {
      var idxStr = this.getAttribute('data-row-index');
      var idxNum = parseInt(idxStr, 10);
      if (!isNaN(idxNum) && currentRows[idxNum]) {
        currentRows[idxNum].name = this.value;
        addNoteMark(currentRows[idxNum], '*'); // 品名修正
        updateRowNoteCell(idxNum);
        rebuildDetailsText();
      }
    };
    tdName.appendChild(nameInput);
    tr.appendChild(tdName);

    // 規格（編集可）
    var tdSpec = document.createElement('td');
    var specInput = document.createElement('input');
    specInput.type = 'text';
    specInput.className = 'spec-input';
    specInput.value = row.spec || '';
    specInput.setAttribute('data-row-index', String(r));
    specInput.oninput = function() {
      var idxStr = this.getAttribute('data-row-index');
      var idxNum = parseInt(idxStr, 10);
      if (!isNaN(idxNum) && currentRows[idxNum]) {
        currentRows[idxNum].spec = this.value;
        rebuildDetailsText();
      }
    };
    tdSpec.appendChild(specInput);
    tr.appendChild(tdSpec);

    // 単位（編集可）
    var tdUnit = document.createElement('td');
    var unitInput = document.createElement('input');
    unitInput.type = 'text';
    unitInput.className = 'unit-input';
    unitInput.value = row.unit || '';
    unitInput.setAttribute('data-row-index', String(r));
    unitInput.oninput = function() {
      var idxStr = this.getAttribute('data-row-index');
      var idxNum = parseInt(idxStr, 10);
      if (!isNaN(idxNum) && currentRows[idxNum]) {
        currentRows[idxNum].unit = this.value;
        addNoteMark(currentRows[idxNum], '*'); // 単位修正
        updateRowNoteCell(idxNum);
        rebuildDetailsText();
      }
    };
    tdUnit.appendChild(unitInput);
    tr.appendChild(tdUnit);

    // 合計数量（編集可）
    var tdQty = document.createElement('td');
    var qtyInput = document.createElement('input');
    qtyInput.type = 'text';
    qtyInput.className = 'qty-input';
    qtyInput.value = row.qty || '';
    qtyInput.setAttribute('data-row-index', String(r));
    qtyInput.oninput = function() {
      var idxStr = this.getAttribute('data-row-index');
      var idxNum = parseInt(idxStr, 10);
      if (!isNaN(idxNum) && currentRows[idxNum]) {
        currentRows[idxNum].qty = this.value;
        recalcRowAndTotals(idxNum); // 数量変更 → 金額＆集計再計算
      }
    };
    tdQty.appendChild(qtyInput);
    tr.appendChild(tdQty);

    // 契約単価（編集可）
    var tdPrice = document.createElement('td');
    var priceInput = document.createElement('input');
    priceInput.type = 'text';
    priceInput.className = 'price-input';
    priceInput.value = row.price || '';
    priceInput.setAttribute('data-row-index', String(r));
    priceInput.oninput = function() {
      var idxStr = this.getAttribute('data-row-index');
      var idxNum = parseInt(idxStr, 10);
      if (!isNaN(idxNum) && currentRows[idxNum]) {
        currentRows[idxNum].price = this.value;
        recalcRowAndTotals(idxNum); // 単価変更 → 金額＆集計再計算
      }
    };
    tdPrice.appendChild(priceInput);
    tr.appendChild(tdPrice);

    // 金額（自動計算表示のみ）
    var tdAmount = document.createElement('td');
    tdAmount.style.textAlign = 'right';
    tdAmount.setAttribute('data-amount-index', String(r));
    tdAmount.appendChild(document.createTextNode(row.amount || ''));
    tr.appendChild(tdAmount);

    // 備考
    var tdNote = document.createElement('td');
    tdNote.setAttribute('data-note-index', String(r));
    tdNote.appendChild(document.createTextNode(row.note || ''));
    tr.appendChild(tdNote);

    tbody.appendChild(tr);
  }
  table.appendChild(tbody);
  resultDiv.appendChild(table);

  summaryDiv.innerHTML = '抽出明細件数: ' + rows.length + ' 件';

  // サマリ（初期表示は原本値＋金額合計）
  displaySummary(originalSummary, false); // false: 初回解析

  // 明細TSV再構築（備考列込み）
  rebuildDetailsText();
}

// -------------------- コピー --------------------

function fallbackCopy(text) {
  var ta = document.createElement('textarea');
  ta.value = text;
  ta.style.position = 'fixed';
  ta.style.left = '-1000px';
  ta.style.top  = '-1000px';
  document.body.appendChild(ta);
  ta.select();
  try {
    document.execCommand('copy');
  } catch (e) {
    // 旧ブラウザ用フォールバック
  }
  document.body.removeChild(ta);
}

function copyTextToClipboard(text, label) {
  var msgSpan = document.getElementById('copyMsg');
  if (!text) {
    msgSpan.innerHTML = label + '用のデータがありません。';
    return;
  }

  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(text).then(function () {
      msgSpan.innerHTML = label + 'をコピーしました。';
    }, function () {
      fallbackCopy(text);
      msgSpan.innerHTML = label + 'をコピーしました。（旧方式）';
    });
  } else {
    fallbackCopy(text);
    msgSpan.innerHTML = label + 'をコピーしました。';
  }
}

function zeroPad2(value) {
  var num = parseInt(value, 10);
  if (isNaN(num)) return '00';
  return num < 10 ? '0' + num : String(num);
}

function getEdgeMajorVersion() {
  var ua = (window.navigator && window.navigator.userAgent) ? window.navigator.userAgent : '';
  var m = ua.match(/Edg\/([0-9]+)/);
  if (!m) return 0;
  return parseInt(m[1], 10) || 0;
}

function isLegacyEdgeCompatibilityMode() {
  var major = getEdgeMajorVersion();
  return major > 0 && major <= 95;
}

function readBlobAsTextWithFileReader(blob, onSuccess, onError) {
  if (!window.FileReader) {
    onError({ name: 'ReadError' });
    return;
  }

  try {
    var reader = new FileReader();
    reader.onload = function() {
      onSuccess(String(reader.result || ''));
    };
    reader.onerror = function() {
      onError({ name: 'ReadError' });
    };
    reader.readAsText(blob);
  } catch (e) {
    onError({ name: 'ReadError' });
  }
}

function readBlobAsTextCompat(blob, onSuccess, onError) {
  if (!blob) {
    onError({ name: 'ReadError' });
    return;
  }

  if (typeof blob.text === 'function') {
    try {
      blob.text().then(function(text) {
        onSuccess(String(text || ''));
      }, function() {
        readBlobAsTextWithFileReader(blob, onSuccess, onError);
      });
      return;
    } catch (e) {
      readBlobAsTextWithFileReader(blob, onSuccess, onError);
      return;
    }
  }

  readBlobAsTextWithFileReader(blob, onSuccess, onError);
}

// -------------------- 自動解析イベント --------------------

(function() {
  function initPageHandlers() {
  var ta           = document.getElementById('src');
  var btnClear     = document.getElementById('btnClear');
  var btnSaveLocal = document.getElementById('btnSaveLocal');
  var btnOpenLocal = document.getElementById('btnOpenLocal');
  var btnCopyAll   = document.getElementById('btnCopyAll');

  // ★これを追加★
  var btnPrintText = document.getElementById('btnPrintText');
  // ★ここまで追加★

  var btnPrint     = document.getElementById('btnPrint');
  var bulkSpecText = document.getElementById('bulkSpecText');
  var btnBulkSpec  = document.getElementById('btnBulkSpec');

  // ★これを追加★
  var btnGyousya   = document.getElementById('btnGyousya');
  // ★ここまで追加★

  var toText       = document.getElementById('toText');
  var vendorText   = document.getElementById('vendorText');
  var billDate     = document.getElementById('billDate');
  var debugLogBody = document.getElementById('debugLogBody');
  var btnCopyLog   = document.getElementById('btnCopyLog');
  var btnClearLog  = document.getElementById('btnClearLog');
  var LOCAL_DATA_FILE_NAME = 'data.json';
  var DEBUG_LOG_LIMIT = 400;

  if (!ta || !btnClear || !btnSaveLocal || !btnOpenLocal || !btnCopyAll || !btnPrint || !bulkSpecText || !btnBulkSpec || !toText || !vendorText || !billDate) {
    if (window.console && console.error) {
      console.error('UI初期化に必要な要素が見つからないため、イベント登録を中断しました。');
    }
    return;
  }

  function formatDebugLogTime(date) {
    return [
      zeroPad2(date.getHours()),
      zeroPad2(date.getMinutes()),
      zeroPad2(date.getSeconds())
    ].join(':') + '.' + ('00' + date.getMilliseconds()).slice(-3);
  }

  function summarizeDebugDetail(detail) {
    if (detail === null || detail === undefined || detail === '') {
      return '';
    }
    if (typeof detail === 'string') {
      return detail;
    }
    try {
      return JSON.stringify(detail);
    } catch (e) {
      return String(detail);
    }
  }

  function describeError(err) {
    if (!err) return 'unknown';
    if (typeof err === 'string') return err;
    var name = err.name || 'Error';
    var message = err.message || '';
    return message ? (name + ': ' + message) : name;
  }

  function appendDebugLog(level, scope, message, detail) {
    var line = '[' + formatDebugLogTime(new Date()) + '] [' + level + '] [' + scope + '] ' + message;
    var detailText = summarizeDebugDetail(detail);
    var entry;

    if (detailText) {
      line += ' | ' + detailText;
    }

    if (!debugLogBody) {
      if (window.console && console.log) {
        console.log(line);
      }
      return;
    }

    if (debugLogBody.textContent === 'まだログはありません。') {
      debugLogBody.innerHTML = '';
    }

    entry = document.createElement('div');
    entry.className = 'debug-log-entry debug-log-' + String(level || 'info').toLowerCase();
    entry.textContent = line;
    debugLogBody.appendChild(entry);

    while (debugLogBody.childNodes.length > DEBUG_LOG_LIMIT) {
      debugLogBody.removeChild(debugLogBody.firstChild);
    }

    debugLogBody.scrollTop = debugLogBody.scrollHeight;
  }

  function logInfo(scope, message, detail) {
    appendDebugLog('info', scope, message, detail);
  }

  function logWarn(scope, message, detail) {
    appendDebugLog('warn', scope, message, detail);
  }

  function logError(scope, message, detail) {
    appendDebugLog('error', scope, message, detail);
  }

  function getDebugLogText() {
    if (!debugLogBody) {
      return '';
    }
    return (debugLogBody.textContent || '').replace(/^\s+|\s+$/g, '');
  }

  if (btnCopyLog) {
    btnCopyLog.onclick = function() {
      var logText = getDebugLogText();
      if (!logText || logText === 'まだログはありません。') {
        setStatusMessage('コピーできるログがありません。');
        logWarn('log', 'ログコピー対象がありません。');
        return;
      }
      copyTextToClipboard(logText, '動作ログ');
      logInfo('log', 'ログをクリップボードへコピーしました。', { length: logText.length });
    };
  }

  if (btnClearLog) {
    btnClearLog.onclick = function() {
      if (debugLogBody) {
        debugLogBody.innerHTML = '';
      }
      logInfo('log', 'ログをクリアしました。');
    };
  }

  logInfo('init', 'UI初期化を開始しました。', {
    version: VERSION_TEXT,
    protocol: window.location ? window.location.protocol : '',
    readyState: document.readyState
  });

  function setStatusMessage(msg) {
    var msgSpan = document.getElementById('copyMsg');
    if (msgSpan) {
      msgSpan.innerHTML = msg || '';
    }
    logInfo('status', 'ステータスメッセージを更新しました。', msg || '');
  }

  function isFileProtocolPage() {
    return window.location && window.location.protocol === 'file:';
  }

  function canUseDirectoryPicker() {
    return !isLegacyEdgeCompatibilityMode() && !!window.showDirectoryPicker;
  }

  function canUseOpenFilePicker() {
    return !isLegacyEdgeCompatibilityMode() && !!window.showOpenFilePicker;
  }

  function saveDataByDownload(jsonText) {
    var fileName = buildSaveFileName();
    var blob = new Blob([jsonText], { type: 'application/json' });
    logInfo('save', 'ダウンロード保存を開始します。', { fileName: fileName, size: jsonText.length });
    if (window.navigator && typeof window.navigator.msSaveOrOpenBlob === 'function') {
      logInfo('save', 'msSaveOrOpenBlob を使用します。', fileName);
      window.navigator.msSaveOrOpenBlob(blob, fileName);
      return;
    }
    if (window.navigator && typeof window.navigator.msSaveBlob === 'function') {
      logInfo('save', 'msSaveBlob を使用します。', fileName);
      window.navigator.msSaveBlob(blob, fileName);
      return;
    }
    if (!window.URL || typeof window.URL.createObjectURL !== 'function') {
      logError('save', 'Blob URL が利用できないため保存できません。');
      throw { name: 'NotSupportedError' };
    }
    var url = URL.createObjectURL(blob);
    var link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    logInfo('save', 'ダウンロード保存を実行しました。', fileName);
  }

  function sanitizeFileNamePart(text) {
    var s = String(text || '');
    s = s.replace(/[\\/:*?"<>|]/g, '');
    s = s.replace(/[\r\n\t]/g, ' ');
    s = s.replace(/^\s+|\s+$/g, '');
    s = s.replace(/\s+/g, ' ');
    return s;
  }

  function formatDateForFileName() {
    var d = (billDate && billDate.value) ? billDate.value : '';
    if (!d) {
      var now = new Date();
      var y = String(now.getFullYear());
      var m = zeroPad2(now.getMonth() + 1);
      var day = zeroPad2(now.getDate());
      return y + m + day;
    }
    return String(d).replace(/-/g, '');
  }

  function getVendorNameForFileName() {
    var v = vendorText && vendorText.value ? String(vendorText.value) : '';
    var lines = v.split(/\r?\n/);
    var firstLine = lines.length ? lines[0] : '';
    firstLine = sanitizeFileNamePart(firstLine);
    if (firstLine) return firstLine;
    if (currentRows && currentRows.length && currentRows[0].vendor) {
      return sanitizeFileNamePart(currentRows[0].vendor);
    }
    return '業者名未設定';
  }

  function getItemNameForFileName() {
    if (currentRows && currentRows.length && currentRows[0].name) {
      var item = sanitizeFileNamePart(currentRows[0].name);
      if (item) return item;
    }
    return '品名未設定';
  }

  function getCountLabelForFileName() {
    var count = currentRows && currentRows.length ? currentRows.length : 0;
    if (count <= 0) return '0件';
    if (count === 1) return '1件';
    return 'ほか' + String(count - 1) + '件';
  }

  function buildSaveFileName() {
    var datePart = formatDateForFileName();
    var vendorPart = getVendorNameForFileName();
    var itemPart = getItemNameForFileName();
    var countPart = getCountLabelForFileName();
    return datePart + '_' + vendorPart + '_' + itemPart + '_' + countPart + '.json';
  }

  function openDataByFileInput(onSuccess, onError) {
    var input = document.createElement('input');
    input.type = 'file';
    input.accept = '.json,application/json';
    logInfo('open', 'input[type=file] による読込を開始します。');
    input.onchange = function() {
      var file = input.files && input.files[0] ? input.files[0] : null;
      if (!file) {
        logWarn('open', 'ファイルが選択されませんでした。');
        onError({ name: 'AbortError' });
        return;
      }
      logInfo('open', 'ファイルが選択されました。', { fileName: file.name || '', size: file.size || 0 });
      readBlobAsTextWithFileReader(file, function(text) {
        try {
          logInfo('open', 'ファイル内容の読込に成功しました。', { fileName: file.name || '', textLength: text.length });
          onSuccess(JSON.parse(text));
        } catch (e) {
          logError('open', 'JSON解析に失敗しました。', { fileName: file.name || '', error: describeError(e) });
          onError({ name: 'SyntaxError' });
        }
      }, function(err) {
        logError('open', 'ファイル読込に失敗しました。', describeError(err));
        onError(err);
      });
    };
    input.click();
  }

  function pickDirectoryHandle(mode, onSuccess, onError) {
    if (!canUseDirectoryPicker()) {
      logWarn('directory', 'Directory Picker が利用できません。', { mode: mode || 'readwrite' });
      onError({ name: 'NotSupportedError' });
      return;
    }

    try {
      logInfo('directory', 'フォルダ選択を要求します。', { mode: mode || 'readwrite' });
      window.showDirectoryPicker({
        id: 'ledger-local-data',
        mode: mode || 'readwrite',
        startIn: 'documents'
      }).then(function(handle) {
        logInfo('directory', 'フォルダが選択されました。', { mode: mode || 'readwrite' });
        onSuccess(handle);
      }, function(err) {
        if (err && err.name === 'TypeError') {
          logWarn('directory', 'オプション付き showDirectoryPicker に失敗したため再試行します。', describeError(err));
          try {
            window.showDirectoryPicker().then(function(handle) {
              logInfo('directory', 'フォルダが選択されました。', { mode: mode || 'readwrite', fallback: true });
              onSuccess(handle);
            }, function(fallbackErr) {
              logError('directory', 'フォルダ選択に失敗しました。', describeError(fallbackErr));
              onError(fallbackErr);
            });
          } catch (fallbackError) {
            logError('directory', 'フォルダ選択の再試行に失敗しました。', describeError(fallbackError));
            onError(fallbackError);
          }
          return;
        }
        logError('directory', 'フォルダ選択に失敗しました。', describeError(err));
        onError(err);
      });
    } catch (err) {
      logError('directory', 'showDirectoryPicker 呼び出しに失敗しました。', describeError(err));
      onError(err);
    }
  }

  function openJsonDataWithPickerCompat(onSuccess, onError) {
    if (!canUseOpenFilePicker()) {
      logWarn('open', 'Open File Picker が利用できないため file input に切り替えます。');
      openDataByFileInput(onSuccess, onError);
      return;
    }

    try {
      logInfo('open', 'Open File Picker による読込を開始します。');
      window.showOpenFilePicker({
        id: 'ledger-local-open',
        multiple: false,
        startIn: 'documents',
        types: [
          {
            description: 'JSONファイル',
            accept: { 'application/json': ['.json'] }
          }
        ]
      }).then(function(fileHandles) {
        var fileHandle = fileHandles && fileHandles[0] ? fileHandles[0] : null;
        if (!fileHandle) {
          logWarn('open', 'ファイルハンドルが取得できませんでした。');
          onError({ name: 'AbortError' });
          return;
        }
        fileHandle.getFile().then(function(file) {
          logInfo('open', 'ファイルハンドルからファイルを取得しました。', { fileName: file.name || '', size: file.size || 0 });
          readBlobAsTextCompat(file, function(text) {
            try {
              logInfo('open', 'ファイル内容の読込に成功しました。', { fileName: file.name || '', textLength: text.length });
              onSuccess(JSON.parse(text));
            } catch (e) {
              logError('open', 'JSON解析に失敗しました。', { fileName: file.name || '', error: describeError(e) });
              onError({ name: 'SyntaxError' });
            }
          }, function(err) {
            logError('open', 'ファイル内容の読込に失敗しました。', describeError(err));
            onError(err);
          });
        }, function(err) {
          logError('open', 'ファイル取得に失敗しました。', describeError(err));
          onError(err);
        });
      }, function(err) {
        if (err && err.name === 'TypeError') {
          logWarn('open', 'オプション付き showOpenFilePicker に失敗したため再試行します。', describeError(err));
          try {
            window.showOpenFilePicker().then(function(fallbackHandles) {
              var fallbackHandle = fallbackHandles && fallbackHandles[0] ? fallbackHandles[0] : null;
              if (!fallbackHandle) {
                logWarn('open', '再試行でもファイルハンドルが取得できませんでした。');
                onError({ name: 'AbortError' });
                return;
              }
              fallbackHandle.getFile().then(function(fallbackFile) {
                logInfo('open', '再試行でファイルを取得しました。', { fileName: fallbackFile.name || '', size: fallbackFile.size || 0 });
                readBlobAsTextCompat(fallbackFile, function(fallbackText) {
                  try {
                    logInfo('open', '再試行でファイル内容の読込に成功しました。', { fileName: fallbackFile.name || '', textLength: fallbackText.length });
                    onSuccess(JSON.parse(fallbackText));
                  } catch (e) {
                    logError('open', '再試行でJSON解析に失敗しました。', { fileName: fallbackFile.name || '', error: describeError(e) });
                    onError({ name: 'SyntaxError' });
                  }
                }, function() {
                  logWarn('open', '再試行でのファイル読込に失敗したため file input に切り替えます。');
                  openDataByFileInput(onSuccess, onError);
                });
              }, function() {
                logWarn('open', '再試行でのファイル取得に失敗したため file input に切り替えます。');
                openDataByFileInput(onSuccess, onError);
              });
            }, function() {
              logWarn('open', '再試行のファイル選択に失敗したため file input に切り替えます。');
              openDataByFileInput(onSuccess, onError);
            });
          } catch (fallbackError) {
            logError('open', 'showOpenFilePicker の再試行に失敗しました。', describeError(fallbackError));
            openDataByFileInput(onSuccess, onError);
          }
          return;
        }
        logError('open', 'ファイル選択に失敗しました。', describeError(err));
        onError(err);
      });
    } catch (err) {
      logError('open', 'showOpenFilePicker 呼び出しに失敗しました。', describeError(err));
      onError(err);
    }
  }

  function handleOpenError(err) {
    logError('open', '読込エラーを処理します。', describeError(err));
    if (err && err.name === 'AbortError') {
      setStatusMessage('読み込みをキャンセルしました。');
      return;
    }
    if (err && err.name === 'NotFoundError') {
      setStatusMessage('ファイルが見つかりません。');
      return;
    }
    if (err && err.name === 'SyntaxError') {
      setStatusMessage('JSONの形式が不正です。');
      return;
    }
    setStatusMessage('読み込みに失敗しました。');
  }

  function createLocalSaveData() {
    var rowsCopy = [];
    var i;
    for (i = 0; i < currentRows.length; i++) {
      rowsCopy.push({
        vendor: currentRows[i].vendor || '',
        no: currentRows[i].no || '',
        name: currentRows[i].name || '',
        spec: currentRows[i].spec || '',
        unit: currentRows[i].unit || '',
        qty: currentRows[i].qty || '',
        price: currentRows[i].price || '',
        amount: currentRows[i].amount || '',
        note: currentRows[i].note || ''
      });
    }
    return {
      format: 'ledger-local-save-v1',
      savedAt: new Date().toISOString(),
      appVersion: VERSION_TEXT,
      srcText: ta.value || '',
      bulkSpecText: bulkSpecText.value || '',
      toText: toText.value || '',
      vendorText: vendorText.value || '',
      billDate: billDate.value || '',
      rows: rowsCopy,
      summary: originalSummary || { base:'', tax:'', total:'', calcBaseStr:'', baseCheckMark:'' }
    };
  }

  function applyLocalSaveData(data) {
    logInfo('open', '保存データを画面へ反映します。', {
      rows: data && data.rows ? data.rows.length : 0,
      hasSourceText: !!(data && data.srcText),
      savedAt: data && data.savedAt ? data.savedAt : ''
    });
    ta.value = data && data.srcText ? data.srcText : '';
    bulkSpecText.value = data && data.bulkSpecText ? data.bulkSpecText : '規格表のとおり';
    toText.value = data && data.toText ? data.toText : '';
    vendorText.value = data && data.vendorText ? data.vendorText : '';
    billDate.value = data && data.billDate ? data.billDate : '';

    if (data && data.rows && data.rows.length) {
      logInfo('open', '保存済み明細をそのまま描画します。', { rows: data.rows.length });
      renderTable(data.rows, data.summary || { base:'',tax:'',total:'',calcBaseStr:'',baseCheckMark:'' });
      return;
    }

    if (ta.value) {
      logInfo('open', '明細が無いため元テキストを再解析します。');
      safeDoParse('applyLocalSaveData', true);
      return;
    }

    logWarn('open', '表示対象データが無いため画面を初期表示へ戻します。');

    document.getElementById('result').innerHTML   = '';
    document.getElementById('summary').innerHTML  = '';
    document.getElementById('headerBox').innerHTML = 'ここに納地・業者名ヘッダーが表示されます。';
    document.getElementById('sumBox').innerHTML   = '';
    lastHeaderText  = '';
    lastDetailsText = '';
    lastSummaryText = '';
    currentRows     = [];
    originalSummary = null;
    lastAllDataText = '';
  }

  function doParse() {
    var text = ta.value || '';
    if (!text.replace(/\s+/g, '')) {
      renderTable([], { base:'',tax:'',total:'',calcBaseStr:'',baseCheckMark:'' });
      return { rows: [], summary: { base:'',tax:'',total:'',calcBaseStr:'',baseCheckMark:'' } };
    }
    // ledger_paste_parser.js のパーサを利用
    if (typeof parseLedgerText !== 'function') {
      throw new Error('parseLedgerText is not available');
    }
    var parsed = parseLedgerText(text);
    var rows    = parsed && parsed.rows ? parsed.rows : [];
    var summary = parsed && parsed.summary ? parsed.summary : { base:'',tax:'',total:'',calcBaseStr:'',baseCheckMark:'' };
    renderTable(rows, summary);
    return { rows: rows, summary: summary };
  }

  function safeDoParse(source, shouldLogSuccess) {
    try {
      var parsed = doParse();
      if (shouldLogSuccess) {
        logInfo('parse', '解析が完了しました。', {
          source: source || 'unknown',
          rows: parsed && parsed.rows ? parsed.rows.length : 0,
          textLength: ta && ta.value ? ta.value.length : 0
        });
      }
    } catch (err) {
      setStatusMessage('貼り付け解析に失敗しました。テキスト形式を確認してください。');
      logError('parse', '解析に失敗しました。', {
        source: source || 'unknown',
        error: describeError(err)
      });
      if (window.console && console.error) {
        console.error(err);
      }
    }
  }

  function scheduleParseWithRetry(source) {
    logInfo('parse', '解析を予約しました。', { source: source || 'unknown' });
    setTimeout(function() {
      safeDoParse((source || 'unknown') + ':retry0', true);
    }, 0);
    setTimeout(function() {
      safeDoParse((source || 'unknown') + ':retry60', true);
    }, 60);
  }

  // 貼り付け時に自動解析
  if (ta.addEventListener) {
    ta.addEventListener('paste', function() {
      scheduleParseWithRetry('paste');
    }, false);
    ta.addEventListener('input', function() {
      safeDoParse('input', false);
    }, false);
    ta.addEventListener('change', function() {
      safeDoParse('change', true);
    }, false);
    ta.addEventListener('keyup', function() {
      safeDoParse('keyup', false);
    }, false);
    ta.addEventListener('drop', function() {
      scheduleParseWithRetry('drop');
    }, false);
  } else {
    ta.onpaste = function() {
      scheduleParseWithRetry('paste');
    };
    ta.oninput = function() {
      safeDoParse('input', false);
    };
    ta.onkeyup = function() {
      safeDoParse('keyup', false);
    };
    ta.onchange = function() {
      safeDoParse('change', true);
    };
  }

  btnClear.onclick = function() {
  logInfo('ui', 'クリアボタンが押されました。');
  ta.value = '';
  document.getElementById('result').innerHTML   = '';
  document.getElementById('summary').innerHTML  = '';
  document.getElementById('headerBox').innerHTML = 'ここに納地・業者名ヘッダーが表示されます。';
  document.getElementById('sumBox').innerHTML   = '';
  document.getElementById('copyMsg').innerHTML  = '';
  bulkSpecText.value = '規格表のとおり'; // 初期値に戻す

  // ★ここでは日付・宛先・業者名はクリアしない
  // dateInput.value = '';
  // atenaText.value = '';
  // vendorText.value = '';

  lastHeaderText  = '';
  lastDetailsText = '';
  lastSummaryText = '';
  currentRows     = [];
  originalSummary = null;
  lastAllDataText = '';
  logInfo('ui', '画面の解析結果をクリアしました。');
};

  btnSaveLocal.onclick = function() {
    var data = createLocalSaveData();
    var jsonText = JSON.stringify(data, null, 2);
    var fileName = buildSaveFileName();

    logInfo('save', '保存ボタンが押されました。', {
      fileName: fileName,
      rows: data && data.rows ? data.rows.length : 0,
      textLength: jsonText.length
    });

    function fallbackSave(messagePrefix) {
      try {
        logWarn('save', '通常のフォルダ保存からダウンロード保存へ切り替えます。', { fileName: fileName, reason: messagePrefix });
        saveDataByDownload(jsonText);
        setStatusMessage(messagePrefix + '（' + fileName + '）。');
      } catch (e) {
        logError('save', 'ダウンロード保存にも失敗しました。', describeError(e));
        setStatusMessage('保存に失敗しました。');
      }
    }

    if (isFileProtocolPage() || !canUseDirectoryPicker()) {
      logWarn('save', 'フォルダ保存が使えない環境のためダウンロード保存を実行します。', {
        protocol: window.location ? window.location.protocol : '',
        canUseDirectoryPicker: canUseDirectoryPicker()
      });
      fallbackSave('JSONファイルとして保存しました');
      return;
    }

    pickDirectoryHandle('readwrite', function(dirHandle) {
      dirHandle.getFileHandle(fileName, { create: true }).then(function(fileHandle) {
        logInfo('save', '保存ファイルハンドルを取得しました。', fileName);
        fileHandle.createWritable().then(function(writable) {
          logInfo('save', '書き込みストリームを開きました。', fileName);
          writable.write(jsonText).then(function() {
            logInfo('save', 'JSON書き込みに成功しました。', { fileName: fileName, size: jsonText.length });
            writable.close().then(function() {
              logInfo('save', '書き込みストリームを閉じました。', fileName);
              setStatusMessage('データを保存しました（' + fileName + '）。');
            }, function(err) {
              if (err && err.name === 'AbortError') {
                logWarn('save', '保存のクローズ処理がキャンセルされました。', describeError(err));
                setStatusMessage('保存をキャンセルしました。');
                return;
              }
              logError('save', '書き込みストリームのクローズに失敗しました。', describeError(err));
              fallbackSave('フォルダ保存できないため、JSONファイル保存に切り替えました');
            });
          }, function(err) {
            if (err && err.name === 'AbortError') {
              logWarn('save', 'JSON書き込みがキャンセルされました。', describeError(err));
              setStatusMessage('保存をキャンセルしました。');
              return;
            }
            logError('save', 'JSON書き込みに失敗しました。', describeError(err));
            fallbackSave('フォルダ保存できないため、JSONファイル保存に切り替えました');
          });
        }, function(err) {
          if (err && err.name === 'AbortError') {
            logWarn('save', '書き込みストリーム作成がキャンセルされました。', describeError(err));
            setStatusMessage('保存をキャンセルしました。');
            return;
          }
          logError('save', '書き込みストリーム作成に失敗しました。', describeError(err));
          fallbackSave('フォルダ保存できないため、JSONファイル保存に切り替えました');
        });
      }, function(err) {
        if (err && err.name === 'AbortError') {
          logWarn('save', '保存先ファイルの取得がキャンセルされました。', describeError(err));
          setStatusMessage('保存をキャンセルしました。');
          return;
        }
        logError('save', '保存先ファイルの取得に失敗しました。', describeError(err));
        fallbackSave('フォルダ保存できないため、JSONファイル保存に切り替えました');
      });
    }, function(err) {
      if (err && err.name === 'AbortError') {
        logWarn('save', '保存先フォルダの選択がキャンセルされました。', describeError(err));
        setStatusMessage('保存をキャンセルしました。');
        return;
      }
      logError('save', '保存先フォルダの選択に失敗しました。', describeError(err));
      fallbackSave('フォルダ保存できないため、JSONファイル保存に切り替えました');
    });
  };

  btnOpenLocal.onclick = function() {
    logInfo('open', '読込ボタンが押されました。');

    function handleLoadedData(data, loadedLabel) {
      if (!data || data.format !== 'ledger-local-save-v1') {
        logError('open', '読み込んだJSONの形式が不正です。', data && data.format ? data.format : 'format missing');
        setStatusMessage('読み込んだデータ形式が不正です。');
        return;
      }
      logInfo('open', '保存データ形式の確認に成功しました。', { format: data.format, savedAt: data.savedAt || '' });
      applyLocalSaveData(data);
      setStatusMessage(loadedLabel);
    }

    function openFallbackFile(reasonLabel) {
      logWarn('open', 'フォールバックの file input 読込に切り替えます。', reasonLabel);
      openDataByFileInput(function(fallbackLoaded) {
        handleLoadedData(fallbackLoaded, reasonLabel);
      }, handleOpenError);
    }

    if (isFileProtocolPage()) {
      logInfo('open', 'file: プロトコルのため file input 読込を使用します。');
      openDataByFileInput(function(data) {
        handleLoadedData(data, 'データを読み込みました（' + LOCAL_DATA_FILE_NAME + '）。');
      }, handleOpenError);
      return;
    }

    if (canUseOpenFilePicker()) {
      logInfo('open', 'Open File Picker での読込を試行します。');
      openJsonDataWithPickerCompat(function(data) {
        handleLoadedData(data, 'データを読み込みました（' + LOCAL_DATA_FILE_NAME + '）。');
      }, function(err) {
        if (err && (err.name === 'SecurityError' || err.name === 'NotAllowedError' || err.name === 'InvalidStateError' || err.name === 'TypeError' || err.name === 'NotSupportedError')) {
          logWarn('open', 'Open File Picker が利用できなかったため file input へ切り替えます。', describeError(err));
          openFallbackFile('フォルダ読込できないため、JSONファイル読込に切り替えました。');
          return;
        }
        handleOpenError(err);
      });
      return;
    }

    if (canUseDirectoryPicker()) {
      logInfo('open', 'Directory Picker で data.json の読込を試行します。');
      pickDirectoryHandle('read', function(dirHandle) {
        dirHandle.getFileHandle(LOCAL_DATA_FILE_NAME).then(function(dirFileHandle) {
          logInfo('open', 'フォルダ内の読込対象ファイルを取得しました。', LOCAL_DATA_FILE_NAME);
          dirFileHandle.getFile().then(function(dirFile) {
            logInfo('open', 'フォルダ内のファイル取得に成功しました。', { fileName: dirFile.name || LOCAL_DATA_FILE_NAME, size: dirFile.size || 0 });
            readBlobAsTextCompat(dirFile, function(dirText) {
              try {
                logInfo('open', 'フォルダ内ファイルの読込に成功しました。', { fileName: dirFile.name || LOCAL_DATA_FILE_NAME, textLength: dirText.length });
                handleLoadedData(JSON.parse(dirText), 'データを読み込みました（' + LOCAL_DATA_FILE_NAME + '）。');
              } catch (e) {
                logError('open', 'フォルダ内ファイルのJSON解析に失敗しました。', { fileName: dirFile.name || LOCAL_DATA_FILE_NAME, error: describeError(e) });
                handleOpenError({ name: 'SyntaxError' });
              }
            }, function(err) {
              if (err && (err.name === 'SecurityError' || err.name === 'NotAllowedError' || err.name === 'InvalidStateError' || err.name === 'TypeError' || err.name === 'NotSupportedError' || err.name === 'ReadError')) {
                logWarn('open', 'フォルダ内ファイルの読込に失敗したため file input に切り替えます。', describeError(err));
                openFallbackFile('フォルダ読込できないため、JSONファイル読込に切り替えました。');
                return;
              }
              handleOpenError(err);
            });
          }, function(err) {
            if (err && (err.name === 'SecurityError' || err.name === 'NotAllowedError' || err.name === 'InvalidStateError' || err.name === 'TypeError' || err.name === 'NotSupportedError')) {
              logWarn('open', 'フォルダ内ファイルの取得に失敗したため file input に切り替えます。', describeError(err));
              openFallbackFile('フォルダ読込できないため、JSONファイル読込に切り替えました。');
              return;
            }
            handleOpenError(err);
          });
        }, function(err) {
          if (err && (err.name === 'SecurityError' || err.name === 'NotAllowedError' || err.name === 'InvalidStateError' || err.name === 'TypeError' || err.name === 'NotSupportedError')) {
            logWarn('open', 'フォルダ内の data.json 取得に失敗したため file input に切り替えます。', describeError(err));
            openFallbackFile('フォルダ読込できないため、JSONファイル読込に切り替えました。');
            return;
          }
          handleOpenError(err);
        });
      }, function(err) {
        if (err && (err.name === 'SecurityError' || err.name === 'NotAllowedError' || err.name === 'InvalidStateError' || err.name === 'TypeError' || err.name === 'NotSupportedError')) {
          logWarn('open', 'フォルダ選択に失敗したため file input に切り替えます。', describeError(err));
          openFallbackFile('フォルダ読込できないため、JSONファイル読込に切り替えました。');
          return;
        }
        handleOpenError(err);
      });
      return;
    }

    openDataByFileInput(function(data) {
      logInfo('open', '最終フォールバックとして file input 読込を実行します。');
      handleLoadedData(data, 'データを読み込みました（' + LOCAL_DATA_FILE_NAME + '）。');
    }, handleOpenError);
  };


  // 規格欄 一括入力
  btnBulkSpec.onclick = function() {
    var v = bulkSpecText.value || '';
    if (!currentRows || !currentRows.length) {
      document.getElementById('copyMsg').innerHTML = '明細がないため規格一括入力は行われません。';
      return;
    }
    var i;
    for (i = 0; i < currentRows.length; i++) {
      currentRows[i].spec = v;
    }
    var inputs = document.getElementsByClassName('spec-input');
    for (i = 0; i < inputs.length; i++) {
      inputs[i].value = v;
    }
    rebuildDetailsText();
    document.getElementById('copyMsg').innerHTML = '規格欄を一括入力しました。';
  };

  // 全データコピー（バージョン＋日付＋宛先＋業者名＋ヘッダー＋明細＋最終ページ集計）
  btnCopyAll.onclick = function() {
    var combined = buildAllDataText();
    copyTextToClipboard(combined, '全データ');
  };

  // ★これを追加★
  // テキスト印刷：全データコピーと同じ内容をプレーンテキストで印刷プレビュー
  if (btnPrintText) {
    btnPrintText.onclick = function() {
      var combined = buildAllDataText();
      if (!combined) {
        alert('印刷するデータがありません。');
        return;
      }

      // 新しいウィンドウでテキストを表示して印刷
      var w = window.open('', '_blank');
      if (!w) {
        alert('ポップアップがブロックされました。ブラウザの設定を確認してください。');
        return;
      }

      // テキストをHTMLとして安全に出すためにエスケープ
      var escaped = combined
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');

      w.document.open();
      w.document.write(
        '<!doctype html>' +
        '<html lang="ja">' +
        '<head>' +
        '<meta charset="UTF-8">' +
        '<title>テキスト印刷</title>' +
        '</head>' +
        '<body>' +
        '<pre style="font-family: monospace; font-size: 12px; white-space: pre-wrap; margin: 20px;">' +
        escaped +
        '</pre>' +
        '<script>window.onload=function(){window.print();};<\/script>' +
        '</body>' +
        '</html>'
      );
      w.document.close();
    };
  }
  // ★ここまで追加★



  // 印刷（print.js 側の printLedgerData を利用）
  btnPrint.onclick = function() {
    var combined = buildAllDataText();
    if (window.printLedgerData) {
      window.printLedgerData(combined);
    } else {
      alert('print.js が読み込まれていません。');
    }
  };


 // ★ここから追加：業者名等反映ボタン★
  if (btnGyousya) {
    btnGyousya.onclick = function () {
  var headerBox = document.getElementById('headerBox');
  var headerText = headerBox ? (headerBox.textContent || '') : '';

  if (!headerText) {
    alert('ヘッダー情報がありません。先に納入台帳テキストを貼り付けて解析してください。');
    return;
  }

  if (window.applyGyousyaFromHeader) {
    window.applyGyousyaFromHeader(headerText, 'toText', 'vendorText');
  } else {
    alert('gyousya.js が読み込まれていません。');
  }
};
  }
  // ★ここまで追加★

  // 日付欄に「当日から3日後」を初期設定
  if (billDate) {
    var today = new Date();
    var threeDaysLater = new Date(today);
    threeDaysLater.setDate(threeDaysLater.getDate() + 3);
    var yyyy = threeDaysLater.getFullYear();
    var mm = zeroPad2(threeDaysLater.getMonth() + 1);
    var dd = zeroPad2(threeDaysLater.getDate());
    billDate.value = yyyy + '-' + mm + '-' + dd;
    logInfo('init', '請求日初期値を設定しました。', billDate.value);
  }

  logInfo('init', 'UI初期化が完了しました。');

  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initPageHandlers);
  } else {
    initPageHandlers();
  }


})();

