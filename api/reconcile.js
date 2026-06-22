import ExcelJS from 'exceljs';

const MONTH_NAMES = ['','January','February','March','April','May','June','July',
  'August','September','October','November','December'];
const MONTH_TEXT = {
  1:['JAN','JANUARY'],2:['FEB','FEBRUARY'],3:['MAR','MARCH'],4:['APR','APRIL'],
  5:['MAY'],6:['JUNE','JUN'],7:['JULY','JUL'],8:['AUG','AUGUST'],
  9:['SEPT','SEP','SEPTEMBER'],10:['OCT','OCTOBER'],
  11:['NOV','NOVEMBER'],12:['DEC','DECEMBER']
};
const ACTUAL_LABELS   = ['ACTUAL ORDER SHIPPED','ACTUAL ORDER','ACTUAL SHIP'];
const FORECAST_LABELS = ['FORECAST'];
const ALL_LABELS = [...ACTUAL_LABELS,...FORECAST_LABELS,'BEG. IN STOCK','INCOMING','ON ORDER','IN STOCK'];

function extractAMG(text) {
  const m = String(text||'').match(/\b([A-Z]{2,4}\d{4,}(?:-CA\d|-TX)?)\b/);
  return m ? m[1] : null;
}

function extractNums(text) {
  const nums = String(text||'').match(/\b(\d{5,10})\b/g) || [];
  return nums.filter(n => !(n.length===4 && (n.startsWith('20')||n.startsWith('19'))));
}

function findMonthCol(rowVals, monthNum, yearNum) {
  const aliases = (MONTH_TEXT[monthNum]||[]).map(a=>a.toUpperCase());
  for (let ci=0; ci<rowVals.length; ci++) {
    const v = rowVals[ci];
    if (v == null) continue;
    if (typeof v === 'string') {
      const vu = v.trim().toUpperCase();
      if (aliases.some(a => vu===a || vu.startsWith(a))) return ci;
    }
    if (v instanceof Date) {
      if (v.getMonth()+1 === monthNum && v.getFullYear() === yearNum) return ci;
    }
  }
  return null;
}

async function loadSAP(buf, monthNum) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  const ws = wb.worksheets[0];
  const headers = [];
  ws.getRow(1).eachCell((cell,ci) => { headers[ci-1] = String(cell.value||'').trim(); });
  const mname = MONTH_NAMES[monthNum].toLowerCase();
  const qColIdx = headers.findIndex(h => h.toLowerCase().includes(mname) && h.toLowerCase().includes('quantity'));
  if (qColIdx < 0) return [{}, `No quantity column for ${MONTH_NAMES[monthNum]}`];
  const data = {};
  ws.eachRow((row, ri) => {
    if (ri === 1) return;
    const itemNo = String(row.getCell(2).value||'').trim();
    if (!itemNo || itemNo==='null') return;
    const qty = row.getCell(qColIdx+1).value;
    if (qty != null) {
      const n = parseFloat(qty);
      if (!isNaN(n)) data[itemNo] = n;
    }
  });
  return [data, null];
}

async function process(sapBuf, fcBuf, tabName, monthNum, yearNum) {
  const [sapData, err] = await loadSAP(sapBuf, monthNum);
  if (err) return [null, err];

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(fcBuf);

  const ws = wb.worksheets.find(s => s.name.toLowerCase().includes(tabName.toLowerCase().trim()));
  if (!ws) {
    const names = wb.worksheets.map(s=>s.name).join(', ');
    return [null, `Tab "${tabName}" not found. Available: ${names}`];
  }

  // Read all rows into array for easy processing
  const allRows = [];
  ws.eachRow((row) => {
    const vals = [];
    row.eachCell({includeEmpty:true}, (cell) => {
      vals.push(cell.value instanceof Object && cell.value.result !== undefined ? cell.value.result : cell.value);
    });
    allRows.push(vals);
  });

  const results = [];
  const updates = []; // {rowIdx (0-based), colIdx (0-based), value}

  let i = 0;
  while (i < allRows.length) {
    const row = allRows[i];
    if (!row || !row[0]) { i++; continue; }
    const cell0 = String(row[0]).trim();
    const amg = extractAMG(cell0);
    if (!amg || ALL_LABELS.some(l => cell0.toUpperCase().startsWith(l))) { i++; continue; }

    const mCol = findMonthCol(row, monthNum, yearNum);
    if (mCol === null) { i++; continue; }

    const itemNums = extractNums(cell0);
    let sapQty = null;
    for (const n of itemNums) { if (sapData[n] !== undefined) { sapQty = sapData[n]; break; } }
    if (sapQty === null && sapData[amg] !== undefined) sapQty = sapData[amg];

    let forecastVal = null, actualRowIdx = null;
    for (let j=i+1; j<Math.min(i+8, allRows.length); j++) {
      const r = allRows[j];
      if (!r || !r[0]) continue;
      const lbl = String(r[0]).trim().toUpperCase();
      if (extractAMG(String(r[0])) && !ALL_LABELS.some(l=>lbl.startsWith(l))) break;
      const raw = mCol < r.length ? r[mCol] : null;
      let val = null;
      if (raw != null && String(raw).trim() !== '' && String(raw).trim() !== 'CUT FROM PO') {
        const n = parseFloat(raw); if (!isNaN(n)) val = n;
      }
      if (FORECAST_LABELS.some(l=>lbl.startsWith(l)) && !lbl.includes('ACTUAL'))
        forecastVal = val ?? 0;
      else if (ACTUAL_LABELS.some(l=>lbl.startsWith(l)))
        actualRowIdx = j;
    }

    let status='no_sap', variance=null, variancePct=null;
    if (sapQty !== null) {
      variance = sapQty - (forecastVal||0);
      variancePct = forecastVal ? parseFloat(((variance/forecastVal)*100).toFixed(1)) : null;
      status = variance>0?'over':(variance<0?'under':'match');
    }

    results.push({amg, item_num:itemNums[0]||null, description:cell0.replace(/\n/g,' ').trim().substring(0,65),
      forecast:forecastVal, actual:sapQty, variance, variance_pct:variancePct, status,
      written: actualRowIdx!==null && sapQty!==null});

    if (actualRowIdx !== null && sapQty !== null)
      updates.push({rowIdx: actualRowIdx, colIdx: mCol, value: sapQty});

    i++;
  }

  // Apply updates — only touch the value, nothing else
  ws.eachRow((row, ri) => {
    const upd = updates.find(u => u.rowIdx === ri-1);
    if (upd) {
      const cell = row.getCell(upd.colIdx+1);
      cell.value = upd.value;
    }
  });

  const outBuf = await wb.xlsx.writeBuffer();
  const b64 = Buffer.from(outBuf).toString('base64');

  return [{
    file_b64: b64,
    sheet: ws.name,
    month: `${MONTH_NAMES[monthNum]} ${yearNum}`,
    items_found: results.length,
    written: updates.length,
    over:   results.filter(r=>r.status==='over').length,
    under:  results.filter(r=>r.status==='under').length,
    match:  results.filter(r=>r.status==='match').length,
    no_sap: results.filter(r=>r.status==='no_sap').length,
    results
  }, null];
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') { res.status(200).end(); return; }
  try {
    const { sap_b64, fc_b64, tab, month, year } = req.body;
    const sapBuf = Buffer.from(sap_b64, 'base64');
    const fcBuf  = Buffer.from(fc_b64,  'base64');
    const [result, err] = await process(sapBuf, fcBuf, tab, parseInt(month), parseInt(year));
    if (err) { res.status(400).json({ error: err }); return; }
    res.status(200).json(result);
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
}
