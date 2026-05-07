// excel-parser.js — 前端 Excel 解析（對應 Python parse_excel()）
// 依賴：SheetJS (XLSX) 已透過 CDN 載入

function parseExcel(workbook, filename) {
  // ── 找報名資料工作表（模糊比對，含「報名」即可）──
  const regSheetName = workbook.SheetNames.find(n => n.includes('報名'));
  if (!regSheetName) {
    return { error: `找不到「報名資料」工作表（現有工作表：${workbook.SheetNames.join(', ')}）` };
  }

  const ws = workbook.Sheets[regSheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  if (rows.length < 2) {
    return { error: 'Excel 沒有資料列' };
  }

  const headers = rows[0];
  const allRows = rows.slice(1);

  // ── helpers ──
  function colidx(keyword) {
    return headers.findIndex(h => h != null && String(h).includes(keyword));
  }

  function safeNum(val) {
    if (val == null) return 0;
    const n = Number(val);
    return isNaN(n) ? 0 : n;
  }

  // 有日期的才是真實資料列（第一欄非空）
  const dataRows = allRows.filter(r => r[0] != null);
  if (dataRows.length === 0) {
    return { error: '找不到有效的報名資料列（第一欄需有日期）' };
  }

  // ── 合計列偵測（最後一個無日期但有實繳金額的列）──
  const paidCol = colidx('實繳金額');
  let totalsRow = null;
  if (paidCol >= 0) {
    for (let i = allRows.length - 1; i >= 0; i--) {
      if (allRows[i][0] == null && safeNum(allRows[i][paidCol]) !== 0) {
        totalsRow = allRows[i];
        break;
      }
    }
  }

  // ── 欄位偵測 ──
  const orderCol = 1;

  // 郵寄費：多欄加總
  const postalCols = headers.reduce((acc, h, i) => {
    if (h == null) return acc;
    const s = String(h);
    if ((s.includes('郵寄') && s.includes('金額')) ||
        (s.includes('現場報到') && (s.includes('金額') || s.includes('費')))) {
      acc.push(i);
    }
    return acc;
  }, []);

  const chipCol = colidx('晶片押金訂單總金額');
  const refundCol = colidx('退費手續費總金額');
  const downgradeCol = colidx('降組不退費總金額');

  // ATM / 超商手續費：多欄加總
  const atmCols = headers.reduce((acc, h, i) => {
    if (h && String(h).includes('ATM') && String(h).includes('手續費')) acc.push(i);
    return acc;
  }, []);
  const cvsCols = headers.reduce((acc, h, i) => {
    if (h && String(h).includes('超商') && String(h).includes('手續費')) acc.push(i);
    return acc;
  }, []);

  // 信用卡線上刷卡總金額
  const creditCardCol = headers.findIndex(h =>
    h && String(h).includes('信用卡') && (String(h).includes('刷卡') || String(h).includes('線上'))
  );

  // 折扣 / 優惠欄偵測（3 種類型）
  // 類型 A：原有折扣欄（含「折」+「金額」+「優惠/折扣」）例如「85折優惠總金額」
  // 類型 B：優惠金額欄（含「優惠」+「金額」）例如「早鳥優惠-菁英組ES總金額」「青年優惠總金額」
  // 類型 C：團報優惠欄（精確比對「團報優惠」）— 無「金額」後綴，合計列通常為 null
  const discountCols = [];
  const discountDetails = [];
  headers.forEach((h, i) => {
    if (h == null) return;
    const s = String(h);
    const isTypeA = s.includes('折') && s.includes('金額') && (s.includes('優惠') || s.includes('折扣'));
    const isTypeB = s.includes('優惠') && s.includes('金額');
    const isTypeC = s === '團報優惠';
    if (isTypeA || isTypeB || isTypeC) {
      discountCols.push(i);
      const label = s.replace('總金額', '').replace('金額', '').trim();
      discountDetails.push({ col: i, label, header: s });
    }
  });

  // ── 財務數字：優先用合計列，否則逐訂單加總 ──
  let actualPaid = 0, postal = 0, chipDeposit = 0, refundFee = 0, downgrade = 0;
  let atmFee = 0, cvsFee = 0, creditCardTotal = 0;

  if (totalsRow) {
    actualPaid = safeNum(totalsRow[paidCol]);
    postal = postalCols.reduce((s, c) => s + safeNum(totalsRow[c]), 0);
    chipDeposit = chipCol >= 0 ? safeNum(totalsRow[chipCol]) : 0;
    refundFee = refundCol >= 0 ? safeNum(totalsRow[refundCol]) : 0;
    downgrade = downgradeCol >= 0 ? safeNum(totalsRow[downgradeCol]) : 0;
    atmFee = atmCols.reduce((s, c) => s + safeNum(totalsRow[c]), 0);
    cvsFee = cvsCols.reduce((s, c) => s + safeNum(totalsRow[c]), 0);
    creditCardTotal = creditCardCol >= 0 ? safeNum(totalsRow[creditCardCol]) : 0;
  } else {
    // 逐訂單加總（每筆訂單只取第一列）
    const seen = new Set();
    for (const row of dataRows) {
      const oid = row[orderCol];
      if (oid && !seen.has(oid)) {
        seen.add(oid);
        if (paidCol >= 0) actualPaid += safeNum(row[paidCol]);
        postalCols.forEach(c => { postal += safeNum(row[c]); });
        if (chipCol >= 0) chipDeposit += safeNum(row[chipCol]);
        if (refundCol >= 0) refundFee += safeNum(row[refundCol]);
        if (downgradeCol >= 0) downgrade += safeNum(row[downgradeCol]);
        atmCols.forEach(c => { atmFee += safeNum(row[c]); });
        cvsCols.forEach(c => { cvsFee += safeNum(row[c]); });
        if (creditCardCol >= 0) creditCardTotal += safeNum(row[creditCardCol]);
      }
    }
  }

  // ── 優惠/折扣金額：一律逐列加總（合計列常為 null）──
  // 團報優惠是每訂單一筆（只取訂單首列），其他優惠是每人一筆（逐列加總）
  let discountTotal = 0;
  const discountBreakdown = {};
  const teamDiscountCol = headers.findIndex(h => h && String(h) === '團報優惠');

  for (const detail of discountDetails) {
    let sum = 0;
    if (detail.col === teamDiscountCol) {
      // 團報優惠：每訂單一筆
      const seen = new Set();
      for (const row of dataRows) {
        const oid = row[orderCol];
        if (oid && !seen.has(oid)) {
          seen.add(oid);
          sum += safeNum(row[detail.col]);
        }
      }
    } else {
      // 早鳥/青年等優惠：每人一筆
      for (const row of dataRows) {
        sum += safeNum(row[detail.col]);
      }
    }
    if (sum !== 0) {
      discountBreakdown[detail.label] = Math.round(sum);
      discountTotal += sum;
    }
  }

  // ── 組別人數 ──
  const eventCol = colidx('參與項目');
  const typeCol = colidx('訂單類型');
  if (eventCol < 0) {
    return { error: '找不到「參與項目」欄位，請確認是否為 iRunner 匯出格式' };
  }

  const registration = {};
  for (const row of dataRows) {
    const event = row[eventCol];
    if (event) {
      const key = String(event);
      registration[key] = (registration[key] || 0) + 1;
    }
  }

  // ── 公關人數 ──
  const prCounts = {};
  for (const k of Object.keys(registration)) prCounts[k] = 0;

  const feeColPr = colidx('報名項目費用');
  const ADDON_PAID_KEYWORD = '需付加價購';
  const REG_PR_KEYWORDS = ['全免費', '免報名費', '公關', '贊助', 'VIP免費'];

  // ── 有付費的訂單 ID 集合（用於排除親子組誤判）──
  const paidOrderIds = new Set();
  if (feeColPr >= 0) {
    for (const row of dataRows) {
      const oid = row[orderCol];
      if (oid && row[feeColPr] != null && safeNum(row[feeColPr]) > 0) {
        paidOrderIds.add(oid);
      }
    }
  }

  // 優先從「免費名單」或「公關名單」工作表
  const FREE_SHEET_NAMES = ['免費名單', '公關名單'];
  const freeSheetName = workbook.SheetNames.find(n => FREE_SHEET_NAMES.includes(n));
  if (freeSheetName) {
    const wsFree = workbook.Sheets[freeSheetName];
    const freeRows = XLSX.utils.sheet_to_json(wsFree, { header: 1, defval: null });
    const freeHeaders = freeRows.length > 0 ? freeRows[0] : [];
    const freeEventCol = freeHeaders.findIndex(h => h && String(h).includes('參與項目'));
    if (freeRows.length > 1 && freeEventCol >= 0) {
      for (let i = 1; i < freeRows.length; i++) {
        if (freeRows[i][0] != null) {
          const e = freeRows[i][freeEventCol] != null ? String(freeRows[i][freeEventCol]) : null;
          if (e && e in prCounts) prCounts[e]++;
        }
      }
    }
    // 免費名單可能漏掉親子組公關（fee=null、訂單無付費成員、且非免費名單已列的人）
    if (feeColPr >= 0) {
      const freeCountByEvent = {};
      if (freeEventCol >= 0) {
        for (let i = 1; i < freeRows.length; i++) {
          if (freeRows[i][0] != null) {
            const e = freeRows[i][freeEventCol] != null ? String(freeRows[i][freeEventCol]) : '';
            freeCountByEvent[e] = (freeCountByEvent[e] || 0) + 1;
          }
        }
      }
      for (const row of dataRows) {
        const event = row[eventCol] != null ? String(row[eventCol]) : '';
        if (event in prCounts && row[feeColPr] == null && !paidOrderIds.has(row[orderCol])) {
          if ((freeCountByEvent[event] || 0) > 0) {
            freeCountByEvent[event]--;
          } else {
            prCounts[event]++;
          }
        }
      }
    }
  } else if (feeColPr >= 0) {
    // 無免費名單：fee=null 且訂單無付費成員 → 公關
    for (const row of dataRows) {
      const event = row[eventCol] != null ? String(row[eventCol]) : '';
      if (event in prCounts && row[feeColPr] == null && !paidOrderIds.has(row[orderCol])) {
        prCounts[event]++;
      }
    }
  } else if (typeCol >= 0) {
    // 備援：訂單類型關鍵字
    for (const row of dataRows) {
      const otype = row[typeCol] != null ? String(row[typeCol]) : '';
      const event = row[eventCol] != null ? String(row[eventCol]) : '';
      if (event in prCounts && REG_PR_KEYWORDS.some(k => otype.includes(k))) {
        prCounts[event]++;
      }
    }
  }

  // ── 加購數量：動態偵測 ──
  const addonQtyCols = [];
  headers.forEach((h, i) => {
    if (h == null) return;
    const s = String(h);
    if (!s.includes('總數量')) return;
    if (s.includes('加購') || s.includes('加價購') ||
        headers.some(h2 => h2 && String(h2) === s.replace('總數量', '總金額'))) {
      addonQtyCols.push([i, h]);
    }
  });

  function findAddonAmtCol(qtyHeader) {
    const key = String(qtyHeader).replace('總數量', '').replace('訂單', '').trim().replace(/-$/, '').trim();
    // 優先：不含「訂單」的金額欄
    let col = headers.findIndex(h =>
      h && String(h).includes('金額') && !String(h).includes('訂單') &&
      String(h).replace('總金額', '').trim().replace(/-$/, '').trim() === key
    );
    if (col >= 0) return col;
    // 備援：訂單總金額
    return headers.findIndex(h =>
      h && String(h).includes('金額') &&
      String(h).replace('總金額', '').replace('訂單', '').trim().replace(/-$/, '').trim() === key
    );
  }

  const addonAmtCols = {};
  addonQtyCols.forEach(([i, h]) => {
    const name = String(h).replace('總數量', '').replace('訂單', '').trim().replace(/-$/, '').trim();
    addonAmtCols[name] = findAddonAmtCol(h);
  });

  const addons = {};
  const addonPrAuto = {};
  const addonPricesAuto = {};

  for (const [i, h] of addonQtyCols) {
    const name = String(h).replace('總數量', '').replace('訂單', '').trim().replace(/-$/, '').trim();
    const totalQty = dataRows.reduce((s, row) => s + safeNum(row[i]), 0);
    addons[name] = Math.round(totalQty);

    // 公關加購數量
    if (feeColPr >= 0) {
      let prQty = 0;
      for (const row of dataRows) {
        if (row[feeColPr] == null &&
            !paidOrderIds.has(row[orderCol]) &&
            !(typeCol >= 0 && String(row[typeCol] || '').includes(ADDON_PAID_KEYWORD))) {
          prQty += safeNum(row[i]);
        }
      }
      if (prQty > 0) addonPrAuto[name] = Math.round(prQty);
    } else if (typeCol >= 0) {
      let prQty = 0;
      for (const row of dataRows) {
        const otype = String(row[typeCol] || '');
        if (otype.includes('全免費') && !otype.includes(ADDON_PAID_KEYWORD)) {
          prQty += safeNum(row[i]);
        }
      }
      if (prQty > 0) addonPrAuto[name] = Math.round(prQty);
    }

    // 推算單價：取最常見的單價
    const amtCol = addonAmtCols[name];
    if (amtCol >= 0) {
      const priceCount = {};
      for (const row of dataRows) {
        const qty = safeNum(row[i]);
        const amt = safeNum(row[amtCol]);
        if (qty > 0 && amt > 0) {
          const up = Math.round(amt / qty);
          priceCount[up] = (priceCount[up] || 0) + 1;
        }
      }
      const entries = Object.entries(priceCount);
      if (entries.length > 0) {
        entries.sort((a, b) => b[1] - a[1]);
        addonPricesAuto[name] = Number(entries[0][0]);
      }
    }
  }

  // ── 組別單價 + 報名費逐列加總 + 繳費人數 ──
  const regPricesAuto = {};
  const regFeeByGroup = {};
  const regPaidCount = {};
  const feeCol = colidx('報名項目費用');
  if (feeCol >= 0) {
    for (const k of Object.keys(registration)) {
      const priceCount = {};
      let feeSum = 0;
      let paidCnt = 0;
      for (const row of dataRows) {
        if (String(row[eventCol] || '') !== k) continue;
        if (row[feeCol] == null) continue;
        const fee = safeNum(row[feeCol]);
        if (fee > 0) {
          feeSum += fee;
          paidCnt++;
          const intFee = Math.round(fee);
          priceCount[intFee] = (priceCount[intFee] || 0) + 1;
        }
      }
      regFeeByGroup[k] = Math.round(feeSum);
      regPaidCount[k] = paidCnt;
      const entries = Object.entries(priceCount);
      if (entries.length > 0) {
        entries.sort((a, b) => b[1] - a[1]);
        regPricesAuto[k] = Number(entries[0][0]);
      }
    }
  }

  // ── 解析檔名 ──
  const filenameInfo = parseFilename(filename);

  return {
    registration,
    pr_counts: prCounts,
    addons,
    reg_prices_auto: regPricesAuto,
    reg_fee_by_group: regFeeByGroup,
    reg_paid_count: regPaidCount,
    addon_prices_auto: addonPricesAuto,
    addon_pr_auto: addonPrAuto,
    financials: {
      actual_paid: Math.round(actualPaid),
      postal: Math.round(postal),
      chip_deposit: Math.round(chipDeposit),
      refund_fee: Math.round(refundFee),
      downgrade: Math.round(downgrade),
      atm_fee: Math.round(atmFee),
      cvs_fee: Math.round(cvsFee),
      credit_card_total: Math.round(creditCardTotal),
      discount_total: Math.round(discountTotal),
      discount_breakdown: discountBreakdown,
    },
    has_totals_row: totalsRow !== null,
    filename: filename,
    filename_info: filenameInfo,
  };
}

// 解析 iRunner 檔名格式：「賽事名稱報名資料-MMDDHHmmss.xlsx」
function parseFilename(filename) {
  const base = filename.replace(/\.(xlsx?)$/i, '');
  const result = { raw: filename, eventName: '', timestamp: '', timestampFormatted: '' };

  // 嘗試匹配尾端的時間戳 -MMDDHHmmss 或 -MMDDHHmm
  const tsMatch = base.match(/-(\d{8,10})$/);
  if (tsMatch) {
    const ts = tsMatch[1];
    result.timestamp = ts;
    const namepart = base.slice(0, base.length - tsMatch[0].length);
    // 移除「報名資料」等後綴取得賽事名稱
    result.eventName = namepart.replace(/報名資料$/, '').replace(/報名名單$/, '').trim();

    // 格式化時間戳
    if (ts.length >= 8) {
      const mm = ts.slice(0, 2);
      const dd = ts.slice(2, 4);
      const HH = ts.slice(4, 6);
      const MM = ts.slice(6, 8);
      const ss = ts.length >= 10 ? ts.slice(8, 10) : '00';
      result.timestampFormatted = `${mm}/${dd} ${HH}:${MM}:${ss}`;
    }
  } else {
    // 無時間戳，嘗試從檔名取賽事名稱
    result.eventName = base
      .replace(/^【[^】]*】/, '')
      .replace(/報名資料.*$/, '')
      .replace(/報名名單.*$/, '')
      .trim();
  }

  return result;
}
