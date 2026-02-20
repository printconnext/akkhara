/**
 * ==================================================================
 * ระบบคำนวณราคางานพิมพ์ (Ultimate Refined Version - Web Port)
 * Ported from Google Apps Script
 * ==================================================================
 */

// --- 1. Database & Config ---

const CONSTANTS = {
  // หมวด: ต้นทุนคงที่
  PRICE_PLATE: 500,        // ค่าเพลท (ต่อแผ่น)
  PRICE_MAKE_READY: 500,   // ค่าขึ้นแท่น (ต่อยก/สี)

  // หมวด: กระดาษตั้งเครื่อง
  MR_PAPER_PER_SIG: 100,   // กระดาษตั้งเครื่อง (แผ่น/ยก)

  // หมวด: ต้นทุนผันแปร
  PRICE_RUN_PER_1000: 150, // ค่ารันพิมพ์ (ต่อ 1,000 รอบ/สี)

  // หมวด: เคลือบปก (ต่อเล่ม)
  FINISH_UV: 2.5,
  FINISH_PVC: 4.0,
  FINISH_SPOT_UV: 5.0,

  // หมวด: เข้าเล่ม
  BIND_PERFECT: 5.0,      // ไสกาว
  BIND_SADDLE: 2.0,       // เย็บมุงหลังคา
  BIND_COLLATE: 5.0       // เก็บเล่ม
};

// Paper Database
const PAPER_DB = [
  { name: "ปอนด์ 100 แกรม", size: "24x35", price: 934.95 },
  { name: "ปอนด์ 100 แกรม", size: "31x43", price: 1483.5 },
  { name: "ปอนด์ 60 แกรม", size: "24x35", price: 577.23 },
  { name: "ปอนด์ 60 แกรม", size: "31x43", price: 915.9 },
  { name: "ปอนด์ 70 แกรม", size: "24x35", price: 654.47 },
  { name: "ปอนด์ 70 แกรม", size: "31x43", price: 1038.45 },
  { name: "ปอนด์ 80 แกรม", size: "24x35", price: 747.96 },
  { name: "ปอนด์ 80 แกรม", size: "31x43", price: 1186.8 },
  { name: "ถนอมสายตา 65 แกรม", size: "24x35", price: 704.4 },
  { name: "ถนอมสายตา 65 แกรม", size: "31x43", price: 1118 },
  { name: "ถนอมสายตา 75 แกรม", size: "24x35", price: 812.8 },
  { name: "ถนอมสายตา 75 แกรม", size: "31x43", price: 1290 },
  { name: "อาร์ตการ์ด 260 แกรม", size: "25x36", price: 3801.2 }, // Special case 25x36
  { name: "อาร์ตการ์ด 260 แกรม", size: "31x43", price: 3801.2 },
  { name: "อาร์ต 85 แกรม", size: "24x35", price: 771.51 },
  { name: "อาร์ต 85 แกรม", size: "31x43", price: 1224.43 },
  { name: "อาร์ต 105 แกรม", size: "24x35", price: 924.63 },
  { name: "อาร์ต 105 แกรม", size: "31x43", price: 1467.38 },
  { name: "อาร์ต 120 แกรม", size: "24x35", price: 1056.9 },
  { name: "อาร์ต 120 แกรม", size: "31x43", price: 1677 },
  { name: "อาร์ต 128 แกรม", size: "24x35", price: 1127.1 },
  { name: "อาร์ต 128 แกรม", size: "31x43", price: 1788.8 },
  { name: "อาร์ต 157 แกรม", size: "24x35", price: 1382.55 },
  { name: "อาร์ต 157 แกรม", size: "31x43", price: 2194.08 }
];

// Helper to get unique paper names for dropdown
const PAPER_NAMES = [...new Set(PAPER_DB.map(p => p.name))];

// --- 2. Calculation Logic ---

function getPaperPrice(name, bookSize) {
  if (!name || name === "(ไม่ใส่/None)") return 0;

  // Logic: ถ้าหนังสือไซส์ B (B5, B6) ให้ใช้กระดาษ 31x43, ถ้าไซส์ A (A4, A5) ให้ใช้ 24x35 หรือ 25x36
  const isSeriesB = bookSize.startsWith("B");
  const targetSheetSize = isSeriesB ? "31x43" : "24x35";

  // 1. Try exact match first
  let paper = PAPER_DB.find(p => p.name === name && p.size.replace(/\s/g, '') === targetSheetSize);

  // 2. Fallback: ถ้าหาไม่เจอ ให้หาที่มีราคา > 0 (อันแรกที่เจอ)
  if (!paper) {
    paper = PAPER_DB.find(p => p.name === name && p.price > 0);
  }

  // 3. Special case for Art Card 260g (25x36) if series A
  if (!paper && !isSeriesB && name.includes("260")) {
    paper = PAPER_DB.find(p => p.name === name && p.size.includes("25x36"));
  }

  return paper ? paper.price : 0;
}

function calculateCost() {
  // 1. Get Inputs
  const qty = parseInt(document.getElementById('qty').value) || 0;
  const size = document.getElementById('size').value;
  const coverPaper = document.getElementById('coverPaper').value;
  const coverColor = document.getElementById('coverColor').value;
  const coverFinish = document.getElementById('coverFinish').value;
  const bindType = document.getElementById('bindType').value;
  const marginPct = parseFloat(document.getElementById('marginPct').value) / 100 || 0;
  const vatPct = parseFloat(document.getElementById('vatPct').value) / 100 || 0;

  // Inner inputs (Array of 3 sets)
  const innerSets = [1, 2, 3].map(i => ({
    paper: document.getElementById(`inner${i}Paper`).value,
    pages: parseInt(document.getElementById(`inner${i}Pages`).value) || 0,
    color: document.getElementById(`inner${i}Color`).value
  }));

  if (qty <= 0) {
    // Show error or just return empty result
    return;
  }

  // 2. Logic Divisors
  let cvDiv = 0, inPaperDiv = 0, signatureDiv = 0;

  if (size === "A4") { cvDiv = 4; inPaperDiv = 16; signatureDiv = 8; }
  else if (size === "A5") { cvDiv = 8; inPaperDiv = 32; signatureDiv = 16; }
  else if (size === "A6") { cvDiv = 16; inPaperDiv = 64; signatureDiv = 32; }
  else if (size === "B5") { cvDiv = 4; inPaperDiv = 32; signatureDiv = 8; }
  else if (size === "B6") { cvDiv = 16; inPaperDiv = 64; signatureDiv = 32; }
  else { cvDiv = 4; inPaperDiv = 16; signatureDiv = 8; } // Default

  // 3. Finishing Costs
  let finishUnitCost = 0;
  if (coverFinish.includes("UV")) finishUnitCost = CONSTANTS.FINISH_UV;
  if (coverFinish.includes("PVC")) finishUnitCost = CONSTANTS.FINISH_PVC;
  if (coverFinish.includes("Spot")) finishUnitCost = CONSTANTS.FINISH_SPOT_UV;

  let bindUnitCost = (bindType === "ไสกาว" ? CONSTANTS.BIND_PERFECT : CONSTANTS.BIND_SADDLE) + CONSTANTS.BIND_COLLATE;

  // --- 4. Cover Check ---
  // สูตรใหม่: ปัดเศษทีละ 0.2 รีม
  const cvActualSheets = qty / cvDiv; // ใช้จริง (ไม่เผื่อเสีย)
  const cvActualReams = cvActualSheets / 500;
  const cvSheets = Math.ceil((qty * 1.05) / cvDiv); // เผื่อเสีย 5%
  const cvReamsRaw = cvSheets / 500;
  const cvReams = Math.ceil(cvReamsRaw * 5) / 5; // ปัดขึ้นทีละ 0.2
  const cvSpareReams = cvReams - cvActualReams; // ส่วนเผื่อเสีย+ปัดเศษ

  const costCvPaper = cvReams * getPaperPrice(coverPaper, size);

  let cvPlates = 0, cvColors = 0;
  let cvColorLabel = '';
  if (coverColor.includes("4/0")) { cvPlates = 4; cvColors = 4; cvColorLabel = '4 สี หน้าเดียว'; }
  else if (coverColor.includes("4/4")) { cvPlates = 8; cvColors = 8; cvColorLabel = '4 สี สองหน้า'; }
  else if (coverColor.includes("2/0")) { cvPlates = 2; cvColors = 2; cvColorLabel = '2 สี หน้าเดียว'; }
  else if (coverColor.includes("2/2")) { cvPlates = 4; cvColors = 4; cvColorLabel = '2 สี สองหน้า'; }
  else { cvPlates = 1; cvColors = 1; cvColorLabel = '1 สี'; }

  const costCvPlate = cvPlates * CONSTANTS.PRICE_PLATE;
  const costCvPrint = (1 * cvColors * CONSTANTS.PRICE_MAKE_READY) + ((cvSheets * cvColors * CONSTANTS.PRICE_RUN_PER_1000) / 1000);

  // --- 5. Inner Loop ---
  let totalInPaper = 0;
  let totalInPlate = 0;
  let totalInPrint = 0;
  let totalInReams = 0;
  let grandTotalPages = 0;
  const innerDetails = []; // Track per-set details

  innerSets.forEach((set, index) => {
    if (set.paper === "(ไม่ใส่/None)" || !set.paper || set.pages === 0) return;

    // Warning: Check signatures (Optional UI alert)
    // if (set.pages % 4 !== 0) console.warn(`Set ${index+1} pages not mod 4`);

    grandTotalPages += set.pages;

    // 5.1 จำนวนยก
    const numSignatures = Math.ceil(set.pages / signatureDiv);

    // 5.2 คำนวณกระดาษรวม
    const mrSheets = numSignatures * CONSTANTS.MR_PAPER_PER_SIG;
    const mrReamsRaw = mrSheets / 500;

    const inSheetsReq = (set.pages * qty) / inPaperDiv;
    const actualReams = inSheetsReq / 500; // ใช้จริง (ไม่เผื่อเสีย)
    const inSheetsTotal = Math.ceil(inSheetsReq * 1.03); // Spare 3%
    const prodReamsRaw = inSheetsTotal / 500;

    // รวมยอดดิบ แล้วปัดเศษตามขนาดห่อกระดาษ
    // ปอนด์/ถนอมสายตา = ห่อละ 1.0 รีม, อาร์ต = ห่อละ 0.5 รีม
    const totalRawReams = mrReamsRaw + prodReamsRaw;
    const isArt = set.paper.includes("อาร์ต");
    const packSize = isArt ? 0.5 : 1.0;
    const sectionReams = Math.ceil(totalRawReams / packSize) * packSize;
    const spareReams = sectionReams - actualReams; // ส่วนเผื่อเสีย+ตั้งเครื่อง+ปัดเศษ

    const sectionPaperCost = sectionReams * getPaperPrice(set.paper, size);

    // 5.3 เพลท
    let inColors = 1;
    // Check color logic
    if (set.color.includes("4 สี") || set.color.includes("CMYK")) inColors = 4;
    else if (set.color.includes("2 สี")) inColors = 2;

    // เพลท: คำนวณจากจำนวนหน้าต่อด้าน (signatureDiv/2)
    // ทุกสีใช้สูตรเดียวกัน: จำนวนหน้า ÷ หน้าต่อด้าน = จำนวนชุดเพลท
    const pagesPerForme = Math.floor(signatureDiv / 2);
    const numFormes = Math.ceil(set.pages / pagesPerForme);

    const sectionPlateCost = (numFormes * inColors) * CONSTANTS.PRICE_PLATE;

    // 5.4 ค่าพิมพ์
    const totalMakeReady = numFormes * inColors * CONSTANTS.PRICE_MAKE_READY;
    const inImpressions = qty * (set.pages / signatureDiv);
    const totalRun = (inImpressions * inColors * CONSTANTS.PRICE_RUN_PER_1000) / 1000;
    const sectionPrintCost = totalMakeReady + totalRun;

    // Track per-set details
    innerDetails.push({
      setIndex: index + 1,
      paper: set.paper,
      pages: set.pages,
      actualReams,
      spareReams,
      reams: sectionReams,
      paperCost: sectionPaperCost,
      numFormes,
      inColors,
      plateCount: numFormes * inColors,
      plateCost: sectionPlateCost
    });

    totalInPaper += sectionPaperCost;
    totalInPlate += sectionPlateCost;
    totalInPrint += sectionPrintCost;
    totalInReams += sectionReams;
  });

  // 6. Summary
  const totalPaper = costCvPaper + totalInPaper;
  const totalPlate = costCvPlate + totalInPlate;
  const totalPrint = costCvPrint + totalInPrint;
  const costCoating = finishUnitCost * qty;
  const costBinding = bindUnitCost * qty;
  const totalFinish = costCoating + costBinding;

  const costTotal = totalPaper + totalPlate + totalPrint + totalFinish;
  const profit = costTotal * marginPct;
  const vat = (costTotal + profit) * vatPct;
  const grandTotal = costTotal + profit + vat;
  const pricePerBook = grandTotal / qty;

  // 7. Render Config
  updateUI({
    size, grandTotalPages, qty,
    // Cover details
    cvActualReams, cvSpareReams, cvReams, costCvPaper, cvPlates, cvColors, cvColorLabel,
    // Inner details
    innerDetails, totalInReams, totalInPaper,
    // Totals
    totalPaper, totalPlate, totalPrint,
    costCoating, costBinding, totalFinish,
    costTotal, profit, marginPct, vat, vatPct, grandTotal, pricePerBook
  });
}

function updateUI(data) {
  // Formatter
  const fmt = (n) => n.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const fmt0 = (n) => n.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 0 });

  // --- Plate Breakdown ---
  let plateHTML = `<div class="detail-item">ปก: ${data.cvColorLabel} (${data.cvPlates} แผ่น)</div>`;
  data.innerDetails.forEach(d => {
    const colorLabel = d.inColors === 4 ? '4 สี' : d.inColors === 2 ? '2 สี' : '1 สี';
    plateHTML += `<div class="detail-item">เนื้อใน ชุด ${d.setIndex}: ${colorLabel} × ${d.numFormes} ชุด = ${d.plateCount} แผ่น</div>`;
  });
  document.getElementById('res-plate-detail').innerHTML = plateHTML;
  document.getElementById('res-plate').textContent = fmt(data.totalPlate);

  // --- Paper Breakdown ---
  document.getElementById('res-paper').textContent = fmt(data.totalPaper);
  let paperHTML = '';
  paperHTML += `<div class="detail-item"><b>ปก: ${fmt(data.costCvPaper)} ฿</b></div>`;
  paperHTML += `<div class="detail-item detail-sub">ใช้จริง ${fmt(data.cvActualReams)} + เผื่อ ${fmt(data.cvSpareReams)} = ${fmt(data.cvReams)} รีม</div>`;
  data.innerDetails.forEach(d => {
    paperHTML += `<div class="detail-item"><b>เนื้อใน ชุด ${d.setIndex}: ${fmt(d.paperCost)} ฿</b></div>`;
    paperHTML += `<div class="detail-item detail-sub">ใช้จริง ${fmt(d.actualReams)} + เผื่อ ${fmt(d.spareReams)} = ${fmt(d.reams)} รีม</div>`;
  });
  if (data.innerDetails.length === 0) {
    paperHTML += `<div class="detail-item">เนื้อใน: ไม่มี</div>`;
  }
  document.getElementById('res-paper-detail').innerHTML = paperHTML;

  // --- Other costs ---
  document.getElementById('res-print').textContent = fmt(data.totalPrint);
  document.getElementById('res-coating').textContent = fmt(data.costCoating);
  document.getElementById('res-binding').textContent = fmt(data.costBinding);

  document.getElementById('res-cost-total').textContent = fmt(data.costTotal);
  document.getElementById('res-cost-per-book').textContent = fmt(data.costTotal / data.qty);

  // Update Quotation Card
  document.getElementById('quo-profit').textContent = fmt(data.profit);
  document.getElementById('quo-profit-label').textContent = `กำไร (${(data.marginPct * 100).toFixed(0)}%)`;

  document.getElementById('quo-vat').textContent = fmt(data.vat);
  document.getElementById('quo-vat-label').textContent = `VAT (${(data.vatPct * 100).toFixed(0)}%)`;

  document.getElementById('quo-grand-total').textContent = fmt(data.grandTotal);
  document.getElementById('quo-price-per-book').textContent = fmt(data.pricePerBook) + " บาท";
}

// --- 3. Initialization ---

document.addEventListener('DOMContentLoaded', () => {
  // Populate Paper Dropdowns
  const paperDropdowns = ['coverPaper', 'inner1Paper', 'inner2Paper', 'inner3Paper'];
  const allPapers = ["(ไม่ใส่/None)", ...PAPER_NAMES];

  paperDropdowns.forEach(id => {
    const el = document.getElementById(id);
    if (!el) return;

    // Clear existing
    el.innerHTML = '';

    allPapers.forEach(p => {
      const option = document.createElement('option');
      option.value = p;
      option.textContent = p;
      el.appendChild(option);
    });

    // Set defaults
    if (id === 'coverPaper') el.value = "อาร์ตการ์ด 260 แกรม";
    if (id === 'inner1Paper') el.value = "อาร์ต 120 แกรม";
  });

  // Attach Event Listeners to ALL inputs
  const inputs = document.querySelectorAll('input, select');
  inputs.forEach(input => {
    input.addEventListener('change', calculateCost);
    input.addEventListener('input', calculateCost); // For sliders/text typing
  });

  // Initial Calculate
  calculateCost();
});
