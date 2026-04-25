
const STORAGE_KEY = "piece-rate-calculator-v1";
const STYLE_IMAGE_DB_NAME = "piece-rate-calculator-images";
const STYLE_IMAGE_STORE = "style-images";
const STORED_IMAGE_PREFIX = "stored-image:";
const DEFAULT_SIZE_LIST = ["XS", "S", "M", "L", "XL", "XXL"];
const DEFAULT_WASHCARE_TEMPLATE_PATH = "C:\\Users\\Lenovo\\Downloads\\SHEIN WASHCARE.btw";
const DEFAULT_WASHCARE_LABEL_WIDTH_MM = 35.5;
const DEFAULT_WASHCARE_LABEL_LENGTH_MM = 100;
const SECTION_ACCESS_TABS = ["styles", "cutting", "production", "acceptance", "dispatch"];
const SECTION_ACCESS_LABELS = {
  styles: "Style Master",
  cutting: "Cutting Entry",
  production: "Production Entry",
  acceptance: "Acceptance Entry",
  dispatch: "Dispatch Entry"
};
const DEFAULTS = {
  styles: [],
  styleImages: {},
  cuttingEntries: [],
  styleProductionEntries: [],
  productionEntries: [],
  acceptanceEntries: [],
  dispatchEntries: [],
  grnStatus: {
    id: "",
    grnFileName: "",
    pdfPageCount: 0,
    grnNumber: "",
    grnDate: "",
    supplier: "",
    supplierCode: "",
    supplierGstin: "",
    vendorInvoiceNo: "",
    poNumber: "",
    poDate: "",
    deliveryNo: "",
    consignee: "",
    warehouse: "",
    transporterName: "",
    consignmentNote: "",
    consignmentDate: "",
    vehicleNumber: "",
    deliveryChallanNo: "",
    totalChallanQty: 0,
    totalReceivedQty: 0,
    totalAcceptedQty: 0,
    totalShortQty: 0,
    items: [],
    packingListFileName: "",
    packingItems: [],
    packingListInvoiceNo: "",
    packingListRawText: "",
    comparisonRows: [],
    updatedAt: ""
  },
  grnReports: [],
  washcareRecords: [],
  payments: [],
  tallyCreditors: {
    endpoint: "",
    company: "",
    fromDate: "",
    asOnDate: "",
    bucketDays: "30,60,90",
    selectedParty: "",
    ledgers: [],
    vouchers: [],
    lastSyncAt: "",
    lastSyncSource: "",
    lastError: ""
  },
  settings: {
    sizes: DEFAULT_SIZE_LIST,
    accessCodes: {
      styles: "",
      cutting: "",
      production: "",
      acceptance: "",
      dispatch: ""
    }
  }
};
const CLOUD_REFRESH_MS = 15000;

let state = clone(DEFAULTS);
let firebaseApp = null;
let firebaseDb = null;
let firebaseStorage = null;
let lastCloudUpdatedAt = "";
let cloudRefreshTimer = null;
let isApplyingRemoteState = false;
let isSyncing = false;
let hasPendingCloudChanges = false;
let pastedStyleImageDataUrl = "";
let activeSharedSection = "";
let pendingStyleImportFile = null;
let styleImageCache = new Map();

const $ = (id) => document.getElementById(id);
const els = {
  sidebar: document.querySelector(".sidebar"),
  sidebarCards: document.querySelectorAll(".sidebar-card"),
  navLinks: document.querySelectorAll(".nav-link"),
  panels: document.querySelectorAll(".tab-panel"),
  styleForm: $("styleForm"),
  styleCardSearch: $("styleCardSearch"),
  runStyleImageDiagnostic: $("runStyleImageDiagnostic"),
  styleImageDiagnosticStatus: $("styleImageDiagnosticStatus"),
  sizeSettingsForm: $("sizeSettingsForm"),
  sizeListInput: $("sizeListInput"),
  accessCodeForm: $("accessCodeForm"),
  accessCodeStatus: $("accessCodeStatus"),
  operationRateRows: $("operationRateRows"),
  addOperationRateBtn: $("addOperationRateBtn"),
  operationRateTemplate: $("operationRateTemplate"),
  styleVariantRows: $("styleVariantRows"),
  addStyleVariantBtn: $("addStyleVariantBtn"),
  styleVariantTemplate: $("styleVariantTemplate"),
  styleCards: $("styleCards"),
  cuttingForm: $("cuttingForm"),
  cuttingStyleSearch: $("cuttingStyleSearch"),
  cuttingStyleSelect: $("cuttingStyleSelect"),
  cuttingSizeRows: $("cuttingSizeRows"),
  cuttingEntriesHead: $("cuttingEntriesHead"),
  cuttingEntriesTable: $("cuttingEntriesTable"),
  cuttingEntriesSearch: $("cuttingEntriesSearch"),
  styleProductionForm: $("styleProductionForm"),
  styleProductionStyleSearch: $("styleProductionStyleSearch"),
  styleProductionStyleSelect: $("styleProductionStyleSelect"),
  styleProductionSizeRows: $("styleProductionSizeRows"),
  styleProductionEntriesTable: $("styleProductionEntriesTable"),
  styleProductionEntriesSearch: $("styleProductionEntriesSearch"),
  styleProductionImportInput: $("styleProductionImportInput"),
  productionForm: $("productionForm"),
  productionStyleSearch: $("productionStyleSearch"),
  productionStyleSelect: $("productionStyleSelect"),
  productionOperationSelect: $("productionOperationSelect"),
  productionSizeSelect: $("productionSizeSelect"),
  productionEntriesTable: $("productionEntriesTable"),
  acceptanceForm: $("acceptanceForm"),
  acceptanceStyleSearch: $("acceptanceStyleSearch"),
  acceptanceStyleSelect: $("acceptanceStyleSelect"),
  acceptanceSizeRows: $("acceptanceSizeRows"),
  acceptanceEntriesTable: $("acceptanceEntriesTable"),
  dispatchForm: $("dispatchForm"),
  dispatchStyleSearch: $("dispatchStyleSearch"),
  dispatchStyleSelect: $("dispatchStyleSelect"),
  dispatchSizeRows: $("dispatchSizeRows"),
  dispatchEntriesTable: $("dispatchEntriesTable"),
  dispatchEntriesSearch: $("dispatchEntriesSearch"),
  grnPdfInput: $("grnPdfInput"),
  grnPackingListInput: $("grnPackingListInput"),
  grnStatusText: $("grnStatusText"),
  grnHeaderSummary: $("grnHeaderSummary"),
  grnSummaryCards: $("grnSummaryCards"),
  grnItemsTable: $("grnItemsTable"),
  grnComparisonTable: $("grnComparisonTable"),
  grnSavedTable: $("grnSavedTable"),
  saveGrnReportBtn: $("saveGrnReportBtn"),
  downloadGrnSheetBtn: $("downloadGrnSheetBtn"),
  downloadAllGrnBtn: $("downloadAllGrnBtn"),
  washcareForm: $("washcareForm"),
  washcareStyleSearch: $("washcareStyleSearch"),
  washcareStyleSelect: $("washcareStyleSelect"),
  washcareSearch: $("washcareSearch"),
  washcareTable: $("washcareTable"),
  washcarePreview: $("washcarePreview"),
  washcareSeedStatus: $("washcareSeedStatus"),
  washcareReportInput: $("washcareReportInput"),
  printWashcareBtn: $("printWashcareBtn"),
  copyWashcareCommandBtn: $("copyWashcareCommandBtn"),
  dashboardStats: $("dashboardStats"),
  pendingWorkflowTable: $("pendingWorkflowTable"),
  styleBillingTable: $("styleBillingTable"),
  workerBillingTable: $("workerBillingTable"),
  styleAmountReportTable: $("styleAmountReportTable"),
  reconciliationTable: $("reconciliationTable"),
  cuttingReportHead: $("cuttingReportHead"),
  cuttingReportTable: $("cuttingReportTable"),
  cuttingReportSearch: $("cuttingReportSearch"),
  dispatchReportTable: $("dispatchReportTable"),
  dispatchReportSearch: $("dispatchReportSearch"),
  operationCostTable: $("operationCostTable"),
  billingReportSearch: $("billingReportSearch"),
  reconciliationSearch: $("reconciliationSearch"),
  reportStartDate: $("reportStartDate"),
  reportEndDate: $("reportEndDate"),
  reportRangeSummary: $("reportRangeSummary"),
  clearReportDate: $("clearReportDate"),
  downloadStyleAmountReport: $("downloadStyleAmountReport"),
  downloadBillingPdf: $("downloadBillingPdf"),
  downloadCuttingReport: $("downloadCuttingReport"),
  downloadInternalChallan: $("downloadInternalChallan"),
  downloadFlowReport: $("downloadFlowReport"),
  paymentForm: $("paymentForm"),
  paymentHistoryTable: $("paymentHistoryTable"),
  tallySyncForm: $("tallySyncForm"),
  tallyEndpoint: $("tallyEndpoint"),
  tallyCompany: $("tallyCompany"),
  tallyFromDate: $("tallyFromDate"),
  tallyAsOnDate: $("tallyAsOnDate"),
  tallyBucketDays: $("tallyBucketDays"),
  tallyPartySearch: $("tallyPartySearch"),
  tallySyncStatus: $("tallySyncStatus"),
  importTallyXmlBtn: $("importTallyXmlBtn"),
  testTallyConnectionBtn: $("testTallyConnectionBtn"),
  tallyXmlInput: $("tallyXmlInput"),
  exportTallyAgeingBtn: $("exportTallyAgeingBtn"),
  exportTallyAgeingCsvBtn: $("exportTallyAgeingCsvBtn"),
  tallySummaryCards: $("tallySummaryCards"),
  tallyCreditorsTable: $("tallyCreditorsTable"),
  tallyInvoiceDetailsTable: $("tallyInvoiceDetailsTable"),
  tallySelectedPartyLabel: $("tallySelectedPartyLabel"),
  tallyBucketHead: $("tallyBucketHead"),
  tallyCurrentBucketHead: $("tallyCurrentBucketHead"),
  exportBtn: $("exportBtn"),
  importInput: $("importInput"),
  styleImportInput: $("styleImportInput"),
  styleImageRecoveryInput: $("styleImageRecoveryInput"),
  styleImageImportInput: $("styleImageImportInput"),
  skipExistingStylesToggle: $("skipExistingStylesToggle"),
  downloadPendingCmtSheet: $("downloadPendingCmtSheet"),
  cmtUpdateImportInput: $("cmtUpdateImportInput"),
  styleImageFile: $("styleImageFile"),
  pasteStyleImageBtn: $("pasteStyleImageBtn"),
  styleImagePasteStatus: $("styleImagePasteStatus"),
  styleFormImagePreview: $("styleFormImagePreview"),
  styleFormImagePreviewImg: $("styleFormImagePreviewImg"),
  cuttingImportInput: $("cuttingImportInput"),
  productionImportInput: $("productionImportInput"),
  storageModeBadge: $("storageModeBadge"),
  syncStatusText: $("syncStatusText"),
  imagePreviewModal: $("imagePreviewModal"),
  imagePreviewModalImg: $("imagePreviewModalImg"),
  closeImagePreview: $("closeImagePreview")
};

init().catch((error) => {
  console.error(error);
  alert("The app could not start correctly. Check the browser console for details.");
});

async function init() {
  state = await loadInitialState();
  normalizeState();
  await loadStoredStyleImages();
  await migrateStoredImageRefsToState();
  await migrateInlineStyleImages();
  configurePdfJs();
  bindTabs();
  seedOperationRows();
  seedStyleVariantRows();
  buildSizeInputs();
  setToday();
  if (els.washcareForm?.templatePath && !clean(els.washcareForm.templatePath.value)) {
    els.washcareForm.templatePath.value = DEFAULT_WASHCARE_TEMPLATE_PATH;
  }
  buildSizeSelect();
  els.sizeListInput.value = getSizes().join(", ");
  populateAccessCodeInputs();
  bindForms();
  applySharedSectionAccess();
  updateStorageModeUi();
  render();
  startCloudRefresh();
}

function bindTabs() {
  els.navLinks.forEach((btn) => btn.addEventListener("click", () => {
    if (activeSharedSection && !isSharedSectionTabAllowed(btn.dataset.tab)) return;
    activateTab(btn.dataset.tab);
  }));
}

function activateTab(tabId) {
  els.navLinks.forEach((n) => n.classList.toggle("active", n.dataset.tab === tabId));
  els.panels.forEach((p) => p.classList.toggle("active", p.id === tabId));
}

function seedOperationRows() {
  if (els.operationRateRows.children.length) return;
  ["Singer", "Overlock", "Kaaj", "Button", "Collar", "Patti"].forEach((name) => addOperationRow(name, ""));
}

function addOperationRow(name, rate) {
  const row = els.operationRateTemplate.content.cloneNode(true);
  row.querySelector('[name="operationName"]').value = name;
  row.querySelector('[name="operationRate"]').value = rate;
  els.operationRateRows.appendChild(row);
}

function seedStyleVariantRows() {
  if (!els.styleVariantRows || els.styleVariantRows.children.length) return;
  addStyleVariantRow("", "");
}

function addStyleVariantRow(color, orderQty) {
  if (!els.styleVariantRows || !els.styleVariantTemplate) return;
  const row = els.styleVariantTemplate.content.cloneNode(true);
  row.querySelector('[name="styleVariantColor"]').value = color;
  row.querySelector('[name="styleVariantOrderQty"]').value = orderQty;
  els.styleVariantRows.appendChild(row);
}

function buildSizeInputs() {
  const sizes = getSizes();
  els.cuttingSizeRows.innerHTML = sizes.map((size) => sizeCard(size, `cut_${size}`)).join("");
  els.styleProductionSizeRows.innerHTML = sizes.map((size) => sizeCard(size, `styleProd_${size}`)).join("");
  els.acceptanceSizeRows.innerHTML = sizes.map((size) => `
    <div class="size-card">
      <h4>${size}</h4>
      <label>Accepted Qty<input type="number" min="0" name="accepted_${size}" placeholder="0"></label>
      <label>Rejected Qty<input type="number" min="0" name="rejected_${size}" placeholder="0"></label>
    </div>`).join("");
  els.dispatchSizeRows.innerHTML = sizes.map((size) => sizeCard(size, `dispatch_${size}`)).join("");
}

function sizeCard(size, name) {
  return `<div class="size-card"><h4>${size}</h4><label>Quantity<input type="number" min="0" name="${name}" placeholder="0"></label></div>`;
}

function setToday() {
  const today = new Date().toISOString().slice(0, 10);
  els.cuttingForm.date.value = els.cuttingForm.date.value || today;
  els.styleProductionForm.date.value = els.styleProductionForm.date.value || today;
  els.productionForm.date.value = els.productionForm.date.value || today;
  els.acceptanceForm.date.value = els.acceptanceForm.date.value || today;
  els.dispatchForm.date.value = els.dispatchForm.date.value || today;
  if (els.paymentForm?.elements?.paymentDate) {
    els.paymentForm.elements.paymentDate.value = els.paymentForm.elements.paymentDate.value || today;
  }
  if (els.tallyAsOnDate) {
    els.tallyAsOnDate.value = els.tallyAsOnDate.value || state.tallyCreditors?.asOnDate || today;
  }
}

function buildSizeSelect() {
  els.productionSizeSelect.innerHTML = `<option value="">Select size</option>` + getSizes().map((s) => `<option value="${s}">${s}</option>`).join("");
}

function bindForms() {
  els.addOperationRateBtn.addEventListener("click", () => addOperationRow("", ""));
  els.addStyleVariantBtn?.addEventListener("click", () => addStyleVariantRow("", ""));
  els.operationRateRows.addEventListener("click", (e) => {
    if (e.target.classList.contains("remove-operation")) e.target.closest(".operation-rate-row").remove();
  });
  els.styleVariantRows?.addEventListener("click", (e) => {
    if (!e.target.classList.contains("remove-style-variant")) return;
    const rows = els.styleVariantRows.querySelectorAll(".style-variant-row");
    if (rows.length <= 1) {
      rows[0]?.querySelector('[name="styleVariantColor"]')?.focus();
      return;
    }
    e.target.closest(".style-variant-row")?.remove();
  });
  els.styleCards.addEventListener("click", handleStyleCardAction);
  els.runStyleImageDiagnostic?.addEventListener("click", runStyleImageDiagnostic);
  els.cuttingReportTable?.addEventListener("click", handleImagePreviewAction);
  els.dispatchReportTable?.addEventListener("click", handleImagePreviewAction);
  els.styleAmountReportTable?.addEventListener("click", handleImagePreviewAction);
  els.cuttingEntriesTable.addEventListener("click", handleCuttingTableAction);
  els.styleProductionEntriesTable.addEventListener("click", handleStyleProductionTableAction);
  els.productionEntriesTable.addEventListener("click", handleProductionTableAction);
  els.acceptanceEntriesTable.addEventListener("click", handleAcceptanceTableAction);
  els.dispatchEntriesTable.addEventListener("click", handleDispatchTableAction);
  els.washcareTable?.addEventListener("click", handleWashcareTableAction);
  els.sizeSettingsForm.addEventListener("submit", saveSizes);
  els.accessCodeForm?.addEventListener("submit", saveAccessCodes);
  els.accessCodeForm?.addEventListener("click", handleAccessLinkCopy);

  els.styleForm.addEventListener("submit", saveStyle);
  els.styleForm.addEventListener("reset", () => {
    setTimeout(() => {
      resetStyleVariantRows();
      clearPastedStyleImage();
      updateStyleFormImagePreview();
    }, 0);
  });
  els.styleForm.addEventListener("paste", handleStyleImagePaste);
  els.styleForm.image?.addEventListener("input", () => {
    if (clean(els.styleForm.image.value)) pastedStyleImageDataUrl = "";
    updateStyleFormImagePreview();
  });
  els.styleImageFile?.addEventListener("change", handleStyleImageFileChange);
  els.cuttingForm.addEventListener("submit", saveCutting);
  els.styleProductionForm.addEventListener("submit", saveStyleProduction);
  els.productionForm.addEventListener("submit", saveProduction);
  els.acceptanceForm.addEventListener("submit", saveAcceptance);
  els.dispatchForm.addEventListener("submit", saveDispatch);
  els.washcareForm?.addEventListener("submit", saveWashcare);
  els.washcareForm?.addEventListener("input", renderWashcarePreview);
  els.washcareForm?.addEventListener("reset", () => {
    window.setTimeout(() => {
      delete els.washcareForm.dataset.editId;
      renderWashcarePreview();
    }, 0);
  });
  els.paymentForm?.addEventListener("submit", savePayment);
  els.paymentForm?.addEventListener("input", syncPaymentFormTotal);
  els.productionStyleSelect.addEventListener("change", renderOperationSelect);
  els.reportStartDate?.addEventListener("change", renderReports);
  els.reportEndDate?.addEventListener("change", renderReports);
  els.clearReportDate.addEventListener("click", clearReportDateFilter);
  els.downloadStyleAmountReport.addEventListener("click", downloadStyleAmountReport);
  els.downloadBillingPdf?.addEventListener("click", downloadBillingPdf);
  els.downloadCuttingReport.addEventListener("click", downloadCuttingReport);
  els.downloadInternalChallan?.addEventListener("click", downloadInternalChallan);
  els.downloadFlowReport.addEventListener("click", downloadFlowReport);
  els.cuttingEntriesSearch?.addEventListener("input", renderCutting);
  els.styleProductionEntriesSearch?.addEventListener("input", renderStyleProduction);
  els.dispatchEntriesSearch?.addEventListener("input", renderDispatch);
  els.washcareSearch?.addEventListener("input", renderWashcare);
  els.billingReportSearch?.addEventListener("input", renderReports);
  els.reconciliationSearch?.addEventListener("input", renderReports);
  els.cuttingReportSearch?.addEventListener("input", renderReports);
  els.dispatchReportSearch?.addEventListener("input", renderReports);
  els.tallySyncForm?.addEventListener("submit", syncTallyCreditors);
  els.testTallyConnectionBtn?.addEventListener("click", testTallyConnection);
  els.importTallyXmlBtn?.addEventListener("click", () => els.tallyXmlInput?.click());
  els.tallyXmlInput?.addEventListener("change", importTallyXmlFile);
  els.exportTallyAgeingBtn?.addEventListener("click", downloadTallyAgeingWorkbook);
  els.exportTallyAgeingCsvBtn?.addEventListener("click", downloadTallyAgeingCsv);
  els.tallyPartySearch?.addEventListener("input", renderTallyCreditors);
  els.tallyBucketDays?.addEventListener("change", saveTallyPreferences);
  els.tallyAsOnDate?.addEventListener("change", saveTallyPreferences);
  els.tallyFromDate?.addEventListener("change", saveTallyPreferences);
  els.tallyEndpoint?.addEventListener("change", saveTallyPreferences);
  els.tallyCompany?.addEventListener("change", saveTallyPreferences);
  els.tallyCreditorsTable?.addEventListener("click", handleTallyCreditorTableClick);
  els.exportBtn.addEventListener("click", exportData);
  els.importInput.addEventListener("change", importData);
  els.styleImportInput.addEventListener("change", importStylesCsv);
  els.styleImageRecoveryInput?.addEventListener("change", importStyleImageSheet);
  els.styleImageImportInput?.addEventListener("change", importStylesCsvFromPendingSelection);
  els.downloadPendingCmtSheet?.addEventListener("click", downloadPendingCmtSheet);
  els.cmtUpdateImportInput?.addEventListener("change", importCmtUpdateSheet);
  els.cuttingImportInput.addEventListener("change", importCuttingCsv);
  els.styleProductionImportInput?.addEventListener("change", importStyleProductionCsv);
  els.productionImportInput.addEventListener("change", importProductionCsv);
  els.pasteStyleImageBtn?.addEventListener("click", pasteStyleImageFromClipboard);
  els.grnPdfInput?.addEventListener("change", importGrnPdf);
  els.grnPackingListInput?.addEventListener("change", importPackingListPdf);
  els.saveGrnReportBtn?.addEventListener("click", saveCurrentGrnReport);
  els.downloadGrnSheetBtn?.addEventListener("click", downloadGrnSheetWorkbook);
  els.downloadAllGrnBtn?.addEventListener("click", downloadAllGrnWorkbook);
  els.washcareStyleSelect?.addEventListener("change", handleWashcareStyleChange);
  els.washcareReportInput?.addEventListener("change", seedWashcareFromReportFile);
  els.printWashcareBtn?.addEventListener("click", handleWashcarePrint);
  els.copyWashcareCommandBtn?.addEventListener("click", copyWashcareCommand);
  bindStyleSearches();
  els.paymentHistoryTable?.addEventListener("click", handlePaymentTableAction);
  els.closeImagePreview?.addEventListener("click", closeImagePreview);
  els.imagePreviewModal?.addEventListener("click", (e) => {
    if (e.target instanceof HTMLElement && e.target.dataset.closeImagePreview === "true") {
      closeImagePreview();
    }
  });
}

function bindStyleSearches() {
  [
    [els.styleCardSearch, null],
    [els.cuttingStyleSearch, els.cuttingStyleSelect],
    [els.styleProductionStyleSearch, els.styleProductionStyleSelect],
    [els.productionStyleSearch, els.productionStyleSelect],
    [els.acceptanceStyleSearch, els.acceptanceStyleSelect],
    [els.dispatchStyleSearch, els.dispatchStyleSelect],
    [els.washcareStyleSearch, els.washcareStyleSelect]
  ].forEach(([input, select]) => {
    input?.addEventListener("input", () => {
      if (select) renderStyleSelect(select, input.value);
      else renderStyles();
    });
  });
}
async function saveStyle(e) {
  e.preventDefault();
  const f = new FormData(els.styleForm);
  const styleNumber = clean(f.get("styleNumber"));
  const editId = els.styleForm.dataset.editId || "";
  if (!styleNumber) return;
  const variants = collectStyleVariants();
  if (!variants.length) {
    alert("Please enter at least one colour.");
    return;
  }
  const duplicateVariant = variants.find((variant, index) => variants.findIndex((item) => normalizeKey(item.color) === normalizeKey(variant.color)) !== index);
  if (duplicateVariant) {
    alert("Please keep each colour only once in the style creation form.");
    return;
  }
  const operations = [...els.operationRateRows.querySelectorAll(".operation-rate-row")].map((row) => ({
    operationName: clean(row.querySelector('[name="operationName"]').value),
    rate: num(row.querySelector('[name="operationRate"]').value)
  })).filter((x) => x.operationName);

  const uploadedFile = els.styleForm.querySelector('[name="imageFile"]').files?.[0];
  const imageValue = clean(f.get("image"));
  const uploadedImage = uploadedFile ? await fileToDataUrl(uploadedFile) : null;
  const existingStyle = editId ? byId(editId) : null;
  const existingImageSrc = styleImageSrc(existingStyle);

  if (editId) {
    const variant = variants[0];
    if (state.styles.some((s) => s.id !== editId && s.styleNumber.toLowerCase() === styleNumber.toLowerCase() && clean(s.color).toLowerCase() === variant.color.toLowerCase())) {
      alert("This style number and color already exists.");
      return;
    }
    const stylePayload = buildStylePayload({
      id: editId,
      styleNumber,
      buyerName: clean(f.get("buyerName")),
      styleName: clean(f.get("styleName")),
      color: variant.color,
      orderQty: variant.orderQty,
      cmtRate: num(f.get("cmtRate")),
      serviceChargePct: num(f.get("serviceChargePct")),
      image: await prepareStyleImageForState(uploadedImage || pastedStyleImageDataUrl || imageValue || existingImageSrc || "", editId),
      notes: clean(f.get("notes")),
      operations
    });
    const index = state.styles.findIndex((s) => s.id === editId);
    if (index >= 0) state.styles[index] = stylePayload;
    delete els.styleForm.dataset.editId;
  } else {
    for (const variant of variants) {
      if (state.styles.some((s) => s.styleNumber.toLowerCase() === styleNumber.toLowerCase() && clean(s.color).toLowerCase() === variant.color.toLowerCase())) {
        alert(`This style number and color already exists: ${variant.color}`);
        return;
      }
    }
    variants.forEach((variant) => {
      const styleId = uid();
      state.styles.push(buildStylePayload({
        id: styleId,
        styleNumber,
        buyerName: clean(f.get("buyerName")),
        styleName: clean(f.get("styleName")),
        color: variant.color,
        orderQty: variant.orderQty,
        cmtRate: num(f.get("cmtRate")),
        serviceChargePct: num(f.get("serviceChargePct")),
        image: "",
        notes: clean(f.get("notes")),
        operations
      }));
    });
    for (const style of state.styles.filter((item) => item.styleNumber === styleNumber && variants.some((variant) => clean(variant.color).toLowerCase() === clean(item.color).toLowerCase()) && !item.image)) {
      style.image = await prepareStyleImageForState(uploadedImage || pastedStyleImageDataUrl || imageValue || "", style.id);
    }
  }

  resetStyleFormState();
  await persistState();
}

async function saveCutting(e) {
  e.preventDefault();
  const f = new FormData(els.cuttingForm);
  const sizes = getSizes();
  const entryId = els.cuttingForm.dataset.editId || uid();
  const payload = {
    id: entryId,
    date: normalizeDateValue(f.get("date")),
    styleId: f.get("styleId"),
    service: clean(f.get("service")),
    remarks: clean(f.get("remarks")),
    quantities: Object.fromEntries(sizes.map((size) => [size, num(els.cuttingSizeRows.querySelector(`[name="cut_${size}"]`).value)]))
  };
  upsertEntry("cuttingEntries", payload, entryId);
  delete els.cuttingForm.dataset.editId;
  els.cuttingForm.reset();
  setToday();
  await persistState();
}

async function saveStyleProduction(e) {
  e.preventDefault();
  const f = new FormData(els.styleProductionForm);
  const sizes = getSizes();
  const entryId = els.styleProductionForm.dataset.editId || uid();
  const quantities = Object.fromEntries(sizes.map((size) => [size, num(els.styleProductionSizeRows.querySelector(`[name="styleProd_${size}"]`).value)]));
  const totalQty = num(f.get("totalQty"));
  const payload = {
    id: entryId,
    date: normalizeDateValue(f.get("date")),
    styleId: f.get("styleId"),
    remarks: clean(f.get("remarks")),
    totalQty,
    quantities
  };
  upsertEntry("styleProductionEntries", payload, entryId);
  delete els.styleProductionForm.dataset.editId;
  els.styleProductionForm.reset();
  setToday();
  await persistState();
}

async function saveProduction(e) {
  e.preventDefault();
  const f = new FormData(els.productionForm);
  const style = byId(f.get("styleId"));
  const operationName = f.get("operationName");
  const op = style?.operations.find((x) => x.operationName === operationName);
  const entryId = els.productionForm.dataset.editId || uid();
  const payload = {
    id: entryId,
    date: normalizeDateValue(f.get("date")),
    styleId: f.get("styleId"),
    operationName,
    operationRate: op ? num(op.rate) : 0,
    workerName: clean(f.get("workerName")),
    size: f.get("size"),
    quantity: num(f.get("quantity")),
    remarks: clean(f.get("remarks"))
  };
  upsertEntry("productionEntries", payload, entryId);
  delete els.productionForm.dataset.editId;
  els.productionForm.reset();
  setToday();
  await persistState();
}

async function saveAcceptance(e) {
  e.preventDefault();
  const f = new FormData(els.acceptanceForm);
  const sizes = getSizes();
  const entryId = els.acceptanceForm.dataset.editId || uid();
  const payload = {
    id: entryId,
    date: normalizeDateValue(f.get("date")),
    styleId: f.get("styleId"),
    remarks: clean(f.get("remarks")),
    items: sizes.map((size) => ({
      size,
      accepted: num(els.acceptanceSizeRows.querySelector(`[name="accepted_${size}"]`).value),
      rejected: num(els.acceptanceSizeRows.querySelector(`[name="rejected_${size}"]`).value)
    }))
  };
  upsertEntry("acceptanceEntries", payload, entryId);
  delete els.acceptanceForm.dataset.editId;
  els.acceptanceForm.reset();
  setToday();
  await persistState();
}

async function saveDispatch(e) {
  e.preventDefault();
  const f = new FormData(els.dispatchForm);
  const sizes = getSizes();
  const entryId = els.dispatchForm.dataset.editId || uid();
  const payload = {
    id: entryId,
    date: normalizeDateValue(f.get("date")),
    styleId: f.get("styleId"),
    remarks: clean(f.get("remarks")),
    quantities: Object.fromEntries(sizes.map((size) => [size, num(els.dispatchSizeRows.querySelector(`[name="dispatch_${size}"]`).value)]))
  };
  upsertEntry("dispatchEntries", payload, entryId);
  delete els.dispatchForm.dataset.editId;
  els.dispatchForm.reset();
  setToday();
  await persistState();
}

async function saveWashcare(e) {
  e.preventDefault();
  const f = new FormData(els.washcareForm);
  const styleId = clean(f.get("styleId"));
  if (!styleId) {
    alert("Please select a style for washcare.");
    return;
  }
  const editId = els.washcareForm.dataset.editId || "";
  const existing = state.washcareRecords.find((record) => record.styleId === styleId && record.id !== editId);
  const templatePath = clean(f.get("templatePath"));
  const manualPrintCommand = clean(f.get("printCommand"));
  const payload = {
    id: editId || uid(),
    styleId,
    reportNumber: clean(f.get("reportNumber")),
    reportDate: normalizeDateValue(f.get("reportDate")),
    labName: clean(f.get("labName")),
    templateName: clean(f.get("templateName")),
    templatePath,
    printMethod: clean(f.get("printMethod")) || "browser",
    printCommand: manualPrintCommand || generateWashcarePrintCommand({ templatePath }),
    labelWidthMm: clampWashcareDimension(f.get("labelWidthMm"), DEFAULT_WASHCARE_LABEL_WIDTH_MM),
    labelLengthMm: clampWashcareDimension(f.get("labelLengthMm"), DEFAULT_WASHCARE_LABEL_LENGTH_MM),
    reportSourceName: clean(f.get("reportSourceName")),
    reportStyleCode: clean(f.get("reportStyleCode")),
    reportColor: clean(f.get("reportColor")),
    reportBrand: clean(f.get("reportBrand")),
    reportSupplier: clean(f.get("reportSupplier")),
    composition: clean(f.get("composition")),
    compositionFontSize: clampWashcareSize(f.get("compositionFontSize"), 22),
    careSymbols: parseWashcareSymbols(f.get("careSymbols")),
    symbolSize: clampWashcareSize(f.get("symbolSize"), 30),
    symbolGap: clampWashcareSize(f.get("symbolGap"), 6),
    washcareText: clean(f.get("washcareText")),
    washcareFontSize: clampWashcareSize(f.get("washcareFontSize"), 16),
    originLine: clean(f.get("originLine")) || "MADE IN INDIA",
    footerLine1: clean(f.get("footerLine1")),
    footerLine2: clean(f.get("footerLine2")),
    notes: clean(f.get("notes")),
    updatedAt: new Date().toISOString()
  };

  if (existing) {
    const confirmed = window.confirm("Washcare already exists for this style. Do you want to replace it?");
    if (!confirmed) return;
    payload.id = existing.id;
  }

  upsertEntry("washcareRecords", payload, payload.id);
  delete els.washcareForm.dataset.editId;
  await persistState();
  populateWashcareFormByStyle(styleId);
}

async function savePayment(e) {
  e.preventDefault();
  const reportRange = getReportDateFilter();
  if (!reportRange.startDate || !reportRange.endDate) {
    alert("Please select both From Date and To Date before saving payment.");
    return;
  }
  const billingRows = styleBillingRows(reportRange).filter((row) => row.billing > 0);
  if (!billingRows.length) {
    alert("No billing rows found for the selected date range.");
    return;
  }

  const f = new FormData(els.paymentForm);
  const baseAmountPaid = num(f.get("baseAmountPaid"));
  const serviceChargePaid = num(f.get("serviceChargePaid"));
  const totalPaid = num(f.get("totalPaid"));
  const existingIndex = state.payments.findIndex((payment) => isSameReportRange(payment, reportRange));
  const payment = {
    id: existingIndex >= 0 ? state.payments[existingIndex].id : uid(),
    paymentDate: normalizeDateValue(f.get("paymentDate")),
    startDate: reportRange.startDate,
    endDate: reportRange.endDate,
    styleIds: billingRows.map((row) => row.styleId),
    baseAmountPaid,
    serviceChargePaid,
    totalPaid,
    notes: clean(f.get("notes"))
  };

  if (!payment.paymentDate) {
    alert("Please enter payment date.");
    return;
  }

  if (existingIndex >= 0) state.payments[existingIndex] = payment;
  else state.payments.push(payment);

  await persistState();
  alert(existingIndex >= 0 ? "Payment updated for the selected date range." : "Payment saved for the selected date range.");
}

function render() {
  renderStyleSelects();
  renderOperationSelect();
  renderStyles();
  renderCutting();
  renderStyleProduction();
  renderProduction();
  renderAcceptance();
  renderDispatch();
  renderGrnStatus();
  renderWashcare();
  renderDashboard();
  renderReports();
  renderTallyForm();
}

function renderStyleSelects() {
  renderStyleSelect(els.cuttingStyleSelect, els.cuttingStyleSearch?.value || "");
  renderStyleSelect(els.styleProductionStyleSelect, els.styleProductionStyleSearch?.value || "");
  renderStyleSelect(els.productionStyleSelect, els.productionStyleSearch?.value || "");
  renderStyleSelect(els.acceptanceStyleSelect, els.acceptanceStyleSearch?.value || "");
  renderStyleSelect(els.dispatchStyleSelect, els.dispatchStyleSearch?.value || "");
  renderStyleSelect(els.washcareStyleSelect, els.washcareStyleSearch?.value || "");
}

function renderStyleSelect(select, searchTerm = "") {
  const current = select.value;
  const filteredStyles = filterStyles(searchTerm);
  const emptyLabel = !state.styles.length
    ? "No style created"
    : filteredStyles.length
      ? "Select style"
      : "No matching style";
  select.innerHTML = `<option value="">${emptyLabel}</option>` +
    filteredStyles.map((s) => `<option value="${s.id}">${esc(styleLabel(s))}</option>`).join("");
  if ([...select.options].some((o) => o.value === current)) {
    select.value = current;
  } else if (filteredStyles.length === 1) {
    select.value = filteredStyles[0].id;
  }
}

function renderOperationSelect() {
  const style = byId(els.productionStyleSelect.value);
  const current = els.productionOperationSelect.value;
  els.productionOperationSelect.innerHTML = `<option value="">${style ? "Select operation" : "Select style first"}</option>` +
    (style ? style.operations.map((o) => `<option value="${escAttr(o.operationName)}">${esc(o.operationName)} (Rs ${fmt(o.rate)})</option>`).join("") : "");
  if ([...els.productionOperationSelect.options].some((o) => o.value === current)) els.productionOperationSelect.value = current;
}

function renderStyles() {
  const filteredStyles = filterStyles(els.styleCardSearch?.value || "");
  updateStyleImageDiagnosticStatus();
  els.styleCards.innerHTML = filteredStyles.length ? filteredStyles.map((style) => `
    <article class="style-card">
      ${styleImageSrc(style) ? `
        <div class="style-media">
          <img class="style-thumb" src="${escAttr(styleImageSrc(style))}" alt="${escAttr(style.styleNumber)}">
          <div class="style-preview-meta">
            <span class="chip">Full image available</span>
            <button type="button" class="ghost small preview-link" data-action="preview-image" data-image-src="${escAttr(styleImageSrc(style))}" data-image-title="${escAttr(styleLabel(style))}">Preview Full Image</button>
          </div>
        </div>` : ""}
      <h4>${esc(style.styleNumber)}${style.styleName ? ` - ${esc(style.styleName)}` : ""}</h4>
      <p><strong>Buyer:</strong> ${esc(style.buyerName || "-")}</p>
      <p><strong>Color:</strong> ${esc(style.color || "-")}</p>
      <p><strong>Order Qty:</strong> ${fmtInt(style.orderQty || 0)}</p>
      <p><strong>Total CMT:</strong> Rs ${fmt(style.cmtRate)}</p>
      <p><strong>Service Charge:</strong> ${fmt(style.serviceChargePct || 0)}%</p>
      <p><strong>Image:</strong> ${styleImageSrc(style) ? "Uploaded" : "-"}</p>
      <div class="chip-row">${style.operations.map((o) => `<span class="chip">${esc(o.operationName)}: Rs ${fmt(o.rate)}</span>`).join("") || "<span class='chip'>No operations</span>"}</div>
      <div class="card-actions">
        <button type="button" class="ghost small" data-action="edit-style" data-style-id="${style.id}">Edit</button>
        <button type="button" class="ghost small" data-action="delete-style" data-style-id="${style.id}">Delete</button>
      </div>
    </article>`).join("") : `<div class="empty-state">${state.styles.length ? "No styles match this search." : "No styles added yet."}</div>`;
}

function renderCutting() {
  renderCuttingEntriesHeader();
  const search = clean(els.cuttingEntriesSearch?.value).toLowerCase();
  const rows = state.cuttingEntries
    .slice()
    .reverse()
    .filter((entry) => {
      const style = byId(entry.styleId);
      return matchesTextSearch([
        entry.date,
        style?.styleNumber,
        style?.color,
        style?.styleName,
        entry.service,
        entry.remarks
      ], search);
    })
    .map((entry) => {
      const style = byId(entry.styleId);
      return `
    <tr><td>${esc(entry.date)}</td><td>${esc(styleLabel(style))}</td><td>${esc(entry.service || "-")}</td><td>${fmtInt(style?.orderQty || 0)}</td>${getSizes().map((size) => `<td>${fmtInt(entry.quantities?.[size] || 0)}</td>`).join("")}<td>${fmtInt(sumObj(entry.quantities))}</td><td>${esc(entry.remarks || "-")}</td><td><button type="button" class="ghost small" data-action="edit-cutting" data-entry-id="${entry.id}">Edit</button> <button type="button" class="ghost small" data-action="delete-cutting" data-entry-id="${entry.id}">Delete</button></td></tr>`;
    });
  els.cuttingEntriesTable.innerHTML = rowsOrEmpty(rows, getSizes().length + 6, "No cutting entries recorded.");
}

function renderStyleProduction() {
  const search = clean(els.styleProductionEntriesSearch?.value).toLowerCase();
  const rows = state.styleProductionEntries.slice().reverse().filter((entry) => {
    const style = byId(entry.styleId);
    return matchesTextSearch([entry.date, style?.styleNumber, style?.color, style?.styleName, entry.remarks], search);
  }).map((entry) => `
    <tr><td>${esc(entry.date)}</td><td>${esc(styleLabel(byId(entry.styleId)))}</td><td>${fmtInt(entryProducedQty(entry))}</td><td>${esc(formatProductionQuantities(entry))}</td><td>${esc(entry.remarks || "-")}</td><td><button type="button" class="ghost small" data-action="edit-style-production" data-entry-id="${entry.id}">Edit</button> <button type="button" class="ghost small" data-action="delete-style-production" data-entry-id="${entry.id}">Delete</button></td></tr>`);
  els.styleProductionEntriesTable.innerHTML = rowsOrEmpty(rows, 6, "No style production entries recorded.");
}

function renderProduction() {
  els.productionEntriesTable.innerHTML = rowsOrEmpty(state.productionEntries.slice().reverse().map((entry) => `
    <tr><td>${esc(entry.date)}</td><td>${esc(styleLabel(byId(entry.styleId)))}</td><td>${esc(entry.operationName)}</td><td>${esc(entry.workerName)}</td><td>${esc(entry.size)}</td><td>${fmtInt(entry.quantity)}</td><td>${esc(entry.remarks || "-")}</td><td><button type="button" class="ghost small" data-action="edit-production" data-entry-id="${entry.id}">Edit</button> <button type="button" class="ghost small" data-action="delete-production" data-entry-id="${entry.id}">Delete</button></td></tr>`), 8, "No production entries recorded.");
}
function renderAcceptance() {
  els.acceptanceEntriesTable.innerHTML = rowsOrEmpty(state.acceptanceEntries.slice().reverse().map((entry) => `
    <tr><td>${esc(entry.date)}</td><td>${esc(styleLabel(byId(entry.styleId)))}</td><td>${fmtInt(entry.items.reduce((s, i) => s + i.accepted, 0))}</td><td>${fmtInt(entry.items.reduce((s, i) => s + i.rejected, 0))}</td><td>${esc(formatAcceptanceItems(entry.items))}</td><td>${esc(entry.remarks || "-")}</td><td><button type="button" class="ghost small" data-action="edit-acceptance" data-entry-id="${entry.id}">Edit</button></td></tr>`), 7, "No acceptance entries recorded.");
}

function renderDispatch() {
  const search = clean(els.dispatchEntriesSearch?.value).toLowerCase();
  const rows = state.dispatchEntries.slice().reverse().filter((entry) => {
    const style = byId(entry.styleId);
    return matchesTextSearch([entry.date, style?.styleNumber, style?.color, style?.styleName, entry.remarks], search);
  }).map((entry) => `
    <tr><td>${esc(entry.date)}</td><td>${esc(styleLabel(byId(entry.styleId)))}</td><td>${fmtInt(sumObj(entry.quantities))}</td><td>${esc(formatQuantities(entry.quantities))}</td><td>${esc(entry.remarks || "-")}</td><td><button type="button" class="ghost small" data-action="edit-dispatch" data-entry-id="${entry.id}">Edit</button> <button type="button" class="ghost small" data-action="delete-dispatch" data-entry-id="${entry.id}">Delete</button></td></tr>`);
  els.dispatchEntriesTable.innerHTML = rowsOrEmpty(rows, 6, "No dispatch entries recorded.");
}

function renderGrnStatus() {
  const grn = state.grnStatus || clone(DEFAULTS.grnStatus);
  if (els.grnStatusText) {
    const itemCount = (grn.items || []).length;
    const itemSummary = itemCount
      ? ` (${fmtInt(itemCount)} items${grn.pdfPageCount ? ` across ${fmtInt(grn.pdfPageCount)} pages` : ""})`
      : "";
    els.grnStatusText.textContent = grn.grnFileName
      ? `Loaded GRN ${grn.grnNumber || "-"} from ${grn.grnFileName}${itemSummary}${grn.packingListFileName ? ` and packing list ${grn.packingListFileName}` : ""}.`
      : "Upload a GRN PDF to extract the complete detail sheet.";
  }
  if (els.grnHeaderSummary) {
    els.grnHeaderSummary.classList.toggle("empty-state", !grn.grnFileName);
    els.grnHeaderSummary.innerHTML = grn.grnFileName ? buildGrnHeaderSummaryMarkup(grn) : "No GRN uploaded yet.";
  }
  if (els.grnSummaryCards) {
    const cards = [
      { label: "GRN Challan Qty", value: fmtInt(grn.totalChallanQty || 0) },
      { label: "Received Qty", value: fmtInt(grn.totalReceivedQty || 0) },
      { label: "Accepted Qty", value: fmtInt(grn.totalAcceptedQty || 0) },
      { label: "Short Qty", value: fmtInt(grn.totalShortQty || 0) },
      { label: "Packing List Qty", value: fmtInt((grn.comparisonRows || []).reduce((sum, row) => sum + num(row.packingQty), 0)) },
      { label: "Invoice Shortage", value: fmtInt((grn.comparisonRows || []).reduce((sum, row) => sum + Math.max(num(row.shortageQty), 0), 0)) }
    ];
    els.grnSummaryCards.innerHTML = cards.map((card) => `<article class="stat-card"><p>${esc(card.label)}</p><strong>${esc(card.value)}</strong></article>`).join("");
  }
  if (els.grnItemsTable) {
    const rows = (grn.items || []).map((item) => `
      <tr>
        <td>${fmtInt(item.serialNo)}</td>
        <td>${esc(item.article || "-")}</td>
        <td>${esc(item.description || "-")}</td>
        <td>${esc(item.ean || "-")}</td>
        <td>${esc(item.vendorArticleNo || "-")}</td>
        <td>${esc(item.uom || "-")}</td>
        <td>${fmtInt(item.challanQty)}</td>
        <td>${fmtInt(item.receivedQty)}</td>
        <td>${fmtInt(item.acceptedQty)}</td>
        <td>${fmtInt(item.shortQty)}</td>
        <td>${esc(item.reason || "-")}</td>
      </tr>`);
    els.grnItemsTable.innerHTML = rowsOrEmpty(rows, 11, "No GRN detail sheet available yet.");
  }
  if (els.grnComparisonTable) {
    const rows = (grn.comparisonRows || []).map((row) => `
      <tr>
        <td>${esc(row.matchLabel || "-")}</td>
        <td>${fmtInt(row.packingQty)}</td>
        <td>${fmtInt(row.challanQty)}</td>
        <td>${fmtInt(row.acceptedQty)}</td>
        <td>${fmtInt(row.shortageQty)}</td>
        <td><span class="status-chip ${num(row.shortageQty) > 0 ? "pending" : "paid"}">${esc(num(row.shortageQty) > 0 ? "Short" : "Matched")}</span></td>
      </tr>`);
    els.grnComparisonTable.innerHTML = rowsOrEmpty(rows, 6, "Upload the packing list later to compare invoice quantity with GRN accepted quantity.");
  }
  if (els.grnSavedTable) {
    const rows = (state.grnReports || [])
      .slice()
      .sort((a, b) => clean(b.updatedAt).localeCompare(clean(a.updatedAt)))
      .map((report) => `
        <tr>
          <td>${esc(report.grnNumber || "-")}</td>
          <td>${esc(report.grnDate || "-")}</td>
          <td>${esc(report.vendorInvoiceNo || "-")}</td>
          <td>${esc(report.poNumber || "-")}</td>
          <td>${esc(report.supplier || "-")}</td>
          <td>${fmtInt(report.totalAcceptedQty || 0)}</td>
          <td>${fmtInt(report.totalShortQty || 0)}</td>
          <td>${esc(formatDateTimeDisplay(report.updatedAt) || "-")}</td>
        </tr>`);
    els.grnSavedTable.innerHTML = rowsOrEmpty(rows, 8, "No saved GRN reports yet.");
  }
}

function buildGrnHeaderSummaryMarkup(grn) {
  return `
    <div class="grn-meta-grid">
      <div><strong>GRN No.</strong><span>${esc(grn.grnNumber || "-")}</span></div>
      <div><strong>GRN Date</strong><span>${esc(grn.grnDate || "-")}</span></div>
      <div><strong>PDF Pages</strong><span>${esc(fmtInt(grn.pdfPageCount || 0))}</span></div>
      <div><strong>Item Count</strong><span>${esc(fmtInt((grn.items || []).length))}</span></div>
      <div><strong>Supplier</strong><span>${esc(grn.supplier || "-")}</span></div>
      <div><strong>Vendor Code</strong><span>${esc(grn.supplierCode || "-")}</span></div>
      <div><strong>Invoice No.</strong><span>${esc(grn.vendorInvoiceNo || "-")}</span></div>
      <div><strong>PO No.</strong><span>${esc(grn.poNumber || "-")}</span></div>
      <div><strong>Delivery No.</strong><span>${esc(grn.deliveryNo || "-")}</span></div>
      <div><strong>Warehouse</strong><span>${esc(grn.warehouse || grn.consignee || "-")}</span></div>
      <div><strong>Transporter</strong><span>${esc(grn.transporterName || "-")}</span></div>
      <div><strong>Vehicle No.</strong><span>${esc(grn.vehicleNumber || "-")}</span></div>
      <div><strong>Consignment</strong><span>${esc(grn.consignmentNote || "-")}</span></div>
      <div><strong>Packing List</strong><span>${esc(grn.packingListFileName || "-")}</span></div>
    </div>
  `;
}

function renderWashcare() {
  renderWashcarePreview();
  if (!els.washcareTable) return;
  const search = clean(els.washcareSearch?.value).toLowerCase();
  const rows = state.washcareRecords
    .slice()
    .sort((a, b) => clean(b.updatedAt).localeCompare(clean(a.updatedAt)))
    .filter((record) => {
      const style = byId(record.styleId);
      return matchesTextSearch([
        style?.styleNumber,
        style?.color,
        style?.styleName,
        record.reportNumber,
        record.composition,
        record.templateName
      ], search);
    })
    .map((record) => {
      const style = byId(record.styleId);
      return `
        <tr>
          <td>${esc(styleLabel(style))}</td>
          <td>${esc(record.reportNumber || "-")}</td>
          <td>${esc(record.reportDate || "-")}</td>
          <td>${esc(record.templateName || record.templatePath || "-")}</td>
          <td>${esc(formatWashcarePrintMethod(record.printMethod))}</td>
          <td>${esc(formatDateTimeDisplay(record.updatedAt) || "-")}</td>
          <td><button type="button" class="ghost small" data-action="edit-washcare" data-entry-id="${record.id}">Edit</button> <button type="button" class="ghost small" data-action="print-washcare" data-entry-id="${record.id}">Print</button> <button type="button" class="ghost small" data-action="delete-washcare" data-entry-id="${record.id}">Delete</button></td>
        </tr>`;
    });
  els.washcareTable.innerHTML = rowsOrEmpty(rows, 7, "No washcare saved yet.");
}

function renderWashcarePreview() {
  if (!els.washcarePreview) return;
  const preview = getWashcarePreviewData();
  if (!preview.style) {
    els.washcarePreview.classList.add("empty-state");
    els.washcarePreview.innerHTML = "Select a style and enter washcare details to preview the label.";
    return;
  }
  els.washcarePreview.classList.remove("empty-state");
  els.washcarePreview.innerHTML = buildWashcarePreviewMarkup(preview);
}

function renderDashboard() {
  const summary = dashboardSummary();
  els.dashboardStats.innerHTML = [
    ["Styles", state.styles.length],
    ["Total Order Qty", summary.totalOrderQty],
    ["Total Cut Qty", summary.totalCutQty],
    ["Total Make Qty", summary.totalMakeQty],
    ["Total Dispatch Qty", summary.totalDispatchQty],
    ["Total Billed (Rs)", fmtInt(Math.round(summary.totalBilled))],
    ["Total Paid (Rs)", summary.totalPaid],
    ["Service Charge Paid (Rs)", summary.totalServiceChargePaid],
    ["Short Cut vs Order %", `${fmt(summary.shortCutPct)}%`],
    ["Ship vs Make %", `${fmt(summary.shipAgainstMakePct)}%`]
  ].map(([label, value]) => `<div class="stat-card"><p>${label}</p><strong>${value}</strong></div>`).join("");

  const billing = styleBillingRows();
  els.styleBillingTable.innerHTML = rowsOrEmpty(billing.map((r) => `
    <tr><td>${esc(r.styleNumber)}</td><td>${esc(r.color || "-")}</td><td>${fmtInt(r.acceptedQty)}</td><td>Rs ${fmt(r.cmtRate)}</td><td>Rs ${fmt(r.billing)}</td></tr>`), 5, "No style billing yet.");

  const workers = workerBillingRows();
  els.workerBillingTable.innerHTML = rowsOrEmpty(workers.map((r) => `
    <tr><td>${esc(r.workerName)}</td><td>${esc(r.operationName)}</td><td>${fmtInt(r.quantity)}</td><td>Rs ${fmt(r.amount)}</td></tr>`), 4, "No worker billing yet.");
  renderPendingWorkflow();
}

function renderReports() {
  renderCuttingReportHeader();
  const reportRange = getReportDateFilter();
  const rangeSummary = reportRange.startDate && reportRange.endDate
    ? `Showing report dates from ${formatDateDisplay(reportRange.startDate)} to ${formatDateDisplay(reportRange.endDate)}.`
    : reportRange.startDate
      ? `Showing report dates from ${formatDateDisplay(reportRange.startDate)} onwards.`
      : reportRange.endDate
        ? `Showing report dates up to ${formatDateDisplay(reportRange.endDate)}.`
        : "Showing all report dates.";
  if (els.reportRangeSummary) {
    els.reportRangeSummary.textContent = rangeSummary;
  }

  const billingSearch = clean(els.billingReportSearch?.value).toLowerCase();
  const styleAmounts = styleBillingRows(reportRange).filter((r) => matchesTextSearch([r.styleNumber, r.color, r.paymentStatusLabel, r.dateLabel], billingSearch));
  els.styleAmountReportTable.innerHTML = rowsOrEmpty(styleAmounts.map((r) => `
    <tr>
      <td>${esc(r.dateLabel)}</td>
      <td>${esc(r.styleNumber)}</td>
      <td>${esc(r.color || "-")}</td>
      <td>${r.image ? `<img class="report-thumb" src="${escAttr(r.image)}" alt="${escAttr(r.styleNumber)}" data-action="preview-image" data-image-src="${escAttr(r.image)}" data-image-title="${escAttr(r.styleNumber)}">` : "-"}</td>
      <td>${fmtInt(r.orderQty)}</td>
      <td>${fmtInt(r.cutQty)}</td>
      <td>${fmtInt(r.producedQty)}</td>
      <td>${fmtInt(r.dispatchQty)}</td>
      <td>${esc(r.cutVsMakeSummary)}</td>
      <td>${esc(r.makeVsDispatchSummary)}</td>
      <td>Rs ${fmt(r.cmtRate)}</td>
      <td>Rs ${fmt(r.baseAmount)}</td>
      <td>Rs ${fmt(r.serviceChargeAmount)}</td>
      <td>Rs ${fmt(r.billing)}</td>
      <td><span class="status-chip ${r.paymentStatusClass}">${esc(r.paymentStatusLabel)}</span></td>
    </tr>`), 15, "No style amount report for the selected date range.");

  const reconciliationSearch = clean(els.reconciliationSearch?.value).toLowerCase();
  const reconciliation = reconciliationRows(reportRange).filter((r) => matchesTextSearch([r.styleNumber, r.color], reconciliationSearch));
  els.reconciliationTable.innerHTML = rowsOrEmpty(reconciliation.map((r) => `
    <tr><td>${esc(r.styleNumber)}</td><td>${fmtInt(r.cutQty)}</td><td>${fmtInt(r.producedQty)}</td><td>${fmtInt(r.acceptedQty)}</td><td>${fmtInt(r.rejectedQty)}</td><td class="${r.balance < 0 ? "text-danger" : "text-success"}">${fmtInt(r.balance)}</td></tr>`), 6, "No reconciliation data yet.");

  const cuttingSearch = clean(els.cuttingReportSearch?.value).toLowerCase();
  const cuttingRows = cuttingReportRows(reportRange).filter((r) => matchesTextSearch([r.date, r.styleNumber, r.color, r.service, r.remarks], cuttingSearch));
  els.cuttingReportTable.innerHTML = rowsOrEmpty(cuttingRows.map((r) => `
    <tr>
      <td>${esc(r.date)}</td>
      <td>${r.image ? `<img class="report-thumb" src="${escAttr(r.image)}" alt="${escAttr(r.styleNumber)}" data-action="preview-image" data-image-src="${escAttr(r.image)}" data-image-title="${escAttr(r.styleNumber)}">` : "-"}</td>
      <td>${esc(r.styleNumber)}</td>
      <td>${esc(r.color || "-")}</td>
      <td>${esc(r.service || "-")}</td>
      ${getSizes().map((size) => `<td>${fmtInt(r.quantities[size] || 0)}</td>`).join("")}
      <td>${fmtInt(r.totalQty)}</td>
      <td>${esc(r.remarks || "-")}</td>
    </tr>`), getSizes().length + 7, "No cutting report for the selected date range.");

  const dispatchSearch = clean(els.dispatchReportSearch?.value).toLowerCase();
  const dispatchRows = dispatchReportRows(reportRange).filter((r) => matchesTextSearch([r.styleNumber, r.color, r.size], dispatchSearch));
  els.dispatchReportTable.innerHTML = rowsOrEmpty(dispatchRows.map((r) => `
    <tr>
      <td>${r.image ? `<img class="report-thumb" src="${escAttr(r.image)}" alt="${escAttr(r.styleNumber)}" data-action="preview-image" data-image-src="${escAttr(r.image)}" data-image-title="${escAttr(r.styleNumber)}">` : "-"}</td>
      <td>${esc(r.styleNumber)}</td>
      <td>${esc(r.color || "-")}</td>
      <td>${esc(r.size)}</td>
      <td>${fmtInt(r.cutQty)}</td>
      <td>${fmtInt(r.makeQty)}</td>
      <td>${fmtInt(r.dispatchQty)}</td>
      <td>Rs ${fmt(r.amount)}</td>
      <td class="${r.balance < 0 ? "text-danger" : "text-success"}">${fmtInt(r.balance)}</td>
    </tr>`), 9, "No size-wise dispatch details yet.");

  const ops = operationCostRows(reportRange);
  els.operationCostTable.innerHTML = rowsOrEmpty(ops.map((r) => `
<tr><td>${esc(r.styleNumber)}</td><td>${esc(r.operationName)}</td><td>${fmtInt(r.quantity)}</td><td>Rs ${fmt(r.rate)}</td><td>Rs ${fmt(r.amount)}</td></tr>`), 5, "No operation costing yet.");

  renderPaymentHistory();
  populatePaymentFormForCurrentRange(styleAmounts, reportRange);
  renderTallyCreditors();
}

function renderTallyForm() {
  if (!els.tallySyncForm) return;
  const config = state.tallyCreditors || {};
  if (els.tallyEndpoint && !els.tallyEndpoint.matches(":focus")) els.tallyEndpoint.value = clean(config.endpoint) || getDefaultTallyEndpoint();
  if (els.tallyCompany && !els.tallyCompany.matches(":focus")) els.tallyCompany.value = clean(config.company);
  if (els.tallyFromDate && !els.tallyFromDate.matches(":focus")) els.tallyFromDate.value = clean(config.fromDate);
  if (els.tallyAsOnDate && !els.tallyAsOnDate.matches(":focus")) els.tallyAsOnDate.value = clean(config.asOnDate) || todayIso();
  if (els.tallyBucketDays && !els.tallyBucketDays.matches(":focus")) els.tallyBucketDays.value = clean(config.bucketDays) || "30,60,90";
}

async function saveTallyPreferences() {
  state.tallyCreditors = {
    ...state.tallyCreditors,
    ...getTallyFormValues()
  };
  await persistState();
}

function getTallyFormValues() {
  return {
    endpoint: clean(els.tallyEndpoint?.value) || getDefaultTallyEndpoint(),
    company: clean(els.tallyCompany?.value),
    fromDate: clean(els.tallyFromDate?.value),
    asOnDate: clean(els.tallyAsOnDate?.value) || todayIso(),
    bucketDays: clean(els.tallyBucketDays?.value) || "30,60,90"
  };
}

async function syncTallyCreditors(e) {
  e?.preventDefault?.();
  const config = getTallyFormValues();
  if (!config.endpoint) {
    alert("Please enter the Tally URL first.");
    return;
  }
  updateTallyStatus("Connecting to Tally and loading supplier vouchers...");
  try {
    const [ledgerXml, voucherXml] = await Promise.all([
      fetchTallyXml(config.endpoint, buildTallyLedgerRequest(config)),
      fetchTallyXml(config.endpoint, buildTallyVoucherRequest(config))
    ]);
    state.tallyCreditors = {
      ...state.tallyCreditors,
      ...config,
      ledgers: parseTallyLedgers(ledgerXml),
      vouchers: parseTallyVouchers(voucherXml),
      lastSyncAt: new Date().toISOString(),
      lastSyncSource: "tally-api",
      lastError: ""
    };
    await persistState();
    const report = getTallyAgeingReport();
    updateTallyStatus(`Loaded ${fmtInt(report.summaryRows.length)} suppliers from Tally on ${formatDateTimeDisplay(state.tallyCreditors.lastSyncAt)}.`);
  } catch (error) {
    state.tallyCreditors = {
      ...state.tallyCreditors,
      ...config,
      lastError: error.message || "Unable to read data from Tally.",
      lastSyncSource: state.tallyCreditors?.lastSyncSource || ""
    };
    await persistState();
    updateTallyStatus(`Direct Tally sync failed: ${state.tallyCreditors.lastError} Use "Import Tally XML" if browser access is blocked.`);
    alert("Direct Tally sync could not complete. If Tally blocks the browser request, export XML from Tally and use Import Tally XML.");
  }
}

async function testTallyConnection() {
  const config = getTallyFormValues();
  updateTallyStatus("Testing Tally connection...");
  try {
    const result = await testTallyEndpoint(config.endpoint);
    const message = result?.ok
      ? `Tally connection successful via proxy. Target: ${result.target || config.endpoint}`
      : `Tally connection failed. ${result?.error || "Unknown error."}`;
    updateTallyStatus(message);
    alert(message);
  } catch (error) {
    const message = error.message || "Tally connection test failed.";
    updateTallyStatus(message);
    alert(message);
  }
}

async function importTallyXmlFile(e) {
  const [file] = Array.from(e?.target?.files || []);
  if (!file) return;
  updateTallyStatus(`Importing ${file.name}...`);
  try {
    const text = await file.text();
    const ledgers = parseTallyLedgers(text);
    const vouchers = parseTallyVouchers(text);
    if (!ledgers.length && !vouchers.length) {
      throw new Error("The XML file did not contain ledger or voucher data.");
    }
    state.tallyCreditors = {
      ...state.tallyCreditors,
      ...getTallyFormValues(),
      ledgers: ledgers.length ? ledgers : state.tallyCreditors.ledgers,
      vouchers: vouchers.length ? vouchers : state.tallyCreditors.vouchers,
      lastSyncAt: new Date().toISOString(),
      lastSyncSource: `xml-import:${file.name}`,
      lastError: ""
    };
    await persistState();
    const report = getTallyAgeingReport();
    updateTallyStatus(`Imported ${fmtInt(report.summaryRows.length)} suppliers from ${file.name}.`);
  } catch (error) {
    state.tallyCreditors = {
      ...state.tallyCreditors,
      lastError: error.message || "XML import failed."
    };
    await persistState();
    updateTallyStatus(`XML import failed: ${state.tallyCreditors.lastError}`);
    alert(state.tallyCreditors.lastError);
  } finally {
    if (els.tallyXmlInput) els.tallyXmlInput.value = "";
  }
}

function renderTallyCreditors() {
  if (!els.tallyCreditorsTable || !els.tallyInvoiceDetailsTable) return;
  const report = getTallyAgeingReport();
  const search = clean(els.tallyPartySearch?.value).toLowerCase();
  const rows = report.summaryRows.filter((row) => !search || row.partyName.toLowerCase().includes(search));
  const selected = rows.find((row) => row.partyName === state.tallyCreditors.selectedParty) || rows[0] || null;
  if (els.tallyCurrentBucketHead) els.tallyCurrentBucketHead.textContent = report.bucketLabels[0] || "0-30 Days";
  if (els.tallyBucketHead) els.tallyBucketHead.textContent = report.bucketLabels.slice(1).join(" / ") || "31+ Days";
  if (els.tallySummaryCards) {
    els.tallySummaryCards.innerHTML = [
      ["Suppliers", fmtInt(report.summaryRows.length)],
      ["Outstanding (Rs)", fmt(report.totalOutstanding)],
      [`${report.bucketLabels.at(-1) || "Oldest"} (Rs)`, fmt(report.oldestBucketOutstanding)],
      ["Unapplied Advances (Rs)", fmt(report.unappliedCredits)],
      ["Last Sync", state.tallyCreditors.lastSyncAt ? formatDateTimeDisplay(state.tallyCreditors.lastSyncAt) : "Not synced"]
    ].map(([label, value]) => `<div class="stat-card"><p>${esc(label)}</p><strong>${esc(value)}</strong></div>`).join("");
  }
  els.tallyCreditorsTable.innerHTML = rowsOrEmpty(rows.map((row) => `
    <tr class="tally-summary-row ${selected?.partyName === row.partyName ? "active" : ""}" data-party="${escAttr(row.partyName)}">
      <td><strong>${esc(row.partyName)}</strong></td>
      <td>Rs ${fmt(row.totalOutstanding)}</td>
      <td>Rs ${fmt(row.bucketAmounts[0] || 0)}</td>
      <td><ul class="bucket-list">${row.bucketLabels.slice(1).map((label, index) => `<li><strong>${esc(label)}:</strong> Rs ${fmt(row.bucketAmounts[index + 1] || 0)}</li>`).join("") || "<li>-</li>"}</ul></td>
      <td>${esc(formatDateDisplay(row.lastPaymentDate) || "-")}</td>
      <td>${fmtInt(row.totalInvoices)}</td>
      <td><span class="tally-invoice-chip">${fmtInt(row.openInvoices.length)} open</span></td>
    </tr>`), 7, "No creditor ageing data loaded yet.");
  renderTallyInvoiceDetails(selected, report);
  const syncLabel = state.tallyCreditors.lastSyncSource
    ? `Last sync source: ${state.tallyCreditors.lastSyncSource}${state.tallyCreditors.lastSyncAt ? ` on ${formatDateTimeDisplay(state.tallyCreditors.lastSyncAt)}` : ""}.`
    : "Connect Tally and load supplier vouchers to build the ageing report.";
  updateTallyStatus(state.tallyCreditors.lastError ? `${syncLabel} ${state.tallyCreditors.lastError}` : syncLabel);
}

function renderTallyInvoiceDetails(selectedRow, report) {
  if (!els.tallyInvoiceDetailsTable || !els.tallySelectedPartyLabel) return;
  if (!selectedRow) {
    els.tallySelectedPartyLabel.textContent = "Select a supplier from the summary to view invoice-level details.";
    els.tallyInvoiceDetailsTable.innerHTML = rowsOrEmpty([], 9, "No invoice details available.");
    return;
  }
  els.tallySelectedPartyLabel.textContent = `${selectedRow.partyName} | Outstanding Rs ${fmt(selectedRow.totalOutstanding)} | ${fmtInt(selectedRow.openInvoices.length)} open invoices`;
  els.tallyInvoiceDetailsTable.innerHTML = rowsOrEmpty(selectedRow.openInvoices.map((invoice) => `
    <tr>
      <td>${esc(formatDateDisplay(invoice.invoiceDate) || "-")}</td>
      <td>${esc(invoice.voucherNumber || "-")}</td>
      <td>${esc(invoice.reference || "-")}</td>
      <td>Rs ${fmt(invoice.originalAmount)}</td>
      <td>Rs ${fmt(invoice.adjustedAmount)}</td>
      <td>Rs ${fmt(invoice.balanceAmount)}</td>
      <td>${fmtInt(invoice.ageDays)}</td>
      <td>${esc(invoice.bucketLabel)}</td>
      <td>${esc(formatDateDisplay(invoice.lastPaymentDate) || "-")}</td>
    </tr>`), 9, "No open invoice balances for this supplier.");
}

function handleTallyCreditorTableClick(e) {
  const row = e.target.closest("[data-party]");
  if (!row) return;
  state.tallyCreditors.selectedParty = row.dataset.party || "";
  renderTallyCreditors();
}

function getTallyAgeingReport() {
  const config = state.tallyCreditors || {};
  return buildTallyAgeingReport({
    asOnDate: clean(config.asOnDate) || todayIso(),
    bucketDays: parseBucketDays(config.bucketDays),
    ledgers: Array.isArray(config.ledgers) ? config.ledgers : [],
    vouchers: Array.isArray(config.vouchers) ? config.vouchers : []
  });
}

function buildTallyAgeingReport({ asOnDate, bucketDays, ledgers, vouchers }) {
  const creditorSet = new Set((ledgers || []).map((ledger) => normalizeKey(ledger.name || ledger.ledgerName || ledger)));
  const parties = new Map();
  (vouchers || []).forEach((voucher) => {
    const candidateEntries = (voucher.entries || []).filter((entry) => {
      const ledgerKey = normalizeKey(entry.ledgerName);
      if (!ledgerKey || !num(entry.amount)) return false;
      if (creditorSet.size && creditorSet.has(ledgerKey)) return true;
      return !!entry.isParty || ledgerKey === normalizeKey(voucher.partyLedgerName);
    });
    candidateEntries.forEach((entry) => collectTallyPartyEntry(parties, voucher, entry));
  });

  const bucketLabels = buildAgeBucketLabels(bucketDays);
  const summaryRows = [...parties.values()]
    .map((party) => finalizePartyAgeing(party, asOnDate, bucketDays, bucketLabels))
    .filter((row) => row.totalOutstanding > 0 || row.unappliedCredit > 0)
    .sort((a, b) => b.totalOutstanding - a.totalOutstanding || a.partyName.localeCompare(b.partyName));

  return {
    asOnDate,
    bucketDays,
    bucketLabels,
    summaryRows,
    totalOutstanding: summaryRows.reduce((sum, row) => sum + row.totalOutstanding, 0),
    unappliedCredits: summaryRows.reduce((sum, row) => sum + row.unappliedCredit, 0),
    oldestBucketOutstanding: summaryRows.reduce((sum, row) => sum + (row.bucketAmounts.at(-1) || 0), 0)
  };
}

function collectTallyPartyEntry(parties, voucher, entry) {
  const partyName = clean(entry.ledgerName || voucher.partyLedgerName);
  if (!partyName) return;
  const party = parties.get(partyName) || {
    partyName,
    invoices: [],
    adjustments: [],
    onAccountCredits: [],
    lastPaymentDate: "",
    totalInvoices: 0
  };
  const entryDirection = num(entry.amount) > 0 ? 1 : num(entry.amount) < 0 ? -1 : 0;
  if (entryDirection < 0) {
    party.lastPaymentDate = maxIsoDate(party.lastPaymentDate, voucher.date);
  }
  const allocations = entry.billAllocations?.length
    ? entry.billAllocations
    : [{ name: voucher.voucherNumber || voucher.masterId || "On Account", billType: entryDirection >= 0 ? "New Ref" : "On Account", amount: Math.abs(num(entry.amount)), dueDate: "" }];

  allocations.forEach((allocation) => {
    const amount = Math.abs(num(allocation.amount)) || Math.abs(num(entry.amount));
    if (!amount) return;
    const type = classifyBillAllocation(allocation.billType, entryDirection);
    if (type === "invoice") {
      party.invoices.push({
        id: `${voucher.id || voucher.masterId || voucher.voucherNumber}-${allocation.name || party.invoices.length}`,
        invoiceDate: voucher.date,
        voucherNumber: voucher.voucherNumber,
        reference: clean(allocation.name) || clean(voucher.voucherNumber) || clean(voucher.masterId),
        originalAmount: amount,
        adjustedAmount: 0,
        balanceAmount: amount,
        dueDate: clean(allocation.dueDate),
        lastPaymentDate: ""
      });
      party.totalInvoices += 1;
      return;
    }
    if (type === "agst-ref") {
      party.adjustments.push({
        reference: clean(allocation.name),
        amount,
        paymentDate: voucher.date,
        voucherNumber: voucher.voucherNumber
      });
      return;
    }
    party.onAccountCredits.push({
      amount,
      paymentDate: voucher.date,
      voucherNumber: voucher.voucherNumber,
      reference: clean(allocation.name)
    });
  });
  parties.set(partyName, party);
}

function finalizePartyAgeing(party, asOnDate, bucketDays, bucketLabels) {
  const invoices = mergePartyInvoices(party.invoices).sort((a, b) => `${clean(a.invoiceDate)}${clean(a.reference)}`.localeCompare(`${clean(b.invoiceDate)}${clean(b.reference)}`));
  const credits = party.onAccountCredits
    .slice()
    .sort((a, b) => `${clean(a.paymentDate)}${clean(a.reference)}`.localeCompare(`${clean(b.paymentDate)}${clean(b.reference)}`));

  party.adjustments
    .slice()
    .sort((a, b) => `${clean(a.paymentDate)}${clean(a.reference)}`.localeCompare(`${clean(b.paymentDate)}${clean(b.reference)}`))
    .forEach((adjustment) => {
      let remaining = adjustment.amount;
      const targets = invoices.filter((invoice) => normalizeKey(invoice.reference) === normalizeKey(adjustment.reference) || normalizeKey(invoice.voucherNumber) === normalizeKey(adjustment.reference));
      for (const invoice of targets) {
        remaining = applyInvoiceAdjustment(invoice, remaining, adjustment.paymentDate);
        if (remaining <= 0.0001) break;
      }
      if (remaining > 0.0001) {
        credits.push({
          amount: remaining,
          paymentDate: adjustment.paymentDate,
          voucherNumber: adjustment.voucherNumber,
          reference: adjustment.reference
        });
      }
    });

  credits.forEach((credit) => {
    let remaining = credit.amount;
    for (const invoice of invoices) {
      if (remaining <= 0.0001) break;
      remaining = applyInvoiceAdjustment(invoice, remaining, credit.paymentDate);
    }
    credit.remainingAmount = remaining;
  });

  const openInvoices = invoices
    .filter((invoice) => invoice.balanceAmount > 0.0001)
    .map((invoice) => {
      const ageDays = diffDays(invoice.dueDate || invoice.invoiceDate, asOnDate);
      const bucketIndex = resolveBucketIndex(ageDays, bucketDays);
      return {
        ...invoice,
        ageDays,
        bucketIndex,
        bucketLabel: bucketLabels[bucketIndex] || bucketLabels.at(-1) || "Open"
      };
    })
    .sort((a, b) => b.ageDays - a.ageDays || clean(a.invoiceDate).localeCompare(clean(b.invoiceDate)));

  const bucketAmounts = bucketLabels.map(() => 0);
  openInvoices.forEach((invoice) => {
    bucketAmounts[invoice.bucketIndex] += invoice.balanceAmount;
  });

  return {
    partyName: party.partyName,
    lastPaymentDate: party.lastPaymentDate,
    totalInvoices: party.totalInvoices || invoices.length,
    totalOutstanding: openInvoices.reduce((sum, invoice) => sum + invoice.balanceAmount, 0),
    openInvoices,
    bucketLabels,
    bucketAmounts,
    unappliedCredit: credits.reduce((sum, credit) => sum + num(credit.remainingAmount), 0)
  };
}

function mergePartyInvoices(invoices) {
  const map = new Map();
  invoices.forEach((invoice) => {
    const key = [clean(invoice.invoiceDate), normalizeKey(invoice.reference), normalizeKey(invoice.voucherNumber)].join("__");
    const current = map.get(key) || { ...invoice };
    if (current !== invoice) {
      current.originalAmount += num(invoice.originalAmount);
      current.balanceAmount += num(invoice.balanceAmount);
    }
    map.set(key, current);
  });
  return [...map.values()];
}

function applyInvoiceAdjustment(invoice, amount, paymentDate) {
  const applied = Math.min(num(amount), num(invoice.balanceAmount));
  if (applied <= 0.0001) return num(amount);
  invoice.adjustedAmount += applied;
  invoice.balanceAmount -= applied;
  invoice.lastPaymentDate = maxIsoDate(invoice.lastPaymentDate, paymentDate);
  return num(amount) - applied;
}

function classifyBillAllocation(billType, entryDirection) {
  const type = normalizeKey(billType);
  if (type.includes("agst")) return "agst-ref";
  if (type.includes("onaccount") || type.includes("advance")) return entryDirection >= 0 ? "invoice" : "credit";
  if (type.includes("newref")) return entryDirection >= 0 ? "invoice" : "credit";
  return entryDirection >= 0 ? "invoice" : "credit";
}

function parseBucketDays(text) {
  const values = String(text || "30,60,90")
    .split(",")
    .map((value) => parseInt(value.trim(), 10))
    .filter((value) => Number.isFinite(value) && value > 0);
  const normalized = [...new Set(values)].sort((a, b) => a - b);
  return normalized.length ? normalized : [30, 60, 90];
}

function buildAgeBucketLabels(bucketDays) {
  const limits = bucketDays.length ? bucketDays : [30, 60, 90];
  return limits.map((limit, index) => {
    const start = index === 0 ? 0 : limits[index - 1] + 1;
    return `${start}-${limit} Days`;
  }).concat(`${limits.at(-1) + 1}+ Days`);
}

function resolveBucketIndex(ageDays, bucketDays) {
  const age = Math.max(num(ageDays), 0);
  const limits = bucketDays.length ? bucketDays : [30, 60, 90];
  const index = limits.findIndex((limit) => age <= limit);
  return index >= 0 ? index : limits.length;
}

async function downloadTallyAgeingWorkbook() {
  const report = getTallyAgeingReport();
  if (!report.summaryRows.length) {
    alert("No creditor ageing data found to export.");
    return;
  }
  const workbook = createWorkbookOrAlert();
  if (!workbook) return;
  const summarySheet = workbook.addWorksheet("Creditor Summary");
  summarySheet.columns = [
    { header: "Party Name", key: "partyName", width: 28 },
    { header: "Total Payable", key: "totalOutstanding", width: 16 },
    ...report.bucketLabels.map((label, index) => ({ header: label, key: `bucket_${index}`, width: 14 })),
    { header: "Last Payment", key: "lastPaymentDate", width: 16 },
    { header: "Open Invoices", key: "openInvoices", width: 14 },
    { header: "Total Invoices", key: "totalInvoices", width: 14 },
    { header: "Unapplied Advance", key: "unappliedCredit", width: 16 }
  ];
  styleWorksheetHeader(summarySheet);
  report.summaryRows.forEach((row) => {
    summarySheet.addRow({
      partyName: row.partyName,
      totalOutstanding: row.totalOutstanding,
      ...Object.fromEntries(row.bucketAmounts.map((amount, index) => [`bucket_${index}`, amount])),
      lastPaymentDate: row.lastPaymentDate,
      openInvoices: row.openInvoices.length,
      totalInvoices: row.totalInvoices,
      unappliedCredit: row.unappliedCredit
    });
  });
  finalizeWorksheet(summarySheet);

  const detailsSheet = workbook.addWorksheet("Invoice Details");
  detailsSheet.columns = [
    { header: "Party Name", key: "partyName", width: 28 },
    { header: "Invoice Date", key: "invoiceDate", width: 14 },
    { header: "Voucher No.", key: "voucherNumber", width: 16 },
    { header: "Reference", key: "reference", width: 18 },
    { header: "Original Amount", key: "originalAmount", width: 16 },
    { header: "Adjusted Amount", key: "adjustedAmount", width: 16 },
    { header: "Balance Amount", key: "balanceAmount", width: 16 },
    { header: "Age Days", key: "ageDays", width: 12 },
    { header: "Bucket", key: "bucketLabel", width: 16 },
    { header: "Last Payment", key: "lastPaymentDate", width: 16 }
  ];
  styleWorksheetHeader(detailsSheet);
  report.summaryRows.forEach((row) => {
    row.openInvoices.forEach((invoice) => {
      detailsSheet.addRow({
        partyName: row.partyName,
        invoiceDate: invoice.invoiceDate,
        voucherNumber: invoice.voucherNumber,
        reference: invoice.reference,
        originalAmount: invoice.originalAmount,
        adjustedAmount: invoice.adjustedAmount,
        balanceAmount: invoice.balanceAmount,
        ageDays: invoice.ageDays,
        bucketLabel: invoice.bucketLabel,
        lastPaymentDate: invoice.lastPaymentDate
      });
    });
  });
  finalizeWorksheet(detailsSheet);
  await downloadWorkbook(`creditor-ageing-${report.asOnDate}.xlsx`, workbook);
}

function downloadTallyAgeingCsv() {
  const report = getTallyAgeingReport();
  if (!report.summaryRows.length) {
    alert("No creditor ageing data found to export.");
    return;
  }
  const csv = [
    ["partyName", "totalPayable", ...report.bucketLabels, "lastPaymentDate", "openInvoices", "totalInvoices", "unappliedAdvance"].join(","),
    ...report.summaryRows.map((row) => [
      csvValue(row.partyName),
      row.totalOutstanding,
      ...row.bucketAmounts.map((amount) => amount),
      csvValue(row.lastPaymentDate),
      row.openInvoices.length,
      row.totalInvoices,
      row.unappliedCredit
    ].join(","))
  ].join("\n");
  downloadTextFile(`creditor-ageing-${report.asOnDate}.csv`, csv, "text/csv");
}

async function fetchTallyXml(endpoint, xmlBody) {
  const response = await fetch(resolveTallyRequestUrl(endpoint), {
    method: "POST",
    headers: { "Content-Type": "text/xml; charset=utf-8" },
    body: xmlBody
  });
  if (!response.ok) {
    const details = await readProxyErrorMessage(response);
    throw new Error(details || `Tally returned HTTP ${response.status}.`);
  }
  return response.text();
}

async function testTallyEndpoint(endpoint) {
  const response = await fetch(resolveTallyTestUrl(endpoint));
  let data = null;
  try {
    data = await response.json();
  } catch {
    data = null;
  }
  if (!response.ok) {
    throw new Error(data?.error || `Tally test failed with HTTP ${response.status}.`);
  }
  return data;
}

function getDefaultTallyEndpoint() {
  const origin = clean(window.location?.origin);
  if (origin && /^https?:\/\/(127\.0\.0\.1|localhost):3100$/i.test(origin)) {
    return `${origin}/tally`;
  }
  return "http://127.0.0.1:9000";
}

function resolveTallyRequestUrl(endpoint) {
  const target = clean(endpoint) || "http://127.0.0.1:9000";
  const origin = clean(window.location?.origin);
  if (origin && /^https?:\/\/(127\.0\.0\.1|localhost):3100$/i.test(origin)) {
    if (/\/tally$/i.test(target)) return target;
    return `${origin}/tally?target=${encodeURIComponent(target)}`;
  }
  return target;
}

function resolveTallyTestUrl(endpoint) {
  const target = clean(endpoint) || "http://127.0.0.1:9000";
  const origin = clean(window.location?.origin);
  if (origin && /^https?:\/\/(127\.0\.0\.1|localhost):3100$/i.test(origin)) {
    return `${origin}/tally-test?target=${encodeURIComponent(target)}`;
  }
  return target;
}

async function readProxyErrorMessage(response) {
  const contentType = response.headers.get("content-type") || "";
  try {
    if (contentType.includes("application/json")) {
      const data = await response.json();
      return data?.details || data?.error || "";
    }
    const text = await response.text();
    return clean(text);
  } catch {
    return "";
  }
}

function buildTallyLedgerRequest(config) {
  return `<?xml version="1.0" encoding="utf-8"?>
<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>Codex Creditors Ledgers</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        ${config.company ? `<SVCURRENTCOMPANY>${escapeXml(config.company)}</SVCURRENTCOMPANY>` : ""}
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="Codex Creditors Ledgers">
            <TYPE>Ledger</TYPE>
            <CHILDOF>Sundry Creditors</CHILDOF>
            <FETCH>Name, Parent, ClosingBalance</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>`;
}

function buildTallyVoucherRequest(config) {
  return `<?xml version="1.0" encoding="utf-8"?>
<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>Codex Creditors Vouchers</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        ${config.company ? `<SVCURRENTCOMPANY>${escapeXml(config.company)}</SVCURRENTCOMPANY>` : ""}
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        ${config.fromDate ? `<SVFROMDATE>${formatTallyDate(config.fromDate)}</SVFROMDATE>` : ""}
        <SVTODATE>${formatTallyDate(config.asOnDate || todayIso())}</SVTODATE>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="Codex Creditors Vouchers">
            <TYPE>Voucher</TYPE>
            <FETCH>Date, VoucherNumber, VoucherTypeName, PartyLedgerName, MasterID, AlterID, Narration, AllLedgerEntries.*, LedgerEntries.*</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>`;
}

function parseTallyLedgers(xmlText) {
  const doc = parseXmlDocument(xmlText);
  const ledgers = [...doc.getElementsByTagName("LEDGER")].map((ledger) => ({
    name: clean(ledger.getAttribute("NAME") || readXmlValue(ledger, "NAME")),
    parent: clean(readXmlValue(ledger, "PARENT")),
    closingBalance: parseAmount(readXmlValue(ledger, "CLOSINGBALANCE"))
  })).filter((ledger) => ledger.name);
  return dedupeByKey(ledgers, (ledger) => normalizeKey(ledger.name));
}

function parseTallyVouchers(xmlText) {
  const doc = parseXmlDocument(xmlText);
  return [...doc.getElementsByTagName("VOUCHER")].map((voucher) => ({
    id: clean(voucher.getAttribute("REMOTEID") || readXmlValue(voucher, "MASTERID") || readXmlValue(voucher, "ALTERID") || readXmlValue(voucher, "GUID")),
    masterId: clean(readXmlValue(voucher, "MASTERID")),
    alterId: clean(readXmlValue(voucher, "ALTERID")),
    date: tallyDateToIso(readXmlValue(voucher, "DATE")),
    voucherNumber: clean(readXmlValue(voucher, "VOUCHERNUMBER")),
    voucherType: clean(readXmlValue(voucher, "VOUCHERTYPENAME")),
    partyLedgerName: clean(readXmlValue(voucher, "PARTYLEDGERNAME")),
    narration: clean(readXmlValue(voucher, "NARRATION")),
    entries: dedupeByKey(
      readXmlLists(voucher, "ALLLEDGERENTRIES.LIST").concat(readXmlLists(voucher, "LEDGERENTRIES.LIST")).map((entry) => ({
        ledgerName: clean(readXmlValue(entry, "LEDGERNAME")),
        amount: parseAmount(readXmlValue(entry, "AMOUNT")),
        isParty: normalizeKey(readXmlValue(entry, "ISPARTYLEDGER")) === "yes",
        billAllocations: readXmlLists(entry, "BILLALLOCATIONS.LIST").map((allocation) => ({
          name: clean(allocation.getAttribute("NAME") || readXmlValue(allocation, "NAME")),
          billType: clean(readXmlValue(allocation, "BILLTYPE") || readXmlValue(allocation, "BILLTYPEOFREF")),
          amount: parseAmount(readXmlValue(allocation, "AMOUNT")),
          dueDate: tallyDateToIso(readXmlValue(allocation, "DUEDATE"))
        }))
      })).filter((entry) => entry.ledgerName),
      (entry) => `${normalizeKey(entry.ledgerName)}__${entry.amount}__${entry.billAllocations.map((allocation) => `${normalizeKey(allocation.name)}:${allocation.amount}`).join("|")}`
    )
  })).filter((voucher) => voucher.date && voucher.entries.length);
}

function parseXmlDocument(xmlText) {
  const doc = new DOMParser().parseFromString(xmlText, "application/xml");
  const error = doc.querySelector("parsererror");
  if (error) {
    throw new Error("Invalid XML received from Tally.");
  }
  return doc;
}

function readXmlLists(node, tagName) {
  return [...node.childNodes].filter((child) => child.nodeType === 1 && child.nodeName.toUpperCase() === tagName.toUpperCase());
}

function readXmlValue(node, tagName) {
  const child = [...node.childNodes].find((item) => item.nodeType === 1 && item.nodeName.toUpperCase() === tagName.toUpperCase());
  return clean(child?.textContent);
}

function updateTallyStatus(text) {
  if (els.tallySyncStatus) els.tallySyncStatus.textContent = text;
}

function collectStyleVariants() {
  return [...(els.styleVariantRows?.querySelectorAll(".style-variant-row") || [])]
    .map((row) => ({
      color: clean(row.querySelector('[name="styleVariantColor"]')?.value),
      orderQty: num(row.querySelector('[name="styleVariantOrderQty"]')?.value)
    }))
    .filter((variant) => variant.color);
}

function buildStylePayload(payload) {
  return {
    id: payload.id || uid(),
    styleNumber: payload.styleNumber,
    buyerName: payload.buyerName,
    styleName: payload.styleName,
    color: payload.color,
    orderQty: payload.orderQty,
    cmtRate: payload.cmtRate,
    serviceChargePct: payload.serviceChargePct,
    image: payload.image || "",
    notes: payload.notes,
    operations: payload.operations
  };
}

function styleImageSrc(style) {
  return resolveReportImageSource(style);
}

function resolveStyleImageValue(style) {
  const rawValue = clean(style?.image || "");
  if (!rawValue) return "";
  if (isStoredImageRef(rawValue)) {
    const styleId = storedImageKey(rawValue);
    return clean(state.styleImages?.[styleId] || "") || styleImageCache.get(styleId) || "";
  }
  return rawValue;
}

function resolveReportImageSource(source = {}) {
  const direct = resolveStyleImageValue(source);
  if (direct) return direct;

  const candidates = [];
  const styleId = clean(source?.styleId || source?.id);
  if (styleId) {
    const linked = byId(styleId);
    if (linked) candidates.push(linked);
  }

  const styleNumber = clean(source?.styleNumber);
  const color = clean(source?.color).toLowerCase();
  if (styleNumber) {
    const matches = state.styles.filter((item) => clean(item.styleNumber).toLowerCase() === styleNumber.toLowerCase());
    const exactColorMatch = matches.find((item) => clean(item.color).toLowerCase() === color);
    if (exactColorMatch) candidates.push(exactColorMatch);
    candidates.push(...matches);
  }

  for (const candidate of candidates) {
    const value = resolveStyleImageValue(candidate);
    if (value) return value;
  }
  return "";
}

function getStyleImageDiagnostics() {
  const summary = {
    total: state.styles.length,
    visible: 0,
    missing: 0,
    legacyRecoverable: 0,
    legacyMissing: 0
  };
  const details = [];
  state.styles.forEach((style) => {
    const rawValue = clean(style?.image || "");
    const resolved = clean(styleImageSrc(style));
    if (resolved) {
      summary.visible += 1;
      return;
    }
    if (!rawValue) {
      summary.missing += 1;
      details.push(`${styleLabel(style)} - no image saved`);
      return;
    }
    if (isStoredImageRef(rawValue)) {
      const styleId = storedImageKey(rawValue);
      const cachedImage = clean(state.styleImages?.[styleId] || "") || styleImageCache.get(styleId) || "";
      if (cachedImage) {
        summary.legacyRecoverable += 1;
        details.push(`${styleLabel(style)} - legacy image can be recovered in this browser`);
      } else {
        summary.legacyMissing += 1;
        details.push(`${styleLabel(style)} - legacy image reference exists but image data is missing`);
      }
      return;
    }
    summary.missing += 1;
    details.push(`${styleLabel(style)} - image path saved but not rendering`);
  });
  return { summary, details };
}

function updateStyleImageDiagnosticStatus() {
  if (!els.styleImageDiagnosticStatus) return;
  const { summary } = getStyleImageDiagnostics();
  els.styleImageDiagnosticStatus.textContent = `Visible: ${summary.visible} | Missing: ${summary.missing} | Legacy recoverable: ${summary.legacyRecoverable} | Legacy missing: ${summary.legacyMissing}`;
}

function runStyleImageDiagnostic() {
  const { summary, details } = getStyleImageDiagnostics();
  updateStyleImageDiagnosticStatus();
  const lines = [
    `Total styles: ${summary.total}`,
    `Visible images: ${summary.visible}`,
    `Missing images: ${summary.missing}`,
    `Legacy recoverable in this browser: ${summary.legacyRecoverable}`,
    `Legacy missing image data: ${summary.legacyMissing}`
  ];
  if (details.length) {
    lines.push("", "Details:", ...details.slice(0, 40));
    if (details.length > 40) lines.push(`...and ${details.length - 40} more`);
  }
  alert(lines.join("\n"));
}

function isStoredImageRef(value) {
  return clean(value).startsWith(STORED_IMAGE_PREFIX);
}

function storedImageRef(styleId) {
  return `${STORED_IMAGE_PREFIX}${styleId}`;
}

function storedImageKey(value) {
  return clean(value).slice(STORED_IMAGE_PREFIX.length);
}

function parseWashcareSymbols(value) {
  const raw = String(value || "");
  const parts = raw
    .split(/\r?\n|,|;|\||\u2022/g)
    .map((item) => clean(item))
    .filter(Boolean);
  return uniqueValues([...parts, ...extractCareSymbolPhrases(raw)]);
}

function formatWashcarePrintMethod(value) {
  if (value === "command") return "Copy Command";
  if (value === "both") return "Preview + Command";
  return "Browser Preview";
}

function configurePdfJs() {
  if (!window.pdfjsLib) return;
  if (window.pdfjsLib.GlobalWorkerOptions) {
    window.pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js";
  }
}

function getWashcareRecordByStyle(styleId) {
  return state.washcareRecords.find((record) => record.styleId === styleId) || null;
}

function populateWashcareForm(record) {
  if (!els.washcareForm) return;
  els.washcareForm.dataset.editId = record?.id || "";
  els.washcareForm.styleId.value = record?.styleId || "";
  els.washcareForm.reportNumber.value = record?.reportNumber || "";
  els.washcareForm.reportDate.value = record?.reportDate || "";
  els.washcareForm.labName.value = record?.labName || "";
  els.washcareForm.templateName.value = record?.templateName || "";
  els.washcareForm.templatePath.value = record?.templatePath || DEFAULT_WASHCARE_TEMPLATE_PATH;
  els.washcareForm.printMethod.value = record?.printMethod || "browser";
  els.washcareForm.printCommand.value = record?.printCommand || "";
  els.washcareForm.labelWidthMm.value = record?.labelWidthMm || DEFAULT_WASHCARE_LABEL_WIDTH_MM;
  els.washcareForm.labelLengthMm.value = record?.labelLengthMm || DEFAULT_WASHCARE_LABEL_LENGTH_MM;
  els.washcareForm.reportSourceName.value = record?.reportSourceName || "";
  els.washcareForm.reportStyleCode.value = record?.reportStyleCode || "";
  els.washcareForm.reportColor.value = record?.reportColor || "";
  els.washcareForm.reportBrand.value = record?.reportBrand || "";
  els.washcareForm.reportSupplier.value = record?.reportSupplier || "";
  els.washcareForm.composition.value = record?.composition || "";
  els.washcareForm.compositionFontSize.value = record?.compositionFontSize || 22;
  els.washcareForm.careSymbols.value = (record?.careSymbols || []).join(", ");
  els.washcareForm.symbolSize.value = record?.symbolSize || 30;
  els.washcareForm.symbolGap.value = record?.symbolGap || 6;
  els.washcareForm.washcareText.value = record?.washcareText || "";
  els.washcareForm.washcareFontSize.value = record?.washcareFontSize || 16;
  els.washcareForm.originLine.value = record?.originLine || "MADE IN INDIA";
  els.washcareForm.footerLine1.value = record?.footerLine1 || "";
  els.washcareForm.footerLine2.value = record?.footerLine2 || "";
  els.washcareForm.notes.value = record?.notes || "";
  renderWashcarePreview();
}

function populateWashcareFormByStyle(styleId) {
  if (!els.washcareForm) return;
  const record = getWashcareRecordByStyle(styleId);
  if (record) {
    populateWashcareForm(record);
    return;
  }
  const currentStyleId = clean(styleId);
  els.washcareForm.reset();
  delete els.washcareForm.dataset.editId;
  els.washcareForm.styleId.value = currentStyleId;
  if (els.washcareForm.printMethod) els.washcareForm.printMethod.value = "browser";
  if (els.washcareForm.templatePath) els.washcareForm.templatePath.value = DEFAULT_WASHCARE_TEMPLATE_PATH;
  if (els.washcareForm.labelWidthMm) els.washcareForm.labelWidthMm.value = DEFAULT_WASHCARE_LABEL_WIDTH_MM;
  if (els.washcareForm.labelLengthMm) els.washcareForm.labelLengthMm.value = DEFAULT_WASHCARE_LABEL_LENGTH_MM;
  if (els.washcareForm.compositionFontSize) els.washcareForm.compositionFontSize.value = 22;
  if (els.washcareForm.symbolSize) els.washcareForm.symbolSize.value = 30;
  if (els.washcareForm.symbolGap) els.washcareForm.symbolGap.value = 6;
  if (els.washcareForm.washcareFontSize) els.washcareForm.washcareFontSize.value = 16;
  if (els.washcareForm.originLine) els.washcareForm.originLine.value = "MADE IN INDIA";
  renderWashcarePreview();
}

function getWashcarePreviewData(record = null) {
  const source = record || (els.washcareForm ? new FormData(els.washcareForm) : null);
  const styleId = record ? record.styleId : clean(source?.get("styleId"));
  const style = byId(styleId);
  if (!style) return { style: null };
  const templatePath = record ? clean(record.templatePath) : clean(source.get("templatePath"));
  const manualPrintCommand = record ? clean(record.printCommand) : clean(source.get("printCommand"));
  return {
    style,
    reportNumber: record ? clean(record.reportNumber) : clean(source.get("reportNumber")),
    reportDate: record ? clean(record.reportDate) : normalizeDateValue(source.get("reportDate")),
    labName: record ? clean(record.labName) : clean(source.get("labName")),
    templateName: record ? clean(record.templateName) : clean(source.get("templateName")),
    templatePath,
    printMethod: record ? clean(record.printMethod) : clean(source.get("printMethod")) || "browser",
    printCommand: manualPrintCommand || generateWashcarePrintCommand({ templatePath }),
    labelWidthMm: record ? num(record.labelWidthMm) : clampWashcareDimension(source.get("labelWidthMm"), DEFAULT_WASHCARE_LABEL_WIDTH_MM),
    labelLengthMm: record ? num(record.labelLengthMm) : clampWashcareDimension(source.get("labelLengthMm"), DEFAULT_WASHCARE_LABEL_LENGTH_MM),
    reportSourceName: record ? clean(record.reportSourceName) : clean(source.get("reportSourceName")),
    reportStyleCode: record ? clean(record.reportStyleCode) : clean(source.get("reportStyleCode")),
    reportColor: record ? clean(record.reportColor) : clean(source.get("reportColor")),
    reportBrand: record ? clean(record.reportBrand) : clean(source.get("reportBrand")),
    reportSupplier: record ? clean(record.reportSupplier) : clean(source.get("reportSupplier")),
    composition: record ? clean(record.composition) : clean(source.get("composition")),
    compositionFontSize: record ? num(record.compositionFontSize) : clampWashcareSize(source.get("compositionFontSize"), 22),
    careSymbols: record ? parseWashcareSymbols((record.careSymbols || []).join(",")) : parseWashcareSymbols(source.get("careSymbols")),
    symbolSize: record ? num(record.symbolSize) : clampWashcareSize(source.get("symbolSize"), 30),
    symbolGap: record ? num(record.symbolGap) : clampWashcareSize(source.get("symbolGap"), 6),
    washcareText: record ? clean(record.washcareText) : clean(source.get("washcareText")),
    washcareFontSize: record ? num(record.washcareFontSize) : clampWashcareSize(source.get("washcareFontSize"), 16),
    originLine: record ? clean(record.originLine) : clean(source.get("originLine")) || "MADE IN INDIA",
    footerLine1: record ? clean(record.footerLine1) : clean(source.get("footerLine1")),
    footerLine2: record ? clean(record.footerLine2) : clean(source.get("footerLine2")),
    notes: record ? clean(record.notes) : clean(source.get("notes"))
  };
}

function buildWashcarePreviewMarkup(preview) {
  const careSymbolEntries = resolveWashcareSymbolEntries(preview.careSymbols, preview.washcareText);
  const washcareLines = formatWashcareLines(preview.washcareText);
  return `
    <div class="washcare-label" style="--washcare-composition-size:${num(preview.compositionFontSize) || 22}px; --washcare-symbol-size:${num(preview.symbolSize) || 30}px; --washcare-symbol-gap:${num(preview.symbolGap) || 6}px; --washcare-text-size:${num(preview.washcareFontSize) || 16}px; --washcare-label-width:${num(preview.labelWidthMm) || DEFAULT_WASHCARE_LABEL_WIDTH_MM}mm; --washcare-label-length:${num(preview.labelLengthMm) || DEFAULT_WASHCARE_LABEL_LENGTH_MM}mm;">
      <div class="washcare-label-composition">${esc((preview.composition || "-").toUpperCase())}</div>
      <div class="washcare-label-icons">${careSymbolEntries.length ? careSymbolEntries.map((entry) => renderWashcareIcon(entry.symbol, entry.phrase)).join("") : ""}</div>
      <div class="washcare-label-text">${washcareLines.length ? washcareLines.map((line) => `<div>${esc(line.toUpperCase())}</div>`).join("") : "<div>-</div>"}</div>
      <div class="washcare-label-footer">
        <div>${esc((preview.originLine || "MADE IN INDIA").toUpperCase())}</div>
        ${preview.footerLine1 ? `<div>${esc(preview.footerLine1.toUpperCase())}</div>` : ""}
        ${preview.footerLine2 ? `<div>${esc(preview.footerLine2.toUpperCase())}</div>` : ""}
      </div>
    </div>
  `;
}

function formatWashcareLines(text) {
  return String(text || "")
    .split(/\r?\n|,/)
    .map((line) => clean(line))
    .filter(Boolean);
}

function deriveCareSymbolsFromText(text) {
  return extractCareSymbolPhrases(text);
}

function clampWashcareSize(value, fallback) {
  const parsed = Number(value);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
}

function clampWashcareDimension(value, fallback) {
  const parsed = Number(value);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
}

function resolveWashcareSymbols(explicitSymbols = [], washcareText = "") {
  return resolveWashcareSymbolEntries(explicitSymbols, washcareText).map((entry) => entry.symbol);
}

function resolveWashcareSymbolEntries(explicitSymbols = [], washcareText = "") {
  const fromExplicit = explicitSymbols
    .map((phrase) => clean(phrase))
    .filter(Boolean)
    .map((phrase) => ({ phrase, symbol: resolveApprovedWashcareSymbol(phrase) }));
  const fromText = extractCareSymbolPhrases(washcareText)
    .map((phrase) => ({ phrase, symbol: resolveApprovedWashcareSymbol(phrase) }));
  return uniqueWashcareSymbolEntries([...fromExplicit, ...fromText]);
}

function resolveApprovedWashcareSymbol(phrase) {
  const normalizedPhrase = normalizeWashcareText(phrase);
  if (/(warm\s*40\s*[^a-z0-9]?\s*c|\b40\s*[^a-z0-9]?\s*c\b|\b40c\b)/i.test(normalizedPhrase)) return "wash-40";
  if (/(machine\s*wash|hand\s*wash|wash\s*(dark\s*colou?rs?\s*)?separately|normal\s*cycle)/i.test(normalizedPhrase)) return "wash";
  return mapPhraseToWashcareSymbol(phrase);
}

function uniqueWashcareSymbolEntries(entries) {
  const seen = new Set();
  return entries.filter((entry) => {
    const key = entry.symbol || normalizeKey(entry.phrase);
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function mapPhraseToWashcareSymbol(phrase) {
  const normalizedPhrase = normalizeWashcareText(phrase);
  const key = normalizeKey(normalizedPhrase);
  if (!key) return "";
  if (key.includes("donotironondecoration")) return "";
  if (/(machine\s*wash|hand\s*wash|wash\s*(dark\s*colou?rs?\s*)?separately|warm\s*40\s*[°º]?\s*c|\b40\s*[°º]?\s*c\b|\b40c\b|normal\s*cycle)/i.test(normalizedPhrase)) return "wash-40";
  if (/(do\s*not\s*bleach|no\s*bleach|\bbleach\b)/i.test(normalizedPhrase)) return "no-bleach";
  if (/(line\s*dry(\s*in)?\s*shade|dry\s*in\s*shade)/i.test(normalizedPhrase)) return "line-dry-shade";
  if (/tumble\s*dry/i.test(normalizedPhrase)) return "tumble-dry-low";
  if (/(iron\s*at\s*low|iron\s*low|low\s*setting|iron\s*.*\blow\b)/i.test(normalizedPhrase)) return "iron-low";
  return "";
}

function extractCareSymbolPhrases(text) {
  const source = normalizeWashcareText(text);
  if (!source) return [];
  const matches = [];
  const phraseMap = [
    ["Machine Wash", /(machine\s*wash|hand\s*wash|wash\s*(dark\s*colou?rs?\s*)?separately|warm\s*40\s*[°º]?\s*c|\b40\s*[°º]?\s*c\b|\b40c\b|normal\s*cycle)/i],
    ["Do Not Bleach", /(do\s*not\s*bleach|no\s*bleach|\bbleach\b)/i],
    ["Line Dry In Shade", /(line\s*dry(\s*in)?\s*shade|dry\s*in\s*shade)/i],
    ["Tumble Dry Low", /tumble\s*dry/i],
    ["Iron Low", /(iron\s*at\s*low|iron\s*low|low\s*setting|iron\s*.*\blow\b)/i]
  ];
  phraseMap.forEach(([label, pattern]) => {
    if (pattern.test(source)) matches.push(label);
  });
  return uniqueValues(matches);
}

function normalizeWashcareText(text) {
  return clean(text).replace(/\s+/g, " ");
}

function uniqueValues(values = []) {
  return [...new Set(values.map((value) => clean(value)).filter(Boolean))];
}

function renderWashcareIcon(symbol, phrase = "") {
  const label = phrase ? ` aria-label="${esc(phrase)}" title="${esc(phrase)}"` : ' aria-hidden="true"';
  if (symbol === "wash-40") {
    return `<svg class="washcare-icon" viewBox="0 0 64 48" fill="none" xmlns="http://www.w3.org/2000/svg" preserveAspectRatio="xMidYMid meet"${label}><path d="M10 12h44l-4 24H14z" fill="none" stroke="currentColor" stroke-width="2.3" stroke-linejoin="round"/><path d="M14 16c4 5 8 5 12 0 4 5 8 5 12 0 4 5 8 5 12 0" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/><text x="32" y="33" text-anchor="middle" font-size="12" font-family="Arial, Helvetica, sans-serif" font-weight="700" fill="currentColor">40C</text></svg>`;
  }
  if (symbol === "wash") {
    return `<svg class="washcare-icon" viewBox="0 0 64 48" fill="none" xmlns="http://www.w3.org/2000/svg" preserveAspectRatio="xMidYMid meet"${label}><path d="M10 12h44l-4 24H14z" fill="none" stroke="currentColor" stroke-width="2.3" stroke-linejoin="round"/><path d="M14 16c4 5 8 5 12 0 4 5 8 5 12 0 4 5 8 5 12 0" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg>`;
  }
  if (symbol === "no-bleach") {
    return `<svg class="washcare-icon" viewBox="0 0 64 48" fill="none" xmlns="http://www.w3.org/2000/svg" preserveAspectRatio="xMidYMid meet"${label}><path d="M32 9l18 30H14z" fill="none" stroke="currentColor" stroke-width="2.3" stroke-linejoin="round"/><path d="M19 12l26 24M45 12L19 36" fill="none" stroke="currentColor" stroke-width="2.3" stroke-linecap="round"/></svg>`;
  }
  if (symbol === "tumble-dry-low") {
    return `<svg class="washcare-icon" viewBox="0 0 64 48" fill="none" xmlns="http://www.w3.org/2000/svg" preserveAspectRatio="xMidYMid meet"${label}><rect x="12" y="8" width="40" height="32" fill="none" stroke="currentColor" stroke-width="2.3"/><circle cx="32" cy="24" r="10" fill="none" stroke="currentColor" stroke-width="2.1"/><circle cx="32" cy="24" r="2.4" fill="currentColor"/></svg>`;
  }
  if (symbol === "iron-low") {
    return `<svg class="washcare-icon" viewBox="0 0 64 48" fill="none" xmlns="http://www.w3.org/2000/svg" preserveAspectRatio="xMidYMid meet"${label}><path d="M14 29.5h36c.9 0 1.6-.7 1.6-1.6 0-6.8-5.4-12.5-12.1-13.2V11H22v5.5h7.2c2.9 0 5.3 1.9 6 4.4l1.8 8.6H14z" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linejoin="round"/><path d="M18 34h28" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"/><circle cx="40" cy="19" r="1.9" fill="currentColor"/></svg>`;
  }
  if (symbol === "line-dry-shade") {
    return `<svg class="washcare-icon" viewBox="0 0 64 48" fill="none" xmlns="http://www.w3.org/2000/svg" preserveAspectRatio="xMidYMid meet"${label}><path d="M12 10h40v28H12z" fill="none" stroke="currentColor" stroke-width="2.2"/><path d="M32 14v19" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"/><path d="M15 12h14L15 26z" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"/></svg>`;
  }
  return "";
}

function resetStyleVariantRows(variants = [{ color: "", orderQty: "" }]) {
  if (!els.styleVariantRows) return;
  els.styleVariantRows.innerHTML = "";
  variants.forEach((variant) => addStyleVariantRow(variant.color, variant.orderQty));
}

function resetStyleFormState() {
  els.styleForm?.reset();
  resetStyleVariantRows();
  clearPastedStyleImage();
  updateStyleFormImagePreview();
  els.operationRateRows.innerHTML = "";
  seedOperationRows();
  setToday();
}

function renderCuttingReportHeader() {
  if (!els.cuttingReportHead) return;
  els.cuttingReportHead.innerHTML = `<tr><th>Date</th><th>Photo</th><th>Style</th><th>Colour</th><th>Service</th>${getSizes().map((size) => `<th>${esc(size)}</th>`).join("")}<th>Total Qty</th><th>Remarks</th></tr>`;
}

function renderCuttingEntriesHeader() {
  if (!els.cuttingEntriesHead) return;
  els.cuttingEntriesHead.innerHTML = `<tr><th>Date</th><th>Style</th><th>Service</th><th>Order Qty</th>${getSizes().map((size) => `<th>${esc(size)}</th>`).join("")}<th>Total Qty</th><th>Remarks</th><th>Action</th></tr>`;
}

function renderPendingWorkflow() {
  if (!els.pendingWorkflowTable) return;
  const rows = state.styles.map((style) => {
    const cutQty = state.cuttingEntries.filter((entry) => entry.styleId === style.id).reduce((sum, entry) => sum + sumObj(entry.quantities), 0);
    const makeQty = state.styleProductionEntries.filter((entry) => entry.styleId === style.id).reduce((sum, entry) => sum + entryProducedQty(entry), 0);
    const dispatchQty = state.dispatchEntries.filter((entry) => entry.styleId === style.id).reduce((sum, entry) => sum + sumObj(entry.quantities), 0);
    let status = "";
    if (!cutQty) status = "Pending Cutting";
    else if (!makeQty) status = "Pending Make";
    else if (!dispatchQty) status = "Pending Dispatch";
    return { styleNumber: style.styleNumber, color: style.color, orderQty: num(style.orderQty), cutQty, makeQty, dispatchQty, status };
  }).filter((row) => row.status).sort((a, b) => a.styleNumber.localeCompare(b.styleNumber));
  els.pendingWorkflowTable.innerHTML = rowsOrEmpty(rows.map((row) => `
    <tr><td>${esc(row.styleNumber)}</td><td>${esc(row.color || "-")}</td><td><span class="status-chip pending">${esc(row.status)}</span></td><td>${fmtInt(row.orderQty)}</td><td>${fmtInt(row.cutQty)}</td><td>${fmtInt(row.makeQty)}</td><td>${fmtInt(row.dispatchQty)}</td></tr>`), 7, "No pending workflow items.");
}

function styleBillingRows(reportDate = "") {
  const reportRange = normalizeReportRangeInput(reportDate);
  const rangePayment = getRangePaymentOverview(reportRange);
  return state.styles.map((style) => {
    const cutQty = state.cuttingEntries
      .filter((e) => e.styleId === style.id && matchesDate(e.date, reportRange))
      .reduce((s, e) => s + sumObj(e.quantities), 0);
    const entries = state.styleProductionEntries.filter((e) => e.styleId === style.id && matchesDate(e.date, reportRange));
    const producedQty = entries.reduce((s, e) => s + entryProducedQty(e), 0);
    const dispatchQty = state.dispatchEntries
      .filter((e) => e.styleId === style.id && matchesDate(e.date, reportRange))
      .reduce((s, e) => s + sumObj(e.quantities), 0);
    const baseAmount = producedQty * num(style.cmtRate);
    const serviceChargePct = num(style.serviceChargePct);
    const serviceChargeAmount = baseAmount * serviceChargePct / 100;
    const billing = baseAmount + serviceChargeAmount;
    const paymentStatus = getStylePaymentStatus(style.id, rangePayment, billing);
    return {
      styleId: style.id,
      styleNumber: style.styleNumber,
      color: style.color,
      image: styleImageSrc(style),
      orderQty: num(style.orderQty),
      cutQty,
      producedQty,
      dispatchQty,
      acceptedQty: producedQty,
      cmtRate: num(style.cmtRate),
      baseAmount,
      serviceChargePct,
      serviceChargeAmount,
      billing,
      cutVsMakeSummary: buildSynopsis(cutQty, producedQty),
      makeVsDispatchSummary: buildSynopsis(producedQty, dispatchQty),
      dateLabel: getReportRangeLabel(reportRange),
      paymentStatusLabel: paymentStatus.label,
      paymentStatusClass: paymentStatus.className
    };
  }).filter((row) => row.cutQty > 0 || row.producedQty > 0 || row.dispatchQty > 0 || (!reportRange.startDate && !reportRange.endDate));
}

function workerBillingRows() {
  const map = new Map();
  state.productionEntries.forEach((e) => {
    const key = `${e.workerName}__${e.operationName}`;
    const row = map.get(key) || { workerName: e.workerName, operationName: e.operationName, quantity: 0, amount: 0 };
    row.quantity += num(e.quantity);
    row.amount += num(e.quantity) * num(e.operationRate);
    map.set(key, row);
  });
  return [...map.values()].sort((a, b) => a.workerName.localeCompare(b.workerName));
}

function reconciliationRows(reportDate = "") {
  const reportRange = normalizeReportRangeInput(reportDate);
  return state.styles.map((style) => {
    const cutQty = state.cuttingEntries.filter((e) => e.styleId === style.id && matchesDate(e.date, reportRange)).reduce((s, e) => s + sumObj(e.quantities), 0);
    const producedQty = state.styleProductionEntries.filter((e) => e.styleId === style.id && matchesDate(e.date, reportRange)).reduce((s, e) => s + entryProducedQty(e), 0);
    const acceptedQty = state.acceptanceEntries.filter((e) => e.styleId === style.id && matchesDate(e.date, reportRange)).reduce((s, e) => s + e.items.reduce((a, i) => a + i.accepted, 0), 0);
    const rejectedQty = state.acceptanceEntries.filter((e) => e.styleId === style.id && matchesDate(e.date, reportRange)).reduce((s, e) => s + e.items.reduce((a, i) => a + i.rejected, 0), 0);
    return { styleNumber: style.styleNumber, color: style.color, cutQty, producedQty, acceptedQty, rejectedQty, balance: cutQty - acceptedQty };
  }).filter((row) => row.cutQty || row.producedQty || row.acceptedQty || row.rejectedQty || (!reportRange.startDate && !reportRange.endDate));
}

function operationCostRows(reportDate = "") {
  const reportRange = normalizeReportRangeInput(reportDate);
  const map = new Map();
  state.productionEntries.filter((e) => matchesDate(e.date, reportRange)).forEach((e) => {
    const key = `${e.styleId}__${e.operationName}`;
    const row = map.get(key) || { styleNumber: byId(e.styleId)?.styleNumber || "-", operationName: e.operationName, quantity: 0, rate: num(e.operationRate), amount: 0 };
    row.quantity += num(e.quantity);
    row.amount += num(e.quantity) * num(e.operationRate);
    map.set(key, row);
  });
  return [...map.values()].sort((a, b) => a.styleNumber.localeCompare(b.styleNumber));
}

function cuttingReportRows(reportDate = "") {
  const reportRange = normalizeReportRangeInput(reportDate);
  return state.cuttingEntries
    .filter((entry) => matchesDate(entry.date, reportRange))
    .slice()
    .sort((a, b) => clean(b.date).localeCompare(clean(a.date)))
    .map((entry) => {
      const style = byId(entry.styleId);
        return {
          date: entry.date,
          image: styleImageSrc(style),
        styleNumber: style?.styleNumber || "-",
        color: style?.color || "",
        service: entry.service || "",
        quantities: getSizeQuantities(entry.quantities),
        totalQty: sumObj(entry.quantities),
        sizeWise: formatQuantities(entry.quantities),
        remarks: entry.remarks || ""
      };
    });
}

function dispatchReportRows(reportDate = "") {
  const reportRange = normalizeReportRangeInput(reportDate);
  const rows = [];
  state.styles.forEach((style) => {
    getSizes().forEach((size) => {
      const cutQty = state.cuttingEntries
        .filter((entry) => entry.styleId === style.id && matchesDate(entry.date, reportRange))
        .reduce((sum, entry) => sum + num(entry.quantities?.[size]), 0);
      const makeQty = state.styleProductionEntries
        .filter((entry) => entry.styleId === style.id && matchesDate(entry.date, reportRange))
        .reduce((sum, entry) => sum + num(entry.quantities?.[size]), 0);
      const dispatchQty = state.dispatchEntries
        .filter((entry) => entry.styleId === style.id && matchesDate(entry.date, reportRange))
        .reduce((sum, entry) => sum + num(entry.quantities?.[size]), 0);
      if (!cutQty && !makeQty && !dispatchQty && (reportRange.startDate || reportRange.endDate)) return;
      if (!cutQty && !makeQty && !dispatchQty) return;
      const amount = (makeQty * num(style.cmtRate)) + ((makeQty * num(style.cmtRate) * num(style.serviceChargePct)) / 100);
      rows.push({
        styleNumber: style.styleNumber,
        color: style.color,
        image: styleImageSrc(style),
        size,
        cutQty,
        makeQty,
        dispatchQty,
        amount,
        balance: cutQty - dispatchQty
      });
    });
  });
  return rows.sort((a, b) => `${a.styleNumber}${a.color}${a.size}`.localeCompare(`${b.styleNumber}${b.color}${b.size}`));
}

function dashboardSummary() {
  const totalOrderQty = state.styles.reduce((sum, style) => sum + num(style.orderQty), 0);
  const totalCutQty = state.cuttingEntries.reduce((sum, entry) => sum + sumObj(entry.quantities), 0);
  const totalMakeQty = state.styleProductionEntries.reduce((sum, entry) => sum + entryProducedQty(entry), 0);
  const totalDispatchQty = state.dispatchEntries.reduce((sum, entry) => sum + sumObj(entry.quantities), 0);
  const totalBilled = styleBillingRows().reduce((sum, row) => sum + row.billing, 0);
  const totalPaid = state.payments.reduce((sum, payment) => sum + num(payment.totalPaid), 0);
  const totalServiceChargePaid = state.payments.reduce((sum, payment) => sum + num(payment.serviceChargePaid), 0);
  return {
    totalOrderQty,
    totalCutQty,
    totalMakeQty,
    totalDispatchQty,
    totalBilled,
    totalPaid,
    totalServiceChargePaid,
    shortCutPct: totalOrderQty ? (Math.max(totalOrderQty - totalCutQty, 0) / totalOrderQty) * 100 : 0,
    shipAgainstMakePct: totalMakeQty ? (totalDispatchQty / totalMakeQty) * 100 : 0
  };
}

function renderPaymentHistory() {
  if (!els.paymentHistoryTable) return;
  const rows = state.payments
    .slice()
    .sort((a, b) => `${clean(b.paymentDate)}${clean(b.endDate)}`.localeCompare(`${clean(a.paymentDate)}${clean(a.endDate)}`))
    .map((payment) => {
      const styleLabels = payment.styleIds?.map((styleId) => byId(styleId)?.styleNumber).filter(Boolean).join(", ") || "-";
      return `
        <tr>
          <td>${esc(formatDateDisplay(payment.paymentDate))}</td>
          <td>${esc(getReportRangeLabel(payment))}</td>
          <td>${esc(styleLabels)}</td>
          <td>Rs ${fmt(payment.baseAmountPaid)}</td>
          <td>Rs ${fmt(payment.serviceChargePaid)}</td>
          <td>Rs ${fmt(payment.totalPaid)}</td>
          <td>${esc(payment.notes || "-")}</td>
          <td><button type="button" class="ghost small" data-action="delete-payment" data-entry-id="${payment.id}">Delete</button></td>
        </tr>`;
    });
  els.paymentHistoryTable.innerHTML = rowsOrEmpty(rows, 8, "No payment entries recorded.");
}

function populatePaymentFormForCurrentRange(styleAmounts, reportRange) {
  if (!els.paymentForm) return;
  const payment = paymentRecordForRange(reportRange);
  const totals = styleAmounts.reduce((sum, row) => {
    sum.baseAmount += row.baseAmount;
    sum.serviceChargeAmount += row.serviceChargeAmount;
    sum.totalBill += row.billing;
    return sum;
  }, { baseAmount: 0, serviceChargeAmount: 0, totalBill: 0 });

  const paymentDateField = els.paymentForm.elements.paymentDate;
  const notesField = els.paymentForm.elements.notes;
  const baseAmountField = els.paymentForm.elements.baseAmountPaid;
  const serviceChargeField = els.paymentForm.elements.serviceChargePaid;
  const totalPaidField = els.paymentForm.elements.totalPaid;

  if (paymentDateField && !clean(paymentDateField.value)) {
    paymentDateField.value = new Date().toISOString().slice(0, 10);
  }

  if (payment) {
    if (paymentDateField) paymentDateField.value = payment.paymentDate || paymentDateField.value;
    if (notesField) notesField.value = payment.notes || "";
    if (baseAmountField) baseAmountField.value = payment.baseAmountPaid || 0;
    if (serviceChargeField) serviceChargeField.value = payment.serviceChargePaid || 0;
    if (totalPaidField) totalPaidField.value = payment.totalPaid || 0;
    return;
  }

  if (notesField) notesField.value = "";
  if (baseAmountField) baseAmountField.value = totals.baseAmount ? totals.baseAmount.toFixed(2) : "0.00";
  if (serviceChargeField) serviceChargeField.value = totals.serviceChargeAmount ? totals.serviceChargeAmount.toFixed(2) : "0.00";
  if (totalPaidField) totalPaidField.value = totals.totalBill ? totals.totalBill.toFixed(2) : "0.00";
}

function syncPaymentFormTotal() {
  if (!els.paymentForm) return;
  const target = document.activeElement;
  const totalField = els.paymentForm.elements.totalPaid;
  if (!totalField || target === totalField) return;
  const baseAmount = num(els.paymentForm.elements.baseAmountPaid?.value);
  const serviceCharge = num(els.paymentForm.elements.serviceChargePaid?.value);
  totalField.value = (baseAmount + serviceCharge).toFixed(2);
}

function normalizeReportRangeInput(value = {}) {
  if (typeof value === "string") {
    const dateValue = normalizeDateValue(value);
    return { startDate: dateValue, endDate: dateValue };
  }
  const startDate = normalizeDateValue(value.startDate);
  const endDate = normalizeDateValue(value.endDate);
  if (startDate && endDate && startDate > endDate) {
    return { startDate: endDate, endDate: startDate };
  }
  return { startDate, endDate };
}

function getReportRangeLabel(range = {}) {
  const normalized = normalizeReportRangeInput(range);
  if (normalized.startDate && normalized.endDate) {
    return `${formatDateDisplay(normalized.startDate)} to ${formatDateDisplay(normalized.endDate)}`;
  }
  if (normalized.startDate) return `From ${formatDateDisplay(normalized.startDate)}`;
  if (normalized.endDate) return `Up to ${formatDateDisplay(normalized.endDate)}`;
  return "All Dates";
}

function formatReportRangeForFilename(range = {}) {
  const normalized = normalizeReportRangeInput(range);
  if (normalized.startDate && normalized.endDate) {
    return `${normalized.startDate}_to_${normalized.endDate}`;
  }
  if (normalized.startDate) return `from_${normalized.startDate}`;
  if (normalized.endDate) return `upto_${normalized.endDate}`;
  return "all-dates";
}

function formatDateDisplay(value) {
  const text = clean(value);
  if (!text) return "-";
  const [year, month, day] = text.split("-");
  if (!year || !month || !day) return text;
  return `${day}/${month}/${year}`;
}

function buildSynopsis(leftQty, rightQty) {
  const difference = num(leftQty) - num(rightQty);
  const label = difference === 0 ? "Balanced" : difference > 0 ? "Short" : "Excess";
  return `${fmtInt(leftQty)} / ${fmtInt(rightQty)} (${label} ${fmtInt(Math.abs(difference))})`;
}

function formatBillingPdfCell(key, row) {
  switch (key) {
    case "orderQty":
    case "cutQty":
    case "producedQty":
    case "dispatchQty":
      return fmtInt(row[key]);
    case "cmtRate":
    case "baseAmount":
    case "serviceChargeAmount":
    case "billing":
      return `Rs ${fmt(row[key])}`;
    default:
      return String(row[key] ?? "-");
  }
}

function paymentRecordForRange(range = {}) {
  const normalized = normalizeReportRangeInput(range);
  if (!normalized.startDate || !normalized.endDate) return null;
  return state.payments.find((payment) => isSameReportRange(payment, normalized)) || null;
}

function rangeStyleBillingAllocationRows(range = {}) {
  const normalized = normalizeReportRangeInput(range);
  return state.styles.map((style) => {
    const producedQty = state.styleProductionEntries
      .filter((entry) => entry.styleId === style.id && matchesDate(entry.date, normalized))
      .reduce((sum, entry) => sum + entryProducedQty(entry), 0);
    const baseAmount = producedQty * num(style.cmtRate);
    const serviceChargeAmount = baseAmount * num(style.serviceChargePct) / 100;
    return {
      styleId: style.id,
      producedQty,
      billing: baseAmount + serviceChargeAmount
    };
  }).filter((row) => row.billing > 0);
}

function getRangePaymentOverview(range = {}) {
  const payment = paymentRecordForRange(range);
  const normalized = normalizeReportRangeInput(range);
  const billRows = rangeStyleBillingAllocationRows(normalized);
  const totalBill = billRows.reduce((sum, row) => sum + row.billing, 0);
  const allocationByStyleId = {};
  let remainingPayment = num(payment?.totalPaid);
  billRows.forEach((row) => {
    let status = "Pending";
    if (remainingPayment >= row.billing) {
      status = "Paid";
      remainingPayment -= row.billing;
    } else if (remainingPayment > 0) {
      status = "Part Paid";
      remainingPayment = 0;
    }
    allocationByStyleId[row.styleId] = status;
  });
  return {
    payment,
    totalBill,
    allocationByStyleId,
    isFullyPaid: Boolean(payment) && num(payment.totalPaid) >= totalBill && totalBill > 0,
    isPartPaid: Boolean(payment) && num(payment.totalPaid) > 0 && num(payment.totalPaid) < totalBill
  };
}

function isSameReportRange(payment = {}, range = {}) {
  const normalized = normalizeReportRangeInput(range);
  return clean(payment.startDate) === normalized.startDate && clean(payment.endDate) === normalized.endDate;
}

function getStylePaymentStatus(styleId, paymentOverview, styleBilling = 0) {
  if (num(styleBilling) <= 0) {
    return { label: "No Bill", className: "pending" };
  }
  if (!paymentOverview?.payment) {
    return { label: "Pending", className: "pending" };
  }
  const status = paymentOverview.allocationByStyleId?.[styleId] || "Pending";
  if (status === "Paid") return { label: "Paid", className: "paid" };
  if (status === "Part Paid") return { label: "Part Paid", className: "pending" };
  return { label: "Pending", className: "pending" };
}

function exportData() {
  const blob = new Blob([JSON.stringify(state, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `piece-rate-data-${new Date().toISOString().slice(0, 10)}.json`;
  a.click();
  URL.revokeObjectURL(url);
}

function importData(e) {
  const file = e.target.files?.[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = async () => {
    try {
      state = { ...DEFAULTS, ...JSON.parse(String(reader.result)) };
      normalizeState();
      await persistState();
      alert("Data imported successfully.");
    } catch {
      alert("Invalid JSON file.");
    }
  };
  reader.readAsText(file);
  els.importInput.value = "";
}

async function importStylesCsv(e) {
  const file = e.target.files?.[0];
  if (!file) return;
  pendingStyleImportFile = file;
  await importPendingStylesCsv();
}

async function importStyleImageSheet(e) {
  const file = e.target.files?.[0];
  if (!file) return;
  try {
    const rows = await parseStyleImportFile(file);
    if (styleImportNeedsLocalImages(rows) && !els.styleImageImportInput?.files?.length) {
      alert("This image recovery file uses image file names. Please select those image files in 'Select Style Images' and upload again.");
      return;
    }
    const importedImageMap = await buildImportedImageMap(els.styleImageImportInput?.files);
    const summary = await importStyleRows(rows, importedImageMap, {
      updateImagesOnly: true
    });
    await persistState();
    alert(`Style image recovery finished.\nUpdated: ${summary.updated}\nSkipped: ${summary.skipped}`);
  } catch (error) {
    console.error("Style image recovery failed:", error);
    alert(`Could not import the style image file.${clean(error?.message) ? `\n${clean(error.message)}` : ""}`);
  } finally {
    if (els.styleImageRecoveryInput) els.styleImageRecoveryInput.value = "";
    if (els.styleImageImportInput) els.styleImageImportInput.value = "";
  }
}

async function importStylesCsvFromPendingSelection() {
  if (!pendingStyleImportFile && !els.styleImportInput?.files?.[0]) return;
  pendingStyleImportFile = pendingStyleImportFile || els.styleImportInput.files[0];
  await importPendingStylesCsv();
}

async function importPendingStylesCsv() {
  const file = pendingStyleImportFile;
  if (!file) return;
  try {
    const rows = await parseStyleImportFile(file);
    if (styleImportNeedsLocalImages(rows) && !els.styleImageImportInput?.files?.length) {
      alert("This style import file uses image file names. Please select those image files in 'Select Style Images' and the import will continue.");
      return;
    }
    const importedImageMap = await buildImportedImageMap(els.styleImageImportInput?.files);
    const summary = await importStyleRows(rows, importedImageMap, {
      skipExisting: Boolean(els.skipExistingStylesToggle?.checked)
    });
    await persistState();
    alert(`Styles import finished.\nCreated: ${summary.created}\nUpdated: ${summary.updated}\nSkipped: ${summary.skipped}`);
  } catch (error) {
    console.error("Style import failed:", error);
    alert(`Could not import the style file.${clean(error?.message) ? `\n${clean(error.message)}` : ""}`);
  }
  pendingStyleImportFile = null;
  els.styleImportInput.value = "";
  if (els.styleImageImportInput) els.styleImageImportInput.value = "";
}

async function importCmtUpdateSheet(e) {
  const file = e.target.files?.[0];
  if (!file) return;
  try {
    const rows = await parseWorkbookOrCsvRows(file);
    const summary = await applyCmtUpdateRows(rows);
    await persistState();
    alert(`CMT update finished.\nUpdated: ${summary.updated}\nSkipped: ${summary.skipped}`);
  } catch (error) {
    console.error("CMT update import failed:", error);
    alert(`Could not import the CMT update sheet.${clean(error?.message) ? `\n${clean(error.message)}` : ""}`);
  }
  if (els.cmtUpdateImportInput) els.cmtUpdateImportInput.value = "";
}

async function downloadPendingCmtSheet() {
  const rows = pendingCmtStyles();
  if (!rows.length) {
    alert("No pending CMT styles found.");
    return;
  }
  const workbook = createWorkbookOrAlert();
  if (!workbook) return;
  const sheet = workbook.addWorksheet("Pending CMT");
  sheet.columns = [
    { header: "styleNumber", key: "styleNumber", width: 20 },
    { header: "styleName", key: "styleName", width: 28 },
    { header: "color", key: "color", width: 18 },
    { header: "buyerName", key: "buyerName", width: 18 },
    { header: "cmtRate", key: "cmtRate", width: 14 },
    { header: "serviceChargePct", key: "serviceChargePct", width: 16 },
    { header: "singerRate", key: "singerRate", width: 14 },
    { header: "overlockRate", key: "overlockRate", width: 16 },
    { header: "flatRate", key: "flatRate", width: 12 },
    { header: "pendingFields", key: "pendingFields", width: 24 }
  ];
  styleWorksheetHeader(sheet);
  rows.forEach((style) => {
    sheet.addRow({
      styleNumber: style.styleNumber,
      styleName: style.styleName || "",
      color: style.color || "",
      buyerName: style.buyerName || "",
      cmtRate: num(style.cmtRate) || "",
      serviceChargePct: num(style.serviceChargePct) || "",
      singerRate: operationRateByName(style, "Singer") || "",
      overlockRate: operationRateByName(style, "Overlock") || "",
      flatRate: operationRateByName(style, "Flat") || "",
      pendingFields: pendingCmtReason(style)
    });
  });
  finalizeWorksheet(sheet);
  await downloadWorkbook(`pending-cmt-styles-${todayIso()}.xlsx`, workbook);
}

async function applyCmtUpdateRows(rows) {
  const summary = { updated: 0, skipped: 0 };
  for (const row of rows) {
    const styleNumber = clean(row.styleNumber);
    if (!styleNumber) continue;
    const color = clean(row.color);
    const matchingStyles = state.styles.filter((style) =>
      style.styleNumber.toLowerCase() === styleNumber.toLowerCase()
      && (!color || clean(style.color).toLowerCase() === color.toLowerCase())
    );
    const style = matchingStyles.length === 1
      ? matchingStyles[0]
      : matchingStyles.find((item) => clean(item.color).toLowerCase() === color.toLowerCase());
    if (!style) {
      summary.skipped += 1;
      continue;
    }
    if (clean(row.buyerName)) style.buyerName = clean(row.buyerName);
    if (clean(row.styleName)) style.styleName = clean(row.styleName);
    if (clean(row.cmtRate)) style.cmtRate = num(row.cmtRate);
    if (clean(row.serviceChargePct)) style.serviceChargePct = num(row.serviceChargePct);
    if (clean(row.singerRate)) upsertStyleOperation(style, "Singer", row.singerRate);
    if (clean(row.overlockRate)) upsertStyleOperation(style, "Overlock", row.overlockRate);
    if (clean(row.flatRate)) upsertStyleOperation(style, "Flat", row.flatRate);
    summary.updated += 1;
  }
  return summary;
}

function styleImportNeedsLocalImages(rows) {
  return rows.some((row) => {
    const imageValue = clean(row.image);
    return imageValue && !isDirectImageSource(imageValue);
  });
}

async function importCuttingCsv(e) {
  const file = e.target.files?.[0];
  if (!file) return;
  try {
    const rows = parseCsv(await file.text());
    const summary = { imported: 0, skipped: 0, zeroQty: 0 };
    rows.forEach((row) => {
      const style = findStyleByNumber(row.styleNumber, row.color);
      if (!style) {
        summary.skipped += 1;
        return;
      }
      const quantities = sizeQuantitiesFromRow(row);
      if (sumObj(quantities) <= 0) summary.zeroQty += 1;
      state.cuttingEntries.push({
        id: uid(),
        date: normalizeDateValue(row.date),
        styleId: style.id,
        service: clean(row.service),
        remarks: clean(row.remarks),
        quantities
      });
      summary.imported += 1;
    });
    await persistState();
    alert(`Cutting data import finished.\nImported: ${summary.imported}\nSkipped: ${summary.skipped}\nRows with 0 qty: ${summary.zeroQty}`);
  } catch {
    alert("Could not import cutting CSV.");
  }
  els.cuttingImportInput.value = "";
}
async function importProductionCsv(e) {
  const file = e.target.files?.[0];
  if (!file) return;
  try {
    const rows = parseCsv(await file.text());
    rows.forEach((row) => {
      const style = findStyleByNumber(row.styleNumber, row.color);
      if (!style) return;
      const operationName = clean(row.operationName);
      const op = style.operations.find((x) => x.operationName.toLowerCase() === operationName.toLowerCase());
      state.productionEntries.push({
        id: uid(),
        date: normalizeDateValue(row.date),
        styleId: style.id,
        operationName,
        operationRate: op ? num(op.rate) : num(row.operationRate),
        workerName: clean(row.workerName),
        size: clean(row.size),
        quantity: num(row.quantity),
        remarks: clean(row.remarks)
      });
    });
    await persistState();
    alert("Production data imported successfully.");
  } catch {
    alert("Could not import production CSV.");
  }
  els.productionImportInput.value = "";
}

async function importStyleProductionCsv(e) {
  const file = e.target.files?.[0];
  if (!file) return;
  try {
    const rows = parseCsv(await file.text());
    const summary = { imported: 0, skipped: 0, zeroQty: 0 };
    rows.forEach((row) => {
      const style = findStyleByNumber(row.styleNumber, row.color);
      if (!style) {
        summary.skipped += 1;
        return;
      }
      const quantities = sizeQuantitiesFromRow(row);
      if (sumObj(quantities) <= 0 && num(row.totalQty) <= 0) summary.zeroQty += 1;
      state.styleProductionEntries.push({
        id: uid(),
        date: normalizeDateValue(row.date),
        styleId: style.id,
        totalQty: num(row.totalQty),
        remarks: clean(row.remarks),
        quantities
      });
      summary.imported += 1;
    });
    await persistState();
    alert(`Style production data import finished.\nImported: ${summary.imported}\nSkipped: ${summary.skipped}\nRows with 0 qty: ${summary.zeroQty}`);
  } catch {
    alert("Could not import style production CSV.");
  }
  if (els.styleProductionImportInput) els.styleProductionImportInput.value = "";
}

function parseCsv(text) {
  const lines = text.replace(/^\uFEFF/, "").split(/\r?\n/).filter((line) => line.trim());
  if (!lines.length) return [];
  const headers = splitCsvLine(lines[0]).map((h) => clean(h));
  return lines.slice(1).map((line) => {
    const values = splitCsvLine(line);
    return Object.fromEntries(headers.map((header, index) => [header, values[index] ?? ""]));
  });
}

const STYLE_IMPORT_FIELD_ALIASES = {
  stylenumber: "styleNumber",
  stylecode: "styleNumber",
  styleno: "styleNumber",
  styleid: "styleNumber",
  vendorstylecode: "styleNumber",
  uvpstyleid: "styleNumber",
  color: "color",
  colour: "color",
  colorname: "color",
  colourname: "color",
  vendorcolorname: "color",
  image: "image",
  imagename: "image",
  imageurl: "image",
  imagepath: "image",
  photourl: "image",
  photo: "image",
  buyer: "buyerName",
  buyername: "buyerName",
  stylename: "styleName",
  styledescription: "styleName",
  description: "styleName",
  orderqty: "orderQty",
  orderquantity: "orderQty",
  qty: "orderQty",
  vendorcapacity: "orderQty",
  cmtrate: "cmtRate",
  totalcmt: "cmtRate",
  servicechargepct: "serviceChargePct",
  servicechargepercent: "serviceChargePct",
  servicechargepercentage: "serviceChargePct",
  notes: "notes",
  remarks: "notes"
};

function splitCsvLine(line) {
  const result = [];
  let current = "";
  let inQuotes = false;
  for (let i = 0; i < line.length; i += 1) {
    const char = line[i];
    const next = line[i + 1];
    if (char === '"' && inQuotes && next === '"') {
      current += '"';
      i += 1;
    } else if (char === '"') {
      inQuotes = !inQuotes;
    } else if (char === "," && !inQuotes) {
      result.push(current);
      current = "";
    } else {
      current += char;
    }
  }
  result.push(current);
  return result.map((value) => value.trim());
}

function extractOperations(row) {
  const operations = [];
  for (let i = 1; i <= 10; i += 1) {
    const name = clean(row[`operation${i}Name`]);
    if (!name) continue;
    operations.push({ operationName: name, rate: num(row[`operation${i}Rate`]) });
  }
  return operations;
}

async function parseStyleImportFile(file) {
  const lowerName = clean(file?.name).toLowerCase();
  if (lowerName.endsWith(".xlsx")) {
    return parseStyleWorkbook(file);
  }
  return parseCsv(await file.text()).map((row) => normalizeStyleImportRow(row));
}

async function parseWorkbookOrCsvRows(file) {
  const lowerName = clean(file?.name).toLowerCase();
  if (lowerName.endsWith(".xlsx")) {
    return parseWorkbookRows(file);
  }
  return parseCsv(await file.text());
}

async function parseStyleWorkbook(file) {
  if (!window.ExcelJS?.Workbook) {
    throw new Error("Excel import is not available.");
  }
  const workbookBuffer = await file.arrayBuffer();
  const workbook = new window.ExcelJS.Workbook();
  await workbook.xlsx.load(workbookBuffer);
  const worksheet = workbook.worksheets.find((sheet) => sheet.actualRowCount > 0);
  if (!worksheet) return [];
  const imageByRow = mergeMaps(
    await extractWorkbookImagesByRowFromZip(workbookBuffer, worksheet.name),
    extractWorkbookImagesByRow(workbook, worksheet)
  );
  const rows = worksheet.getSheetValues().slice(1).map((row) => Array.isArray(row) ? row.slice(1) : []);
  const headerRowIndex = rows.findIndex((row) => detectStyleImportHeader(row));
  if (headerRowIndex === -1) return [];
  const headers = rows[headerRowIndex].map((cell) => clean(excelCellToText(cell)));
  const dataRows = rows.slice(headerRowIndex + 1);
  const normalizedHeaders = headers.map((header) => normalizeImportHeader(header));
  if (normalizedHeaders.includes("vendorstylecode")) {
    return parseVendorStyleWorkbookRows(headers, dataRows, headerRowIndex + 2, imageByRow);
  }
  return dataRows
    .map((row, index) => normalizeStyleImportRow(
      Object.fromEntries(headers.map((header, columnIndex) => [header, excelCellToText(row[columnIndex])])),
      imageByRow.get(headerRowIndex + 2 + index) || ""
    ))
    .filter((row) => Object.values(row).some((value) => clean(value)));
}

async function parseWorkbookRows(file) {
  if (!window.ExcelJS?.Workbook) {
    throw new Error("Excel import is not available.");
  }
  const workbook = new window.ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());
  const worksheet = workbook.worksheets.find((sheet) => sheet.actualRowCount > 0);
  if (!worksheet) return [];
  const rows = worksheet.getSheetValues().slice(1).map((row) => Array.isArray(row) ? row.slice(1) : []);
  const headerRow = rows.find((row) => row.some((cell) => clean(excelCellToText(cell))));
  if (!headerRow) return [];
  const headerIndex = rows.indexOf(headerRow);
  const headers = headerRow.map((cell) => clean(excelCellToText(cell)));
  return rows.slice(headerIndex + 1)
    .map((row) => Object.fromEntries(headers.map((header, index) => [header, excelCellToText(row[index])])))
    .filter((row) => Object.values(row).some((value) => clean(value)));
}

function parseVendorStyleWorkbookRows(headers, dataRows, startRowNumber = 2, imageByRow = new Map()) {
  let lastImage = "";
  let lastStyleNumber = "";
  return dataRows.map((row, index) => {
    const raw = Object.fromEntries(headers.map((header, index) => [header, excelCellToText(row[index])]));
    const sheetRowNumber = startRowNumber + index;
    const imageFromSheet = imageByRow.get(sheetRowNumber) || clean(raw.IMAGE || raw.Image || raw.image);
    const styleNumber = clean(raw["UVP Style ID"] || raw["Vendor Style Code"]);
    if (styleNumber !== lastStyleNumber) lastImage = "";
    if (imageFromSheet) lastImage = imageFromSheet;
    lastStyleNumber = styleNumber;
    const styleName = [clean(raw.Brick), clean(raw.Fabric)].filter(Boolean).join(" - ");
    const notes = [
      raw["Vendor Style Code"] ? `Vendor Style Code: ${clean(raw["Vendor Style Code"])}` : "",
      raw["Content Fabric"] ? `Content Fabric: ${clean(raw["Content Fabric"])}` : "",
      raw.Fabric ? `Fabric: ${clean(raw.Fabric)}` : "",
      raw["HSN Code"] ? `HSN Code: ${clean(raw["HSN Code"])}` : ""
    ].filter(Boolean).join(" | ");
    return normalizeStyleImportRow({
      styleNumber,
      buyerName: "",
      styleName,
      color: clean(raw["Vendor Color Name"]),
      orderQty: excelCellToText(raw["Vendor Capacity"]),
      cmtRate: "",
      serviceChargePct: "",
      image: imageFromSheet || lastImage,
      notes
    });
  }).filter((row) => row.styleNumber && row.color);
}

async function importStyleRows(rows, importedImageMap = new Map(), options = {}) {
  const summary = { created: 0, updated: 0, skipped: 0 };
  const rowsWithStyleImages = applySharedImagesByStyleNumber(rows, importedImageMap);
  for (const row of rowsWithStyleImages) {
    const styleNumber = clean(row.styleNumber);
    if (!styleNumber) continue;
    const matchingStyles = state.styles.filter((s) => s.styleNumber.toLowerCase() === styleNumber.toLowerCase());
    const existing = matchingStyles[0];
    const color = clean(row.color);
    const variant = matchingStyles.find((s) => clean(s.color).toLowerCase() === color.toLowerCase());
    const imageRecoveryOnly = isImageRecoveryImportRow(row);
    const updateImagesOnly = Boolean(options.updateImagesOnly);
    if (options.skipExisting && variant) {
      summary.skipped += 1;
      continue;
    }
    if ((imageRecoveryOnly || updateImagesOnly) && !matchingStyles.length) {
      summary.skipped += 1;
      continue;
    }
    const currentStyle = variant || existing || {};
    const importedOperations = extractOperations(row);
    const operations = importedOperations.length ? importedOperations : (Array.isArray(currentStyle.operations) ? currentStyle.operations : []);
    const resolvedImage = resolveImportedImage(row.image, importedImageMap);
    if (updateImagesOnly && !resolvedImage) {
      summary.skipped += 1;
      continue;
    }
    if (!color && resolvedImage && matchingStyles.length > 1 && (imageRecoveryOnly || updateImagesOnly)) {
      for (const style of matchingStyles) {
        style.image = await prepareStyleImageForState(resolvedImage || styleImageSrc(style) || style.image || "", style.id);
      }
      summary.updated += matchingStyles.length;
      continue;
    }
    if ((imageRecoveryOnly || updateImagesOnly) && (variant || existing)) {
      const targetStyle = variant || existing;
      targetStyle.image = await prepareStyleImageForState(resolvedImage || styleImageSrc(targetStyle) || targetStyle.image || "", targetStyle.id);
      summary.updated += 1;
      continue;
    }
    if (updateImagesOnly) {
      summary.skipped += 1;
      continue;
    }
    const styleId = variant?.id || uid();
    const payload = {
      id: styleId,
      styleNumber,
      buyerName: clean(row.buyerName) || clean(currentStyle.buyerName),
      styleName: clean(row.styleName) || clean(currentStyle.styleName),
      color,
      orderQty: hasImportValue(row.orderQty) ? num(row.orderQty) : (rowSizeTotal(row) || num(currentStyle.orderQty)),
      cmtRate: hasImportValue(row.cmtRate) ? num(row.cmtRate) : num(currentStyle.cmtRate),
      serviceChargePct: hasImportValue(row.serviceChargePct) ? num(row.serviceChargePct) : num(currentStyle.serviceChargePct),
      image: await prepareStyleImageForState(resolvedImage || styleImageSrc(variant) || styleImageSrc(existing) || existing?.image || "", styleId),
      notes: clean(row.notes) || clean(currentStyle.notes),
      operations
    };
    if (variant) {
      Object.assign(variant, payload);
      summary.updated += 1;
    } else if (existing && !color) {
      Object.assign(existing, payload);
      summary.updated += 1;
    } else {
      state.styles.push(payload);
      summary.created += 1;
    }
  }
  return summary;
}

function hasImportValue(value) {
  return clean(value) !== "";
}

function isImageRecoveryImportRow(row = {}) {
  return !clean(row.color)
    && !clean(row.buyerName)
    && !clean(row.styleName)
    && !clean(row.orderQty)
    && !clean(row.cmtRate)
    && !clean(row.serviceChargePct)
    && !clean(row.notes)
    && !extractOperations(row).length;
}

function applySharedImagesByStyleNumber(rows, importedImageMap = new Map()) {
  const imageByStyleNumber = new Map();
  rows.forEach((row) => {
    const styleNumber = normalizeKey(row.styleNumber);
    if (!styleNumber || imageByStyleNumber.has(styleNumber)) return;
    const resolvedImage = resolveImportedImage(row.image, importedImageMap);
    if (resolvedImage) imageByStyleNumber.set(styleNumber, resolvedImage);
  });
  return rows.map((row) => {
    const styleNumber = normalizeKey(row.styleNumber);
    const sharedImage = imageByStyleNumber.get(styleNumber);
    return sharedImage ? { ...row, image: sharedImage } : row;
  });
}

function detectStyleImportHeader(row = []) {
  const headers = row.map((cell) => normalizeImportHeader(excelCellToText(cell))).filter(Boolean);
  return headers.includes("stylenumber") || headers.includes("vendorstylecode");
}

function normalizeImportHeader(value) {
  return clean(value).toLowerCase().replace(/[^a-z0-9]+/g, "");
}

function normalizeStyleImportRow(row = {}, fallbackImage = "") {
  const normalized = {};
  Object.entries(row || {}).forEach(([rawKey, rawValue]) => {
    const key = clean(rawKey);
    const value = rawValue ?? "";
    if (!key) return;
    const normalizedKey = normalizeImportHeader(key);
    const mappedKey = STYLE_IMPORT_FIELD_ALIASES[normalizedKey] || normalizeOperationImportKey(normalizedKey) || key;
    if (!hasImportValue(normalized[mappedKey]) || hasImportValue(value)) {
      normalized[mappedKey] = value;
    }
  });
  if (fallbackImage && !clean(normalized.image)) {
    normalized.image = fallbackImage;
  }
  return normalized;
}

function normalizeOperationImportKey(normalizedKey) {
  const match = normalizedKey.match(/^operation(\d+)(name|rate)$/);
  if (!match) return "";
  return `operation${match[1]}${match[2] === "name" ? "Name" : "Rate"}`;
}

function operationRateByName(style, operationName) {
  const target = normalizeOperationName(operationName);
  return num(style?.operations?.find((operation) => normalizeOperationName(operation.operationName) === target)?.rate);
}

function normalizeOperationName(value) {
  return clean(value).toLowerCase().replace(/[^a-z0-9]+/g, "");
}

function upsertStyleOperation(style, operationName, rate) {
  const normalized = normalizeOperationName(operationName);
  const index = style.operations.findIndex((operation) => normalizeOperationName(operation.operationName) === normalized);
  if (index >= 0) {
    style.operations[index].operationName = operationName;
    style.operations[index].rate = num(rate);
  } else {
    style.operations.push({ operationName, rate: num(rate) });
  }
}

function pendingCmtStyles() {
  return state.styles.filter((style) => {
    const hasBuyer = Boolean(clean(style.buyerName));
    const hasCmt = num(style.cmtRate) > 0;
    const hasSinger = operationRateByName(style, "Singer") > 0;
    const hasOverlock = operationRateByName(style, "Overlock") > 0;
    const hasFlat = operationRateByName(style, "Flat") > 0;
    return !hasBuyer || !hasCmt || !hasSinger || !hasOverlock || !hasFlat;
  });
}

function pendingCmtReason(style) {
  const reasons = [];
  if (!clean(style.buyerName)) reasons.push("Buyer");
  if (num(style.cmtRate) <= 0) reasons.push("Total CMT");
  if (operationRateByName(style, "Singer") <= 0) reasons.push("Singer");
  if (operationRateByName(style, "Overlock") <= 0) reasons.push("Overlock");
  if (operationRateByName(style, "Flat") <= 0) reasons.push("Flat");
  return reasons.join(", ");
}

function excelCellToText(value) {
  if (value == null) return "";
  if (typeof value === "object") {
    if ("text" in value && clean(value.text)) return value.text;
    if ("result" in value && value.result != null) return excelCellToText(value.result);
    if ("richText" in value && Array.isArray(value.richText)) return value.richText.map((part) => part.text || "").join("");
    if ("hyperlink" in value && value.hyperlink) return String(value.hyperlink);
    if ("formula" in value && value.formula) return String(value.formula);
  }
  return String(value).trim();
}

function extractWorkbookImagesByRow(workbook, worksheet) {
  const map = new Map();
  const images = typeof worksheet.getImages === "function" ? worksheet.getImages() : [];
  images.forEach((image) => {
    try {
      const rowNumber = imageAnchorRowNumber(image);
      const dataUrl = workbookImageToDataUrl(workbook, image.imageId);
      if (rowNumber && dataUrl) map.set(rowNumber, dataUrl);
    } catch (error) {
      console.warn("Skipping workbook image during style import:", error);
    }
  });
  return map;
}

async function extractWorkbookImagesByRowFromZip(workbookBuffer, worksheetName = "") {
  const map = new Map();
  if (!window.JSZip || !window.DOMParser) return map;
  try {
    const zip = await window.JSZip.loadAsync(workbookBuffer);
    const sheetPath = await findWorkbookSheetPath(zip, worksheetName);
    if (!sheetPath) return map;
    const sheetXml = await readZipText(zip, sheetPath);
    const sheetDoc = parseXml(sheetXml);
    const drawingRelId = [...sheetDoc.getElementsByTagNameNS("*", "drawing")]
      .map((node) => node.getAttribute("r:id") || node.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id"))
      .find(Boolean);
    if (!drawingRelId) return map;

    const sheetRelsPath = relationshipPath(sheetPath);
    const sheetRelTarget = await findRelationshipTarget(zip, sheetRelsPath, drawingRelId);
    const drawingPath = resolveZipPath(sheetPath, sheetRelTarget);
    if (!drawingPath) return map;

    const drawingXml = await readZipText(zip, drawingPath);
    const drawingDoc = parseXml(drawingXml);
    const drawingRelsPath = relationshipPath(drawingPath);
    const mediaByRelId = await readRelationshipTargets(zip, drawingRelsPath);
    const anchors = [
      ...drawingDoc.getElementsByTagNameNS("*", "oneCellAnchor"),
      ...drawingDoc.getElementsByTagNameNS("*", "twoCellAnchor")
    ];
    for (const anchor of anchors) {
      const from = anchor.getElementsByTagNameNS("*", "from")[0];
      const rowNode = from?.getElementsByTagNameNS("*", "row")?.[0];
      const rowNumber = Number(rowNode?.textContent || "") + 1;
      const blip = anchor.getElementsByTagNameNS("*", "blip")[0];
      const embedId = blip?.getAttribute("r:embed") || blip?.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "embed");
      const mediaTarget = mediaByRelId.get(embedId);
      const mediaPath = resolveZipPath(drawingPath, mediaTarget);
      const mediaFile = mediaPath ? zip.file(mediaPath) : null;
      if (!rowNumber || !mediaFile) continue;
      const extension = clean(mediaPath.split(".").pop() || "png").toLowerCase() || "png";
      const base64 = await mediaFile.async("base64");
      if (base64) map.set(rowNumber, `data:image/${extension};base64,${base64}`);
    }
  } catch (error) {
    console.warn("Could not extract workbook images from the XLSX package:", error);
  }
  return map;
}

async function findWorkbookSheetPath(zip, worksheetName = "") {
  const workbookXml = await readZipText(zip, "xl/workbook.xml");
  const workbookDoc = parseXml(workbookXml);
  const rels = await readRelationshipTargets(zip, "xl/_rels/workbook.xml.rels");
  const sheets = [...workbookDoc.getElementsByTagNameNS("*", "sheet")];
  const targetSheet = sheets.find((sheet) => clean(sheet.getAttribute("name")) === clean(worksheetName)) || sheets[0];
  const relId = targetSheet?.getAttribute("r:id") || targetSheet?.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
  return resolveZipPath("xl/workbook.xml", rels.get(relId));
}

async function findRelationshipTarget(zip, relsPath, relId) {
  const targets = await readRelationshipTargets(zip, relsPath);
  return targets.get(relId) || "";
}

async function readRelationshipTargets(zip, relsPath) {
  const map = new Map();
  const xml = await readZipText(zip, relsPath);
  if (!xml) return map;
  const doc = parseXml(xml);
  [...doc.getElementsByTagNameNS("*", "Relationship")].forEach((rel) => {
    map.set(rel.getAttribute("Id"), rel.getAttribute("Target"));
  });
  return map;
}

async function readZipText(zip, path) {
  const file = path ? zip.file(path) : null;
  return file ? file.async("text") : "";
}

function parseXml(xml) {
  return new DOMParser().parseFromString(xml, "application/xml");
}

function relationshipPath(path) {
  const parts = clean(path).split("/");
  const fileName = parts.pop();
  return `${parts.join("/")}/_rels/${fileName}.rels`;
}

function resolveZipPath(fromPath, targetPath) {
  const target = clean(targetPath);
  if (!target) return "";
  if (target.startsWith("/")) return target.slice(1);
  const stack = clean(fromPath).split("/");
  stack.pop();
  target.split("/").forEach((part) => {
    if (!part || part === ".") return;
    if (part === "..") stack.pop();
    else stack.push(part);
  });
  return stack.join("/");
}

function mergeMaps(...maps) {
  const merged = new Map();
  maps.forEach((map) => {
    map.forEach((value, key) => {
      if (value) merged.set(key, value);
    });
  });
  return merged;
}

function imageAnchorRowNumber(image) {
  const nativeRow = image?.range?.tl?.nativeRow;
  if (typeof nativeRow === "number") return nativeRow + 1;
  const directRow = image?.range?.tl?.row;
  if (typeof directRow === "number") return directRow + 1;
  const alternate = image?.range?.br?.nativeRow;
  if (typeof alternate === "number") return alternate + 1;
  return 0;
}

function workbookImageToDataUrl(workbook, imageId) {
  const media = (workbook.model?.media || []).find((item) => item.index === imageId || item.id === imageId);
  if (!media) return "";
  const extension = clean(media.extension || "png").toLowerCase() || "png";
  if (clean(media.base64)) {
    return media.base64.startsWith("data:image/")
      ? media.base64
      : `data:image/${extension};base64,${media.base64}`;
  }
  const binary = normalizeBinaryData(media.buffer || media.data);
  if (binary) {
    return `data:image/${extension};base64,${bytesToBase64(binary)}`;
  }
  return "";
}

function bytesToBase64(bufferLike) {
  const bytes = bufferLike instanceof Uint8Array ? bufferLike : new Uint8Array(bufferLike);
  let binary = "";
  const chunkSize = 32768;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    const chunk = bytes.subarray(i, i + chunkSize);
    binary += String.fromCharCode(...chunk);
  }
  return btoa(binary);
}

function normalizeBinaryData(value) {
  if (!value) return null;
  if (value instanceof Uint8Array) return value;
  if (value instanceof ArrayBuffer) return new Uint8Array(value);
  if (Array.isArray(value)) return new Uint8Array(value);
  if (ArrayBuffer.isView(value)) return new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
  if (typeof value === "object" && Array.isArray(value.data)) return new Uint8Array(value.data);
  if (typeof value === "string") {
    const text = value.startsWith("data:") ? value.split(",", 2)[1] || "" : value;
    if (!text) return null;
    const binary = atob(text);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
    return bytes;
  }
  return null;
}

function sizeQuantitiesFromRow(row) {
  const valuesByNormalizedHeader = normalizedRowValues(row);
  return Object.fromEntries(getSizes().map((size) => [size, num(valueForSizeHeader(valuesByNormalizedHeader, size))]));
}

function findStyleByNumber(styleNumber, color = "") {
  const styleNumberText = normalizeStyleLookupValue(styleNumber);
  const colorText = normalizeStyleLookupValue(color);
  return state.styles.find((style) => normalizeStyleLookupValue(style.styleNumber) === styleNumberText && (!colorText || normalizeStyleLookupValue(style.color) === colorText));
}

function rowSizeTotal(row = {}) {
  const valuesByNormalizedHeader = normalizedRowValues(row);
  return getSizes().reduce((total, size) => total + num(valueForSizeHeader(valuesByNormalizedHeader, size)), 0);
}

function normalizedRowValues(row = {}) {
  const values = new Map();
  Object.entries(row || {}).forEach(([key, value]) => {
    const normalizedKey = normalizeSizeHeader(key);
    if (normalizedKey && (!values.has(normalizedKey) || hasImportValue(value))) {
      values.set(normalizedKey, value);
    }
  });
  return values;
}

function normalizeSizeHeader(value) {
  return clean(value).toLowerCase().replace(/[^a-z0-9]+/g, "");
}

function valueForSizeHeader(valuesByNormalizedHeader, size) {
  const candidates = sizeHeaderCandidates(size);
  const matchedKey = candidates.find((key) => valuesByNormalizedHeader.has(key));
  return matchedKey ? valuesByNormalizedHeader.get(matchedKey) : "";
}

function sizeHeaderCandidates(size) {
  const text = clean(size);
  const parts = text.split(/[\/|]/).map((part) => clean(part)).filter(Boolean);
  return [...new Set([
    normalizeSizeHeader(text),
    ...parts.map((part) => normalizeSizeHeader(part))
  ].filter(Boolean))];
}

function normalizeStyleLookupValue(value) {
  return clean(value).toLowerCase().replace(/\s+/g, " ").replace(/[\u2010-\u2015]/g, "-").trim();
}

function handleStyleCardAction(e) {
  const button = e.target.closest("[data-action]");
  if (!button) return;
  const styleId = button.dataset.styleId;
  if (button.dataset.action === "edit-style") editStyle(styleId);
  if (button.dataset.action === "delete-style") void deleteStyle(styleId);
  if (button.dataset.action === "preview-image") openImagePreview(button.dataset.imageSrc, button.dataset.imageTitle);
}

function editStyle(styleId) {
  const style = byId(styleId);
  if (!style) return;
  clearPastedStyleImage();
  els.styleForm.dataset.editId = style.id;
  els.styleForm.styleNumber.value = style.styleNumber;
  els.styleForm.buyerName.value = style.buyerName || "";
  els.styleForm.styleName.value = style.styleName || "";
  els.styleForm.cmtRate.value = style.cmtRate || "";
  els.styleForm.serviceChargePct.value = style.serviceChargePct || "";
  const imageSource = styleImageSrc(style);
  els.styleForm.image.value = imageSource && !imageSource.startsWith("data:") ? imageSource : "";
  els.styleForm.querySelector('[name="imageFile"]').value = "";
  els.styleForm.notes.value = style.notes || "";
  resetStyleVariantRows([{ color: style.color || "", orderQty: style.orderQty || "" }]);
  els.operationRateRows.innerHTML = "";
  (style.operations.length ? style.operations : [{ operationName: "", rate: "" }]).forEach((op) => addOperationRow(op.operationName, op.rate));
  updateStyleFormImagePreview(imageSource || "");
  els.styleForm.scrollIntoView({ behavior: "smooth", block: "start" });
}

async function deleteStyle(styleId) {
  const style = byId(styleId);
  if (!style) return;
  if (hasStyleTransactions(styleId)) {
    alert("This style cannot be deleted because cutting, production, or acceptance data already exists for it.");
    return;
  }
  const confirmed = window.confirm(`Delete style ${style.styleNumber}?`);
  if (!confirmed) return;
  state.styles = state.styles.filter((s) => s.id !== styleId);
  await deleteStoredStyleImage(styleId);
  if (els.styleForm.dataset.editId === styleId) {
    delete els.styleForm.dataset.editId;
    resetStyleFormState();
  }
  await persistState();
}

function hasStyleTransactions(styleId) {
  return state.cuttingEntries.some((e) => e.styleId === styleId)
    || state.styleProductionEntries.some((e) => e.styleId === styleId)
    || state.productionEntries.some((e) => e.styleId === styleId)
    || state.acceptanceEntries.some((e) => e.styleId === styleId)
    || state.dispatchEntries.some((e) => e.styleId === styleId)
    || state.washcareRecords.some((e) => e.styleId === styleId);
}

function handleCuttingTableAction(e) {
  const button = e.target.closest("[data-action]");
  if (!button) return;
  if (button.dataset.action === "edit-cutting") editCutting(button.dataset.entryId);
  if (button.dataset.action === "delete-cutting") void deleteEntry("cuttingEntries", button.dataset.entryId, "cutting entry");
}

function handleStyleProductionTableAction(e) {
  const button = e.target.closest("[data-action]");
  if (!button) return;
  if (button.dataset.action === "edit-style-production") editStyleProduction(button.dataset.entryId);
  if (button.dataset.action === "delete-style-production") void deleteEntry("styleProductionEntries", button.dataset.entryId, "style production entry");
}

function handleProductionTableAction(e) {
  const button = e.target.closest("[data-action]");
  if (!button) return;
  if (button.dataset.action === "edit-production") editProduction(button.dataset.entryId);
  if (button.dataset.action === "delete-production") void deleteEntry("productionEntries", button.dataset.entryId, "production entry");
}

function handleAcceptanceTableAction(e) {
  const button = e.target.closest("[data-action='edit-acceptance']");
  if (!button) return;
  editAcceptance(button.dataset.entryId);
}

function handleDispatchTableAction(e) {
  const button = e.target.closest("[data-action]");
  if (!button) return;
  if (button.dataset.action === "edit-dispatch") editDispatch(button.dataset.entryId);
  if (button.dataset.action === "delete-dispatch") void deleteEntry("dispatchEntries", button.dataset.entryId, "dispatch entry");
}

function handlePaymentTableAction(e) {
  const button = e.target.closest("[data-action='delete-payment']");
  if (!button) return;
  void deleteEntry("payments", button.dataset.entryId, "payment entry");
}

function handleWashcareTableAction(e) {
  const button = e.target.closest("[data-action]");
  if (!button) return;
  if (button.dataset.action === "edit-washcare") editWashcare(button.dataset.entryId);
  if (button.dataset.action === "delete-washcare") void deleteEntry("washcareRecords", button.dataset.entryId, "washcare record");
  if (button.dataset.action === "print-washcare") handleWashcarePrint(button.dataset.entryId);
}

async function deleteEntry(collectionName, entryId, label) {
  const entry = state[collectionName]?.find((item) => item.id === entryId);
  if (!entry) return;
  const confirmed = window.confirm(`Delete this ${label}?`);
  if (!confirmed) return;
  state[collectionName] = state[collectionName].filter((item) => item.id !== entryId);
  await persistState();
}

function editWashcare(entryId) {
  const record = state.washcareRecords.find((item) => item.id === entryId);
  if (!record) return;
  populateWashcareForm(record);
  els.washcareForm?.scrollIntoView({ behavior: "smooth", block: "start" });
}

function handleWashcareStyleChange() {
  const styleId = clean(els.washcareStyleSelect?.value);
  if (!styleId) {
    delete els.washcareForm.dataset.editId;
    renderWashcarePreview();
    return;
  }
  populateWashcareFormByStyle(styleId);
}

async function importGrnPdf(e) {
  const files = [...(e.target.files || [])];
  if (!files.length) return;
  let importedCount = 0;
  const failures = [];
  for (const file of files) {
    if (els.grnStatusText) {
      els.grnStatusText.textContent = `Reading GRN PDF: ${file.name}`;
    }
    try {
      const extracted = await extractGrnReportDetails(file);
      const report = {
        ...clone(DEFAULTS.grnStatus),
        ...extracted,
        id: uid(),
        grnFileName: file.name,
        updatedAt: new Date().toISOString()
      };
      state.grnStatus = report;
      upsertGrnReport(report);
      importedCount += 1;
    } catch (error) {
      console.error(error);
      failures.push(`${file.name}${clean(error?.message) ? ` (${clean(error.message)})` : ""}`);
    }
  }
  await persistState();
  if (els.grnStatusText) {
    if (importedCount && !failures.length) {
      els.grnStatusText.textContent = importedCount === 1
        ? `Loaded and saved 1 GRN report.`
        : `Loaded and saved ${importedCount} GRN reports.`;
    } else if (importedCount && failures.length) {
      els.grnStatusText.textContent = `Saved ${importedCount} GRN reports. Failed to read: ${failures.join(", ")}.`;
    } else {
      els.grnStatusText.textContent = "Could not read the selected GRN PDF files automatically. Please try the launcher file and upload again.";
    }
  }
  if (!importedCount) {
    alert("Could not read the selected GRN PDF files.");
  }
  if (els.grnPdfInput) els.grnPdfInput.value = "";
}

async function saveCurrentGrnReport() {
  if (!state.grnStatus?.grnFileName || !(state.grnStatus.items || []).length) {
    alert("Please upload a GRN PDF first.");
    return;
  }
  const report = {
    ...clone(DEFAULTS.grnStatus),
    ...state.grnStatus,
    id: clean(state.grnStatus.id) || uid(),
    updatedAt: new Date().toISOString()
  };
  state.grnStatus = report;
  upsertGrnReport(report);
  await persistState();
  if (els.grnStatusText) {
    els.grnStatusText.textContent = `GRN ${report.grnNumber || "-"} saved successfully.`;
  }
}

function upsertGrnReport(report) {
  const reportKey = normalizeKey(report.grnNumber) || normalizeKey(report.vendorInvoiceNo) || clean(report.id);
  const index = (state.grnReports || []).findIndex((item) => {
    const itemKey = normalizeKey(item.grnNumber) || normalizeKey(item.vendorInvoiceNo) || clean(item.id);
    return itemKey && itemKey === reportKey;
  });
  if (index >= 0) state.grnReports[index] = report;
  else state.grnReports.push(report);
}

async function importPackingListPdf(e) {
  const file = e.target.files?.[0];
  if (!file) return;
  if (!state.grnStatus?.grnFileName) {
    alert("Please upload the GRN PDF first.");
    if (els.grnPackingListInput) els.grnPackingListInput.value = "";
    return;
  }
  if (els.grnStatusText) {
    els.grnStatusText.textContent = `Reading packing list file: ${file.name}`;
  }
  try {
    const extracted = /\.xlsx$/i.test(file.name)
      ? await extractPackingListWorkbookDetails(file)
      : await extractPackingListDetails(file);
    const comparisonRows = buildGrnComparisonRows(state.grnStatus.items, extracted.items);
    state.grnStatus = {
      ...state.grnStatus,
      packingListFileName: file.name,
      packingItems: extracted.items,
      packingListInvoiceNo: extracted.invoiceNo,
      packingListRawText: extracted.rawText,
      comparisonRows,
      updatedAt: new Date().toISOString()
    };
    await persistState();
  } catch (error) {
    console.error(error);
    if (els.grnStatusText) {
      els.grnStatusText.textContent = "Could not read the packing list file automatically. We can refine this once you share one sample packing list.";
    }
    alert("Could not read the selected packing list file.");
  } finally {
    if (els.grnPackingListInput) els.grnPackingListInput.value = "";
  }
}

async function extractPdfPlainText(file) {
  if (canUseServerPdfExtraction()) {
    try {
      return await extractPdfPlainTextViaServer(file);
    } catch (error) {
      console.error("Server-side PDF extraction failed, falling back to browser parser.", error);
    }
  }
  if (!window.pdfjsLib?.getDocument) {
    throw new Error("PDF library not loaded.");
  }
  const arrayBuffer = await file.arrayBuffer();
  const loadingTask = window.pdfjsLib.getDocument({
    data: arrayBuffer,
    isEvalSupported: false,
    useWorkerFetch: false
  });
  const pdf = await loadingTask.promise;
  const pageTexts = [];
  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const content = await page.getTextContent();
    pageTexts.push(buildPdfPageText(content.items || []));
  }
  const pages = normalizePdfPages(pageTexts);
  const fullText = pages.map((page) => page.text).join("\n");
  return {
    fullText,
    normalized: fullText.replace(/\s+/g, " ").trim(),
    pages
  };
}

function canUseServerPdfExtraction() {
  const protocol = clean(window.location?.protocol);
  if (!["http:", "https:"].includes(protocol)) return false;
  return true;
}

async function extractPdfPlainTextViaServer(file) {
  const response = await fetch("/extract-pdf-text", {
    method: "POST",
    headers: {
      "Content-Type": "application/pdf"
    },
    body: file
  });
  const data = await response.json().catch(() => null);
  if (!response.ok || !data) {
    throw new Error(data?.details || data?.error || `Server extraction failed with HTTP ${response.status}.`);
  }
  const fullText = clean(data.rawText || data.text);
  const normalized = clean(data.normalizedText || data.text);
  const pages = normalizePdfPages(
    Array.isArray(data.rawPages) && data.rawPages.length ? data.rawPages : [fullText],
    Array.isArray(data.normalizedPages) ? data.normalizedPages : []
  );
  return {
    fullText,
    normalized,
    pages
  };
}

async function extractGrnReportDetails(file) {
  const pdfText = await extractPdfPlainText(file);
  const { fullText, normalized } = pdfText;
  const items = parseGrnItems(pdfText);
  if (!items.length) {
    throw new Error("No GRN item rows were found in the PDF.");
  }
  return {
    pdfPageCount: Array.isArray(pdfText.pages) ? pdfText.pages.length : 0,
    grnNumber: extractReportField(normalized, /Goods Receipt Note No\.\s*:?\s*([A-Za-z0-9/-]+)/i),
    grnDate: normalizeDateValue(extractReportField(normalized, /Goods Receipt Note No\.\s*:?\s*[A-Za-z0-9/-]+\s+Date:?\s*([0-9./-]+)/i)),
    supplier: extractReportField(normalized, /Supplier\s*:?\s*(.+?)\s+E-\d+/i) || extractReportField(normalized, /NEXTGEN FAST FASHION LIMITED\s+Supplier\s*:?\s*(.+?)\s+Code/i),
    supplierCode: extractReportField(normalized, /Code\s*:?\s*([A-Za-z0-9/-]+)/i),
    supplierGstin: extractReportField(normalized, /GSTIN NO\s*:?\s*([A-Z0-9]+)/i),
    vendorInvoiceNo: extractReportField(normalized, /Vendor invoice no\s*:?\s*([A-Za-z0-9/-]+)/i),
    poNumber: extractReportField(normalized, /PO Number\s*:?\s*([A-Za-z0-9/-]+)/i),
    poDate: normalizeDateValue(extractReportField(normalized, /PO Number\s*:?\s*[A-Za-z0-9/-]+\s+Date\s*:?\s*([0-9./-]+)/i)),
    deliveryNo: extractReportField(normalized, /Delivery No\s*:?\s*([A-Za-z0-9/-]+)/i),
    consignee: extractReportField(normalized, /Consignee\s*:?\s*(.+?)\s+GSTIN NO/i),
    warehouse: extractReportField(normalized, /Delivery Addr:\s*(.+?)\s+GSTIN NO/i) || extractReportField(normalized, /Consignee\s*:?\s*(.+?)\s+NEXTGEN FAST FASHION/i),
    transporterName: extractReportField(normalized, /Transporter Details:\s*Name\s*:?\s*(.+?)\s+Consignment Note/i),
    consignmentNote: extractReportField(normalized, /Consignment Note\s*:?\s*([A-Za-z0-9/-]+)/i),
    consignmentDate: normalizeDateValue(extractReportField(normalized, /C\/N Date\s*:?\s*([0-9./-]+)/i)),
    vehicleNumber: extractReportField(normalized, /Truck\/ Lorry\/ Carrier No:\s*([A-Za-z0-9-]+)/i),
    deliveryChallanNo: extractReportField(normalized, /Delivery Challan No:\s*([A-Za-z0-9/-]+)/i),
    totalChallanQty: items.reduce((sum, item) => sum + num(item.challanQty), 0),
    totalReceivedQty: items.reduce((sum, item) => sum + num(item.receivedQty), 0),
    totalAcceptedQty: items.reduce((sum, item) => sum + num(item.acceptedQty), 0),
    totalShortQty: items.reduce((sum, item) => sum + num(item.shortQty), 0),
    items,
    comparisonRows: [],
    packingItems: [],
    packingListFileName: "",
    packingListInvoiceNo: "",
    packingListRawText: fullText
  };
}

function buildPdfPageText(items = []) {
  const lines = [];
  let currentLine = [];
  let currentY = null;
  sortedPdfTextItems(items).forEach((item) => {
    const text = clean(item?.str);
    const y = Array.isArray(item?.transform) ? Number(item.transform[5]) : null;
    const shouldBreakLine = currentLine.length
      && currentY !== null
      && y !== null
      && Math.abs(y - currentY) > 2;
    if (shouldBreakLine) {
      lines.push(currentLine.join(" ").replace(/\s+/g, " ").trim());
      currentLine = [];
    }
    if (text) currentLine.push(text);
    if (y !== null) currentY = y;
    if (item?.hasEOL && currentLine.length) {
      lines.push(currentLine.join(" ").replace(/\s+/g, " ").trim());
      currentLine = [];
      currentY = null;
    }
  });
  if (currentLine.length) {
    lines.push(currentLine.join(" ").replace(/\s+/g, " ").trim());
  }
  return lines.join("\n");
}

function sortedPdfTextItems(items = []) {
  return [...items].sort((a, b) => {
    const ay = Array.isArray(a?.transform) ? Number(a.transform[5]) : 0;
    const by = Array.isArray(b?.transform) ? Number(b.transform[5]) : 0;
    if (Math.abs(by - ay) > 2) return by - ay;
    const ax = Array.isArray(a?.transform) ? Number(a.transform[4]) : 0;
    const bx = Array.isArray(b?.transform) ? Number(b.transform[4]) : 0;
    return ax - bx;
  });
}

function normalizePdfPages(rawPages = [], normalizedPages = []) {
  return rawPages
    .map((pageText, index) => {
      const text = cleanPdfMultilineText(pageText);
      const normalized = clean(normalizedPages[index] || text.replace(/\s+/g, " "));
      return {
        pageNumber: index + 1,
        text,
        normalized
      };
    })
    .filter((page) => page.text);
}

function cleanPdfMultilineText(value) {
  return String(value || "")
    .replace(/\r/g, "")
    .replace(/\u00a0/g, " ")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n[ \t]+/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function sliceGrnTableSection(text) {
  const source = String(text || "");
  const start = source.search(/S No\s+Article/i);
  if (start < 0) return source;
  const section = source.slice(start);
  const markerPatterns = [/Total Qty UOM wise/i, /Reason\s*:/i];
  let end = section.length;
  markerPatterns.forEach((pattern) => {
    const index = section.search(pattern);
    if (index >= 0) end = Math.min(end, index);
  });
  return section.slice(0, end).trim();
}

function parseGrnItems({ fullText = "", normalized = "", pages = [] } = {}) {
  const pageItems = pages.flatMap((page) => {
    const rawTableText = sliceGrnTableSection(page?.text);
    const normalizedTableText = sliceGrnTableSection(page?.normalized || page?.text);
    const lineItems = parseGrnItemsFromLines(rawTableText);
    return lineItems.length ? lineItems : parseGrnItemsFromNormalizedText(normalizedTableText);
  });
  if (pageItems.length) return dedupeGrnItems(pageItems);
  const rawTableText = sliceGrnTableSection(fullText);
  const lineItems = parseGrnItemsFromLines(rawTableText);
  if (lineItems.length) return lineItems;
  const normalizedTableText = sliceGrnTableSection(normalized || fullText);
  return parseGrnItemsFromNormalizedText(normalizedTableText);
}

function parseGrnItemsFromLines(text) {
  const items = [];
  const pattern = /^\s*(\d+)\s+(\d{6,})\s+(.+?)\s+(\d{8,14})\s+(?:(\S+)\s+)?(EA|PCS|PC|NOS)\s+(\d+)\s+(\d+)\s+(\d+)(?:\s+(\d+))?(?:\s+(.+?))?\s*$/i;
  const lines = String(text || "")
    .split(/\r?\n/)
    .map((line) => line.replace(/\s+/g, " ").trim())
    .filter(Boolean);
  for (const line of lines) {
    const match = line.match(pattern);
    if (!match) continue;
    const [, serialNo, article, description, ean, vendorArticleNo, uom, challanQty, receivedQty, acceptedQty, shortQty, reason] = match;
    items.push(createGrnItem({
      serialNo,
      article,
      description,
      ean,
      vendorArticleNo,
      uom,
      challanQty,
      receivedQty,
      acceptedQty,
      shortQty,
      reason
    }));
  }
  return dedupeGrnItems(items);
}

function parseGrnItemsFromNormalizedText(text) {
  const pattern = /(\d+)\s+(\d{6,})\s+(.+?)\s+(\d{8,14})\s+(?:(\S+)\s+)?(EA|PCS|PC|NOS)\s+(\d+)\s+(\d+)\s+(\d+)(?:\s+(\d+))?(?:\s+([A-Za-z][A-Za-z.\s]+?))?(?=\s+\d+\s+\d{6,}|\s*$)/gi;
  const items = [];
  let match;
  while ((match = pattern.exec(text))) {
    const [, serialNo, article, description, ean, vendorArticleNo, uom, challanQty, receivedQty, acceptedQty, shortQty, reason] = match;
    items.push(createGrnItem({
      serialNo,
      article,
      description,
      ean,
      vendorArticleNo,
      uom,
      challanQty,
      receivedQty,
      acceptedQty,
      shortQty,
      reason
    }));
  }
  return dedupeGrnItems(items);
}

function createGrnItem({
  serialNo,
  article,
  description,
  ean,
  vendorArticleNo,
  uom,
  challanQty,
  receivedQty,
  acceptedQty,
  shortQty,
  reason
}) {
  return {
    serialNo: num(serialNo),
    article: clean(article),
    description: clean(String(description || "").replace(/\s+/g, " ")),
    ean: clean(ean),
    vendorArticleNo: clean(vendorArticleNo),
    uom: clean(uom) || "EA",
    challanQty: num(challanQty),
    receivedQty: num(receivedQty),
    acceptedQty: num(acceptedQty),
    shortQty: num(shortQty) || Math.max(num(challanQty) - num(acceptedQty), 0),
    reason: clean(reason) || (num(challanQty) > num(acceptedQty) ? "Short Recvd." : "")
  };
}

function dedupeGrnItems(items) {
  const seen = new Set();
  return items.filter((item) => {
    const key = [
      num(item.serialNo),
      normalizeKey(item.article),
      normalizeKey(item.ean),
      num(item.challanQty),
      num(item.acceptedQty)
    ].join("__");
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

async function extractPackingListDetails(file) {
  const { fullText, normalized } = await extractPdfPlainText(file);
  return {
    invoiceNo: extractReportField(normalized, /(?:invoice|inv\.?)\s*(?:no|number)\s*:?\s*([A-Za-z0-9/-]+)/i) || extractReportField(normalized, /delivery challan no\s*:?\s*([A-Za-z0-9/-]+)/i),
    items: parsePackingListItems(normalized),
    rawText: fullText
  };
}

async function extractPackingListWorkbookDetails(file) {
  if (!window.ExcelJS?.Workbook) {
    throw new Error("Excel library not loaded.");
  }
  const workbook = new window.ExcelJS.Workbook();
  const buffer = await file.arrayBuffer();
  await workbook.xlsx.load(buffer);
  const sheet = workbook.worksheets[0];
  if (!sheet) {
    throw new Error("Packing list workbook is empty.");
  }

  let invoiceNo = "";
  let invoiceDate = "";
  let asnNo = "";
  let poNumber = "";
  let headerRowNumber = 0;

  for (let rowNumber = 1; rowNumber <= Math.min(sheet.rowCount, 20); rowNumber += 1) {
    const values = getWorksheetRowValues(sheet.getRow(rowNumber));
    const joined = values.join(" | ");
    if (!invoiceNo && /invoice\s*no/i.test(joined)) invoiceNo = valueAfterLabel(values, /invoice\s*no/i);
    if (!invoiceDate && /invoice\s*date/i.test(joined)) invoiceDate = valueAfterLabel(values, /invoice\s*date/i);
    if (!asnNo && /asn\s*no/i.test(joined)) asnNo = valueAfterLabel(values, /asn\s*no/i);
    if (!poNumber && /po\s*no/i.test(joined)) poNumber = valueAfterLabel(values, /po\s*no/i);
    if (!headerRowNumber && values.some((value) => /rr\s*sku\s*code/i.test(value)) && values.some((value) => /\bqty\b/i.test(value))) {
      headerRowNumber = rowNumber;
    }
  }

  if (!headerRowNumber) {
    throw new Error("Could not find the packing list item table.");
  }

  const headerValues = getWorksheetRowValues(sheet.getRow(headerRowNumber)).map((value) => normalizeKey(value));
  const itemCol = findHeaderIndex(headerValues, ["item"]);
  const skuCol = findHeaderIndex(headerValues, ["skucode"]);
  const rrSkuCol = findHeaderIndex(headerValues, ["rrskucode"]);
  const qtyCol = findHeaderIndex(headerValues, ["qty"]);

  const items = [];
  for (let rowNumber = headerRowNumber + 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    const values = getWorksheetRowValues(row);
    const description = values[itemCol] || "";
    const article = values[rrSkuCol] || "";
    const vendorArticleNo = values[skuCol] || "";
    const qty = num(values[qtyCol] || 0);
    if (!clean(description) && !clean(article) && !clean(vendorArticleNo) && !qty) continue;
    items.push({
      article: clean(article),
      description: clean(description),
      ean: "",
      vendorArticleNo: clean(vendorArticleNo),
      uom: "EA",
      qty
    });
  }

  return {
    invoiceNo: clean(invoiceNo),
    invoiceDate: normalizeDateValue(invoiceDate),
    asnNo: clean(asnNo),
    poNumber: clean(poNumber),
    items: dedupePackingItems(items),
    rawText: `${sheet.name} workbook import`
  };
}

function getWorksheetRowValues(row) {
  return Array.from({ length: row.cellCount }, (_, index) => clean(cellText(row.getCell(index + 1).value)));
}

function cellText(value) {
  if (value == null) return "";
  if (typeof value === "object") {
    if (typeof value.text === "string") return value.text;
    if (typeof value.result !== "undefined") return String(value.result ?? "");
    if (typeof value.richText !== "undefined") return value.richText.map((part) => part.text || "").join("");
  }
  return String(value);
}

function valueAfterLabel(values, labelPattern) {
  const index = values.findIndex((value) => labelPattern.test(value));
  if (index < 0) return "";
  return clean(values[index + 1] || values[index + 2] || "");
}

function findHeaderIndex(headers, aliases) {
  const index = headers.findIndex((value) => aliases.includes(value));
  return index >= 0 ? index : 0;
}

function parsePackingListItems(text) {
  const results = [];
  const articlePattern = /(\d{12,})\s+(.+?)\s+(\d{13})\s+([A-Za-z0-9/-]+)\s+(EA|PCS|PC|NOS)\s+(\d+)(?!\s+\d)/gi;
  let match;
  while ((match = articlePattern.exec(text))) {
    const [, article, description, ean, vendorArticleNo, uom, qty] = match;
    results.push({
      article: clean(article),
      description: clean(description.replace(/\s+/g, " ")),
      ean: clean(ean),
      vendorArticleNo: clean(vendorArticleNo),
      uom: clean(uom),
      qty: num(qty)
    });
  }
  if (results.length) return dedupePackingItems(results);

  const sizePattern = /([A-Za-z0-9 .,&/-]+?,\s*[A-Za-z]{1,4})\s+(\d+)(?!\s+\d)/gi;
  while ((match = sizePattern.exec(text))) {
    const [, description, qty] = match;
    results.push({
      article: "",
      description: clean(description.replace(/\s+/g, " ")),
      ean: "",
      vendorArticleNo: "",
      uom: "EA",
      qty: num(qty)
    });
  }
  return dedupePackingItems(results);
}

function dedupePackingItems(items = []) {
  const map = new Map();
  items.forEach((item) => {
    const key = buildItemMatchKey(item);
    if (!key) return;
    const current = map.get(key) || { ...item, qty: 0 };
    current.qty += num(item.qty);
    map.set(key, current);
  });
  return [...map.values()];
}

function buildGrnComparisonRows(grnItems = [], packingItems = []) {
  const grnMap = new Map();
  grnItems.forEach((item) => {
    const key = buildItemMatchKey(item);
    if (!key) return;
    const current = grnMap.get(key) || { ...item, challanQty: 0, acceptedQty: 0 };
    current.challanQty += num(item.challanQty);
    current.acceptedQty += num(item.acceptedQty);
    current.description = current.description || item.description;
    grnMap.set(key, current);
  });
  return packingItems.map((item) => {
    const key = buildItemMatchKey(item);
    const matched = grnMap.get(key) || {};
    const packingQty = num(item.qty);
    const acceptedQty = num(matched.acceptedQty);
    return {
      matchKey: key,
      matchLabel: item.description || item.vendorArticleNo || item.article || item.ean || key,
      packingQty,
      challanQty: num(matched.challanQty),
      acceptedQty,
      shortageQty: Math.max(packingQty - acceptedQty, 0)
    };
  });
}

function buildItemMatchKey(item = {}) {
  return normalizeKey(item.article) || normalizeKey(item.ean) || normalizeKey(item.vendorArticleNo) || normalizeKey(item.description);
}

async function downloadGrnSheetWorkbook() {
  const grn = state.grnStatus || {};
  if (!(grn.items || []).length) {
    alert("No GRN detail sheet available to export yet.");
    return;
  }
  const workbook = createWorkbookOrAlert();
  if (!workbook) return;
  const sheet = workbook.addWorksheet("GRN Detail Sheet");
  sheet.columns = [
    { header: "GRN No", key: "grnNumber", width: 16 },
    { header: "GRN Date", key: "grnDate", width: 14 },
    { header: "Invoice No", key: "vendorInvoiceNo", width: 18 },
    { header: "PO No", key: "poNumber", width: 16 },
    { header: "Article", key: "article", width: 16 },
    { header: "Description", key: "description", width: 38 },
    { header: "EAN", key: "ean", width: 18 },
    { header: "Vendor Article", key: "vendorArticleNo", width: 18 },
    { header: "UOM", key: "uom", width: 8 },
    { header: "Challan Qty", key: "challanQty", width: 12 },
    { header: "Received Qty", key: "receivedQty", width: 12 },
    { header: "Accepted Qty", key: "acceptedQty", width: 12 },
    { header: "Short Qty", key: "shortQty", width: 10 },
    { header: "Reason", key: "reason", width: 16 }
  ];
  styleWorksheetHeader(sheet);
  grn.items.forEach((item) => {
    sheet.addRow({
      grnNumber: grn.grnNumber,
      grnDate: grn.grnDate,
      vendorInvoiceNo: grn.vendorInvoiceNo,
      poNumber: grn.poNumber,
      article: item.article,
      description: item.description,
      ean: item.ean,
      vendorArticleNo: item.vendorArticleNo,
      uom: item.uom,
      challanQty: item.challanQty,
      receivedQty: item.receivedQty,
      acceptedQty: item.acceptedQty,
      shortQty: item.shortQty,
      reason: item.reason
    });
  });
  if ((grn.comparisonRows || []).length) {
    const compare = workbook.addWorksheet("Shortage Compare");
    compare.columns = [
      { header: "Match Key", key: "matchLabel", width: 36 },
      { header: "Packing Qty", key: "packingQty", width: 14 },
      { header: "GRN Challan Qty", key: "challanQty", width: 14 },
      { header: "GRN Accepted Qty", key: "acceptedQty", width: 14 },
      { header: "Shortage Qty", key: "shortageQty", width: 12 }
    ];
    styleWorksheetHeader(compare);
    grn.comparisonRows.forEach((row) => compare.addRow(row));
  }
  await downloadWorkbook(`grn-detail-sheet-${clean(grn.grnNumber) || todayIso()}.xlsx`, workbook);
}

async function downloadAllGrnWorkbook() {
  const reports = (state.grnReports || []).length
    ? state.grnReports
    : ((state.grnStatus?.items || []).length ? [state.grnStatus] : []);
  if (!reports.length) {
    alert("No saved GRN reports available to export.");
    return;
  }
  const workbook = createWorkbookOrAlert();
  if (!workbook) return;

  const summary = workbook.addWorksheet("GRN Summary");
  summary.columns = [
    { header: "GRN No", key: "grnNumber", width: 16 },
    { header: "GRN Date", key: "grnDate", width: 14 },
    { header: "Invoice No", key: "vendorInvoiceNo", width: 18 },
    { header: "PO No", key: "poNumber", width: 16 },
    { header: "Supplier", key: "supplier", width: 28 },
    { header: "Challan Qty", key: "totalChallanQty", width: 12 },
    { header: "Received Qty", key: "totalReceivedQty", width: 12 },
    { header: "Accepted Qty", key: "totalAcceptedQty", width: 12 },
    { header: "Short Qty", key: "totalShortQty", width: 10 },
    { header: "Packing List", key: "packingListFileName", width: 24 },
    { header: "Updated", key: "updatedAt", width: 22 }
  ];
  styleWorksheetHeader(summary);
  reports.forEach((report) => summary.addRow({
    grnNumber: report.grnNumber,
    grnDate: report.grnDate,
    vendorInvoiceNo: report.vendorInvoiceNo,
    poNumber: report.poNumber,
    supplier: report.supplier,
    totalChallanQty: report.totalChallanQty,
    totalReceivedQty: report.totalReceivedQty,
    totalAcceptedQty: report.totalAcceptedQty,
    totalShortQty: report.totalShortQty,
    packingListFileName: report.packingListFileName,
    updatedAt: formatDateTimeDisplay(report.updatedAt)
  }));

  const detail = workbook.addWorksheet("GRN Details");
  detail.columns = [
    { header: "GRN No", key: "grnNumber", width: 16 },
    { header: "Invoice No", key: "vendorInvoiceNo", width: 18 },
    { header: "Article", key: "article", width: 16 },
    { header: "Description", key: "description", width: 38 },
    { header: "EAN", key: "ean", width: 18 },
    { header: "Vendor Article", key: "vendorArticleNo", width: 18 },
    { header: "UOM", key: "uom", width: 8 },
    { header: "Challan Qty", key: "challanQty", width: 12 },
    { header: "Received Qty", key: "receivedQty", width: 12 },
    { header: "Accepted Qty", key: "acceptedQty", width: 12 },
    { header: "Short Qty", key: "shortQty", width: 10 },
    { header: "Reason", key: "reason", width: 16 }
  ];
  styleWorksheetHeader(detail);
  reports.forEach((report) => {
    (report.items || []).forEach((item) => detail.addRow({
      grnNumber: report.grnNumber,
      vendorInvoiceNo: report.vendorInvoiceNo,
      article: item.article,
      description: item.description,
      ean: item.ean,
      vendorArticleNo: item.vendorArticleNo,
      uom: item.uom,
      challanQty: item.challanQty,
      receivedQty: item.receivedQty,
      acceptedQty: item.acceptedQty,
      shortQty: item.shortQty,
      reason: item.reason
    }));
  });

  const compareRows = reports.flatMap((report) => (report.comparisonRows || []).map((row) => ({
    grnNumber: report.grnNumber,
    invoiceNo: report.vendorInvoiceNo,
    matchLabel: row.matchLabel,
    packingQty: row.packingQty,
    challanQty: row.challanQty,
    acceptedQty: row.acceptedQty,
    shortageQty: row.shortageQty
  })));
  if (compareRows.length) {
    const compare = workbook.addWorksheet("Shortage Compare");
    compare.columns = [
      { header: "GRN No", key: "grnNumber", width: 16 },
      { header: "Invoice No", key: "invoiceNo", width: 18 },
      { header: "Match Key", key: "matchLabel", width: 36 },
      { header: "Packing Qty", key: "packingQty", width: 14 },
      { header: "GRN Challan Qty", key: "challanQty", width: 14 },
      { header: "GRN Accepted Qty", key: "acceptedQty", width: 14 },
      { header: "Shortage Qty", key: "shortageQty", width: 12 }
    ];
    styleWorksheetHeader(compare);
    compareRows.forEach((row) => compare.addRow(row));
  }

  await downloadWorkbook(`all-grn-reports-${todayIso()}.xlsx`, workbook);
}

async function seedWashcareFromReportFile(e) {
  const file = e.target.files?.[0];
  if (!file || !els.washcareForm) return;
  if (els.washcareSeedStatus) {
    els.washcareSeedStatus.textContent = `Reading test report: ${file.name}`;
  }
  try {
    const extracted = await extractWashcareReportDetails(file);
    els.washcareForm.reportSourceName.value = file.name;
    if (extracted.reportNumber && !clean(els.washcareForm.reportNumber.value)) els.washcareForm.reportNumber.value = extracted.reportNumber;
    if (extracted.reportDate && !clean(els.washcareForm.reportDate.value)) els.washcareForm.reportDate.value = extracted.reportDate;
    if (extracted.labName && !clean(els.washcareForm.labName.value)) els.washcareForm.labName.value = extracted.labName;
    if (extracted.reportStyleCode) els.washcareForm.reportStyleCode.value = extracted.reportStyleCode;
    if (extracted.reportColor) els.washcareForm.reportColor.value = extracted.reportColor;
    if (extracted.reportBrand) els.washcareForm.reportBrand.value = extracted.reportBrand;
    if (extracted.reportSupplier) els.washcareForm.reportSupplier.value = extracted.reportSupplier;
    if (extracted.composition) els.washcareForm.composition.value = extracted.composition;
    if (extracted.washcareText) els.washcareForm.washcareText.value = extracted.washcareText;
    if (extracted.careSymbols.length) els.washcareForm.careSymbols.value = extracted.careSymbols.join(", ");
    if (!clean(els.washcareForm.templatePath.value)) els.washcareForm.templatePath.value = DEFAULT_WASHCARE_TEMPLATE_PATH;
    renderWashcarePreview();
    if (els.washcareSeedStatus) {
      els.washcareSeedStatus.textContent = `Seeded washcare from ${file.name}. Please confirm the style mapping and print command.`;
    }
  } catch (error) {
    console.error(error);
    if (els.washcareSeedStatus) {
      els.washcareSeedStatus.textContent = "Could not read this test report automatically in local file mode. Try the launcher file or fill the washcare manually.";
    }
    alert("Could not read the selected test report PDF. Please try opening the app with 'Launch Piece Rate Calculator.cmd' and upload the PDF again.");
  } finally {
    if (els.washcareReportInput) els.washcareReportInput.value = "";
  }
}

async function extractWashcareReportDetails(file) {
  if (/^https?:\/\/(127\.0\.0\.1|localhost):3100/i.test(clean(window.location?.origin))) {
    try {
      return await extractWashcareReportDetailsViaServer(file);
    } catch (error) {
      console.error("Server-side PDF extraction failed, falling back to browser parser.", error);
    }
  }
  const { normalized } = await extractPdfPlainText(file);
  const submittedWashcare = extractReportField(normalized, /Submitted Wash Care\s*:?\s*(.+?)\s*Test Name/i);
  const composition = extractReportField(normalized, /Fiber Content\s*:?\s*(.+?)\s*Fabric Code/i);
  const reportNumber = extractReportField(normalized, /\b(RRLD\d{8,})\b/i);
  const reportDate = normalizeDateValue(extractReportField(normalized, /Received date\s+(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})/i) || extractReportField(normalized, /Report No\.\s*:?\s*(\d{1,2}\s+[A-Za-z]{3}\s+\d{4})/i));
  const labName = normalized.includes("Reliance Trends PRODUCT TESTING LABORATORY".replace(/\s+/g, " "))
    ? "Reliance Trends Product Testing Laboratory"
    : "Testing Laboratory";
  return {
    reportNumber,
    reportDate,
    labName,
    reportStyleCode: extractReportField(normalized, /Style Code\s*:?\s*(.+?)\s*No\.?\s*of Sample/i),
    reportColor: extractReportField(normalized, /Color\s*:?\s*(.+?)\s*Style Code/i),
    reportBrand: extractReportField(normalized, /Brand\s*:?\s*(.+?)\s*Supplier\s*:/i),
    reportSupplier: extractReportField(normalized, /Supplier\s*:?\s*(.+?)\s*Trf ID/i) || extractReportField(normalized, /Mill Supplier\s*:?\s*(.+?)\s*Type of Testing/i),
    composition,
    washcareText: submittedWashcare,
    careSymbols: parseWashcareSymbols((submittedWashcare || "").replace(/,\s*/g, ", "))
  };
}

async function extractWashcareReportDetailsViaServer(file) {
  const response = await fetch("/extract-washcare-report", {
    method: "POST",
    headers: {
      "Content-Type": "application/pdf"
    },
    body: file
  });
  const data = await response.json().catch(() => null);
  if (!response.ok || !data) {
    throw new Error(data?.details || data?.error || `Server extraction failed with HTTP ${response.status}.`);
  }
  return {
    reportNumber: clean(data.reportNumber),
    reportDate: normalizeDateValue(data.reportDate),
    labName: clean(data.labName),
    reportStyleCode: clean(data.reportStyleCode),
    reportColor: clean(data.reportColor),
    reportBrand: clean(data.reportBrand),
    reportSupplier: clean(data.reportSupplier),
    composition: clean(data.composition),
    washcareText: clean(data.washcareText),
    careSymbols: parseWashcareSymbols(clean(data.washcareText).replace(/,\s*/g, ", "))
  };
}

function extractReportField(text, pattern) {
  const match = text.match(pattern);
  return clean(match?.[1] || "");
}

function generateWashcarePrintCommand({ templatePath } = {}) {
  const resolvedTemplatePath = clean(templatePath) || DEFAULT_WASHCARE_TEMPLATE_PATH;
  if (!resolvedTemplatePath) return "";
  const escapedTemplatePath = resolvedTemplatePath.replace(/"/g, '""');
  return `bartend.exe /F="${escapedTemplatePath}" /P /X`;
}

async function copyWashcareCommand(recordOrEntryId = null) {
  const preview = typeof recordOrEntryId === "string"
    ? getWashcarePreviewData(state.washcareRecords.find((item) => item.id === recordOrEntryId) || null)
    : getWashcarePreviewData();
  if (!preview.style) {
    alert("Please select a style first.");
    return;
  }
  const command = clean(preview.printCommand);
  if (!command) {
    alert("No .btw template path is available yet. Please set the Template Path first.");
    return;
  }
  if (els.washcareForm?.printCommand && typeof recordOrEntryId !== "string" && !clean(els.washcareForm.printCommand.value)) {
    els.washcareForm.printCommand.value = command;
  }
  try {
    await navigator.clipboard.writeText(command);
    if (els.washcareSeedStatus) {
      els.washcareSeedStatus.textContent = `Print command copied to clipboard: ${command}`;
    }
  } catch {
    window.prompt("Copy the print command below:", command);
  }
}

async function handleWashcarePrint(recordOrEntryId = null) {
  const record = typeof recordOrEntryId === "string"
    ? state.washcareRecords.find((item) => item.id === recordOrEntryId) || null
    : null;
  const preview = getWashcarePreviewData(record);
  if (!preview.style) {
    alert("Please select a style first.");
    return;
  }
  const method = preview.printMethod || "browser";
  if (method === "command") {
    await copyWashcareCommand(recordOrEntryId);
    return;
  }
  const printWindow = window.open("", "_blank", "width=900,height=700");
  if (!printWindow) {
    alert("Pop-up blocked. Please allow pop-ups and try again.");
    return;
  }
  printWindow.document.write(`<!DOCTYPE html>
  <html lang="en">
  <head>
    <meta charset="UTF-8">
    <title>Washcare Print</title>
    <style>
      @page { size: ${num(preview.labelWidthMm) || DEFAULT_WASHCARE_LABEL_WIDTH_MM}mm ${num(preview.labelLengthMm) || DEFAULT_WASHCARE_LABEL_LENGTH_MM}mm; margin: 0; }
      html, body { margin: 0; padding: 0; background: #fff; color: #000; font-family: Arial, Helvetica, sans-serif; }
      body { display: flex; justify-content: center; }
      .sheet { width: ${num(preview.labelWidthMm) || DEFAULT_WASHCARE_LABEL_WIDTH_MM}mm; min-height: ${num(preview.labelLengthMm) || DEFAULT_WASHCARE_LABEL_LENGTH_MM}mm; box-sizing: border-box; display: flex; }
      .washcare-label { width: 100%; min-height: 100%; box-sizing: border-box; padding: 6mm 3mm; display: flex; flex-direction: column; align-items: center; justify-content: space-between; gap: 5mm; text-align: center; }
      .washcare-label-composition { font-size: var(--washcare-composition-size, 22px); font-weight: 700; line-height: 1.2; }
      .washcare-label-icons { min-height: 10mm; display: flex; align-items: center; justify-content: center; gap: var(--washcare-symbol-gap, 6px); flex-wrap: nowrap; }
      .washcare-icon { width: var(--washcare-symbol-size, 30px); height: calc(var(--washcare-symbol-size, 30px) * 0.8); color: #000; }
      .washcare-symbol-text { font-size: 2.4mm; font-weight: 700; }
      .washcare-label-text { font-size: var(--washcare-text-size, 16px); font-weight: 700; line-height: 1.35; letter-spacing: .02em; }
      .washcare-label-footer { margin-top: auto; font-size: 4mm; font-weight: 700; line-height: 1.35; }
    </style>
  </head>
  <body>
    <div class="sheet">${buildWashcarePreviewMarkup(preview)}</div>
  </body>
  </html>`);
  printWindow.document.close();
  printWindow.focus();
  printWindow.print();
  if (method === "both") {
    await copyWashcareCommand(recordOrEntryId);
  }
}

function editCutting(entryId) {
  const entry = state.cuttingEntries.find((item) => item.id === entryId);
  if (!entry) return;
  els.cuttingForm.dataset.editId = entry.id;
  els.cuttingForm.date.value = entry.date || "";
  els.cuttingForm.styleId.value = entry.styleId || "";
  els.cuttingForm.service.value = entry.service || "";
  els.cuttingForm.remarks.value = entry.remarks || "";
  fillSizeQuantities(els.cuttingSizeRows, "cut_", entry.quantities);
  els.cuttingForm.scrollIntoView({ behavior: "smooth", block: "start" });
}

function editStyleProduction(entryId) {
  const entry = state.styleProductionEntries.find((item) => item.id === entryId);
  if (!entry) return;
  els.styleProductionForm.dataset.editId = entry.id;
  els.styleProductionForm.date.value = entry.date || "";
  els.styleProductionForm.styleId.value = entry.styleId || "";
  els.styleProductionForm.totalQty.value = entry.totalQty || "";
  els.styleProductionForm.remarks.value = entry.remarks || "";
  fillSizeQuantities(els.styleProductionSizeRows, "styleProd_", entry.quantities);
  els.styleProductionForm.scrollIntoView({ behavior: "smooth", block: "start" });
}

function editProduction(entryId) {
  const entry = state.productionEntries.find((item) => item.id === entryId);
  if (!entry) return;
  els.productionForm.dataset.editId = entry.id;
  els.productionForm.date.value = entry.date || "";
  els.productionForm.styleId.value = entry.styleId || "";
  renderOperationSelect();
  els.productionForm.operationName.value = entry.operationName || "";
  els.productionForm.workerName.value = entry.workerName || "";
  els.productionForm.size.value = entry.size || "";
  els.productionForm.quantity.value = entry.quantity || "";
  els.productionForm.remarks.value = entry.remarks || "";
  els.productionForm.scrollIntoView({ behavior: "smooth", block: "start" });
}
function editAcceptance(entryId) {
  const entry = state.acceptanceEntries.find((item) => item.id === entryId);
  if (!entry) return;
  els.acceptanceForm.dataset.editId = entry.id;
  els.acceptanceForm.date.value = entry.date || "";
  els.acceptanceForm.styleId.value = entry.styleId || "";
  els.acceptanceForm.remarks.value = entry.remarks || "";
  entry.items.forEach((item) => {
    const acceptedInput = els.acceptanceSizeRows.querySelector(`[name="accepted_${item.size}"]`);
    const rejectedInput = els.acceptanceSizeRows.querySelector(`[name="rejected_${item.size}"]`);
    if (acceptedInput) acceptedInput.value = item.accepted || "";
    if (rejectedInput) rejectedInput.value = item.rejected || "";
  });
  els.acceptanceForm.scrollIntoView({ behavior: "smooth", block: "start" });
}

function editDispatch(entryId) {
  const entry = state.dispatchEntries.find((item) => item.id === entryId);
  if (!entry) return;
  els.dispatchForm.dataset.editId = entry.id;
  els.dispatchForm.date.value = entry.date || "";
  els.dispatchForm.styleId.value = entry.styleId || "";
  els.dispatchForm.remarks.value = entry.remarks || "";
  fillSizeQuantities(els.dispatchSizeRows, "dispatch_", entry.quantities);
  els.dispatchForm.scrollIntoView({ behavior: "smooth", block: "start" });
}

function fillSizeQuantities(container, prefix, quantities = {}) {
  getSizes().forEach((size) => {
    const input = container.querySelector(`[name="${prefix}${size}"]`);
    if (input) input.value = quantities[size] || "";
  });
}

function upsertEntry(collectionName, payload, entryId) {
  const index = state[collectionName].findIndex((item) => item.id === entryId);
  if (index >= 0) state[collectionName][index] = payload;
  else state[collectionName].push(payload);
}

function styleLabel(style) {
  if (!style) return "-";
  return style.color ? `${style.styleNumber} - ${style.color}` : style.styleNumber;
}

function filterStyles(searchTerm = "") {
  const query = normalizeStyleSearch(searchTerm);
  if (!query) {
    return state.styles.slice().sort((a, b) => styleLabel(a).localeCompare(styleLabel(b)));
  }
  return state.styles
    .filter((style) => styleMatchesSearch(style, query))
    .sort((a, b) => styleLabel(a).localeCompare(styleLabel(b)));
}

function styleMatchesSearch(style, query) {
  const styleNumber = normalizeStyleSearch(style?.styleNumber || "");
  if (styleNumber.includes(query)) return true;
  const styleDigits = digitsOnly(styleNumber);
  const queryDigits = digitsOnly(query);
  return queryDigits.length >= 2 && styleDigits.endsWith(queryDigits);
}

function normalizeStyleSearch(value) {
  return clean(value).toLowerCase().replace(/\s+/g, "");
}

function digitsOnly(value) {
  return String(value || "").replace(/\D+/g, "");
}

function formatQuantities(quantities = {}) {
  return Object.entries(quantities).filter(([, qty]) => num(qty) > 0).map(([size, qty]) => `${size}: ${fmtInt(qty)}`).join(", ") || "-";
}

function formatProductionQuantities(entry = {}) {
  const sizeWise = formatQuantities(entry.quantities);
  if (sizeWise !== "-") return sizeWise;
  return num(entry.totalQty) > 0 ? `Total only: ${fmtInt(entry.totalQty)}` : "-";
}

function entryProducedQty(entry = {}) {
  const sizeWiseTotal = sumObj(entry.quantities);
  return sizeWiseTotal > 0 ? sizeWiseTotal : num(entry.totalQty);
}

function formatAcceptanceItems(items = []) {
  return items.filter((item) => num(item.accepted) > 0 || num(item.rejected) > 0).map((item) => `${item.size} A:${fmtInt(item.accepted)} R:${fmtInt(item.rejected)}`).join(", ") || "-";
}

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(new Error("Could not read file"));
    reader.readAsDataURL(file);
  });
}

function handleStyleImagePaste(e) {
  const clipboardItems = [...(e.clipboardData?.items || [])];
  const imageItem = clipboardItems.find((item) => item.type.startsWith("image/"));
  if (!imageItem) return;
  e.preventDefault();
  const file = imageItem.getAsFile();
  if (!file) return;
  void storePastedStyleImage(file);
}

async function pasteStyleImageFromClipboard() {
  if (!navigator.clipboard?.read) {
    setStyleImagePasteStatus("Clipboard paste button is not supported in this browser. Use Ctrl + V after copying the image.");
    return;
  }
  try {
    const clipboardItems = await navigator.clipboard.read();
    for (const item of clipboardItems) {
      const imageType = item.types.find((type) => type.startsWith("image/"));
      if (!imageType) continue;
      const blob = await item.getType(imageType);
      await storePastedStyleImage(blob);
      return;
    }
    setStyleImagePasteStatus("No image found in the clipboard.");
  } catch (error) {
    console.error(error);
    setStyleImagePasteStatus("Clipboard access was blocked. Try the button again or use Ctrl + V.");
  }
}

async function storePastedStyleImage(fileOrBlob) {
  pastedStyleImageDataUrl = await blobToDataUrl(fileOrBlob);
  const imageInput = els.styleForm?.querySelector('[name="imageFile"]');
  if (imageInput) imageInput.value = "";
  if (els.styleForm?.image) els.styleForm.image.value = "";
  setStyleImagePasteStatus("Clipboard image added. Save Style to keep it.");
  updateStyleFormImagePreview();
}

function clearPastedStyleImage() {
  pastedStyleImageDataUrl = "";
  setStyleImagePasteStatus("You can also click here after copying a screenshot and press Ctrl + V.");
}

function setStyleImagePasteStatus(message) {
  if (els.styleImagePasteStatus) els.styleImagePasteStatus.textContent = message;
}

async function handleStyleImageFileChange() {
  const file = els.styleImageFile?.files?.[0];
  if (!file) {
    updateStyleFormImagePreview();
    return;
  }
  pastedStyleImageDataUrl = await fileToDataUrl(file);
  if (els.styleForm?.image) els.styleForm.image.value = "";
  setStyleImagePasteStatus("Uploaded image selected. Save Style to keep it.");
  updateStyleFormImagePreview();
}

function updateStyleFormImagePreview(explicitSource = "") {
  if (!els.styleFormImagePreview || !els.styleFormImagePreviewImg) return;
  const source = explicitSource || pastedStyleImageDataUrl || clean(els.styleForm?.image?.value || "");
  if (!source) {
    els.styleFormImagePreview.hidden = true;
    els.styleFormImagePreviewImg.removeAttribute("src");
    return;
  }
  els.styleFormImagePreview.hidden = false;
  els.styleFormImagePreviewImg.src = source;
}

async function buildImportedImageMap(fileList) {
  const files = [...(fileList || [])];
  const entries = await Promise.all(files.map(async (file) => {
    const dataUrl = await fileToDataUrl(file);
    return [normalizeImageLookupKey(file.name), dataUrl];
  }));
  return new Map(entries);
}

function resolveImportedImage(imageValue, importedImageMap) {
  const rawValue = clean(imageValue);
  if (!rawValue) return "";
  if (isDirectImageSource(rawValue)) return rawValue;
  const lookupKeys = [
    normalizeImageLookupKey(rawValue),
    normalizeImageLookupKey(fileNameFromPath(rawValue))
  ].filter(Boolean);
  for (const key of lookupKeys) {
    if (importedImageMap.has(key)) return importedImageMap.get(key);
  }
  return "";
}

function isDirectImageSource(value) {
  return /^data:image\//i.test(value) || /^(https?:)?\/\//i.test(value) || value.startsWith("./") || value.startsWith("../") || value.startsWith("/");
}

function fileNameFromPath(value) {
  return clean(value).split(/[\\/]/).pop() || "";
}

function normalizeImageLookupKey(value) {
  return clean(value).toLowerCase();
}

async function saveSizes(e) {
  e.preventDefault();
  const sizes = parseSizeList(els.sizeListInput.value);
  if (!sizes.length) {
    alert("Please enter at least one size.");
    return;
  }
  state.settings = state.settings || {};
  state.settings.sizes = sizes;
  buildSizeInputs();
  buildSizeSelect();
  await persistState();
  alert("Sizes updated successfully.");
}

async function saveAccessCodes(e) {
  e.preventDefault();
  state.settings = state.settings || {};
  state.settings.accessCodes = state.settings.accessCodes || {};
  SECTION_ACCESS_TABS.forEach((section) => {
    state.settings.accessCodes[section] = clean(els.accessCodeForm?.elements?.[section]?.value || "");
  });
  await persistState();
  setAccessCodeStatus("Access codes saved.");
}

async function handleAccessLinkCopy(e) {
  const button = e.target.closest("[data-copy-access-link]");
  if (!button) return;
  const section = button.dataset.copyAccessLink;
  const code = getSectionAccessCode(section);
  if (!code) {
    setAccessCodeStatus(`Set and save a code for ${SECTION_ACCESS_LABELS[section]} first.`);
    return;
  }
  const shareUrl = buildSectionAccessUrl(section, code);
  try {
    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(shareUrl);
      setAccessCodeStatus(`${SECTION_ACCESS_LABELS[section]} link copied.`);
    } else {
      setAccessCodeStatus(shareUrl);
    }
  } catch {
    setAccessCodeStatus(shareUrl);
  }
}

function populateAccessCodeInputs() {
  if (!els.accessCodeForm) return;
  SECTION_ACCESS_TABS.forEach((section) => {
    const input = els.accessCodeForm.elements?.[section];
    if (input) input.value = getSectionAccessCode(section);
  });
}

function setAccessCodeStatus(message) {
  if (els.accessCodeStatus) els.accessCodeStatus.textContent = message;
}

function getSectionAccessCode(section) {
  return clean(state.settings?.accessCodes?.[section] || "");
}

function buildSectionAccessUrl(section, code) {
  const url = new URL(window.location.href);
  url.searchParams.set("section", section);
  url.searchParams.set("code", code);
  return url.toString();
}

function isSharedSectionTabAllowed(tabId) {
  return !activeSharedSection || tabId === "dashboard" || tabId === "reports" || tabId === activeSharedSection;
}

function applySharedSectionAccess() {
  const url = new URL(window.location.href);
  const requestedSection = clean(url.searchParams.get("section"));
  const requestedCode = clean(url.searchParams.get("code"));
  activeSharedSection = "";
  if (!requestedSection || !requestedCode || !SECTION_ACCESS_TABS.includes(requestedSection)) {
    clearSharedSectionAccess();
    return;
  }
  if (requestedCode !== getSectionAccessCode(requestedSection)) {
    clearSharedSectionAccess();
    setAccessCodeStatus(`Shared link code for ${SECTION_ACCESS_LABELS[requestedSection]} is invalid.`);
    alert("This shared section link is invalid or the access code does not match.");
    return;
  }
  activeSharedSection = requestedSection;
  if (els.sidebar) els.sidebar.classList.toggle("shared-mode", true);
  els.sidebarCards.forEach((card) => {
    card.hidden = true;
  });
  els.navLinks.forEach((btn) => {
    const allowed = isSharedSectionTabAllowed(btn.dataset.tab);
    btn.hidden = !allowed;
    btn.style.display = allowed ? "" : "none";
  });
  els.panels.forEach((panel) => {
    const allowed = isSharedSectionTabAllowed(panel.id);
    panel.hidden = !allowed;
    panel.style.display = allowed ? "" : "none";
  });
  activateTab(requestedSection);
}

function clearSharedSectionAccess() {
  if (els.sidebar) els.sidebar.classList.toggle("shared-mode", false);
  els.sidebarCards.forEach((card) => {
    card.hidden = false;
  });
  els.navLinks.forEach((btn) => {
    btn.hidden = false;
    btn.style.display = "";
  });
  els.panels.forEach((panel) => {
    panel.hidden = false;
    panel.style.display = "";
  });
  const activeButton = [...els.navLinks].find((btn) => btn.classList.contains("active")) || els.navLinks[0];
  if (activeButton) activateTab(activeButton.dataset.tab);
}

function getSizes() {
  return (state.settings?.sizes?.length ? state.settings.sizes : DEFAULT_SIZE_LIST).map((size) => clean(size)).filter(Boolean);
}

function getSizeQuantities(quantities = {}) {
  return Object.fromEntries(getSizes().map((size) => [size, num(quantities?.[size])]));
}

function parseSizeList(value) {
  return [...new Set(String(value || "").split(",").map((item) => clean(item)).filter(Boolean))];
}

function normalizeState() {
  state = { ...clone(DEFAULTS), ...state };
  state.styleImages = state.styleImages && typeof state.styleImages === "object" ? { ...state.styleImages } : {};
  state.settings = state.settings || {};
  if (!Array.isArray(state.settings.sizes) || !state.settings.sizes.length) {
    state.settings.sizes = [...DEFAULT_SIZE_LIST];
  }
  state.settings.accessCodes = { ...clone(DEFAULTS.settings.accessCodes), ...(state.settings.accessCodes || {}) };
  state.styles = (state.styles || []).map((style) => ({
    ...style,
    orderQty: num(style.orderQty),
    cmtRate: num(style.cmtRate),
    serviceChargePct: num(style.serviceChargePct),
    operations: Array.isArray(style.operations) ? style.operations : []
  }));
  state.cuttingEntries = (Array.isArray(state.cuttingEntries) ? state.cuttingEntries : []).map((entry) => ({
    ...entry,
    date: normalizeDateValue(entry.date),
    service: clean(entry.service)
  }));
  state.styleProductionEntries = (Array.isArray(state.styleProductionEntries) ? state.styleProductionEntries : []).map((entry) => ({
    ...entry,
    date: normalizeDateValue(entry.date),
    totalQty: num(entry.totalQty)
  }));
  state.productionEntries = (Array.isArray(state.productionEntries) ? state.productionEntries : []).map((entry) => ({
    ...entry,
    date: normalizeDateValue(entry.date)
  }));
  state.acceptanceEntries = (Array.isArray(state.acceptanceEntries) ? state.acceptanceEntries : []).map((entry) => ({
    ...entry,
    date: normalizeDateValue(entry.date)
  }));
  state.dispatchEntries = (Array.isArray(state.dispatchEntries) ? state.dispatchEntries : []).map((entry) => ({
    ...entry,
    date: normalizeDateValue(entry.date)
  }));
  state.grnStatus = { ...clone(DEFAULTS.grnStatus), ...(state.grnStatus || {}) };
  state.grnReports = (Array.isArray(state.grnReports) ? state.grnReports : []).map((report) => ({
    ...clone(DEFAULTS.grnStatus),
    ...report
  }));
  state.grnStatus.grnDate = normalizeDateValue(state.grnStatus.grnDate);
  state.grnStatus.poDate = normalizeDateValue(state.grnStatus.poDate);
  state.grnStatus.consignmentDate = normalizeDateValue(state.grnStatus.consignmentDate);
  state.grnStatus.items = Array.isArray(state.grnStatus.items) ? state.grnStatus.items.map((item) => ({
    ...item,
    serialNo: num(item.serialNo),
    challanQty: num(item.challanQty),
    receivedQty: num(item.receivedQty),
    acceptedQty: num(item.acceptedQty),
    shortQty: num(item.shortQty)
  })) : [];
  state.grnStatus.packingItems = Array.isArray(state.grnStatus.packingItems) ? state.grnStatus.packingItems.map((item) => ({
    ...item,
    qty: num(item.qty)
  })) : [];
  state.grnStatus.comparisonRows = Array.isArray(state.grnStatus.comparisonRows) ? state.grnStatus.comparisonRows.map((row) => ({
    ...row,
    packingQty: num(row.packingQty),
    challanQty: num(row.challanQty),
    acceptedQty: num(row.acceptedQty),
    shortageQty: num(row.shortageQty)
  })) : [];
  state.grnStatus.totalChallanQty = num(state.grnStatus.totalChallanQty || state.grnStatus.items.reduce((sum, item) => sum + num(item.challanQty), 0));
  state.grnStatus.totalReceivedQty = num(state.grnStatus.totalReceivedQty || state.grnStatus.items.reduce((sum, item) => sum + num(item.receivedQty), 0));
  state.grnStatus.totalAcceptedQty = num(state.grnStatus.totalAcceptedQty || state.grnStatus.items.reduce((sum, item) => sum + num(item.acceptedQty), 0));
  state.grnStatus.totalShortQty = num(state.grnStatus.totalShortQty || state.grnStatus.items.reduce((sum, item) => sum + num(item.shortQty), 0));
  state.grnReports = state.grnReports.map((report) => ({
    ...report,
    grnDate: normalizeDateValue(report.grnDate),
    poDate: normalizeDateValue(report.poDate),
    consignmentDate: normalizeDateValue(report.consignmentDate),
    items: Array.isArray(report.items) ? report.items.map((item) => ({
      ...item,
      serialNo: num(item.serialNo),
      challanQty: num(item.challanQty),
      receivedQty: num(item.receivedQty),
      acceptedQty: num(item.acceptedQty),
      shortQty: num(item.shortQty)
    })) : [],
    packingItems: Array.isArray(report.packingItems) ? report.packingItems.map((item) => ({
      ...item,
      qty: num(item.qty)
    })) : [],
    comparisonRows: Array.isArray(report.comparisonRows) ? report.comparisonRows.map((row) => ({
      ...row,
      packingQty: num(row.packingQty),
      challanQty: num(row.challanQty),
      acceptedQty: num(row.acceptedQty),
      shortageQty: num(row.shortageQty)
    })) : [],
    totalChallanQty: num(report.totalChallanQty || (report.items || []).reduce((sum, item) => sum + num(item.challanQty), 0)),
    totalReceivedQty: num(report.totalReceivedQty || (report.items || []).reduce((sum, item) => sum + num(item.receivedQty), 0)),
    totalAcceptedQty: num(report.totalAcceptedQty || (report.items || []).reduce((sum, item) => sum + num(item.acceptedQty), 0)),
    totalShortQty: num(report.totalShortQty || (report.items || []).reduce((sum, item) => sum + num(item.shortQty), 0))
  }));
  state.washcareRecords = (Array.isArray(state.washcareRecords) ? state.washcareRecords : []).map((record) => ({
    ...record,
    reportDate: normalizeDateValue(record.reportDate),
    careSymbols: Array.isArray(record.careSymbols) ? record.careSymbols.map((item) => clean(item)).filter(Boolean) : parseWashcareSymbols(record.careSymbols),
    printMethod: clean(record.printMethod) || "browser",
    templatePath: clean(record.templatePath) || DEFAULT_WASHCARE_TEMPLATE_PATH,
    labelWidthMm: clampWashcareDimension(record.labelWidthMm, DEFAULT_WASHCARE_LABEL_WIDTH_MM),
    labelLengthMm: clampWashcareDimension(record.labelLengthMm, DEFAULT_WASHCARE_LABEL_LENGTH_MM),
    reportSourceName: clean(record.reportSourceName),
    reportStyleCode: clean(record.reportStyleCode),
    reportColor: clean(record.reportColor),
    reportBrand: clean(record.reportBrand),
    reportSupplier: clean(record.reportSupplier),
    compositionFontSize: clampWashcareSize(record.compositionFontSize, 22),
    symbolSize: clampWashcareSize(record.symbolSize, 30),
    symbolGap: clampWashcareSize(record.symbolGap, 6),
    washcareFontSize: clampWashcareSize(record.washcareFontSize, 16),
    originLine: clean(record.originLine) || "MADE IN INDIA",
    footerLine1: clean(record.footerLine1),
    footerLine2: clean(record.footerLine2),
    updatedAt: clean(record.updatedAt)
  }));
  state.payments = (Array.isArray(state.payments) ? state.payments : []).map((payment) => ({
    ...payment,
    paymentDate: normalizeDateValue(payment.paymentDate),
    startDate: normalizeDateValue(payment.startDate),
    endDate: normalizeDateValue(payment.endDate),
    styleIds: Array.isArray(payment.styleIds) ? payment.styleIds : [],
    baseAmountPaid: num(payment.baseAmountPaid),
    serviceChargePaid: num(payment.serviceChargePaid),
    totalPaid: num(payment.totalPaid)
  }));
  state.tallyCreditors = {
    ...clone(DEFAULTS.tallyCreditors),
    ...(state.tallyCreditors || {}),
    ledgers: Array.isArray(state.tallyCreditors?.ledgers) ? state.tallyCreditors.ledgers : [],
    vouchers: Array.isArray(state.tallyCreditors?.vouchers) ? state.tallyCreditors.vouchers : []
  };
}

function getReportDateFilter() {
  return normalizeReportRangeInput({
    startDate: clean(els.reportStartDate?.value),
    endDate: clean(els.reportEndDate?.value)
  });
}

function clearReportDateFilter() {
  if (els.reportStartDate) els.reportStartDate.value = "";
  if (els.reportEndDate) els.reportEndDate.value = "";
  renderReports();
}

function matchesDate(entryDate, reportDate) {
  const range = normalizeReportRangeInput(reportDate);
  const dateValue = normalizeDateValue(entryDate);
  if (!dateValue) return false;
  if (range.startDate && dateValue < range.startDate) return false;
  if (range.endDate && dateValue > range.endDate) return false;
  return true;
}

function downloadStyleAmountReport() {
  const reportDate = getReportDateFilter();
  const rows = styleBillingRows(reportDate);
  if (!rows.length) {
    alert("No style amount report found for the selected date range.");
    return;
  }
  const csv = [
    ["reportDates", "styleNumber", "color", "orderQty", "cutQty", "makeQty", "dispatchQty", "cutVsMake", "makeVsDispatch", "cmtRate", "amount", "serviceChargeAmount", "totalBillAmount", "status"].join(","),
    ...rows.map((row) => [
      csvValue(row.dateLabel),
      csvValue(row.styleNumber),
      csvValue(row.color || ""),
      row.orderQty,
      row.cutQty,
      row.producedQty,
      row.dispatchQty,
      csvValue(row.cutVsMakeSummary),
      csvValue(row.makeVsDispatchSummary),
      row.cmtRate,
      row.baseAmount,
      row.serviceChargeAmount,
      row.billing,
      csvValue(row.paymentStatusLabel)
    ].join(","))
  ].join("\n");
  downloadTextFile(`style-amount-report-${formatReportRangeForFilename(reportDate)}.csv`, csv, "text/csv");
}

async function downloadBillingPdf() {
  const reportRange = getReportDateFilter();
  const rows = styleBillingRows(reportRange);
  if (!rows.length) {
    alert("No billing report found for the selected date range.");
    return;
  }
  const jsPdfCtor = window.jspdf?.jsPDF;
  if (!jsPdfCtor) {
    alert("PDF export library could not load. Please refresh and try again.");
    return;
  }

  try {
    const pdf = new jsPdfCtor({
      orientation: "landscape",
      unit: "mm",
      format: "a4"
    });
    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();
    const left = 8;
    const top = 10;
    const columns = [
      { key: "styleNumber", label: "Style", width: 24 },
      { key: "color", label: "Color", width: 16 },
      { key: "image", label: "Image", width: 20 },
      { key: "orderQty", label: "Ord Qty", width: 13 },
      { key: "cutQty", label: "Cut Qty", width: 13 },
      { key: "producedQty", label: "Make Qty", width: 15 },
      { key: "dispatchQty", label: "Disp Qty", width: 15 },
      { key: "cutVsMakeSummary", label: "Cut/Make", width: 29 },
      { key: "makeVsDispatchSummary", label: "Make/Disp", width: 29 },
      { key: "cmtRate", label: "CMT", width: 12 },
      { key: "baseAmount", label: "Amount", width: 18 },
      { key: "serviceChargeAmount", label: "Service", width: 18 },
      { key: "billing", label: "Bill Amt", width: 18 },
      { key: "paymentStatusLabel", label: "Status", width: 15 }
    ];

    let y = top;
    const drawHeader = () => {
      pdf.setFont("helvetica", "bold");
      pdf.setFontSize(16);
      pdf.text("ENVOGUE CLOTHING", pageWidth / 2, y, { align: "center" });
      pdf.setFontSize(11);
      pdf.text("Billing Report", pageWidth / 2, y + 6, { align: "center" });
      pdf.setFont("helvetica", "normal");
      pdf.setFontSize(9);
      pdf.text(`Date of Reports: ${getReportRangeLabel(reportRange)}`, left, y + 12);
      y += 18;

      let currentX = left;
      const headerHeight = 12;
      pdf.setFillColor(245, 232, 214);
      pdf.setDrawColor(143, 59, 32);
      columns.forEach((column) => {
        pdf.rect(currentX, y, column.width, headerHeight, "FD");
        currentX += column.width;
      });
      currentX = left;
      pdf.setFont("helvetica", "bold");
      pdf.setFontSize(6);
      pdf.setTextColor(43, 33, 23);
      columns.forEach((column) => {
        pdf.text(String(column.label), currentX + (column.width / 2), y + 7, {
          align: "center",
          maxWidth: column.width - 2
        });
        currentX += column.width;
      });
      pdf.setTextColor(0, 0, 0);
      y += 12;
    };

    drawHeader();

    for (const row of rows) {
      const rowHeight = 16;
      if (y + rowHeight > pageHeight - 18) {
        pdf.addPage("a4", "landscape");
        y = top;
        drawHeader();
      }
      let currentX = left;
      pdf.setFont("helvetica", "normal");
      pdf.setFontSize(7);
      for (const column of columns) {
        pdf.rect(currentX, y, column.width, rowHeight);
        if (column.key === "image") {
          const imageData = await getPdfSafeImageData(row.image);
          if (imageData) {
            const fit = await calculateImageFit(column.width - 2, rowHeight - 2, imageData);
            pdf.addImage(imageData, "PNG", currentX + 1 + fit.xOffset, y + 1 + fit.yOffset, fit.width, fit.height);
          } else {
            pdf.text("-", currentX + (column.width / 2), y + 9, { align: "center" });
          }
        } else {
          const value = formatBillingPdfCell(column.key, row);
          pdf.text(pdf.splitTextToSize(value, column.width - 2), currentX + 1, y + 4);
        }
        currentX += column.width;
      }
      y += rowHeight;
    }

    const totals = rows.reduce((sum, row) => {
      sum.baseAmount += row.baseAmount;
      sum.serviceChargeAmount += row.serviceChargeAmount;
      sum.billing += row.billing;
      return sum;
    }, { baseAmount: 0, serviceChargeAmount: 0, billing: 0 });
    if (y + 16 > pageHeight - 10) {
      pdf.addPage("a4", "landscape");
      y = top;
      drawHeader();
    }
    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(9);
    pdf.text(`Total Amount: Rs ${fmt(totals.baseAmount)}`, left, y + 8);
    pdf.text(`Total Service Charge: Rs ${fmt(totals.serviceChargeAmount)}`, left + 82, y + 8);
    pdf.text(`Total Bill Amount: Rs ${fmt(totals.billing)}`, left + 180, y + 8);

    pdf.save(`billing-report-${formatReportRangeForFilename(reportRange)}.pdf`);
  } catch (error) {
    console.error("Billing PDF export failed:", error);
    alert(`Billing PDF could not be created.${clean(error?.message) ? `\n${clean(error.message)}` : ""}`);
  }
}

async function downloadCuttingReport() {
  const reportDate = getReportDateFilter();
  const rows = cuttingReportRows(reportDate);
  if (!rows.length) {
    alert("No cutting report found for the selected date range.");
    return;
  }
  const workbook = createWorkbookOrAlert();
  if (!workbook) return;
  const sheet = workbook.addWorksheet("Cutting Report");
  const sizeColumns = getSizes();
  sheet.columns = [
    { header: "Date", key: "date", width: 14 },
    { header: "Photo", key: "photo", width: 16 },
    { header: "Style", key: "styleNumber", width: 18 },
    { header: "Colour", key: "color", width: 16 },
    { header: "Service", key: "service", width: 18 },
    ...sizeColumns.map((size) => ({ header: size, key: `size_${size}`, width: 10 })),
    { header: "Total Qty", key: "totalQty", width: 12 },
    { header: "Remarks", key: "remarks", width: 24 }
  ];
  styleWorksheetHeader(sheet);
  for (const row of rows) {
    const excelRow = sheet.addRow({
      date: row.date,
      styleNumber: row.styleNumber,
      color: row.color || "",
      service: row.service || "",
      ...Object.fromEntries(sizeColumns.map((size) => [`size_${size}`, row.quantities[size] || 0])),
      totalQty: row.totalQty,
      remarks: row.remarks || ""
    });
    excelRow.height = 62;
    await addWorksheetImage(sheet, row.image, excelRow.number, 2);
  }
  finalizeWorksheet(sheet);
  await downloadWorkbook(`cutting-report-${formatReportRangeForFilename(reportDate)}.xlsx`, workbook);
}

async function downloadInternalChallan() {
  const reportDate = getReportDateFilter();
  const rows = cuttingReportRows(reportDate);
  if (!rows.length) {
    alert("No cutting entries found for internal challan.");
    return;
  }
  const pdf = createPdfDocumentOrAlert();
  if (!pdf) return;
  try {
    for (let index = 0; index < rows.length; index += 1) {
      const row = rows[index];
      if (index > 0) pdf.addPage([148, 210], "portrait");
      await buildInternalChallanPdfPage(pdf, row, index + 1);
    }
    pdf.save(`internal-challan-${formatReportRangeForFilename(reportDate)}.pdf`);
  } catch (error) {
    console.error("Internal challan PDF export failed:", error);
    alert(`Internal challan PDF could not be created.${clean(error?.message) ? `\n${clean(error.message)}` : ""}`);
  }
}

async function downloadFlowReport() {
  const reportDate = getReportDateFilter();
  const rows = dispatchReportRows(reportDate);
  if (!rows.length) {
    alert("No cut-make-dispatch report found for the selected date range.");
    return;
  }
  const workbook = createWorkbookOrAlert();
  if (!workbook) return;
  const sheet = workbook.addWorksheet("Cut Make Dispatch");
  sheet.columns = [
    { header: "Photo", key: "photo", width: 16 },
    { header: "Style", key: "styleNumber", width: 18 },
    { header: "Colour", key: "color", width: 16 },
    { header: "Size", key: "size", width: 10 },
    { header: "Cut Qty", key: "cutQty", width: 11 },
    { header: "Make Qty", key: "makeQty", width: 11 },
    { header: "Dispatch Qty", key: "dispatchQty", width: 13 },
    { header: "Amount", key: "amount", width: 14 },
    { header: "Balance", key: "balance", width: 11 }
  ];
  styleWorksheetHeader(sheet);
  for (const row of rows) {
    const excelRow = sheet.addRow({
      styleNumber: row.styleNumber,
      color: row.color || "",
      size: row.size,
      cutQty: row.cutQty,
      makeQty: row.makeQty,
      dispatchQty: row.dispatchQty,
      amount: row.amount,
      balance: row.balance
    });
    excelRow.height = 62;
    await addWorksheetImage(sheet, row.image, excelRow.number, 1);
  }
  finalizeWorksheet(sheet);
  await downloadWorkbook(`cut-make-dispatch-${formatReportRangeForFilename(reportDate)}.xlsx`, workbook);
}

function handleImagePreviewAction(e) {
  const trigger = e.target.closest("[data-action='preview-image']");
  if (!trigger) return;
  openImagePreview(trigger.dataset.imageSrc, trigger.dataset.imageTitle);
}

function openImagePreview(src, title = "Style Preview") {
  if (!src || !els.imagePreviewModal || !els.imagePreviewModalImg) return;
  els.imagePreviewModalImg.src = src;
  els.imagePreviewModalImg.alt = title;
  const titleEl = $("imagePreviewTitle");
  if (titleEl) titleEl.textContent = title;
  els.imagePreviewModal.hidden = false;
}

function closeImagePreview() {
  if (!els.imagePreviewModal || !els.imagePreviewModalImg) return;
  els.imagePreviewModal.hidden = true;
  els.imagePreviewModalImg.removeAttribute("src");
}

function createWorkbookOrAlert() {
  if (!window.ExcelJS?.Workbook) {
    alert("XLSX export library could not load. Please refresh and try again.");
    return null;
  }
  const workbook = new window.ExcelJS.Workbook();
  workbook.creator = "Piece Rate Calculator";
  workbook.created = new Date();
  return workbook;
}

function styleWorksheetHeader(sheet) {
  const headerRow = sheet.getRow(1);
  headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
  headerRow.alignment = { vertical: "middle", horizontal: "center" };
  headerRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF8F3B20" } };
  headerRow.height = 24;
}

function finalizeWorksheet(sheet) {
  sheet.views = [{ state: "frozen", ySplit: 1 }];
  sheet.eachRow((row) => {
    row.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
  });
}

function buildInternalChallanNumber(row, serialNumber) {
  const styleCode = clean(row.styleNumber || "STYLE").replace(/[^a-z0-9]/gi, "").toUpperCase().slice(0, 8) || "STYLE";
  const dateCode = clean(row.date || "").replace(/-/g, "") || "DATE";
  return `${dateCode}-${styleCode}-${String(serialNumber).padStart(2, "0")}`;
}

function createPdfDocumentOrAlert() {
  const jsPdfCtor = window.jspdf?.jsPDF;
  if (!jsPdfCtor) {
    alert("PDF export library could not load. Please refresh and try again.");
    return null;
  }
  return new jsPdfCtor({
    orientation: "portrait",
    unit: "mm",
    format: [148, 210]
  });
}

async function buildInternalChallanPdfPage(pdf, row, serialNumber) {
  const pageWidth = 148;
  const pageHeight = 210;
  const margin = 8;
  const contentWidth = pageWidth - (margin * 2);
  const leftColWidth = 88;
  const rightColX = margin + leftColWidth + 4;
  const challanNumber = buildInternalChallanNumber(row, serialNumber);
  const quantities = getNonZeroQuantityPairs(row.quantities);

  pdf.setDrawColor(185, 164, 140);
  pdf.setLineWidth(0.4);
  pdf.rect(4, 4, pageWidth - 8, pageHeight - 8);

  pdf.setFont("helvetica", "bold");
  pdf.setFontSize(14);
  pdf.text("ENVOGUE CLOTHING", pageWidth / 2, 14, { align: "center" });
  pdf.setFontSize(12);
  pdf.text("Internal Challan", pageWidth / 2, 21, { align: "center" });

  pdf.setFontSize(9);
  pdf.text(`Challan No: ${challanNumber}`, margin, 29);
  pdf.text(`Date: ${row.date || "-"}`, pageWidth - margin, 29, { align: "right" });

  drawPdfFieldBox(pdf, margin, 34, contentWidth, 10, "Issue To", "Production Department");
  drawPdfFieldBox(pdf, margin, 44, contentWidth, 10, "Purpose", "Stitching");
  drawPdfFieldBox(pdf, margin, 54, leftColWidth, 10, "Style", row.styleNumber || "-");
  drawPdfFieldBox(pdf, rightColX, 54, contentWidth - leftColWidth - 4, 10, "Colour", row.color || "-");
  drawPdfFieldBox(pdf, margin, 64, leftColWidth, 10, "Service", row.service || "-");
  drawPdfFieldBox(pdf, rightColX, 64, contentWidth - leftColWidth - 4, 10, "Total Qty", fmtInt(row.totalQty || 0));
  drawPdfFieldBox(pdf, margin, 74, contentWidth, 12, "Remarks", row.remarks || "-");

  const imageBottomY = await drawChallanImageBox(pdf, row.image, rightColX, 88, contentWidth - leftColWidth - 4, 36);

  const tableTop = 88;
  const tableWidth = leftColWidth;
  pdf.setFillColor(143, 59, 32);
  pdf.setTextColor(255, 255, 255);
  pdf.rect(margin, tableTop, tableWidth, 8, "F");
  pdf.setFont("helvetica", "bold");
  pdf.setFontSize(9);
  pdf.text("Size", margin + 10, tableTop + 5.4, { align: "center" });
  pdf.text("Quantity", margin + tableWidth - 14, tableTop + 5.4, { align: "center" });
  pdf.setTextColor(0, 0, 0);

  let currentY = tableTop + 8;
  if (quantities.length) {
    quantities.forEach(([size, qty]) => {
      pdf.rect(margin, currentY, tableWidth * 0.45, 7);
      pdf.rect(margin + (tableWidth * 0.45), currentY, tableWidth * 0.55, 7);
      pdf.setFont("helvetica", "normal");
      pdf.text(String(size), margin + (tableWidth * 0.225), currentY + 4.6, { align: "center" });
      pdf.text(fmtInt(qty), margin + (tableWidth * 0.45) + (tableWidth * 0.275), currentY + 4.6, { align: "center" });
      currentY += 7;
    });
  } else {
    pdf.rect(margin, currentY, tableWidth, 10);
    pdf.text("No size-wise details", margin + (tableWidth / 2), currentY + 6, { align: "center" });
    currentY += 10;
  }

  const signatureY = Math.max(currentY + 18, imageBottomY + 18, 175);
  pdf.line(margin, signatureY, margin + 46, signatureY);
  pdf.line(pageWidth - margin - 46, signatureY, pageWidth - margin, signatureY);
  pdf.setFont("helvetica", "normal");
  pdf.setFontSize(9);
  pdf.text("Cutting Incharge Signature", margin + 23, signatureY + 5, { align: "center" });
  pdf.text("Production Received By", pageWidth - margin - 23, signatureY + 5, { align: "center" });
}

function drawPdfFieldBox(pdf, x, y, width, height, label, value) {
  pdf.rect(x, y, width, height);
  pdf.setFont("helvetica", "bold");
  pdf.setFontSize(8);
  pdf.text(`${label}:`, x + 2, y + 3.8);
  pdf.setFont("helvetica", "normal");
  pdf.setFontSize(9);
  const text = pdf.splitTextToSize(String(value || "-"), width - 4);
  pdf.text(text, x + 2, y + 7.4);
}

async function drawChallanImageBox(pdf, source, x, y, width, height) {
  pdf.rect(x, y, width, height);
  pdf.setFont("helvetica", "bold");
  pdf.setFontSize(8);
  pdf.text("Image", x + 2, y + 4);

  const imageData = await getPdfSafeImageData(source);
  if (!imageData) {
    pdf.setFont("helvetica", "normal");
    pdf.setFontSize(9);
    pdf.text("No image", x + (width / 2), y + (height / 2) + 1, { align: "center" });
    return y + height;
  }

  const fit = await calculateImageFit(width - 4, height - 8, imageData);
  pdf.addImage(imageData, "PNG", x + 2 + fit.xOffset, y + 6 + fit.yOffset, fit.width, fit.height);
  return y + height;
}

async function getPdfSafeImageData(source) {
  const imageData = await imageSourceToBase64(source);
  if (!imageData) return "";
  return rasterizeImageDataForPdf(imageData);
}

function rasterizeImageDataForPdf(imageData) {
  return new Promise((resolve) => {
    const image = new Image();
    image.onload = () => {
      try {
        const canvas = document.createElement("canvas");
        canvas.width = image.naturalWidth || 1;
        canvas.height = image.naturalHeight || 1;
        const ctx = canvas.getContext("2d");
        if (!ctx) {
          resolve(/^data:image\/png/i.test(imageData) ? imageData : "");
          return;
        }
        ctx.drawImage(image, 0, 0);
        resolve(canvas.toDataURL("image/png"));
      } catch {
        resolve(/^data:image\/png/i.test(imageData) ? imageData : "");
      }
    };
    image.onerror = () => resolve(/^data:image\/png/i.test(imageData) ? imageData : "");
    image.src = imageData;
  });
}

async function calculateImageFit(maxWidth, maxHeight, imageData) {
  const { width: sourceWidth, height: sourceHeight } = await getImageDimensions(imageData);
  const ratio = Math.min(maxWidth / sourceWidth, maxHeight / sourceHeight);
  const width = Math.max(1, sourceWidth * ratio);
  const height = Math.max(1, sourceHeight * ratio);
  return {
    width,
    height,
    xOffset: (maxWidth - width) / 2,
    yOffset: (maxHeight - height) / 2
  };
}

function getImageDimensions(imageData) {
  return new Promise((resolve) => {
    const image = new Image();
    image.onload = () => resolve({
      width: image.naturalWidth || 1,
      height: image.naturalHeight || 1
    });
    image.onerror = () => resolve({ width: 1, height: 1 });
    image.src = imageData;
  });
}

function getNonZeroQuantityPairs(quantities = {}) {
  return Object.entries(quantities).filter(([, qty]) => num(qty) > 0);
}

async function addWorksheetImage(sheet, source, rowNumber, columnNumber) {
  const imageData = await imageSourceToBase64(source);
  if (!imageData) return;
  const extension = imageData.startsWith("data:image/png") ? "png" : "jpeg";
  const imageId = sheet.workbook.addImage({
    base64: imageData,
    extension
  });
  sheet.addImage(imageId, {
    tl: { col: columnNumber - 1 + 0.15, row: rowNumber - 1 + 0.15 },
    ext: { width: 52, height: 52 }
  });
}

async function imageSourceToBase64(source) {
  const value = clean(source);
  if (!value) return "";
  if (value.startsWith("data:image/")) return value;
  try {
    const response = await fetch(value);
    if (!response.ok) return "";
    const blob = await response.blob();
    return await blobToDataUrl(blob);
  } catch {
    return await imageUrlToDataUrl(value);
  }
}

function blobToDataUrl(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(new Error("Could not read blob"));
    reader.readAsDataURL(blob);
  });
}

function imageUrlToDataUrl(url) {
  return new Promise((resolve) => {
    const image = new Image();
    image.crossOrigin = "anonymous";
    image.onload = () => {
      try {
        const canvas = document.createElement("canvas");
        canvas.width = image.naturalWidth || 1;
        canvas.height = image.naturalHeight || 1;
        const ctx = canvas.getContext("2d");
        if (!ctx) {
          resolve("");
          return;
        }
        ctx.drawImage(image, 0, 0);
        resolve(canvas.toDataURL("image/png"));
      } catch {
        resolve("");
      }
    };
    image.onerror = () => resolve("");
    image.src = url;
  });
}

async function downloadWorkbook(filename, workbook) {
  const buffer = await workbook.xlsx.writeBuffer();
  downloadBlob(filename, buffer, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
}

function downloadBlob(filename, content, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

function downloadTextFile(filename, content, type) {
  downloadBlob(filename, content, type);
}

function csvValue(value) {
  const text = String(value ?? "");
  return `"${text.replaceAll('"', '""')}"`;
}

async function persistState() {
  normalizeState();
  if (cloudEnabled()) {
    hasPendingCloudChanges = true;
  }
  saveLocalState();
  updateStorageModeUi();
  render();
  if (cloudEnabled()) {
    await syncStateToCloud();
  }
}

async function loadInitialState() {
  const localState = loadLocalState();
  if (!cloudEnabled()) {
    setSyncStatus("Cloud sync is not configured yet.");
    return localState;
  }

  setSyncStatus("Checking cloud data...");
  try {
    const remote = await fetchCloudState();
    if (hasPendingCloudChanges && hasMeaningfulState(localState)) {
      setSyncStatus("Using newer local changes saved in this browser. Cloud sync will retry automatically.");
      return localState;
    }
    if (remote?.state) {
      lastCloudUpdatedAt = clean(remote.updatedAt);
      saveLocalState(remote.state);
      hasPendingCloudChanges = false;
      setSyncStatus(`Connected to cloud. Last sync ${formatSyncTime(remote.updatedAt)}.`);
      return remote.state;
    }

    setSyncStatus("Cloud storage is empty. Your next save will create the shared dataset.");
    return localState;
  } catch (error) {
    console.error(error);
    setSyncStatus(`Cloud sync could not connect. Using local browser data only. ${formatCloudError(error)}`);
    return localState;
  }
}

function cloudEnabled() {
  const cfg = window.APP_CONFIG?.firebase;
  return Boolean(
    isConfiguredValue(cfg?.apiKey) &&
    isConfiguredValue(cfg?.authDomain) &&
    isConfiguredValue(cfg?.projectId) &&
    isConfiguredValue(cfg?.databaseURL) &&
    isConfiguredValue(cfg?.storageBucket) &&
    window.firebase?.initializeApp
  );
}

function looksLikeFirebaseWebAppId(value) {
  return /^\d+:[^:]+:[^:]+:[A-Za-z0-9_-]+$/i.test(clean(value));
}

function getFirebaseInitConfig() {
  const cfg = window.APP_CONFIG?.firebase || {};
  const initConfig = {
    apiKey: clean(cfg.apiKey),
    authDomain: clean(cfg.authDomain),
    projectId: clean(cfg.projectId),
    databaseURL: clean(cfg.databaseURL),
    storageBucket: clean(cfg.storageBucket)
  };
  if (looksLikeFirebaseWebAppId(cfg.appId)) {
    initConfig.appId = clean(cfg.appId);
  }
  if (isConfiguredValue(cfg.messagingSenderId)) {
    initConfig.messagingSenderId = clean(cfg.messagingSenderId);
  }
  if (isConfiguredValue(cfg.measurementId)) {
    initConfig.measurementId = clean(cfg.measurementId);
  }
  return initConfig;
}

function getFirebaseApp() {
  if (!cloudEnabled()) return null;
  if (!firebaseApp) {
    const cfg = getFirebaseInitConfig();
    firebaseApp = window.firebase.apps?.length ? window.firebase.app() : window.firebase.initializeApp(cfg);
    firebaseDb = window.firebase.database(firebaseApp);
    firebaseStorage = window.firebase.storage(firebaseApp);
  }
  return firebaseApp;
}

function getFirebaseDatabase() {
  return getFirebaseApp() ? firebaseDb : null;
}

function getFirebaseStorage() {
  return getFirebaseApp() ? firebaseStorage : null;
}

function getCloudAppId() {
  return clean(window.APP_CONFIG?.firebase?.cloudAppId || window.APP_CONFIG?.firebase?.appId) || "piece-rate-main";
}

function getCloudStatePath() {
  return `apps/${getCloudAppId()}`;
}

async function fetchCloudState() {
  const db = getFirebaseDatabase();
  if (!db) return null;
  const snapshot = await db.ref(getCloudStatePath()).once("value");
  const data = snapshot.val();
  if (!data?.state_json) return null;
  return {
    state: { ...clone(DEFAULTS), ...decodeFirebaseValue(data.state_json) },
    updatedAt: data.updated_at || ""
  };
}

async function syncStateToCloud() {
  if (isSyncing || isApplyingRemoteState) return;
  isSyncing = true;
  setSyncStatus("Syncing data to cloud...");
  try {
    const db = getFirebaseDatabase();
    if (!db) throw new Error("Firebase is not configured.");
    const snapshot = encodeFirebaseValue(cloudStateSnapshot(state));
    const payload = {
      state_json: snapshot,
      updated_at: new Date().toISOString()
    };
    await db.ref(getCloudStatePath()).set(payload);
    lastCloudUpdatedAt = payload.updated_at;
    hasPendingCloudChanges = false;
    saveLocalState();
    setSyncStatus(`Connected to cloud. Last sync ${formatSyncTime(payload.updated_at)}.`);
  } catch (error) {
    console.error(error);
    hasPendingCloudChanges = true;
    saveLocalState();
    setSyncStatus(`Cloud sync failed. Local changes are kept in this browser and will retry automatically. ${formatCloudError(error)}`);
  } finally {
    isSyncing = false;
  }
}

function startCloudRefresh() {
  if (!cloudEnabled()) return;
  if (cloudRefreshTimer) window.clearInterval(cloudRefreshTimer);
  cloudRefreshTimer = window.setInterval(() => {
    void refreshFromCloud();
  }, CLOUD_REFRESH_MS);
}

async function refreshFromCloud() {
  if (!cloudEnabled() || isSyncing) return;
  try {
    if (hasPendingCloudChanges) {
      setSyncStatus("Cloud sync pending. Local changes stay on this device until upload succeeds.");
      await syncStateToCloud();
      if (hasPendingCloudChanges) return;
    }
    const remote = await fetchCloudState();
    if (!remote?.state) return;
    if (!remote.updatedAt || remote.updatedAt === lastCloudUpdatedAt) return;
    if (JSON.stringify(remote.state) === JSON.stringify(state)) {
      lastCloudUpdatedAt = remote.updatedAt;
      setSyncStatus(`Connected to cloud. Last sync ${formatSyncTime(remote.updatedAt)}.`);
      return;
    }
    isApplyingRemoteState = true;
    state = { ...clone(DEFAULTS), ...remote.state };
    normalizeState();
    saveLocalState();
    lastCloudUpdatedAt = remote.updatedAt;
    hasPendingCloudChanges = false;
    updateStorageModeUi();
    render();
    setSyncStatus(`Cloud changes received. Last sync ${formatSyncTime(remote.updatedAt)}.`);
  } catch (error) {
    console.error(error);
    setSyncStatus(`Cloud sync check failed. Working from local data until the next refresh. ${formatCloudError(error)}`);
  } finally {
    isApplyingRemoteState = false;
  }
}

function updateStorageModeUi() {
  if (!els.storageModeBadge || !els.syncStatusText) return;
  if (cloudEnabled()) {
    els.storageModeBadge.textContent = "Shared cloud sync enabled";
    els.storageModeBadge.classList.add("is-online");
  } else {
    els.storageModeBadge.textContent = "Local browser only";
    els.storageModeBadge.classList.remove("is-online");
  }
}

function setSyncStatus(message) {
  if (els.syncStatusText) {
    els.syncStatusText.textContent = message;
  }
}

function formatCloudError(error) {
  const message = clean(error?.message || error);
  return message ? message.slice(0, 180) : "Unknown cloud error.";
}

function isConfiguredValue(value) {
  const text = clean(value);
  return Boolean(text && !/^YOUR_/i.test(text));
}

function formatSyncTime(value) {
  if (!value) return "just now";
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? "just now" : date.toLocaleString("en-IN");
}

function saveLocalState(nextState = state) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(localStateSnapshot(nextState)));
}

function cloudStateSnapshot(sourceState = state) {
  const snapshot = localStateSnapshot(sourceState);
  delete snapshot.__pendingCloudChanges;
  snapshot.styleImages = buildPersistedStyleImageMap(sourceState);
  return snapshot;
}

const FIREBASE_KEY_PREFIX = "__k__";

function encodeFirebaseValue(value) {
  if (Array.isArray(value)) {
    return value.map((item) => encodeFirebaseValue(item));
  }
  if (!value || typeof value !== "object") {
    return value;
  }
  return Object.fromEntries(
    Object.entries(value).map(([key, entryValue]) => [
      encodeFirebaseKey(key),
      encodeFirebaseValue(entryValue)
    ])
  );
}

function decodeFirebaseValue(value) {
  if (Array.isArray(value)) {
    return value.map((item) => decodeFirebaseValue(item));
  }
  if (!value || typeof value !== "object") {
    return value;
  }
  return Object.fromEntries(
    Object.entries(value).map(([key, entryValue]) => [
      decodeFirebaseKey(key),
      decodeFirebaseValue(entryValue)
    ])
  );
}

function encodeFirebaseKey(key) {
  return `${FIREBASE_KEY_PREFIX}${encodeURIComponent(String(key ?? ""))}`;
}

function decodeFirebaseKey(key) {
  const text = String(key ?? "");
  if (!text.startsWith(FIREBASE_KEY_PREFIX)) return text;
  try {
    return decodeURIComponent(text.slice(FIREBASE_KEY_PREFIX.length));
  } catch {
    return text;
  }
}

function localStateSnapshot(sourceState = state) {
  const snapshot = clone(sourceState);
  snapshot.__pendingCloudChanges = hasPendingCloudChanges;
  snapshot.styleImages = {};
  snapshot.styles = (snapshot.styles || []).map((style) => {
    if (!/^data:image\//i.test(clean(style.image))) return style;
    return { ...style, image: storedImageRef(style.id) };
  });
  return snapshot;
}

function buildPersistedStyleImageMap(sourceState = state) {
  const imageMap = {};
  const styles = Array.isArray(sourceState?.styles) ? sourceState.styles : [];
  for (const style of styles) {
    const imageValue = clean(style?.image);
    if (!imageValue) continue;
    const styleId = isStoredImageRef(imageValue) ? storedImageKey(imageValue) : clean(style?.id);
    if (!styleId) continue;
    if (/^data:image\//i.test(imageValue)) {
      imageMap[styleId] = imageValue;
      continue;
    }
    if (!isStoredImageRef(imageValue)) continue;
    const cachedImage = clean(sourceState?.styleImages?.[styleId] || "") || styleImageCache.get(styleId) || "";
    if (/^data:image\//i.test(cachedImage)) {
      imageMap[styleId] = cachedImage;
    }
  }
  return imageMap;
}

function loadLocalState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      hasPendingCloudChanges = false;
      return clone(DEFAULTS);
    }
    const parsed = JSON.parse(raw);
    hasPendingCloudChanges = Boolean(parsed?.__pendingCloudChanges);
    if (parsed && typeof parsed === "object") {
      delete parsed.__pendingCloudChanges;
    }
    return { ...clone(DEFAULTS), ...parsed };
  } catch {
    hasPendingCloudChanges = false;
    return clone(DEFAULTS);
  }
}

function hasMeaningfulState(sourceState = state) {
  return JSON.stringify(localStateSnapshot(sourceState)) !== JSON.stringify(localStateSnapshot(DEFAULTS));
}

async function loadStoredStyleImages() {
  const entries = await listStoredStyleImages();
  styleImageCache = new Map(entries);
}

async function migrateStoredImageRefsToState() {
  let changed = false;
  state.styleImages = state.styleImages && typeof state.styleImages === "object" ? state.styleImages : {};
  for (const style of state.styles) {
    const imageValue = clean(style?.image);
    if (!isStoredImageRef(imageValue)) continue;
    const styleId = storedImageKey(imageValue);
    const cachedImage = clean(state.styleImages[styleId] || "") || styleImageCache.get(styleId) || "";
    if (!cachedImage) continue;
    state.styleImages[styleId] = cachedImage;
    if (!styleImageCache.has(styleId)) await putStoredStyleImage(styleId, cachedImage);
    changed = true;
  }
  if (changed) saveLocalState();
}

async function migrateInlineStyleImages() {
  let changed = false;
  for (const style of state.styles) {
    const imageValue = clean(style.image);
    if (!/^data:image\//i.test(imageValue) && !(cloudEnabled() && isStoredImageRef(imageValue))) continue;
    style.image = await prepareStyleImageForState(imageValue, style.id);
    changed = true;
  }
  if (changed) {
    if (cloudEnabled()) {
      hasPendingCloudChanges = true;
    }
    saveLocalState();
    if (cloudEnabled()) {
      await syncStateToCloud();
    }
  }
}

async function prepareStyleImageForState(imageValue, styleId) {
  const value = clean(imageValue);
  if (!styleId) return value;
  state.styleImages = state.styleImages && typeof state.styleImages === "object" ? state.styleImages : {};
  if (cloudEnabled()) {
    if (!value) {
      delete state.styleImages[styleId];
      await deleteStoredStyleImage(styleId);
      return "";
    }
    if (/^data:image\//i.test(value)) {
      try {
        const downloadUrl = await uploadStyleImageToCloud(styleId, value);
        delete state.styleImages[styleId];
        await deleteStoredStyleImage(styleId);
        return downloadUrl;
      } catch (error) {
        console.error("Cloud style image upload failed:", error);
        state.styleImages[styleId] = value;
        await putStoredStyleImage(styleId, value);
        return storedImageRef(styleId);
      }
    }
    if (isStoredImageRef(value)) {
      const cachedImage = clean(state.styleImages[styleId] || "") || styleImageCache.get(styleId) || "";
      if (cachedImage) {
        try {
          const downloadUrl = await uploadStyleImageToCloud(styleId, cachedImage);
          delete state.styleImages[styleId];
          await deleteStoredStyleImage(styleId);
          return downloadUrl;
        } catch (error) {
          console.error("Cloud style image upload failed:", error);
          state.styleImages[styleId] = cachedImage;
          if (!styleImageCache.has(styleId)) await putStoredStyleImage(styleId, cachedImage);
          return value;
        }
      }
      return "";
    }
    delete state.styleImages[styleId];
    await deleteStoredStyleImage(styleId);
    return value;
  }
  if (!value) {
    delete state.styleImages[styleId];
    await deleteStoredStyleImage(styleId);
    return "";
  }
  if (/^data:image\//i.test(value)) {
    state.styleImages[styleId] = value;
    await putStoredStyleImage(styleId, value);
    return storedImageRef(styleId);
  }
  if (isStoredImageRef(value)) {
    const cachedImage = clean(state.styleImages[styleId] || "") || styleImageCache.get(styleId) || "";
    if (cachedImage) {
      state.styleImages[styleId] = cachedImage;
      if (!styleImageCache.has(styleId)) await putStoredStyleImage(styleId, cachedImage);
      return value;
    }
    return "";
  }
  delete state.styleImages[styleId];
  await deleteStoredStyleImage(styleId);
  return value;
}

async function putStoredStyleImage(styleId, dataUrl) {
  styleImageCache.set(styleId, dataUrl);
  await withStyleImageStore("readwrite", (store, resolve, reject) => {
    const request = store.put({ id: styleId, dataUrl });
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error || new Error("Could not save style image"));
  });
}

async function deleteStoredStyleImage(styleId) {
  styleImageCache.delete(styleId);
  await withStyleImageStore("readwrite", (store, resolve, reject) => {
    const request = store.delete(styleId);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error || new Error("Could not delete style image"));
  });
}

async function listStoredStyleImages() {
  return withStyleImageStore("readonly", (store, resolve, reject) => {
    if (typeof store.getAll === "function") {
      const request = store.getAll();
      request.onsuccess = () => resolve((request.result || []).map((item) => [item.id, item.dataUrl]));
      request.onerror = () => reject(request.error || new Error("Could not load style images"));
      return;
    }
    const request = store.openCursor();
    const rows = [];
    request.onsuccess = () => {
      const cursor = request.result;
      if (!cursor) {
        resolve(rows);
        return;
      }
      rows.push([cursor.value.id, cursor.value.dataUrl]);
      cursor.continue();
    };
    request.onerror = () => reject(request.error || new Error("Could not load style images"));
  });
}

async function uploadStyleImageToCloud(styleId, dataUrl) {
  const storage = getFirebaseStorage();
  if (!storage) throw new Error("Firebase Storage is not configured.");
  const extension = inferImageExtensionFromDataUrl(dataUrl);
  const ref = storage.ref().child(`style-images/${getCloudAppId()}/${styleId}/${Date.now()}.${extension}`);
  const snapshot = await ref.putString(dataUrl, "data_url");
  return snapshot.ref.getDownloadURL();
}

function inferImageExtensionFromDataUrl(dataUrl) {
  const value = clean(dataUrl).toLowerCase();
  if (value.startsWith("data:image/png")) return "png";
  if (value.startsWith("data:image/webp")) return "webp";
  if (value.startsWith("data:image/gif")) return "gif";
  return "jpg";
}

function withStyleImageStore(mode, executor) {
  return openStyleImageDb().then((db) => new Promise((resolve, reject) => {
    const tx = db.transaction(STYLE_IMAGE_STORE, mode);
    const store = tx.objectStore(STYLE_IMAGE_STORE);
    tx.onerror = () => reject(tx.error || new Error("Image storage failed"));
    executor(store, resolve, reject);
  }));
}

function openStyleImageDb() {
  return new Promise((resolve, reject) => {
    if (!window.indexedDB) {
      reject(new Error("This browser does not support image storage."));
      return;
    }
    const request = window.indexedDB.open(STYLE_IMAGE_DB_NAME, 1);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(STYLE_IMAGE_STORE)) {
        db.createObjectStore(STYLE_IMAGE_STORE, { keyPath: "id" });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error("Could not open image storage"));
  });
}

function rowsOrEmpty(rows, colspan, msg) { return rows.length ? rows.join("") : `<tr><td colspan="${colspan}" class="empty-state">${esc(msg)}</td></tr>`; }
function byId(id) { return state.styles.find((s) => s.id === id); }
function clone(v) { return JSON.parse(JSON.stringify(v)); }
function todayIso() { return new Date().toISOString().slice(0, 10); }
function uid() { return window.crypto?.randomUUID ? window.crypto.randomUUID() : `${Date.now()}-${Math.random().toString(16).slice(2)}`; }
function clean(v) { return String(v || "").trim(); }
function normalizeDateValue(value) {
  const text = clean(value);
  if (!text) return "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  if (/^\d{8}$/.test(text)) return `${text.slice(0, 4)}-${text.slice(4, 6)}-${text.slice(6, 8)}`;
  const slashOrDash = text.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (slashOrDash) {
    let [, first, second, year] = slashOrDash;
    if (year.length === 2) year = `20${year}`;
    const day = first.padStart(2, "0");
    const month = second.padStart(2, "0");
    return `${year}-${month}-${day}`;
  }
  const yearFirst = text.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (yearFirst) {
    const [, year, month, day] = yearFirst;
    return `${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
  }
  const parsed = new Date(text);
  if (!Number.isNaN(parsed.getTime())) {
    const year = parsed.getFullYear();
    const month = `${parsed.getMonth() + 1}`.padStart(2, "0");
    const day = `${parsed.getDate()}`.padStart(2, "0");
    return `${year}-${month}-${day}`;
  }
  return text;
}
function num(v) { return Number(v || 0); }
function parseAmount(v) { return Number(String(v || "0").replace(/,/g, "")) || 0; }
function sumObj(obj) { return Object.values(obj || {}).reduce((s, v) => s + num(v), 0); }
function fmt(v) { return num(v).toLocaleString("en-IN", { minimumFractionDigits: Number.isInteger(num(v)) ? 0 : 2, maximumFractionDigits: 2 }); }
function fmtInt(v) { return num(v).toLocaleString("en-IN", { maximumFractionDigits: 0 }); }
function normalizeKey(v) { return clean(v).toLowerCase().replace(/[^a-z0-9]+/g, ""); }
function dedupeByKey(items, getKey) { const map = new Map(); items.forEach((item) => { const key = getKey(item); if (key && !map.has(key)) map.set(key, item); }); return [...map.values()]; }
function matchesTextSearch(values, search) {
  const needle = clean(search).toLowerCase();
  if (!needle) return true;
  return values.some((value) => clean(value).toLowerCase().includes(needle));
}
function diffDays(fromDate, toDate) {
  const from = new Date(fromDate);
  const to = new Date(toDate);
  if (Number.isNaN(from.getTime()) || Number.isNaN(to.getTime())) return 0;
  return Math.max(Math.floor((to - from) / 86400000), 0);
}
function maxIsoDate(a, b) { return clean(a) > clean(b) ? clean(a) : clean(b); }
function tallyDateToIso(value) {
  const raw = clean(value);
  return /^\d{8}$/.test(raw) ? `${raw.slice(0, 4)}-${raw.slice(4, 6)}-${raw.slice(6, 8)}` : raw;
}
function formatTallyDate(value) { return clean(value).replaceAll("-", ""); }
function formatDateTimeDisplay(value) {
  if (!value) return "";
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? value : date.toLocaleString("en-IN");
}
function escapeXml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}
function esc(v) { return String(v).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;").replaceAll("'", "&#39;"); }
function escAttr(v) { return esc(v); }
