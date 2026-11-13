import React, { useState, useEffect, useCallback, useMemo } from "react";
import axios from "axios";

let resolvedApiBase = null;
async function apiRequest(method, endpoint, data = undefined) {
  const host = typeof window !== "undefined" && window.location.hostname ? window.location.hostname : "10.10.0.15";
  const inRodSubpath = typeof window !== "undefined" && window.location.pathname.toLowerCase().startsWith("/rod");
  const candidates = resolvedApiBase
    ? [resolvedApiBase]
    : [
        inRodSubpath ? "/rod/api/rodtep" : "/api/rodtep",
        `http://${host}:5065/api/rodtep`,
        "http://10.10.0.15:5065/api/rodtep"
      ];

  let lastErr = null;
  for (const base of candidates) {
    try {
      const res = await axios({
        method,
        url: `${base}${endpoint}`,
        data,
        headers: { Accept: "application/json" }
      });
      if (typeof res.data === "string" && res.data.trim().startsWith("<")) {
        throw new Error("Received HTML instead of JSON from proxy");
      }
      resolvedApiBase = base;
      return res;
    } catch (err) {
      lastErr = err;
      if (err.response && err.response.status === 400) {
        throw err;
      }
    }
  }
  throw lastErr;
}

const SAMPLE_EXCEL_TSV = `EXPORTER,LIC NO,DATED,SCHEME,PORT CODE,  SCRIP AMOUNT  ,
THE PARADISE LINEN,2610023656,22.10.2026,ROSCTL,INMAA1,298820
STR GARMENTS,2611005776,05.11.2026,ROSCTL,INTUT6,291326
STR GARMENTS,2611005714,05.11.2026,ROSCTL,INMAA1,248477
VKS HYTECH PRIVATE LIMITED,2608030729,14.08.2026,RODTEP,INDEL4,208634
VISHNU ENTERPRISES,2611009820,06.11.2026,ROSCTL,INCOK1,205502
PAKIZA FASHION POINT,2611010269,07.11.2026,ROSCTL,INNSA1,184981
SIDDHRAJ CHEMICALS LLP,2611001420,03.11.2026,RODTEP,INHZA1,128933
MAULIK LEATHERCRAFTS PRIVATE LIMITED,2611002219,05.11.2026,RODTEP,INMAA4,125964
VISHNU ENTERPRISES,2611009819,06.11.2026,RODTEP,INCOK1,113853
STR GARMENTS,2611005755,05.11.2026,ROSCTL,INMAA4,113794
STR GARMENTS,2611005594,05.11.2026,ROSCTL,INCOK1,89189
NEW INDIA EXIM,2609024500,12.09.2026,RODTEP,INSBI6,68698
VISHNU ENTERPRISES,2611001974,04.11.2026,RODTEP,INCOK1,59148
SHREE VANI TRADING CO,2604178808,29.04.2026,RODTEP,INNSA1,57729
PAKIZA FASHION POINT,2611010268,07.11.2026,RODTEP,INNSA1,56182
SIDDHRAJ CHEMICALS LLP,2611001446,03.11.2026,RODTEP,INMUN1,50900
ESSCO INTERNATIONAL,2610047851,31.10.2026,RODTEP,INNSA1,42699
STR GARMENTS,2611005985,05.11.2026,ROSCTL,INBLR4,40187
M.B.INTERNATIONAL,2611000122,02.11.2026,RODTEP,INMUN1,33335
MAULIK LEATHERCRAFTS PRIVATE LIMITED,2611002249,05.11.2026,ROSCTL,INMAA4,29794
ILEX ODYSSEY LLP,2611001396,03.11.2026,RODTEP,INCOK1,29085
NEW INDIA EXIM,2606045474,24.06.2026,RODTEP,INMUN1,25495
SARITA HANDA EXPORTS PRIVATE LIMITED,2610047266,30.10.2026,RODTEP,INBOM4,19846
THE PARADISE LINEN,2611007929,05.11.2026,ROSCTL,INMAA1,9840
SIDDHRAJ CHEMICALS LLP,2611001442,03.11.2026,RODTEP,INNSA1,9324
MOHAN MOTORS,2610046108,29.10.2026,RODTEP,INWFD6,5056
BELIEVE COSMETICS PRIVATE LIMITED,2504000070,01.04.2025,RODTEP,INNSA1,5003
,,,,,2551794,`;

function formatDateDMY(isoOrRaw) {
  if (!isoOrRaw) return "";
  const s = String(isoOrRaw).trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[3]}.${m[2]}.${m[1]}`;
  return s;
}

function downloadChaExcel({ blNo, grpoDocNum, boeNo, draftBoeAmount, attachedScrips, allScrips = [] }) {
  const cleanBl = String(blNo || "BL").trim();
  const scripLookup = new Map((allScrips || []).map((s) => [String(s.LicNo).trim(), s]));

  let totalUsed = 0;
  const rowsHtml = (attachedScrips || [])
    .map((u, idx) => {
      const lic = String(u.LicNo || "").trim();
      const s = scripLookup.get(lic) || {};
      const exporter = u.ExporterName || s.ExporterName || "";
      const licDate = formatDateDMY(u.LicDate || s.LicDate || "");
      const scheme = u.Scheme || s.Scheme || "RODTEP";
      const portCode = u.PortCode || s.PortCode || "";
      const origAmt = Number(u.ScripAmount || s.ScripAmount || u.AmountUsed || 0);
      const usedAmt = Number(u.AmountUsed || 0);
      totalUsed += usedAmt;

      return `
        <tr>
          <td style="text-align:center;">${idx + 1}</td>
          <td style="font-weight:bold;">${cleanBl}</td>
          <td>${exporter}</td>
          <td style="mso-number-format:\@; font-weight:bold; text-align:center;">${lic}</td>
          <td style="text-align:center;">${licDate}</td>
          <td style="text-align:center;">${scheme}</td>
          <td style="text-align:center;">${portCode}</td>
          <td style="text-align:right; mso-number-format:#,##0.00;">${origAmt.toFixed(2)}</td>
          <td style="text-align:right; font-weight:bold; background:#e6f4ea; mso-number-format:#,##0.00;">${usedAmt.toFixed(2)}</td>
        </tr>
      `;
    })
    .join("");

  const html = `
    <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
    <head>
      <meta charset="UTF-8">
      <!--[if gte mso 9]>
      <xml>
        <x:ExcelWorkbook>
          <x:ExcelWorksheets>
            <x:ExcelWorksheet>
              <x:Name>CHA RoDTEP Scrips</x:Name>
              <x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions>
            </x:ExcelWorksheet>
          </x:ExcelWorksheets>
        </x:ExcelWorkbook>
      </xml>
      <![endif]-->
    </head>
    <body>
      <table border="1" cellspacing="0" cellpadding="6" style="border-collapse:collapse; font-family:Calibri, Arial, sans-serif; font-size:11pt;">
        <tr style="background:#1a1a1a; color:#ffffff; font-weight:bold; font-size:13pt;">
          <td colspan="9">BHAVYA POLYFILMS PRIVATE LIMITED — RoDTEP / RoSCTL SCRIP DETAILS FOR CHA (BOE FILING)</td>
        </tr>
        <tr style="background:#f4f2ef; font-weight:bold;">
          <td colspan="2">BL Number:</td>
          <td colspan="2" style="color:#c8102e; font-size:12pt;">${cleanBl}</td>
          <td>GRPO / BOE Ref:</td>
          <td colspan="2">${grpoDocNum ? "GRPO #" + grpoDocNum : ""} ${boeNo ? "| BOE: " + boeNo : "(Draft BOE)"}</td>
          <td>Draft BOE Duty (₹):</td>
          <td style="text-align:right; mso-number-format:#,##0.00;">${Number(draftBoeAmount || totalUsed).toFixed(2)}</td>
        </tr>
        <tr><td colspan="9"></td></tr>
        <tr style="background:#c8102e; color:#ffffff; font-weight:bold; text-align:center;">
          <th>Sr. No.</th>
          <th>BL Number</th>
          <th>Exporter Name</th>
          <th>RoDTEP / RoSCTL Lic No.</th>
          <th>Scrip Date</th>
          <th>Scheme</th>
          <th>Port Code</th>
          <th>Scrip Face Value (₹)</th>
          <th>Amount to Use in BOE (₹)</th>
        </tr>
        ${rowsHtml}
        <tr style="background:#f4f2ef; font-weight:bold; font-size:11.5pt;">
          <td colspan="8" style="text-align:right;">TOTAL RoDTEP / RoSCTL SCRIP AMOUNT TO UTILIZE FOR BL ${cleanBl}:</td>
          <td style="text-align:right; background:#d1fae5; mso-number-format:#,##0.00;">${totalUsed.toFixed(2)}</td>
        </tr>
      </table>
    </body>
    </html>
  `;

  const blob = new Blob([html], { type: "application/vnd.ms-excel;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  const safeBl = cleanBl.replace(/[^a-zA-Z0-9_-]/g, "_");
  a.href = url;
  a.download = `CHA_RoDTEP_Scrips_BL_${safeBl}.xls`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function cleanBlForExcel(v) { return String(v || "").trim(); }
function formatINR(val) {
  const n = Number(val || 0);
  if (isNaN(n)) return "0.00";
  return n.toLocaleString("en-IN", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

function parseClipboardScrips(rawText, defaultPort = "Mundra") {
  if (!rawText || !String(rawText).trim()) return [];
  const lines = String(rawText)
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .split("\n")
    .map((l) => l.trim())
    .filter(Boolean);

  const parsedRows = [];
  let headerMap = null; // { exporter, licNo, licDate, scheme, portCode, amount, portLoc }

  const isDateToken = (v) =>
    /^\d{1,2}[.\-/]\d{1,2}[.\-/]\d{2,4}$/.test(v) ||
    /^\d{4}[.\-/]\d{1,2}[.\-/]\d{1,2}$/.test(v);
  const isLicNoToken = (v) => /^\d{8,12}$/.test(v.replace(/\s+/g, ""));
  const isSchemeToken = (v) => /^(RODTEP|ROSCTL|ROSTL|SEIS|MEIS|DEPB)$/i.test(v);
  const isPortCodeToken = (v) => /^IN[A-Z0-9]{4}$/i.test(v);
  const isPortLocToken = (v) => /^(CHENNAI|MUNDRA|MAA|MUN)$/i.test(v);

  for (const line of lines) {
    const delimiter = line.includes("\t") ? "\t" : ",";
    const cols = line.split(delimiter).map((c) => c.trim());
    if (cols.length < 2) continue;

    const joinedUpper = cols.join(" ").toUpperCase();
    // Detect Header Row and build dynamic column index map
    if (
      joinedUpper.includes("EXPORTER") ||
      joinedUpper.includes("LIC NO") ||
      joinedUpper.includes("LICENCE") ||
      joinedUpper.includes("SCRIP NO") ||
      joinedUpper.includes("SCRIP AMOUNT")
    ) {
      headerMap = {};
      cols.forEach((colName, idx) => {
        const u = colName.toUpperCase();
        if (u.includes("EXPORT") || u.includes("PARTY") || u.includes("SELLER") || u.includes("NAME")) {
          if (headerMap.exporter === undefined) headerMap.exporter = idx;
        } else if (u.includes("LIC") || u.includes("SCRIP NO") || u.includes("SCRIPT NO") || u.includes("CERT")) {
          if (headerMap.licNo === undefined) headerMap.licNo = idx;
        } else if (u.includes("DATE") || u.includes("DT")) {
          if (headerMap.licDate === undefined) headerMap.licDate = idx;
        } else if (u.includes("SCHEME") || u.includes("TYPE")) {
          if (headerMap.scheme === undefined) headerMap.scheme = idx;
        } else if (u.includes("PORT CODE") || u.includes("EDI") || u.includes("CUSTOM")) {
          if (headerMap.portCode === undefined) headerMap.portCode = idx;
        } else if (u.includes("AMOUNT") || u.includes("VALUE") || u.includes("DUTY") || u.includes("CREDIT") || u.includes("INR")) {
          if (headerMap.amount === undefined) headerMap.amount = idx;
        } else if (u === "PORT" || u.includes("LOCATION") || u.includes("LOC")) {
          if (headerMap.portLoc === undefined) headerMap.portLoc = idx;
        }
      });
      continue;
    }

    let exporter = "";
    let licNo = "";
    let licDate = "";
    let scheme = "";
    let portCode = "";
    let amount = 0;
    let rawPortLoc = "";

    // 1. If a Header Row was detected and matches columns, extract via headerMap first
    if (headerMap && Object.keys(headerMap).length >= 2) {
      if (headerMap.exporter !== undefined) exporter = cols[headerMap.exporter] || "";
      if (headerMap.licNo !== undefined) licNo = cols[headerMap.licNo] || "";
      if (headerMap.licDate !== undefined) licDate = cols[headerMap.licDate] || "";
      if (headerMap.scheme !== undefined) scheme = (cols[headerMap.scheme] || "").toUpperCase();
      if (headerMap.portCode !== undefined) portCode = (cols[headerMap.portCode] || "").toUpperCase();
      if (headerMap.amount !== undefined) {
        const rawA = (cols[headerMap.amount] || "").replace(/[₹,\s]/g, "");
        if (!isNaN(Number(rawA))) amount = Number(rawA);
      }
      if (headerMap.portLoc !== undefined) rawPortLoc = cols[headerMap.portLoc] || "";
    }

    // 2. Smart Pattern-Based Auto-Detection (handles any column order, missing headers, or extra Sr.No columns)
    if (!licNo || amount <= 0) {
      const usedIndices = new Set();

      // Pass A: Find LicNo (8-12 digit number), Date, Scheme, PortCode, PortLocation
      cols.forEach((cell, i) => {
        if (!cell) return;
        const clean = cell.trim();
        if (!licDate && isDateToken(clean)) {
          licDate = clean;
          usedIndices.add(i);
        } else if (!scheme && isSchemeToken(clean)) {
          scheme = clean.toUpperCase() === "ROSTL" ? "ROSCTL" : clean.toUpperCase();
          usedIndices.add(i);
        } else if (!portCode && isPortCodeToken(clean)) {
          portCode = clean.toUpperCase();
          usedIndices.add(i);
        } else if (!rawPortLoc && isPortLocToken(clean)) {
          rawPortLoc = clean;
          usedIndices.add(i);
        } else if (!licNo && isLicNoToken(clean)) {
          licNo = clean.replace(/\s+/g, "");
          usedIndices.add(i);
        }
      });

      // Pass B: Find Amount (numeric value > 0, ignoring small Sr.No in col 0) and Exporter (text)
      cols.forEach((cell, i) => {
        if (usedIndices.has(i) || !cell) return;
        const clean = cell.trim();
        const numCandidate = clean.replace(/[₹,\s]/g, "");
        if (/^-?\d+(\.\d+)?$/.test(numCandidate)) {
          const val = Number(numCandidate);
          // Ignore serial numbers (e.g. 1..200 in first column when other columns exist)
          if (i === 0 && val >= 1 && val <= 500 && cols.length >= 4) {
            return;
          }
          if (amount <= 0 && val > 0) {
            amount = val;
            usedIndices.add(i);
          }
        } else if (!exporter && /[A-Za-z]{2,}/.test(clean)) {
          exporter = clean;
          usedIndices.add(i);
        }
      });
    }

    // Skip summary/total rows where LIC NO is still empty (e.g. ",,,,,2551794,")
    if (!licNo || isNaN(amount) || amount <= 0) continue;

    let portLocation = defaultPort;
    const upperLoc = (rawPortLoc || "").toUpperCase();
    if (upperLoc.includes("CHENNAI") || upperLoc === "MAA") {
      portLocation = "Chennai";
    } else if (upperLoc.includes("MUNDRA") || upperLoc === "MUN") {
      portLocation = "Mundra";
    } else if (rawPortLoc) {
      portLocation = rawPortLoc;
    }

    parsedRows.push({
      ExporterName: exporter || "—",
      LicNo: licNo,
      LicDate: licDate,
      Scheme: scheme || "RODTEP",
      PortCode: portCode,
      ScripAmount: amount,
      PortLocation: portLocation
    });
  }

  return parsedRows;
}

export default function App() {
  const [activeTab, setActiveTab] = useState("module1"); // "module1", "module2", "ledger"
  const [loading, setLoading] = useState(false);
  const [banner, setBanner] = useState(null);

  // Master Data from Backend
  const [vendors, setVendors] = useState([]);
  const [apInvoices, setApInvoices] = useState([]);
  const [grpos, setGrpos] = useState([]);
  const [hbls, setHbls] = useState([]);
  const [scrips, setScrips] = useState([]);
  const [utilizations, setUtilizations] = useState([]);
  const [glReconPendingRows, setGlReconPendingRows] = useState([]);
  const [glAllApRows, setGlAllApRows] = useState([]);
  const [glAllBoeCreditRows, setGlAllBoeCreditRows] = useState([]);
  const [glReconPendingTotal, setGlReconPendingTotal] = useState(0);
  const [glTotalPurchasedFromStart, setGlTotalPurchasedFromStart] = useState(0);
  const [glListViewMode, setGlListViewMode] = useState("PENDING_ONLY"); // "ALL_AP" | "PENDING_ONLY" | "BOE_CREDITS"

  // --- MODULE 1 STATE (Excel Paste -> Auto-Create AP Invoice) ---
  const [apMode, setApMode] = useState("CREATE_AP"); // "SAVE_ONLY" | "LINK_EXISTING" | "CREATE_AP"
  const [selectedCardCode, setSelectedCardCode] = useState("V00593");
  const [numAtCard, setNumAtCard] = useState("");
  const [docDate, setDocDate] = useState(() => new Date().toISOString().split("T")[0]);
  const [docDueDate, setDocDueDate] = useState(() => new Date().toISOString().split("T")[0]);
  const [discountAmount, setDiscountAmount] = useState("");
  const [discountPercent, setDiscountPercent] = useState("");
  const [apComments, setApComments] = useState("");
  const [selectedExistingApDocEntry, setSelectedExistingApDocEntry] = useState("");
  const [rawPasteText, setRawPasteText] = useState("");
  const [pastedScrips, setPastedScrips] = useState([]);
  const [submittingAp, setSubmittingAp] = useState(false);
  const [expandedApDocs, setExpandedApDocs] = useState({});
  const [expandedGrpoDocs, setExpandedGrpoDocs] = useState({});
  const [grpoListViewMode, setGrpoListViewMode] = useState("ATTACHED_ONLY"); // "ATTACHED_ONLY" | "PENDING_ONLY" | "ALL_GRPO"
  const [customDutyAccounts, setCustomDutyAccounts] = useState({ bpAccounts: [], glAccounts: [] });
  const [activeJeGrpoKey, setActiveJeGrpoKey] = useState(null);
  const [jeFormDate, setJeFormDate] = useState(() => new Date().toISOString().split("T")[0]);
  const [jeFormBoeNo, setJeFormBoeNo] = useState("");
  const [jeFormAmount, setJeFormAmount] = useState("");
  const [jeFormDutyCode, setJeFormDutyCode] = useState("V00040");
  const [jeFormDutyType, setJeFormDutyType] = useState("BP");
  const [jeFormDirection, setJeFormDirection] = useState("CREDIT_370001_DEBIT_DUTY");
  const [jeFormMemo, setJeFormMemo] = useState("");
  const [jeFormReconcileAp, setJeFormReconcileAp] = useState(true);
  const [submittingJe, setSubmittingJe] = useState(false);
  const toggleExpandGrpo = (key) => setExpandedGrpoDocs((prev) => ({ ...prev, [key]: !prev[key] }));
  const toggleExpandAp = (key) => setExpandedApDocs((prev) => ({ ...prev, [key]: !prev[key] }));

  // --- MODULE 2 STATE (BL + GRPO + Port + Draft BOE Amount -> Port-Locked Checkbox Scrip Selection) ---
  const [selectedGrpoDocEntry, setSelectedGrpoDocEntry] = useState("");
  const [selectedGrpoDocNum, setSelectedGrpoDocNum] = useState("");
  const [blNo, setBlNo] = useState("");
  const [selectedPort, setSelectedPort] = useState("Mundra"); // "Mundra" or "Chennai"
  const [draftBoeAmount, setDraftBoeAmount] = useState("");
  const [boeNo, setBoeNo] = useState("");
  const [isBoeGenerated, setIsBoeGenerated] = useState(false);
  const [blRemarks, setBlRemarks] = useState("");
  const [portViewMode, setPortViewMode] = useState("SHOW_ALL_LOCK_OTHER"); // "SHOW_ALL_LOCK_OTHER" or "ONLY_MATCHING_PORT"
  const [checkedScrips, setCheckedScrips] = useState({}); // { [LicNo]: amountUsed }
  const [submittingBl, setSubmittingBl] = useState(false);
  const [module2Search, setModule2Search] = useState("");
  const [ledgerSearch, setLedgerSearch] = useState("");

  const fetchAllData = useCallback(async () => {
    setLoading(true);
    try {
      const [vRes, apRes, grpoRes, scripRes, dutyRes] = await Promise.all([
        apiRequest("GET", "/vendors"),
        apiRequest("GET", "/ap-invoices"),
        apiRequest("GET", "/grpo-bl-list"),
        apiRequest("GET", "/scrips"),
        apiRequest("GET", "/custom-duty-accounts").catch(() => ({ data: { bpAccounts: [], glAccounts: [] } }))
      ]);
      if (dutyRes && dutyRes.data) {
        setCustomDutyAccounts(dutyRes.data);
      }
      setVendors(vRes.data || []);
      setApInvoices(apRes.data || []);
      setGrpos(grpoRes.data?.grpos || []);
      setHbls(grpoRes.data?.hbls || []);
      setScrips(scripRes.data?.scrips || []);
      setUtilizations(scripRes.data?.utilizations || []);
      setGlReconPendingRows(scripRes.data?.glReconPendingRows || []);
      setGlAllApRows(scripRes.data?.glAllApRows || []);
      setGlAllBoeCreditRows(scripRes.data?.glAllBoeCreditRows || []);
      setGlReconPendingTotal(Number(scripRes.data?.glReconPendingTotal || 0));
      setGlTotalPurchasedFromStart(Number(scripRes.data?.glTotalPurchasedFromStart || 0));
    } catch (err) {
      console.error("Failed to load RoDTEP data:", err);
      setBanner({
        type: "error",
        text: err.response?.data?.error || err.message || "Failed to load data from server."
      });
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    fetchAllData();
  }, [fetchAllData]);

  // Computed totals for Module 1 pasted scrips
  const pastedSummary = useMemo(() => {
    let total = 0;
    let mundraTotal = 0;
    let chennaiTotal = 0;
    let mundraCount = 0;
    let chennaiCount = 0;

    for (const s of pastedScrips) {
      const amt = Number(s.ScripAmount || 0);
      total += amt;
      if (String(s.PortLocation).toUpperCase() === "CHENNAI") {
        chennaiTotal += amt;
        chennaiCount++;
      } else {
        mundraTotal += amt;
        mundraCount++;
      }
    }

    const disc = Math.abs(Number(discountAmount || 0));
    const netPayable = Math.max(0, total - disc);

    return {
      count: pastedScrips.length,
      total: Number(total.toFixed(2)),
      mundraTotal: Number(mundraTotal.toFixed(2)),
      chennaiTotal: Number(chennaiTotal.toFixed(2)),
      mundraCount,
      chennaiCount,
      discount: Number(disc.toFixed(2)),
      netPayable: Number(netPayable.toFixed(2))
    };
  }, [pastedScrips, discountAmount]);

  // Per-AP Invoice Linked Scrips & Unlinked Balance Map
  const apLinkMap = useMemo(() => {
    const map = {};
    const getScripsForAp = (docEntry, docNum) =>
      (scrips || []).filter(
        (s) =>
          (docEntry && Number(s.APDocEntry) === Number(docEntry)) ||
          (docNum && Number(s.APDocNum) === Number(docNum))
      );

    for (const inv of apInvoices || []) {
      const list = getScripsForAp(inv.DocEntry, inv.DocNum);
      const linkedTotal = Number(
        list.reduce((sum, s) => sum + Number(s.ScripAmount || 0), 0).toFixed(2)
      );
      const origDebit = Number(inv.ScripFaceValue || 0);
      const pendingRecon = Number(inv.InternalReconPending || 0);
      const targetAmt = pendingRecon > 0 ? pendingRecon : origDebit;
      const unlinkedBalance = Math.max(0, Number((targetAmt - linkedTotal).toFixed(2)));
      const isFullyLinked = linkedTotal > 0 && unlinkedBalance <= 0.5;
      map[`ENTRY_${inv.DocEntry}`] = {
        scrips: list,
        linkedTotal,
        targetAmt,
        unlinkedBalance,
        isFullyLinked
      };
      if (inv.DocNum) {
        map[`NUM_${inv.DocNum}`] = map[`ENTRY_${inv.DocEntry}`];
      }
    }
    return map;
  }, [apInvoices, scrips]);

  // Only show AP invoices in Link dropdowns that still have unlinked balance > 0
  const linkableApInvoices = useMemo(() => {
    return (apInvoices || []).filter((inv) => {
      const info = apLinkMap[`ENTRY_${inv.DocEntry}`];
      return !info || !info.isFullyLinked;
    });
  }, [apInvoices, apLinkMap]);

  // Per-GRPO / BL Attached Scrips Map & Summary List
  const { grpoAttachmentRows, unattachedGrpos } = useMemo(() => {
    const usedUtilIds = new Set();
    const rows = [];

    const findMatchedSapBoeJe = ({ matchingUtils, gBl, gBoe, gNum, totalAttached, draftBoe }) => {
      const utilWithJe = (matchingUtils || []).find(
        (u) => u.JENum || u.JETransId || String(u.Remarks || "").includes("[JE#")
      );
      const isMarkedReconInUtils = (matchingUtils || []).some(
        (u) => !!u.ReconciledInSAP || String(u.Remarks || "").includes(":RECONCILED")
      );
      const targetAmt = Number(totalAttached || draftBoe || 0);
      const cleanBl = String(gBl || "").trim().toUpperCase();
      const cleanBoe = String(gBoe || "").trim().toUpperCase();
      const cleanGrpo = gNum ? String(gNum).trim() : "";

      // Match strictly by explicit JE tag on this GRPO's rows or GRPO #<num> in SAP JDT1 (NEVER by BLNo)
      const matchedCreditRow = (glAllBoeCreditRows || [])
        .slice()
        .reverse()
        .find((j) => {
          if (utilWithJe?.JETransId && Number(j.TransId) === Number(utilWithJe.JETransId)) return true;
          if (
            utilWithJe?.JENum &&
            (String(j.BaseRef) === String(utilWithJe.JENum) || String(j.TransId) === String(utilWithJe.JENum))
          ) {
            return true;
          }
          const text = `${j.Ref1 || ""} ${j.Ref2 || ""} ${j.LineMemo || ""} ${j.HeaderRef1 || ""} ${j.HeaderRef2 || ""} ${j.HeaderMemo || ""}`.toUpperCase();
          if (
            cleanGrpo &&
            (String(j.Ref1 || "").trim().toUpperCase() === "GRPO#" + cleanGrpo ||
              String(j.Ref2 || "").trim() === cleanGrpo ||
              text.includes("GRPO#" + cleanGrpo) ||
              text.includes("GRPO #" + cleanGrpo))
          ) {
            return true;
          }
          return false;
        });

      if (matchedCreditRow) {
        const isRecon =
          Number(matchedCreditRow.BalDueCred ?? 0) <= 0.5 &&
          Number(matchedCreditRow.BalDueDeb ?? 0) <= 0.5;
        return {
          JENum: matchedCreditRow.BaseRef || matchedCreditRow.TransId,
          JETransId: matchedCreditRow.TransId,
          JEDate: matchedCreditRow.RefDate || utilWithJe?.JEDate || null,
          ReconciledInSAP: isRecon
        };
      }

      if (utilWithJe) {
        return {
          JENum: utilWithJe.JENum || utilWithJe.JETransId,
          JETransId: utilWithJe.JETransId || utilWithJe.JENum,
          JEDate: utilWithJe.JEDate || null,
          ReconciledInSAP: isMarkedReconInUtils
        };
      }

      return {
        JENum: null,
        JETransId: null,
        JEDate: null,
        ReconciledInSAP: false
      };
    };

    for (const g of grpos || []) {
      const gEntry = Number(g.GRPODocEntry || 0);
      const gNum = Number(g.GRPODocNum || 0);
      const gBl = String(g.BLNo || "").trim().toUpperCase();

      const matchingUtils = (utilizations || []).filter((u) => {
        if (usedUtilIds.has(u.Id)) return false;
        const uEntry = Number(u.GRPODocEntry || 0);
        const uNum = Number(u.GRPODocNum || 0);
        if (uEntry || uNum) {
          if (gEntry && uEntry === gEntry) return true;
          if (gNum && uNum === gNum) return true;
          return false;
        }
        if (gBl && String(u.BLNo || "").trim().toUpperCase() === gBl) return true;
        return false;
      });

      matchingUtils.forEach((u) => usedUtilIds.add(u.Id));

      const totalAttached = Number(
        matchingUtils.reduce((sum, u) => sum + Number(u.AmountUsed || 0), 0).toFixed(2)
      );
      const draftBoe = matchingUtils.length > 0
        ? Math.max(...matchingUtils.map((u) => Number(u.DraftBOEAmount || 0)), Number(g.BoEValue || 0))
        : Number(g.BoEValue || 0);
      const isAttached = matchingUtils.length > 0 && totalAttached > 0;

      const resolvedBl = (matchingUtils[0] && matchingUtils[0].BLNo) ? matchingUtils[0].BLNo : (g.BLNo || "");
      const resolvedBoe = g.BOENo || (matchingUtils[0] ? matchingUtils[0].BOENo : "");
      const jeInfo = isAttached
        ? findMatchedSapBoeJe({
            matchingUtils,
            gBl: resolvedBl,
            gBoe: resolvedBoe,
            gNum: g.GRPODocNum,
            totalAttached,
            draftBoe
          })
        : { JENum: null, JETransId: null, JEDate: null, ReconciledInSAP: false };

      rows.push({
        rowKey: `GRPO_${g.GRPODocEntry}`,
        GRPODocEntry: g.GRPODocEntry,
        GRPODocNum: g.GRPODocNum,
        GRPODate: g.GRPODate,
        CardCode: g.CardCode,
        CardName: g.CardName,
        VendorRefNo: g.VendorRefNo,
        BLNo: resolvedBl,
        BOENo: resolvedBoe,
        DraftBOEAmount: draftBoe,
        totalAttached,
        attachedScrips: matchingUtils,
        isAttached,
        JENum: jeInfo.JENum,
        JETransId: jeInfo.JETransId,
        JEDate: jeInfo.JEDate,
        ReconciledInSAP: jeInfo.ReconciledInSAP
      });
    }

    // Also include any BL attachments in utilizations that were entered manually without matching a row in grpos
    const orphanByBl = {};
    for (const u of utilizations || []) {
      if (usedUtilIds.has(u.Id)) continue;
      const key = String(u.BLNo || u.GRPODocNum || u.Id).trim().toUpperCase();
      if (!orphanByBl[key]) {
        orphanByBl[key] = [];
      }
      orphanByBl[key].push(u);
    }

    for (const [blKey, list] of Object.entries(orphanByBl)) {
      const first = list[0];
      const totalAttached = Number(
        list.reduce((sum, u) => sum + Number(u.AmountUsed || 0), 0).toFixed(2)
      );
      const draftBoe = Math.max(...list.map((u) => Number(u.DraftBOEAmount || 0)));
      const jeInfo = findMatchedSapBoeJe({
        matchingUtils: list,
        gBl: first.BLNo,
        gBoe: first.BOENo || "",
        gNum: first.GRPODocNum,
        totalAttached,
        draftBoe
      });
      rows.push({
        rowKey: `BL_${blKey}`,
        GRPODocEntry: first.GRPODocEntry || null,
        GRPODocNum: first.GRPODocNum || null,
        GRPODate: first.CreatedAt ? String(first.CreatedAt).slice(0, 10) : "—",
        CardCode: "",
        CardName: "Direct BL Attachment",
        VendorRefNo: "",
        BLNo: first.BLNo,
        BOENo: first.BOENo || "",
        DraftBOEAmount: draftBoe,
        totalAttached,
        attachedScrips: list,
        isAttached: true,
        JENum: jeInfo.JENum,
        JETransId: jeInfo.JETransId,
        JEDate: jeInfo.JEDate,
        ReconciledInSAP: jeInfo.ReconciledInSAP
      });
    }

    const unattached = (grpos || []).filter((g) => {
      const r = rows.find((x) => Number(x.GRPODocEntry) === Number(g.GRPODocEntry));
      return !r || !r.isAttached;
    });

    return { grpoAttachmentRows: rows, unattachedGrpos: unattached };
  }, [grpos, utilizations, glAllBoeCreditRows]);

  // Existing LicNo set to warn on duplicates in Module 1
  const existingLicSet = useMemo(() => {
    return new Set((scrips || []).map((s) => String(s.LicNo).trim()));
  }, [scrips]);

  const handlePasteTextChange = (text) => {
    setRawPasteText(text);
    const rows = parseClipboardScrips(text);
    if (rows.length > 0) {
      setPastedScrips(rows);
      if (discountPercent && !isNaN(Number(discountPercent))) {
        const tot = rows.reduce((acc, r) => acc + Number(r.ScripAmount || 0), 0);
        setDiscountAmount(String(Math.round((tot * Number(discountPercent)) / 100)));
      }
    }
  };

  const handleLoadAttachedExcelSample = () => {
    handlePasteTextChange(SAMPLE_EXCEL_TSV);
    setSelectedCardCode("V00593");
    setNumAtCard("SCRP15/11/26-27");
    setDocDate("2026-11-11");
    setDocDueDate("2026-12-11");
    setDiscountAmount("20414");
    setBanner({
      type: "info",
      text: "Loaded the 27 RoDTEP / RoSCTL scrips from your attached Excel sheet (Face Value: ₹25,51,794.00 | Gain 410023: -₹20,414.00 | Net AP Invoice: ₹25,31,380.00)."
    });
  };

  const handleSelectExistingAp = (docEntryStr) => {
    setSelectedExistingApDocEntry(docEntryStr);
    const found = apInvoices.find((a) => String(a.DocEntry) === String(docEntryStr));
    if (found) {
      setSelectedCardCode(found.CardCode);
      setNumAtCard(found.NumAtCard || "");
      setDocDate(found.DocDate || "");
      setDocDueDate(found.DocDueDate || found.DocDate || "");
      setDiscountAmount(String(Math.abs(Number(found.GainDiscountAmount || 0))));
    }
  };

  const handleUpdatePastedRow = (idx, field, val) => {
    setPastedScrips((prev) => {
      const copy = [...prev];
      copy[idx] = {
        ...copy[idx],
        [field]: field === "ScripAmount" ? Number(String(val).replace(/,/g, "")) || 0 : val
      };
      return copy;
    });
  };

  const handleRemovePastedRow = (idx) => {
    setPastedScrips((prev) => prev.filter((_, i) => i !== idx));
  };

  const handleSubmitModule1 = async (e) => {
    e.preventDefault();
    setBanner(null);

    if (pastedScrips.length === 0) {
      setBanner({ type: "error", text: "Please paste RoDTEP scrip rows from Excel first." });
      return;
    }

    const selectedVendorObj = vendors.find((v) => v.CardCode === selectedCardCode);
    const existingApObj = apInvoices.find(
      (a) => String(a.DocEntry) === String(selectedExistingApDocEntry)
    );

    if (apMode === "LINK_EXISTING" && !existingApObj) {
      setBanner({ type: "error", text: "Please select an existing SAP A/P Invoice to link." });
      return;
    }

    setSubmittingAp(true);
    try {
      const payload = {
        mode: apMode,
        cardCode: selectedCardCode,
        cardName: selectedVendorObj?.CardName || existingApObj?.CardName || "",
        numAtCard,
        docDate,
        docDueDate,
        discountAmount: Number(discountAmount || 0),
        comments: apComments,
        existingApDocEntry: existingApObj?.DocEntry || null,
        existingApDocNum: existingApObj?.DocNum || null,
        scrips: pastedScrips
      };

      const res = await apiRequest("POST", "/create-ap-and-scrips", payload);
      const data = res.data;

      setBanner({
        type: "success",
        text:
          apMode === "CREATE_AP"
            ? `Created SAP A/P Invoice #${data.apDocNum} (DocEntry: ${data.apDocEntry}) via Service Layer for ₹${formatINR(data.netInvoiceTotal)} and saved ${data.scripsCount} RoDTEP scrips!`
            : apMode === "SAVE_ONLY"
            ? `Saved ${data.scripsCount} Pending RoDTEP scrips (₹${formatINR(data.totalScripFaceValue)}) to Table 1! You can use them on BLs immediately and link their SAP A/P Invoice # anytime later.`
            : `Linked ${data.scripsCount} RoDTEP scrips (₹${formatINR(data.totalScripFaceValue)}) to SAP A/P Invoice #${data.apDocNum} (DocEntry: ${data.apDocEntry})!`
      });

      setPastedScrips([]);
      setRawPasteText("");
      await fetchAllData();
    } catch (err) {
      setBanner({
        type: "error",
        text: err.response?.data?.error || err.message || "Failed to save A/P Invoice & Scrips."
      });
    } finally {
      setSubmittingAp(false);
    }
  };

  // --- MODULE 2 LOGIC (BL + GRPO + Port + Draft BOE Amount) ---
  const handleSelectGrpo = (grpoDocEntryStr) => {
    setSelectedGrpoDocEntry(grpoDocEntryStr);
    if (!grpoDocEntryStr) {
      setSelectedGrpoDocNum("");
      return;
    }
    const g = grpos.find((x) => String(x.GRPODocEntry) === String(grpoDocEntryStr));
    if (g) {
      setSelectedGrpoDocNum(String(g.GRPODocNum || ""));
      setBlNo(g.BLNo || "");
      if (g.DetectedPort) {
        setSelectedPort(g.DetectedPort);
        // Clear any checked scrips that don't match the new port
        setCheckedScrips({});
      }
      if (g.BOENo) {
        setBoeNo(g.BOENo);
        setIsBoeGenerated(true);
      } else {
        setBoeNo("");
        setIsBoeGenerated(false);
      }
    }
  };

  const handlePortChange = (newPort) => {
    setSelectedPort(newPort);
    // Automatically uncheck any scrip that does not belong to newPort so wrong port can never remain checked
    setCheckedScrips((prev) => {
      const cleaned = {};
      for (const [lic, amt] of Object.entries(prev)) {
        const scripObj = scrips.find((s) => s.LicNo === lic);
        if (
          scripObj &&
          String(scripObj.PortLocation || "").toUpperCase() === newPort.toUpperCase()
        ) {
          cleaned[lic] = amt;
        }
      }
      return cleaned;
    });
  };

  // Pending scrips where BalanceAmount > 0 (calculated on-the-fly by backend)
  const pendingScrips = useMemo(() => {
    return (scrips || []).filter((s) => Number(s.BalanceAmount || 0) > 0.001);
  }, [scrips]);

  const displayedPendingScrips = useMemo(() => {
    const q = module2Search.trim().toLowerCase();
    const filtered = pendingScrips.filter((s) => {
      if (!q) return true;
      return (
        String(s.LicNo || "").toLowerCase().includes(q) ||
        String(s.ExporterName || "").toLowerCase().includes(q) ||
        String(s.PortCode || "").toLowerCase().includes(q) ||
        String(s.Scheme || "").toLowerCase().includes(q) ||
        String(s.APDocNum || "").toLowerCase().includes(q) ||
        String(s.NumAtCard || "").toLowerCase().includes(q) ||
        String(s.ScripAmount || "").includes(q) ||
        String(s.BalanceAmount || "").includes(q)
      );
    });

    return [...filtered].sort((a, b) =>
      String(a.LicDate || "").localeCompare(String(b.LicDate || ""))
    );
  }, [pendingScrips, module2Search]);

  const handleToggleScripCheckbox = (scrip) => {

    setCheckedScrips((prev) => {
      const next = { ...prev };
      if (next[scrip.LicNo] !== undefined) {
        delete next[scrip.LicNo];
      } else {
        // Auto-fill the exact remaining balance of the scrip
        next[scrip.LicNo] = Number(scrip.BalanceAmount || 0);
      }
      return next;
    });
  };

  const handleChangeScripUsedAmount = (scrip, rawVal) => {
    const maxBal = Number(scrip.BalanceAmount || 0);
    let val = Number(String(rawVal).replace(/,/g, ""));
    if (isNaN(val) || val < 0) val = 0;
    if (val > maxBal) val = maxBal;

    setCheckedScrips((prev) => ({
      ...prev,
      [scrip.LicNo]: val
    }));
  };

  // Auto-fill FIFO matching Draft BOE Amount
  const handleAutoAllocateDraftBoe = () => {
    const target = Number(String(draftBoeAmount || "").replace(/,/g, ""));
    if (!target || isNaN(target) || target <= 0) {
      setBanner({
        type: "error",
        text: "Please enter the Draft BOE Amount first to auto-select matching scrips."
      });
      return;
    }

    let remainingNeeded = target;
    const nextChecked = {};

    for (const s of displayedPendingScrips) {
      if (remainingNeeded <= 0.001) break;
      const avail = Number(s.BalanceAmount || 0);
      if (avail <= 0) continue;

      const useAmt = Math.min(avail, Number(remainingNeeded.toFixed(2)));
      nextChecked[s.LicNo] = Number(useAmt.toFixed(2));
      remainingNeeded = Number((remainingNeeded - useAmt).toFixed(2));
    }

    setCheckedScrips(nextChecked);
  };

  const allocationSummary = useMemo(() => {
    let totalUtilized = 0;
    let checkedCount = 0;
    for (const amt of Object.values(checkedScrips)) {
      const n = Number(amt || 0);
      if (n > 0) {
        totalUtilized += n;
        checkedCount++;
      }
    }
    totalUtilized = Number(totalUtilized.toFixed(2));
    const targetBoe = Number(Number(String(draftBoeAmount || 0).replace(/,/g, "")).toFixed(2));
    const diff = Number((targetBoe - totalUtilized).toFixed(2));
    const isExactMatch = targetBoe > 0 && Math.abs(diff) <= 0.01;

    return {
      checkedCount,
      totalUtilized,
      targetBoe,
      diff,
      isExactMatch
    };
  }, [checkedScrips, draftBoeAmount]);

  const handleSubmitModule2 = async (e) => {
    e.preventDefault();
    setBanner(null);

    if (!blNo.trim()) {
      setBanner({ type: "error", text: "Please select or enter a BL Number." });
      return;
    }

    const allocations = Object.entries(checkedScrips)
      .filter(([_, amt]) => Number(amt) > 0)
      .map(([licNo, amt]) => {
        const s = scrips.find((x) => x.LicNo === licNo) || {};
        return {
          LicNo: licNo,
          AmountUsed: Number(amt),
          PortLocation: s.PortLocation || selectedPort,
          ExporterName: s.ExporterName,
          Scheme: s.Scheme,
          PortCode: s.PortCode,
          ScripAmount: s.ScripAmount,
          APDocNum: s.APDocNum
        };
      });

    if (allocations.length === 0) {
      setBanner({
        type: "error",
        text: "Please check at least one RoDTEP scrip from the table."
      });
      return;
    }

    // When BOE is not yet generated (Draft stage), require that Total Utilized == Draft BOE Amount
    const mustMatchDraftBoe = !isBoeGenerated || Number(draftBoeAmount || 0) > 0;
    if (mustMatchDraftBoe) {
      if (allocationSummary.targetBoe <= 0) {
        setBanner({
          type: "error",
          text: "Please enter the Draft BOE Amount so the system can verify that total Scrip Amount Utilized equals the Draft BOE Amount."
        });
        return;
      }
      if (!allocationSummary.isExactMatch) {
        setBanner({
          type: "error",
          text: `Amount Utilized (₹${formatINR(allocationSummary.totalUtilized)}) must equal the Draft BOE Amount (₹${formatINR(allocationSummary.targetBoe)}). Difference: ₹${formatINR(allocationSummary.diff)}. Adjust the editable scrip amount column to match.`
        });
        return;
      }
    }

    setSubmittingBl(true);
    try {
      const payload = {
        blNo: blNo.trim(),
        portLocation: selectedPort,
        grpoDocEntry: selectedGrpoDocEntry ? Number(selectedGrpoDocEntry) : null,
        grpoDocNum: selectedGrpoDocNum ? Number(selectedGrpoDocNum) : null,
        draftBoeAmount: allocationSummary.targetBoe,
        boeNo: boeNo.trim(),
        enforceDraftBoeMatch: mustMatchDraftBoe,
        remarks: blRemarks.trim(),
        allocations
      };

      const res = await apiRequest("POST", "/allocate-bl", payload);
      downloadChaExcel({
        blNo: cleanBlForExcel(blNo),
        grpoDocNum: selectedGrpoDocNum,
        boeNo: boeNo.trim(),
        draftBoeAmount: allocationSummary.targetBoe,
        attachedScrips: allocations,
        allScrips: scrips
      });
      setBanner({
        type: "success",
        text: `Attached ${res.data.count} RoDTEP scrip(s) totaling ₹${formatINR(res.data.totalAllocated)} to BL ${res.data.blNo}${res.data.grpoDocNum ? ` (GRPO #${res.data.grpoDocNum})` : ""} & downloaded CHA Excel sheet!`
      });

      setCheckedScrips({});
      setSelectedGrpoDocEntry("");
      setSelectedGrpoDocNum("");
      setBlNo("");
      setDraftBoeAmount("");
      await fetchAllData();
    } catch (err) {
      setBanner({
        type: "error",
        text: err.response?.data?.error || err.message || "Failed to attach scrips to BL."
      });
    } finally {
      setSubmittingBl(false);
    }
  };

  const getGrpoApBreakdown = (r) => {
    const byAp = new Map();
    for (const u of r.attachedScrips || []) {
      const sc = scrips.find((s) => s.LicNo === u.LicNo) || {};
      const apEntry = sc.APDocEntry || u.APDocEntry || null;
      const apNum = sc.APDocNum || u.APDocNum || null;
      const invObj =
        apInvoices.find(
          (a) =>
            (apEntry && Number(a.DocEntry) === Number(apEntry)) ||
            (apNum && Number(a.DocNum) === Number(apNum))
        ) || {};
      const numAtCard = sc.NumAtCard || invObj.NumAtCard || "";
      const cardName = sc.CardName || invObj.CardName || "";
      const key = apEntry ? `ENTRY_${apEntry}` : apNum ? `NUM_${apNum}` : "UNLINKED";

      const prev = byAp.get(key) || {
        apDocEntry: apEntry,
        apDocNum: apNum,
        numAtCard,
        cardName,
        licNos: [],
        amount: 0
      };
      prev.licNos.push(u.LicNo);
      prev.amount = Number((prev.amount + Number(u.AmountUsed || 0)).toFixed(2));
      byAp.set(key, prev);
    }
    return Array.from(byAp.values());
  };

  const handleOpenJeForm = (r) => {
    if (activeJeGrpoKey === r.rowKey) {
      setActiveJeGrpoKey(null);
      return;
    }
    setActiveJeGrpoKey(r.rowKey);
    setJeFormBoeNo(r.BOENo || "");
    setJeFormAmount(String(r.totalAttached || r.DraftBOEAmount || 0));
    setJeFormDate(new Date().toISOString().split("T")[0]);
    setJeFormDutyCode("V00040");
    setJeFormDutyType("BP");
    setJeFormDirection("CREDIT_370001_DEBIT_DUTY");
    setJeFormMemo(
      r.BOENo
        ? `BCD of BOE No ${r.BOENo}`
        : `BCD of BL ${r.BLNo || ""}${r.GRPODocNum ? " GRPO #" + r.GRPODocNum : ""}`
    );
    setJeFormReconcileAp(true);
  };

  const handlePostGrpoJe = async (r) => {
    setSubmittingJe(true);
    setBanner(null);
    try {
      const linkedApDocEntries = Array.from(
        new Set(
          (r.attachedScrips || [])
            .map((u) => {
              const sc = scrips.find((s) => s.LicNo === u.LicNo);
              return sc?.APDocEntry || null;
            })
            .filter(Boolean)
        )
      );

      const isFormOpenForThisRow = activeJeGrpoKey === r.rowKey;
      const finalBoeNo = isFormOpenForThisRow ? jeFormBoeNo : (r.BOENo || "");
      const finalDate = isFormOpenForThisRow && jeFormDate ? jeFormDate : new Date().toISOString().split("T")[0];
      const finalAmt = isFormOpenForThisRow && Number(jeFormAmount) > 0
        ? Number(jeFormAmount)
        : Number(r.totalAttached || r.DraftBOEAmount || 0);
      const finalDutyCode = isFormOpenForThisRow && jeFormDutyCode ? jeFormDutyCode : "V00040";
      const finalDutyType = isFormOpenForThisRow && jeFormDutyType ? jeFormDutyType : "BP";
      const finalDirection = isFormOpenForThisRow && jeFormDirection ? jeFormDirection : "CREDIT_370001_DEBIT_DUTY";
      const finalMemo = isFormOpenForThisRow && jeFormMemo
        ? jeFormMemo
        : (finalBoeNo
            ? `BCD of BOE No ${finalBoeNo}`
            : `BCD of BL ${r.BLNo || ""}${r.GRPODocNum ? " GRPO #" + r.GRPODocNum : ""}`);

      const apReconBreakdown = getGrpoApBreakdown(r);
      const res = await apiRequest("POST", "/post-grpo-je", {
        grpoDocEntry: r.GRPODocEntry,
        grpoDocNum: r.GRPODocNum,
        blNo: r.BLNo,
        boeNo: finalBoeNo,
        jeDate: finalDate,
        amount: finalAmt,
        customDutyCode: finalDutyCode,
        customDutyAccountType: finalDutyType,
        jeDirection: finalDirection,
        memo: finalMemo,
        reconcileWithAp: true,
        reconcileApDocEntries: linkedApDocEntries,
        apReconBreakdown
      });

      const d = res.data;
      setBanner({
        type: "success",
        text: `Posted SAP Journal Entry #${d.jdtNum || d.transId} for ₹${formatINR(d.amount)} (${
          d.jeDirection === "DEBIT_370001_CREDIT_DUTY"
            ? `Debit 370001 / Credit ${d.customDutyCode}`
            : `Debit ${d.customDutyCode} / Credit 370001`
        })${d.reconStatus?.message ? " · " + d.reconStatus.message : ""}!`
      });
      setActiveJeGrpoKey(null);
      await fetchAllData();
    } catch (err) {
      setBanner({
        type: "error",
        text: err.response?.data?.error || err.message || "Failed to post Custom Duty JE."
      });
    } finally {
      setSubmittingJe(false);
    }
  };

  const handleReconcileExistingJe = async (r) => {
    setSubmittingJe(true);
    setBanner(null);
    try {
      const apReconBreakdown = getGrpoApBreakdown(r);
      const res = await apiRequest("POST", "/post-grpo-je", {
        grpoDocEntry: r.GRPODocEntry,
        grpoDocNum: r.GRPODocNum,
        blNo: r.BLNo,
        boeNo: r.BOENo || "",
        jeDate: r.JEDate || new Date().toISOString().split("T")[0],
        amount: Number(r.totalAttached || r.DraftBOEAmount || 0),
        reconcileWithAp: true,
        reconcileOnlyExistingJe: true,
        existingJeNum: r.JENum,
        existingJeTransId: r.JETransId,
        apReconBreakdown
      });
      const d = res.data;
      setBanner({
        type: d.reconStatus?.reconciled ? "success" : "error",
        text: d.reconStatus?.message || `Processed reconciliation for JE #${r.JENum}.`
      });
      await fetchAllData();
    } catch (err) {
      setBanner({
        type: "error",
        text: err.response?.data?.error || err.message || "Failed to reconcile existing JE."
      });
    } finally {
      setSubmittingJe(false);
    }
  };

  const handleCancelReconAndReverseJe = async (r, detachScrips = true) => {
    const confirmMsg = detachScrips
      ? `Cancel SAP Internal Reconciliation, reverse SAP JE #${r.JENum} via Service Layer, and detach all ${r.attachedScrips.length} scrip(s) from ${r.GRPODocNum ? "GRPO #" + r.GRPODocNum : "BL " + r.BLNo} so you can attach them to the right GRPO?`
      : `Cancel SAP Internal Reconciliation and reverse SAP JE #${r.JENum} via Service Layer (keeping scrips attached)?`;
    if (!window.confirm(confirmMsg)) return;

    setSubmittingJe(true);
    setBanner(null);
    try {
      const res = await apiRequest("POST", "/cancel-grpo-je-and-detach", {
        grpoDocEntry: r.GRPODocEntry,
        grpoDocNum: r.GRPODocNum,
        blNo: r.BLNo,
        jeNum: r.JENum,
        jeTransId: r.JETransId,
        detachScrips
      });
      setBanner({
        type: "success",
        text: res.data?.message || "Cancelled reconciliation, reversed JE, and restored scrip balances!"
      });
      await fetchAllData();
    } catch (err) {
      setBanner({
        type: "error",
        text: err.response?.data?.error || err.message || "Failed to cancel reconciliation / reverse JE."
      });
    } finally {
      setSubmittingJe(false);
    }
  };

  const handleDetachAllForGrpo = async (r) => {
    try {
      const ids = (r.attachedScrips || []).map((u) => u.Id).filter(Boolean);
      await apiRequest("POST", "/detach-grpo", {
        grpoDocEntry: r.GRPODocEntry || null,
        grpoDocNum: r.GRPODocNum || null,
        blNo: r.BLNo || "",
        ids
      });
      await fetchAllData();
      setBanner({
        type: "info",
        text: `Detached all ${ids.length} scrip(s) from ${
          r.GRPODocNum ? "GRPO #" + r.GRPODocNum : "BL " + r.BLNo
        }. On-the-fly scrip balances have been restored!`
      });
    } catch (err) {
      setBanner({
        type: "error",
        text: err.response?.data?.error || err.message || "Failed to detach scrips."
      });
    }
  };

  const handleDeleteUtilization = async (id) => {
    try {
      await apiRequest("DELETE", `/utilization/${id}`);
      await fetchAllData();
      setBanner({
        type: "info",
        text: "Removed BL scrip attachment. On-the-fly Scrip Balance has been restored."
      });
    } catch (err) {
      setBanner({
        type: "error",
        text: err.response?.data?.error || err.message
      });
    }
  };

  const handleLinkApToExistingScrip = async (licNo, apDocEntryStr) => {
    try {
      const foundAp = apInvoices.find((a) => String(a.DocEntry) === String(apDocEntryStr));
      await apiRequest("PATCH", "/scrip-ap-link", {
        licNo,
        apDocEntry: foundAp ? foundAp.DocEntry : null,
        apDocNum: foundAp ? foundAp.DocNum : null,
        cardCode: foundAp ? foundAp.CardCode : "",
        cardName: foundAp ? foundAp.CardName : "",
        numAtCard: foundAp ? foundAp.NumAtCard : "",
        apDocDate: foundAp ? foundAp.DocDate : null
      });
      await fetchAllData();
      setBanner({
        type: "success",
        text: foundAp
          ? `Linked Scrip ${licNo} to SAP A/P Invoice #${foundAp.DocNum} (${foundAp.NumAtCard || foundAp.CardName})!`
          : `Unlinked A/P Invoice from Scrip ${licNo}.`
      });
    } catch (err) {
      setBanner({
        type: "error",
        text: err.response?.data?.error || err.message || "Failed to update AP Invoice link."
      });
    }
  };

  // Top-level KPI summary across all scrips in Table 1 + Table 2
  const globalKpis = useMemo(() => {
    let totalPurchased = 0;
    let totalUtilized = 0;
    let totalBalance = 0;
    let mundraBalance = 0;
    let chennaiBalance = 0;

    for (const s of scrips) {
      const amt = Number(s.ScripAmount || 0);
      const util = Number(s.UtilizedAmount || 0);
      const bal = Number(s.BalanceAmount || 0);
      totalPurchased += amt;
      totalUtilized += util;
      totalBalance += bal;
      if (String(s.PortLocation || "").toUpperCase() === "CHENNAI") {
        chennaiBalance += bal;
      } else {
        mundraBalance += bal;
      }
    }

    return {
      scripCount: scrips.length,
      pendingCount: pendingScrips.length,
      totalPurchased,
      totalUtilized,
      totalBalance,
      mundraBalance,
      chennaiBalance
    };
  }, [scrips, pendingScrips]);

  return (
    <div className="min-h-screen flex flex-col bg-[#faf9f7] text-[#333333]">
      {/* Top Header */}
      <header className="sticky top-0 z-40 border-b border-black/[0.08] bg-[#faf9f7]/95 backdrop-blur-md">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-[72px] flex items-center justify-between gap-4">
          <div className="flex items-center gap-3">
            <div className="w-3 h-8 bg-[#c8102e] rounded-[2px]" />
            <div>
              <div className="text-[15px] font-extrabold tracking-[0.06em] text-[#1a1a1a] uppercase">
                BHAVYA POLYFILMS · RoDTEP / RoSCTL SCRIP ENGINE
              </div>
              <div className="text-[11px] font-mono-code text-[#6b6863]">
                SAP HANA (BHAVYA_POLYFILMS_01) · G/L 370001 &amp; 410023 · On-The-Fly Balance
              </div>
            </div>
          </div>

          <nav className="flex items-center gap-2">
            <button
              type="button"
              onClick={() => setActiveTab("module1")}
              className={`px-4 py-2 rounded-[4px] text-[11.5px] font-bold uppercase tracking-[0.08em] transition-all cursor-pointer ${
                activeTab === "module1"
                  ? "bg-[#c8102e] text-white shadow-sm"
                  : "bg-[#f4f2ef] text-[#444444] hover:bg-[#eae6df]"
              }`}
            >
              1. AP Invoice &amp; Excel Paste
            </button>
            <button
              type="button"
              onClick={() => setActiveTab("module2")}
              className={`px-4 py-2 rounded-[4px] text-[11.5px] font-bold uppercase tracking-[0.08em] transition-all cursor-pointer ${
                activeTab === "module2"
                  ? "bg-[#c8102e] text-white shadow-sm"
                  : "bg-[#f4f2ef] text-[#444444] hover:bg-[#eae6df]"
              }`}
            >
              2. Attach Scrips to BL / GRPO ({pendingScrips.length})
            </button>
            <button
              type="button"
              onClick={() => setActiveTab("ledger")}
              className={`px-4 py-2 rounded-[4px] text-[11.5px] font-bold uppercase tracking-[0.08em] transition-all cursor-pointer ${
                activeTab === "ledger"
                  ? "bg-[#c8102e] text-white shadow-sm"
                  : "bg-[#f4f2ef] text-[#444444] hover:bg-[#eae6df]"
              }`}
            >
              3. Scrip Balance &amp; BL Ledger
            </button>
          </nav>
        </div>
      </header>

      <main className="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-6 space-y-6">
        {/* Global Live Balance KPI Strip */}
        <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
          <div className="rp-card p-3.5 border-l-4 border-l-[#c8102e]">
            <div className="text-[10px] font-bold uppercase tracking-[0.08em] text-[#c8102e]">
              G/L 370001 Recon Pending
            </div>
            <div className="text-[17px] font-extrabold font-mono-code text-[#1a1a1a] mt-1">
              ₹{formatINR(glReconPendingTotal)}
            </div>
            <div className="text-[10.5px] text-[#8a8780] mt-0.5">
              {glReconPendingRows.length} Open AP Invoices in SAP
            </div>
          </div>

          <div className="rp-card p-3.5">
            <div className="text-[10px] font-bold uppercase tracking-[0.08em] text-[#6b6863]">
              Scrips Entered (Table 1)
            </div>
            <div className="text-[17px] font-extrabold font-mono-code text-[#1a1a1a] mt-1">
              ₹{formatINR(globalKpis.totalPurchased)}
            </div>
            <div className="text-[10.5px] text-[#8a8780] mt-0.5">
              {globalKpis.scripCount} Licences Captured
            </div>
          </div>

          <div className="rp-card p-3.5">
            <div className="text-[10px] font-bold uppercase tracking-[0.08em] text-[#6b6863]">
              Utilized on BL / BOE (On-The-Fly)
            </div>
            <div className="text-[17px] font-extrabold font-mono-code text-amber-700 mt-1">
              ₹{formatINR(globalKpis.totalUtilized)}
            </div>
            <div className="text-[10.5px] text-[#8a8780] mt-0.5">
              {utilizations.length} BL Allocations
            </div>
          </div>

          <div className="rp-card p-3.5 border-l-4 border-l-emerald-600">
            <div className="text-[10px] font-bold uppercase tracking-[0.08em] text-emerald-800">
              Net Available Scrip Balance
            </div>
            <div className="text-[17px] font-extrabold font-mono-code text-emerald-700 mt-1">
              ₹{formatINR(globalKpis.totalBalance)}
            </div>
            <div className="text-[10.5px] text-[#8a8780] mt-0.5">
              {globalKpis.pendingCount} Pending Scrips Available
            </div>
          </div>
        </div>

        {/* Status / Alert Banner */}
        {banner && (
          <div
            className={`p-4 rounded-[6px] border flex items-center justify-between gap-4 text-[13px] font-medium ${
              banner.type === "error"
                ? "bg-red-50 border-red-300 text-red-900"
                : banner.type === "success"
                ? "bg-emerald-50 border-emerald-300 text-emerald-900"
                : "bg-blue-50 border-blue-200 text-blue-900"
            }`}
          >
            <span>{banner.text}</span>
            <button
              type="button"
              onClick={() => setBanner(null)}
              className="text-xs font-bold uppercase underline cursor-pointer shrink-0"
            >
              Dismiss
            </button>
          </div>
        )}

        {/* ===================================================================== */}
        {/* MODULE 1: PASTE EXCEL SCRIPS & CREATE SAP A/P INVOICE VIA SERVICE LAYER */}
        {/* ===================================================================== */}
        {activeTab === "module1" && (
          <form onSubmit={handleSubmitModule1} className="space-y-6">
            {/* TOP CARD: START FROM SCRATCH WITH EXCEL -> CREATE NEW SAP A/P INVOICE (OR LINK EXISTING) */}
            <div className="rp-card p-6 space-y-5 border-t-4 border-t-[#c8102e]">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-black/[0.08] pb-4">
                <div>
                  <span className="sec-label">Module 1 · Start Here with Your RoDTEP Excel Sheet</span>
                  <h2 className="text-xl font-extrabold text-[#1a1a1a] mt-1">
                    {apMode === "CREATE_AP"
                      ? "Start from Scratch: Paste RoDTEP Excel & Create New SAP A/P Invoice"
                      : apMode === "LINK_EXISTING"
                      ? "Link Pasted RoDTEP Excel Scrips to Existing SAP A/P Invoice"
                      : "Import Pending RoDTEP Scrips Only (Link SAP A/P Invoice Later)"}
                  </h2>
                </div>

                <div className="flex flex-wrap items-center gap-2">
                  <button
                    type="button"
                    onClick={() => setApMode("CREATE_AP")}
                    className={`px-3.5 py-2 rounded-[4px] text-[11px] font-bold uppercase tracking-wider cursor-pointer ${
                      apMode === "CREATE_AP"
                        ? "bg-[#c8102e] text-white shadow-sm"
                        : "bg-[#f4f2ef] text-[#555] hover:bg-[#e8e4dd]"
                    }`}
                  >
                    1. Create New SAP A/P Invoice (From Excel)
                  </button>
                  <button
                    type="button"
                    onClick={() => setApMode("LINK_EXISTING")}
                    className={`px-3.5 py-2 rounded-[4px] text-[11px] font-bold uppercase tracking-wider cursor-pointer ${
                      apMode === "LINK_EXISTING"
                        ? "bg-[#1a1a1a] text-white shadow-sm"
                        : "bg-[#f4f2ef] text-[#555] hover:bg-[#e8e4dd]"
                    }`}
                  >
                    2. Link Excel to Existing SAP AP Invoice
                  </button>
                  <button
                    type="button"
                    onClick={() => setApMode("SAVE_ONLY")}
                    className={`px-3.5 py-2 rounded-[4px] text-[11px] font-bold uppercase tracking-wider cursor-pointer ${
                      apMode === "SAVE_ONLY"
                        ? "bg-[#1a1a1a] text-white shadow-sm"
                        : "bg-[#f4f2ef] text-[#555] hover:bg-[#e8e4dd]"
                    }`}
                  >
                    3. Import Scrips Only (Link AP # Later)
                  </button>
                </div>
              </div>

              {/* A/P Invoice Header Fields */}
              {apMode === "CREATE_AP" ? (
                <div className="space-y-4">
                  <div className="bg-amber-50/80 border border-amber-200 p-3.5 rounded-[6px] text-[12.5px] text-amber-950">
                    <strong>How to enter a New A/P Invoice from scratch using only your RoDTEP Excel sheet:</strong>
                    <ol className="list-decimal list-inside mt-1 space-y-0.5 text-[12px]">
                      <li>
                        Select the <strong>Vendor</strong>, enter the <strong>Vendor Bill Ref (NumAtCard)</strong>, <strong>Posting Date</strong>, and <strong>Gain / Discount Amount (G/L 410023)</strong> below.
                      </li>
                      <li>
                        Copy the rows from your RoDTEP Scrips Excel sheet and paste (<strong>Ctrl+V</strong>) into the box below — the system automatically sums the Scrip Face Value for <strong>G/L 370001</strong>.
                      </li>
                      <li>
                        Click <strong>&ldquo;Create SAP A/P Invoice &amp; Save Scrips&rdquo;</strong> — it posts the new Service A/P Invoice in SAP B1 and automatically links all the pasted RoDTEP scrips to that new A/P Invoice in one step!
                      </li>
                    </ol>
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4 bg-[#f4f2ef]/70 p-4 rounded-[6px] border border-black/[0.08]">
                    <div className="lg:col-span-2">
                      <label className="block text-[11px] font-bold uppercase tracking-wider text-[#444] mb-1">
                        1. Select Vendor (CardCode) *
                      </label>
                      <select
                        value={selectedCardCode}
                        onChange={(e) => setSelectedCardCode(e.target.value)}
                        className="w-full px-3 py-2 bg-white border border-black/20 rounded-[4px] text-[13px] font-medium"
                        required
                      >
                        {vendors.map((v) => (
                          <option key={v.CardCode} value={v.CardCode}>
                            {v.CardCode} — {v.CardName} {v.IsScripVendor ? "★" : ""}
                          </option>
                        ))}
                      </select>
                    </div>

                    <div>
                      <label className="block text-[11px] font-bold uppercase tracking-wider text-[#444] mb-1">
                        2. Vendor Bill No (NumAtCard) *
                      </label>
                      <input
                        type="text"
                        value={numAtCard}
                        onChange={(e) => setNumAtCard(e.target.value)}
                        placeholder="e.g. SCRP15/11/26-27"
                        className="w-full px-3 py-2 bg-white border border-black/20 rounded-[4px] text-[13px] font-mono-code"
                      />
                    </div>

                    <div>
                      <label className="block text-[11px] font-bold uppercase tracking-wider text-[#444] mb-1">
                        3. Posting Date (DocDate) *
                      </label>
                      <input
                        type="date"
                        value={docDate}
                        onChange={(e) => setDocDate(e.target.value)}
                        className="w-full px-3 py-2 bg-white border border-black/20 rounded-[4px] text-[13px] font-mono-code"
                        required
                      />
                    </div>

                    <div>
                      <label className="block text-[11px] font-bold uppercase tracking-wider text-[#444] mb-1">
                        4. Gain / Discount (G/L 410023 ₹)
                      </label>
                      <div className="flex gap-1.5">
                        <input
                          type="number"
                          step="0.01"
                          value={discountPercentageOrAmtDisplay(discountPercent)}
                          onChange={(e) => {
                            const pct = e.target.value;
                            setDiscountPercent(pct);
                            if (pct && !isNaN(Number(pct)) && pastedSummary.total > 0) {
                              setDiscountAmount(
                                String(Math.round((pastedSummary.total * Number(pct)) / 100))
                              );
                            }
                          }}
                          placeholder="%"
                          title="Optional Discount % calculator"
                          className="w-16 px-2 py-2 bg-white border border-black/20 rounded-[4px] text-[12px] font-mono-code"
                        />
                        <input
                          type="number"
                          step="0.01"
                          value={discountAmount}
                          onChange={(e) => setDiscountAmount(e.target.value)}
                          placeholder="e.g. 20414"
                          className="flex-1 px-3 py-2 bg-white border border-black/20 rounded-[4px] text-[13px] font-mono-code"
                        />
                      </div>
                    </div>
                  </div>
                </div>
              ) : apMode === "LINK_EXISTING" ? (
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4 bg-[#f4f2ef]/70 p-4 rounded-[6px] border border-black/[0.08]">
                  <div>
                    <label className="block text-[11px] font-bold uppercase tracking-wider text-[#444] mb-1.5">
                      Select Existing SAP A/P Invoice with Unlinked Balance (Fully Linked Invoices Are Hidden)
                    </label>
                    <select
                      value={selectedExistingApDocEntry}
                      onChange={(e) => handleSelectExistingAp(e.target.value)}
                      className="w-full px-3 py-2.5 bg-white border border-black/20 rounded-[4px] text-[13px] font-mono-code"
                    >
                      <option value="">-- Choose SAP A/P Invoice to Link --</option>
                      {linkableApInvoices.map((inv) => {
                        const info = apLinkMap[`ENTRY_${inv.DocEntry}`];
                        const unlinked = info ? info.unlinkedBalance : Number(inv.InternalReconPending || inv.ScripFaceValue || 0);
                        return (
                          <option key={inv.DocEntry} value={inv.DocEntry}>
                            #{inv.DocNum} | {inv.DocDate} | {inv.CardName} | Ref: {inv.NumAtCard || "-"} | Unlinked Bal: ₹{formatINR(unlinked)}
                          </option>
                        );
                      })}
                    </select>
                  </div>

                  {selectedExistingApDocEntry && (
                    <div className="flex flex-col justify-center text-[12.5px] font-mono-code bg-white p-3 rounded border border-black/10">
                      <div>
                        <strong>Vendor:</strong> {selectedCardCode} | <strong>Ref:</strong>{" "}
                        {numAtCard || "-"} | <strong>Date:</strong> {docDate}
                      </div>
                      <div className="mt-1">
                        <strong>Gain/Discount (410023):</strong> -₹{formatINR(discountAmount)}
                      </div>
                    </div>
                  )}
                </div>
              ) : (
                <div className="bg-emerald-50/70 border border-emerald-200 p-3.5 rounded-[6px] text-[12.5px] text-emerald-950 flex flex-wrap items-center justify-between gap-2">
                  <div>
                    <strong>Quick Pending Import Mode:</strong> Paste your pending RoDTEP scrips below and save them directly. You can link them to an SAP A/P Invoice # anytime later.
                  </div>
                </div>
              )}

              {/* Excel Copy-Paste Area */}
              <div className="space-y-2">
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <label className="text-[11px] font-bold uppercase tracking-wider text-[#444]">
                    {apMode === "CREATE_AP"
                      ? "5. Copy & Paste RoDTEP Scrips from Excel (EXPORTER | LIC NO | DATED | SCHEME | PORT CODE | SCRIP AMOUNT)"
                      : "Copy & Paste RoDTEP Scrips from Excel (EXPORTER | LIC NO | DATED | SCHEME | PORT CODE | SCRIP AMOUNT)"}
                  </label>
                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      onClick={handleLoadAttachedExcelSample}
                      className="rp-btn rp-btn-outline !py-1.5 !px-3 !text-[10.5px]"
                    >
                      Load Attached 27-Scrip Excel Sheet (₹25,51,794)
                    </button>
                    {pastedScrips.length > 0 && (
                      <button
                        type="button"
                        onClick={() => {
                          setPastedScrips([]);
                          setRawPasteText("");
                        }}
                        className="rp-btn rp-btn-outline !py-1.5 !px-3 !text-[10.5px]"
                      >
                        Clear Table
                      </button>
                    )}
                  </div>
                </div>

                <textarea
                  rows={4}
                  value={rawPasteText}
                  onChange={(e) => handlePasteTextChange(e.target.value)}
                  placeholder="Click here and press Ctrl+V / Cmd+V after copying rows from your RoDTEP Excel sheet..."
                  className="w-full p-3 bg-[#faf9f7] border border-dashed border-black/30 rounded-[4px] text-[12px] font-mono-code focus:border-[#c8102e] focus:bg-white outline-none"
                />
              </div>

              {/* Pasted Scrips Preview & Editable Grid */}
              {pastedScrips.length > 0 && (
                <div className="space-y-4">
                  {/* Summary Bar */}
                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 bg-[#1c1c1c] text-white p-4 rounded-[6px]">
                    <div>
                      <div className="text-[10px] uppercase tracking-wider text-gray-400">
                        Line 1 · G/L 370001 (Face Value)
                      </div>
                      <div className="text-[16px] font-bold font-mono-code text-white mt-0.5">
                        ₹{formatINR(pastedSummary.total)}
                      </div>
                      <div className="text-[10px] text-gray-400">{pastedSummary.count} Scrips</div>
                    </div>

                    <div>
                      <div className="text-[10px] uppercase tracking-wider text-gray-400">
                        Line 2 · G/L 410023 (Gain/Disc)
                      </div>
                      <div className="text-[16px] font-bold font-mono-code text-amber-400 mt-0.5">
                        -₹{formatINR(pastedSummary.discount)}
                      </div>
                      <div className="text-[10px] text-gray-400">TaxCode: EXEMPTED</div>
                    </div>

                    <div>
                      <div className="text-[10px] uppercase tracking-wider text-gray-400">
                        Net AP Invoice Total
                      </div>
                      <div className="text-[16px] font-bold font-mono-code text-emerald-400 mt-0.5">
                        ₹{formatINR(pastedSummary.netPayable)}
                      </div>
                      <div className="text-[10px] text-gray-400">Series 239 (Service)</div>
                    </div>
                  </div>

                  {/* Table */}
                  <div className="overflow-x-auto border border-black/10 rounded-[6px] max-h-[420px]">
                    <table className="w-full text-left border-collapse text-[12px]">
                      <thead className="bg-[#f4f2ef] sticky top-0 z-10 border-b border-black/10 text-[10.5px] font-bold uppercase tracking-wider text-[#555]">
                        <tr>
                          <th className="py-2.5 px-3">#</th>
                          <th className="py-2.5 px-3">Exporter</th>
                          <th className="py-2.5 px-3">Lic No</th>
                          <th className="py-2.5 px-3">Dated</th>
                          <th className="py-2.5 px-3">Scheme</th>
                          <th className="py-2.5 px-3">Port Code</th>
                          <th className="py-2.5 px-3 text-right">Scrip Amount (₹)</th>
                          {apMode === "SAVE_ONLY" && (
                            <th className="py-2.5 px-3">SAP AP # (Optional / Later)</th>
                          )}
                          <th className="py-2.5 px-3 text-center">Del</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-black/[0.06] bg-white">
                        {pastedScrips.map((row, idx) => {
                          const alreadyExists = existingLicSet.has(String(row.LicNo).trim());
                          return (
                            <tr key={idx} className="hover:bg-[#faf9f7]">
                              <td className="py-2 px-3 font-mono-code text-[#888]">{idx + 1}</td>
                              <td className="py-2 px-3">
                                <input
                                  type="text"
                                  value={row.ExporterName}
                                  onChange={(e) =>
                                    handleUpdatePastedRow(idx, "ExporterName", e.target.value)
                                  }
                                  className="w-full px-2 py-1 border border-transparent hover:border-black/15 focus:border-black/30 rounded text-[12px]"
                                />
                              </td>
                              <td className="py-2 px-3">
                                <div className="flex items-center gap-1.5">
                                  <input
                                    type="text"
                                    value={row.LicNo}
                                    onChange={(e) =>
                                      handleUpdatePastedRow(idx, "LicNo", e.target.value)
                                    }
                                    className="w-32 px-2 py-1 border border-transparent hover:border-black/15 focus:border-black/30 rounded font-mono-code font-bold text-[12px]"
                                  />
                                  {alreadyExists && (
                                    <span className="px-1.5 py-0.5 text-[9.5px] font-bold bg-amber-100 text-amber-800 rounded">
                                      Saved
                                    </span>
                                  )}
                                </div>
                              </td>
                              <td className="py-2 px-3 font-mono-code">{row.LicDate}</td>
                              <td className="py-2 px-3">
                                <span
                                  className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                                    row.Scheme === "ROSCTL"
                                      ? "bg-amber-100 text-amber-900"
                                      : "bg-emerald-100 text-emerald-900"
                                  }`}
                                >
                                  {row.Scheme}
                                </span>
                              </td>
                              <td className="py-2 px-3 font-mono-code">{row.PortCode}</td>
                              <td className="py-2 px-3 text-right">
                                <input
                                  type="number"
                                  step="0.01"
                                  value={row.ScripAmount}
                                  onChange={(e) =>
                                    handleUpdatePastedRow(idx, "ScripAmount", e.target.value)
                                  }
                                  className="w-28 px-2 py-1 text-right border border-black/15 rounded font-mono-code font-bold text-[12px]"
                                />
                              </td>
                              {apMode === "SAVE_ONLY" && (
                                <td className="py-2 px-3">
                                  <select
                                    value={row.APDocEntry || ""}
                                    onChange={(e) => {
                                      const entry = e.target.value;
                                      const found = apInvoices.find((a) => String(a.DocEntry) === String(entry));
                                      setPastedScrips((prev) => {
                                        const copy = [...prev];
                                        copy[idx] = {
                                          ...copy[idx],
                                          APDocEntry: found ? found.DocEntry : null,
                                          APDocNum: found ? found.DocNum : null,
                                          CardCode: found ? found.CardCode : "",
                                          CardName: found ? found.CardName : "",
                                          NumAtCard: found ? found.NumAtCard : "",
                                          APDocDate: found ? found.DocDate : ""
                                        };
                                        return copy;
                                      });
                                    }}
                                    className="px-2 py-1 bg-white border border-black/15 rounded font-mono-code text-[11px] max-w-[180px]"
                                  >
                                    <option value="">-- Link AP Later --</option>
                                    {linkableApInvoices.map((inv) => (
                                      <option key={inv.DocEntry} value={inv.DocEntry}>
                                        #{inv.DocNum} ({inv.NumAtCard || inv.DocDate}) · Pending ₹{formatINR(inv.InternalReconPending)}
                                      </option>
                                    ))}
                                  </select>
                                </td>
                              )}
                              <td className="py-2 px-3 text-center">
                                <button
                                  type="button"
                                  onClick={() => handleRemovePastedRow(idx)}
                                  className="text-red-600 hover:text-red-800 font-bold px-1.5 cursor-pointer"
                                  title="Remove row"
                                >
                                  ×
                                </button>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>

                  {/* Submit Action */}
                  <div className="flex items-center justify-end gap-3 pt-2">
                    <button
                      type="submit"
                      disabled={submittingAp}
                      className="rp-btn rp-btn-red !py-3 !px-6 !text-[12px]"
                    >
                      {submittingAp
                        ? "Posting to SAP..."
                        : apMode === "CREATE_AP"
                        ? `Create SAP A/P Invoice (₹${formatINR(pastedSummary.netPayable)}) & Save ${pastedSummary.count} Linked Scrips`
                        : apMode === "SAVE_ONLY"
                        ? `Save ${pastedSummary.count} Pending Scrips (₹${formatINR(pastedSummary.total)}) — Link AP Invoice Later`
                        : `Link ${pastedSummary.count} Scrips (₹${formatINR(pastedSummary.total)}) to Selected A/P Invoice`}
                    </button>
                  </div>
                </div>
              )}
            </div>

            {/* SECOND CARD: SAP A/P INVOICES & EXPANDABLE LINKED RoDTEP SCRIPS */}
            <div className="rp-card p-5 space-y-3">
              <div className="flex flex-wrap items-center justify-between gap-3 border-b border-black/[0.08] pb-3">
                <div>
                  <span className="sec-label">SAP HANA Live · G/L 370001 (RODTEP Scrip Account) &amp; Linked Scrips</span>
                  <h3 className="text-[15px] font-extrabold text-[#1a1a1a]">
                    {glListViewMode === "ALL_AP"
                      ? `Complete RoDTEP A/P Invoice List (${glAllApRows.length} Invoices · Total ₹${formatINR(glTotalPurchasedFromStart)})`
                      : glListViewMode === "PENDING_ONLY"
                      ? `Internal Reconciliation Pending A/P Invoices (${glReconPendingRows.length} Open · Pending ₹${formatINR(glReconPendingTotal)})`
                      : `All BOE Duty Utilization Credit JEs (${glAllBoeCreditRows.length} Entries)`}
                  </h3>
                </div>

                <div className="flex flex-wrap items-center gap-1.5">
                  <button
                    type="button"
                    onClick={() => setGlListViewMode("LINKED_AP")}
                    className={`px-3 py-1.5 rounded-[4px] text-[10.5px] font-bold uppercase tracking-wider cursor-pointer ${
                      glListViewMode === "LINKED_AP"
                        ? "bg-emerald-700 text-white"
                        : "bg-emerald-50 text-emerald-900 border border-emerald-300 hover:bg-emerald-100"
                    }`}
                  >
                    ✓ AP Invoices with Linked RoDTEP ({
                      glAllApRows.filter((r) =>
                        scrips.some(
                          (s) =>
                            (r.APDocEntry && Number(s.APDocEntry) === Number(r.APDocEntry)) ||
                            (r.APDocNum && Number(s.APDocNum) === Number(r.APDocNum))
                        )
                      ).length
                    })
                  </button>
                  <button
                    type="button"
                    onClick={() => setGlListViewMode("PENDING_ONLY")}
                    className={`px-3 py-1.5 rounded-[4px] text-[10.5px] font-bold uppercase tracking-wider cursor-pointer ${
                      glListViewMode === "PENDING_ONLY"
                        ? "bg-[#c8102e] text-white"
                        : "bg-[#f4f2ef] text-[#444] hover:bg-[#e5e1da]"
                    }`}
                  >
                    Pending Recon Only ({glReconPendingRows.length})
                  </button>
                  <button
                    type="button"
                    onClick={() => setGlListViewMode("ALL_AP")}
                    className={`px-3 py-1.5 rounded-[4px] text-[10.5px] font-bold uppercase tracking-wider cursor-pointer ${
                      glListViewMode === "ALL_AP"
                        ? "bg-[#c8102e] text-white"
                        : "bg-[#f4f2ef] text-[#444] hover:bg-[#e5e1da]"
                    }`}
                  >
                    Complete AP List ({glAllApRows.length})
                  </button>
                  <button
                    type="button"
                    onClick={() => setGlListViewMode("BOE_CREDITS")}
                    className={`px-3 py-1.5 rounded-[4px] text-[10.5px] font-bold uppercase tracking-wider cursor-pointer ${
                      glListViewMode === "BOE_CREDITS"
                        ? "bg-[#c8102e] text-white"
                        : "bg-[#f4f2ef] text-[#444] hover:bg-[#e5e1da]"
                    }`}
                  >
                    All BOE Credit JEs ({glAllBoeCreditRows.length})
                  </button>
                </div>
              </div>

              <div className="overflow-x-auto max-h-[520px] border border-black/10 rounded-[4px]">
                {glListViewMode === "BOE_CREDITS" ? (
                  <table className="w-full text-left border-collapse text-[12px]">
                    <thead className="bg-[#f4f2ef] sticky top-0 z-10 border-b border-black/10 text-[10.5px] font-bold uppercase tracking-wider text-[#555]">
                      <tr>
                        <th className="py-2 px-3">#</th>
                        <th className="py-2 px-3">JE #</th>
                        <th className="py-2 px-3">Date</th>
                        <th className="py-2 px-3">BOE / Duty Memo (LineMemo)</th>
                        <th className="py-2 px-3 text-right">Credit Utilized (₹)</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-black/[0.06] bg-white">
                      {glAllBoeCreditRows.map((r, idx) => (
                        <tr key={`${r.TransId}-${r.Line_ID}`} className="hover:bg-[#faf9f7]">
                          <td className="py-2 px-3 font-mono-code text-[#888]">{idx + 1}</td>
                          <td className="py-2 px-3 font-mono-code font-bold">#{r.BaseRef}</td>
                          <td className="py-2 px-3 font-mono-code">{r.RefDate}</td>
                          <td className="py-2 px-3 font-medium">{r.LineMemo || "-"}</td>
                          <td className="py-2 px-3 text-right font-mono-code font-bold text-amber-800">
                            ₹{formatINR(r.OriginalCredit)}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                ) : (
                  <table className="w-full text-left border-collapse text-[12px]">
                    <thead className="bg-[#f4f2ef] sticky top-0 z-10 border-b border-black/10 text-[10.5px] font-bold uppercase tracking-wider text-[#555]">
                      <tr>
                        <th className="py-2 px-3">#</th>
                        <th className="py-2 px-3">AP #</th>
                        <th className="py-2 px-3">Date</th>
                        <th className="py-2 px-3">Vendor</th>
                        <th className="py-2 px-3">Vendor Ref (NumAtCard)</th>
                        <th className="py-2 px-3 text-right">Original 370001 Debit</th>
                        <th className="py-2 px-3 text-right">Pending in SAP (BalDueDeb)</th>
                        <th className="py-2 px-3 text-right">Linked Scrips Total</th>
                        <th className="py-2 px-3 text-right">Unlinked Balance</th>
                        <th className="py-2 px-3 text-center">Linked Scrips / Action</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-black/[0.06] bg-white">
                      {(glListViewMode === "ALL_AP"
                        ? glAllApRows
                        : glListViewMode === "LINKED_AP"
                        ? glAllApRows.filter((r) =>
                            scrips.some(
                              (s) =>
                                (r.APDocEntry && Number(s.APDocEntry) === Number(r.APDocEntry)) ||
                                (r.APDocNum && Number(s.APDocNum) === Number(r.APDocNum))
                            )
                          )
                        : glReconPendingRows
                      ).map((r, idx) => {
                        const rowKey = `${r.TransId}-${r.Line_ID}`;
                        const linkedList = scrips.filter(
                          (s) =>
                            (r.APDocEntry && Number(s.APDocEntry) === Number(r.APDocEntry)) ||
                            (r.APDocNum && Number(s.APDocNum) === Number(r.APDocNum))
                        );
                        const enteredForAp = Number(
                          linkedList.reduce((sum, s) => sum + Number(s.ScripAmount || 0), 0).toFixed(2)
                        );
                        const origDebit = Number(r.OriginalDebit || 0);
                        const pendingRecon = Number(r.PendingReconBalance || 0);
                        const targetAmt = pendingRecon > 0 ? pendingRecon : origDebit;
                        const unlinkedBalance = Math.max(0, Number((targetAmt - enteredForAp).toFixed(2)));
                        const isFullyLinked = enteredForAp > 0 && unlinkedBalance <= 0.5;
                        const isExpanded = expandedApDocs[rowKey] !== undefined ? expandedApDocs[rowKey] : glListViewMode === "LINKED_AP" && linkedList.length > 0;

                        return (
                          <React.Fragment key={rowKey}>
                            <tr className={`hover:bg-[#faf9f7] ${isExpanded ? "bg-blue-50/40" : ""}`}>
                              <td className="py-2.5 px-3 font-mono-code text-[#888]">{idx + 1}</td>
                              <td className="py-2.5 px-3 font-mono-code font-bold text-[#c8102e]">
                                #{r.APDocNum || r.BaseRef}
                              </td>
                              <td className="py-2.5 px-3 font-mono-code">{r.RefDate}</td>
                              <td className="py-2.5 px-3 font-medium">
                                {r.CardName} <span className="text-[11px] text-[#777]">({r.CardCode})</span>
                              </td>
                              <td className="py-2.5 px-3 font-mono-code">{r.NumAtCard || "-"}</td>
                              <td className="py-2.5 px-3 text-right font-mono-code font-semibold">
                                ₹{formatINR(origDebit)}
                              </td>
                              <td
                                className={`py-2.5 px-3 text-right font-mono-code font-bold ${
                                  pendingRecon > 0 ? "text-emerald-700" : "text-[#999]"
                                }`}
                              >
                                ₹{formatINR(pendingRecon)}
                              </td>
                              <td className="py-2.5 px-3 text-right font-mono-code">
                                {enteredForAp > 0 ? (
                                  <span className="text-blue-700 font-bold">
                                    ₹{formatINR(enteredForAp)} ({linkedList.length})
                                  </span>
                                ) : (
                                  <span className="text-[#999]">₹0.00</span>
                                )}
                              </td>
                              <td className="py-2.5 px-3 text-right font-mono-code font-bold">
                                {isFullyLinked ? (
                                  <span className="text-emerald-700">₹0.00 ✓</span>
                                ) : (
                                  <span className="text-amber-700">₹{formatINR(unlinkedBalance)}</span>
                                )}
                              </td>
                              <td className="py-2.5 px-3 text-center">
                                <div className="flex items-center justify-center gap-1.5 flex-wrap">
                                  {linkedList.length > 0 && (
                                    <button
                                      type="button"
                                      onClick={() => toggleExpandAp(rowKey)}
                                      className="px-2.5 py-1 bg-blue-600 hover:bg-blue-700 text-white rounded text-[10.5px] font-bold uppercase tracking-wider cursor-pointer transition-colors"
                                    >
                                      {isExpanded ? `Hide Scrips (${linkedList.length}) ▲` : `View Scrips (${linkedList.length}) ▼`}
                                    </button>
                                  )}

                                  {isFullyLinked ? (
                                    <span className="px-2 py-0.5 bg-emerald-100 text-emerald-800 border border-emerald-300 rounded text-[10px] font-bold uppercase">
                                      ✓ Linked
                                    </span>
                                  ) : (
                                    r.APDocEntry && (
                                      <button
                                        type="button"
                                        onClick={() => {
                                          setApMode("LINK_EXISTING");
                                          handleSelectExistingAp(String(r.APDocEntry));
                                          window.scrollTo({ top: 0, behavior: "smooth" });
                                        }}
                                        className="px-2.5 py-1 bg-[#1a1a1a] hover:bg-[#c8102e] text-white rounded text-[10.5px] font-bold uppercase tracking-wider cursor-pointer transition-colors"
                                      >
                                        + Link Scrips
                                      </button>
                                    )
                                  )}
                                </div>
                              </td>
                            </tr>

                            {/* Expandable Child Table showing all Linked RoDTEP Scrips for this A/P Invoice */}
                            {isExpanded && linkedList.length > 0 && (
                              <tr className="bg-[#f8fafc]">
                                <td colSpan={10} className="p-3 border-l-4 border-l-blue-600">
                                  <div className="text-[11px] font-bold uppercase tracking-wider text-blue-900 mb-2 flex items-center justify-between">
                                    <span>
                                      Linked RoDTEP / RoSCTL Scrips under SAP A/P Invoice #{r.APDocNum || r.BaseRef} ({r.NumAtCard || r.CardName}) — {linkedList.length} Scrips Totaling ₹{formatINR(enteredForAp)}
                                    </span>
                                  </div>
                                  <div className="overflow-x-auto border border-blue-200 rounded bg-white">
                                    <table className="w-full text-left border-collapse text-[11.5px]">
                                      <thead className="bg-blue-50 text-[10px] font-bold uppercase tracking-wider text-blue-950 border-b border-blue-200">
                                        <tr>
                                          <th className="py-1.5 px-2.5">#</th>
                                          <th className="py-1.5 px-2.5">RoDTEP Lic No</th>
                                          <th className="py-1.5 px-2.5">Dated</th>
                                          <th className="py-1.5 px-2.5">Exporter</th>
                                          <th className="py-1.5 px-2.5">Scheme</th>
                                          <th className="py-1.5 px-2.5">Port Code</th>
                                          <th className="py-1.5 px-2.5 text-right">Scrip Amount (₹)</th>
                                          <th className="py-1.5 px-2.5 text-right">Utilized on BL (₹)</th>
                                          <th className="py-1.5 px-2.5 text-right">Available Balance (₹)</th>
                                        </tr>
                                      </thead>
                                      <tbody className="divide-y divide-blue-100">
                                        {linkedList.map((sc, sIdx) => (
                                          <tr key={sc.LicNo} className="hover:bg-blue-50/30">
                                            <td className="py-1.5 px-2.5 font-mono-code text-[#777]">{sIdx + 1}</td>
                                            <td className="py-1.5 px-2.5 font-mono-code font-bold text-[#1a1a1a]">{sc.LicNo}</td>
                                            <td className="py-1.5 px-2.5 font-mono-code">{sc.LicDate || "-"}</td>
                                            <td className="py-1.5 px-2.5">{sc.ExporterName}</td>
                                            <td className="py-1.5 px-2.5 font-bold text-[10.5px]">{sc.Scheme}</td>
                                            <td className="py-1.5 px-2.5 font-mono-code">{sc.PortCode}</td>
                                            <td className="py-1.5 px-2.5 text-right font-mono-code font-semibold">
                                              ₹{formatINR(sc.ScripAmount)}
                                            </td>
                                            <td className="py-1.5 px-2.5 text-right font-mono-code text-amber-700">
                                              ₹{formatINR(sc.UtilizedAmount)}
                                            </td>
                                            <td className="py-1.5 px-2.5 text-right font-mono-code font-bold text-emerald-700">
                                              ₹{formatINR(sc.BalanceAmount)}
                                            </td>
                                          </tr>
                                        ))}
                                      </tbody>
                                    </table>
                                  </div>
                                </td>
                              </tr>
                            )}
                          </React.Fragment>
                        );
                      })}
                    </tbody>
                  </table>
                )}
              </div>
            </div>
          </form>
        )}

        {/* ===================================================================== */}
        {/* MODULE 2: ATTACH SCRIPS TO BL / GRPO WITH PORT LOCK & DRAFT BOE CHECK */}
        {/* ===================================================================== */}
        {activeTab === "module2" && (
          <form onSubmit={handleSubmitModule2} className="space-y-6">
            <div className="rp-card p-6 space-y-5">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-black/[0.08] pb-4">
                <div>
                  <span className="sec-label">Module 2 · Table 2 (BHV_RODTEP_BL)</span>
                  <h2 className="text-xl font-extrabold text-[#1a1a1a] mt-1">
                    Attach Pending RoDTEP Scrips to BL / GRPO &amp; Verify Draft BOE Amount
                  </h2>
                </div>


              </div>

              {/* Main Header Form: GRPO Selection, BL Number, Draft BOE Amount */}
              <div className="grid grid-cols-1 md:grid-cols-4 gap-4 bg-[#f4f2ef]/80 p-4 rounded-[6px] border border-black/[0.08]">
                <div className="md:col-span-2">
                  <label className="block text-[11px] font-bold uppercase tracking-wider text-[#444] mb-1">
                    1. Select Pending Import GRPO (Already-Attached GRPOs Are Hidden)
                  </label>
                  <select
                    value={selectedGrpoDocEntry}
                    onChange={(e) => handleSelectGrpo(e.target.value)}
                    className="w-full px-3 py-2 bg-white border border-black/20 rounded-[4px] text-[12.5px] font-mono-code"
                  >
                    <option value="">-- Select Unattached Import GRPO (or enter BL manually) --</option>
                    {unattachedGrpos.map((g) => (
                      <option key={g.GRPODocEntry} value={g.GRPODocEntry}>
                        GRPO #{g.GRPODocNum} ({g.GRPODate}) | BL: {g.BLNo || "N/A"} |{" "}
                        {g.BOENo ? `BOE: ${g.BOENo}` : "BOE Pending (Draft)"} | {g.CardName}
                      </option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-[11px] font-bold uppercase tracking-wider text-[#444] mb-1">
                    2. BL Number *
                  </label>
                  <input
                    type="text"
                    list="sap-hbl-list"
                    value={blNo}
                    onChange={(e) => {
                      const val = e.target.value;
                      setBlNo(val);
                      const matchedHbl = hbls.find(
                        (h) => String(h.BLNo).toUpperCase() === val.trim().toUpperCase()
                      );
                      if (matchedHbl && matchedHbl.GRPODocEntry) {
                        setSelectedGrpoDocEntry(String(matchedHbl.GRPODocEntry));
                        setSelectedGrpoDocNum(String(matchedHbl.GRPODocNum || ""));
                      }
                    }}
                    placeholder="Enter or pick BL No..."
                    className="w-full px-3 py-2 bg-white border border-black/20 rounded-[4px] text-[13px] font-mono-code font-bold"
                    required
                  />
                  <datalist id="sap-hbl-list">
                    {hbls.map((h) => (
                      <option key={h.HBLDocEntry} value={h.BLNo}>
                        {h.GRPODocNum ? `GRPO #${h.GRPODocNum}` : ""} {h.CardName || ""}
                      </option>
                    ))}
                  </datalist>
                </div>

                <div>
                  <label className="block text-[11px] font-bold uppercase tracking-wider text-[#444] mb-1">
                    3. Draft BOE / Duty Amount (₹) *
                  </label>
                  <div className="flex gap-1.5">
                    <input
                      type="number"
                      step="0.01"
                      value={draftBoeAmount}
                      onChange={(e) => setDraftBoeAmount(e.target.value)}
                      placeholder="e.g. 545000"
                      className="w-full px-3 py-2 bg-white border border-black/20 rounded-[4px] text-[13px] font-mono-code font-bold"
                      required
                    />
                    <button
                      type="button"
                      onClick={handleAutoAllocateDraftBoe}
                      className="px-2.5 py-2 bg-[#1a1a1a] text-white rounded-[4px] text-[10.5px] font-bold uppercase tracking-wider shrink-0 cursor-pointer hover:bg-black"
                      title="Automatically check oldest matching port scrips and adjust partial amount on last scrip to match Draft BOE Amount"
                    >
                      Auto-Pick
                    </button>
                  </div>
                </div>
              </div>

              {/* Live Validation Bar: Draft BOE Amount vs. Scrips Utilized */}
              <div
                className={`p-4 rounded-[6px] border flex flex-wrap items-center justify-between gap-4 ${
                  allocationSummary.targetBoe > 0 && allocationSummary.isExactMatch
                    ? "bg-emerald-950 text-white border-emerald-700"
                    : "bg-[#1c1c1c] text-white border-black"
                }`}
              >
                <div className="flex flex-wrap items-center gap-6">
                  <div>
                    <div className="text-[10px] uppercase tracking-wider text-gray-400">
                      Selected BL &amp; GRPO
                    </div>
                    <div className="text-[14px] font-bold font-mono-code">
                      {blNo || "—"}
                      {selectedGrpoDocNum ? ` · GRPO #${selectedGrpoDocNum}` : ""}
                    </div>
                  </div>

                  <div>
                    <div className="text-[10px] uppercase tracking-wider text-gray-400">
                      Draft BOE Target Amount
                    </div>
                    <div className="text-[16px] font-extrabold font-mono-code text-white">
                      ₹{formatINR(allocationSummary.targetBoe)}
                    </div>
                  </div>

                  <div>
                    <div className="text-[10px] uppercase tracking-wider text-gray-400">
                      Checked Scrips Utilized ({allocationSummary.checkedCount})
                    </div>
                    <div className="text-[16px] font-extrabold font-mono-code text-amber-300">
                      ₹{formatINR(allocationSummary.totalUtilized)}
                    </div>
                  </div>

                  <div>
                    <div className="text-[10px] uppercase tracking-wider text-gray-400">
                      Variance (Draft BOE − Utilized)
                    </div>
                    <div
                      className={`text-[16px] font-extrabold font-mono-code ${
                        allocationSummary.isExactMatch
                          ? "text-emerald-400"
                          : "text-red-400"
                      }`}
                    >
                      ₹{formatINR(allocationSummary.diff)}{" "}
                      {allocationSummary.isExactMatch && "✓ MATCHED"}
                    </div>
                  </div>
                </div>

                <div className="flex items-center gap-2">
                  {Object.keys(checkedScrips).length > 0 && (
                    <button
                      type="button"
                      onClick={() => setCheckedScrips({})}
                      className="px-3 py-2 rounded bg-white/10 hover:bg-white/20 text-white text-[11px] font-bold uppercase cursor-pointer"
                    >
                      Uncheck All
                    </button>
                  )}
                  <button
                    type="submit"
                    disabled={submittingBl || allocationSummary.checkedCount === 0}
                    className="rp-btn rp-btn-red !py-2.5 !px-5 !text-[11.5px]"
                  >
                    {submittingBl
                      ? "Saving BL Attachment..."
                      : `Attach ${allocationSummary.checkedCount} Scrip(s) to BL`}
                  </button>
                </div>
              </div>

              {/* Search Bar for Pending RoDTEP Scrips */}
              <div className="flex flex-wrap items-center justify-between gap-3">
                <div className="relative flex-1 min-w-[260px]">
                  <input
                    type="text"
                    value={module2Search}
                    onChange={(e) => setModule2Search(e.target.value)}
                    placeholder="Search pending scrips by RoDTEP Lic No, Exporter, Amount, Port, or AP # (No AP Invoice selection needed)..."
                    className="w-full px-3.5 py-2 bg-white border border-black/20 rounded-[4px] text-[12.5px] font-mono-code"
                  />
                </div>
                <div className="text-[11.5px] text-[#666] font-mono-code">
                  Showing <strong>{displayedPendingScrips.length}</strong> of <strong>{pendingScrips.length}</strong> pending scrips
                </div>
              </div>

              {/* Pending RoDTEP Scrips Table with Checkbox & Editable Amount */}
              <div className="overflow-x-auto border border-black/10 rounded-[6px] max-h-[480px]">
                <table className="w-full text-left border-collapse text-[12.5px]">
                  <thead className="bg-[#f4f2ef] sticky top-0 z-10 border-b border-black/10 text-[10.5px] font-bold uppercase tracking-wider text-[#555]">
                    <tr>
                      <th className="py-3 px-3 text-center w-14">Check</th>
                      <th className="py-3 px-3">RoDTEP Lic No</th>
                      <th className="py-3 px-3">Dated</th>
                      <th className="py-3 px-3">Scheme</th>
                      <th className="py-3 px-3">Exporter</th>
                      <th className="py-3 px-3">AP Inv #</th>
                      <th className="py-3 px-3 text-right">Original Amt</th>
                      <th className="py-3 px-3 text-right">Available Balance</th>
                      <th className="py-3 px-3 text-right w-44">Amount to Use (₹)</th>
                      <th className="py-3 px-3 text-right">Balance Left After</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-black/[0.06] bg-white">
                    {displayedPendingScrips.length === 0 ? (
                      <tr>
                        <td colSpan={10} className="py-8 text-center text-[#8a8780]">
                          No pending RoDTEP scrips found. Go to Tab 1 to paste your Excel scrips.
                        </td>
                      </tr>
                    ) : (
                      displayedPendingScrips.map((scrip) => {
                        const isChecked = checkedScrips[scrip.LicNo] !== undefined;
                        const currentUsed = isChecked ? Number(checkedScrips[scrip.LicNo] || 0) : 0;
                        const availBal = Number(scrip.BalanceAmount || 0);
                        const afterBal = Math.max(0, Number((availBal - currentUsed).toFixed(2)));

                        return (
                          <tr
                            key={scrip.LicNo}
                            onClick={() => {
                              if (!isChecked) {
                                handleToggleScripCheckbox(scrip);
                              }
                            }}
                            className={`transition-colors ${
                              isChecked
                                ? "bg-emerald-50/70"
                                : "hover:bg-[#faf9f7] cursor-pointer"
                            }`}
                          >
                            <td
                              className="py-2.5 px-3 text-center"
                              onClick={(e) => e.stopPropagation()}
                            >
                              <input
                                type="checkbox"
                                checked={isChecked}
                                onChange={() => handleToggleScripCheckbox(scrip)}
                                title="Check to attach this RoDTEP scrip"
                                className="w-4 h-4 accent-[#c8102e] cursor-pointer"
                              />
                            </td>

                            <td className="py-2.5 px-3 font-mono-code font-bold text-[#1a1a1a]">
                              {scrip.LicNo}
                            </td>
                            <td className="py-2.5 px-3 font-mono-code text-[12px]">
                              {scrip.LicDate || "-"}
                            </td>
                            <td className="py-2.5 px-3">
                              <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-[#f4f2ef] text-[#333]">
                                {scrip.Scheme} · {scrip.PortCode}
                              </span>
                            </td>
                            <td className="py-2.5 px-3 text-[12px] truncate max-w-[190px]">
                              {scrip.ExporterName}
                            </td>
                            <td className="py-2.5 px-3 font-mono-code text-[12px]">
                              {scrip.APDocNum ? `#${scrip.APDocNum}` : "-"}
                            </td>
                            <td className="py-2.5 px-3 text-right font-mono-code text-[#666]">
                              ₹{formatINR(scrip.ScripAmount)}
                            </td>
                            <td className="py-2.5 px-3 text-right font-mono-code font-bold text-[#1a1a1a]">
                              ₹{formatINR(availBal)}
                            </td>
                            <td
                              className="py-2.5 px-3 text-right"
                              onClick={(e) => e.stopPropagation()}
                            >
                              {isChecked ? (
                                <input
                                  type="number"
                                  step="0.01"
                                  min="0.01"
                                  max={availBal}
                                  value={checkedScrips[scrip.LicNo]}
                                  onChange={(e) =>
                                    handleChangeScripUsedAmount(scrip, e.target.value)
                                  }
                                  className="w-36 px-2.5 py-1 text-right bg-white border-2 border-[#c8102e] rounded font-mono-code font-bold text-[13px] text-[#1a1a1a]"
                                />
                              ) : (
                                <span className="font-mono-code text-[#aaa]">—</span>
                              )}
                            </td>
                            <td className="py-2.5 px-3 text-right font-mono-code font-semibold">
                              {isChecked ? (
                                <span
                                  className={
                                    afterBal === 0 ? "text-gray-500" : "text-amber-700 font-bold"
                                  }
                                >
                                  ₹{formatINR(afterBal)}
                                </span>
                              ) : (
                                <span className="text-[#888]">₹{formatINR(availBal)}</span>
                              )}
                            </td>
                          </tr>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>
            </div>

            {/* SECOND CARD IN MODULE 2: CHECK GRPOs / BLs AGAINST WHICH ATTACHMENT HAS ALREADY HAPPENED */}
            <div className="rp-card p-6 space-y-4">
              <div className="flex flex-wrap items-center justify-between gap-3 border-b border-black/[0.08] pb-3">
                <div>
                  <span className="sec-label">GRPO &amp; BL Attachment Tracker · Parent–Child View</span>
                  <h3 className="text-[16px] font-extrabold text-[#1a1a1a] mt-0.5">
                    {grpoListViewMode === "ATTACHED_ONLY"
                      ? `GRPOs / BLs with Attached RoDTEP Scrips (${grpoAttachmentRows.filter((r) => r.isAttached).length})`
                      : grpoListViewMode === "PENDING_ONLY"
                      ? `Pending Import GRPOs Without Scrip Attachment (${grpoAttachmentRows.filter((r) => !r.isAttached).length})`
                      : `All Import GRPOs & BLs (${grpoAttachmentRows.length})`}
                  </h3>
                </div>

                <div className="flex flex-wrap items-center gap-1.5">
                  <button
                    type="button"
                    onClick={() => setGrpoListViewMode("ATTACHED_ONLY")}
                    className={`px-3 py-1.5 rounded-[4px] text-[10.5px] font-bold uppercase tracking-wider cursor-pointer ${
                      grpoListViewMode === "ATTACHED_ONLY"
                        ? "bg-[#c8102e] text-white"
                        : "bg-[#f4f2ef] text-[#444] hover:bg-[#e5e1da]"
                    }`}
                  >
                    ✓ Attached GRPOs / BLs ({grpoAttachmentRows.filter((r) => r.isAttached).length})
                  </button>
                  <button
                    type="button"
                    onClick={() => setGrpoListViewMode("PENDING_ONLY")}
                    className={`px-3 py-1.5 rounded-[4px] text-[10.5px] font-bold uppercase tracking-wider cursor-pointer ${
                      grpoListViewMode === "PENDING_ONLY"
                        ? "bg-[#c8102e] text-white"
                        : "bg-[#f4f2ef] text-[#444] hover:bg-[#e5e1da]"
                    }`}
                  >
                    Pending GRPOs ({grpoAttachmentRows.filter((r) => !r.isAttached).length})
                  </button>
                  <button
                    type="button"
                    onClick={() => setGrpoListViewMode("ALL_GRPO")}
                    className={`px-3 py-1.5 rounded-[4px] text-[10.5px] font-bold uppercase tracking-wider cursor-pointer ${
                      grpoListViewMode === "ALL_GRPO"
                        ? "bg-[#c8102e] text-white"
                        : "bg-[#f4f2ef] text-[#444] hover:bg-[#e5e1da]"
                    }`}
                  >
                    All Import GRPOs ({grpoAttachmentRows.length})
                  </button>
                </div>
              </div>

              {banner && (
                <div
                  className={`p-3 rounded-[6px] border text-[12.5px] font-semibold flex items-center justify-between gap-3 ${
                    banner.type === "error"
                      ? "bg-red-50 border-red-300 text-red-900"
                      : banner.type === "info"
                      ? "bg-amber-50 border-amber-300 text-amber-900"
                      : "bg-emerald-50 border-emerald-400 text-emerald-900"
                  }`}
                >
                  <span>{banner.text}</span>
                  <button
                    type="button"
                    onClick={() => setBanner(null)}
                    className="text-[11px] font-bold uppercase underline cursor-pointer"
                  >
                    Dismiss
                  </button>
                </div>
              )}

              <div className="overflow-x-auto max-h-[500px] border border-black/10 rounded-[6px]">
                <table className="w-full text-left border-collapse text-[12px]">
                  <thead className="bg-[#f4f2ef] sticky top-0 z-10 border-b border-black/10 text-[10.5px] font-bold uppercase tracking-wider text-[#555]">
                    <tr>
                      <th className="py-2.5 px-3">#</th>
                      <th className="py-2.5 px-3">GRPO #</th>
                      <th className="py-2.5 px-3">GRPO Date</th>
                      <th className="py-2.5 px-3">BL Number</th>
                      <th className="py-2.5 px-3">Vendor</th>
                      <th className="py-2.5 px-3">BOE No</th>
                      <th className="py-2.5 px-3 text-right">Draft BOE / Duty Amt</th>
                      <th className="py-2.5 px-3 text-right">Scrips Attached (₹)</th>
                      <th className="py-2.5 px-3 text-center">Posted</th>
                      <th className="py-2.5 px-3 text-center">Reconciled</th>
                      <th className="py-2.5 px-3 text-center">Attached Scrips / Action</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-black/[0.06] bg-white">
                    {(() => {
                      const filteredRows =
                        grpoListViewMode === "ATTACHED_ONLY"
                          ? grpoAttachmentRows.filter((r) => r.isAttached)
                          : grpoListViewMode === "PENDING_ONLY"
                          ? grpoAttachmentRows.filter((r) => !r.isAttached)
                          : grpoAttachmentRows;

                      if (filteredRows.length === 0) {
                        return (
                          <tr>
                            <td colSpan={11} className="py-8 text-center text-[#8a8780]">
                              {grpoListViewMode === "ATTACHED_ONLY"
                                ? "No GRPOs or BLs have attached RoDTEP scrips yet. Select a GRPO above, check scrips, and click Attach."
                                : "No matching Import GRPOs found."}
                            </td>
                          </tr>
                        );
                      }

                      return filteredRows.map((r, idx) => {
                        const isExpanded = !!expandedGrpoDocs[r.rowKey];
                        return (
                          <React.Fragment key={r.rowKey}>
                            <tr className={`hover:bg-[#faf9f7] ${isExpanded ? "bg-blue-50/40" : ""}`}>
                              <td className="py-2.5 px-3 font-mono-code text-[#888]">{idx + 1}</td>
                              <td className="py-2.5 px-3 font-mono-code font-bold text-[#c8102e]">
                                {r.GRPODocNum ? `#${r.GRPODocNum}` : "—"}
                              </td>
                              <td className="py-2.5 px-3 font-mono-code">{r.GRPODate || "—"}</td>
                              <td className="py-2.5 px-3 font-mono-code font-bold text-[#1a1a1a]">
                                {r.BLNo || "—"}
                              </td>
                              <td className="py-2.5 px-3 font-medium truncate max-w-[200px]">
                                {r.CardName || "—"}
                              </td>
                              <td className="py-2.5 px-3 font-mono-code">
                                {r.BOENo ? (
                                  <span className="font-semibold text-[#1a1a1a]">{r.BOENo}</span>
                                ) : (
                                  <span className="text-amber-700 font-semibold">Draft / Pending</span>
                                )}
                              </td>
                              <td className="py-2.5 px-3 text-right font-mono-code font-semibold">
                                {r.DraftBOEAmount > 0 ? `₹${formatINR(r.DraftBOEAmount)}` : "—"}
                              </td>
                              <td className="py-2.5 px-3 text-right font-mono-code">
                                {r.isAttached ? (
                                  <span className="text-emerald-700 font-bold">
                                    ₹{formatINR(r.totalAttached)} ({r.attachedScrips.length} Scrips)
                                  </span>
                                ) : (
                                  <span className="text-[#999]">₹0.00</span>
                                )}
                              </td>
                              {/* Flag 1: Posted (Yes / No) */}
                              <td className="py-2.5 px-3 text-center">
                                {r.JENum ? (
                                  <span className="inline-flex items-center gap-1 px-2.5 py-1 bg-emerald-100 text-emerald-900 border border-emerald-400 rounded text-[10.5px] font-extrabold uppercase whitespace-nowrap">
                                    Yes (JE #{r.JENum})
                                  </span>
                                ) : (
                                  <span className="inline-flex items-center px-2.5 py-1 bg-rose-100 text-rose-800 border border-rose-300 rounded text-[10.5px] font-extrabold uppercase whitespace-nowrap">
                                    No
                                  </span>
                                )}
                              </td>
                              {/* Flag 2: Reconciled (Yes / No) */}
                              <td className="py-2.5 px-3 text-center">
                                {r.ReconciledInSAP ? (
                                  <span className="inline-flex items-center gap-1 px-2.5 py-1 bg-emerald-100 text-emerald-900 border border-emerald-400 rounded text-[10.5px] font-extrabold uppercase whitespace-nowrap">
                                    Yes
                                  </span>
                                ) : (
                                  <span className="inline-flex items-center px-2.5 py-1 bg-amber-100 text-amber-900 border border-amber-300 rounded text-[10.5px] font-extrabold uppercase whitespace-nowrap">
                                    No
                                  </span>
                                )}
                              </td>
                              {/* 3 Color-Coded Action / Status Options */}
                              <td className="py-2.5 px-3 text-center">
                                <div className="flex items-center justify-center gap-1.5 flex-wrap">
                                  {r.isAttached ? (
                                    <>
                                      <button
                                        type="button"
                                        onClick={() =>
                                          downloadChaExcel({
                                            blNo: r.BLNo,
                                            grpoDocNum: r.GRPODocNum,
                                            boeNo: r.BOENo,
                                            draftBoeAmount: r.DraftBOEAmount || r.totalAttached,
                                            attachedScrips: r.attachedScrips,
                                            allScrips: scrips
                                          })
                                        }
                                        className="px-2.5 py-1 bg-[#c8102e] hover:bg-[#a00c24] text-white rounded text-[10.5px] font-bold uppercase tracking-wider cursor-pointer transition-colors"
                                        title="Download Excel sheet for CHA with BL Number and attached RoDTEP Scrips"
                                      >
                                        ⬇ CHA Excel
                                      </button>
                                      <button
                                        type="button"
                                        onClick={() => toggleExpandGrpo(r.rowKey)}
                                        className="px-2.5 py-1 bg-blue-600 hover:bg-blue-700 text-white rounded text-[10.5px] font-bold uppercase tracking-wider cursor-pointer transition-colors"
                                      >
                                        {isExpanded
                                          ? `Hide Scrips (${r.attachedScrips.length}) ▲`
                                          : `View Scrips (${r.attachedScrips.length}) ▼`}
                                      </button>

                                      {/* OPTION 3: Posted = Yes & Reconciled = Yes -> Locked · Posted JE and Reconciled (Purple) + Undo/Cancel */}
                                      {r.JENum && r.ReconciledInSAP ? (
                                        <>
                                          <span
                                            className="px-2.5 py-1 bg-purple-700 text-white border border-purple-900 rounded text-[10.5px] font-bold uppercase tracking-wider whitespace-nowrap select-none"
                                            title="Both SAP Journal Entry and Internal Reconciliation on G/L 370001 are completed and locked"
                                          >
                                            🔒 Locked · Posted JE and Reconciled
                                          </span>
                                          <button
                                            type="button"
                                            disabled={submittingJe}
                                            onClick={() => handleCancelReconAndReverseJe(r, true)}
                                            className="px-2 py-1 bg-red-50 hover:bg-red-700 text-red-800 hover:text-white border border-red-300 rounded text-[10px] font-bold uppercase tracking-wider cursor-pointer transition-colors whitespace-nowrap"
                                            title="Cancel SAP Internal Reconciliation, reverse SAP JE via Service Layer, and detach scrips to select the right GRPO"
                                          >
                                            {submittingJe ? "Cancelling..." : "↺ Cancel Recon, Reverse JE & Detach"}
                                          </button>
                                        </>
                                      ) : r.JENum && !r.ReconciledInSAP ? (
                                        /* OPTION 2: Posted = Yes & Reconciled = No -> Reconcile (Amber/Orange) + Reverse JE & Detach */
                                        <>
                                          <button
                                            type="button"
                                            disabled={submittingJe}
                                            onClick={() => handleReconcileExistingJe(r)}
                                            className="px-2.5 py-1 bg-amber-600 hover:bg-amber-700 text-white rounded text-[10.5px] font-bold uppercase tracking-wider cursor-pointer transition-colors whitespace-nowrap"
                                            title="JE is already posted in SAP. Click to reconcile JE with linked A/P Invoice(s) on G/L 370001"
                                          >
                                            {submittingJe ? "Reconciling..." : "Reconcile"}
                                          </button>
                                          <button
                                            type="button"
                                            disabled={submittingJe}
                                            onClick={() => handleCancelReconAndReverseJe(r, true)}
                                            className="px-2 py-1 bg-red-50 hover:bg-red-700 text-red-800 hover:text-white border border-red-300 rounded text-[10px] font-bold uppercase tracking-wider cursor-pointer transition-colors whitespace-nowrap"
                                            title="Reverse SAP JE via Service Layer and detach scrips to select the right GRPO"
                                          >
                                            {submittingJe ? "Cancelling..." : "↺ Reverse JE & Detach"}
                                          </button>
                                        </>
                                      ) : (
                                        /* OPTION 1: Posted = No & Reconciled = No -> Post JE and Reconcile (Emerald Green) + Detach All */
                                        <>
                                          <button
                                            type="button"
                                            disabled={submittingJe}
                                            onClick={() => handlePostGrpoJe(r)}
                                            className="px-2.5 py-1 bg-emerald-700 hover:bg-emerald-800 text-white rounded text-[10.5px] font-bold uppercase tracking-wider cursor-pointer transition-colors whitespace-nowrap"
                                            title="Directly post SAP Custom Duty Journal Entry (Debit V00040 / Credit 370001) and reconcile with linked A/P Invoice(s) via SAP Service Layer"
                                          >
                                            {submittingJe ? "Posting JE & Reconciling..." : "Post JE and Reconcile"}
                                          </button>
                                          <button
                                            type="button"
                                            onClick={() => handleDetachAllForGrpo(r)}
                                            className="px-2 py-1 bg-gray-200 hover:bg-red-600 text-[#333] hover:text-white rounded text-[10.5px] font-bold uppercase tracking-wider cursor-pointer transition-colors whitespace-nowrap"
                                            title="Detach all scrips from this GRPO/BL and restore their balance"
                                          >
                                            ✕ Detach All
                                          </button>
                                        </>
                                      )}
                                    </>
                                  ) : (
                                    <button
                                      type="button"
                                      onClick={() => {
                                        if (r.GRPODocEntry) {
                                          handleSelectGrpo(String(r.GRPODocEntry));
                                        } else if (r.BLNo) {
                                          setBlNo(r.BLNo);
                                        }
                                        window.scrollTo({ top: 0, behavior: "smooth" });
                                      }}
                                      className="px-2.5 py-1 bg-[#1a1a1a] hover:bg-[#c8102e] text-white rounded text-[10.5px] font-bold uppercase tracking-wider cursor-pointer transition-colors"
                                    >
                                      + Select GRPO to Attach
                                    </button>
                                  )}
                                </div>
                              </td>
                            </tr>

                            {/* Inline Panel to Post Custom Duty <-> RoDTEP 370001 JE & Reconcile with AP Invoice */}
                            {activeJeGrpoKey === r.rowKey && (
                              <tr className="bg-emerald-50/70">
                                <td colSpan={11} className="p-4 border-l-4 border-l-emerald-700">
                                  <div className="bg-white p-4 rounded-[6px] border border-emerald-300 space-y-4">
                                    <div className="flex flex-wrap items-center justify-between gap-2 border-b border-black/10 pb-2">
                                      <div>
                                        <span className="text-[10px] font-bold uppercase tracking-wider text-emerald-800">
                                          SAP Service Layer · JournalEntries &amp; InternalReconciliations (G/L 370001)
                                        </span>
                                        <h4 className="text-[14px] font-extrabold text-[#1a1a1a]">
                                          Post Custom Duty Journal Entry &amp; Reconcile AP Invoice for{" "}
                                          {r.GRPODocNum ? `GRPO #${r.GRPODocNum}` : `BL ${r.BLNo}`}
                                        </h4>
                                      </div>
                                    </div>

                                    <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3">
                                      <div>
                                        <label className="block text-[10.5px] font-bold uppercase text-[#444] mb-1">
                                          1. Entry Direction (G/L 370001 vs Custom Duty)
                                        </label>
                                        <select
                                          value={jeFormDirection}
                                          onChange={(e) => setJeFormDirection(e.target.value)}
                                          className="w-full px-2.5 py-2 bg-white border border-black/20 rounded text-[11.5px] font-bold"
                                        >
                                          <option value="CREDIT_370001_DEBIT_DUTY">
                                            Debit Custom Duty / Credit RoDTEP 370001 (SAP BOE Standard — Reconciles with AP Debit)
                                          </option>
                                          <option value="DEBIT_370001_CREDIT_DUTY">
                                            Debit RoDTEP 370001 / Credit Custom Duty
                                          </option>
                                        </select>
                                      </div>

                                      <div>
                                        <label className="block text-[10.5px] font-bold uppercase text-[#444] mb-1">
                                          2. Custom Duty BP / G/L Account
                                        </label>
                                        <div className="flex gap-1">
                                          <select
                                            value={jeFormDutyType}
                                            onChange={(e) => {
                                              const t = e.target.value;
                                              setJeFormDutyType(t);
                                              setJeFormDutyCode(t === "BP" ? "V00040" : "520053");
                                            }}
                                            className="px-2 py-2 bg-[#f4f2ef] border border-black/20 rounded text-[11px] font-bold"
                                          >
                                            <option value="BP">BP</option>
                                            <option value="GL">G/L</option>
                                          </select>
                                          <input
                                            type="text"
                                            list="custom-duty-code-list"
                                            value={jeFormDutyCode}
                                            onChange={(e) => setJeFormDutyCode(e.target.value)}
                                            placeholder="V00040 or 520053"
                                            className="w-full px-2.5 py-2 bg-white border border-black/20 rounded text-[12px] font-mono-code font-bold"
                                          />
                                          <datalist id="custom-duty-code-list">
                                            <option value="V00040">V00040 — Custom Duty BP (Used in 57 BOE JEs)</option>
                                            <option value="520053">520053 — Custom Duty G/L</option>
                                            {(customDutyAccounts.bpAccounts || []).map((a) => (
                                              <option key={`BP_${a.Code}`} value={a.Code}>
                                                {a.Code} — {a.Name} (BP)
                                              </option>
                                            ))}
                                            {(customDutyAccounts.glAccounts || []).map((a) => (
                                              <option key={`GL_${a.Code}`} value={a.Code}>
                                                {a.Code} — {a.Name} (G/L)
                                              </option>
                                            ))}
                                          </datalist>
                                        </div>
                                      </div>

                                      <div>
                                        <label className="block text-[10.5px] font-bold uppercase text-[#444] mb-1">
                                          3. BOE Number &amp; JE Date
                                        </label>
                                        <div className="flex gap-1.5">
                                          <input
                                            type="text"
                                            value={jeFormBoeNo}
                                            onChange={(e) => {
                                              const v = e.target.value;
                                              setJeFormBoeNo(v);
                                              if (v) setJeFormMemo(`BCD of BOE No ${v}`);
                                            }}
                                            placeholder="BOE No"
                                            className="w-1/2 px-2 py-2 bg-white border border-black/20 rounded text-[12px] font-mono-code"
                                          />
                                          <input
                                            type="date"
                                            value={jeFormDate}
                                            onChange={(e) => setJeFormDate(e.target.value)}
                                            className="w-1/2 px-2 py-2 bg-white border border-black/20 rounded text-[11.5px] font-mono-code"
                                          />
                                        </div>
                                      </div>

                                      <div>
                                        <label className="block text-[10.5px] font-bold uppercase text-[#444] mb-1">
                                          4. JE Amount (₹)
                                        </label>
                                        <input
                                          type="number"
                                          step="0.01"
                                          value={jeFormAmount}
                                          onChange={(e) => setJeFormAmount(e.target.value)}
                                          className="w-full px-2.5 py-2 bg-white border border-black/20 rounded text-[12.5px] font-mono-code font-bold"
                                        />
                                      </div>

                                      <div>
                                        <label className="block text-[10.5px] font-bold uppercase text-[#444] mb-1">
                                          5. Remarks / LineMemo
                                        </label>
                                        <input
                                          type="text"
                                          value={jeFormMemo}
                                          onChange={(e) => setJeFormMemo(e.target.value)}
                                          placeholder="BCD of BOE No ..."
                                          className="w-full px-2.5 py-2 bg-white border border-black/20 rounded text-[12px] font-mono-code"
                                        />
                                      </div>
                                    </div>

                                    {/* Live A/P Invoice Internal Reconciliation Breakdown from Attached Scrips */}
                                    <div className="bg-[#f8fafc] border border-emerald-200 rounded p-3 space-y-2">
                                      <div className="text-[11px] font-bold uppercase tracking-wider text-emerald-950">
                                        SAP Internal Reconciliation Breakdown on G/L 370001 (Auto-Mapped from Attached RoDTEP Scrips):
                                      </div>
                                      <div className="overflow-x-auto">
                                        <table className="w-full text-left border-collapse text-[11.5px] bg-white border border-black/10 rounded">
                                          <thead className="bg-[#f4f2ef] text-[10px] font-bold uppercase text-[#555] border-b border-black/10">
                                            <tr>
                                              <th className="py-1.5 px-2.5">SAP AP DocNum (Ref1)</th>
                                              <th className="py-1.5 px-2.5">Vendor Bill Ref / NumAtCard (Ref2)</th>
                                              <th className="py-1.5 px-2.5">Vendor</th>
                                              <th className="py-1.5 px-2.5">Attached Scrips (Lic No)</th>
                                              <th className="py-1.5 px-2.5 text-right">Amount to Reconcile on 370001 (₹)</th>
                                            </tr>
                                          </thead>
                                          <tbody className="divide-y divide-black/[0.06]">
                                            {getGrpoApBreakdown(r).map((b, bIdx) => (
                                              <tr key={bIdx}>
                                                <td className="py-1.5 px-2.5 font-mono-code font-bold text-[#c8102e]">
                                                  {b.apDocNum ? `AP #${b.apDocNum}` : "Unlinked (Oldest Open AP)"}
                                                </td>
                                                <td className="py-1.5 px-2.5 font-mono-code font-bold text-[#1a1a1a]">
                                                  {b.numAtCard || "—"}
                                                </td>
                                                <td className="py-1.5 px-2.5">{b.cardName || "—"}</td>
                                                <td className="py-1.5 px-2.5 font-mono-code text-[11px]">
                                                  {b.licNos.join(", ")}
                                                </td>
                                                <td className="py-1.5 px-2.5 text-right font-mono-code font-bold text-emerald-700">
                                                  ₹{formatINR(b.amount)}
                                                </td>
                                              </tr>
                                            ))}
                                          </tbody>
                                        </table>
                                      </div>
                                    </div>

                                    <div className="flex flex-wrap items-center justify-between gap-3 pt-2 border-t border-black/10">
                                      <label className="flex items-center gap-2 text-[12px] font-medium cursor-pointer">
                                        <input
                                          type="checkbox"
                                          checked={jeFormReconcileAp}
                                          onChange={(e) => setJeFormReconcileAp(e.target.checked)}
                                          className="w-4 h-4 accent-emerald-700"
                                        />
                                        <span>
                                          <strong>Also Reconcile with Linked SAP A/P Invoice(s) on G/L 370001</strong> (via SAP Internal Reconciliation)
                                        </span>
                                      </label>

                                      <div className="flex items-center gap-2">
                                        <button
                                          type="button"
                                          onClick={() => setActiveJeGrpoKey(null)}
                                          className="px-3 py-2 bg-[#f4f2ef] hover:bg-[#e5e1da] text-[#333] rounded text-[11px] font-bold uppercase cursor-pointer"
                                        >
                                          Cancel
                                        </button>
                                        <button
                                          type="button"
                                          disabled={submittingJe}
                                          onClick={() => handlePostGrpoJe(r)}
                                          className="px-4 py-2 bg-emerald-700 hover:bg-emerald-800 text-white rounded text-[11.5px] font-bold uppercase tracking-wider cursor-pointer"
                                        >
                                          {submittingJe
                                            ? "Posting JE & Reconciling in SAP..."
                                            : `Post SAP JE (₹${formatINR(jeFormAmount)}) & Reconcile`}
                                        </button>
                                      </div>
                                    </div>
                                  </div>
                                </td>
                              </tr>
                            )}

                            {/* Expandable Child Row listing all RoDTEP Scrips attached to this GRPO / BL */}
                            {isExpanded && r.attachedScrips.length > 0 && (
                              <tr className="bg-[#f8fafc]">
                                <td colSpan={11} className="p-3 border-l-4 border-l-blue-600">
                                  <div className="text-[11px] font-bold uppercase tracking-wider text-blue-900 mb-2">
                                    Attached RoDTEP / RoSCTL Scrips for{" "}
                                    {r.GRPODocNum ? `GRPO #${r.GRPODocNum}` : ""} (BL: {r.BLNo || "—"}) —{" "}
                                    {r.attachedScrips.length} Scrips Totaling ₹{formatINR(r.totalAttached)}
                                  </div>
                                  <div className="overflow-x-auto border border-blue-200 rounded bg-white">
                                    <table className="w-full text-left border-collapse text-[11.5px]">
                                      <thead className="bg-blue-50 text-[10px] font-bold uppercase tracking-wider text-blue-950 border-b border-blue-200">
                                        <tr>
                                          <th className="py-1.5 px-2.5">#</th>
                                          <th className="py-1.5 px-2.5">RoDTEP Lic No</th>
                                          <th className="py-1.5 px-2.5">Scheme</th>
                                          <th className="py-1.5 px-2.5">Exporter</th>
                                          <th className="py-1.5 px-2.5">Port Code</th>
                                          <th className="py-1.5 px-2.5">AP Inv #</th>
                                          <th className="py-1.5 px-2.5 text-right">Original Scrip Amt (₹)</th>
                                          <th className="py-1.5 px-2.5 text-right">Amount Used on This GRPO/BL (₹)</th>
                                          <th className="py-1.5 px-2.5 text-center">Action</th>
                                        </tr>
                                      </thead>
                                      <tbody className="divide-y divide-blue-100">
                                        {r.attachedScrips.map((u, uIdx) => (
                                          <tr key={u.Id} className="hover:bg-blue-50/30">
                                            <td className="py-1.5 px-2.5 font-mono-code text-[#777]">{uIdx + 1}</td>
                                            <td className="py-1.5 px-2.5 font-mono-code font-bold text-[#c8102e]">
                                              {u.LicNo}
                                            </td>
                                            <td className="py-1.5 px-2.5 font-bold text-[10.5px]">{u.Scheme}</td>
                                            <td className="py-1.5 px-2.5">{u.ExporterName}</td>
                                            <td className="py-1.5 px-2.5 font-mono-code">{u.PortCode || "—"}</td>
                                            <td className="py-1.5 px-2.5 font-mono-code">
                                              {u.APDocNum ? `#${u.APDocNum}` : "—"}
                                            </td>
                                            <td className="py-1.5 px-2.5 text-right font-mono-code text-[#666]">
                                              ₹{formatINR(u.ScripAmount)}
                                            </td>
                                            <td className="py-1.5 px-2.5 text-right font-mono-code font-bold text-emerald-700">
                                              ₹{formatINR(u.AmountUsed)}
                                            </td>
                                            <td className="py-1.5 px-2.5 text-center">
                                              {r.JENum ? (
                                                <span className="text-purple-700 text-[10.5px] font-bold uppercase">
                                                  🔒 Locked
                                                </span>
                                              ) : (
                                                <button
                                                  type="button"
                                                  onClick={() => handleDeleteUtilization(u.Id)}
                                                  className="text-red-600 hover:underline text-[11px] font-bold cursor-pointer"
                                                >
                                                  Detach
                                                </button>
                                              )}
                                            </td>
                                          </tr>
                                        ))}
                                      </tbody>
                                    </table>
                                  </div>
                                </td>
                              </tr>
                            )}
                          </React.Fragment>
                        );
                      });
                    })()}
                  </tbody>
                </table>
              </div>
            </div>
          </form>
        )}

        {/* ===================================================================== */}
        {/* TAB 3: ON-THE-FLY SCRIP BALANCE LEDGER & BL UTILIZATION HISTORY       */}
        {/* ===================================================================== */}
        {activeTab === "ledger" && (
          <div className="space-y-6">
            {/* Table 2: BL <-> Scrip Linking History */}
            <div className="rp-card p-6 space-y-4">
              <div className="flex items-center justify-between border-b border-black/[0.08] pb-3">
                <div>
                  <span className="sec-label">Table 2 · BHV_RODTEP_BL</span>
                  <h3 className="text-lg font-extrabold text-[#1a1a1a] mt-0.5">
                    BL / GRPO / Draft BOE Scrip Attachments ({utilizations.length})
                  </h3>
                </div>
              </div>

              <div className="overflow-x-auto border border-black/10 rounded-[6px]">
                <table className="w-full text-left border-collapse text-[12px]">
                  <thead className="bg-[#f4f2ef] border-b border-black/10 text-[10.5px] font-bold uppercase tracking-wider text-[#555]">
                    <tr>
                      <th className="py-2.5 px-3">BL Number</th>
                      <th className="py-2.5 px-3">GRPO #</th>
                      <th className="py-2.5 px-3">BOE No</th>
                      <th className="py-2.5 px-3 text-right">Draft BOE Amt</th>
                      <th className="py-2.5 px-3">RoDTEP Lic No</th>
                      <th className="py-2.5 px-3">Scheme / Exporter</th>
                      <th className="py-2.5 px-3 text-right">Amount Used (₹)</th>
                      <th className="py-2.5 px-3 text-center">Action</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-black/[0.06] bg-white">
                    {utilizations.length === 0 ? (
                      <tr>
                        <td colSpan={9} className="py-6 text-center text-[#8a8780]">
                          No BL attachments recorded yet.
                        </td>
                      </tr>
                    ) : (
                      utilizations.map((u) => (
                        <tr key={u.Id} className="hover:bg-[#faf9f7]">
                          <td className="py-2.5 px-3 font-mono-code font-bold text-[#1a1a1a]">
                            {u.BLNo}
                          </td>
                          <td className="py-2.5 px-3 font-mono-code">
                            {u.GRPODocNum ? `#${u.GRPODocNum}` : "—"}
                          </td>
                          <td className="py-2.5 px-3 font-mono-code">
                            {u.BOENo || <span className="text-amber-700 font-semibold">Draft</span>}
                          </td>
                          <td className="py-2.5 px-3 text-right font-mono-code">
                            ₹{formatINR(u.DraftBOEAmount)}
                          </td>
                          <td className="py-2.5 px-3 font-mono-code font-bold text-[#c8102e]">
                            {u.LicNo}
                          </td>
                          <td className="py-2.5 px-3 text-[11.5px]">
                            {u.Scheme} · {u.ExporterName}
                          </td>
                          <td className="py-2.5 px-3 text-right font-mono-code font-bold text-[#1a1a1a]">
                            ₹{formatINR(u.AmountUsed)}
                          </td>
                          <td className="py-2.5 px-3 text-center">
                            <button
                              type="button"
                              onClick={() => handleDeleteUtilization(u.Id)}
                              className="text-red-600 hover:underline text-[11px] font-bold cursor-pointer"
                            >
                              Remove
                            </button>
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Table 1: Master Scrip Inventory with On-the-Fly Calculated Balance */}
            <div className="rp-card p-6 space-y-4">
              <div className="flex items-center justify-between border-b border-black/[0.08] pb-3">
                <div>
                  <span className="sec-label">
                    Table 1 · BHV_RODTEP_SCRIP (Balance Calculated On-The-Fly)
                  </span>
                  <h3 className="text-lg font-extrabold text-[#1a1a1a] mt-0.5">
                    All RoDTEP / RoSCTL Scrips ({scrips.length}) — Link or Change SAP A/P Invoice # Anytime
                  </h3>
                </div>
                <div className="w-full sm:w-80">
                  <input
                    type="text"
                    value={ledgerSearch}
                    onChange={(e) => setLedgerSearch(e.target.value)}
                    placeholder="Search Lic No, Exporter, Amount, AP #..."
                    className="w-full px-3 py-1.5 bg-white border border-black/20 rounded-[4px] text-[12px] font-mono-code"
                  />
                </div>
              </div>

              <div className="overflow-x-auto border border-black/10 rounded-[6px]">
                <table className="w-full text-left border-collapse text-[12px]">
                  <thead className="bg-[#f4f2ef] border-b border-black/10 text-[10.5px] font-bold uppercase tracking-wider text-[#555]">
                    <tr>
                      <th className="py-2.5 px-3">Lic No</th>
                      <th className="py-2.5 px-3">Dated</th>
                      <th className="py-2.5 px-3">Exporter</th>
                      <th className="py-2.5 px-3">Scheme</th>
                      <th className="py-2.5 px-3">Port Code</th>
                      <th className="py-2.5 px-3">AP Inv #</th>
                      <th className="py-2.5 px-3 text-right">Scrip Amount</th>
                      <th className="py-2.5 px-3 text-right">Utilized (On-The-Fly)</th>
                      <th className="py-2.5 px-3 text-right">Balance (On-The-Fly)</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-black/[0.06] bg-white">
                    {scrips
                      .filter((s) => {
                        const q = ledgerSearch.trim().toLowerCase();
                        if (!q) return true;
                        return (
                          String(s.LicNo || "").toLowerCase().includes(q) ||
                          String(s.ExporterName || "").toLowerCase().includes(q) ||
                          String(s.PortLocation || "").toLowerCase().includes(q) ||
                          String(s.APDocNum || "").toLowerCase().includes(q) ||
                          String(s.NumAtCard || "").toLowerCase().includes(q) ||
                          String(s.ScripAmount || "").includes(q)
                        );
                      })
                      .map((s) => {
                      const bal = Number(s.BalanceAmount || 0);
                      const util = Number(s.UtilizedAmount || 0);
                      return (
                        <tr key={s.LicNo} className="hover:bg-[#faf9f7]">
                          <td className="py-2 px-3 font-mono-code font-bold">{s.LicNo}</td>
                          <td className="py-2 px-3 font-mono-code">{s.LicDate}</td>
                          <td className="py-2 px-3">{s.ExporterName}</td>
                          <td className="py-2 px-3 font-bold text-[11px]">{s.Scheme}</td>
                          <td className="py-2 px-3 font-mono-code">{s.PortCode}</td>
                          <td className="py-2 px-3 font-mono-code">
                            <select
                              value={s.APDocEntry || ""}
                              onChange={(e) => handleLinkApToExistingScrip(s.LicNo, e.target.value)}
                              className={`px-2 py-1 rounded text-[11px] font-mono-code border ${
                                s.APDocNum
                                  ? "bg-emerald-50 border-emerald-300 text-emerald-900 font-bold"
                                  : "bg-amber-50 border-amber-300 text-amber-900"
                              }`}
                            >
                              <option value="">-- Unlinked (Link AP Later) --</option>
                              {apInvoices
                                .filter((inv) => {
                                  if (Number(s.APDocEntry) === Number(inv.DocEntry)) return true;
                                  const info = apLinkMap[`ENTRY_${inv.DocEntry}`];
                                  return !info || !info.isFullyLinked;
                                })
                                .map((inv) => (
                                  <option key={inv.DocEntry} value={inv.DocEntry}>
                                    #{inv.DocNum} | {inv.NumAtCard || inv.DocDate} | ₹{formatINR(inv.InternalReconPending)}
                                  </option>
                                ))}
                            </select>
                          </td>
                          <td className="py-2 px-3 text-right font-mono-code">
                            ₹{formatINR(s.ScripAmount)}
                          </td>
                          <td className="py-2 px-3 text-right font-mono-code text-amber-700">
                            ₹{formatINR(util)}
                          </td>
                          <td
                            className={`py-2 px-3 text-right font-mono-code font-bold ${
                              bal === 0 ? "text-gray-400" : "text-emerald-700"
                            }`}
                          >
                            ₹{formatINR(bal)}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}

function discountPercentageOrAmtDisplay(val) {
  return val;
}
