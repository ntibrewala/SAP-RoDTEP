import { useState, useEffect, useCallback, useMemo } from "react";
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
THE PARADISE LINEN,2610023656,22.10.2026,ROSCTL,INMAA1,298820,Chennai
STR GARMENTS,2611005776,05.11.2026,ROSCTL,INTUT6,291326,Chennai
STR GARMENTS,2611005714,05.11.2026,ROSCTL,INMAA1,248477,Chennai
VKS HYTECH PRIVATE LIMITED,2608030729,14.08.2026,RODTEP,INDEL4,208634,Mundra
VISHNU ENTERPRISES,2611009820,06.11.2026,ROSCTL,INCOK1,205502,Mundra
PAKIZA FASHION POINT,2611010269,07.11.2026,ROSCTL,INNSA1,184981,Mundra
SIDDHRAJ CHEMICALS LLP,2611001420,03.11.2026,RODTEP,INHZA1,128933,Mundra
MAULIK LEATHERCRAFTS PRIVATE LIMITED,2611002219,05.11.2026,RODTEP,INMAA4,125964,Mundra
VISHNU ENTERPRISES,2611009819,06.11.2026,RODTEP,INCOK1,113853,Mundra
STR GARMENTS,2611005755,05.11.2026,ROSCTL,INMAA4,113794,Mundra
STR GARMENTS,2611005594,05.11.2026,ROSCTL,INCOK1,89189,Chennai
NEW INDIA EXIM,2609024500,12.09.2026,RODTEP,INSBI6,68698,Mundra
VISHNU ENTERPRISES,2611001974,04.11.2026,RODTEP,INCOK1,59148,Mundra
SHREE VANI TRADING CO,2604178808,29.04.2026,RODTEP,INNSA1,57729,Mundra
PAKIZA FASHION POINT,2611010268,07.11.2026,RODTEP,INNSA1,56182,Mundra
SIDDHRAJ CHEMICALS LLP,2611001446,03.11.2026,RODTEP,INMUN1,50900,Mundra
ESSCO INTERNATIONAL,2610047851,31.10.2026,RODTEP,INNSA1,42699,Mundra
STR GARMENTS,2611005985,05.11.2026,ROSCTL,INBLR4,40187,Chennai
M.B.INTERNATIONAL,2611000122,02.11.2026,RODTEP,INMUN1,33335,Mundra
MAULIK LEATHERCRAFTS PRIVATE LIMITED,2611002249,05.11.2026,ROSCTL,INMAA4,29794,Mundra
ILEX ODYSSEY LLP,2611001396,03.11.2026,RODTEP,INCOK1,29085,Mundra
NEW INDIA EXIM,2606045474,24.06.2026,RODTEP,INMUN1,25495,Mundra
SARITA HANDA EXPORTS PRIVATE LIMITED,2610047266,30.10.2026,RODTEP,INBOM4,19846,Mundra
THE PARADISE LINEN,2611007929,05.11.2026,ROSCTL,INMAA1,9840,Chennai
SIDDHRAJ CHEMICALS LLP,2611001442,03.11.2026,RODTEP,INNSA1,9324,Mundra
MOHAN MOTORS,2610046108,29.10.2026,RODTEP,INWFD6,5056,Mundra
BELIEVE COSMETICS PRIVATE LIMITED,2504000070,01.04.2025,RODTEP,INNSA1,5003,Mundra
,,,,,2551794,`;

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

  for (const line of lines) {
    // Support both Excel Tab-separated (TSV) and Comma-separated (CSV)
    const delimiter = line.includes("\t") ? "\t" : ",";
    const cols = line.split(delimiter).map((c) => c.trim());

    if (cols.length < 2) continue;

    const firstUpper = (cols[0] || "").toUpperCase();
    const secondUpper = (cols[1] || "").toUpperCase();

    // Skip header row
    if (
      firstUpper.includes("EXPORTER") ||
      secondUpper.includes("LIC NO") ||
      secondUpper.includes("LICENCE")
    ) {
      continue;
    }

    const exporter = cols[0] || "";
    const licNo = cols[1] || "";
    const licDate = cols[2] || "";
    const scheme = (cols[3] || "RODTEP").toUpperCase();
    const portCode = (cols[4] || "").toUpperCase();
    const rawAmt = (cols[5] || "").replace(/,/g, "").trim();
    const rawPortLoc = cols[6] || "";

    // Skip summary/total rows where LIC NO is empty (e.g. ",,,,,2551794,")
    if (!licNo) continue;

    const amount = Number(rawAmt);
    if (isNaN(amount) || amount <= 0) continue;

    let portLocation = defaultPort;
    const upperLoc = rawPortLoc.toUpperCase();
    if (upperLoc.includes("CHENNAI")) {
      portLocation = "Chennai";
    } else if (upperLoc.includes("MUNDRA")) {
      portLocation = "Mundra";
    } else if (rawPortLoc) {
      portLocation = rawPortLoc;
    }

    parsedRows.push({
      ExporterName: exporter,
      LicNo: licNo,
      LicDate: licDate,
      Scheme: scheme,
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

  // --- MODULE 1 STATE (Excel Paste -> Auto-Create AP Invoice) ---
  const [apMode, setApMode] = useState("CREATE_AP"); // "CREATE_AP" or "LINK_EXISTING"
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

  const fetchAllData = useCallback(async () => {
    setLoading(true);
    try {
      const [vRes, apRes, grpoRes, scripRes] = await Promise.all([
        apiRequest("GET", "/vendors"),
        apiRequest("GET", "/ap-invoices"),
        apiRequest("GET", "/grpo-bl-list"),
        apiRequest("GET", "/scrips")
      ]);
      setVendors(vRes.data || []);
      setApInvoices(apRes.data || []);
      setGrpos(grpoRes.data?.grpos || []);
      setHbls(grpoRes.data?.hbls || []);
      setScrips(scripRes.data?.scrips || []);
      setUtilizations(scripRes.data?.utilizations || []);
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
      if (g.BLNo) setBlNo(g.BLNo);
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
    if (portViewMode === "ONLY_MATCHING_PORT") {
      return pendingScrips.filter(
        (s) => String(s.PortLocation || "").toUpperCase() === selectedPort.toUpperCase()
      );
    }
    // Sort matching port first, or by date, while showing all pending scrips
    return [...pendingScrips].sort((a, b) => {
      const aMatch =
        String(a.PortLocation || "").toUpperCase() === selectedPort.toUpperCase() ? 0 : 1;
      const bMatch =
        String(b.PortLocation || "").toUpperCase() === selectedPort.toUpperCase() ? 0 : 1;
      if (aMatch !== bMatch) return aMatch - bMatch;
      return String(a.LicDate || "").localeCompare(String(b.LicDate || ""));
    });
  }, [pendingScrips, portViewMode, selectedPort]);

  const handleToggleScripCheckbox = (scrip) => {
    const isPortMatch =
      String(scrip.PortLocation || "").toUpperCase() === selectedPort.toUpperCase();
    if (!isPortMatch) return; // Strictly locked if port doesn't match

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

    const matchingPortScrips = pendingScrips.filter(
      (s) => String(s.PortLocation || "").toUpperCase() === selectedPort.toUpperCase()
    );

    let remainingNeeded = target;
    const nextChecked = {};

    for (const s of matchingPortScrips) {
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
      setBanner({
        type: "success",
        text: `Attached ${res.data.count} RoDTEP scrip(s) totaling ₹${formatINR(res.data.totalAllocated)} to BL ${res.data.blNo}${res.data.grpoDocNum ? ` (GRPO #${res.data.grpoDocNum})` : ""} for Port ${res.data.portLocation}!`
      });

      setCheckedScrips({});
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
        <div className="grid grid-cols-2 md:grid-cols-5 gap-3">
          <div className="rp-card p-3.5">
            <div className="text-[10px] font-bold uppercase tracking-[0.1em] text-[#6b6863]">
              Total Scrips Purchased
            </div>
            <div className="text-[17px] font-extrabold font-mono-code text-[#1a1a1a] mt-1">
              ₹{formatINR(globalKpis.totalPurchased)}
            </div>
            <div className="text-[10.5px] text-[#8a8780] mt-0.5">
              {globalKpis.scripCount} Licences in Table 1
            </div>
          </div>

          <div className="rp-card p-3.5">
            <div className="text-[10px] font-bold uppercase tracking-[0.1em] text-[#6b6863]">
              Utilized on BL / BOE (On-The-Fly)
            </div>
            <div className="text-[17px] font-extrabold font-mono-code text-amber-700 mt-1">
              ₹{formatINR(globalKpis.totalUtilized)}
            </div>
            <div className="text-[10.5px] text-[#8a8780] mt-0.5">
              {utilizations.length} BL Allocation rows
            </div>
          </div>

          <div className="rp-card p-3.5 border-l-4 border-l-emerald-600">
            <div className="text-[10px] font-bold uppercase tracking-[0.1em] text-emerald-800">
              Total Net Available Balance
            </div>
            <div className="text-[17px] font-extrabold font-mono-code text-emerald-700 mt-1">
              ₹{formatINR(globalKpis.totalBalance)}
            </div>
            <div className="text-[10.5px] text-[#8a8780] mt-0.5">
              {globalKpis.pendingCount} Pending Scrips
            </div>
          </div>

          <div className="rp-card p-3.5 border-l-4 border-l-blue-600">
            <div className="text-[10px] font-bold uppercase tracking-[0.1em] text-blue-800">
              Mundra Port Pending
            </div>
            <div className="text-[17px] font-extrabold font-mono-code text-blue-700 mt-1">
              ₹{formatINR(globalKpis.mundraBalance)}
            </div>
            <div className="text-[10.5px] text-[#8a8780] mt-0.5">
              Available for Mundra BLs
            </div>
          </div>

          <div className="rp-card p-3.5 border-l-4 border-l-purple-600">
            <div className="text-[10px] font-bold uppercase tracking-[0.1em] text-purple-800">
              Chennai Port Pending
            </div>
            <div className="text-[17px] font-extrabold font-mono-code text-purple-700 mt-1">
              ₹{formatINR(globalKpis.chennaiBalance)}
            </div>
            <div className="text-[10.5px] text-[#8a8780] mt-0.5">
              Available for Chennai BLs
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
            <div className="rp-card p-6 space-y-5">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 border-b border-black/[0.08] pb-4">
                <div>
                  <span className="sec-label">Module 1 · Table 1 (BHV_RODTEP_SCRIP)</span>
                  <h2 className="text-xl font-extrabold text-[#1a1a1a] mt-1">
                    Purchase RoDTEP / RoSCTL Scrips &amp; Create SAP A/P Invoice
                  </h2>
                </div>

                <div className="flex items-center gap-2">
                  <button
                    type="button"
                    onClick={() => setApMode("CREATE_AP")}
                    className={`px-3.5 py-2 rounded-[4px] text-[11px] font-bold uppercase tracking-wider cursor-pointer ${
                      apMode === "CREATE_AP"
                        ? "bg-[#1a1a1a] text-white"
                        : "bg-[#f4f2ef] text-[#555] hover:bg-[#e8e4dd]"
                    }`}
                  >
                    Create New AP Invoice (Service Layer)
                  </button>
                  <button
                    type="button"
                    onClick={() => setApMode("LINK_EXISTING")}
                    className={`px-3.5 py-2 rounded-[4px] text-[11px] font-bold uppercase tracking-wider cursor-pointer ${
                      apMode === "LINK_EXISTING"
                        ? "bg-[#1a1a1a] text-white"
                        : "bg-[#f4f2ef] text-[#555] hover:bg-[#e8e4dd]"
                    }`}
                  >
                    Link to Existing SAP AP Invoice
                  </button>
                </div>
              </div>

              {/* A/P Invoice Header Fields */}
              {apMode === "LINK_EXISTING" ? (
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4 bg-[#f4f2ef]/70 p-4 rounded-[6px] border border-black/[0.08]">
                  <div>
                    <label className="block text-[11px] font-bold uppercase tracking-wider text-[#444] mb-1.5">
                      Select Existing SAP A/P Invoice (G/L 370001 - RODTEP Scrip Account)
                    </label>
                    <select
                      value={selectedExistingApDocEntry}
                      onChange={(e) => handleSelectExistingAp(e.target.value)}
                      className="w-full px-3 py-2.5 bg-white border border-black/20 rounded-[4px] text-[13px] font-mono-code"
                    >
                      <option value="">-- Choose SAP A/P Invoice --</option>
                      {apInvoices.map((inv) => (
                        <option key={inv.DocEntry} value={inv.DocEntry}>
                          #{inv.DocNum} | {inv.DocDate} | {inv.CardName} ({inv.CardCode}) | Ref:{" "}
                          {inv.NumAtCard || "-"} | Scrip 370001: ₹{formatINR(inv.ScripFaceValue)}
                        </option>
                      ))}
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
                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4">
                  <div className="lg:col-span-2">
                    <label className="block text-[11px] font-bold uppercase tracking-wider text-[#444] mb-1">
                      Vendor (CardCode) *
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
                      Vendor Ref No (NumAtCard)
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
                      Posting Date (DocDate)
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
                      Gain / Discount (G/L 410023 ₹)
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
              )}

              {/* Excel Copy-Paste Area */}
              <div className="space-y-2">
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <label className="text-[11px] font-bold uppercase tracking-wider text-[#444]">
                    Copy &amp; Paste from Excel (EXPORTER | LIC NO | DATED | SCHEME | PORT CODE | SCRIP AMOUNT | Port)
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
                  placeholder="Click here and press Ctrl+V / Cmd+V after copying rows from Excel..."
                  className="w-full p-3 bg-[#faf9f7] border border-dashed border-black/30 rounded-[4px] text-[12px] font-mono-code focus:border-[#c8102e] focus:bg-white outline-none"
                />
              </div>

              {/* Pasted Scrips Preview & Editable Grid */}
              {pastedScrips.length > 0 && (
                <div className="space-y-4">
                  {/* Summary Bar */}
                  <div className="grid grid-cols-2 sm:grid-cols-5 gap-3 bg-[#1c1c1c] text-white p-4 rounded-[6px]">
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

                    <div>
                      <div className="text-[10px] uppercase tracking-wider text-blue-300">
                        Mundra Scrips ({pastedSummary.mundraCount})
                      </div>
                      <div className="text-[15px] font-bold font-mono-code text-blue-300 mt-0.5">
                        ₹{formatINR(pastedSummary.mundraTotal)}
                      </div>
                    </div>

                    <div>
                      <div className="text-[10px] uppercase tracking-wider text-purple-300">
                        Chennai Scrips ({pastedSummary.chennaiCount})
                      </div>
                      <div className="text-[15px] font-bold font-mono-code text-purple-300 mt-0.5">
                        ₹{formatINR(pastedSummary.chennaiTotal)}
                      </div>
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
                          <th className="py-2.5 px-3">Port Location</th>
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
                              <td className="py-2 px-3">
                                <select
                                  value={row.PortLocation}
                                  onChange={(e) =>
                                    handleUpdatePastedRow(idx, "PortLocation", e.target.value)
                                  }
                                  className={`px-2 py-1 rounded text-[11px] font-bold border ${
                                    String(row.PortLocation).toUpperCase() === "CHENNAI"
                                      ? "bg-purple-50 text-purple-800 border-purple-200"
                                      : "bg-blue-50 text-blue-800 border-blue-200"
                                  }`}
                                >
                                  <option value="Mundra">Mundra</option>
                                  <option value="Chennai">Chennai</option>
                                </select>
                              </td>
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
                        ? `Create SAP A/P Invoice (₹${formatINR(pastedSummary.netPayable)}) & Save ${pastedSummary.count} Scrips`
                        : `Link ${pastedSummary.count} Scrips (₹${formatINR(pastedSummary.total)}) to Selected A/P Invoice`}
                    </button>
                  </div>
                </div>
              )}
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

                {/* Port Filter / Lock Mode Toggle */}
                <div className="flex items-center gap-2">
                  <button
                    type="button"
                    onClick={() => setPortViewMode("SHOW_ALL_LOCK_OTHER")}
                    className={`px-3 py-1.5 rounded text-[11px] font-bold uppercase cursor-pointer ${
                      portViewMode === "SHOW_ALL_LOCK_OTHER"
                        ? "bg-[#1a1a1a] text-white"
                        : "bg-[#f4f2ef] text-[#555]"
                    }`}
                  >
                    Show All Pending (Lock Other Port)
                  </button>
                  <button
                    type="button"
                    onClick={() => setPortViewMode("ONLY_MATCHING_PORT")}
                    className={`px-3 py-1.5 rounded text-[11px] font-bold uppercase cursor-pointer ${
                      portViewMode === "ONLY_MATCHING_PORT"
                        ? "bg-[#1a1a1a] text-white"
                        : "bg-[#f4f2ef] text-[#555]"
                    }`}
                  >
                    Show Only {selectedPort} Scrips
                  </button>
                </div>
              </div>

              {/* Main Header Form: GRPO Selection, BL Number, Port, Draft BOE Amount */}
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-5 gap-4 bg-[#f4f2ef]/80 p-4 rounded-[6px] border border-black/[0.08]">
                <div className="lg:col-span-2">
                  <label className="block text-[11px] font-bold uppercase tracking-wider text-[#444] mb-1">
                    1. Select Import GRPO (Captures GRPO # &amp; Auto-Picks BL / Port)
                  </label>
                  <select
                    value={selectedGrpoDocEntry}
                    onChange={(e) => handleSelectGrpo(e.target.value)}
                    className="w-full px-3 py-2 bg-white border border-black/20 rounded-[4px] text-[12.5px] font-mono-code"
                  >
                    <option value="">-- Optional: Select Import GRPO (or enter BL manually) --</option>
                    {grpos.map((g) => (
                      <option key={g.GRPODocEntry} value={g.GRPODocEntry}>
                        GRPO #{g.GRPODocNum} | BL: {g.BLNo || "N/A"} | Port: {g.DetectedPort} |{" "}
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
                      if (matchedHbl) {
                        if (matchedHbl.DetectedPort) handlePortChange(matchedHbl.DetectedPort);
                        if (matchedHbl.GRPODocEntry) {
                          setSelectedGrpoDocEntry(String(matchedHbl.GRPODocEntry));
                          setSelectedGrpoDocNum(String(matchedHbl.GRPODocNum || ""));
                        }
                      }
                    }}
                    placeholder="Enter or pick BL No..."
                    className="w-full px-3 py-2 bg-white border border-black/20 rounded-[4px] text-[13px] font-mono-code font-bold"
                    required
                  />
                  <datalist id="sap-hbl-list">
                    {hbls.map((h) => (
                      <option key={h.HBLDocEntry} value={h.BLNo}>
                        {h.DetectedPort ? `Port: ${h.DetectedPort}` : ""}{" "}
                        {h.GRPODocNum ? `| GRPO #${h.GRPODocNum}` : ""}
                      </option>
                    ))}
                  </datalist>
                </div>

                <div>
                  <label className="block text-[11px] font-bold uppercase tracking-wider text-[#444] mb-1">
                    3. Port * (Locks Other Port)
                  </label>
                  <select
                    value={selectedPort}
                    onChange={(e) => handlePortChange(e.target.value)}
                    className={`w-full px-3 py-2 rounded-[4px] text-[13px] font-extrabold border ${
                      selectedPort === "Chennai"
                        ? "bg-purple-50 text-purple-900 border-purple-300"
                        : "bg-blue-50 text-blue-900 border-blue-300"
                    }`}
                  >
                    <option value="Mundra">Mundra</option>
                    <option value="Chennai">Chennai</option>
                  </select>
                </div>

                <div>
                  <label className="block text-[11px] font-bold uppercase tracking-wider text-[#444] mb-1">
                    4. Draft BOE / Duty Amount (₹) *
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
                      Selected BL &amp; Port
                    </div>
                    <div className="text-[14px] font-bold font-mono-code">
                      {blNo || "—"} · <span className="text-amber-300">{selectedPort}</span>
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

              {/* Pending RoDTEP Scrips Table with Checkbox & Editable Amount */}
              <div className="overflow-x-auto border border-black/10 rounded-[6px] max-h-[480px]">
                <table className="w-full text-left border-collapse text-[12.5px]">
                  <thead className="bg-[#f4f2ef] sticky top-0 z-10 border-b border-black/10 text-[10.5px] font-bold uppercase tracking-wider text-[#555]">
                    <tr>
                      <th className="py-3 px-3 text-center w-14">Check</th>
                      <th className="py-3 px-3">Port</th>
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
                        <td colSpan={11} className="py-8 text-center text-[#8a8780]">
                          No pending RoDTEP scrips found. Go to Tab 1 to paste your Excel scrips.
                        </td>
                      </tr>
                    ) : (
                      displayedPendingScrips.map((scrip) => {
                        const isPortMatch =
                          String(scrip.PortLocation || "").toUpperCase() ===
                          selectedPort.toUpperCase();
                        const isChecked = checkedScrips[scrip.LicNo] !== undefined;
                        const currentUsed = isChecked ? Number(checkedScrips[scrip.LicNo] || 0) : 0;
                        const availBal = Number(scrip.BalanceAmount || 0);
                        const afterBal = Math.max(0, Number((availBal - currentUsed).toFixed(2)));

                        return (
                          <tr
                            key={scrip.LicNo}
                            onClick={() => {
                              if (isPortMatch && !isChecked) {
                                handleToggleScripCheckbox(scrip);
                              }
                            }}
                            className={`transition-colors ${
                              !isPortMatch
                                ? "bg-gray-100/80 opacity-55 cursor-not-allowed"
                                : isChecked
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
                                disabled={!isPortMatch}
                                onChange={() => handleToggleScripCheckbox(scrip)}
                                title={
                                  !isPortMatch
                                    ? `Locked: This scrip belongs to ${scrip.PortLocation} and cannot be checked when posting for ${selectedPort}`
                                    : "Check to attach this RoDTEP scrip"
                                }
                                className="w-4 h-4 accent-[#c8102e] cursor-pointer disabled:cursor-not-allowed"
                              />
                            </td>

                            <td className="py-2.5 px-3">
                              <span
                                className={`inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10.5px] font-bold ${
                                  String(scrip.PortLocation).toUpperCase() === "CHENNAI"
                                    ? "bg-purple-100 text-purple-900"
                                    : "bg-blue-100 text-blue-900"
                                }`}
                              >
                                {scrip.PortLocation}
                                {!isPortMatch && " 🔒"}
                              </span>
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
                      <th className="py-2.5 px-3">Port</th>
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
                          <td className="py-2.5 px-3">
                            <span
                              className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                                String(u.PortLocation).toUpperCase() === "CHENNAI"
                                  ? "bg-purple-100 text-purple-900"
                                  : "bg-blue-100 text-blue-900"
                              }`}
                            >
                              {u.PortLocation}
                            </span>
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
                    All RoDTEP / RoSCTL Scrips ({scrips.length})
                  </h3>
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
                      <th className="py-2.5 px-3">Port</th>
                      <th className="py-2.5 px-3">AP Inv #</th>
                      <th className="py-2.5 px-3 text-right">Scrip Amount</th>
                      <th className="py-2.5 px-3 text-right">Utilized (On-The-Fly)</th>
                      <th className="py-2.5 px-3 text-right">Balance (On-The-Fly)</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-black/[0.06] bg-white">
                    {scrips.map((s) => {
                      const bal = Number(s.BalanceAmount || 0);
                      const util = Number(s.UtilizedAmount || 0);
                      return (
                        <tr key={s.LicNo} className="hover:bg-[#faf9f7]">
                          <td className="py-2 px-3 font-mono-code font-bold">{s.LicNo}</td>
                          <td className="py-2 px-3 font-mono-code">{s.LicDate}</td>
                          <td className="py-2 px-3">{s.ExporterName}</td>
                          <td className="py-2 px-3 font-bold text-[11px]">{s.Scheme}</td>
                          <td className="py-2 px-3 font-mono-code">{s.PortCode}</td>
                          <td className="py-2 px-3">
                            <span
                              className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                                String(s.PortLocation).toUpperCase() === "CHENNAI"
                                  ? "bg-purple-100 text-purple-900"
                                  : "bg-blue-100 text-blue-900"
                              }`}
                            >
                              {s.PortLocation}
                            </span>
                          </td>
                          <td className="py-2 px-3 font-mono-code">
                            {s.APDocNum ? `#${s.APDocNum} (${s.NumAtCard || ""})` : "—"}
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
