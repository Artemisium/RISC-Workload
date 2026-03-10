import { useState, useCallback } from "react";

const MEMBERS = ["AK", "SK", "MR", "DX"];

const LOAN_GROUPS = {
  Construction: ["construction"],
  "Refi / Acquisition": ["refinanc", "refi", "acquisition"],
};

function classifyLoan(typeStr) {
  if (!typeStr) return "Other";
  const lower = typeStr.toLowerCase();
  for (const [group, keywords] of Object.entries(LOAN_GROUPS)) {
    if (keywords.some((k) => lower.includes(k))) return group;
  }
  return "Other";
}

function fuzzyMatch(header, keywords) {
  const h = (header || "").toLowerCase().replace(/[^a-z0-9]/g, "");
  return keywords.some((k) => h.includes(k));
}

function findColIndex(headers, keywords, excludeIndices = []) {
  for (let i = 0; i < headers.length; i++) {
    if (excludeIndices.includes(i)) continue;
    if (fuzzyMatch(headers[i], keywords)) return i;
  }
  return -1;
}

function getWorkWeeks(count = 3) {
  const today = new Date();
  const day = today.getDay();
  const mondayOffset = day === 0 ? 1 : -(day - 1);
  const thisMonday = new Date(today);
  thisMonday.setDate(today.getDate() + mondayOffset);
  thisMonday.setHours(0, 0, 0, 0);
  const weeks = [];
  for (let w = 0; w < count; w++) {
    const days = [];
    for (let d = 0; d < 5; d++) {
      const date = new Date(thisMonday);
      date.setDate(thisMonday.getDate() + w * 7 + d);
      days.push(date);
    }
    weeks.push(days);
  }
  return weeks;
}

function dateKey(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

function parseExcelDate(val) {
  if (!val) return null;
  if (typeof val === "number") {
    const epoch = new Date(1899, 11, 30);
    epoch.setDate(epoch.getDate() + val);
    return epoch;
  }
  const d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}

function fmtDate(d) {
  const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  return `${months[d.getMonth()]} ${d.getDate()}`;
}

function fmtWeekLabel(days) {
  return `${fmtDate(days[0])} – ${fmtDate(days[4])}`;
}

function businessDaysBetween(from, to) {
  const start = new Date(from); start.setHours(0,0,0,0);
  const end = new Date(to); end.setHours(0,0,0,0);
  if (end <= start) return 0;
  let count = 0;
  const cur = new Date(start);
  while (cur < end) {
    cur.setDate(cur.getDate() + 1);
    const dow = cur.getDay();
    if (dow !== 0 && dow !== 6) count++;
  }
  return count;
}

const GROUP_COLORS = {
  Construction: { bg: "#FFF3E0", border: "#F57C00", text: "#E65100", badge: "#FF9800" },
  "Refi / Acquisition": { bg: "#E3F2FD", border: "#1E88E5", text: "#0D47A1", badge: "#2196F3" },
  PM: { bg: "#E8F5E9", border: "#43A047", text: "#1B5E20", badge: "#2E7D32" },
  Other: { bg: "#F3E5F5", border: "#8E24AA", text: "#4A148C", badge: "#AB47BC" },
};

const GROUPS = ["Construction", "Refi / Acquisition", "Other"];
const ALL_GROUPS = ["Construction", "Refi / Acquisition", "Other", "PM"];
const DAY_NAMES = ["Mon", "Tue", "Wed", "Thu", "Fri"];

export default function RISCWorkload() {
  const [loans, setLoans] = useState(null);
  const [error, setError] = useState(null);
  const [loading, setLoading] = useState(false);
  const [expandedMember, setExpandedMember] = useState(null);

  const processFile = useCallback(async (file) => {
    setLoading(true);
    setError(null);
    try {
      if (!window.XLSX) {
        await new Promise((res, rej) => {
          const s = document.createElement("script");
          s.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
          s.onload = res;
          s.onerror = () => rej(new Error("Failed to load SheetJS"));
          document.head.appendChild(s);
        });
      }
      const data = await file.arrayBuffer();
      const wb = window.XLSX.read(data, { type: "array", cellDates: false });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = window.XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      if (rows.length < 2) throw new Error("File appears empty");

      const headers = rows[0].map(String);
      const pmStatusCol = findColIndex(headers, ["openorclosed", "openorclose", "openclosed"]);
      const exclude = pmStatusCol !== -1 ? [pmStatusCol] : [];
      const teamCol = findColIndex(headers, ["team", "assign", "member", "initials", "processor", "analyst"], exclude);
      const typeCol = findColIndex(headers, ["type", "loan type", "loantype", "product"], exclude);
      const dateCol = findColIndex(headers, ["fund", "funding", "close", "closing", "fundingdate", "closedate"], exclude);
      const nameCol = findColIndex(headers, ["borrowername", "borrower"], exclude);
      const lenderCol = headers.findIndex((h, i) => {
        if (exclude.includes(i)) return false;
        const stripped = h.toLowerCase().replace(/[^a-z0-9]/g, "");
        return stripped === "client" || stripped === "lender" || stripped === "lendername";
      });
      const amountCol = findColIndex(headers, ["amount", "loan amount", "value", "principal", "size"], exclude);
      const checkInCol = findColIndex(headers, ["check in", "checkin", "next check", "nextcheck"], exclude);
      const statusCol = findColIndex(headers, ["status"], exclude);
      const addressCol = findColIndex(headers, ["address", "property", "location", "site"], exclude);

      if (teamCol === -1) throw new Error("Could not find a Team/Assigned column");
      if (typeCol === -1) throw new Error("Could not find a Loan Type column");

      const parsed = [];
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const member = String(row[teamCol] || "").trim().toUpperCase();
        if (!MEMBERS.includes(member)) continue;
        const rawType = String(row[typeCol] || "");
        const fundDate = dateCol !== -1 ? parseExcelDate(row[dateCol]) : null;
        const status = statusCol !== -1 ? String(row[statusCol] || "") : "";
        const pmStatus = pmStatusCol !== -1 ? String(row[pmStatusCol] || "").toLowerCase().trim() : "";
        const pmOpen = pmStatus === "open";
        const statusOpen = status.toLowerCase().includes("open");
        if (!statusOpen && !pmOpen) continue;
        const borrower = nameCol !== -1 ? String(row[nameCol] || "") : "";
        const amount = amountCol !== -1 ? row[amountCol] : null;
        const checkInDate = checkInCol !== -1 ? parseExcelDate(row[checkInCol]) : null;
        const address = addressCol !== -1 ? String(row[addressCol] || "") : "";
        const lender = lenderCol !== -1 ? String(row[lenderCol] || "") : "";
        parsed.push({
          member, loanType: rawType, group: classifyLoan(rawType),
          fundDate, fundDateKey: statusOpen ? (fundDate ? dateKey(fundDate) : null) : null,
          checkInDate, checkInDateKey: checkInDate ? dateKey(checkInDate) : null,
          pmOpen, statusOpen, status, borrower, address, lender,
          amount: typeof amount === "number" ? amount : parseFloat(String(amount).replace(/[^0-9.-]/g, "")) || 0,
        });
      }
      setLoans(parsed);
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  }, []);

  const handleDrop = useCallback((e) => { e.preventDefault(); const file = e.dataTransfer?.files?.[0]; if (file) processFile(file); }, [processFile]);
  const handleFileInput = useCallback((e) => { const file = e.target.files?.[0]; if (file) processFile(file); }, [processFile]);

  const weeks = getWorkWeeks(3);

  if (!loans) {
    return (
      <div style={S.container}>
        <div style={S.header}>
          <h1 style={S.title}>RISC Team Workload</h1>
          <p style={S.subtitle}>Upload the archive file to view assignments</p>
        </div>
        <div style={S.dropZone} onDrop={handleDrop} onDragOver={(e) => e.preventDefault()}>
          <div style={S.dropIcon}>📂</div>
          <p style={S.dropText}>Drop your Excel file here</p>
          <p style={S.dropHint}>or</p>
          <label style={S.fileButton}>
            Browse Files
            <input type="file" accept=".xlsx,.xls,.csv" onChange={handleFileInput} style={{ display: "none" }} />
          </label>
          {loading && <p style={S.loadingText}>Processing…</p>}
          {error && <p style={S.errorText}>{error}</p>}
        </div>
      </div>
    );
  }

  const today = new Date(); today.setHours(0,0,0,0);

  const memberData = MEMBERS.map((m) => {
    const myLoans = loans.filter((l) => l.member === m);
    const openLoans = myLoans.filter((l) => l.statusOpen);
    const counts = {};
    GROUPS.forEach((g) => (counts[g] = openLoans.filter((l) => l.group === g).length));
    const pmCount = myLoans.filter((l) => l.pmOpen).length;
    const dailyMap = {};
    weeks.flat().forEach((d) => {
      const dk = dateKey(d);
      const dayFundLoans = openLoans.filter((l) => l.fundDateKey === dk);
      const dayCheckInLoans = myLoans.filter((l) => l.pmOpen && l.checkInDateKey === dk);
      const dayCounts = {};
      GROUPS.forEach((g) => (dayCounts[g] = dayFundLoans.filter((l) => l.group === g).length));
      dayCounts["PM"] = dayCheckInLoans.length;
      dailyMap[dk] = dayCounts;
    });
    const rushLoans = openLoans
      .filter((l) => {
        if (!l.fundDate) return false;
        const fd = new Date(l.fundDate); fd.setHours(0,0,0,0);
        if (fd < today) return false;
        return businessDaysBetween(today, fd) <= 4;
      })
      .sort((a, b) => a.fundDate - b.fundDate);
    return { member: m, loans: myLoans, counts, pmCount, total: openLoans.length, dailyMap, rushLoans };
  });

  return (
    <div style={S.container} className="risc-container">
      <style>{RESPONSIVE_CSS}</style>
      <div className="risc-header" style={S.header}>
        <h1 style={S.title}>RISC Team Workload</h1>
        <div style={S.headerRight}>
          <div style={S.legend}>
            {ALL_GROUPS.map((g) => (
              <span key={g} style={S.legendItem}>
                <span style={{ ...S.legendDot, background: GROUP_COLORS[g].badge }} />
                {g}
              </span>
            ))}
          </div>
          <label style={S.reloadBtn}>
            ↻ Reload
            <input type="file" accept=".xlsx,.xls,.csv" onChange={handleFileInput} style={{ display: "none" }} />
          </label>
        </div>
      </div>
      {error && <p style={S.errorText}>{error}</p>}

      <div style={S.memberList}>
        {memberData.map(({ member, counts, pmCount, total, dailyMap, rushLoans, loans: memberLoans }) => {
          const isExpanded = expandedMember === member;
          const allCounts = { ...counts, PM: pmCount };
          return (
            <div key={member} className="member-row" style={S.memberRowOuter}>
              {/* Main workload card */}
              <div className="member-main" style={S.card}>
                <div style={S.cardHeader} onClick={() => setExpandedMember(isExpanded ? null : member)}>
                  <div style={S.memberRow}>
                    <span style={S.memberBadge}>{member}</span>
                    <span style={S.totalBadge}>{total} loans</span>
                  </div>
                  <span style={S.expandArrow}>{isExpanded ? "▲" : "▼"}</span>
                </div>
                <div style={S.groupRow}>
                  {ALL_GROUPS.map((g) => (
                    <div key={g} style={{ ...S.groupPill, background: GROUP_COLORS[g].bg, borderColor: GROUP_COLORS[g].border }}>
                      <span style={{ ...S.groupCount, color: GROUP_COLORS[g].text }}>{allCounts[g]}</span>
                      <span style={{ ...S.groupLabel, color: GROUP_COLORS[g].text }}>{g}</span>
                    </div>
                  ))}
                </div>
                <div style={S.weeksSection}>
                  {weeks.map((weekDays, wi) => {
                    const weekFundTotal = weekDays.reduce((sum, d) => {
                      const dc = dailyMap[dateKey(d)];
                      return sum + GROUPS.reduce((s, g) => s + (dc?.[g] || 0), 0);
                    }, 0);
                    const weekPmTotal = weekDays.reduce((sum, d) => {
                      const dc = dailyMap[dateKey(d)];
                      return sum + (dc?.["PM"] || 0);
                    }, 0);
                    return (
                      <div key={wi}>
                        <div style={S.weekHeader}>
                          <span style={S.weekLabel}>{wi === 0 ? "This Week" : wi === 1 ? "Next Week" : "Week After"}</span>
                          <span style={S.weekRange}>{fmtWeekLabel(weekDays)}</span>
                          <div style={S.weekTotals}>
                            {weekFundTotal > 0 && <span style={S.weekTotal}>{weekFundTotal} funding</span>}
                            {weekPmTotal > 0 && <span style={S.weekTotalPm}>{weekPmTotal} PM follow-up{weekPmTotal !== 1 ? "s" : ""}</span>}
                          </div>
                        </div>
                        <div style={S.dayGrid}>
                          {weekDays.map((d, di) => {
                            const dk = dateKey(d);
                            const dc = dailyMap[dk] || {};
                            const dayTotal = ALL_GROUPS.reduce((s, g) => s + (dc[g] || 0), 0);
                            const isToday = dateKey(new Date()) === dk;
                            const isPast = d < today;
                            return (
                              <div key={di} style={{ ...S.dayCell, ...(isToday ? S.todayCell : {}), ...(isPast && !isToday ? S.pastCell : {}) }}>
                                <div style={{ ...S.dayName, ...(isToday ? { color: "#1565C0", fontWeight: 700 } : {}) }}>{DAY_NAMES[di]}</div>
                                <div style={S.dayDate}>{fmtDate(d)}</div>
                                {dayTotal === 0 ? (
                                  <div style={S.emptyDay}>—</div>
                                ) : (
                                  <div style={S.dayPills}>
                                    {ALL_GROUPS.map((g) => dc[g] > 0 && (
                                      <span key={g} style={{ ...S.dayPill, background: GROUP_COLORS[g].badge }} title={g}>{dc[g]}</span>
                                    ))}
                                  </div>
                                )}
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    );
                  })}
                </div>
                {isExpanded && (
                  <div style={S.loanList}>
                    <div style={S.loanListHeader}>All Loans</div>
                    {memberLoans.length === 0 ? (
                      <p style={{ color: "#999", fontSize: 13, padding: "8px 0" }}>No loans found</p>
                    ) : (
                      memberLoans.sort((a, b) => {
                        if (a.fundDate && b.fundDate) return a.fundDate - b.fundDate;
                        if (a.fundDate) return -1;
                        return 1;
                      }).map((l, i) => (
                        <div key={i} style={S.loanRow}>
                          <span style={{ ...S.loanDot, background: GROUP_COLORS[l.group].badge }} />
                          <span style={S.loanBorrower}>{l.borrower || "—"}</span>
                          <span style={S.loanType}>{l.loanType}</span>
                          {l.pmOpen && l.checkInDate && <span style={{ ...S.loanDate, color: GROUP_COLORS.PM.text }}>✓ {fmtDate(l.checkInDate)}</span>}
                          <span style={S.loanDate}>{l.fundDate ? fmtDate(l.fundDate) : "No date"}</span>
                          {l.pmOpen && <span style={{ ...S.loanStatus, background: GROUP_COLORS.PM.bg, color: GROUP_COLORS.PM.text }}>PM</span>}
                          {l.status && <span style={S.loanStatus}>{l.status}</span>}
                        </div>
                      ))
                    )}
                  </div>
                )}
              </div>

              {/* Rush card */}
              <div className="member-rush" style={S.rushCard}>
                <div style={S.rushBar}>
                  <span style={S.rushIcon}>⚡</span>
                  <span style={S.rushTitle}>Rush</span>
                  <span style={S.rushBadge}>{rushLoans.length}</span>
                </div>
                {rushLoans.length === 0 ? (
                  <div style={S.rushEmpty}>No rush loans</div>
                ) : (
                  <div style={S.rushList}>
                    {rushLoans.map((l, i) => {
                      const bd = businessDaysBetween(today, l.fundDate);
                      return (
                        <div key={i} style={S.rushItem}>
                          <div style={S.rushItemTop}>
                            <span style={{ ...S.rushDot, background: GROUP_COLORS[l.group].badge }} />
                            <span style={S.rushBorrower}>{l.borrower || "—"}</span>
                          </div>
                          {l.address && <div style={S.rushAddress}>{l.address}</div>}
                          <div style={S.rushMeta}>
                            <span style={S.rushDate}>{fmtDate(l.fundDate)}</span>
                            <span style={{
                              ...S.rushDays,
                              background: bd <= 1 ? "#FFEBEE" : bd <= 2 ? "#FFF3E0" : "#FFF8E1",
                              color: bd <= 1 ? "#C62828" : bd <= 2 ? "#E65100" : "#F57F17",
                            }}>
                              {bd === 0 ? "TODAY" : bd === 1 ? "1 day" : `${bd} days`}
                            </span>
                            {l.lender && <span style={S.rushLender}>{l.lender}</span>}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

const RESPONSIVE_CSS = `
  .risc-header { display: flex; flex-direction: column; gap: 8px; }
  .member-row { display: flex; flex-direction: column; gap: 12px; }
  .member-main { flex: 1; min-width: 0; }
  .member-rush { }
  @media (min-width: 860px) {
    .risc-container { padding: 20px 24px !important; }
    .risc-header { flex-direction: row !important; align-items: center !important; justify-content: space-between !important; }
    .member-row { flex-direction: row !important; align-items: flex-start !important; }
    .member-main { flex: 0 1 55% !important; }
    .member-rush { flex: 1 1 0 !important; min-width: 260px !important; flex-shrink: 0 !important; position: sticky !important; top: 20px !important; }
  }
  @media (min-width: 1200px) {
    .member-main { flex: 0 1 52% !important; }
  }
`;

const S = {
  container: { fontFamily: "'IBM Plex Sans', 'Segoe UI', system-ui, sans-serif", background: "#F5F5F0", minHeight: "100vh", padding: "12px", color: "#1A1A1A" },
  header: { marginBottom: 16 },
  title: { fontSize: 20, fontWeight: 800, letterSpacing: "-0.5px", margin: 0, color: "#111" },
  subtitle: { fontSize: 13, color: "#777", marginTop: 2 },
  headerRight: { display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" },
  legend: { display: "flex", gap: 8, fontSize: 11, color: "#555", flexWrap: "wrap" },
  legendItem: { display: "flex", alignItems: "center", gap: 3 },
  legendDot: { width: 7, height: 7, borderRadius: "50%", display: "inline-block" },
  reloadBtn: { fontSize: 12, fontWeight: 600, color: "#444", background: "#E8E8E4", border: "1px solid #D0D0CC", borderRadius: 6, padding: "5px 12px", cursor: "pointer" },
  dropZone: { border: "2px dashed #C0C0BC", borderRadius: 16, padding: "48px 20px", textAlign: "center", background: "#FAFAF8", maxWidth: 400, margin: "40px auto" },
  dropIcon: { fontSize: 40, marginBottom: 10 },
  dropText: { fontSize: 15, fontWeight: 600, color: "#333", margin: "0 0 4px" },
  dropHint: { fontSize: 12, color: "#999", margin: "4px 0 10px" },
  fileButton: { display: "inline-block", padding: "10px 24px", background: "#1A1A1A", color: "#fff", borderRadius: 8, fontSize: 14, fontWeight: 600, cursor: "pointer" },
  loadingText: { marginTop: 16, color: "#666", fontSize: 13 },
  errorText: { color: "#C62828", fontSize: 13, marginTop: 8, fontWeight: 500 },

  memberList: { display: "flex", flexDirection: "column", gap: 16 },
  memberRowOuter: {},

  card: { background: "#fff", borderRadius: 10, border: "1px solid #E0E0DC", overflow: "hidden" },
  cardHeader: { display: "flex", justifyContent: "space-between", alignItems: "center", padding: "12px 14px 8px", cursor: "pointer", userSelect: "none" },
  memberRow: { display: "flex", alignItems: "center", gap: 8 },
  memberBadge: { fontSize: 18, fontWeight: 800, letterSpacing: "1px", color: "#111" },
  totalBadge: { fontSize: 11, fontWeight: 600, color: "#666", background: "#F0F0EC", borderRadius: 10, padding: "2px 8px" },
  expandArrow: { fontSize: 11, color: "#999" },
  groupRow: { display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 5, padding: "0 14px 12px" },
  groupPill: { borderRadius: 6, border: "1px solid", padding: "6px 4px", textAlign: "center" },
  groupCount: { display: "block", fontSize: 18, fontWeight: 800, lineHeight: 1 },
  groupLabel: { display: "block", fontSize: 9, fontWeight: 600, marginTop: 3, textTransform: "uppercase", letterSpacing: "0.2px" },
  weeksSection: { padding: "0 14px 12px", display: "flex", flexDirection: "column", gap: 8 },
  weekHeader: { display: "flex", alignItems: "center", gap: 6, marginBottom: 4, flexWrap: "wrap" },
  weekLabel: { fontSize: 10, fontWeight: 700, color: "#333", textTransform: "uppercase", letterSpacing: "0.5px" },
  weekRange: { fontSize: 10, color: "#999" },
  weekTotals: { display: "flex", gap: 4, marginLeft: "auto", flexWrap: "wrap", justifyContent: "flex-end" },
  weekTotal: { fontSize: 9, fontWeight: 700, color: "#fff", background: "#555", borderRadius: 8, padding: "1px 7px", whiteSpace: "nowrap" },
  weekTotalPm: { fontSize: 9, fontWeight: 700, color: GROUP_COLORS.PM.text, background: GROUP_COLORS.PM.bg, border: `1px solid ${GROUP_COLORS.PM.border}`, borderRadius: 8, padding: "1px 7px", whiteSpace: "nowrap" },
  dayGrid: { display: "grid", gridTemplateColumns: "repeat(5, 1fr)", gap: 3 },
  dayCell: { background: "#FAFAF8", borderRadius: 5, padding: "4px 2px", textAlign: "center", minHeight: 48, border: "1px solid #EEEEEA" },
  todayCell: { background: "#E3F2FD", border: "1.5px solid #90CAF9" },
  pastCell: { opacity: 0.5 },
  dayName: { fontSize: 9, fontWeight: 600, color: "#888", textTransform: "uppercase", letterSpacing: "0.3px" },
  dayDate: { fontSize: 9, color: "#aaa", marginBottom: 3 },
  emptyDay: { fontSize: 12, color: "#ccc", marginTop: 2 },
  dayPills: { display: "flex", gap: 2, justifyContent: "center", flexWrap: "wrap" },
  dayPill: { color: "#fff", fontSize: 10, fontWeight: 700, borderRadius: 3, padding: "1px 5px", lineHeight: "16px" },
  loanList: { borderTop: "1px solid #EEEEEA", padding: "10px 14px 14px", maxHeight: 220, overflowY: "auto" },
  loanListHeader: { fontSize: 10, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.5px", color: "#888", marginBottom: 6 },
  loanRow: { display: "flex", alignItems: "center", gap: 6, padding: "4px 0", borderBottom: "1px solid #F5F5F0", fontSize: 12, flexWrap: "wrap" },
  loanDot: { width: 6, height: 6, borderRadius: "50%", flexShrink: 0 },
  loanBorrower: { fontWeight: 600, color: "#222", flex: 1, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
  loanType: { color: "#777", fontSize: 11, flex: "0 0 auto" },
  loanDate: { color: "#555", fontSize: 11, fontWeight: 500, textAlign: "right", flex: "0 0 auto" },
  loanStatus: { fontSize: 9, fontWeight: 600, color: "#888", background: "#F0F0EC", borderRadius: 4, padding: "1px 5px" },

  // Rush card
  rushCard: { background: "#fff", borderRadius: 10, border: "1px solid #E0E0DC", padding: 14, overflow: "hidden" },
  rushBar: { display: "flex", alignItems: "center", gap: 5, marginBottom: 8, paddingBottom: 8, borderBottom: "1px solid #F0F0EC" },
  rushIcon: { fontSize: 14 },
  rushTitle: { fontSize: 14, fontWeight: 800, color: "#C62828", letterSpacing: "-0.3px" },
  rushBadge: { marginLeft: "auto", fontSize: 12, fontWeight: 800, color: "#C62828", background: "#FFEBEE", borderRadius: 10, padding: "0px 8px" },
  rushEmpty: { fontSize: 12, color: "#bbb", textAlign: "center", padding: "16px 0" },
  rushList: { display: "flex", flexDirection: "column", gap: 0 },
  rushItem: { padding: "8px 0", borderBottom: "1px solid #F5F5F0" },
  rushItemTop: { display: "flex", alignItems: "center", gap: 5, marginBottom: 2 },
  rushDot: { width: 6, height: 6, borderRadius: "50%", flexShrink: 0 },
  rushBorrower: { fontSize: 12, fontWeight: 700, color: "#111", flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
  rushLender: { fontSize: 10, color: "#999", fontWeight: 500, marginLeft: "auto" },
  rushAddress: { fontSize: 10, color: "#777", marginLeft: 11, marginBottom: 3, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
  rushMeta: { display: "flex", alignItems: "center", gap: 5, marginLeft: 11 },
  rushDate: { fontSize: 10, fontWeight: 600, color: "#444" },
  rushDays: { fontSize: 9, fontWeight: 700, borderRadius: 4, padding: "1px 5px" },
};
