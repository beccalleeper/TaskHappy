import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

// ── helpers ───────────────────────────────────────────────────────────────────
const uid = () => Math.random().toString(36).slice(2, 10);
const today = () => new Date().toISOString().split("T")[0];
const daysBetween = (a, b) => {
  if (!a || !b) return null;
  return Math.round((new Date(b) - new Date(a)) / 86400000);
};

// Compute average actual cadence from completion history (need ≥2 entries)
const avgCadence = (history) => {
  if (!history || history.length < 2) return null;
  const sorted = [...history].sort((a, b) => new Date(a.date) - new Date(b.date));
  const gaps = [];
  for (let i = 1; i < sorted.length; i++) {
    const g = daysBetween(sorted[i - 1].date, sorted[i].date);
    if (g > 0) gaps.push(g);
  }
  if (!gaps.length) return null;
  return Math.round(gaps.reduce((s, g) => s + g, 0) / gaps.length);
};

// Effective frequency: blend goal with reality (weight toward reality as history grows)
const effectiveFreq = (task) => {
  const goal = task.frequencyDays;
  const avg = avgCadence(task.history);
  if (!goal && !avg) return null;
  if (!goal) return avg;
  if (!avg) return goal;
  const entries = (task.history || []).length;
  // After 10+ completions, weight 60% reality / 40% goal
  const realWeight = Math.min(0.6, entries * 0.06);
  return Math.round(goal * (1 - realWeight) + avg * realWeight);
};

const lastCompletedDate = (task) => {
  if (task.history && task.history.length > 0) {
    return [...task.history].sort((a, b) => new Date(b.date) - new Date(a.date))[0].date;
  }
  return task.lastCompleted || null;
};

const daysOverdue = (task) => {
  const last = lastCompletedDate(task);
  const freq = effectiveFreq(task);
  if (!last) return 999;
  if (!freq) return 0;
  return Math.round(daysBetween(last, today()) - freq);
};

const nextDueDate = (task) => {
  const last = lastCompletedDate(task);
  const freq = effectiveFreq(task);
  if (!last || !freq) return null;
  const d = new Date(last);
  d.setDate(d.getDate() + freq);
  return d.toISOString().split("T")[0];
};

// Status color/label is based on how far through its own frequency cycle a task
// is (percent elapsed), not on absolute day counts. This way a task done every
// 2 days that was just completed today reads as green/on-track ("Due in 2d"),
// while a yearly task 2 days from due reads as "coming due" rather than green.
const overdueLabel = (task) => {
  const freq = effectiveFreq(task);
  const d = daysOverdue(task); // positive = overdue, 0 = due today, negative = due in future
  if (!lastCompletedDate(task)) return { label: "Never done", color: "#FF6B6B" };
  if (!freq) return { label: "No schedule", color: "#aaa" };

  const percentElapsed = (freq + d) / freq; // 0 = just completed, 1 = due today, >1 = overdue

  if (percentElapsed <= 0.5) return { label: `Due in ${Math.abs(d)}d`, color: "#06D6A0" }; // first half of cycle — green, on track
  if (percentElapsed < 1) return { label: `Due in ${Math.abs(d)}d`, color: "#B6E388" }; // back half of cycle — coming due, light green/yellow
  if (percentElapsed === 1) return { label: "Due today!", color: "#FFD93D" }; // due exactly today — yellow
  if (percentElapsed < 1.25) return { label: `${d}d overdue`, color: "#FFC857" }; // slightly overdue — light orange/yellow
  if (percentElapsed < 1.75) return { label: `${d}d overdue`, color: "#FF9F43" }; // overdue — orange
  return { label: `${d}d overdue`, color: "#FF6B6B" }; // well overdue — red
};

const nextAssignee = (task, users) => {
  if (task.assignMode === "fixed") return task.assignedTo;
  if (task.assignMode === "any") return null;
  const last = task.history && task.history.length > 0
    ? [...task.history].sort((a, b) => new Date(b.date) - new Date(a.date))[0].completedBy
    : task.lastCompletedBy;
  if (!last || users.length < 2) return users[0] || null;
  const idx = users.indexOf(last);
  return users[(idx + 1) % users.length];
};

// Status slider (acts as a "days until due" window):
// 0=overdue only (excludes due-today), 1=due tomorrow or sooner, 2=due in 2 days or sooner,
// 3=due in 3 days or sooner, 4=due in 4 days or sooner, 5=due within the week, 6=all
const STATUS_FILTER_STEPS = [
  { label: "Overdue", short: "Overdue" },
  { label: "Due Today", short: "Today" },
  { label: "Due Tomorrow", short: "Tomorrow" },
  { label: "Due in 2 Days", short: "2 Days" },
  { label: "Due in 3 Days", short: "3 Days" },
  { label: "Due This Week", short: "This Week" },
  { label: "All Tasks", short: "All" },
];

// daysOverdue: positive = overdue, 0 = due today, negative = due in future (-1 = tomorrow, etc.)
// Step 0 ("Overdue") shows ONLY tasks that are strictly overdue (d > 0), or never done.
// Steps 1-5 widen the window to include tasks due today through that many days out,
// while overdue tasks remain visible at every step since they're always due.
const matchesStatusFilter = (task, step) => {
  if (step === 6) return true; // All
  const d = daysOverdue(task);
  if (d === 999) return true; // never-done tasks always show
  if (step === 0) return d > 0; // Overdue only — excludes due-today (d === 0)
  if (d >= 0) return true; // overdue or due-today tasks show at steps 1-5
  switch (step) {
    case 1: return d >= -1; // Due Today + Tomorrow
    case 2: return d >= -2; // Due within 2 days
    case 3: return d >= -3; // Due within 3 days
    case 4: return d >= -4; // Due within 4 days
    case 5: return d >= -7; // Due within the week
    default: return true;
  }
};

// ── default data ──────────────────────────────────────────────────────────────
const DEFAULT_CATEGORIES = [
  { id: "car-stuff", name: "Car Stuff", emoji: "🚗", color: "#4ECDC4" },
  { id: "personal-care", name: "Personal Care", emoji: "✨", color: "#C77DFF" },
  { id: "housework", name: "House Work", emoji: "🏠", color: "#FF6B6B" },
];

const TASK_SUGGESTIONS = {
  "Car Stuff": [
    { name: "Oil change", freq: 90 }, { name: "Tire rotation", freq: 180 }, { name: "Car wash", freq: 14 },
    { name: "Update car registration", freq: 365 }, { name: "Check tire pressure", freq: 30 },
    { name: "Review car insurance", freq: 180 },
  ],
  "Personal Care": [
    { name: "Journal", freq: 1 }, { name: "Meditate", freq: 1 }, { name: "Therapy session", freq: 14 },
    { name: "Workout", freq: 2 }, { name: "Run / walk", freq: 2 }, { name: "Stretch", freq: 1 },
    { name: "Dentist appointment", freq: 180 }, { name: "Doctor checkup", freq: 365 }, { name: "Take vitamins", freq: 1 },
    { name: "Haircut", freq: 42 }, { name: "Nail trim / manicure", freq: 14 }, { name: "Skincare routine", freq: 1 },
    { name: "Eyebrows", freq: 21 }, { name: "Eye exam", freq: 365 }, { name: "Refill prescriptions", freq: 30 },
  ],
  "House Work": [
    { name: "Vacuum", freq: 7 }, { name: "Mop floors", freq: 7 }, { name: "Wipe countertops", freq: 3 },
    { name: "Clean toilet", freq: 7 }, { name: "Take out trash", freq: 7 }, { name: "Do laundry", freq: 7 },
    { name: "Wash dishes", freq: 1 }, { name: "Clean oven", freq: 90 }, { name: "Wash windows", freq: 90 },
    { name: "Deep scrub bathrooms", freq: 30 }, { name: "Change HVAC filter", freq: 90 },
    { name: "Test smoke detectors", freq: 180 }, { name: "Declutter a drawer", freq: 30 },
    { name: "Mow lawn", freq: 7 }, { name: "Water plants", freq: 2 }, { name: "Weed garden beds", freq: 14 },
  ],
  "Adulting": [
    { name: "Pay electric bill", freq: 30 }, { name: "Pay rent / mortgage", freq: 30 }, { name: "Review budget", freq: 30 },
    { name: "Check credit score", freq: 90 }, { name: "Review subscriptions", freq: 90 }, { name: "File taxes", freq: 365 },
    { name: "Review health insurance", freq: 365 }, { name: "Renew passport", freq: 3650 }, { name: "Register to vote", freq: 365 },
  ],
};

const FREQ_PRESETS = [
  { label: "Daily", days: 1 }, { label: "Every 2 days", days: 2 },
  { label: "Weekly", days: 7 }, { label: "Every 2 weeks", days: 14 },
  { label: "Monthly", days: 30 }, { label: "Every 3 months", days: 90 },
  { label: "Every 6 months", days: 180 }, { label: "Yearly", days: 365 },
];

const freqLabel = (days) => {
  if (!days) return "No schedule";
  const p = FREQ_PRESETS.find((f) => f.days === days);
  if (p) return p.label;
  if (days < 7) return `Every ${days} days`;
  if (days < 30) return `Every ${Math.round(days / 7)} weeks`;
  if (days < 365) return `Every ${Math.round(days / 30)} months`;
  return `Every ${Math.round(days / 365)} years`;
};

// ── CSV/XLSX import/export ────────────────────────────────────────────────────
const IMPORT_HEADERS = ["Task Name", "Category", "Frequency (Days)", "Last Completed (YYYY-MM-DD)", "Assignment Mode (any/fixed/alternate)", "Assigned To"];
const IMPORT_SAMPLE_ROW = ["Vacuum living room", "House Work", "7", "2026-06-08", "alternate", "Becca"];

const csvEscape = (val) => {
  const s = String(val ?? "");
  if (s.includes(",") || s.includes('"') || s.includes("\n")) {
    return `"${s.replace(/"/g, '""')}"`;
  }
  return s;
};

const downloadCSVTemplate = () => {
  const rows = [IMPORT_HEADERS, IMPORT_SAMPLE_ROW];
  const csv = rows.map((r) => r.map(csvEscape).join(",")).join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "taskhappy_import_template.csv";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
};

const downloadXLSXTemplate = () => {
  const ws = XLSX.utils.aoa_to_sheet([IMPORT_HEADERS, IMPORT_SAMPLE_ROW]);
  ws["!cols"] = IMPORT_HEADERS.map((h) => ({ wch: Math.max(h.length, 16) }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Tasks");
  const wbArray = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbArray], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "taskhappy_import_template.xlsx";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
};

// Basic CSV line parser that handles quoted fields containing commas
const parseCSVLine = (line) => {
  const result = [];
  let cur = "";
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (inQuotes) {
      if (ch === '"' && line[i + 1] === '"') { cur += '"'; i++; }
      else if (ch === '"') { inQuotes = false; }
      else { cur += ch; }
    } else {
      if (ch === '"') inQuotes = true;
      else if (ch === ",") { result.push(cur); cur = ""; }
      else cur += ch;
    }
  }
  result.push(cur);
  return result.map((s) => s.trim());
};

// Converts a 2D array of cell values (rows of columns, including header row)
// into row objects matching IMPORT_HEADERS. Skips the header row and the
// sample row (matched by exact task name) if present.
const rowsToTaskObjects = (rows) => {
  if (rows.length < 2) return [];
  const dataRows = rows.slice(1); // skip header
  return dataRows
    .map((cols) => (cols || []).map((c) => (c === null || c === undefined ? "" : String(c).trim())))
    .filter((cols) => cols[0] && cols[0].trim() && cols[0].trim() !== IMPORT_SAMPLE_ROW[0])
    .map((cols) => ({
      name: cols[0] || "",
      category: cols[1] || "",
      frequencyDays: cols[2] || "",
      lastCompleted: cols[3] || "",
      assignMode: (cols[4] || "any").trim().toLowerCase(),
      assignedTo: cols[5] || "",
    }));
};

// Validates the raw 2D array of cells before conversion and returns a specific,
// actionable error message if something looks wrong — or null if it's fine.
const validateImportRows = (rows) => {
  if (!rows || rows.length === 0) {
    return "The file appears to be empty. Please use the downloaded template and don't remove the header row.";
  }
  if (rows.length === 1) {
    return "Only a header row was found — there are no task rows below it. Add at least one row of task data and try again.";
  }
  const header = (rows[0] || []).map((c) => String(c ?? "").trim());
  const firstHeader = (header[0] || "").toLowerCase();
  if (!firstHeader.includes("task")) {
    return `The first row doesn't look like the template header (expected something like "${IMPORT_HEADERS[0]}", found "${header[0] || "(empty)"}"). Make sure you're uploading the downloaded template with its header row intact, and that your data starts on row 2.`;
  }
  if (header.length < IMPORT_HEADERS.length) {
    return `This file only has ${header.length} column${header.length === 1 ? "" : "s"}, but the template has ${IMPORT_HEADERS.length} (${IMPORT_HEADERS.join(", ")}). Please use the downloaded template and don't remove any columns.`;
  }
  return null;
};

// Parses CSV text into a 2D array of raw cell values (including header row).
const parseCSVRows = (text) => {
  const lines = text.split(/\r?\n/).filter((l) => l.trim().length > 0);
  return lines.map(parseCSVLine);
};

// Parses an .xlsx/.xls ArrayBuffer into a 2D array of raw cell values (including header row).
const parseXLSXRows = (arrayBuffer) => {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  if (!wb.SheetNames || wb.SheetNames.length === 0) {
    throw new Error("This Excel file doesn't contain any sheets.");
  }
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: "" });
};

// ── Confetti ──────────────────────────────────────────────────────────────────
function Confetti({ x, y }) {
  const pieces = Array.from({ length: 14 }, (_, i) => ({
    id: i,
    color: ["#FFD93D","#FF6B6B","#4ECDC4","#C77DFF","#06D6A0","#FF9F43","#74B9FF"][i % 7],
    angle: (i / 14) * 360,
    dist: 45 + Math.random() * 35,
  }));
  return (
    <div style={{ position: "fixed", left: x, top: y, pointerEvents: "none", zIndex: 9999 }}>
      {pieces.map((p) => (
        <div key={p.id} style={{
          position: "absolute", width: 8, height: 8, borderRadius: "2px",
          background: p.color, animation: `confettiFly 0.75s ease-out forwards`,
          "--angle": `${p.angle}deg`, "--dist": `${p.dist}px`,
        }} />
      ))}
    </div>
  );
}

// ── TaskHappy Mark ────────────────────────────────────────────────────────────
// A checklist square (rounded purple outline, green check, coral dot, lavender bars)
// with a happy-face badge in the corner — matches the TaskHappy app icon.
function TaskHappyMark({ size = 40 }) {
  return (
    <svg width={size} height={size} viewBox="0 0 100 100" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ flexShrink: 0 }}>
      {/* Checklist square outline */}
      <rect x="8" y="8" width="68" height="68" rx="18" stroke="#C77DFF" strokeWidth="9" fill="none" />
      {/* Green checkmark */}
      <path d="M22 38 L32 48 L52 26" stroke="#5FD9A4" strokeWidth="9" strokeLinecap="round" strokeLinejoin="round" fill="none" />
      {/* Lavender bar (top) */}
      <rect x="46" y="32" width="22" height="9" rx="4.5" fill="#DCC9F7" />
      {/* Coral dot */}
      <circle cx="27" cy="60" r="9" fill="#FF8A80" />
      {/* Lavender bar (bottom) */}
      <rect x="42" y="55" width="14" height="9" rx="4.5" fill="#DCC9F7" />
      {/* Happy face badge */}
      <circle cx="76" cy="76" r="22" fill="#FFD15C" stroke="#FFFEF7" strokeWidth="3" />
      <path d="M65 70 a5 5 0 0 1 9 0" stroke="#1A1A2E" strokeWidth="3.5" strokeLinecap="round" fill="none" />
      <path d="M79 70 a5 5 0 0 1 9 0" stroke="#1A1A2E" strokeWidth="3.5" strokeLinecap="round" fill="none" />
      <path d="M68 80 a9 7 0 0 0 16 0" stroke="#1A1A2E" strokeWidth="3.5" strokeLinecap="round" fill="none" />
    </svg>
  );
}

// ── Main App ──────────────────────────────────────────────────────────────────
export default function App() {
  const [categories, setCategories] = useState(() => { try { return JSON.parse(localStorage.getItem("tm_cats")) || DEFAULT_CATEGORIES; } catch { return DEFAULT_CATEGORIES; } });
  const [tasks, setTasks] = useState(() => { try { return JSON.parse(localStorage.getItem("tm_tasks")) || []; } catch { return []; } });
  const [users, setUsers] = useState(() => { try { return JSON.parse(localStorage.getItem("tm_users")) || ["Me"]; } catch { return ["Me"]; } });
  const [view, setView] = useState("dashboard");
  const [selectedCat, setSelectedCat] = useState("all");
  const [filterUser, setFilterUser] = useState("all");
  const [statusFilter, setStatusFilter] = useState(6); // 0=overdue,1=today,2=tomorrow,3=2days,4=3days,5=week,6=all
  const [groupBy, setGroupBy] = useState("category");
  const [viewMode, setViewMode] = useState("cutesy"); // cutesy | compact
  const [showAddTask, setShowAddTask] = useState(false);
  const [showAddCat, setShowAddCat] = useState(false);
  const [editTask, setEditTask] = useState(null);
  const [historyTask, setHistoryTask] = useState(null);
  const [completeTask, setCompleteTask] = useState(null); // task awaiting date input
  const [confetti, setConfetti] = useState([]);
  const [animatingIds, setAnimatingIds] = useState(new Set());

  useEffect(() => { localStorage.setItem("tm_cats", JSON.stringify(categories)); }, [categories]);
  useEffect(() => { localStorage.setItem("tm_tasks", JSON.stringify(tasks)); }, [tasks]);
  useEffect(() => { localStorage.setItem("tm_users", JSON.stringify(users)); }, [users]);

  const fireConfetti = (e) => {
    const rect = e.currentTarget.getBoundingClientRect();
    const id = uid();
    setConfetti((c) => [...c, { id, x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 }]);
    setTimeout(() => setConfetti((c) => c.filter((x) => x.id !== id)), 900);
  };

  // Called when user taps check button — open date prompt
  const handleCheckClick = (task, e) => {
    if (animatingIds.has(task.id)) return;
    fireConfetti(e);
    setCompleteTask({ task, buttonEvent: e });
  };

  // Called after user confirms date in CompleteModal
  const handleConfirmComplete = (task, completionDate, completedBy) => {
    setAnimatingIds((s) => new Set([...s, task.id]));
    const historyEntry = { id: uid(), date: completionDate, completedBy };
    setTimeout(() => {
      setTasks((ts) =>
        ts.map((t) =>
          t.id === task.id
            ? {
                ...t,
                lastCompleted: completionDate,
                lastCompletedBy: completedBy,
                completions: (t.completions || 0) + 1,
                history: [...(t.history || []), historyEntry],
              }
            : t
        )
      );
      setAnimatingIds((s) => { const n = new Set(s); n.delete(task.id); return n; });
    }, 850);
    setCompleteTask(null);
  };

  const updateTaskHistory = (taskId, newHistory) => {
    setTasks((ts) =>
      ts.map((t) => {
        if (t.id !== taskId) return t;
        const sorted = [...newHistory].sort((a, b) => new Date(b.date) - new Date(a.date));
        return { ...t, history: newHistory, lastCompleted: sorted[0]?.date || null, completions: newHistory.length };
      })
    );
  };

  // Import rows parsed from a CSV/XLSX upload. Each row: { name, category,
  // frequencyDays, lastCompleted, assignMode, assignedTo }
  // Categories that don't exist yet are created automatically (default styling).
  const handleImportTasks = (rows) => {
    let cats = [...categories];
    const newTasks = [];

    rows.forEach((row) => {
      if (!row.name || !row.name.trim()) return;

      // Match category by name (case-insensitive), or create a new one
      let cat = cats.find((c) => c.name.toLowerCase() === (row.category || "").toLowerCase());
      if (!cat && row.category && row.category.trim()) {
        cat = {
          id: uid(),
          name: row.category.trim(),
          emoji: "📌",
          color: "#FFD93D",
        };
        cats = [...cats, cat];
      }

      const assignMode = ["any", "fixed", "alternate"].includes(row.assignMode) ? row.assignMode : "any";
      const assignedTo = users.includes(row.assignedTo) ? row.assignedTo : (users[0] || "");
      const lastCompleted = row.lastCompleted && row.lastCompleted.trim() ? row.lastCompleted.trim() : "";
      const history = lastCompleted ? [{ id: uid(), date: lastCompleted, completedBy: assignedTo || null }] : [];

      newTasks.push({
        id: uid(),
        name: row.name.trim(),
        categoryId: cat ? cat.id : "",
        frequencyDays: row.frequencyDays && parseInt(row.frequencyDays) > 0 ? parseInt(row.frequencyDays) : null,
        lastCompleted,
        lastCompletedBy: history[0]?.completedBy || null,
        assignMode,
        assignedTo,
        completions: history.length,
        history,
      });
    });

    setCategories(cats);
    setTasks((ts) => [...ts, ...newTasks]);
    return newTasks.length;
  };

  const visibleTasks = tasks
    .filter((t) => selectedCat === "all" || t.categoryId === selectedCat)
    .filter((t) => {
      if (filterUser === "all") return true;
      if (t.assignMode === "any") return true;
      const nx = nextAssignee(t, users);
      return nx === filterUser || t.assignedTo === filterUser;
    })
    .filter((t) => matchesStatusFilter(t, statusFilter))
    .sort((a, b) => daysOverdue(b) - daysOverdue(a));

  const groupedTasks = groupBy === "category"
    ? categories.reduce((acc, cat) => { const ts = visibleTasks.filter((t) => t.categoryId === cat.id); if (ts.length) acc.push({ cat, tasks: ts }); return acc; }, [])
    : [{ cat: null, tasks: visibleTasks }];

  return (
    <div style={{ minHeight: "100vh", background: "#FFFEF7", fontFamily: "'Nunito', sans-serif", display: "flex", flexDirection: "column" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800;900&display=swap');
        * { box-sizing: border-box; margin: 0; padding: 0; }
        @keyframes confettiFly {
          0% { transform: translate(0,0) rotate(0deg) scale(1); opacity: 1; }
          100% { transform: translate(calc(cos(var(--angle)) * var(--dist)), calc(sin(var(--angle)) * var(--dist))) rotate(360deg) scale(0); opacity: 0; }
        }
        @keyframes strikethrough { 0% { width: 0; } 100% { width: 100%; } }
        @keyframes taskFadeOut {
          0% { opacity: 1; transform: scale(1); }
          40% { opacity: 0.6; transform: scale(0.99); }
          100% { opacity: 0; transform: scale(0.95) translateX(24px); }
        }
        @keyframes slideIn { from { transform: translateY(-10px); opacity: 0; } to { transform: translateY(0); opacity: 1; } }
        @keyframes bounceIn { 0% { transform: scale(0.82); opacity: 0; } 60% { transform: scale(1.04); } 100% { transform: scale(1); opacity: 1; } }
        .task-card { animation: slideIn 0.25s ease; transition: transform 0.15s, box-shadow 0.15s; }
        .task-card.completing { animation: taskFadeOut 0.85s ease forwards; pointer-events: none; }
        .task-card:hover { transform: translateY(-1px); box-shadow: 0 6px 20px rgba(0,0,0,0.1); }
        .strike-line { position: absolute; top: 50%; left: 0; height: 3px; border-radius: 2px; background: currentColor; animation: strikethrough 0.4s cubic-bezier(.4,0,.2,1) forwards; }
        input, select, textarea { font-family: 'Nunito', sans-serif; }
        ::-webkit-scrollbar { width: 6px; } ::-webkit-scrollbar-thumb { background: #ddd; border-radius: 3px; }
        .pill-btn { transition: all 0.15s; cursor: pointer; border: none; font-family: 'Nunito', sans-serif; font-weight: 700; }
        .pill-btn:hover { filter: brightness(0.92); } .pill-btn:active { transform: scale(0.96); }
        .modal-overlay { position: fixed; inset: 0; background: rgba(0,0,0,0.4); display: flex; align-items: center; justify-content: center; z-index: 1000; padding: 16px; }
        .modal { background: #fff; border-radius: 20px; padding: 28px; width: 100%; max-width: 520px; max-height: 90vh; overflow-y: auto; animation: bounceIn 0.3s ease; }
        .form-group { margin-bottom: 16px; }
        .form-label { display: block; font-weight: 700; color: #333; margin-bottom: 6px; font-size: 14px; }
        .form-input { width: 100%; padding: 10px 14px; border: 2px solid #eee; border-radius: 12px; font-size: 14px; font-weight: 600; color: #333; outline: none; transition: border-color 0.15s; background: #fff; }
        .form-input:focus { border-color: #FFD93D; }
        .suggestion-chip { display: inline-flex; align-items: center; gap: 4px; padding: 4px 10px; border-radius: 20px; background: #f5f5f5; font-size: 12px; font-weight: 700; cursor: pointer; border: none; margin: 3px; font-family: 'Nunito', sans-serif; transition: all 0.15s; }
        .suggestion-chip:hover { background: #FFD93D; color: #1A1A2E; }
        .status-slider { -webkit-appearance: none; appearance: none; height: 6px; border-radius: 4px; background: linear-gradient(90deg, #FF6B6B, #FFD93D, #4ECDC4); cursor: pointer; outline: none; }
        .status-slider::-webkit-slider-thumb { -webkit-appearance: none; appearance: none; width: 22px; height: 22px; border-radius: 50%; background: #fff; border: 3px solid #FF9F43; box-shadow: 0 2px 6px rgba(0,0,0,0.2); cursor: pointer; transition: transform 0.1s; }
        .status-slider::-webkit-slider-thumb:hover { transform: scale(1.15); }
        .status-slider::-moz-range-thumb { width: 22px; height: 22px; border-radius: 50%; background: #fff; border: 3px solid #FF9F43; box-shadow: 0 2px 6px rgba(0,0,0,0.2); cursor: pointer; }
        .history-row { display: flex; align-items: center; gap: 10px; padding: 8px 0; border-bottom: 1px solid #f5f5f5; animation: slideIn 0.2s ease; }
        .history-row:last-child { border-bottom: none; }
      `}</style>

      {/* Header */}
      <div style={{ background: "linear-gradient(135deg, #FFD93D 0%, #FF6B6B 50%, #C77DFF 100%)", padding: "16px 24px", display: "flex", alignItems: "center", justifyContent: "space-between", boxShadow: "0 4px 20px rgba(0,0,0,0.1)" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <TaskHappyMark size={40} />
          <div>
            <div style={{ fontSize: 22, fontWeight: 900, color: "#fff", textShadow: "0 2px 8px rgba(0,0,0,0.15)" }}>TaskHappy</div>
            <div style={{ fontSize: 12, color: "rgba(255,255,255,0.85)", fontWeight: 600 }}>Smarter to-do lists. Happier you.</div>
          </div>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          {["dashboard","manage","settings"].map((v) => (
            <button key={v} className="pill-btn" onClick={() => setView(v)}
              style={{ padding: "8px 16px", borderRadius: 20, background: view === v ? "#fff" : "rgba(255,255,255,0.25)", color: view === v ? "#FF6B6B" : "#fff", fontSize: 13 }}>
              {v === "dashboard" ? "📊 Dashboard" : v === "manage" ? "⚙️ Manage" : "👤 Settings"}
            </button>
          ))}
        </div>
      </div>

      {confetti.map((c) => <Confetti key={c.id} x={c.x} y={c.y} />)}

      {view === "dashboard" && (
        <DashboardView
          tasks={tasks} categories={categories} users={users}
          selectedCat={selectedCat} setSelectedCat={setSelectedCat}
          filterUser={filterUser} setFilterUser={setFilterUser}
          statusFilter={statusFilter} setStatusFilter={setStatusFilter}
          groupBy={groupBy} setGroupBy={setGroupBy}
          viewMode={viewMode} setViewMode={setViewMode}
          groupedTasks={groupedTasks} visibleTasks={visibleTasks}
          animatingIds={animatingIds}
          onCheckClick={handleCheckClick}
          onAddTask={() => setShowAddTask(true)}
          onEditTask={setEditTask}
          onViewHistory={setHistoryTask}
        />
      )}
      {view === "manage" && (
        <ManageView categories={categories} tasks={tasks} users={users}
          onAddTask={() => setShowAddTask(true)}
          onEditTask={setEditTask}
          onDeleteTask={(id) => setTasks((ts) => ts.filter((t) => t.id !== id))}
          onAddCat={() => setShowAddCat(true)}
          onDeleteCat={(id) => { setCategories((c) => c.filter((x) => x.id !== id)); setTasks((ts) => ts.filter((t) => t.categoryId !== id)); }}
          onEditCat={(cat) => setCategories((cs) => cs.map((c) => c.id === cat.id ? cat : c))}
          onViewHistory={setHistoryTask}
          onImportTasks={handleImportTasks}
        />
      )}
      {view === "settings" && (
        <SettingsView users={users} setUsers={setUsers} tasks={tasks} setTasks={setTasks} />
      )}

      {completeTask && (
        <CompleteModal
          task={completeTask.task} users={users}
          onConfirm={(date, who) => handleConfirmComplete(completeTask.task, date, who)}
          onClose={() => setCompleteTask(null)}
        />
      )}
      {historyTask && (
        <HistoryModal
          task={historyTask} users={users} categories={categories}
          onSave={(newHistory) => { updateTaskHistory(historyTask.id, newHistory); setHistoryTask((t) => ({ ...t, history: newHistory })); }}
          onClose={() => setHistoryTask(null)}
        />
      )}
      {showAddTask && (
        <TaskModal categories={categories} users={users}
          onSave={(t) => { setTasks((ts) => [...ts, { ...t, id: uid(), completions: 0, history: [] }]); setShowAddTask(false); }}
          onClose={() => setShowAddTask(false)}
        />
      )}
      {editTask && (
        <TaskModal task={editTask} categories={categories} users={users}
          onSave={(t) => { setTasks((ts) => ts.map((x) => x.id === t.id ? t : x)); setEditTask(null); }}
          onClose={() => setEditTask(null)}
        />
      )}
      {showAddCat && (
        <CategoryModal
          onSave={(c) => { setCategories((cs) => [...cs, { ...c, id: uid() }]); setShowAddCat(false); }}
          onClose={() => setShowAddCat(false)}
        />
      )}
    </div>
  );
}

// ── Status legend chip ───────────────────────────────────────────────────────
function StatusLegendItem({ color, label, value }) {
  return (
    <span style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 12, fontWeight: 700, color: "#666" }}>
      <span style={{ width: 9, height: 9, borderRadius: "50%", background: color, display: "inline-block" }} />
      {label} <span style={{ color: "#1A1A2E", fontWeight: 800 }}>{value}</span>
    </span>
  );
}

// ── Dashboard View ────────────────────────────────────────────────────────────
function DashboardView({ tasks, categories, users, selectedCat, setSelectedCat, filterUser, setFilterUser, statusFilter, setStatusFilter, groupBy, setGroupBy, viewMode, setViewMode, groupedTasks, visibleTasks, animatingIds, onCheckClick, onAddTask, onEditTask, onViewHistory }) {
  const total = tasks.length;
  const overdueTasks = tasks.filter((t) => daysOverdue(t) > 0).length;
  const dueSoonTasks = tasks.filter((t) => { const d = daysOverdue(t); return d <= 0 && d >= -3; }).length; // due today through 3 days out
  const onTrack = total - overdueTasks - dueSoonTasks;

  const overduePct = total ? (overdueTasks / total) * 100 : 0;
  const dueSoonPct = total ? (dueSoonTasks / total) * 100 : 0;
  const onTrackPct = total ? (onTrack / total) * 100 : 0;

  return (
    <div style={{ flex: 1, padding: "24px", maxWidth: viewMode === "compact" ? 1200 : 900, margin: "0 auto", width: "100%" }}>
      {/* Status overview bar */}
      <div style={{ background: "#fff", borderRadius: 16, padding: "16px 20px", marginBottom: 20, boxShadow: "0 4px 16px rgba(0,0,0,0.07)" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 10, flexWrap: "wrap", gap: 6 }}>
          <span style={{ fontSize: 13, fontWeight: 800, color: "#888", textTransform: "uppercase", letterSpacing: 0.5 }}>📊 Task Status</span>
          <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
            <StatusLegendItem color="#FF6B6B" label="Overdue" value={overdueTasks} />
            <StatusLegendItem color="#FFD93D" label="Due Soon" value={dueSoonTasks} />
            <StatusLegendItem color="#06D6A0" label="On Track" value={onTrack} />
            <span style={{ fontSize: 12, fontWeight: 700, color: "#ccc" }}>· {total} total</span>
          </div>
        </div>
        {total === 0 ? (
          <div style={{ height: 14, borderRadius: 8, background: "#f0f0f0" }} />
        ) : (
          <div style={{ display: "flex", height: 14, borderRadius: 8, overflow: "hidden", boxShadow: "inset 0 1px 3px rgba(0,0,0,0.08)" }}>
            {overduePct > 0 && (
              <div style={{ width: `${overduePct}%`, background: "linear-gradient(90deg, #FF6B6B, #FF9F43)", transition: "width 0.4s ease" }} />
            )}
            {dueSoonPct > 0 && (
              <div style={{ width: `${dueSoonPct}%`, background: "linear-gradient(90deg, #FFD93D, #B6E388)", transition: "width 0.4s ease" }} />
            )}
            {onTrackPct > 0 && (
              <div style={{ width: `${onTrackPct}%`, background: "linear-gradient(90deg, #B6E388, #06D6A0)", transition: "width 0.4s ease" }} />
            )}
          </div>
        )}
      </div>

      <div style={{ background: "#fff", borderRadius: 16, padding: "16px 20px", marginBottom: 20, boxShadow: "0 4px 16px rgba(0,0,0,0.07)" }}>
        <div style={{ display: "flex", gap: 12, flexWrap: "wrap", alignItems: "center" }}>
          <div style={{ display: "flex", gap: 6, flexWrap: "wrap", flex: 1 }}>
            <button className="pill-btn" onClick={() => setSelectedCat("all")}
              style={{ padding: "6px 14px", borderRadius: 20, background: selectedCat === "all" ? "#FFD93D" : "#f5f5f5", color: selectedCat === "all" ? "#1A1A2E" : "#555", fontSize: 13 }}>All</button>
            {categories.map((c) => (
              <button key={c.id} className="pill-btn" onClick={() => setSelectedCat(c.id)}
                style={{ padding: "6px 14px", borderRadius: 20, background: selectedCat === c.id ? c.color : "#f5f5f5", color: selectedCat === c.id ? "#fff" : "#555", fontSize: 13 }}>
                {c.emoji} {c.name}
              </button>
            ))}
          </div>
          <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
            <select className="form-input" value={filterUser} onChange={(e) => setFilterUser(e.target.value)}
              style={{ padding: "6px 12px", borderRadius: 20, fontSize: 13, width: "auto" }}>
              <option value="all">👤 All Users</option>
              {users.map((u) => <option key={u} value={u}>{u}</option>)}
            </select>
            {viewMode === "cutesy" && (
              <button className="pill-btn" onClick={() => setGroupBy(g => g === "category" ? "none" : "category")}
                style={{ padding: "6px 14px", borderRadius: 20, background: groupBy === "category" ? "#C77DFF" : "#f5f5f5", color: groupBy === "category" ? "#fff" : "#555", fontSize: 13 }}>
                {groupBy === "category" ? "📂 Grouped" : "📋 Flat"}
              </button>
            )}
            <div style={{ display: "flex", borderRadius: 20, background: "#f5f5f5", padding: 3 }}>
              <button className="pill-btn" onClick={() => setViewMode("cutesy")}
                style={{ padding: "5px 12px", borderRadius: 17, background: viewMode === "cutesy" ? "#fff" : "transparent", color: viewMode === "cutesy" ? "#1A1A2E" : "#aaa", fontSize: 13, boxShadow: viewMode === "cutesy" ? "0 2px 6px rgba(0,0,0,0.08)" : "none" }}>
                🌈 Cutesy
              </button>
              <button className="pill-btn" onClick={() => setViewMode("compact")}
                style={{ padding: "5px 12px", borderRadius: 17, background: viewMode === "compact" ? "#fff" : "transparent", color: viewMode === "compact" ? "#1A1A2E" : "#aaa", fontSize: 13, boxShadow: viewMode === "compact" ? "0 2px 6px rgba(0,0,0,0.08)" : "none" }}>
                📊 Compact
              </button>
            </div>
          </div>
        </div>

        {/* Status slider */}
        <div style={{ marginTop: 16, paddingTop: 16, borderTop: "1px solid #f5f5f5" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
            <span style={{ fontSize: 12, fontWeight: 800, color: "#888", textTransform: "uppercase", letterSpacing: 0.5 }}>📍 Showing</span>
            <span style={{ fontSize: 13, fontWeight: 800, color: STATUS_FILTER_STEPS[statusFilter].short === "Overdue" ? "#FF6B6B" : "#1A1A2E", background: STATUS_FILTER_STEPS[statusFilter].short === "Overdue" ? "#FFE5E5" : "#FFF8E1", padding: "3px 12px", borderRadius: 10 }}>
              {STATUS_FILTER_STEPS[statusFilter].label}
            </span>
          </div>
          <input
            type="range"
            min={0}
            max={6}
            step={1}
            value={statusFilter}
            onChange={(e) => setStatusFilter(parseInt(e.target.value))}
            className="status-slider"
            style={{
              width: "100%",
              accentColor: "#FF9F43",
            }}
          />
          <div style={{ display: "flex", justifyContent: "space-between", marginTop: 4 }}>
            {STATUS_FILTER_STEPS.map((s, i) => (
              <span key={i}
                onClick={() => setStatusFilter(i)}
                style={{
                  fontSize: 10, fontWeight: statusFilter === i ? 800 : 600,
                  color: statusFilter === i ? "#FF9F43" : "#ccc",
                  cursor: "pointer", textAlign: "center", flex: 1,
                  whiteSpace: "nowrap", userSelect: "none",
                }}>
                {s.short}
              </span>
            ))}
          </div>
        </div>
      </div>


      {visibleTasks.length === 0 ? (
        <div style={{ textAlign: "center", padding: "60px 20px", color: "#bbb" }}>
          <div style={{ fontSize: 48, marginBottom: 16 }}>🌟</div>
          <div style={{ fontSize: 20, fontWeight: 800, marginBottom: 8 }}>No tasks here!</div>
          <div style={{ fontWeight: 600, marginBottom: 20 }}>Add some tasks to get started.</div>
          <button className="pill-btn" onClick={onAddTask}
            style={{ padding: "12px 28px", borderRadius: 24, background: "#FFD93D", color: "#1A1A2E", fontSize: 15 }}>+ Add First Task</button>
        </div>
      ) : viewMode === "compact" ? (
        <div>
          <TaskTable tasks={visibleTasks} categories={categories} users={users}
            animatingIds={animatingIds} onCheckClick={onCheckClick} onEdit={onEditTask} onViewHistory={onViewHistory} />
          <button className="pill-btn" onClick={onAddTask}
            style={{ width: "100%", padding: "14px", borderRadius: 16, background: "linear-gradient(135deg, #FFD93D, #FF9F43)", color: "#fff", fontSize: 15, marginTop: 12 }}>
            + Add Task
          </button>
        </div>
      ) : (
        <div>
          {groupedTasks.map(({ cat, tasks: gTasks }) => (
            <div key={cat?.id || "all"} style={{ marginBottom: 28 }}>
              {cat && (
                <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
                  <span style={{ fontSize: 22 }}>{cat.emoji}</span>
                  <span style={{ fontSize: 17, fontWeight: 800, color: cat.color }}>{cat.name}</span>
                  <div style={{ flex: 1, height: 2, background: cat.color, borderRadius: 2, opacity: 0.3 }} />
                  <span style={{ fontSize: 12, fontWeight: 700, color: "#aaa" }}>{gTasks.length} tasks</span>
                </div>
              )}
              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                {gTasks.map((task) => (
                  <TaskCard key={task.id} task={task} categories={categories} users={users}
                    completing={animatingIds.has(task.id)}
                    onCheckClick={(e) => onCheckClick(task, e)}
                    onEdit={() => onEditTask(task)}
                    onViewHistory={() => onViewHistory(task)}
                  />
                ))}
              </div>
            </div>
          ))}
          <button className="pill-btn" onClick={onAddTask}
            style={{ width: "100%", padding: "14px", borderRadius: 16, background: "linear-gradient(135deg, #FFD93D, #FF9F43)", color: "#fff", fontSize: 15, marginTop: 8 }}>
            + Add Task
          </button>
        </div>
      )}
    </div>
  );
}

// ── Task Card ─────────────────────────────────────────────────────────────────
function TaskCard({ task, categories, users, completing, onCheckClick, onEdit, onViewHistory }) {
  const cat = categories.find((c) => c.id === task.categoryId);
  const { label, color } = overdueLabel(task);
  const nx = nextAssignee(task, users);
  const avg = avgCadence(task.history);
  const eff = effectiveFreq(task);
  const due = nextDueDate(task);
  const last = lastCompletedDate(task);
  const historyCount = (task.history || []).length;

  return (
    <div className={`task-card ${completing ? "completing" : ""}`}
      style={{ background: "#fff", borderRadius: 16, padding: "16px 20px", boxShadow: "0 2px 12px rgba(0,0,0,0.06)", display: "flex", alignItems: "center", gap: 16, borderLeft: `5px solid ${cat?.color || "#ddd"}` }}>

      <button onClick={onCheckClick} className="pill-btn"
        style={{ width: 38, height: 38, borderRadius: "50%", flexShrink: 0, background: completing ? "#06D6A0" : "#f5f5f5", color: completing ? "#fff" : "#ccc", fontSize: 20, display: "flex", alignItems: "center", justifyContent: "center" }}>
        {completing ? "✓" : "○"}
      </button>

      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{ position: "relative", display: "inline-block" }}>
          <span style={{ fontSize: 15, fontWeight: 800, color: completing ? "#bbb" : "#1A1A2E", transition: "color 0.3s" }}>
            {cat?.emoji} {task.name}
          </span>
          {completing && <span className="strike-line" style={{ color: cat?.color || "#FF6B6B" }} />}
        </div>
        <div style={{ display: "flex", gap: 8, marginTop: 5, flexWrap: "wrap", alignItems: "center" }}>
          {cat && <span style={{ fontSize: 11, fontWeight: 700, color: cat.color, background: cat.color + "18", padding: "2px 8px", borderRadius: 10 }}>{cat.name}</span>}
          {eff && <span style={{ fontSize: 11, fontWeight: 700, color: "#aaa" }}>🎯 {freqLabel(task.frequencyDays)}</span>}
          {avg && avg !== task.frequencyDays && (
            <span style={{ fontSize: 11, fontWeight: 700, color: "#FF9F43", background: "#FF9F4318", padding: "2px 8px", borderRadius: 10 }}>
              📊 avg {freqLabel(avg)}
            </span>
          )}
          {last && <span style={{ fontSize: 11, color: "#bbb", fontWeight: 600 }}>Last: {last}</span>}
          {due && <span style={{ fontSize: 11, color: "#4ECDC4", fontWeight: 700 }}>Due: {due}</span>}
        </div>
      </div>

      <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-end", gap: 5, flexShrink: 0 }}>
        <span style={{ fontSize: 12, fontWeight: 800, color, background: color + "18", padding: "3px 10px", borderRadius: 10, whiteSpace: "nowrap" }}>{label}</span>
        {nx && <span style={{ fontSize: 11, fontWeight: 700, color: "#C77DFF", background: "#C77DFF18", padding: "2px 8px", borderRadius: 10 }}>👤 {nx}</span>}
        {task.assignMode === "any" && <span style={{ fontSize: 11, fontWeight: 700, color: "#4ECDC4", background: "#4ECDC418", padding: "2px 8px", borderRadius: 10 }}>Anyone</span>}
        <div style={{ display: "flex", gap: 4 }}>
          <button className="pill-btn" onClick={onViewHistory}
            style={{ fontSize: 11, padding: "3px 8px", borderRadius: 10, background: "#f0f0ff", color: "#C77DFF" }}>
            📅 {historyCount}
          </button>
          <button className="pill-btn" onClick={onEdit}
            style={{ fontSize: 11, padding: "3px 8px", borderRadius: 10, background: "#f5f5f5", color: "#888" }}>✏️</button>
        </div>
      </div>
    </div>
  );
}

// ── Task Table (compact view) ───────────────────────────────────────────────────
const TABLE_COLS = [
  { key: "status", label: "", width: 44 },
  { key: "history", label: "", width: 40 },
  { key: "edit", label: "", width: 40 },
  { key: "name", label: "Task", width: undefined },
  { key: "due", label: "Due Date", width: 110 },
  { key: "status_label", label: "Status", width: 120 },
  { key: "category", label: "Category", width: 140 },
  { key: "goal", label: "Goal", width: 110 },
  { key: "avg", label: "Avg", width: 100 },
  { key: "last", label: "Last Done", width: 110 },
  { key: "assignee", label: "Assigned", width: 110 },
];

function TaskTable({ tasks, categories, users, animatingIds, onCheckClick, onEdit, onViewHistory }) {
  const [sortKey, setSortKey] = useState("due");
  const [sortDir, setSortDir] = useState("asc"); // asc = most overdue/soonest first

  const sorted = [...tasks].sort((a, b) => {
    let cmp = 0;
    switch (sortKey) {
      case "name": cmp = a.name.localeCompare(b.name); break;
      case "category": {
        const ca = categories.find((c) => c.id === a.categoryId)?.name || "";
        const cb = categories.find((c) => c.id === b.categoryId)?.name || "";
        cmp = ca.localeCompare(cb); break;
      }
      case "goal": cmp = (a.frequencyDays || 99999) - (b.frequencyDays || 99999); break;
      case "avg": cmp = (avgCadence(a.history) || 99999) - (avgCadence(b.history) || 99999); break;
      case "last": {
        const la = lastCompletedDate(a), lb = lastCompletedDate(b);
        cmp = (la ? new Date(la).getTime() : -Infinity) - (lb ? new Date(lb).getTime() : -Infinity); break;
      }
      case "assignee": {
        const aa = nextAssignee(a, users) || a.assignedTo || ""; const ab = nextAssignee(b, users) || b.assignedTo || "";
        cmp = String(aa).localeCompare(String(ab)); break;
      }
      case "due":
      default:
        // most overdue first => higher daysOverdue first
        cmp = daysOverdue(b) - daysOverdue(a);
        break;
    }
    return sortDir === "asc" ? cmp : -cmp;
  });

  const toggleSort = (key) => {
    if (key === "status" || key === "history" || key === "edit") return;
    if (sortKey === key) {
      setSortDir((d) => (d === "asc" ? "desc" : "asc"));
    } else {
      setSortKey(key);
      // sensible default direction per column
      setSortDir(key === "due" ? "asc" : "asc");
    }
  };

  return (
    <div style={{ background: "#fff", borderRadius: 16, boxShadow: "0 2px 12px rgba(0,0,0,0.06)", overflow: "hidden" }}>
      <div style={{ overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#FFFBE8", borderBottom: "2px solid #f0f0f0" }}>
              {TABLE_COLS.map((col) => (
                <th key={col.key}
                  onClick={() => toggleSort(col.key)}
                  style={{
                    textAlign: col.key === "name" ? "left" : "left",
                    padding: "10px 12px",
                    fontWeight: 800, color: "#888", fontSize: 11, textTransform: "uppercase", letterSpacing: 0.5,
                    cursor: (col.key === "status" || col.key === "history" || col.key === "edit") ? "default" : "pointer",
                    whiteSpace: "nowrap", userSelect: "none",
                    width: col.width,
                  }}>
                  {col.label}
                  {sortKey === col.key && (sortDir === "asc" ? " ▲" : " ▼")}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {sorted.map((task) => (
              <TaskTableRow key={task.id} task={task} categories={categories} users={users}
                completing={animatingIds.has(task.id)}
                onCheckClick={(e) => onCheckClick(task, e)}
                onEdit={() => onEdit(task)}
                onViewHistory={() => onViewHistory(task)}
              />
            ))}
          </tbody>
        </table>
      </div>
      {sorted.length === 0 && (
        <div style={{ textAlign: "center", padding: 40, color: "#bbb", fontWeight: 700 }}>No tasks match this filter.</div>
      )}
    </div>
  );
}

function TaskTableRow({ task, categories, users, completing, onCheckClick, onEdit, onViewHistory }) {
  const cat = categories.find((c) => c.id === task.categoryId);
  const { label, color } = overdueLabel(task);
  const nx = nextAssignee(task, users);
  const avg = avgCadence(task.history);
  const due = nextDueDate(task);
  const last = lastCompletedDate(task);

  return (
    <tr style={{
      borderBottom: "1px solid #f5f5f5",
      opacity: completing ? 0.4 : 1,
      transition: "opacity 0.5s, background 0.15s",
      background: "#fff",
    }}
    onMouseEnter={(e) => e.currentTarget.style.background = "#FFFEF7"}
    onMouseLeave={(e) => e.currentTarget.style.background = "#fff"}>
      <td style={{ padding: "8px 12px" }}>
        <button onClick={onCheckClick} className="pill-btn"
          style={{ width: 28, height: 28, borderRadius: "50%", background: completing ? "#06D6A0" : "#f5f5f5", color: completing ? "#fff" : "#ccc", fontSize: 14, display: "flex", alignItems: "center", justifyContent: "center" }}>
          {completing ? "✓" : "○"}
        </button>
      </td>
      <td style={{ padding: "8px 6px" }}>
        <button className="pill-btn" onClick={onViewHistory}
          style={{ fontSize: 11, padding: "3px 8px", borderRadius: 8, background: "#f0f0ff", color: "#C77DFF" }}>📅</button>
      </td>
      <td style={{ padding: "8px 6px" }}>
        <button className="pill-btn" onClick={onEdit}
          style={{ fontSize: 11, padding: "3px 8px", borderRadius: 8, background: "#f5f5f5", color: "#888" }}>✏️</button>
      </td>
      <td style={{ padding: "8px 12px", fontWeight: 800, color: "#1A1A2E", whiteSpace: "nowrap", position: "relative" }}>
        <span style={{ textDecoration: completing ? "line-through" : "none", transition: "text-decoration 0.3s" }}>
          {cat?.emoji} {task.name}
        </span>
      </td>
      <td style={{ padding: "8px 12px", color: "#4ECDC4", fontWeight: 700, whiteSpace: "nowrap" }}>{due || "—"}</td>
      <td style={{ padding: "8px 12px", whiteSpace: "nowrap" }}>
        <span style={{ fontSize: 12, fontWeight: 800, color, background: color + "18", padding: "3px 10px", borderRadius: 10 }}>{label}</span>
      </td>
      <td style={{ padding: "8px 12px", whiteSpace: "nowrap" }}>
        {cat && <span style={{ fontSize: 11, fontWeight: 700, color: cat.color, background: cat.color + "18", padding: "2px 8px", borderRadius: 10 }}>{cat.name}</span>}
      </td>
      <td style={{ padding: "8px 12px", color: "#666", fontWeight: 700, whiteSpace: "nowrap" }}>{freqLabel(task.frequencyDays)}</td>
      <td style={{ padding: "8px 12px", whiteSpace: "nowrap" }}>
        {avg ? <span style={{ fontSize: 12, fontWeight: 700, color: "#FF9F43" }}>{freqLabel(avg)}</span> : <span style={{ color: "#ddd" }}>—</span>}
      </td>
      <td style={{ padding: "8px 12px", color: "#aaa", fontWeight: 600, whiteSpace: "nowrap" }}>{last || "—"}</td>
      <td style={{ padding: "8px 12px", whiteSpace: "nowrap" }}>
        {nx ? (
          <span style={{ fontSize: 11, fontWeight: 700, color: "#C77DFF", background: "#C77DFF18", padding: "2px 8px", borderRadius: 10 }}>👤 {nx}</span>
        ) : task.assignMode === "any" ? (
          <span style={{ fontSize: 11, fontWeight: 700, color: "#4ECDC4", background: "#4ECDC418", padding: "2px 8px", borderRadius: 10 }}>Anyone</span>
        ) : "—"}
      </td>
    </tr>
  );
}


function CompleteModal({ task, users, onConfirm, onClose }) {
  const [date, setDate] = useState(today());
  const [who, setWho] = useState(users[0] || "");

  return (
    <div className="modal-overlay" onClick={(e) => e.target === e.currentTarget && onClose()}>
      <div className="modal" style={{ maxWidth: 380 }}>
        <div style={{ textAlign: "center", marginBottom: 20 }}>
          <div style={{ fontSize: 40, marginBottom: 8 }}>🎉</div>
          <h3 style={{ fontSize: 19, fontWeight: 900, color: "#1A1A2E", marginBottom: 4 }}>Task Complete!</h3>
          <p style={{ fontSize: 14, color: "#888", fontWeight: 600 }}>{task.name}</p>
        </div>

        <div className="form-group">
          <label className="form-label">When did you complete it?</label>
          <input className="form-input" type="date" value={date} max={today()}
            onChange={(e) => setDate(e.target.value)} />
        </div>

        {users.length > 1 && (
          <div className="form-group">
            <label className="form-label">Who completed it?</label>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              {users.map((u) => (
                <button key={u} className="pill-btn" onClick={() => setWho(u)}
                  style={{ flex: 1, padding: "9px 12px", borderRadius: 12, background: who === u ? "#C77DFF" : "#f5f5f5", color: who === u ? "#fff" : "#555", fontSize: 14 }}>
                  👤 {u}
                </button>
              ))}
            </div>
          </div>
        )}

        <div style={{ display: "flex", gap: 10, marginTop: 8 }}>
          <button className="pill-btn" onClick={onClose}
            style={{ flex: 1, padding: "12px", borderRadius: 14, background: "#f5f5f5", color: "#666" }}>Cancel</button>
          <button className="pill-btn" onClick={() => onConfirm(date, who)}
            style={{ flex: 2, padding: "12px", borderRadius: 14, background: "linear-gradient(135deg, #06D6A0, #4ECDC4)", color: "#fff", fontSize: 15 }}>
            ✅ Log Completion
          </button>
        </div>
      </div>
    </div>
  );
}

// ── History Modal ─────────────────────────────────────────────────────────────
function HistoryModal({ task, users, categories, onSave, onClose }) {
  const cat = categories.find((c) => c.id === task.categoryId);
  const [history, setHistory] = useState(() =>
    [...(task.history || [])].sort((a, b) => new Date(b.date) - new Date(a.date))
  );
  const [newDate, setNewDate] = useState(today());
  const [newWho, setNewWho] = useState(users[0] || "");
  const [editingId, setEditingId] = useState(null);
  const [editDate, setEditDate] = useState("");
  const [editWho, setEditWho] = useState("");

  const avg = avgCadence(history);
  const goalFreq = task.frequencyDays;
  const eff = effectiveFreq({ ...task, history });

  const addEntry = () => {
    if (!newDate) return;
    const entry = { id: uid(), date: newDate, completedBy: newWho };
    const updated = [...history, entry].sort((a, b) => new Date(b.date) - new Date(a.date));
    setHistory(updated);
    onSave(updated);
  };

  const deleteEntry = (id) => {
    const updated = history.filter((h) => h.id !== id);
    setHistory(updated);
    onSave(updated);
  };

  const startEdit = (entry) => {
    setEditingId(entry.id);
    setEditDate(entry.date);
    setEditWho(entry.completedBy || "");
  };

  const saveEdit = () => {
    const updated = history.map((h) => h.id === editingId ? { ...h, date: editDate, completedBy: editWho } : h)
      .sort((a, b) => new Date(b.date) - new Date(a.date));
    setHistory(updated);
    onSave(updated);
    setEditingId(null);
  };

  return (
    <div className="modal-overlay" onClick={(e) => e.target === e.currentTarget && onClose()}>
      <div className="modal" style={{ maxWidth: 500 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 20 }}>
          <div>
            <h3 style={{ fontSize: 18, fontWeight: 900, color: "#1A1A2E" }}>
              {cat?.emoji} {task.name}
            </h3>
            <p style={{ fontSize: 12, color: "#aaa", fontWeight: 600, marginTop: 2 }}>Completion History</p>
          </div>
          <button className="pill-btn" onClick={onClose} style={{ padding: "6px 12px", borderRadius: 10, background: "#f5f5f5", color: "#888" }}>✕</button>
        </div>

        {/* Stats */}
        <div style={{ display: "flex", gap: 10, marginBottom: 20, flexWrap: "wrap" }}>
          <div style={{ flex: 1, background: "#f8f4ff", borderRadius: 12, padding: "10px 14px" }}>
            <div style={{ fontSize: 11, fontWeight: 700, color: "#C77DFF", marginBottom: 2 }}>GOAL</div>
            <div style={{ fontSize: 15, fontWeight: 800, color: "#1A1A2E" }}>{freqLabel(goalFreq)}</div>
          </div>
          {avg && (
            <div style={{ flex: 1, background: "#fff8ed", borderRadius: 12, padding: "10px 14px" }}>
              <div style={{ fontSize: 11, fontWeight: 700, color: "#FF9F43", marginBottom: 2 }}>YOUR AVERAGE</div>
              <div style={{ fontSize: 15, fontWeight: 800, color: "#1A1A2E" }}>{freqLabel(avg)}</div>
            </div>
          )}
          {eff && (
            <div style={{ flex: 1, background: "#edfff8", borderRadius: 12, padding: "10px 14px" }}>
              <div style={{ fontSize: 11, fontWeight: 700, color: "#06D6A0", marginBottom: 2 }}>EFFECTIVE</div>
              <div style={{ fontSize: 15, fontWeight: 800, color: "#1A1A2E" }}>{freqLabel(eff)}</div>
            </div>
          )}
          <div style={{ flex: 1, background: "#f0f9ff", borderRadius: 12, padding: "10px 14px" }}>
            <div style={{ fontSize: 11, fontWeight: 700, color: "#4ECDC4", marginBottom: 2 }}>COMPLETIONS</div>
            <div style={{ fontSize: 15, fontWeight: 800, color: "#1A1A2E" }}>{history.length}×</div>
          </div>
        </div>

        {avg && goalFreq && avg > goalFreq * 1.2 && (
          <div style={{ background: "#FFF5E5", border: "1px solid #FFD93D", borderRadius: 12, padding: "10px 14px", marginBottom: 16, fontSize: 13, fontWeight: 600, color: "#FF9F43" }}>
            💡 Your average ({freqLabel(avg)}) is longer than your goal ({freqLabel(goalFreq)}). The effective schedule adjusts to keep things realistic while nudging you toward your goal.
          </div>
        )}
        {avg && goalFreq && avg < goalFreq * 0.8 && (
          <div style={{ background: "#EDFFF8", border: "1px solid #06D6A0", borderRadius: 12, padding: "10px 14px", marginBottom: 16, fontSize: 13, fontWeight: 600, color: "#06D6A0" }}>
            🌟 You're doing this more often than your goal — keep it up!
          </div>
        )}

        {/* Add new entry */}
        <div style={{ background: "#f9f9f9", borderRadius: 14, padding: 14, marginBottom: 16 }}>
          <div style={{ fontSize: 13, fontWeight: 800, color: "#555", marginBottom: 10 }}>+ Add Completion</div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "flex-end" }}>
            <div style={{ flex: 2 }}>
              <label className="form-label">Date</label>
              <input className="form-input" type="date" value={newDate} max={today()} onChange={(e) => setNewDate(e.target.value)} />
            </div>
            {users.length > 1 && (
              <div style={{ flex: 2 }}>
                <label className="form-label">Who</label>
                <select className="form-input" value={newWho} onChange={(e) => setNewWho(e.target.value)}>
                  {users.map((u) => <option key={u} value={u}>{u}</option>)}
                </select>
              </div>
            )}
            <button className="pill-btn" onClick={addEntry}
              style={{ padding: "10px 16px", borderRadius: 12, background: "#FFD93D", color: "#1A1A2E", fontSize: 13, height: 42 }}>
              Add
            </button>
          </div>
        </div>

        {/* History list */}
        <div style={{ maxHeight: 280, overflowY: "auto" }}>
          {history.length === 0 ? (
            <div style={{ textAlign: "center", padding: "20px", color: "#bbb", fontWeight: 700 }}>No history yet</div>
          ) : (
            history.map((entry, i) => (
              <div key={entry.id} className="history-row">
                <div style={{ width: 28, height: 28, borderRadius: "50%", background: "#f0f0ff", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 13, fontWeight: 800, color: "#C77DFF", flexShrink: 0 }}>
                  {history.length - i}
                </div>
                {editingId === entry.id ? (
                  <>
                    <input className="form-input" type="date" value={editDate} max={today()} onChange={(e) => setEditDate(e.target.value)} style={{ flex: 2 }} />
                    {users.length > 1 && (
                      <select className="form-input" value={editWho} onChange={(e) => setEditWho(e.target.value)} style={{ flex: 2 }}>
                        {users.map((u) => <option key={u} value={u}>{u}</option>)}
                      </select>
                    )}
                    <button className="pill-btn" onClick={saveEdit} style={{ padding: "5px 12px", borderRadius: 10, background: "#06D6A0", color: "#fff", fontSize: 12 }}>Save</button>
                    <button className="pill-btn" onClick={() => setEditingId(null)} style={{ padding: "5px 10px", borderRadius: 10, background: "#f5f5f5", color: "#888", fontSize: 12 }}>✕</button>
                  </>
                ) : (
                  <>
                    <div style={{ flex: 1 }}>
                      <div style={{ fontWeight: 800, color: "#1A1A2E", fontSize: 14 }}>{entry.date}</div>
                      {entry.completedBy && <div style={{ fontSize: 12, color: "#aaa", fontWeight: 600 }}>by {entry.completedBy}</div>}
                    </div>
                    {i > 0 && (() => {
                      const prev = history[i - 1];
                      const gap = daysBetween(entry.date, prev.date);
                      return gap > 0 ? <span style={{ fontSize: 11, color: "#aaa", fontWeight: 700, whiteSpace: "nowrap" }}>{gap}d gap</span> : null;
                    })()}
                    <button className="pill-btn" onClick={() => startEdit(entry)} style={{ padding: "4px 10px", borderRadius: 10, background: "#f5f5f5", color: "#888", fontSize: 12 }}>✏️</button>
                    <button className="pill-btn" onClick={() => deleteEntry(entry.id)} style={{ padding: "4px 8px", borderRadius: 10, background: "#FFE5E5", color: "#FF6B6B", fontSize: 12 }}>✕</button>
                  </>
                )}
              </div>
            ))
          )}
        </div>

        <button className="pill-btn" onClick={onClose}
          style={{ width: "100%", padding: "12px", borderRadius: 14, background: "#f5f5f5", color: "#666", marginTop: 16, fontSize: 15 }}>
          Done
        </button>
      </div>
    </div>
  );
}

// ── Manage View ───────────────────────────────────────────────────────────────
function ManageView({ categories, tasks, users, onAddTask, onEditTask, onDeleteTask, onAddCat, onDeleteCat, onEditCat, onViewHistory, onImportTasks }) {
  const [editingCat, setEditingCat] = useState(null);
  const [importResult, setImportResult] = useState(null); // { count } | { error }
  const [importing, setImporting] = useState(false);
  const fileInputRef = useRef(null);

  const handleFileChange = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const name = file.name.toLowerCase();
    const isExcel = name.endsWith(".xlsx") || name.endsWith(".xls");

    setImporting(true);
    setImportResult(null);
    const reader = new FileReader();

    reader.onload = (ev) => {
      try {
        let rawRows;
        if (isExcel) {
          rawRows = parseXLSXRows(ev.target.result);
        } else {
          rawRows = parseCSVRows(String(ev.target.result || ""));
        }

        const validationError = validateImportRows(rawRows);
        if (validationError) {
          setImportResult({ error: validationError });
          return;
        }

        const rows = rowsToTaskObjects(rawRows);
        if (rows.length === 0) {
          setImportResult({ error: "No new tasks found. Make sure your task rows are below the header, each has a name in the first column, and the sample row (\"Vacuum living room\") has been replaced or removed." });
        } else {
          const count = onImportTasks(rows);
          setImportResult({ count });
        }
      } catch (err) {
        const detail = err && err.message ? err.message : String(err);
        if (isExcel) {
          setImportResult({ error: `Couldn't read this Excel file (${detail}). Make sure it's a .xlsx file saved from Excel, Google Sheets, or Numbers — not a renamed CSV — and that it isn't password-protected.` });
        } else {
          setImportResult({ error: `Couldn't read this CSV file (${detail}). Make sure it's a plain text .csv file with comma-separated columns matching the template.` });
        }
      } finally {
        setImporting(false);
      }
    };
    reader.onerror = () => { setImportResult({ error: "The file couldn't be opened — it may be too large, corrupted, or in a format your browser can't read." }); setImporting(false); };

    if (isExcel) reader.readAsArrayBuffer(file);
    else reader.readAsText(file);

    e.target.value = ""; // allow re-selecting the same file
  };

  return (
    <div style={{ flex: 1, padding: "24px", maxWidth: 900, margin: "0 auto", width: "100%" }}>
      {/* Bulk Import */}
      <div style={{ marginBottom: 32, background: "#fff", borderRadius: 16, padding: "18px 20px", boxShadow: "0 2px 12px rgba(0,0,0,0.07)", borderTop: "4px solid #06D6A0" }}>
        <h2 style={{ fontSize: 18, fontWeight: 900, color: "#1A1A2E", marginBottom: 6 }}>📥 Bulk Add Tasks</h2>
        <p style={{ fontSize: 13, color: "#999", fontWeight: 600, marginBottom: 14 }}>
          Download a template, fill in your recurring tasks, and upload it to add them all at once. CSV and Excel (.xlsx) files are both supported.
        </p>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
          <button className="pill-btn" onClick={downloadCSVTemplate}
            style={{ padding: "10px 20px", borderRadius: 14, background: "#4ECDC4", color: "#fff", fontSize: 14 }}>
            ⬇️ CSV Template
          </button>
          <button className="pill-btn" onClick={downloadXLSXTemplate}
            style={{ padding: "10px 20px", borderRadius: 14, background: "#74B9FF", color: "#fff", fontSize: 14 }}>
            ⬇️ Excel Template
          </button>
          <button className="pill-btn" onClick={() => fileInputRef.current?.click()} disabled={importing}
            style={{ padding: "10px 20px", borderRadius: 14, background: "#06D6A0", color: "#fff", fontSize: 14, opacity: importing ? 0.6 : 1 }}>
            {importing ? "Importing..." : "⬆️ Upload Filled File"}
          </button>
          <input ref={fileInputRef} type="file" accept=".csv,text/csv,.xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel" onChange={handleFileChange} style={{ display: "none" }} />
        </div>
        {importResult && (
          <div style={{ marginTop: 14, padding: "10px 14px", borderRadius: 12, fontSize: 13, fontWeight: 700, lineHeight: 1.5,
            background: importResult.error ? "#FFE5E5" : "#EDFFF8", color: importResult.error ? "#FF6B6B" : "#06D6A0" }}>
            {importResult.error ? `⚠️ ${importResult.error}` : `🎉 Added ${importResult.count} task${importResult.count === 1 ? "" : "s"} from your file!`}
          </div>
        )}
      </div>

      <div style={{ marginBottom: 32 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
          <h2 style={{ fontSize: 20, fontWeight: 900, color: "#1A1A2E" }}>📂 Categories</h2>
          <button className="pill-btn" onClick={onAddCat}
            style={{ padding: "8px 18px", borderRadius: 20, background: "#C77DFF", color: "#fff", fontSize: 13 }}>+ Add Category</button>
        </div>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 12 }}>
          {categories.map((cat) => (
            <div key={cat.id} style={{ background: "#fff", borderRadius: 16, padding: "14px 18px", boxShadow: "0 2px 12px rgba(0,0,0,0.07)", borderTop: `4px solid ${cat.color}`, minWidth: 160 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 8 }}>
                <div style={{ fontSize: 24 }}>{cat.emoji}</div>
                <div style={{ display: "flex", gap: 4 }}>
                  <button className="pill-btn" onClick={() => setEditingCat(cat)} style={{ fontSize: 11, padding: "2px 8px", borderRadius: 8, background: "#f5f5f5", color: "#888" }}>✏️</button>
                  <button className="pill-btn" onClick={() => onDeleteCat(cat.id)} style={{ fontSize: 11, padding: "2px 8px", borderRadius: 8, background: "#FFE5E5", color: "#FF6B6B" }}>✕</button>
                </div>
              </div>
              <div style={{ fontWeight: 800, color: cat.color, marginBottom: 4 }}>{cat.name}</div>
              <div style={{ fontSize: 12, color: "#aaa", fontWeight: 600 }}>{tasks.filter((t) => t.categoryId === cat.id).length} tasks</div>
            </div>
          ))}
        </div>
      </div>

      <div>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
          <h2 style={{ fontSize: 20, fontWeight: 900, color: "#1A1A2E" }}>📝 All Tasks</h2>
          <button className="pill-btn" onClick={onAddTask}
            style={{ padding: "8px 18px", borderRadius: 20, background: "#FFD93D", color: "#1A1A2E", fontSize: 13 }}>+ Add Task</button>
        </div>
        {tasks.length === 0 ? (
          <div style={{ textAlign: "center", padding: 40, color: "#bbb", fontWeight: 700 }}>No tasks yet. Add one!</div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {tasks.map((task) => {
              const cat = categories.find((c) => c.id === task.categoryId);
              const avg = avgCadence(task.history);
              return (
                <div key={task.id} style={{ background: "#fff", borderRadius: 12, padding: "12px 16px", boxShadow: "0 2px 8px rgba(0,0,0,0.06)", display: "flex", alignItems: "center", gap: 12 }}>
                  <span style={{ fontSize: 18 }}>{cat?.emoji}</span>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontWeight: 800, color: "#1A1A2E", fontSize: 14 }}>{task.name}</div>
                    <div style={{ fontSize: 12, color: "#aaa", fontWeight: 600 }}>
                      {cat?.name} · Goal: {freqLabel(task.frequencyDays)}
                      {avg ? ` · Avg: ${freqLabel(avg)}` : ""}
                    </div>
                  </div>
                  <div style={{ fontSize: 12, color: "#C77DFF", fontWeight: 700 }}>
                    {task.assignMode === "any" ? "Anyone" : task.assignMode === "alternate" ? "Alternating" : task.assignedTo}
                  </div>
                  <button className="pill-btn" onClick={() => onViewHistory(task)}
                    style={{ fontSize: 12, padding: "5px 10px", borderRadius: 10, background: "#f0f0ff", color: "#C77DFF" }}>📅 {(task.history || []).length}</button>
                  <button className="pill-btn" onClick={() => onEditTask(task)}
                    style={{ fontSize: 12, padding: "5px 12px", borderRadius: 10, background: "#f5f5f5", color: "#555" }}>✏️</button>
                  <button className="pill-btn" onClick={() => onDeleteTask(task.id)}
                    style={{ fontSize: 12, padding: "5px 10px", borderRadius: 10, background: "#FFE5E5", color: "#FF6B6B" }}>🗑️</button>
                </div>
              );
            })}
          </div>
        )}
      </div>

      {editingCat && (
        <CategoryModal cat={editingCat}
          onSave={(c) => { onEditCat(c); setEditingCat(null); }}
          onClose={() => setEditingCat(null)}
        />
      )}
    </div>
  );
}

// ── Settings View ─────────────────────────────────────────────────────────────
function SettingsView({ users, setUsers, tasks, setTasks }) {
  const [newUser, setNewUser] = useState("");
  return (
    <div style={{ flex: 1, padding: "24px", maxWidth: 600, margin: "0 auto", width: "100%" }}>
      <h2 style={{ fontSize: 20, fontWeight: 900, color: "#1A1A2E", marginBottom: 20 }}>👤 Household Members</h2>
      <div style={{ background: "#fff", borderRadius: 16, padding: 20, boxShadow: "0 2px 12px rgba(0,0,0,0.07)", marginBottom: 20 }}>
        {users.map((u) => (
          <div key={u} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "10px 0", borderBottom: "1px solid #f5f5f5" }}>
            <span style={{ fontWeight: 700, color: "#1A1A2E" }}>👤 {u}</span>
            {users.length > 1 && (
              <button className="pill-btn" onClick={() => setUsers(users.filter((x) => x !== u))}
                style={{ padding: "4px 12px", borderRadius: 10, background: "#FFE5E5", color: "#FF6B6B", fontSize: 12 }}>Remove</button>
            )}
          </div>
        ))}
        <div style={{ display: "flex", gap: 8, marginTop: 14 }}>
          <input className="form-input" value={newUser} onChange={(e) => setNewUser(e.target.value)}
            placeholder="Add household member..." style={{ flex: 1 }}
            onKeyDown={(e) => { if (e.key === "Enter" && newUser.trim()) { setUsers([...users, newUser.trim()]); setNewUser(""); } }} />
          <button className="pill-btn" onClick={() => { if (newUser.trim()) { setUsers([...users, newUser.trim()]); setNewUser(""); } }}
            style={{ padding: "10px 20px", borderRadius: 12, background: "#FFD93D", color: "#1A1A2E", fontSize: 14 }}>Add</button>
        </div>
      </div>
      <div style={{ background: "#FFF5E5", borderRadius: 16, padding: 20, border: "2px solid #FFD93D" }}>
        <div style={{ fontWeight: 800, color: "#FF9F43", marginBottom: 8 }}>⚠️ Danger Zone</div>
        <button className="pill-btn"
          onClick={() => { if (confirm("Reset all task history?")) setTasks((ts) => ts.map((t) => ({ ...t, lastCompleted: null, lastCompletedBy: null, completions: 0, history: [] }))); }}
          style={{ padding: "8px 18px", borderRadius: 12, background: "#FF6B6B", color: "#fff", fontSize: 13 }}>
          Reset All Task History
        </button>
      </div>
    </div>
  );
}

// ── Task Modal ────────────────────────────────────────────────────────────────
function TaskModal({ task, categories, users, onSave, onClose }) {
  const [name, setName] = useState(task?.name || "");
  const [categoryId, setCategoryId] = useState(task?.categoryId || categories[0]?.id || "");
  const [frequencyDays, setFrequencyDays] = useState(task?.frequencyDays ?? null);
  const [customInput, setCustomInput] = useState(task?.frequencyDays ? String(task.frequencyDays) : "");
  const [lastCompleted, setLastCompleted] = useState(task?.lastCompleted || "");
  const [assignMode, setAssignMode] = useState(task?.assignMode || "any");
  const [assignedTo, setAssignedTo] = useState(task?.assignedTo || users[0] || "");
  const [suggestions, setSuggestions] = useState([]);

  const cat = categories.find((c) => c.id === categoryId);

  useEffect(() => {
    const catSuggestions = cat && TASK_SUGGESTIONS[cat.name] ? TASK_SUGGESTIONS[cat.name] : [];
    if (name.length < 2 && catSuggestions.length) {
      setSuggestions(catSuggestions);
    } else if (name.length > 1) {
      const all = Object.values(TASK_SUGGESTIONS).flat();
      setSuggestions(all.filter((s) => s.name.toLowerCase().includes(name.toLowerCase())).slice(0, 6));
    } else {
      setSuggestions([]);
    }
  }, [categoryId, name]);

  const setFreq = (days) => {
    setFrequencyDays(days);
    setCustomInput(days ? String(days) : "");
  };

  const handleCustomInput = (val) => {
    setCustomInput(val);
    const n = parseInt(val);
    setFrequencyDays(n > 0 ? n : null);
  };

  const handleSave = () => {
    if (!name.trim()) return;
    let history = task?.history || [];
    // If this is a new task and a "last completed" date was entered, log it as the first history entry
    if (!task && lastCompleted) {
      history = [...history, { id: uid(), date: lastCompleted, completedBy: assignedTo || users[0] || null }];
    }
    onSave({ id: task?.id, name: name.trim(), categoryId, frequencyDays, lastCompleted, assignMode, assignedTo, lastCompletedBy: task?.lastCompletedBy, completions: history.length, history });
  };

  return (
    <div className="modal-overlay" onClick={(e) => e.target === e.currentTarget && onClose()}>
      <div className="modal">
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
          <h3 style={{ fontSize: 20, fontWeight: 900, color: "#1A1A2E" }}>{task ? "✏️ Edit Task" : "✨ Add Task"}</h3>
          <button className="pill-btn" onClick={onClose} style={{ padding: "6px 12px", borderRadius: 10, background: "#f5f5f5", color: "#888" }}>✕</button>
        </div>

        <div className="form-group">
          <label className="form-label">Task Name</label>
          <input className="form-input" value={name} onChange={(e) => setName(e.target.value)} placeholder="What needs to be done?" />
          {suggestions.length > 0 && name.length < 2 && (
            <div style={{ marginTop: 8 }}>
              <div style={{ fontSize: 12, color: "#aaa", fontWeight: 700, marginBottom: 4 }}>💡 Suggestions:</div>
              {suggestions.map((s) => (
                <button key={s.name} className="suggestion-chip" onClick={() => { setName(s.name); if (s.freq) setFreq(s.freq); }}>
                  {s.name} <span style={{ color: "#aaa" }}>· {freqLabel(s.freq)}</span>
                </button>
              ))}
            </div>
          )}
        </div>

        <div className="form-group">
          <label className="form-label">Category</label>
          <select className="form-input" value={categoryId} onChange={(e) => setCategoryId(e.target.value)}>
            {categories.map((c) => <option key={c.id} value={c.id}>{c.emoji} {c.name}</option>)}
          </select>
        </div>

        <div className="form-group">
          <label className="form-label">How often? 🔁</label>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 10 }}>
            {FREQ_PRESETS.map((p) => (
              <button key={p.days} className="pill-btn" onClick={() => setFreq(p.days)}
                style={{ padding: "5px 12px", borderRadius: 16, background: frequencyDays === p.days ? "#FFD93D" : "#f5f5f5", color: frequencyDays === p.days ? "#1A1A2E" : "#666", fontSize: 12 }}>
                {p.label}
              </button>
            ))}
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <span style={{ fontSize: 13, fontWeight: 600, color: "#aaa", whiteSpace: "nowrap" }}>Custom days:</span>
            <input className="form-input" type="number" min={1} value={customInput} placeholder="e.g. 45"
              onChange={(e) => handleCustomInput(e.target.value)}
              style={{ width: 90 }} />
            {frequencyDays && !FREQ_PRESETS.find(p => p.days === frequencyDays) && (
              <span style={{ fontSize: 12, color: "#aaa", fontWeight: 700 }}>≈ {freqLabel(frequencyDays)}</span>
            )}
          </div>
        </div>

        <div className="form-group">
          <label className="form-label">Last Completed</label>
          <input className="form-input" type="date" value={lastCompleted} onChange={(e) => setLastCompleted(e.target.value)} max={today()} />
        </div>

        <div className="form-group">
          <label className="form-label">Assignment 👤</label>
          <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
            {[["any","Anyone can do it"],["fixed","One person only"],["alternate","Alternate turns"]].map(([mode, lbl]) => (
              <button key={mode} className="pill-btn" onClick={() => setAssignMode(mode)}
                style={{ flex: 1, padding: "8px 6px", borderRadius: 12, background: assignMode === mode ? "#4ECDC4" : "#f5f5f5", color: assignMode === mode ? "#fff" : "#555", fontSize: 11, textAlign: "center" }}>
                {mode === "any" ? "🤝" : mode === "fixed" ? "📌" : "🔄"}<br />{lbl}
              </button>
            ))}
          </div>
          {(assignMode === "fixed" || assignMode === "alternate") && (
            <select className="form-input" value={assignedTo} onChange={(e) => setAssignedTo(e.target.value)}>
              {users.map((u) => <option key={u} value={u}>{u}</option>)}
            </select>
          )}
        </div>

        <div style={{ display: "flex", gap: 10, marginTop: 8 }}>
          <button className="pill-btn" onClick={onClose}
            style={{ flex: 1, padding: "12px", borderRadius: 14, background: "#f5f5f5", color: "#666", fontSize: 15 }}>Cancel</button>
          <button className="pill-btn" onClick={handleSave}
            style={{ flex: 2, padding: "12px", borderRadius: 14, background: "linear-gradient(135deg, #FFD93D, #FF9F43)", color: "#fff", fontSize: 15 }}>
            {task ? "Save Changes" : "✨ Add Task"}
          </button>
        </div>
      </div>
    </div>
  );
}

// ── Category Modal ────────────────────────────────────────────────────────────
function CategoryModal({ cat, onSave, onClose }) {
  const [name, setName] = useState(cat?.name || "");
  const [emoji, setEmoji] = useState(cat?.emoji || "📌");
  const [color, setColor] = useState(cat?.color || "#FFD93D");
  const COLORS = ["#FFD93D","#FF6B6B","#4ECDC4","#C77DFF","#06D6A0","#FF9F43","#74B9FF","#FD79A8","#A29BFE","#55EFC4"];
  const EMOJIS = ["📌","⭐","🎯","💪","🏡","🌿","🚗","💰","🎨","🎵","🐾","🌸","🔥","❄️","🎮","📚","🍳","💊","🧹","🌟"];

  return (
    <div className="modal-overlay" onClick={(e) => e.target === e.currentTarget && onClose()}>
      <div className="modal" style={{ maxWidth: 380 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
          <h3 style={{ fontSize: 18, fontWeight: 900, color: "#1A1A2E" }}>{cat ? "Edit Category" : "New Category"}</h3>
          <button className="pill-btn" onClick={onClose} style={{ padding: "6px 12px", borderRadius: 10, background: "#f5f5f5", color: "#888" }}>✕</button>
        </div>
        <div className="form-group">
          <label className="form-label">Name</label>
          <input className="form-input" value={name} onChange={(e) => setName(e.target.value)} placeholder="Category name..." />
        </div>
        <div className="form-group">
          <label className="form-label">Emoji Icon</label>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
            {EMOJIS.map((e) => (
              <button key={e} className="pill-btn" onClick={() => setEmoji(e)}
                style={{ width: 36, height: 36, borderRadius: 10, background: emoji === e ? color : "#f5f5f5", fontSize: 18, border: emoji === e ? `2px solid ${color}` : "2px solid transparent" }}>
                {e}
              </button>
            ))}
          </div>
          <div style={{ marginTop: 8, display: "flex", alignItems: "center", gap: 8 }}>
            <span style={{ fontSize: 13, fontWeight: 600, color: "#aaa" }}>Custom:</span>
            <input className="form-input" value={emoji} onChange={(e) => setEmoji(e.target.value)} style={{ width: 60 }} maxLength={2} />
          </div>
        </div>
        <div className="form-group">
          <label className="form-label">Color</label>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 8 }}>
            {COLORS.map((c) => (
              <button key={c} className="pill-btn" onClick={() => setColor(c)}
                style={{ width: 32, height: 32, borderRadius: "50%", background: c, border: color === c ? "3px solid #1A1A2E" : "3px solid transparent" }} />
            ))}
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <span style={{ fontSize: 13, fontWeight: 600, color: "#aaa" }}>Custom:</span>
            <input type="color" value={color} onChange={(e) => setColor(e.target.value)} style={{ width: 40, height: 32, border: "none", cursor: "pointer", borderRadius: 8 }} />
            <span style={{ fontSize: 13, fontWeight: 600, color: "#aaa" }}>{color}</span>
          </div>
        </div>
        <div style={{ display: "flex", gap: 10 }}>
          <button className="pill-btn" onClick={onClose} style={{ flex: 1, padding: "12px", borderRadius: 14, background: "#f5f5f5", color: "#666" }}>Cancel</button>
          <button className="pill-btn" onClick={() => { if (name.trim()) onSave({ ...(cat || {}), name: name.trim(), emoji, color }); }}
            style={{ flex: 2, padding: "12px", borderRadius: 14, background: color, color: "#fff", fontSize: 15 }}>
            {emoji} Save
          </button>
        </div>
      </div>
    </div>
  );
}
