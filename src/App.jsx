import { useEffect, useMemo, useState } from 'react';
import * as XLSX from 'xlsx';
import {
  BarChart, Bar,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, Cell
} from 'recharts';
import html2canvas from 'html2canvas';
import './App.css';

const COLORS = ['#0088FE', '#00C49F', '#FFBB28', '#FF8042', '#8884D8', '#82CA9D', '#FFC658', '#FF6B6B', '#2ECC71', '#E74C3C'];

const TIME_SLOTS = [
  { label: '10:00-11:00', startMin: 10 * 60, endMin: 11 * 60, color: '#4A90E2' },
  { label: '11:00-12:00', startMin: 11 * 60, endMin: 12 * 60, color: '#50C878' },
  { label: '12:00-13:00', startMin: 12 * 60, endMin: 13 * 60, color: '#F5A623' },
  { label: '13:00-14:00', startMin: 13 * 60, endMin: 14 * 60, color: '#E94B3C' },
  { label: '14:00-15:00', startMin: 14 * 60, endMin: 15 * 60, color: '#9B59B6' },
  { label: '15:00-16:00', startMin: 15 * 60, endMin: 16 * 60, color: '#1ABC9C' },
  { label: '16:00-17:00', startMin: 16 * 60, endMin: 17 * 60, color: '#E67E22' },
  { label: '17:00-18:00', startMin: 17 * 60, endMin: 18 * 60, color: '#34495E' },
  { label: '18:00-19:00', startMin: 18 * 60, endMin: 19 * 60, color: '#7F8C8D' },
];

const TIME_SLOT_COLORS = TIME_SLOTS.reduce((acc, s) => {
  acc[s.label] = s.color;
  return acc;
}, {});

const normalizeHeader = (h) => String(h ?? '').trim().toLowerCase().replace(/\s+/g, ' ');
const findColumnName = (headers, candidates) => {
  const map = new Map(headers.map(h => [normalizeHeader(h), h]));
  for (const c of candidates) {
    const key = normalizeHeader(c);
    if (map.has(key)) return map.get(key);
  }
  return null;
};

const excelSerialToDate = (serial) => {
  if (typeof serial !== 'number' || Number.isNaN(serial)) return null;

  if (XLSX?.SSF?.parse_date_code) {
    const d = XLSX.SSF.parse_date_code(serial);
    if (!d) return null;
    return new Date(d.y, (d.m ?? 1) - 1, d.d ?? 1, d.H ?? 0, d.M ?? 0, d.S ?? 0);
  }

  const epoch = new Date(Date.UTC(1899, 11, 30));
  const ms = Math.round(serial * 24 * 60 * 60 * 1000);
  return new Date(epoch.getTime() + ms);
};

const timeToMinutes = (timeObj) => {
  if (timeObj == null) return null;

  if (timeObj instanceof Date && !Number.isNaN(timeObj.getTime())) {
    return timeObj.getHours() * 60 + timeObj.getMinutes();
  }

  if (typeof timeObj === 'number' && !Number.isNaN(timeObj)) {
    const frac = timeObj >= 1 ? (timeObj - Math.floor(timeObj)) : timeObj;
    const totalMinutes = Math.round(frac * 24 * 60);
    if (totalMinutes < 0 || totalMinutes >= 24 * 60) return null;
    return totalMinutes;
  }

  const s = String(timeObj).trim();
  const m = s.match(/(\d{1,2})\s*:\s*(\d{2})(?:\s*:\s*(\d{2}))?\s*(AM|PM)?/i);
  if (!m) return null;

  let hh = parseInt(m[1], 10);
  const mm = parseInt(m[2], 10);
  const ampm = (m[4] || '').toUpperCase();

  if (ampm === 'AM' && hh === 12) hh = 0;
  if (ampm === 'PM' && hh < 12) hh += 12;

  if (hh < 0 || hh > 23 || mm < 0 || mm > 59) return null;
  return hh * 60 + mm;
};

const durationToHours = (val) => {
  if (val == null) return null;

  if (typeof val === 'number' && !Number.isNaN(val)) {
    return val * 24;
  }

  if (val instanceof Date && !Number.isNaN(val.getTime())) {
    return (val.getHours() * 60 + val.getMinutes()) / 60;
  }

  const s = String(val).trim();
  if (!s) return null;

  let m = s.match(/(\d+)\s*day[s]?\s*(\d{1,2}):(\d{2})(?::(\d{2}))?/i);
  if (m) {
    const days = parseInt(m[1], 10);
    const hh = parseInt(m[2], 10);
    const mm = parseInt(m[3], 10);
    const ss = m[4] ? parseInt(m[4], 10) : 0;
    return days * 24 + hh + (mm / 60) + (ss / 3600);
  }

  m = s.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if (m) {
    const hh = parseInt(m[1], 10);
    const mm = parseInt(m[2], 10);
    const ss = m[3] ? parseInt(m[3], 10) : 0;
    return hh + (mm / 60) + (ss / 3600);
  }

  return null;
};

const getTimeSlot = (timeObj) => {
  const mins = timeToMinutes(timeObj);
  if (mins == null) return null;
  const slot = TIME_SLOTS.find(s => mins >= s.startMin && mins < s.endMin);
  return slot ? slot.label : null;
};

const dateToMonthKey = (dateObj) => {
  if (dateObj == null) return null;

  if (dateObj instanceof Date && !Number.isNaN(dateObj.getTime())) {
    return dateObj.toLocaleString('default', { month: 'long', year: 'numeric' });
  }

  if (typeof dateObj === 'number' && !Number.isNaN(dateObj)) {
    const d = excelSerialToDate(dateObj);
    return d ? d.toLocaleString('default', { month: 'long', year: 'numeric' }) : null;
  }

  const s = String(dateObj).trim();
  if (!s) return null;

  const d = new Date(s);
  if (!Number.isNaN(d.getTime())) {
    return d.toLocaleString('default', { month: 'long', year: 'numeric' });
  }

  return null;
};

const stableMonthSort = (months) => {
  const ms = [...months];
  ms.sort((a, b) => {
    const da = new Date(a);
    const db = new Date(b);
    const ta = Number.isNaN(da.getTime()) ? 0 : da.getTime();
    const tb = Number.isNaN(db.getTime()) ? 0 : db.getTime();
    return ta - tb;
  });
  return ms;
};

const buildCounts = (rows, keyFn) => {
  const map = new Map();
  for (const r of rows) {
    const k = keyFn(r);
    if (!k) continue;
    map.set(k, (map.get(k) || 0) + 1);
  }
  return map;
};

const buildAverages = (rows, groupKeyFn, valueFn) => {
  const acc = new Map(); // key -> {sum, n}
  for (const r of rows) {
    const k = groupKeyFn(r);
    if (!k) continue;
    const v = valueFn(r);
    if (v == null || Number.isNaN(v)) continue;
    if (!acc.has(k)) acc.set(k, { sum: 0, n: 0 });
    const cur = acc.get(k);
    cur.sum += v;
    cur.n += 1;
  }
  const out = new Map();
  for (const [k, { sum, n }] of acc.entries()) {
    out.set(k, n > 0 ? sum / n : 0);
  }
  return out;
};

const mapToChartData = (m, keyName, valName, { sortByValueDesc = true } = {}) => {
  const arr = Array.from(m.entries()).map(([k, v]) => ({ [keyName]: k, [valName]: v }));
  if (sortByValueDesc) arr.sort((a, b) => (b[valName] || 0) - (a[valName] || 0));
  return arr;
};

const parseSemesterFromText = (text) => {
  const t = String(text || '');
  const m = t.match(/\b(Fall|Spring|Summer|Winter)\s*(Semester\s*)?(\d{4})\b/i);
  if (!m) return null;
  const term = m[1].charAt(0).toUpperCase() + m[1].slice(1).toLowerCase();
  const year = m[3];
  return { term, year, label: `${term} Semester ${year}` };
};

const inferSemesterLabel = (fileName, sheetNames) => {
  const candidates = [fileName, ...(sheetNames || [])].filter(Boolean);
  for (const c of candidates) {
    const parsed = parseSemesterFromText(c);
    if (parsed) return parsed.label;
  }
  return 'Semester';
};

function App() {
  const [workbook, setWorkbook] = useState(null);
  const [sheetNames, setSheetNames] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState('');
  const [sheetData, setSheetData] = useState([]);
  const [columns, setColumns] = useState([]);
  const [charts, setCharts] = useState([]);
  const [isTutoringData, setIsTutoringData] = useState(false);

  const [uploadedFileName, setUploadedFileName] = useState('');
  const semesterLabel = useMemo(() => inferSemesterLabel(uploadedFileName, sheetNames), [uploadedFileName, sheetNames]);

  const [tutoringMonthOptions, setTutoringMonthOptions] = useState([]);
  const [selectedTutoringMonth, setSelectedTutoringMonth] = useState('');
  const [viewMode, setViewMode] = useState('month'); // month | semester

  const tutoringColumnNames = useMemo(() => {
    if (!columns || columns.length === 0) return { date: null, signIn: null, tutor: null, subject: null, duration: null };

    const date = findColumnName(columns, ['Date', 'Session Date', 'Visit Date']);
    const signIn = findColumnName(columns, ['Sign in Time', 'Sign-in Time', 'Signin Time', 'Sign In Time']);
    const tutor = findColumnName(columns, ['Tutor', 'Tutors']);
    const subject = findColumnName(columns, ['Subject/Class', 'Subject', 'Course', 'Course Name', 'Class']);
    const duration = findColumnName(columns, ['Time', 'Duration', 'Total Time', 'Tutoring Time']);

    return { date, signIn, tutor, subject, duration };
  }, [columns]);

  const handleFileUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setUploadedFileName(file.name);

    const reader = new FileReader();
    reader.onload = (event) => {
      const data = new Uint8Array(event.target.result);
      const wb = XLSX.read(data, { type: 'array', cellDates: true });

      setWorkbook(wb);
      setSheetNames(wb.SheetNames);

      setSelectedSheet('');
      setSheetData([]);
      setColumns([]);
      setCharts([]);
      setIsTutoringData(false);
      setTutoringMonthOptions([]);
      setSelectedTutoringMonth('');
      setViewMode('month');
    };
    reader.readAsArrayBuffer(file);
  };

  const handleSheetSelect = (sheetName) => {
    if (!workbook) return;

    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: null });

    setSelectedSheet(sheetName);
    setSheetData(jsonData);
    setCharts([]);
    setViewMode('month');

    if (jsonData.length > 0) {
      const cols = Object.keys(jsonData[0] || {});
      setColumns(cols);

      const date = findColumnName(cols, ['Date', 'Session Date', 'Visit Date']);
      const signIn = findColumnName(cols, ['Sign in Time', 'Sign-in Time', 'Signin Time', 'Sign In Time']);
      const tutor = findColumnName(cols, ['Tutor', 'Tutors']);
      const subject = findColumnName(cols, ['Subject/Class', 'Subject', 'Course', 'Course Name', 'Class']);

      const hasTutoringColumns = Boolean(date && signIn && tutor && subject);
      setIsTutoringData(hasTutoringColumns);

      if (hasTutoringColumns) {
        const months = new Set();
        for (const row of jsonData) {
          const mk = dateToMonthKey(row?.[date]);
          if (mk) months.add(mk);
        }
        const sorted = stableMonthSort(Array.from(months));
        setTutoringMonthOptions(sorted);
        setSelectedTutoringMonth(sorted[0] || '');
      } else {
        setTutoringMonthOptions([]);
        setSelectedTutoringMonth('');
      }
    }
  };

  const filterRowsByMonth = (rows, dateCol, monthKey) => {
    if (!monthKey) return rows;
    return rows.filter(r => dateToMonthKey(r?.[dateCol]) === monthKey);
  };

  const buildTutoringCharts = (rows, labelText, colNames, idPrefix) => {
    const signInCol = colNames.signIn;
    const tutorCol = colNames.tutor;
    const subjectCol = colNames.subject;
    const durationCol = colNames.duration;

    // 1) Hourly students
    const slotCounts = new Map(TIME_SLOTS.map(s => [s.label, 0]));
    for (const r of rows) {
      const slot = getTimeSlot(r?.[signInCol]);
      if (!slot) continue;
      slotCounts.set(slot, (slotCounts.get(slot) || 0) + 1);
    }
    const slotChartData = TIME_SLOTS.map(s => ({ timeSlot: s.label, students: slotCounts.get(s.label) || 0 }));

    // 2) Tutor count
    const tutorCounts = buildCounts(rows, r => String(r?.[tutorCol] ?? '').trim());
    const tutorChartData = mapToChartData(tutorCounts, 'tutor', 'students');

    // 3) Subject count
    const subjectCounts = buildCounts(rows, r => String(r?.[subjectCol] ?? '').trim());
    const subjectChartData = mapToChartData(subjectCounts, 'subject', 'students');

    // 4) Tutor avg hours per session
    const tutorAvgHours = buildAverages(
      rows,
      r => String(r?.[tutorCol] ?? '').trim(),
      r => durationToHours(r?.[durationCol])
    );
    const tutorAvgHoursData = mapToChartData(tutorAvgHours, 'tutor', 'avgHours');

    // 5) Subject avg hours per session
    const subjectAvgHours = buildAverages(
      rows,
      r => String(r?.[subjectCol] ?? '').trim(),
      r => durationToHours(r?.[durationCol])
    );
    const subjectAvgHoursData = mapToChartData(subjectAvgHours, 'subject', 'avgHours');

    return [
      {
        id: `${idPrefix}-timeslot`,
        type: 'bar',
        title: `Hourly Number of Students in ${labelText}`,
        data: slotChartData,
        dataKey: 'students',
        nameKey: 'timeSlot',
        yLabel: 'Number of Students',
        colorMode: 'timeSlot',
      },
      {
        id: `${idPrefix}-tutor-count`,
        type: 'bar',
        title: `Tutor vs Number of Students in ${labelText}`,
        data: tutorChartData,
        dataKey: 'students',
        nameKey: 'tutor',
        yLabel: 'Number of Students',
      },
      {
        id: `${idPrefix}-subject-count`,
        type: 'bar',
        title: `Subject vs Number of Students in ${labelText}`,
        data: subjectChartData,
        dataKey: 'students',
        nameKey: 'subject',
        yLabel: 'Number of Students',
      },
      {
        id: `${idPrefix}-tutor-avg-hours`,
        type: 'bar',
        title: `Average Tutoring Hours per Session by Tutor in ${labelText}`,
        data: tutorAvgHoursData,
        dataKey: 'avgHours',
        nameKey: 'tutor',
        yLabel: 'Average Hours',
      },
      {
        id: `${idPrefix}-subject-avg-hours`,
        type: 'bar',
        title: `Average Tutoring Hours per Session by Subject in ${labelText}`,
        data: subjectAvgHoursData,
        dataKey: 'avgHours',
        nameKey: 'subject',
        yLabel: 'Average Hours',
      },
    ];
  };

  const generateMonthCharts = () => {
    if (!isTutoringData || !selectedTutoringMonth) return;

    const rows = filterRowsByMonth(sheetData, tutoringColumnNames.date, selectedTutoringMonth);
    setCharts(buildTutoringCharts(rows, selectedTutoringMonth, tutoringColumnNames, 'month'));
    setViewMode('month');
  };

  const generateSemesterCharts = () => {
    if (!isTutoringData) return;

    setCharts(buildTutoringCharts(sheetData, semesterLabel, tutoringColumnNames, 'semester'));
    setViewMode('semester');
  };

  // Keep month charts synced when month changes and month view already displayed
  useEffect(() => {
    if (!isTutoringData) return;
    if (charts.length === 0) return;
    if (viewMode !== 'month') return;
    generateMonthCharts();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedTutoringMonth]);

  const downloadChartAsPNG = async (chartId) => {
    const element = document.getElementById(chartId);
    if (!element) return;

    const canvas = await html2canvas(element, { backgroundColor: '#ffffff', scale: 2 });
    const link = document.createElement('a');
    link.download = `${chartId}.png`;
    link.href = canvas.toDataURL('image/png');
    link.click();
  };

  const downloadChartAsSVG = (chartId) => {
    const element = document.getElementById(chartId);
    if (!element) return;

    const svg = element.querySelector('svg');
    if (!svg) return;

    const serializer = new XMLSerializer();
    const source = serializer.serializeToString(svg);

    const blob = new Blob([source], { type: 'image/svg+xml;charset=utf-8' });
    const url = URL.createObjectURL(blob);

    const link = document.createElement('a');
    link.href = url;
    link.download = `${chartId}.svg`;
    link.click();

    URL.revokeObjectURL(url);
  };

  const renderChart = (chart) => {
    if (!chart || !chart.data) return null;

    return (
      <ResponsiveContainer width="100%" height={430}>
        <BarChart data={chart.data} margin={{ top: 20, right: 30, left: 20, bottom: 90 }}>
          <CartesianGrid strokeDasharray="3 3" />
          <XAxis dataKey={chart.nameKey} angle={-28} textAnchor="end" interval={0} height={90} />
          <YAxis />
          <Tooltip />
          <Legend />
          <Bar dataKey={chart.dataKey} name={chart.yLabel}>
            {chart.data.map((entry, index) => (
              <Cell
                key={`cell-${index}`}
                fill={chart.colorMode === 'timeSlot'
                  ? (TIME_SLOT_COLORS[entry[chart.nameKey]] || COLORS[index % COLORS.length])
                  : (COLORS[index % COLORS.length])}
              />
            ))}
          </Bar>
        </BarChart>
      </ResponsiveContainer>
    );
  };

  return (
    <div className="App">
      <header className="header">
        <div className="header-content">
          <h1>📊 Excel Analyzer</h1>
          <p>Tutoring analytics: monthly and semester performance</p>
        </div>
      </header>

      <div className="main-content">
        <div className="upload-section">
          <h2>Upload Excel File</h2>
          <input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} className="file-input" />
        </div>

        {sheetNames.length > 0 && (
          <div className="sheet-selection">
            <h2>Select Sheet</h2>
            <div className="sheet-buttons">
              {sheetNames.map((name) => (
                <button
                  key={name}
                  className={`sheet-button ${selectedSheet === name ? 'active' : ''}`}
                  onClick={() => handleSheetSelect(name)}
                >
                  {name}
                </button>
              ))}
            </div>
          </div>
        )}

        {workbook && (
          <div className="actions-panel">
            <h2>Generate Graphs</h2>

            {!selectedSheet && (
              <p className="info-text">
                Select a sheet first. For your file, choose a monthly sheet such as “Sep. 2025” or “Oct. 2025”.
              </p>
            )}

            {selectedSheet && !isTutoringData && (
              <p className="info-text">
                This sheet does not look like a tutoring log. Select a monthly tutoring sheet such as “Sep. 2025”.
              </p>
            )}

            {selectedSheet && isTutoringData && (
              <>
                <div className="tutoring-controls">
                  <label className="control-label" htmlFor="monthSelect">Month</label>
                  <select
                    id="monthSelect"
                    className="control-select"
                    value={selectedTutoringMonth}
                    onChange={(e) => setSelectedTutoringMonth(e.target.value)}
                  >
                    {tutoringMonthOptions.map(m => (<option key={m} value={m}>{m}</option>))}
                  </select>
                </div>

                <p className="info-text subtle">
                  Monthly chart titles follow “Month YYYY”. Semester chart titles follow “{semesterLabel}”.
                </p>
              </>
            )}

            <div className="generate-buttons-row">
              <button
                className={`generate-primary ${viewMode === 'month' ? 'active' : ''}`}
                onClick={generateMonthCharts}
                disabled={!selectedSheet || !isTutoringData || !selectedTutoringMonth}
              >
                Generate Monthly Graphs
              </button>

              <button
                className={`generate-secondary ${viewMode === 'semester' ? 'active' : ''}`}
                onClick={generateSemesterCharts}
                disabled={!selectedSheet || !isTutoringData}
              >
                Generate Semester Graphs
              </button>
            </div>
          </div>
        )}

        {charts.length > 0 && (
          <div className="charts-container">
            <h2>
              {viewMode === 'semester'
                ? `Semester Consolidated Charts: ${semesterLabel}`
                : `Monthly Charts: ${selectedTutoringMonth}`}
            </h2>

            {charts.map(chart => (
              <div key={chart.id} className="chart-card">
                <div className="chart-header">
                  <h3>{chart.title}</h3>
                  <div className="download-buttons">
                    <button onClick={() => downloadChartAsPNG(chart.id)}>📥 PNG</button>
                    <button onClick={() => downloadChartAsSVG(chart.id)}>📥 SVG</button>
                  </div>
                </div>
                <div id={chart.id} className="chart-content">
                  {renderChart(chart)}
                </div>
              </div>
            ))}
          </div>
        )}

        {!workbook && (
          <div className="empty-state">
            <p>Upload an Excel file to get started.</p>
          </div>
        )}
      </div>

      <footer className="footer">
        <p>Built with React, Recharts, and SheetJS</p>
      </footer>
    </div>
  );
}

export default App;