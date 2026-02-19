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
    return days * 24 + hh + mm / 60;
  }

  m = s.match(/(\d{1,2})\s*:\s*(\d{2})(?:\s*:\s*(\d{2}))?/);
  if (m) {
    const hh = parseInt(m[1], 10);
    const mm = parseInt(m[2], 10);
    return hh + mm / 60;
  }

  return null;
};

const getTimeSlot = (timeObj) => {
  const minutes = timeToMinutes(timeObj);
  if (minutes == null) return null;

  for (const slot of TIME_SLOTS) {
    if (minutes >= slot.startMin && minutes < slot.endMin) {
      return slot.label;
    }
  }
  return null;
};

const dateToMonthKey = (val) => {
  if (!val) return null;

  let d = null;
  if (val instanceof Date && !Number.isNaN(val.getTime())) {
    d = val;
  } else if (typeof val === 'string') {
    d = new Date(val);
    if (Number.isNaN(d.getTime())) return null;
  }

  if (!d) return null;

  const yyyy = d.getFullYear();
  const m = d.getMonth();
  const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 
                      'July', 'August', 'September', 'October', 'November', 'December'];
  return `${monthNames[m]} ${yyyy}`;
};

const parseSemesterFromText = (text) => {
  if (!text) return null;
  const s = String(text).toLowerCase();

  const yearMatch = s.match(/20\d{2}/);
  const year = yearMatch ? yearMatch[0] : null;

  if (s.includes('spring')) {
    return { season: 'Spring', year, label: year ? `Spring ${year}` : 'Spring Semester' };
  }
  if (s.includes('fall')) {
    return { season: 'Fall', year, label: year ? `Fall ${year}` : 'Fall Semester' };
  }
  if (s.includes('summer')) {
    return { season: 'Summer', year, label: year ? `Summer ${year}` : 'Summer Semester' };
  }
  if (s.includes('winter')) {
    return { season: 'Winter', year, label: year ? `Winter ${year}` : 'Winter Semester' };
  }

  return null;
};

const inferSemesterLabel = (fileName, sheetNames) => {
  const candidates = [fileName, ...(sheetNames || [])].filter(Boolean);
  for (const c of candidates) {
    const parsed = parseSemesterFromText(c);
    if (parsed) return parsed.label;
  }
  return 'Semester';
};

const stableMonthSort = (monthKeys) => {
  const order = ['January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'];

  return monthKeys.slice().sort((a, b) => {
    const [mA, yA] = a.split(' ');
    const [mB, yB] = b.split(' ');
    const yDiff = parseInt(yA, 10) - parseInt(yB, 10);
    if (yDiff !== 0) return yDiff;
    return order.indexOf(mA) - order.indexOf(mB);
  });
};

const buildCounts = (rows, keyFn) => {
  const map = new Map();
  for (const r of rows) {
    const key = keyFn(r);
    if (!key) continue;
    map.set(key, (map.get(key) || 0) + 1);
  }
  return map;
};

const buildAverages = (rows, keyFn, valFn) => {
  const sums = new Map();
  const counts = new Map();

  for (const r of rows) {
    const key = keyFn(r);
    const val = valFn(r);
    if (!key || val == null || Number.isNaN(val)) continue;

    sums.set(key, (sums.get(key) || 0) + val);
    counts.set(key, (counts.get(key) || 0) + 1);
  }

  const avgMap = new Map();
  for (const [k, sum] of sums) {
    const cnt = counts.get(k) || 1;
    avgMap.set(k, sum / cnt);
  }
  return avgMap;
};

const mapToChartData = (map, nameKey, valueKey) => {
  return Array.from(map.entries())
    .map(([k, v]) => ({ [nameKey]: k, [valueKey]: v }))
    .filter(x => x[nameKey])
    .sort((a, b) => b[valueKey] - a[valueKey]);
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
  const [viewMode, setViewMode] = useState('none'); // none | month | semester

  const tutoringColumnNames = useMemo(() => {
    if (!columns || columns.length === 0) return { date: null, signIn: null, tutor: null, subject: null, duration: null };

    const date = findColumnName(columns, ['Date', 'Session Date', 'Visit Date']);
    const signIn = findColumnName(columns, ['Sign in Time', 'Sign-in Time', 'Signin Time', 'Sign In Time']);
    const tutor = findColumnName(columns, ['Tutor', 'Tutors']);
    const subject = findColumnName(columns, ['Subject/Class', 'Subject', 'Course', 'Course Name', 'Class']);
    const duration = findColumnName(columns, ['Time', 'Duration', 'Total Time', 'Tutoring Time', 'Time Tutored']);

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
      setViewMode('none');
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
    setViewMode('none');

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

    if (!signInCol || !tutorCol || !subjectCol) {
      console.warn('Missing required columns for tutoring charts');
      return [];
    }

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

    const chartsToReturn = [
      {
        id: `${idPrefix}-timeslot`,
        title: `Students by Time Slot (${labelText})`,
        data: slotChartData,
        dataKey: 'students',
        nameKey: 'timeSlot',
        yLabel: 'Number of Students',
        colorMode: 'timeSlot',
      },
      {
        id: `${idPrefix}-tutor-count`,
        title: `Sessions per Tutor (${labelText})`,
        data: tutorChartData,
        dataKey: 'students',
        nameKey: 'tutor',
        yLabel: 'Number of Sessions',
        colorMode: 'default',
      },
      {
        id: `${idPrefix}-subject-count`,
        title: `Sessions per Subject (${labelText})`,
        data: subjectChartData,
        dataKey: 'students',
        nameKey: 'subject',
        yLabel: 'Number of Sessions',
        colorMode: 'default',
      },
    ];

    // 4 & 5) Duration charts if duration column exists
    if (durationCol) {
      // Tutor avg hours per session
      const tutorAvgHours = buildAverages(
        rows,
        r => String(r?.[tutorCol] ?? '').trim(),
        r => durationToHours(r?.[durationCol])
      );
      const tutorAvgHoursData = mapToChartData(tutorAvgHours, 'tutor', 'avgHours');

      // Subject avg hours per session
      const subjectAvgHours = buildAverages(
        rows,
        r => String(r?.[subjectCol] ?? '').trim(),
        r => durationToHours(r?.[durationCol])
      );
      const subjectAvgHoursData = mapToChartData(subjectAvgHours, 'subject', 'avgHours');

      chartsToReturn.push(
        {
          id: `${idPrefix}-tutor-avg-hours`,
          title: `Average Hours per Tutor (${labelText})`,
          data: tutorAvgHoursData,
          dataKey: 'avgHours',
          nameKey: 'tutor',
          yLabel: 'Average Hours',
          colorMode: 'default',
        },
        {
          id: `${idPrefix}-subject-avg-hours`,
          title: `Average Hours per Subject (${labelText})`,
          data: subjectAvgHoursData,
          dataKey: 'avgHours',
          nameKey: 'subject',
          yLabel: 'Average Hours',
          colorMode: 'default',
        }
      );
    }

    return chartsToReturn;
  };

  const generateMonthCharts = () => {
    if (!isTutoringData || !selectedTutoringMonth) return;

    const rows = filterRowsByMonth(sheetData, tutoringColumnNames.date, selectedTutoringMonth);
    const generatedCharts = buildTutoringCharts(rows, selectedTutoringMonth, tutoringColumnNames, 'month');
    console.log('Generated month charts:', generatedCharts.length, generatedCharts);
    setCharts(generatedCharts);
    setViewMode('month');
  };

  // FIXED: Semester charts should aggregate ALL tutoring sheets from the workbook
  const generateSemesterCharts = () => {
    if (!workbook) return;

    // Find all sheets that look like tutoring data
    const allTutoringRows = [];
    
    for (const sheetName of sheetNames) {
      // Skip summary/data sheets
      if (sheetName.toLowerCase().includes('data from') || 
          sheetName.toLowerCase().includes('list of') ||
          sheetName.toLowerCase().includes('schedule')) {
        continue;
      }

      try {
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: null });

        if (jsonData.length === 0) continue;

        const cols = Object.keys(jsonData[0] || {});
        const date = findColumnName(cols, ['Date', 'Session Date', 'Visit Date']);
        const signIn = findColumnName(cols, ['Sign in Time', 'Sign-in Time', 'Signin Time', 'Sign In Time']);
        const tutor = findColumnName(cols, ['Tutor', 'Tutors']);
        const subject = findColumnName(cols, ['Subject/Class', 'Subject', 'Course', 'Course Name', 'Class']);

        // If this sheet has tutoring columns, add its data
        if (date && signIn && tutor && subject) {
          allTutoringRows.push(...jsonData);
        }
      } catch (err) {
        console.warn(`Skipping sheet ${sheetName}:`, err);
      }
    }

    if (allTutoringRows.length === 0) {
      alert('No tutoring data sheets found in this workbook');
      return;
    }

    const generatedCharts = buildTutoringCharts(allTutoringRows, semesterLabel, tutoringColumnNames, 'semester');
    console.log('Generated semester charts:', generatedCharts.length, generatedCharts);
    setCharts(generatedCharts);
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
    if (!chart || !chart.data || chart.data.length === 0) {
      console.warn('Chart has no data:', chart?.id);
      return <p style={{ padding: '20px', color: '#999' }}>No data available for this chart</p>;
    }

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
            <h2>Select Sheet (for Monthly View)</h2>
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
            <p className="info-text subtle" style={{ marginTop: '12px' }}>
              Select a monthly sheet (e.g., "Feb.", "March") to generate monthly graphs, or skip and generate semester graphs for all data.
            </p>
          </div>
        )}

        {workbook && (
          <div className="actions-panel">
            <h2>Generate Graphs</h2>

            {/* Semester button - ALWAYS VISIBLE when file is uploaded */}
            <div style={{ marginBottom: '20px', paddingBottom: '20px', borderBottom: '1px solid var(--border)' }}>
              <h3 style={{ fontSize: '1.05rem', marginBottom: '10px', color: 'var(--blue-900)' }}>Semester Overview</h3>
              <p className="info-text subtle" style={{ marginBottom: '12px' }}>
                Consolidates data from all tutoring sheets in the workbook into semester-wide charts.
              </p>
              <button
                className="generate-primary"
                onClick={generateSemesterCharts}
                style={{ width: 'auto' }}
              >
                📊 Generate {semesterLabel} Graphs
              </button>
            </div>

            {/* Monthly section - only if sheet selected */}
            {selectedSheet && isTutoringData && (
              <>
                <h3 style={{ fontSize: '1.05rem', marginBottom: '10px', color: 'var(--blue-900)' }}>Monthly View</h3>
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

                <div style={{ marginTop: '14px' }}>
                  <button
                    className="generate-secondary"
                    onClick={generateMonthCharts}
                    disabled={!selectedTutoringMonth}
                  >
                    📅 Generate Monthly Graphs
                  </button>
                </div>
              </>
            )}

            {selectedSheet && !isTutoringData && (
              <div style={{ marginTop: '14px' }}>
                <p className="info-text">
                  Selected sheet "{selectedSheet}" does not appear to be a tutoring log. Try selecting a monthly sheet like "Feb." or "March", or use the Semester Overview button above.
                </p>
              </div>
            )}
          </div>
        )}

        {charts.length > 0 && (
          <div className="charts-container">
            <h2>
              {viewMode === 'semester'
                ? `Semester Consolidated Charts: ${semesterLabel}`
                : viewMode === 'month'
                ? `Monthly Charts: ${selectedTutoringMonth}`
                : 'Generated Charts'}
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