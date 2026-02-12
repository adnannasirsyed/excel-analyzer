import { useEffect, useMemo, useState } from 'react';
import * as XLSX from 'xlsx';
import {
  BarChart, Bar, LineChart, Line, PieChart, Pie, ScatterChart, Scatter,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, Cell
} from 'recharts';
import html2canvas from 'html2canvas';
import './App.css';

// Custom color palettes
const COLORS = ['#0088FE', '#00C49F', '#FFBB28', '#FF8042', '#8884D8', '#82CA9D', '#FFC658', '#FF6B6B'];

// Hour-long time slots from 10:30 to 18:30 (inclusive start, exclusive end)
const TIME_SLOTS = [
  { label: '10:30-11:30', startMin: 10 * 60 + 30, endMin: 11 * 60 + 30, color: '#4A90E2' },
  { label: '11:30-12:30', startMin: 11 * 60 + 30, endMin: 12 * 60 + 30, color: '#50C878' },
  { label: '12:30-13:30', startMin: 12 * 60 + 30, endMin: 13 * 60 + 30, color: '#F5A623' },
  { label: '13:30-14:30', startMin: 13 * 60 + 30, endMin: 14 * 60 + 30, color: '#E94B3C' },
  { label: '14:30-15:30', startMin: 14 * 60 + 30, endMin: 15 * 60 + 30, color: '#9B59B6' },
  { label: '15:30-16:30', startMin: 15 * 60 + 30, endMin: 16 * 60 + 30, color: '#1ABC9C' },
  { label: '16:30-17:30', startMin: 16 * 60 + 30, endMin: 17 * 60 + 30, color: '#E67E22' },
  { label: '17:30-18:30', startMin: 17 * 60 + 30, endMin: 18 * 60 + 30, color: '#34495E' },
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
  // Excel serial date (Windows): days since 1899-12-30
  // SheetJS provides SSF.parse_date_code which is more robust for edge cases,
  // but keep a fallback for environments where SSF isn't present.
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

  // 1) Date object
  if (timeObj instanceof Date && !Number.isNaN(timeObj.getTime())) {
    return timeObj.getHours() * 60 + timeObj.getMinutes();
  }

  // 2) Excel numeric time (fraction of day) or datetime serial (>= 1)
  if (typeof timeObj === 'number' && !Number.isNaN(timeObj)) {
    const frac = timeObj >= 1 ? (timeObj - Math.floor(timeObj)) : timeObj;
    const totalMinutes = Math.round(frac * 24 * 60);
    if (totalMinutes < 0 || totalMinutes >= 24 * 60) return null;
    return totalMinutes;
  }

  // 3) Object with hour/minute
  if (typeof timeObj === 'object') {
    const hour = ('hour' in timeObj ? timeObj.hour : ('hours' in timeObj ? timeObj.hours : null));
    const minute = ('minute' in timeObj ? timeObj.minute : ('minutes' in timeObj ? timeObj.minutes : null));
    if (hour != null && minute != null) {
      const h = Number(hour);
      const m = Number(minute);
      if (!Number.isNaN(h) && !Number.isNaN(m)) return h * 60 + m;
    }
  }

  // 4) String formats: "10:45", "10:45 AM", "10:45:00", "10:45:00 PM"
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

// Returns:
// {
//   bySlot: { [slotLabel]: { total: number, byMonth: { [monthKey]: number } } },
//   months: string[]
// }
const processTutoringData = (data, signInColName, dateColName) => {
  const bySlot = {};
  const monthSet = new Set();

  for (const row of data) {
    const signInTime = row?.[signInColName];
    const dateVal = row?.[dateColName];

    const slot = getTimeSlot(signInTime);
    const monthKey = dateToMonthKey(dateVal);

    if (!slot) continue;

    if (!bySlot[slot]) bySlot[slot] = { total: 0, byMonth: {} };
    bySlot[slot].total += 1;

    if (monthKey) {
      bySlot[slot].byMonth[monthKey] = (bySlot[slot].byMonth[monthKey] || 0) + 1;
      monthSet.add(monthKey);
    }
  }

  // Stable month ordering: chronological by parsing "Month YYYY"
  const months = Array.from(monthSet);
  months.sort((a, b) => {
    const da = new Date(a);
    const db = new Date(b);
    const ta = Number.isNaN(da.getTime()) ? 0 : da.getTime();
    const tb = Number.isNaN(db.getTime()) ? 0 : db.getTime();
    return ta - tb;
  });

  // Ensure all time slots exist (even if 0)
  for (const s of TIME_SLOTS) {
    if (!bySlot[s.label]) bySlot[s.label] = { total: 0, byMonth: {} };
  }

  return { bySlot, months };
};

function App() {
  const [workbook, setWorkbook] = useState(null);
  const [sheetNames, setSheetNames] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState('');
  const [sheetData, setSheetData] = useState([]);
  const [columns, setColumns] = useState([]);
  const [charts, setCharts] = useState([]);
  const [isTutoringData, setIsTutoringData] = useState(false);

  // Tutoring-specific state
  const [tutoringMonthOptions, setTutoringMonthOptions] = useState([]);
  const [selectedTutoringMonth, setSelectedTutoringMonth] = useState('ALL');

  const tutoringColumnNames = useMemo(() => {
    if (!columns || columns.length === 0) return { signIn: null, date: null };
    const signIn = findColumnName(columns, ['Sign in Time', 'Sign-in Time', 'Signin Time', 'Sign In Time']);
    const date = findColumnName(columns, ['Date', 'Session Date', 'Visit Date']);
    return { signIn, date };
  }, [columns]);

  const handleFileUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const data = new Uint8Array(event.target.result);
      // cellDates helps SheetJS convert date-formatted cells into Date objects,
      // but we still keep robust numeric parsing for mixed spreadsheets.
      const wb = XLSX.read(data, { type: 'array', cellDates: true });

      setWorkbook(wb);
      setSheetNames(wb.SheetNames);
      setSelectedSheet('');
      setSheetData([]);
      setColumns([]);
      setCharts([]);
      setIsTutoringData(false);
      setTutoringMonthOptions([]);
      setSelectedTutoringMonth('ALL');
    };
    reader.readAsArrayBuffer(file);
  };

  const handleSheetSelect = (sheetName) => {
    if (!workbook) return;

    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: null });

    setSelectedSheet(sheetName);
    setSheetData(jsonData);

    if (jsonData.length > 0) {
      const cols = Object.keys(jsonData[0] || {});
      setColumns(cols);

      const signIn = findColumnName(cols, ['Sign in Time', 'Sign-in Time', 'Signin Time', 'Sign In Time']);
      const date = findColumnName(cols, ['Date', 'Session Date', 'Visit Date']);

      const hasTutoringColumns = Boolean(signIn && date);
      setIsTutoringData(hasTutoringColumns);

      if (hasTutoringColumns) {
        const processed = processTutoringData(jsonData, signIn, date);
        setTutoringMonthOptions(processed.months);
        setSelectedTutoringMonth('ALL');
      } else {
        setTutoringMonthOptions([]);
        setSelectedTutoringMonth('ALL');
      }
    }
  };

  const getNumericColumns = () => {
    if (!sheetData || sheetData.length === 0) return [];
    const sample = sheetData[0];

    return Object.keys(sample).filter(col => {
      return sheetData.some(row => typeof row[col] === 'number' && !Number.isNaN(row[col]));
    });
  };

  const generateCharts = () => {
    const newCharts = [];

    // Specialized tutoring charts
    if (isTutoringData && tutoringColumnNames.signIn && tutoringColumnNames.date) {
      const { bySlot, months } = processTutoringData(sheetData, tutoringColumnNames.signIn, tutoringColumnNames.date);

      // MAIN CHART: Students by time slot for a selected month (or ALL)
      const timeSlotChartData = TIME_SLOTS.map(s => {
        const slotData = bySlot[s.label] || { total: 0, byMonth: {} };
        const students = selectedTutoringMonth === 'ALL'
          ? (slotData.total || 0)
          : (slotData.byMonth?.[selectedTutoringMonth] || 0);
        return { timeSlot: s.label, students };
      });

      const mainTitle = selectedTutoringMonth === 'ALL'
        ? 'Number of Students by Time Slot (All Months)'
        : `Number of Students by Time Slot (${selectedTutoringMonth})`;

      newCharts.push({
        id: 'tutoring-timeslot-month-bar',
        type: 'bar',
        title: mainTitle,
        data: timeSlotChartData,
        dataKey: 'students',
        nameKey: 'timeSlot',
        xLabel: 'Time Slot',
        yLabel: 'Number of Students',
      });

      // Optional: Stacked bar by month (useful to spot seasonal shifts)
      if (months.length > 1) {
        const monthStackData = months.map(m => {
          const row = { month: m };
          for (const s of TIME_SLOTS) {
            row[s.label] = bySlot[s.label]?.byMonth?.[m] || 0;
          }
          return row;
        });

        newCharts.push({
          id: 'tutoring-month-stacked',
          type: 'stackedBar',
          title: 'Students per Month (Stacked by Time Slot)',
          data: monthStackData,
          nameKey: 'month',
          xLabel: 'Month',
          yLabel: 'Number of Students',
          stacks: TIME_SLOTS.map(s => s.label)
        });
      }

      setCharts(newCharts);
      return;
    }

    // Generic charts for non-tutoring sheets
    const numericCols = getNumericColumns();
    if (numericCols.length === 0) {
      setCharts([]);
      return;
    }

    // Build a simple frequency chart for the first numeric column
    const col = numericCols[0];
    const data = sheetData.map((row, idx) => ({
      index: idx + 1,
      value: row[col]
    }));

    newCharts.push({
      id: 'generic-line',
      type: 'line',
      title: `Line Chart: ${col}`,
      data,
      dataKey: 'value',
      nameKey: 'index',
      xLabel: 'Row',
      yLabel: col
    });

    setCharts(newCharts);
  };

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

    const common = {
      width: '100%',
      height: 400,
    };

    if (chart.type === 'bar') {
      return (
        <ResponsiveContainer {...common}>
          <BarChart data={chart.data} margin={{ top: 20, right: 30, left: 20, bottom: 40 }}>
            <CartesianGrid strokeDasharray="3 3" />
            <XAxis dataKey={chart.nameKey} angle={-20} textAnchor="end" interval={0} />
            <YAxis />
            <Tooltip />
            <Legend />
            <Bar dataKey={chart.dataKey} name={chart.yLabel}>
              {chart.data.map((entry, index) => (
                <Cell
                  key={`cell-${index}`}
                  fill={TIME_SLOT_COLORS[entry[chart.nameKey]] || COLORS[index % COLORS.length]}
                />
              ))}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      );
    }

    if (chart.type === 'stackedBar') {
      return (
        <ResponsiveContainer {...common}>
          <BarChart data={chart.data} margin={{ top: 20, right: 30, left: 20, bottom: 40 }}>
            <CartesianGrid strokeDasharray="3 3" />
            <XAxis dataKey={chart.nameKey} angle={-20} textAnchor="end" interval={0} />
            <YAxis />
            <Tooltip />
            <Legend />
            {chart.stacks.map((k, idx) => (
              <Bar key={k} dataKey={k} stackId="a" fill={TIME_SLOT_COLORS[k] || COLORS[idx % COLORS.length]} />
            ))}
          </BarChart>
        </ResponsiveContainer>
      );
    }

    if (chart.type === 'line') {
      return (
        <ResponsiveContainer {...common}>
          <LineChart data={chart.data} margin={{ top: 20, right: 30, left: 20, bottom: 40 }}>
            <CartesianGrid strokeDasharray="3 3" />
            <XAxis dataKey={chart.nameKey} />
            <YAxis />
            <Tooltip />
            <Legend />
            <Line type="monotone" dataKey={chart.dataKey} dot={false} />
          </LineChart>
        </ResponsiveContainer>
      );
    }

    if (chart.type === 'pie') {
      return (
        <ResponsiveContainer {...common}>
          <PieChart>
            <Tooltip />
            <Legend />
            <Pie data={chart.data} dataKey={chart.dataKey} nameKey={chart.nameKey} outerRadius={140} label>
              {chart.data.map((entry, index) => (
                <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
              ))}
            </Pie>
          </PieChart>
        </ResponsiveContainer>
      );
    }

    if (chart.type === 'scatter') {
      return (
        <ResponsiveContainer {...common}>
          <ScatterChart margin={{ top: 20, right: 30, left: 20, bottom: 40 }}>
            <CartesianGrid />
            <XAxis type="number" dataKey={chart.xKey} name={chart.xLabel} />
            <YAxis type="number" dataKey={chart.yKey} name={chart.yLabel} />
            <Tooltip cursor={{ strokeDasharray: '3 3' }} />
            <Legend />
            <Scatter data={chart.data} />
          </ScatterChart>
        </ResponsiveContainer>
      );
    }

    return null;
  };

  // If charts already exist, keep them synced when month changes
  useEffect(() => {
    if (charts.length === 0) return;
    if (!isTutoringData) return;
    generateCharts();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedTutoringMonth]);

  return (
    <div className="App">
      <header className="header">
        <div className="header-content">
          <h1>📊 Excel Analyzer</h1>
          <p>Upload an Excel file and generate insightful charts</p>
        </div>
      </header>

      <div className="main-content">
        {/* Upload Section */}
        <div className="upload-section">
          <h2>Upload Excel File</h2>
          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={handleFileUpload}
            className="file-input"
          />
        </div>

        {/* Sheet Selection */}
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

        {/* Generate Section */}
        {selectedSheet && sheetData.length > 0 && (
          <div className="generate-section">
            {isTutoringData && tutoringMonthOptions.length > 0 && (
              <div className="tutoring-controls">
                <label className="control-label" htmlFor="monthSelect">
                  Month
                </label>
                <select
                  id="monthSelect"
                  className="control-select"
                  value={selectedTutoringMonth}
                  onChange={(e) => setSelectedTutoringMonth(e.target.value)}
                >
                  <option value="ALL">All months</option>
                  {tutoringMonthOptions.map(m => (
                    <option key={m} value={m}>{m}</option>
                  ))}
                </select>
              </div>
            )}

            <button className="generate-button" onClick={generateCharts}>
              🎨 Generate Charts
            </button>

            <p className="info-text">
              Found {sheetData.length} rows and {columns.length} columns
              {isTutoringData && <span className="tutoring-badge"> • Tutoring Data Detected</span>}
            </p>

            {isTutoringData && tutoringColumnNames.signIn && tutoringColumnNames.date && (
              <p className="info-text subtle">
                Using columns: <strong>{tutoringColumnNames.date}</strong> (date) and <strong>{tutoringColumnNames.signIn}</strong> (sign-in time)
              </p>
            )}
          </div>
        )}

        {/* Charts Display */}
        {charts.length > 0 && (
          <div className="charts-container">
            <h2>Generated Charts ({charts.length})</h2>
            {charts.map(chart => (
              <div key={chart.id} className="chart-card">
                <div className="chart-header">
                  <h3>{chart.title}</h3>
                  <div className="download-buttons">
                    <button onClick={() => downloadChartAsPNG(chart.id)}>
                      📥 PNG
                    </button>
                    <button onClick={() => downloadChartAsSVG(chart.id)}>
                      📥 SVG
                    </button>
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
            <p>👆 Upload an Excel file to get started</p>
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