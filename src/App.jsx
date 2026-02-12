import { useState } from 'react';
import * as XLSX from 'xlsx';
import { 
  BarChart, Bar, LineChart, Line, PieChart, Pie, ScatterChart, Scatter,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, Cell 
} from 'recharts';
import html2canvas from 'html2canvas';
import './App.css';

// Custom color palettes
const COLORS = ['#0088FE', '#00C49F', '#FFBB28', '#FF8042', '#8884D8', '#82CA9D', '#FFC658', '#FF6B6B'];
const TIME_SLOT_COLORS = {
  '10:30-11:30': '#4A90E2',
  '11:30-12:30': '#50C878',
  '12:30-13:30': '#F5A623',
  '13:30-14:30': '#E94B3C',
  '14:30-15:30': '#9B59B6',
  '15:30-16:30': '#1ABC9C',
  '16:30-17:30': '#E67E22',
  '17:30-18:30': '#34495E'
};

function App() {
  const [workbook, setWorkbook] = useState(null);
  const [sheetNames, setSheetNames] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState('');
  const [sheetData, setSheetData] = useState([]);
  const [columns, setColumns] = useState([]);
  const [charts, setCharts] = useState([]);
  const [isTutoringData, setIsTutoringData] = useState(false);

  // Convert time string to hour (for grouping)
  const getTimeSlot = (timeObj) => {
    if (!timeObj) return null;
    
    let hour, minute;
    
    // Handle datetime.time objects (from Excel)
    if (typeof timeObj === 'object' && timeObj !== null) {
      // If it's a Date object
      if (timeObj instanceof Date) {
        hour = timeObj.getHours();
        minute = timeObj.getMinutes();
      }
      // If it has hour/minute properties
      else if ('hour' in timeObj || 'hours' in timeObj) {
        hour = timeObj.hour || timeObj.hours || 0;
        minute = timeObj.minute || timeObj.minutes || 0;
      }
      else {
        // Convert to string and parse
        const timeStr = timeObj.toString();
        const timeMatch = timeStr.match(/(\d{1,2}):(\d{2})/);
        if (!timeMatch) return null;
        hour = parseInt(timeMatch[1]);
        minute = parseInt(timeMatch[2]);
      }
    }
    // Handle string format
    else {
      const timeMatch = timeObj.toString().match(/(\d{1,2}):(\d{2})/);
      if (!timeMatch) return null;
      hour = parseInt(timeMatch[1]);
      minute = parseInt(timeMatch[2]);
    }
    
    // Define time slots (1-hour intervals)
    if (hour === 10 && minute >= 30) return '10:30-11:30';
    if (hour === 11 && minute < 30) return '10:30-11:30';
    if (hour === 11 && minute >= 30) return '11:30-12:30';
    if (hour === 12 && minute < 30) return '11:30-12:30';
    if (hour === 12 && minute >= 30) return '12:30-13:30';
    if (hour === 13 && minute < 30) return '12:30-13:30';
    if (hour === 13 && minute >= 30) return '13:30-14:30';
    if (hour === 14 && minute < 30) return '13:30-14:30';
    if (hour === 14 && minute >= 30) return '14:30-15:30';
    if (hour === 15 && minute < 30) return '14:30-15:30';
    if (hour === 15 && minute >= 30) return '15:30-16:30';
    if (hour === 16 && minute < 30) return '15:30-16:30';
    if (hour === 16 && minute >= 30) return '16:30-17:30';
    if (hour === 17 && minute < 30) return '16:30-17:30';
    if (hour === 17 && minute >= 30) return '17:30-18:30';
    if (hour === 18 && minute < 30) return '17:30-18:30';
    
    return null;
  };

  // Process tutoring data to create time slot analysis
  const processTutoringData = (data) => {
    const monthlyTimeSlots = {};
    
    data.forEach(row => {
      const signInTime = row['Sign in Time'];
      const date = row['Date'];
      
      if (signInTime && date) {
        const timeSlot = getTimeSlot(signInTime);
        
        if (timeSlot) {
          // Extract month from date
          let month = '';
          if (date instanceof Date) {
            month = date.toLocaleString('default', { month: 'long', year: 'numeric' });
          } else if (typeof date === 'string') {
            const dateObj = new Date(date);
            if (!isNaN(dateObj)) {
              month = dateObj.toLocaleString('default', { month: 'long', year: 'numeric' });
            }
          }
          
          if (!monthlyTimeSlots[timeSlot]) {
            monthlyTimeSlots[timeSlot] = { count: 0, months: {} };
          }
          monthlyTimeSlots[timeSlot].count++;
          
          if (month) {
            if (!monthlyTimeSlots[timeSlot].months[month]) {
              monthlyTimeSlots[timeSlot].months[month] = 0;
            }
            monthlyTimeSlots[timeSlot].months[month]++;
          }
        }
      }
    });
    
    return monthlyTimeSlots;
  };

  // Handle file upload
  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const data = new Uint8Array(event.target.result);
      const wb = XLSX.read(data, { type: 'array' });
      
      setWorkbook(wb);
      setSheetNames(wb.SheetNames);
      setSelectedSheet('');
      setSheetData([]);
      setCharts([]);
    };
    reader.readAsArrayBuffer(file);
  };

  // Handle sheet selection
  const handleSheetSelect = (sheetName) => {
    if (!workbook) return;

    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: null });
    
    setSelectedSheet(sheetName);
    setSheetData(jsonData);
    
    // Extract column names
    if (jsonData.length > 0) {
      const cols = Object.keys(jsonData[0]);
      setColumns(cols);
      
      // Check if this is tutoring data (has Sign in Time column)
      const hasTutoringColumns = cols.includes('Sign in Time') && cols.includes('Date');
      setIsTutoringData(hasTutoringColumns);
    }
  };

  // Detect numeric columns
  const getNumericColumns = () => {
    if (sheetData.length === 0) return [];
    
    return columns.filter(col => {
      const values = sheetData.slice(0, 10).map(row => row[col]);
      const numericCount = values.filter(v => v !== null && !isNaN(Number(v))).length;
      return numericCount > values.length * 0.5; // At least 50% numeric
    });
  };

  // Generate all possible charts
  const generateCharts = () => {
    const newCharts = [];
    
    // If this is tutoring data, create specialized tutoring charts
    if (isTutoringData) {
      const timeSlotData = processTutoringData(sheetData);
      
      // Chart 1: Students per Time Slot (Bar Chart) - THE MAIN CHART YOU REQUESTED
      const timeSlotChartData = Object.keys(timeSlotData)
        .sort()
        .map(slot => ({
          timeSlot: slot,
          students: timeSlotData[slot].count
        }));
      
      if (timeSlotChartData.length > 0) {
        newCharts.push({
          id: 'tutoring-timeslot-bar',
          type: 'bar',
          title: 'Number of Students by Tutoring Time Slot',
          data: timeSlotChartData,
          dataKey: 'students',
          nameKey: 'timeSlot',
          xLabel: 'Time Slot',
          yLabel: 'Number of Students',
          color: '#4A90E2'
        });
      }
      
      // Chart 2: Trend Line of Students per Time Slot
      if (timeSlotChartData.length > 0) {
        newCharts.push({
          id: 'tutoring-timeslot-line',
          type: 'line',
          title: 'Student Attendance Trend Across Time Slots',
          data: timeSlotChartData,
          dataKey: 'students',
          nameKey: 'timeSlot',
          xLabel: 'Time Slot',
          yLabel: 'Number of Students',
          color: '#50C878'
        });
      }
      
      // Chart 3: Subject/Class Distribution (if available)
      if (columns.includes('Subject/Class')) {
        const subjectCounts = {};
        sheetData.forEach(row => {
          const subject = row['Subject/Class'];
          if (subject) {
            subjectCounts[subject] = (subjectCounts[subject] || 0) + 1;
          }
        });
        
        const subjectData = Object.keys(subjectCounts)
          .map(subject => ({
            name: subject,
            value: subjectCounts[subject]
          }))
          .sort((a, b) => b.value - a.value)
          .slice(0, 10); // Top 10 subjects
        
        if (subjectData.length > 0) {
          newCharts.push({
            id: 'subject-distribution-bar',
            type: 'bar',
            title: 'Top 10 Most Popular Subjects/Classes',
            data: subjectData,
            dataKey: 'value',
            nameKey: 'name',
            xLabel: 'Subject/Class',
            yLabel: 'Number of Sessions',
            color: '#F5A623'
          });
        }
      }
      
      // Chart 4: Tutor workload (if available)
      if (columns.includes('Tutor')) {
        const tutorCounts = {};
        sheetData.forEach(row => {
          const tutor = row['Tutor'];
          if (tutor) {
            tutorCounts[tutor] = (tutorCounts[tutor] || 0) + 1;
          }
        });
        
        const tutorData = Object.keys(tutorCounts)
          .map(tutor => ({
            name: tutor,
            value: tutorCounts[tutor]
          }))
          .sort((a, b) => b.value - a.value);
        
        if (tutorData.length > 0) {
          newCharts.push({
            id: 'tutor-workload-bar',
            type: 'bar',
            title: 'Tutoring Sessions per Tutor',
            data: tutorData,
            dataKey: 'value',
            nameKey: 'name',
            xLabel: 'Tutor',
            yLabel: 'Number of Sessions',
            color: '#E94B3C'
          });
        }
      }
      
      // Chart 5: Daily attendance pattern
      if (columns.includes('Date')) {
        const dateCounts = {};
        sheetData.forEach(row => {
          const date = row['Date'];
          if (date) {
            let dateStr = '';
            if (date instanceof Date) {
              dateStr = date.toLocaleDateString();
            } else if (typeof date === 'string') {
              const dateObj = new Date(date);
              if (!isNaN(dateObj)) {
                dateStr = dateObj.toLocaleDateString();
              }
            }
            if (dateStr) {
              dateCounts[dateStr] = (dateCounts[dateStr] || 0) + 1;
            }
          }
        });
        
        const dateData = Object.keys(dateCounts)
          .sort((a, b) => new Date(a) - new Date(b))
          .map(date => ({
            date: date,
            students: dateCounts[date]
          }));
        
        if (dateData.length > 0) {
          newCharts.push({
            id: 'daily-attendance-line',
            type: 'line',
            title: 'Daily Student Attendance',
            data: dateData,
            dataKey: 'students',
            nameKey: 'date',
            xLabel: 'Date',
            yLabel: 'Number of Students',
            color: '#9B59B6'
          });
        }
      }
      
    } else {
      // General data analysis (original logic)
      const numericCols = getNumericColumns();
      const categoricalCols = columns.filter(col => !numericCols.includes(col));
      
      // Bar charts for each numeric column by categorical column
      numericCols.forEach(numCol => {
        categoricalCols.slice(0, 2).forEach(catCol => {
          const aggregated = {};
          sheetData.forEach(row => {
            const category = row[catCol];
            const value = Number(row[numCol]);
            if (category && !isNaN(value)) {
              if (!aggregated[category]) {
                aggregated[category] = { sum: 0, count: 0 };
              }
              aggregated[category].sum += value;
              aggregated[category].count += 1;
            }
          });

          const chartData = Object.keys(aggregated).slice(0, 10).map(key => ({
            name: String(key),
            value: aggregated[key].sum / aggregated[key].count
          }));

          if (chartData.length > 0) {
            newCharts.push({
              id: `bar-${numCol}-${catCol}`,
              type: 'bar',
              title: `Average ${numCol} by ${catCol}`,
              data: chartData,
              dataKey: 'value',
              nameKey: 'name',
              xLabel: catCol,
              yLabel: `Average ${numCol}`,
              color: '#8884d8'
            });
          }
        });
      });

      // Line charts for numeric trends
      numericCols.slice(0, 3).forEach(numCol => {
        const chartData = sheetData.slice(0, 20).map((row, idx) => ({
          name: String(idx + 1),
          value: Number(row[numCol]) || 0
        }));

        newCharts.push({
          id: `line-${numCol}`,
          type: 'line',
          title: `${numCol} Trend`,
          data: chartData,
          dataKey: 'value',
          nameKey: 'name',
          xLabel: 'Record Number',
          yLabel: numCol,
          color: '#8884d8'
        });
      });

      // Pie charts for categorical distributions
      categoricalCols.slice(0, 2).forEach(catCol => {
        const counts = {};
        sheetData.forEach(row => {
          const val = row[catCol];
          if (val) {
            counts[val] = (counts[val] || 0) + 1;
          }
        });

        const chartData = Object.keys(counts).slice(0, 8).map(key => ({
          name: String(key),
          value: counts[key]
        }));

        if (chartData.length > 1) {
          newCharts.push({
            id: `pie-${catCol}`,
            type: 'pie',
            title: `${catCol} Distribution`,
            data: chartData,
            dataKey: 'value',
            nameKey: 'name'
          });
        }
      });

      // Scatter plots for numeric vs numeric
      if (numericCols.length >= 2) {
        for (let i = 0; i < Math.min(2, numericCols.length - 1); i++) {
          const xCol = numericCols[i];
          const yCol = numericCols[i + 1];
          
          const chartData = sheetData.slice(0, 50).map(row => ({
            x: Number(row[xCol]) || 0,
            y: Number(row[yCol]) || 0
          })).filter(d => d.x !== 0 || d.y !== 0);

          if (chartData.length > 0) {
            newCharts.push({
              id: `scatter-${xCol}-${yCol}`,
              type: 'scatter',
              title: `${yCol} vs ${xCol}`,
              data: chartData,
              xLabel: xCol,
              yLabel: yCol
            });
          }
        }
      }
    }

    setCharts(newCharts);
  };

  // Download chart as PNG
  const downloadChartAsPNG = async (chartId) => {
    const element = document.getElementById(chartId);
    if (!element) return;

    const canvas = await html2canvas(element, {
      backgroundColor: '#ffffff',
      scale: 2
    });
    
    canvas.toBlob((blob) => {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${chartId}.png`;
      a.click();
      URL.revokeObjectURL(url);
    });
  };

  // Download chart as SVG
  const downloadChartAsSVG = (chartId) => {
    const element = document.getElementById(chartId);
    if (!element) return;

    const svgElement = element.querySelector('svg');
    if (!svgElement) return;

    const svgData = new XMLSerializer().serializeToString(svgElement);
    const blob = new Blob([svgData], { type: 'image/svg+xml' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${chartId}.svg`;
    a.click();
    URL.revokeObjectURL(url);
  };

  // Render individual chart
  const renderChart = (chart) => {
    const commonProps = {
      width: 600,
      height: 400,
      data: chart.data,
      margin: { top: 20, right: 30, left: 20, bottom: 80 }
    };

    switch (chart.type) {
      case 'bar':
        return (
          <ResponsiveContainer width="100%" height={400}>
            <BarChart {...commonProps}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis 
                dataKey={chart.nameKey} 
                angle={-45} 
                textAnchor="end" 
                height={120}
                label={{ value: chart.xLabel || '', position: 'insideBottom', offset: -10 }}
              />
              <YAxis label={{ value: chart.yLabel || '', angle: -90, position: 'insideLeft' }} />
              <Tooltip />
              <Legend wrapperStyle={{ paddingTop: '20px' }} />
              <Bar 
                dataKey={chart.dataKey} 
                fill={chart.color || '#8884d8'} 
                name={chart.yLabel || chart.dataKey}
              />
            </BarChart>
          </ResponsiveContainer>
        );

      case 'line':
        return (
          <ResponsiveContainer width="100%" height={400}>
            <LineChart {...commonProps}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis 
                dataKey={chart.nameKey}
                angle={-45}
                textAnchor="end"
                height={120}
                label={{ value: chart.xLabel || '', position: 'insideBottom', offset: -10 }}
              />
              <YAxis label={{ value: chart.yLabel || '', angle: -90, position: 'insideLeft' }} />
              <Tooltip />
              <Legend wrapperStyle={{ paddingTop: '20px' }} />
              <Line 
                type="monotone" 
                dataKey={chart.dataKey} 
                stroke={chart.color || '#8884d8'}
                strokeWidth={2}
                dot={{ fill: chart.color || '#8884d8', r: 4 }}
                name={chart.yLabel || chart.dataKey}
              />
            </LineChart>
          </ResponsiveContainer>
        );

      case 'pie':
        return (
          <ResponsiveContainer width="100%" height={400}>
            <PieChart>
              <Pie
                data={chart.data}
                dataKey={chart.dataKey}
                nameKey={chart.nameKey}
                cx="50%"
                cy="50%"
                outerRadius={120}
                label={({ name, percent }) => `${name}: ${(percent * 100).toFixed(1)}%`}
              >
                {chart.data.map((entry, index) => (
                  <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                ))}
              </Pie>
              <Tooltip />
              <Legend wrapperStyle={{ paddingTop: '20px' }} />
            </PieChart>
          </ResponsiveContainer>
        );

      case 'scatter':
        return (
          <ResponsiveContainer width="100%" height={400}>
            <ScatterChart {...commonProps}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis 
                dataKey="x" 
                name={chart.xLabel}
                label={{ value: chart.xLabel || '', position: 'insideBottom', offset: -10 }}
              />
              <YAxis 
                dataKey="y" 
                name={chart.yLabel}
                label={{ value: chart.yLabel || '', angle: -90, position: 'insideLeft' }}
              />
              <Tooltip cursor={{ strokeDasharray: '3 3' }} />
              <Legend wrapperStyle={{ paddingTop: '20px' }} />
              <Scatter name={chart.title} data={chart.data} fill="#8884d8" />
            </ScatterChart>
          </ResponsiveContainer>
        );

      default:
        return null;
    }
  };

  return (
    <div className="App">
      <header className="header">
        <h1>📊 Excel Data Analyzer</h1>
        <p>Upload your Excel file, select a sheet, and generate interactive charts</p>
      </header>

      <div className="container">
        {/* File Upload */}
        <div className="upload-section">
          <label htmlFor="file-upload" className="upload-button">
            Choose Excel File
          </label>
          <input
            id="file-upload"
            type="file"
            accept=".xlsx,.xls"
            onChange={handleFileUpload}
            style={{ display: 'none' }}
          />
        </div>

        {/* Sheet Selection */}
        {sheetNames.length > 0 && (
          <div className="sheet-selection">
            <h3>Select a Sheet:</h3>
            <div className="sheet-buttons">
              {sheetNames.map(name => (
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

        {/* Generate Charts Button */}
        {selectedSheet && sheetData.length > 0 && (
          <div className="generate-section">
            <button className="generate-button" onClick={generateCharts}>
              🎨 Generate Charts
            </button>
            <p className="info-text">
              Found {sheetData.length} rows and {columns.length} columns
              {isTutoringData && <span className="tutoring-badge"> • Tutoring Data Detected</span>}
            </p>
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

        {/* Empty State */}
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