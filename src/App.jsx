import { useState } from 'react';
import * as XLSX from 'xlsx';
import { 
  BarChart, Bar, LineChart, Line, PieChart, Pie, ScatterChart, Scatter,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, Cell 
} from 'recharts';
import html2canvas from 'html2canvas';
import './App.css';

const COLORS = ['#0088FE', '#00C49F', '#FFBB28', '#FF8042', '#8884D8', '#82CA9D', '#FFC658', '#FF6B6B'];

function App() {
  const [workbook, setWorkbook] = useState(null);
  const [sheetNames, setSheetNames] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState('');
  const [sheetData, setSheetData] = useState([]);
  const [columns, setColumns] = useState([]);
  const [charts, setCharts] = useState([]);

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
    }
  };

  // Detect numeric columns
  const getNumericColumns = () => {
    if (sheetData.length === 0) return [];
    
    return columns.filter(col => {
      const values = sheetData.slice(0, 10).map(row => row[col]);
      const numericCount = values.filter(v => v !== null && !isNaN(Number(v))).length;
      return numericCount > values.length * 0.5;
    });
  };

  // Generate all possible charts
  const generateCharts = () => {
    const numericCols = getNumericColumns();
    const categoricalCols = columns.filter(col => !numericCols.includes(col));
    
    const newCharts = [];

    // 1. Bar charts
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
            title: `${numCol} by ${catCol}`,
            data: chartData,
            dataKey: 'value',
            nameKey: 'name'
          });
        }
      });
    });

    // 2. Line charts
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
        nameKey: 'name'
      });
    });

    // 3. Pie charts
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

    // 4. Scatter plots
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
      margin: { top: 20, right: 30, left: 20, bottom: 60 }
    };

    switch (chart.type) {
      case 'bar':
        return (
          <ResponsiveContainer width="100%" height={400}>
            <BarChart {...commonProps}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey={chart.nameKey} angle={-45} textAnchor="end" height={100} />
              <YAxis />
              <Tooltip />
              <Legend />
              <Bar dataKey={chart.dataKey} fill="#8884d8" />
            </BarChart>
          </ResponsiveContainer>
        );

      case 'line':
        return (
          <ResponsiveContainer width="100%" height={400}>
            <LineChart {...commonProps}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey={chart.nameKey} />
              <YAxis />
              <Tooltip />
              <Legend />
              <Line type="monotone" dataKey={chart.dataKey} stroke="#8884d8" />
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
                label
              >
                {chart.data.map((entry, index) => (
                  <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                ))}
              </Pie>
              <Tooltip />
              <Legend />
            </PieChart>
          </ResponsiveContainer>
        );

      case 'scatter':
        return (
          <ResponsiveContainer width="100%" height={400}>
            <ScatterChart {...commonProps}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="x" name={chart.xLabel} />
              <YAxis dataKey="y" name={chart.yLabel} />
              <Tooltip cursor={{ strokeDasharray: '3 3' }} />
              <Legend />
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
