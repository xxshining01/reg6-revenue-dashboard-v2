import React, { useState, useEffect, useMemo, useRef } from 'react';
import Papa from 'papaparse';
import { 
  BarChart, Bar, AreaChart, Area, PieChart, Pie, Cell, 
  XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, ResponsiveContainer, Legend, ComposedChart, Line
} from 'recharts';
import { Activity, DollarSign, TrendingUp, TrendingDown, RefreshCw, AlertCircle, Filter, Target, MapPin, Layers, ChevronRight, ChevronDown, Sparkles, Sun, Moon, Download } from 'lucide-react';
import html2canvas from 'html2canvas';
import { MapContainer, TileLayer, GeoJSON } from 'react-leaflet';
import 'leaflet/dist/leaflet.css';

const SHEET_CSV_URL = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQtmCb551xMV17lEECaAvPBySZ43zIrHT2jbz84udDmB9cvwiPYUmwogIdxranN_J3fheWXJZLrj2hV/pub?gid=1362457951&single=true&output=csv';

const COLORS = ['#0d9488', '#ec4899', '#f59e0b', '#3b82f6', '#ef4444', '#84cc16', '#8b5cf6'];
const MONTH_NAMES = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];

const PROVINCE_MAP_EN_TH = {
  'Nakhon Sawan': 'นครสวรรค์',  'Uthai Thani': 'อุทัยธานี', 'Kamphaeng Phet': 'กำแพงเพชร',
  'Tak': 'ตาก', 'Sukhothai': 'สุโขทัย', 'Phitsanulok': 'พิษณุโลก',
  'Phichit': 'พิจิตร', 'Phetchabun': 'เพชรบูรณ์'
};

const GaugeChart = ({ title, actual, target, isIncome, theme }) => {
  const pct = target > 0 ? (actual / target) * 100 : 0;
  
  let color = '#3b82f6';
  if (isIncome) {
    if (pct >= 100) color = '#10b981';
    else if (pct >= 80) color = '#facc15';
    else color = '#ef4444';
  } else {
    if (pct <= 100) color = '#10b981';
    else if (pct <= 110) color = '#facc15';
    else color = '#ef4444';
  }

  const displayPct = Math.min(pct, 100);
  const data = [
    { name: 'ผลงาน', value: displayPct },
    { name: 'ส่วนต่าง', value: 100 - displayPct }
  ];

  return (
    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', width: '100%', padding: '1rem', background: 'var(--bg-highlight)', borderRadius: '12px' }}>
      <h4 style={{ fontSize: '0.9rem', fontWeight: '500', color: 'var(--text-secondary)', marginBottom: '0.75rem', textAlign: 'center' }}>{title}</h4>
      <div style={{ width: '100%', height: '140px', position: 'relative' }}>
        <ResponsiveContainer width="100%" height="100%">
          <PieChart>
            <Pie
              data={data}
              cx="50%"
              cy="100%"
              startAngle={180}
              endAngle={0}
              innerRadius={90}
              outerRadius={115}
              dataKey="value"
              stroke="none"
              isAnimationActive={true}
            >
              <Cell fill={color} />
              <Cell fill="var(--glass-border)" />
            </Pie>
          </PieChart>
        </ResponsiveContainer>
        <div style={{ position: 'absolute', bottom: '15px', left: '0', right: '0', textAlign: 'center' }}>
          <div style={{ fontSize: '2rem', fontWeight: '700', color: color, lineHeight: '1' }}>{pct.toFixed(1)}%</div>
        </div>
      </div>
      <div style={{ marginTop: '1rem', fontSize: '0.8rem', color: 'var(--text-secondary)', textAlign: 'center' }}>
        เป้าหมาย: ฿{target.toLocaleString(undefined, { maximumFractionDigits: 0 })}
      </div>
    </div>
  );
};

const Dashboard = () => {
  const dashboardRef = useRef(null);

  const captureAndDownload = async (filename) => {
    if (!dashboardRef.current) return;
    try {
      const canvas = await html2canvas(dashboardRef.current, {
        scale: 2,
        useCORS: true,
        backgroundColor: theme === 'dark' ? '#09090b' : '#f8fafc',
      });
      const image = canvas.toDataURL("image/png", 1.0);
      const link = document.createElement("a");
      link.download = filename;
      link.href = image;
      link.click();
    } catch (err) {
      console.error("Error exporting image:", err);
    }
  };

  const handleConfirmExport = async () => {
    try {
      if (exportMode === 'basic') {
         setIsExportModalOpen(false);
         setIsExporting(true);
         setExportProgress({ current: 1, total: 1 });
         await delay(800);
         await captureAndDownload(`dashboard-export-${selectedYear}-month${selectedMonth}.png`);
      } else if (exportMode === 'advanced') {
         if (selectedExportProvinces.length === 0) {
           alert("กรุณาเลือกอย่างน้อย 1 จังหวัด");
           return;
         }
         setIsExportModalOpen(false);
         setIsExporting(true);
         
         const originalProv = selectedProvince;
         
         for (let i = 0; i < selectedExportProvinces.length; i++) {
            const prov = selectedExportProvinces[i];
            setExportProgress({ current: i + 1, total: selectedExportProvinces.length });
            
            setSelectedProvince(prov);
            await delay(2500); 
            
            await captureAndDownload(`dashboard-${prov}-${selectedYear}-month${selectedMonth}.png`);
         }
         
         setSelectedProvince(originalProv);
      }
    } catch (err) {
      console.error("Export Error:", err);
      alert("เกิดข้อผิดพลาดในการแคปหน้าจอ");
    } finally {
      setIsExporting(false);
    }
  };

  const [rawData, setRawData] = useState([]);
  const [geoData, setGeoData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);

  // Filters & State
  const [activeTab, setActiveTab] = useState('income');
  const [theme, setTheme] = useState('dark');
  const [isExportModalOpen, setIsExportModalOpen] = useState(false);
  const [exportMode, setExportMode] = useState('basic');
  const [selectedExportProvinces, setSelectedExportProvinces] = useState([]);
  const [isExporting, setIsExporting] = useState(false);
  const [exportProgress, setExportProgress] = useState({ current: 0, total: 0 });

  const delay = (ms) => new Promise(res => setTimeout(res, ms));

  useEffect(() => {
    document.body.setAttribute('data-theme', theme);
  }, [theme]);
  const [selectedYear, setSelectedYear] = useState('');
  const [availableYears, setAvailableYears] = useState([]);
  const [selectedMonth, setSelectedMonth] = useState('ทั้งหมด');
  
  const [selectedProvince, setSelectedProvince] = useState('ทั้งหมด');
  const [availableProvinces, setAvailableProvinces] = useState([]);
  const [selectedOffice, setSelectedOffice] = useState('ทั้งหมด');
  const [availableOffices, setAvailableOffices] = useState([]);
  
  // Drill-down States
  const [expandedBGs, setExpandedBGs] = useState({});
  const [expandedProvs, setExpandedProvs] = useState({});

  const cleanNumber = (val) => {
    if (!val) return 0;
    if (typeof val === 'number') return val;
    return parseFloat(val.replace(/,/g, '')) || 0;
  };

  const fetchGeoJSON = async () => {
    try {
      const res = await fetch('/region6.geojson');
      const data = await res.json();
      setGeoData(data);
    } catch (err) {
      console.error("Failed to load geojson", err);
    }
  };

  const fetchData = () => {
    setLoading(true); setError(null);
    Papa.parse(SHEET_CSV_URL, {
      download: true, header: true, skipEmptyLines: true,
      complete: (results) => {
        const parsed = results.data.map(row => ({
          year: parseInt(row['ปี พ.ศ.']),
          month: parseInt(row['Month']),
          province: row['จังหวัด'],
          office: row['ที่ทำการ'],
          category: row['หมวดหมู่'],
          businessGroup: row['Business Group'],
          evmService: row['EVM Service'],
          actual: cleanNumber(row['ผลงานปีนี้']),
          target: cleanNumber(row['เป้าหมาย']),
          prevActual: cleanNumber(row['ผลงานปีก่อน'])
        })).filter(r => !isNaN(r.year));
        
        setRawData(parsed);
        const years = [...new Set(parsed.map(r => r.year))].sort((a,b) => b - a);
        setAvailableYears(years);
        if (years.length > 0) setSelectedYear(years[0]);

        const targetProvincesTh = Object.values(PROVINCE_MAP_EN_TH);
        const provinces = [...new Set(parsed.map(r => r.province))].filter(p => targetProvincesTh.includes(p)).sort();
        setAvailableProvinces(provinces);

        setLoading(false);
      },
      error: () => { setLoading(false); setError("ไม่สามารถดึงข้อมูลจาก Google Sheets ได้"); }
    });
  };

  useEffect(() => { fetchData(); fetchGeoJSON(); }, []);

  useEffect(() => {
    if (!rawData.length) return;
    if (selectedProvince === 'ทั้งหมด') {
      const targetProvincesTh = Object.values(PROVINCE_MAP_EN_TH);
      const offices = [...new Set(rawData.filter(r => targetProvincesTh.includes(r.province)).map(r => r.office))].filter(Boolean).sort();
      setAvailableOffices(offices);
    } else {
      const offices = [...new Set(rawData.filter(r => r.province === selectedProvince).map(r => r.office))].filter(Boolean).sort();
      setAvailableOffices(offices);
    }
    setSelectedOffice('ทั้งหมด'); 
  }, [selectedProvince, rawData]);

  const processed = useMemo(() => {
    if (!rawData.length || !selectedYear) return null;

    let filtered = rawData.filter(r => r.year === selectedYear);
    const targetProvincesTh = Object.values(PROVINCE_MAP_EN_TH);
    filtered = filtered.filter(r => targetProvincesTh.includes(r.province));

    if (selectedMonth !== 'ทั้งหมด') filtered = filtered.filter(r => r.month === parseInt(selectedMonth));
    if (selectedProvince !== 'ทั้งหมด') filtered = filtered.filter(r => r.province === selectedProvince);
    if (selectedOffice !== 'ทั้งหมด') filtered = filtered.filter(r => r.office === selectedOffice);

    const targetCategory = activeTab === 'income' ? 'รายได้' : 'ค่าใช้จ่าย';
    const tabFiltered = filtered.filter(r => r.category === targetCategory);

    let tabActual = 0, tabTarget = 0, tabPrev = 0;
    const monthlyMap = Array.from({ length: 12 }, (_, i) => ({
      name: MONTH_NAMES[i], actual: 0, target: 0, prev: 0
    }));

    const bgMap = {};
    const bgToEvmMap = {};
    const locationMap = {};

    tabFiltered.forEach(row => {
      tabActual += row.actual; tabTarget += row.target; tabPrev += row.prevActual;
      
      if (row.month >= 1 && row.month <= 12) {
        monthlyMap[row.month - 1].actual += row.actual;
        monthlyMap[row.month - 1].target += row.target;
        monthlyMap[row.month - 1].prev += row.prevActual;
      }

      const bg = row.businessGroup || 'อื่นๆ';
      if (!bgMap[bg]) bgMap[bg] = { name: bg.substring(0,35), rawName: bg, actual: 0, target: 0, prev: 0 };
      bgMap[bg].actual += row.actual;
      bgMap[bg].target += row.target;
      bgMap[bg].prev += row.prevActual;

      const evm = row.evmService || 'ไม่ระบุ EVM';
      if (!bgToEvmMap[bg]) bgToEvmMap[bg] = {};
      if (!bgToEvmMap[bg][evm]) bgToEvmMap[bg][evm] = { name: evm, actual: 0, target: 0, prev: 0 };
      bgToEvmMap[bg][evm].actual += row.actual;
      bgToEvmMap[bg][evm].target += row.target;
      bgToEvmMap[bg][evm].prev += row.prevActual;

      const provKey = row.province;
      const officeKey = row.office || 'ไม่ระบุที่ทำการ';
      if (provKey) {
         if (!locationMap[provKey]) locationMap[provKey] = { name: provKey, actual: 0, target: 0, offices: {} };
         locationMap[provKey].actual += row.actual;
         locationMap[provKey].target += row.target;
         
         if (!locationMap[provKey].offices[officeKey]) locationMap[provKey].offices[officeKey] = { name: officeKey, actual: 0, target: 0 };
         locationMap[provKey].offices[officeKey].actual += row.actual;
         locationMap[provKey].offices[officeKey].target += row.target;
      }
    });

    const bgDataList = Object.values(bgMap).sort((a,b) => b.actual - a.actual);
    
    // Sort EVMs within BG
    const hierarchicalData = bgDataList.map(bg => {
       const evms = bgToEvmMap[bg.rawName] ? Object.values(bgToEvmMap[bg.rawName]).sort((a,b) => b.actual - a.actual) : [];
       return { ...bg, evms };
    });

    const hierarchicalLocationData = Object.values(locationMap)
      .sort((a,b) => b.actual - a.actual)
      .map(prov => ({
         ...prov,
         offices: Object.values(prov.offices).sort((a,b) => b.actual - a.actual)
      }));

    return {
      totals: { actual: tabActual, target: tabTarget, prev: tabPrev },
      monthlyData: monthlyMap,
      hierarchicalData,
      hierarchicalLocationData,
      provinceAgg: locationMap
    };

  }, [rawData, selectedYear, selectedMonth, selectedProvince, selectedOffice, activeTab]);

  const formatAmt = (val) => `฿${(val/1000000).toFixed(1)}M`;
  const formatFullAmt = (val) => `฿${val.toLocaleString(undefined, { maximumFractionDigits: 0 })}`;

  const isIncome = activeTab === 'income';
  const themeColor = isIncome ? '#2dd4bf' : '#fb7185';
  
  
  const generateAIInsight = () => {
    if (!processed || !processed.totals) return 'กำลังวิเคราะห์ข้อมูล...';
    const pct = processed.totals.target > 0 ? (processed.totals.actual / processed.totals.target) * 100 : 0;
    
    const typeStr = isIncome ? "รายได้" : "ค่าใช้จ่าย";
    const statusStr = isIncome 
      ? (pct >= 100 ? "ทะลุเป้าหมายได้อย่างยอดเยี่ยม" : (pct >= 80 ? "อยู่ในเกณฑ์ที่ดีแต่ยังต้องผลักดันอีกเล็กน้อย" : "ยังต่ำกว่าเป้าหมายที่ควรจะเป็นพอสมควร")) 
      : (pct <= 100 ? "สามารถควบคุมได้ดีเยี่ยม" : (pct <= 110 ? "เริ่มสูงกว่าเป้าหมาย ต้องเฝ้าระวัง" : "เกินเป้าหมายที่ตั้งไว้ค่อนข้างมาก ให้ตรวจสอบเพื่อคุมรายจ่ายด่วน"));
      
    let topProv = "ไม่มีข้อมูล";
    const provs = Object.values(processed.provinceAgg).filter(p => p.target > 0);
    if (provs.length > 0) {
      if (isIncome) provs.sort((a,b) => (b.actual/b.target) - (a.actual/a.target));
      else provs.sort((a,b) => (a.actual/a.target) - (b.actual/b.target));
      topProv = provs[0].name;
    }

    let topBG = "ไม่มีข้อมูล";
    if (processed.hierarchicalData.length > 0) {
      topBG = processed.hierarchicalData[0].name;
    }

    const actStr = `฿${(processed.totals.actual / 1000000).toFixed(1)}M`;
    return `ในภาพรวม ${typeStr}สะสมอยู่ที่ ${actStr} คิดเป็น ${pct.toFixed(1)}% ของเป้าหมาย ถือว่า${statusStr} โดยมีกลุ่มธุรกิจหลักที่ขับเคลื่อนคือ "${topBG}" และจังหวัดที่มีประสิทธิภาพเมื่อเปรียบเทียบกับเป้าหมายได้ดีที่สุดคือ "จังหวัด${topProv}"`;
  };

  // Toggle expansion blocks
  const toggleBG = (bgName) => setExpandedBGs(prev => ({ ...prev, [bgName]: !prev[bgName] }));
  const toggleProv = (provName) => setExpandedProvs(prev => ({ ...prev, [provName]: !prev[provName] }));

  const getProvinceStyle = (feature) => {
    if (!processed) return { fillColor: theme === 'light' ? '#e2e8f0' : '#333', weight: 1, opacity: 1, color: '#000', fillOpacity: 0.7 };
    const thName = PROVINCE_MAP_EN_TH[feature.properties.NAME_1];
    const data = processed.provinceAgg[thName];

    if (!data || data.target === 0) return { fillColor: theme === 'light' ? '#e2e8f0' : '#333', weight: 1, opacity: 1, color: '#000', fillOpacity: 0.7 };

    const pct = (data.actual / data.target) * 100;
    
    let fillColor = '#333';
    const isGood = isIncome ? pct >= 100 : pct <= 100;
    const isWarning = isIncome ? (pct >= 80 && pct < 100) : (pct > 100 && pct <= 110);

    if (isGood) fillColor = '#10b981';
    else if (isWarning) fillColor = '#facc15';
    else fillColor = '#ef4444';

    const isTargetProv = selectedProvince === 'ทั้งหมด' || selectedProvince === thName;
    return { fillColor, weight: 1, opacity: 1, color: theme === 'light' ? 'rgba(0,0,0,0.1)' : 'rgba(255,255,255,0.2)', fillOpacity: isTargetProv ? 0.8 : 0.2 };
  };

  const onEachFeature = (feature, layer) => {
    const thName = PROVINCE_MAP_EN_TH[feature.properties.NAME_1];
    const data = processed?.provinceAgg[thName];

    let popupContent = `<b>${thName}</b><br/>ไม่พบข้อมูล`;
    if (data) {
      const pct = data.target > 0 ? ((data.actual / data.target) * 100).toFixed(1) : 0;
      popupContent = `
        <div style="font-family: Outfit, sans-serif;">
          <b style="font-size: 14px;">จังหวัด${thName}</b><br/>
          ผลงาน: ฿${data.actual.toLocaleString(undefined, { maximumFractionDigits: 0 })}<br/>
          เป้าหมาย: ฿${data.target.toLocaleString(undefined, { maximumFractionDigits: 0 })}<br/>
          ความสำเร็จ: <b>${pct}%</b>
        </div>
      `;
    }
    layer.bindTooltip(popupContent, { sticky: true });
  };

  const getPerfColor = (actual, target) => {
     if(target === 0) return 'var(--text-secondary)';
     const pct = (actual / target) * 100;
     const isGood = isIncome ? pct >= 100 : pct <= 100;
     const isWarn = isIncome ? (pct >= 80 && pct < 100) : (pct > 100 && pct <= 110);
     if(isGood) return '#10b981';
     if(isWarn) return '#facc15';
     return '#ef4444';
  };

  const MiniBar = ({ pct }) => (
    <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', flex: 1, justifyContent: 'flex-end' }}>
      <span style={{ fontSize: '0.75rem', color: 'var(--text-secondary)', display: 'inline-block', width: '35px', textAlign: 'right' }}>
         {pct.toFixed(1)}%
      </span>
      <div style={{ width: '40px', height: '6px', background: 'var(--line-color)', borderRadius: '3px', overflow: 'hidden' }}>
        <div style={{ width: `${Math.min(pct, 100)}%`, height: '100%', background: themeColor, borderRadius: '3px' }}></div>
      </div>
    </div>
  );

  return (
    <div ref={dashboardRef} style={{ display: 'flex', flexDirection: 'column', minHeight: '100vh', gap: '1rem', paddingBottom: '3rem' }}>
      
      {/* App Header w/ Tabs */}
      <div className="glass-panel" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '1rem 1.5rem', borderRadius: '16px' }}>
        <h1 style={{ fontWeight: '700', fontSize: '1.5rem', margin: 0, color: 'var(--text-primary)' }}>Dashboard แสดงผลรายได้/ค่าใช้จ่ายของที่ทำการในสังกัด ปข.6</h1>
        
        <div style={{ display: 'flex', gap: '0.5rem' }}>
          <div style={{ display: 'flex', background: 'var(--bg-panel-tertiary)', borderRadius: '12px', padding: '0.25rem' }}>
            <button 
              onClick={() => setActiveTab('income')} 
              style={{ 
                padding: '0.5rem 1.5rem', borderRadius: '8px', border: 'none', cursor: 'pointer', fontWeight: '600',
                background: isIncome ? '#0d9488' : 'transparent', color: isIncome ? '#fff' : 'var(--text-secondary)', transition: 'all 0.3s'
              }}
            >
              รายได้
            </button>
            <button 
              onClick={() => setActiveTab('expense')} 
              style={{ 
                padding: '0.5rem 1.5rem', borderRadius: '8px', border: 'none', cursor: 'pointer', fontWeight: '600',
                background: !isIncome ? '#e11d48' : 'transparent', color: !isIncome ? '#fff' : 'var(--text-secondary)', transition: 'all 0.3s'
              }}
            >
              ค่าใช้จ่าย
            </button>
          </div>
          
          <button 
            onClick={() => setTheme(theme === 'dark' ? 'light' : 'dark')}
            style={{ 
              display: 'flex', alignItems: 'center', justifyContent: 'center',
              width: '42px', height: '42px', borderRadius: '12px', border: '1px solid var(--glass-border)',
              background: 'var(--bg-panel-tertiary)', color: 'var(--text-primary)', cursor: 'pointer',
              transition: 'all 0.3s'
            }}
            title={theme === 'dark' ? 'เปิดโหมดสว่าง' : 'เปิดโหมดมืด'}
          >
            {theme === 'dark' ? <Sun size={20} /> : <Moon size={20} />}
          </button>
          <button 
            onClick={() => setIsExportModalOpen(true)}
            style={{ 
              display: 'flex', alignItems: 'center', justifyContent: 'center',
              width: '42px', height: '42px', borderRadius: '12px', border: '1px solid var(--glass-border)',
              background: 'var(--bg-panel-tertiary)', color: 'var(--text-primary)', cursor: 'pointer',
              transition: 'all 0.3s'
            }}
            title="บันทึกภาพหน้าจอ"
          >
            <Download size={20} />
          </button>
        </div>
      </div>

      {error && (
        <div className="glass-panel" style={{ background: 'rgba(239, 68, 68, 0.1)', color: 'var(--danger)' }}>
          <AlertCircle size={24} style={{ display: 'inline', marginRight: '0.5rem', verticalAlign: 'middle' }} /> {error}
        </div>
      )}

      {loading && !processed ? (
        <div className="flex-center" style={{ height: '50vh', width: '100%' }}>
           <div className="text-gradient" style={{ fontSize: '2rem', fontWeight: 'bold', animation: 'pulse 1.5s infinite' }}>กำลังโหลดข้อมูล...</div>
        </div>
      
      ) : processed && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '1rem' }}>
          
          {/* AI Daily Insights */}
          <div className="glass-panel" style={{ 
            background: 'linear-gradient(90deg, rgba(59, 130, 246, 0.15) 0%, rgba(139, 92, 246, 0.15) 100%)', 
            border: '1px solid rgba(139, 92, 246, 0.3)', 
            padding: '1.25rem 1.5rem', 
            borderRadius: '16px',
            display: 'flex', 
            gap: '1.25rem', 
            alignItems: 'center' 
          }}>
            <div style={{ background: 'rgba(139, 92, 246, 0.25)', padding: '0.8rem', borderRadius: '12px', color: theme === 'dark' ? '#c4b5fd' : '#8b5cf6', flexShrink: 0 }}>
               <Sparkles size={28} />
            </div>
            <div style={{ flex: 1 }}>
               <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem', marginBottom: '0.5rem' }}>
                  <h3 style={{ fontSize: '1.15rem', fontWeight: '700', color: 'var(--text-primary)', margin: 0 }}>AI Daily Insights</h3>
                  <span style={{ fontSize: '0.75rem', background: 'rgba(255,255,255,0.15)', color: 'var(--text-secondary)', padding: '2px 8px', borderRadius: '12px', fontWeight: '500' }}>บทวิเคราะห์โดย AI</span>
               </div>
               <p style={{ margin: 0, color: 'var(--text-primary)', lineHeight: '1.6', fontSize: '0.95rem' }}>
                 {generateAIInsight()}
               </p>
            </div>
          </div>

          <div style={{ display: 'flex', gap: '1rem', alignItems: 'flex-start' }}>
          
          {/* LEFT SIDEBAR (Fixed Width) */}
          <div style={{ flex: '0 0 320px', display: 'flex', flexDirection: 'column', gap: '1rem' }}>
             
             {/* Totals Block */}
            <div style={{ 
              background: `linear-gradient(180deg, ${isIncome ? 'rgba(13,148,136,0.15)' : 'rgba(225,29,72,0.15)'} 0%, ${theme === 'dark' ? 'rgba(9,9,11,0.5)' : 'rgba(255,255,255,0.4)'} 100%)`,
              border: '1px solid var(--glass-border)', borderRadius: '20px', padding: '1.5rem', display: 'flex', flexDirection: 'column', gap: '1rem', boxShadow: 'var(--glass-shadow)'
            }}>
              <div>
                <h2 style={{ fontSize: '1.5rem', fontWeight: '700', color: themeColor, marginBottom: '0.25rem' }}>ภาพรวม{isIncome ? 'รายได้' : 'ค่าใช้จ่าย'}</h2>
                <p style={{ color: 'var(--text-secondary)', fontSize: '0.875rem' }}>ปี {selectedYear} {selectedMonth !== 'ทั้งหมด' ? `เดือน ${MONTH_NAMES[parseInt(selectedMonth)-1]}` : ''}</p>
              </div>

              <div style={{ background: `rgba(${isIncome ? '45,212,191' : '244,63,94'},0.15)`, padding: '1.25rem', borderRadius: '16px', border: `1px solid ${themeColor}40` }}>
                <div style={{ color: themeColor, fontSize: '0.875rem', marginBottom: '0.25rem', fontWeight: '500' }}>รวมยอดสะสม</div>
                <div style={{ fontSize: '2.5rem', fontWeight: '700', color: 'var(--text-primary)' }}>{formatAmt(processed.totals.actual)}</div>
              </div>

              <div style={{ display: 'flex', flexDirection: 'column', gap: '1rem', marginTop: '1rem' }}>
                  <GaugeChart title="เทียบเป้าหมายปีนี้" actual={processed.totals.actual} target={processed.totals.target} isIncome={isIncome} theme={theme} />
                  <GaugeChart title="เทียบส่วนปีก่อนหน้า" actual={processed.totals.actual} target={processed.totals.prev} isIncome={isIncome} theme={theme} />
              </div>
            </div>

             {/* Top Provinces Map */}
             <div className="glass-panel" style={{ padding: '1rem', display: 'flex', flexDirection: 'column', height: '400px' }}>
                <h3 style={{ fontSize: '1rem', fontWeight: '600', marginBottom: '1rem', color: 'var(--text-primary)' }}>แผนที่ผลงานจังหวัด</h3>
                <div style={{ flex: 1, borderRadius: '12px', overflow: 'hidden', background: 'var(--bg-panel)' }}>
                  <MapContainer center={[16.2, 99.8]} zoom={6.5} style={{ height: '100%', width: '100%' }} zoomControl={false} scrollWheelZoom={false}>
                    <TileLayer url={`https://{s}.basemaps.cartocdn.com/${theme === 'dark' ? 'dark_nolabels' : 'light_nolabels'}/{z}/{x}/{y}{r}.png`} />
                    {geoData && <GeoJSON key={activeTab + selectedYear + selectedMonth} data={geoData} style={getProvinceStyle} onEachFeature={onEachFeature} />}
                  </MapContainer>
                </div>
             </div>
          </div>

          {/* MAIN CONTENT AREA */}
          <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: '1rem', minWidth: '600px' }}>
            
            {/* Top Filter Bar */}
            <div className="glass-panel" style={{ padding: '0.75rem 1.25rem', display: 'flex', gap: '1.25rem', alignItems: 'center', flexWrap: 'wrap', borderRadius: '16px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                <span style={{ color: 'var(--text-secondary)', fontSize: '0.875rem' }}>ปี พ.ศ.</span>
                <select value={selectedYear} onChange={e => setSelectedYear(parseInt(e.target.value))} className="filter-select">
                  {availableYears.map(y => <option key={y} value={y}>{y}</option>)}
                </select>
              </div>

              <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                <span style={{ color: 'var(--text-secondary)', fontSize: '0.875rem' }}>เดือน</span>
                <select value={selectedMonth} onChange={e => setSelectedMonth(e.target.value)} className="filter-select">
                  <option value="ทั้งหมด">ทุกเดือน</option>
                  {MONTH_NAMES.map((m, i) => <option key={i} value={i+1}>{m}</option>)}
                </select>
              </div>

              <div style={{ width: '1px', height: '20px', background: 'var(--line-color)' }}></div>

              <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                <span style={{ color: 'var(--text-secondary)', fontSize: '0.875rem' }}>พื้นที่</span>
                <select value={selectedProvince} onChange={e => setSelectedProvince(e.target.value)} className="filter-select">
                  <option value="ทั้งหมด">ทุกจังหวัด (8 จังหวัด)</option>
                  {availableProvinces.map(p => <option key={p} value={p}>{p}</option>)}
                </select>
                <select value={selectedOffice} onChange={e => setSelectedOffice(e.target.value)} className="filter-select" style={{maxWidth: '180px'}}>
                  <option value="ทั้งหมด">ทุกที่ทำการ</option>
                  {availableOffices.map(o => <option key={o} value={o}>{o}</option>)}
                </select>
              </div>

              <button className="glass-button" onClick={fetchData} disabled={loading} style={{ marginLeft: 'auto', padding: '0.5rem 1rem' }}>
                <RefreshCw size={16} className={loading ? 'animate-spin' : ''} style={{ animation: loading ? 'spin 1s linear infinite' : 'none' }} />
                <span>รีเฟรชข้อมูล</span>
              </button>
            </div>

            {/* Monthly Trend Chart */}
            <div className="glass-panel" style={{ padding: '1.25rem', height: '320px', display: 'flex', flexDirection: 'column' }}>
                <h3 style={{ fontSize: '1rem', fontWeight: '600', marginBottom: '1rem', color: 'var(--text-primary)' }}>
                  แนวโน้มผลงานรายเดือน (ม.ค. - ธ.ค.)
                </h3>
                <div style={{ flex: 1, minHeight: 0 }}>
                  <ResponsiveContainer width="100%" height="100%">
                    <ComposedChart data={processed.monthlyData} margin={{ top: 10, right: 10, left: 0, bottom: 0 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke={theme === 'dark' ? 'rgba(255,255,255,0.1)' : 'rgba(0,0,0,0.1)'} vertical={false} />
                      <XAxis dataKey="name" stroke={theme === 'dark' ? '#a1a1aa' : '#475569'} tick={{ fill: theme === 'dark' ? '#a1a1aa' : '#475569', fontSize: 11 }} axisLine={false} tickLine={false} />
                      <YAxis stroke={theme === 'dark' ? '#a1a1aa' : '#475569'} tick={{ fill: theme === 'dark' ? '#a1a1aa' : '#475569', fontSize: 11 }} axisLine={false} tickLine={false} tickFormatter={formatAmt} />
                      <RechartsTooltip formatter={(v) => formatFullAmt(v)} contentStyle={{ backgroundColor: 'var(--tooltip-bg)', border: 'none', borderRadius: '8px' }} />
                      <Legend wrapperStyle={{ fontSize: '11px', paddingTop: '10px' }} />
                      <Bar dataKey="actual" name="ผลงาน" fill={themeColor} radius={[4, 4, 0, 0]} barSize={35} />
                      <Line type="monotone" dataKey="target" name="เป้าหมาย" stroke="#fcd34d" strokeWidth={3} dot={{ r: 4 }} />
                      <Line type="monotone" dataKey="prev" name="ปีก่อน" stroke="#6b7280" strokeWidth={2} strokeDasharray="5 5" dot={false} />
                    </ComposedChart>
                  </ResponsiveContainer>
                </div>
            </div>

            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1rem' }}>
                
                {/* Hierarchical Breakdown (Business Group -> EVM Service) */}
                <div className="glass-panel" style={{ padding: '1.5rem', display: 'flex', flexDirection: 'column', flex: 1 }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1rem' }}>
                       <h3 style={{ fontSize: '1.1rem', fontWeight: '600', color: themeColor }}>เจาะลึกกลุ่มธุรกิจ (Business Group)</h3>
                    </div>

                    <div style={{ display: 'flex', gap: '1rem', padding: '0.75rem 1rem', borderBottom: '1px solid var(--line-color)', fontWeight: '600', color: 'var(--text-secondary)', fontSize: '0.8rem' }}>
                        <div style={{ flex: 2.5 }}>กลุ่มธุรกิจ / บริการ</div>
                        <div style={{ flex: 1, textAlign: 'right' }}>ผลงานจริง</div>
                        <div style={{ flex: 1, textAlign: 'right' }}>% สำเร็จ</div>
                    </div>

                    <div style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem', marginTop: '0.5rem' }}>
                        {processed.hierarchicalData.map((bg, idx) => {
                           const isExpanded = !!expandedBGs[bg.name];
                           const pct = bg.target > 0 ? (bg.actual / bg.target) * 100 : 0;
                           
                           return (
                             <div key={idx} style={{ background: 'var(--bg-highlight)', borderRadius: '8px', overflow: 'hidden' }}>
                               <div onClick={() => toggleBG(bg.name)} style={{ display: 'flex', gap: '1rem', padding: '1rem', cursor: 'pointer', alignItems: 'center' }} className="hover:bg-white/5 transition-colors">
                                  <div style={{ flex: 2.5, display: 'flex', alignItems: 'center', gap: '0.5rem', fontWeight: '600', color: 'var(--text-primary)', fontSize: '0.95rem' }}>
                                     {isExpanded ? <ChevronDown size={18} color={themeColor} /> : <ChevronRight size={18} color="var(--text-secondary)" />}
                                     <span title={bg.name} style={{ whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', maxWidth: '180px' }}>{bg.name}</span>
                                  </div>
                                  <div style={{ flex: 1, textAlign: 'right', fontWeight: 'bold' }}>{formatAmt(bg.actual)}</div>
                                  <div style={{ flex: 1, textAlign: 'right', fontWeight: 'bold', color: getPerfColor(bg.actual, bg.target) }}>{pct.toFixed(0)}%</div>
                               </div>

                               {isExpanded && (
                                 <div style={{ background: 'var(--bg-panel-secondary)', padding: '0.5rem 0' }}>
                                   {bg.evms.map((evm, eidx) => {
                                      let evmPct = evm.target > 0 ? (evm.actual / evm.target) * 100 : 0;
                                      return (
                                         <div key={eidx} style={{ display: 'flex', gap: '0.5rem', padding: '0.5rem 1rem 0.5rem 2.5rem', fontSize: '0.85rem', alignItems: 'center', borderBottom: eidx !== bg.evms.length - 1 ? '1px solid var(--line-color-faint)' : 'none' }}>
                                            <div style={{ flex: 2.5, color: 'var(--text-secondary)', display: 'flex', alignItems: 'center' }}>
                                              <span style={{ display: 'inline-block', width: '4px', height: '4px', borderRadius: '50%', background: 'var(--text-secondary)', marginRight: '0.5rem' }}></span>
                                              <span title={evm.name} style={{ whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', maxWidth: '160px' }}>{evm.name}</span>
                                            </div>
                                            <div style={{ flex: 1, textAlign: 'right', color: 'var(--text-secondary)' }}>{formatFullAmt(evm.actual)}</div>
                                            <div style={{ flex: 1, textAlign: 'right', color: getPerfColor(evm.actual, evm.target) }}>{evmPct.toFixed(0)}%</div>
                                         </div>
                                      )
                                   })}
                                 </div>
                               )}
                             </div>
                           );
                        })}
                    </div>
                </div>

                {/* Hierarchical Breakdown (Province -> Office) with Proportions*/}
                <div className="glass-panel" style={{ padding: '1.5rem', display: 'flex', flexDirection: 'column', flex: 1 }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1rem' }}>
                       <h3 style={{ fontSize: '1.1rem', fontWeight: '600', color: themeColor }}>เจาะลึกจังหวัด (Provinces)</h3>
                    </div>

                    <div style={{ display: 'flex', gap: '0.5rem', padding: '0.75rem 1rem', borderBottom: '1px solid var(--line-color)', fontWeight: '600', color: 'var(--text-secondary)', fontSize: '0.8rem' }}>
                        <div style={{ flex: 2.2 }}>จังหวัด / ที่ทำการ</div>
                        <div style={{ flex: 0.8, textAlign: 'right' }}>ผลงานจริง</div>
                        <div style={{ flex: 0.8, textAlign: 'right' }}>% สำเร็จ</div>
                        <div style={{ flex: 1, textAlign: 'right' }}>สัดส่วนรวม</div>
                    </div>

                    <div style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem', marginTop: '0.5rem' }}>
                        {processed.hierarchicalLocationData.map((prov, idx) => {
                           const isExpanded = !!expandedProvs[prov.name];
                           const pct = prov.target > 0 ? (prov.actual / prov.target) * 100 : 0;
                           const provProp = processed.totals.actual > 0 ? (prov.actual / processed.totals.actual) * 100 : 0;
                           
                           return (
                             <div key={idx} style={{ background: 'var(--bg-highlight)', borderRadius: '8px', overflow: 'hidden' }}>
                               <div onClick={() => toggleProv(prov.name)} style={{ display: 'flex', gap: '0.5rem', padding: '1rem 0.5rem', cursor: 'pointer', alignItems: 'center' }} className="hover:bg-white/5 transition-colors">
                                  <div style={{ flex: 2.2, display: 'flex', alignItems: 'center', gap: '0.5rem', fontWeight: '600', color: 'var(--text-primary)', fontSize: '0.95rem' }}>
                                     {isExpanded ? <ChevronDown size={18} color={themeColor} /> : <ChevronRight size={18} color="var(--text-secondary)" />}
                                     <span title={prov.name} style={{ whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', maxWidth: '140px' }}>{prov.name}</span>
                                  </div>
                                  <div style={{ flex: 0.8, textAlign: 'right', fontWeight: 'bold', fontSize: '0.9rem' }}>{formatAmt(prov.actual)}</div>
                                  <div style={{ flex: 0.8, textAlign: 'right', fontWeight: 'bold', fontSize: '0.9rem', color: getPerfColor(prov.actual, prov.target) }}>{pct.toFixed(0)}%</div>
                                  <MiniBar pct={provProp} />
                               </div>

                               {isExpanded && (
                                 <div style={{ background: 'var(--bg-panel-secondary)', padding: '0.5rem 0' }}>
                                   {prov.offices.map((office, oidx) => {
                                      let officePct = office.target > 0 ? (office.actual / office.target) * 100 : 0;
                                      let officeProp = processed.totals.actual > 0 ? (office.actual / processed.totals.actual) * 100 : 0;
                                      return (
                                         <div key={oidx} style={{ display: 'flex', gap: '0.5rem', padding: '0.5rem 0.5rem 0.5rem 2.5rem', fontSize: '0.85rem', alignItems: 'center', borderBottom: oidx !== prov.offices.length - 1 ? '1px solid var(--line-color-faint)' : 'none' }}>
                                            <div style={{ flex: 2.2, color: 'var(--text-secondary)', display: 'flex', alignItems: 'center' }}>
                                              <span style={{ display: 'inline-block', width: '4px', height: '4px', borderRadius: '50%', background: 'var(--text-secondary)', marginRight: '0.5rem' }}></span>
                                              <span title={office.name} style={{ whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', maxWidth: '120px' }}>{office.name}</span>
                                            </div>
                                            <div style={{ flex: 0.8, textAlign: 'right', color: 'var(--text-secondary)' }}>{formatFullAmt(office.actual)}</div>
                                            <div style={{ flex: 0.8, textAlign: 'right', color: getPerfColor(office.actual, office.target) }}>{officePct.toFixed(0)}%</div>
                                            <MiniBar pct={officeProp} />
                                         </div>
                                      )
                                   })}
                                 </div>
                               )}
                             </div>
                           );
                        })}
                    </div>
                </div>

            </div>

          </div>
        </div>
        </div>
      )}

            {/* Export Modal */}
      {isExportModalOpen && (
        <div data-html2canvas-ignore="true" style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, background: 'rgba(0,0,0,0.6)', backdropFilter: 'blur(4px)', zIndex: 9999, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
          <div className="glass-panel" style={{ width: '450px', maxWidth: '90vw', background: theme === 'dark' ? '#18181b' : '#ffffff', color: 'var(--text-primary)', border: '1px solid var(--glass-border)' }}>
            <h2 style={{ fontSize: '1.25rem', fontWeight: '700', marginBottom: '1rem' }}>ตัวเลือกการบันทึกภาพหน้าจอ</h2>
            
            <div style={{ display: 'flex', flexDirection: 'column', gap: '1rem', marginBottom: '1.5rem' }}>
               <label style={{ display: 'flex', alignItems: 'center', gap: '0.75rem', cursor: 'pointer', padding: '0.75rem', borderRadius: '8px', background: exportMode === 'basic' ? 'var(--bg-highlight-hover)' : 'transparent', border: '1px solid', borderColor: exportMode === 'basic' ? themeColor : 'var(--glass-border)' }}>
                 <input type="radio" name="exportMode" value="basic" checked={exportMode === 'basic'} onChange={() => setExportMode('basic')} style={{ accentColor: themeColor, width: '18px', height: '18px' }} />
                 <div>
                   <div style={{ fontWeight: '600' }}>โหมดพื้นฐาน (Basic)</div>
                   <div style={{ fontSize: '0.85rem', color: 'var(--text-secondary)' }}>แคปภาพหน้าจอในสถานะปัจจุบัน 1 ภาพ</div>
                 </div>
               </label>
               
               <label style={{ display: 'flex', alignItems: 'center', gap: '0.75rem', cursor: 'pointer', padding: '0.75rem', borderRadius: '8px', background: exportMode === 'advanced' ? 'var(--bg-highlight-hover)' : 'transparent', border: '1px solid', borderColor: exportMode === 'advanced' ? themeColor : 'var(--glass-border)' }}>
                 <input type="radio" name="exportMode" value="advanced" checked={exportMode === 'advanced'} onChange={() => setExportMode('advanced')} style={{ accentColor: themeColor, width: '18px', height: '18px' }} />
                 <div>
                   <div style={{ fontWeight: '600' }}>โหมดขั้นสูง (Advanced Batch Export)</div>
                   <div style={{ fontSize: '0.85rem', color: 'var(--text-secondary)' }}>แคปภาพแยกรายจังหวัดแบบอัตโนมัติ</div>
                 </div>
               </label>
               
               {exportMode === 'advanced' && (
                 <div style={{ background: 'var(--bg-panel-secondary)', padding: '1rem', borderRadius: '8px', maxHeight: '200px', overflowY: 'auto' }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '0.5rem' }}>
                      <span style={{ fontWeight: '600', fontSize: '0.9rem' }}>เลือกจังหวัดที่ต้องการ:</span>
                      <div style={{ display: 'flex', gap: '0.5rem' }}>
                        <button onClick={() => setSelectedExportProvinces([...availableProvinces])} style={{ background: 'none', border: 'none', color: themeColor, fontSize: '0.8rem', cursor: 'pointer', textDecoration: 'underline' }}>เลือกทั้งหมด</button>
                        <button onClick={() => setSelectedExportProvinces([])} style={{ background: 'none', border: 'none', color: 'var(--text-secondary)', fontSize: '0.8rem', cursor: 'pointer', textDecoration: 'underline' }}>ล้างค่า</button>
                      </div>
                    </div>
                    
                    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '0.5rem' }}>
                      {availableProvinces.map(p => (
                        <label key={p} style={{ display: 'flex', alignItems: 'center', gap: '0.5rem', fontSize: '0.9rem', cursor: 'pointer', color: 'var(--text-primary)' }}>
                          <input type="checkbox" checked={selectedExportProvinces.includes(p)} onChange={(e) => {
                            if (e.target.checked) setSelectedExportProvinces(prev => [...prev, p]);
                            else setSelectedExportProvinces(prev => prev.filter(x => x !== p));
                          }} style={{ accentColor: themeColor }} />
                          {p}
                        </label>
                      ))}
                    </div>
                 </div>
               )}
            </div>
            
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '0.75rem' }}>
              <button 
                onClick={() => setIsExportModalOpen(false)} 
                style={{ padding: '0.6rem 1.25rem', borderRadius: '8px', border: '1px solid var(--glass-border)', background: 'transparent', color: 'var(--text-primary)', cursor: 'pointer', fontWeight: '500' }}>
                ยกเลิก
              </button>
              <button 
                onClick={handleConfirmExport} 
                className="glass-button" style={{ background: themeColor, borderColor: themeColor, padding: '0.6rem 1.25rem', opacity: isExporting ? 0.7 : 1 }} disabled={isExporting}>
                <Download size={16} /> เริ่มดาวน์โหลด
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Export Loading Overlay */}
      {isExporting && (
        <div data-html2canvas-ignore="true" style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, background: 'rgba(0,0,0,0.8)', backdropFilter: 'blur(8px)', zIndex: 10000, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', color: '#fff' }}>
           <RefreshCw size={48} style={{ animation: 'spin 1.5s linear infinite', marginBottom: '1.5rem', color: themeColor }} />
           <h2 style={{ fontSize: '1.5rem', fontWeight: '700', marginBottom: '0.5rem' }}>กำลังบันทึกภาพหน้าจอ...</h2>
           {exportMode === 'advanced' && (
             <p style={{ fontSize: '1.1rem', color: '#a1a1aa' }}>
               ภาพที่ {exportProgress.current} จากทั้งหมด {exportProgress.total} ภาพ
             </p>
           )}
           <p style={{ fontSize: '0.9rem', color: '#ef4444', marginTop: '1rem' }}>*กรุณาอย่าสลับหน้าจอหรือปิดเบราว์เซอร์ระหว่างดำเนินการ</p>
        </div>
      )}

      <style>{`

        @keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
        @keyframes pulse { 0% { opacity: 0.5; } 50% { opacity: 1; } 100% { opacity: 0.5; } }
        .filter-select { background: var(--select-bg); padding: 0.35rem 0.5rem; border-radius: 8px; color: var(--text-primary); border: 1px solid var(--glass-border); outline: none; font-family: Outfit; font-weight: 500;}
        .hover\\:bg-white\\/5:hover { background-color: var(--line-color-subtle); }
        .transition-colors { transition-property: color, background-color, border-color; transition-timing-function: cubic-bezier(0.4, 0, 0.2, 1); transition-duration: 150ms; }
      `}</style>
    </div>
  );
};

export default Dashboard;
