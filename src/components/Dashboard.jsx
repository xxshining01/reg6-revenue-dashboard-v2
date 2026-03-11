import React, { useState, useEffect, useMemo } from 'react';
import Papa from 'papaparse';
import { 
  BarChart, Bar, AreaChart, Area, PieChart, Pie, Cell, 
  XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, ResponsiveContainer, Legend, ComposedChart, Line
} from 'recharts';
import { Activity, DollarSign, TrendingUp, TrendingDown, RefreshCw, AlertCircle, Filter, Target, MapPin, Layers, ChevronRight, ChevronDown } from 'lucide-react';
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

const GaugeChart = ({ title, actual, target, isIncome }) => {
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
    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', width: '100%', padding: '1rem', background: 'rgba(255,255,255,0.02)', borderRadius: '12px' }}>
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
              <Cell fill="rgba(255,255,255,0.08)" />
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
  const [rawData, setRawData] = useState([]);
  const [geoData, setGeoData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);

  // Filters & State
  const [activeTab, setActiveTab] = useState('income');
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
  
  // Toggle expansion blocks
  const toggleBG = (bgName) => setExpandedBGs(prev => ({ ...prev, [bgName]: !prev[bgName] }));
  const toggleProv = (provName) => setExpandedProvs(prev => ({ ...prev, [provName]: !prev[provName] }));

  const getProvinceStyle = (feature) => {
    if (!processed) return { fillColor: '#333', weight: 1, opacity: 1, color: '#000', fillOpacity: 0.7 };
    const thName = PROVINCE_MAP_EN_TH[feature.properties.NAME_1];
    const data = processed.provinceAgg[thName];

    if (!data || data.target === 0) return { fillColor: '#333', weight: 1, opacity: 1, color: '#000', fillOpacity: 0.7 };

    const pct = (data.actual / data.target) * 100;
    
    let fillColor = '#333';
    const isGood = isIncome ? pct >= 100 : pct <= 100;
    const isWarning = isIncome ? (pct >= 80 && pct < 100) : (pct > 100 && pct <= 110);

    if (isGood) fillColor = '#10b981';
    else if (isWarning) fillColor = '#facc15';
    else fillColor = '#ef4444';

    const isTargetProv = selectedProvince === 'ทั้งหมด' || selectedProvince === thName;
    return { fillColor, weight: 1, opacity: 1, color: 'rgba(255,255,255,0.2)', fillOpacity: isTargetProv ? 0.8 : 0.2 };
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
     if(target === 0) return '#fff';
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
      <div style={{ width: '40px', height: '6px', background: 'rgba(255,255,255,0.1)', borderRadius: '3px', overflow: 'hidden' }}>
        <div style={{ width: `${Math.min(pct, 100)}%`, height: '100%', background: themeColor, borderRadius: '3px' }}></div>
      </div>
    </div>
  );

  return (
    <div style={{ display: 'flex', flexDirection: 'column', minHeight: '100vh', gap: '1rem', paddingBottom: '3rem' }}>
      
      {/* App Header w/ Tabs */}
      <div className="glass-panel" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '1rem 1.5rem', borderRadius: '16px' }}>
        <h1 style={{ fontWeight: '700', fontSize: '1.5rem', margin: 0, color: '#fff' }}>ระบบรายงานผลประกอบการ (8 จังหวัดภาคเหนือตอนล่าง)</h1>
        
        <div style={{ display: 'flex', background: 'rgba(0,0,0,0.3)', borderRadius: '12px', padding: '0.25rem' }}>
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
        <div style={{ display: 'flex', gap: '1rem', alignItems: 'flex-start' }}>
          
          {/* LEFT SIDEBAR (Fixed Width) */}
          <div style={{ flex: '0 0 320px', display: 'flex', flexDirection: 'column', gap: '1rem' }}>
             
             {/* Totals Block */}
            <div style={{ 
              background: `linear-gradient(180deg, ${isIncome ? 'rgba(13,148,136,0.15)' : 'rgba(225,29,72,0.15)'} 0%, rgba(9,9,11,0.5) 100%)`,
              border: '1px solid var(--glass-border)', borderRadius: '20px', padding: '1.5rem', display: 'flex', flexDirection: 'column', gap: '1rem', boxShadow: '0 8px 32px 0 rgba(0, 0, 0, 0.37)'
            }}>
              <div>
                <h2 style={{ fontSize: '1.5rem', fontWeight: '700', color: themeColor, marginBottom: '0.25rem' }}>ภาพรวม{isIncome ? 'รายได้' : 'ค่าใช้จ่าย'}</h2>
                <p style={{ color: 'var(--text-secondary)', fontSize: '0.875rem' }}>ปี {selectedYear} {selectedMonth !== 'ทั้งหมด' ? `เดือน ${MONTH_NAMES[parseInt(selectedMonth)-1]}` : ''}</p>
              </div>

              <div style={{ background: `rgba(${isIncome ? '45,212,191' : '244,63,94'},0.15)`, padding: '1.25rem', borderRadius: '16px', border: `1px solid ${themeColor}40` }}>
                <div style={{ color: themeColor, fontSize: '0.875rem', marginBottom: '0.25rem', fontWeight: '500' }}>รวมยอดสะสม</div>
                <div style={{ fontSize: '2.5rem', fontWeight: '700', color: '#fff' }}>{formatAmt(processed.totals.actual)}</div>
              </div>

              <div style={{ display: 'flex', flexDirection: 'column', gap: '1rem', marginTop: '1rem' }}>
                  <GaugeChart title="เทียบเป้าหมายปีนี้" actual={processed.totals.actual} target={processed.totals.target} isIncome={isIncome} />
                  <GaugeChart title="เทียบส่วนปีก่อนหน้า" actual={processed.totals.actual} target={processed.totals.prev} isIncome={isIncome} />
              </div>
            </div>

             {/* Top Provinces Map */}
             <div className="glass-panel" style={{ padding: '1rem', display: 'flex', flexDirection: 'column', height: '400px' }}>
                <h3 style={{ fontSize: '1rem', fontWeight: '600', marginBottom: '1rem', color: 'var(--text-primary)' }}>แผนที่ผลงานจังหวัด</h3>
                <div style={{ flex: 1, borderRadius: '12px', overflow: 'hidden', background: '#0a0a0a' }}>
                  <MapContainer center={[16.2, 99.8]} zoom={6.5} style={{ height: '100%', width: '100%' }} zoomControl={false} scrollWheelZoom={false}>
                    <TileLayer url="https://{s}.basemaps.cartocdn.com/dark_nolabels/{z}/{x}/{y}{r}.png" />
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
                  {availableYears.map(y => <option key={y} value={y} style={{color: '#000'}}>{y}</option>)}
                </select>
              </div>

              <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                <span style={{ color: 'var(--text-secondary)', fontSize: '0.875rem' }}>เดือน</span>
                <select value={selectedMonth} onChange={e => setSelectedMonth(e.target.value)} className="filter-select">
                  <option value="ทั้งหมด" style={{color: '#000'}}>ทุกเดือน</option>
                  {MONTH_NAMES.map((m, i) => <option key={i} value={i+1} style={{color: '#000'}}>{m}</option>)}
                </select>
              </div>

              <div style={{ width: '1px', height: '20px', background: 'rgba(255,255,255,0.1)' }}></div>

              <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                <span style={{ color: 'var(--text-secondary)', fontSize: '0.875rem' }}>พื้นที่</span>
                <select value={selectedProvince} onChange={e => setSelectedProvince(e.target.value)} className="filter-select">
                  <option value="ทั้งหมด" style={{color: '#000'}}>ทุกจังหวัด (8 จังหวัด)</option>
                  {availableProvinces.map(p => <option key={p} value={p} style={{color: '#000'}}>{p}</option>)}
                </select>
                <select value={selectedOffice} onChange={e => setSelectedOffice(e.target.value)} className="filter-select" style={{maxWidth: '180px'}}>
                  <option value="ทั้งหมด" style={{color: '#000'}}>ทุกที่ทำการ</option>
                  {availableOffices.map(o => <option key={o} value={o} style={{color: '#000'}}>{o}</option>)}
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
                      <CartesianGrid strokeDasharray="3 3" stroke="rgba(255,255,255,0.05)" vertical={false} />
                      <XAxis dataKey="name" stroke="var(--text-secondary)" tick={{ fill: 'var(--text-secondary)', fontSize: 11 }} axisLine={false} tickLine={false} />
                      <YAxis stroke="var(--text-secondary)" tick={{ fill: 'var(--text-secondary)', fontSize: 11 }} axisLine={false} tickLine={false} tickFormatter={formatAmt} />
                      <RechartsTooltip formatter={(v) => formatFullAmt(v)} contentStyle={{ backgroundColor: 'rgba(9, 9, 11, 0.95)', border: 'none', borderRadius: '8px' }} />
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

                    <div style={{ display: 'flex', gap: '1rem', padding: '0.75rem 1rem', borderBottom: '1px solid rgba(255,255,255,0.1)', fontWeight: '600', color: 'var(--text-secondary)', fontSize: '0.8rem' }}>
                        <div style={{ flex: 2.5 }}>กลุ่มธุรกิจ / บริการ</div>
                        <div style={{ flex: 1, textAlign: 'right' }}>ผลงานจริง</div>
                        <div style={{ flex: 1, textAlign: 'right' }}>% สำเร็จ</div>
                    </div>

                    <div style={{ display: 'flex', flexDirection: 'column', gap: '0.5rem', marginTop: '0.5rem' }}>
                        {processed.hierarchicalData.map((bg, idx) => {
                           const isExpanded = !!expandedBGs[bg.name];
                           const pct = bg.target > 0 ? (bg.actual / bg.target) * 100 : 0;
                           
                           return (
                             <div key={idx} style={{ background: 'rgba(255,255,255,0.02)', borderRadius: '8px', overflow: 'hidden' }}>
                               <div onClick={() => toggleBG(bg.name)} style={{ display: 'flex', gap: '1rem', padding: '1rem', cursor: 'pointer', alignItems: 'center' }} className="hover:bg-white/5 transition-colors">
                                  <div style={{ flex: 2.5, display: 'flex', alignItems: 'center', gap: '0.5rem', fontWeight: '600', color: '#fff', fontSize: '0.95rem' }}>
                                     {isExpanded ? <ChevronDown size={18} color={themeColor} /> : <ChevronRight size={18} color="var(--text-secondary)" />}
                                     <span title={bg.name} style={{ whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', maxWidth: '180px' }}>{bg.name}</span>
                                  </div>
                                  <div style={{ flex: 1, textAlign: 'right', fontWeight: 'bold' }}>{formatAmt(bg.actual)}</div>
                                  <div style={{ flex: 1, textAlign: 'right', fontWeight: 'bold', color: getPerfColor(bg.actual, bg.target) }}>{pct.toFixed(0)}%</div>
                               </div>

                               {isExpanded && (
                                 <div style={{ background: 'rgba(0,0,0,0.2)', padding: '0.5rem 0' }}>
                                   {bg.evms.map((evm, eidx) => {
                                      let evmPct = evm.target > 0 ? (evm.actual / evm.target) * 100 : 0;
                                      return (
                                         <div key={eidx} style={{ display: 'flex', gap: '0.5rem', padding: '0.5rem 1rem 0.5rem 2.5rem', fontSize: '0.85rem', alignItems: 'center', borderBottom: eidx !== bg.evms.length - 1 ? '1px solid rgba(255,255,255,0.03)' : 'none' }}>
                                            <div style={{ flex: 2.5, color: 'var(--text-secondary)', display: 'flex', alignItems: 'center' }}>
                                              <span style={{ display: 'inline-block', width: '4px', height: '4px', borderRadius: '50%', background: 'var(--text-secondary)', marginRight: '0.5rem' }}></span>
                                              <span title={evm.name} style={{ whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', maxWidth: '160px' }}>{evm.name}</span>
                                            </div>
                                            <div style={{ flex: 1, textAlign: 'right', color: '#ccc' }}>{formatFullAmt(evm.actual)}</div>
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

                    <div style={{ display: 'flex', gap: '0.5rem', padding: '0.75rem 1rem', borderBottom: '1px solid rgba(255,255,255,0.1)', fontWeight: '600', color: 'var(--text-secondary)', fontSize: '0.8rem' }}>
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
                             <div key={idx} style={{ background: 'rgba(255,255,255,0.02)', borderRadius: '8px', overflow: 'hidden' }}>
                               <div onClick={() => toggleProv(prov.name)} style={{ display: 'flex', gap: '0.5rem', padding: '1rem 0.5rem', cursor: 'pointer', alignItems: 'center' }} className="hover:bg-white/5 transition-colors">
                                  <div style={{ flex: 2.2, display: 'flex', alignItems: 'center', gap: '0.5rem', fontWeight: '600', color: '#fff', fontSize: '0.95rem' }}>
                                     {isExpanded ? <ChevronDown size={18} color={themeColor} /> : <ChevronRight size={18} color="var(--text-secondary)" />}
                                     <span title={prov.name} style={{ whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', maxWidth: '140px' }}>{prov.name}</span>
                                  </div>
                                  <div style={{ flex: 0.8, textAlign: 'right', fontWeight: 'bold', fontSize: '0.9rem' }}>{formatAmt(prov.actual)}</div>
                                  <div style={{ flex: 0.8, textAlign: 'right', fontWeight: 'bold', fontSize: '0.9rem', color: getPerfColor(prov.actual, prov.target) }}>{pct.toFixed(0)}%</div>
                                  <MiniBar pct={provProp} />
                               </div>

                               {isExpanded && (
                                 <div style={{ background: 'rgba(0,0,0,0.2)', padding: '0.5rem 0' }}>
                                   {prov.offices.map((office, oidx) => {
                                      let officePct = office.target > 0 ? (office.actual / office.target) * 100 : 0;
                                      let officeProp = processed.totals.actual > 0 ? (office.actual / processed.totals.actual) * 100 : 0;
                                      return (
                                         <div key={oidx} style={{ display: 'flex', gap: '0.5rem', padding: '0.5rem 0.5rem 0.5rem 2.5rem', fontSize: '0.85rem', alignItems: 'center', borderBottom: oidx !== prov.offices.length - 1 ? '1px solid rgba(255,255,255,0.03)' : 'none' }}>
                                            <div style={{ flex: 2.2, color: 'var(--text-secondary)', display: 'flex', alignItems: 'center' }}>
                                              <span style={{ display: 'inline-block', width: '4px', height: '4px', borderRadius: '50%', background: 'var(--text-secondary)', marginRight: '0.5rem' }}></span>
                                              <span title={office.name} style={{ whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', maxWidth: '120px' }}>{office.name}</span>
                                            </div>
                                            <div style={{ flex: 0.8, textAlign: 'right', color: '#ccc' }}>{formatFullAmt(office.actual)}</div>
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
      )}

      <style>{`
        @keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
        @keyframes pulse { 0% { opacity: 0.5; } 50% { opacity: 1; } 100% { opacity: 0.5; } }
        .filter-select { background: rgba(255,255,255,0.1); padding: 0.35rem 0.5rem; border-radius: 8px; color: #fff; border: 1px solid rgba(255,255,255,0.1); outline: none; font-family: Outfit; font-weight: 500;}
        .hover\\:bg-white\\/5:hover { background-color: rgba(255, 255, 255, 0.05); }
        .transition-colors { transition-property: color, background-color, border-color; transition-timing-function: cubic-bezier(0.4, 0, 0.2, 1); transition-duration: 150ms; }
      `}</style>
    </div>
  );
};

export default Dashboard;
