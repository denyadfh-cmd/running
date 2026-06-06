import { useState, useMemo, useEffect, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  AreaChart, Area, LineChart, Line, BarChart, Bar,
  XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer
} from "recharts";

/* ─── helpers ─── */
const MONTHS_ID = ["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];
const MONTHS_SHORT = ["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];
const DAYS_ID = ["Minggu","Senin","Selasa","Rabu","Kamis","Jumat","Sabtu"];
const CY = new Date().getFullYear();

const fmtPace = (p) => {
  if (!p) return "—";
  const m = Math.floor(p), s = Math.round((p - m) * 60);
  return `${m}:${s.toString().padStart(2,"0")}`;
};
const fmtDate = (iso) => {
  const [y,mo,d] = iso.split("-");
  return `${d}/${mo}/${y}`;
};
const pct = (a,b) => b ? (((a-b)/b)*100).toFixed(1) : null;
const LS_KEY = (y) => `rundata_${y}`;

/* ─── localStorage helpers ─── */
const loadYear = (y) => {
  try { const d = localStorage.getItem(LS_KEY(y)); return d ? JSON.parse(d) : []; }
  catch { return []; }
};
const saveYear = (y, data) => {
  try { localStorage.setItem(LS_KEY(y), JSON.stringify(data)); } catch {}
};
const loadYearList = () => {
  try {
    const raw = localStorage.getItem("run_years");
    return raw ? JSON.parse(raw) : [CY];
  } catch { return [CY]; }
};
const saveYearList = (list) => {
  try { localStorage.setItem("run_years", JSON.stringify(list)); } catch {}
};

/* ─── excel export ─── */
const exportExcel = (sessions, year) => {
  const wb = XLSX.utils.book_new();
  const sorted = [...sessions].sort((a,b)=>new Date(a.date)-new Date(b.date));

  // ── Sheet 1: Log Sesi ──
  const logRows = sorted.map((s,i) => {
    const d = new Date(s.date);
    return [
      i+1,
      fmtDate(s.date),
      DAYS_ID[d.getDay()],
      MONTHS_ID[d.getMonth()],
      +s.distance.toFixed(1),
      `${fmtPace(s.pace)} /km`,
      s.duration,
    ];
  });
  const logHeaders = ["No","Tanggal","Hari","Bulan","Jarak (km)","Pace","Durasi (menit)"];
  const logData = [logHeaders, ...logRows];
  const ws1 = XLSX.utils.aoa_to_sheet(logData);

  // column widths
  ws1["!cols"] = [
    {wch:5},{wch:12},{wch:10},{wch:12},{wch:12},{wch:14},{wch:16}
  ];
  ws1["!freeze"] = {xSplit:0, ySplit:1};

  // header style (SheetJS CE supports limited styling via cell format)
  const headerStyle = {
    font:{bold:true,color:{rgb:"FFFFFF"},sz:11},
    fill:{fgColor:{rgb:"1a472a"}},
    alignment:{horizontal:"center",vertical:"center"},
    border:{top:{style:"thin"},bottom:{style:"thin"},left:{style:"thin"},right:{style:"thin"}}
  };
  const dataStyle = (row) => ({
    font:{sz:11},
    fill:{fgColor:{rgb: row%2===0 ? "FFFFFF" : "F5F5F5"}},
    alignment:{vertical:"center"},
    border:{top:{style:"thin"},bottom:{style:"thin"},left:{style:"thin"},right:{style:"thin"}}
  });

  logHeaders.forEach((_,ci) => {
    const ref = XLSX.utils.encode_cell({r:0,c:ci});
    if (ws1[ref]) ws1[ref].s = headerStyle;
  });
  logRows.forEach((row,ri) => {
    row.forEach((_,ci) => {
      const ref = XLSX.utils.encode_cell({r:ri+1,c:ci});
      if (ws1[ref]) ws1[ref].s = dataStyle(ri);
    });
  });

  XLSX.utils.book_append_sheet(wb, ws1, "Log Sesi");

  // ── Sheet 2: Ringkasan Bulanan ──
  const monthly = MONTHS_ID.map((name,idx)=>{
    const ms = sorted.filter(s=>new Date(s.date).getMonth()===idx);
    const dist = +ms.reduce((a,b)=>a+b.distance,0).toFixed(1);
    const pace = ms.length ? +(ms.reduce((a,b)=>a+b.pace,0)/ms.length).toFixed(2) : 0;
    const maxD = ms.length ? +Math.max(...ms.map(s=>s.distance)).toFixed(1) : 0;
    return {name, dist, count:ms.length, pace, maxD};
  });
  const bestMonth = monthly.reduce((a,b)=>b.dist>a.dist?b:a, monthly[0]);

  const mHeaders = ["Bulan","Total Jarak (km)","Jumlah Sesi","Avg Pace","Jarak Terpanjang (km)","Tren vs Bulan Lalu"];
  const mRows = monthly.map((m,i)=>{
    const prev = monthly[i-1];
    let tren = "—";
    if (prev && prev.dist > 0 && m.dist > 0) {
      const d = pct(m.dist, prev.dist);
      tren = parseFloat(d) >= 0 ? `↑ +${d}%` : `↓ ${d}%`;
    }
    return [m.name, m.dist||"", m.count||"", m.pace ? `${fmtPace(m.pace)} /km` : "—", m.maxD||"", tren];
  });
  const ws2 = XLSX.utils.aoa_to_sheet([mHeaders, ...mRows]);
  ws2["!cols"] = [{wch:14},{wch:18},{wch:14},{wch:14},{wch:22},{wch:20}];
  ws2["!freeze"] = {xSplit:0, ySplit:1};

  const mHeaderStyle = {
    font:{bold:true,color:{rgb:"FFFFFF"},sz:11},
    fill:{fgColor:{rgb:"1a3a5c"}},
    alignment:{horizontal:"center",vertical:"center"},
    border:{top:{style:"thin"},bottom:{style:"thin"},left:{style:"thin"},right:{style:"thin"}}
  };
  mHeaders.forEach((_,ci)=>{
    const ref = XLSX.utils.encode_cell({r:0,c:ci});
    if (ws2[ref]) ws2[ref].s = mHeaderStyle;
  });
  mRows.forEach((row,ri)=>{
    const isBest = monthly[ri].name === bestMonth.name && bestMonth.dist > 0;
    row.forEach((_,ci)=>{
      const ref = XLSX.utils.encode_cell({r:ri+1,c:ci});
      if (ws2[ref]) ws2[ref].s = {
        font:{sz:11, bold: isBest},
        fill:{fgColor:{rgb: isBest ? "FFFDE7" : (ri%2===0?"FFFFFF":"F5F5F5")}},
        alignment:{vertical:"center"},
        border:{top:{style:"thin"},bottom:{style:"thin"},left:{style:"thin"},right:{style:"thin"}}
      };
    });
  });

  XLSX.utils.book_append_sheet(wb, ws2, "Ringkasan Bulanan");

  // ── Sheet 3: Statistik Tahunan ──
  const total = +sorted.reduce((a,b)=>a+b.distance,0).toFixed(1);
  const avgP = sorted.length ? +(sorted.reduce((a,b)=>a+b.pace,0)/sorted.length).toFixed(2) : 0;
  const fastest = sorted.length ? sorted.reduce((a,b)=>b.pace<a.pace?b:a) : null;
  const longest = sorted.length ? sorted.reduce((a,b)=>b.distance>a.distance?b:a) : null;
  const bestMo = monthly.reduce((a,b)=>b.dist>a.dist?b:a, monthly[0]);
  const mostActive = monthly.reduce((a,b)=>b.count>a.count?b:a, monthly[0]);

  // consistency: % of weeks with at least 1 run
  const weeksWithRun = new Set(sorted.map(s=>{
    const diff = (new Date(s.date)-new Date(`${year}-01-01`))/(7*86400000);
    return Math.floor(diff);
  })).size;
  const totalWeeks = 52;
  const consistency = `${Math.round((weeksWithRun/totalWeeks)*100)}%`;

  const statRows = [
    ["Tahun", year],
    ["Total Jarak (km)", total || "—"],
    ["Total Sesi", sorted.length || "—"],
    ["Rata-rata Pace", avgP ? `${fmtPace(avgP)} /km` : "—"],
    ["Pace Terbaik (tercepat)", fastest ? `${fmtPace(fastest.pace)} /km` : "—"],
    ["Jarak Terpanjang", longest ? `${longest.distance} km` : "—"],
    ["Bulan Terbaik (jarak)", bestMo.dist > 0 ? `${bestMo.name} — ${bestMo.dist} km` : "—"],
    ["Bulan Teraktif (sesi)", mostActive.count > 0 ? `${mostActive.name} — ${mostActive.count} sesi` : "—"],
    ["Konsistensi", sorted.length ? consistency : "—"],
  ];
  const ws3 = XLSX.utils.aoa_to_sheet(statRows);
  ws3["!cols"] = [{wch:30},{wch:24}];

  statRows.forEach((_,ri)=>{
    const rA = XLSX.utils.encode_cell({r:ri,c:0});
    const rB = XLSX.utils.encode_cell({r:ri,c:1});
    if (ws3[rA]) ws3[rA].s = {
      font:{bold:true,sz:11},
      fill:{fgColor:{rgb:"F0F0F0"}},
      border:{top:{style:"thin"},bottom:{style:"thin"},left:{style:"thin"},right:{style:"thin"}},
      alignment:{vertical:"center"}
    };
    if (ws3[rB]) ws3[rB].s = {
      font:{sz:11,color:{rgb:"1a6b3a"},bold:true},
      alignment:{horizontal:"right",vertical:"center"},
      border:{top:{style:"thin"},bottom:{style:"thin"},left:{style:"thin"},right:{style:"thin"}}
    };
  });

  XLSX.utils.book_append_sheet(wb, ws3, "Statistik Tahunan");

  // ── write file ──
  const today = new Date();
  const dd = String(today.getDate()).padStart(2,"0");
  const mm = String(today.getMonth()+1).padStart(2,"0");
  const filename = `Lari_${year}_${mm}-${dd}.xlsx`;
  XLSX.writeFile(wb, filename);
};

/* ══════════════════════════════════════════════════ */
export default function App() {
  const [yearList, setYearList]   = useState(() => loadYearList());
  const [year, setYear]           = useState(CY);
  const [allData, setAllData]     = useState(() => {
    const init = {};
    loadYearList().forEach(y => { init[y] = loadYear(y); });
    return init;
  });
  const [tab, setTab]             = useState("dashboard");
  const [showForm, setShowForm]   = useState(false);
  const [editId, setEditId]       = useState(null);
  const [form, setForm]           = useState({date:"",distance:"",pace:""});
  const [formErr, setFormErr]     = useState("");
  const [cmpA, setCmpA]           = useState(0);
  const [cmpB, setCmpB]           = useState(1);
  const [showAddYear, setShowAddYear] = useState(false);
  const [newYearInput, setNewYearInput] = useState("");
  const [deleteConfirm, setDeleteConfirm] = useState(null);

  // sessions for active year, sorted
  const sessions = useMemo(() =>
    (allData[year]||[]).slice().sort((a,b)=>new Date(a.date)-new Date(b.date)),
    [allData, year]
  );

  // save to localStorage whenever data changes
  useEffect(() => {
    Object.entries(allData).forEach(([y,d]) => saveYear(+y, d));
  }, [allData]);

  useEffect(() => { saveYearList(yearList); }, [yearList]);

  // switch year
  const switchYear = (y) => {
    setYear(y);
    if (!allData[y]) {
      setAllData(prev => ({...prev, [y]: loadYear(y)}));
    }
    setTab("dashboard");
    setShowForm(false);
  };

  // add new year
  const addYear = () => {
    const y = parseInt(newYearInput);
    if (!y || y < 1900 || y > 2200) { setNewYearInput(""); return; }
    if (!yearList.includes(y)) {
      const newList = [...yearList, y].sort((a,b)=>a-b);
      setYearList(newList);
      setAllData(prev => ({...prev, [y]: loadYear(y)}));
    }
    switchYear(y);
    setShowAddYear(false);
    setNewYearInput("");
  };

  // save session
  const saveSession = () => {
    if (!form.date || !form.distance || !form.pace) { setFormErr("Semua field wajib diisi"); return; }
    const dist = parseFloat(form.distance);
    const pace = parseFloat(form.pace);
    if (isNaN(dist)||dist<=0||isNaN(pace)||pace<=0) { setFormErr("Jarak dan pace harus angka positif"); return; }
    const d = new Date(form.date);
    const entry = {
      id: editId || `s-${Date.now()}`,
      date: form.date,
      month: d.getMonth(),
      distance: +dist.toFixed(2),
      pace: +pace.toFixed(2),
      duration: Math.round(dist * pace),
    };
    setAllData(prev => {
      const list = (prev[year]||[]).filter(s=>s.id!==entry.id);
      return {...prev, [year]: [...list, entry]};
    });
    setForm({date:"",distance:"",pace:""}); setEditId(null);
    setShowForm(false); setFormErr("");
  };

  const deleteSession = (id) => {
    setAllData(prev => ({...prev, [year]: (prev[year]||[]).filter(s=>s.id!==id)}));
    setDeleteConfirm(null);
  };

  const startEdit = (s) => {
    setForm({date:s.date, distance:String(s.distance), pace:String(s.pace)});
    setEditId(s.id); setShowForm(true); setTab("dashboard"); setFormErr("");
  };

  /* stats */
  const yearly = useMemo(()=>{
    if (!sessions.length) return null;
    const total = +sessions.reduce((a,b)=>a+b.distance,0).toFixed(1);
    const avgP  = +(sessions.reduce((a,b)=>a+b.pace,0)/sessions.length).toFixed(2);
    const best  = sessions.reduce((a,b)=>b.distance>a.distance?b:a);
    return {total, avgPace:avgP, count:sessions.length, best};
  },[sessions]);

  const monthly = useMemo(()=>MONTHS_SHORT.map((name,idx)=>{
    const ms = sessions.filter(s=>s.month===idx);
    const dist = +ms.reduce((a,b)=>a+b.distance,0).toFixed(1);
    const pace = ms.length ? +(ms.reduce((a,b)=>a+b.pace,0)/ms.length).toFixed(2) : 0;
    return {name, month:idx, dist, pace, count:ms.length};
  }),[sessions]);

  const cumulative = useMemo(()=>{
    let cum=0;
    return sessions.map(s=>{cum+=s.distance;return{date:s.date.slice(5),cum:+cum.toFixed(1)};});
  },[sessions]);

  const weekly = useMemo(()=>{
    const map={};
    sessions.forEach(s=>{
      const w=Math.floor((new Date(s.date)-new Date(`${year}-01-01`))/6048e5);
      const k=`W${w+1}`;
      map[k]=(map[k]||0)+s.distance;
    });
    return Object.entries(map).map(([k,v])=>({k,dist:+v.toFixed(1)}));
  },[sessions,year]);

  const mA=monthly[cmpA], mB=monthly[cmpB];

  const trendLabel=(curr,prev)=>{
    const d=pct(curr,prev);
    if(d===null||!prev) return {text:"—",color:"#556"};
    const n=parseFloat(d);
    if(n>10)  return {text:`↑↑ +${d}%`,color:"#00ff9d"};
    if(n>0)   return {text:`↑ +${d}%`, color:"#7cfc00"};
    if(n<-10) return {text:`↓↓ ${d}%`, color:"#ff4d6d"};
    return          {text:`↓ ${d}%`,   color:"#ff9f45"};
  };

  /* ═══ RENDER ═══ */
  return (
    <div style={{minHeight:"100vh",maxWidth:480,margin:"0 auto",background:"#080c10",
      color:"#e8f0fe",fontFamily:"'Barlow Condensed',sans-serif",
      display:"flex",flexDirection:"column",position:"relative",overflowX:"hidden"}}>

      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@300;400;600;700;800&family=Barlow:wght@300;400;500&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        ::-webkit-scrollbar{display:none}
        .card{background:#0e1520;border:1px solid #1a2535;border-radius:12px;padding:16px}
        .lbl{font-size:10px;font-weight:700;letter-spacing:2px;text-transform:uppercase;color:#445}
        .big{font-family:'Barlow Condensed';font-weight:800;font-size:34px;line-height:1}
        .inp{background:#0e1520;border:1px solid #2a3545;border-radius:8px;color:#e8f0fe;
             padding:10px 12px;width:100%;font-family:'Barlow';font-size:14px;outline:none}
        .inp:focus{border-color:#00ff9d66}
        .btn-g{background:#00ff9d;color:#080c10;border:none;border-radius:8px;
               padding:11px 0;width:100%;font-family:'Barlow Condensed';font-weight:800;
               font-size:16px;letter-spacing:1px;cursor:pointer}
        .btn-ghost{background:transparent;border:1px solid #2a3545;border-radius:8px;
                   padding:9px 0;width:100%;font-family:'Barlow Condensed';font-weight:700;
                   font-size:14px;color:#889;cursor:pointer;letter-spacing:1px}
        .btn-export{background:transparent;border:1px solid #00ff9d;border-radius:8px;
                    padding:7px 12px;font-family:'Barlow Condensed';font-weight:700;
                    font-size:13px;color:#00ff9d;cursor:pointer;letter-spacing:1px;
                    display:flex;align-items:center;gap:5px;transition:background .15s}
        .btn-export:hover{background:#00ff9d18}
        .btn-export:disabled{opacity:.35;cursor:not-allowed;border-color:#445;color:#445}
        .chip{background:#0e1520;border:1px solid #1a2535;border-radius:20px;
              padding:5px 12px;cursor:pointer;font-family:'Barlow Condensed';
              font-weight:700;font-size:13px;letter-spacing:1px;color:#889;transition:all .15s}
        .chip.a{border-color:#00ff9d;color:#00ff9d;background:#00ff9d11}
        .chip.b{border-color:#7eb8f7;color:#7eb8f7;background:#7eb8f711}
        .yr-tab{background:transparent;border:none;padding:7px 14px;cursor:pointer;
                font-family:'Barlow Condensed';font-weight:700;font-size:15px;
                letter-spacing:1px;border-bottom:2px solid transparent;transition:all .15s;color:#445}
        .yr-tab.active{color:#00ff9d;border-bottom-color:#00ff9d}
        .yr-tab:hover{color:#aaa}
        .del-btn{background:transparent;border:none;color:#ff4d6d;cursor:pointer;
                 font-size:16px;padding:4px 8px;font-weight:700}
        .edit-btn{background:transparent;border:none;color:#7eb8f7;cursor:pointer;
                  font-family:'Barlow Condensed';font-weight:700;font-size:12px;
                  letter-spacing:1px;padding:4px 8px}
        .hcell{width:11px;height:11px;border-radius:2px;flex-shrink:0;cursor:pointer}
        @keyframes up{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}
        .fade{animation:up .3s ease forwards}
        .modal-bg{position:fixed;inset:0;background:#00000088;z-index:200;
                  display:flex;align-items:center;justify-content:center;padding:20px}
        .modal{background:#0e1520;border:1px solid #2a3545;border-radius:16px;
               padding:24px;width:100%;max-width:320px}
      `}</style>

      {/* ── HEADER ── */}
      <div style={{padding:"14px 16px 0",borderBottom:"1px solid #1a2535",background:"#080c10",
        position:"sticky",top:0,zIndex:50}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10}}>
          <div>
            <div style={{fontSize:9,fontWeight:700,letterSpacing:3,color:"#00ff9d"}}>RUN TRACKER</div>
            <div style={{fontFamily:"'Barlow Condensed'",fontWeight:800,fontSize:22,lineHeight:1}}>
              PERFORMA LARI <span style={{color:"#00ff9d"}}>{year}</span>
            </div>
          </div>
          <div style={{display:"flex",gap:8,alignItems:"center"}}>
            <button className="btn-export" disabled={!sessions.length}
              onClick={()=>exportExcel(sessions,year)} title={!sessions.length?"Belum ada data":""}>
              📊 Export
            </button>
            <button onClick={()=>{setEditId(null);setForm({date:"",distance:"",pace:""});
              setFormErr("");setShowForm(s=>!s);}}
              style={{background:showForm?"#ff4d6d":"#00ff9d",color:"#080c10",border:"none",
                borderRadius:8,width:34,height:34,fontWeight:800,fontSize:20,cursor:"pointer",
                display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
              {showForm?"×":"+"}
            </button>
          </div>
        </div>

        {/* year tabs */}
        <div style={{display:"flex",gap:0,overflowX:"auto",alignItems:"center"}}>
          {yearList.map(y=>(
            <button key={y} className={`yr-tab ${year===y?"active":""}`} onClick={()=>switchYear(y)}>{y}</button>
          ))}
          <button onClick={()=>setShowAddYear(true)}
            style={{background:"transparent",border:"none",color:"#445",cursor:"pointer",
              padding:"7px 12px",fontFamily:"'Barlow Condensed'",fontWeight:700,fontSize:18,
              lineHeight:1,flexShrink:0}}>＋</button>
        </div>
      </div>

      {/* ── FORM ── */}
      {showForm && (
        <div style={{padding:"12px 16px",borderBottom:"1px solid #1a2535",background:"#0a1018"}} className="fade">
          <div style={{fontSize:11,fontWeight:700,letterSpacing:2,color:"#00ff9d",marginBottom:10}}>
            {editId?"EDIT SESI":"TAMBAH SESI — "+year}
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:8}}>
            <div>
              <div className="lbl" style={{marginBottom:4}}>Tanggal</div>
              <input type="date" className="inp" value={form.date}
                onChange={e=>setForm({...form,date:e.target.value})}/>
            </div>
            <div>
              <div className="lbl" style={{marginBottom:4}}>Jarak km</div>
              <input type="number" placeholder="10.5" className="inp" value={form.distance}
                onChange={e=>setForm({...form,distance:e.target.value})} step="0.1" min="0"/>
            </div>
            <div>
              <div className="lbl" style={{marginBottom:4}}>Pace /km</div>
              <input type="number" placeholder="5.5" className="inp" value={form.pace}
                onChange={e=>setForm({...form,pace:e.target.value})} step="0.1" min="0"/>
            </div>
          </div>
          {formErr && <div style={{fontSize:12,color:"#ff4d6d",marginBottom:8}}>{formErr}</div>}
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
            <button className="btn-ghost" onClick={()=>{setShowForm(false);setEditId(null);setFormErr("");}}>Batal</button>
            <button className="btn-g" onClick={saveSession}>Simpan</button>
          </div>
        </div>
      )}

      {/* ── CONTENT ── */}
      <div style={{flex:1,overflowY:"auto",paddingBottom:80}}>

        {/* ═══ DASHBOARD ═══ */}
        {tab==="dashboard" && (
          <div style={{padding:"14px 16px",display:"flex",flexDirection:"column",gap:12}} className="fade">
            {!sessions.length && <EmptyState year={year}/>}

            {yearly && <>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                {[
                  {label:"Total Jarak",val:yearly.total,unit:" km",color:"#00ff9d"},
                  {label:"Total Sesi",val:yearly.count,unit:"x",color:"#7eb8f7"},
                  {label:"Avg Pace",val:fmtPace(yearly.avgPace),unit:" /km",color:"#ffb347",sm:true},
                  {label:"Terpanjang",val:yearly.best.distance,unit:" km",color:"#ff6b9d"},
                ].map(c=>(
                  <div key={c.label} className="card">
                    <div className="lbl">{c.label}</div>
                    <div className="big" style={{color:c.color,fontSize:c.sm?26:34}}>
                      {c.val}<span style={{fontSize:14,color:"#556"}}>{c.unit}</span>
                    </div>
                  </div>
                ))}
              </div>

              {cumulative.length>1 && (
                <div className="card">
                  <div className="lbl" style={{marginBottom:10}}>Akumulasi Jarak {year}</div>
                  <ResponsiveContainer width="100%" height={150}>
                    <AreaChart data={cumulative} margin={{top:4,right:4,left:-20,bottom:0}}>
                      <defs>
                        <linearGradient id="cg" x1="0" y1="0" x2="0" y2="1">
                          <stop offset="0%" stopColor="#00ff9d" stopOpacity={0.3}/>
                          <stop offset="100%" stopColor="#00ff9d" stopOpacity={0}/>
                        </linearGradient>
                      </defs>
                      <CartesianGrid strokeDasharray="3 3" stroke="#1a2535"/>
                      <XAxis dataKey="date" tick={{fill:"#445",fontSize:9}} tickLine={false}
                        interval={Math.max(Math.floor(cumulative.length/5),1)}/>
                      <YAxis tick={{fill:"#445",fontSize:9}} tickLine={false} axisLine={false}/>
                      <Tooltip contentStyle={{background:"#0e1520",border:"1px solid #1a2535",
                        borderRadius:8,fontFamily:"'Barlow Condensed'",fontSize:12}}
                        formatter={v=>[`${v} km`,"Total"]}/>
                      <Area type="monotone" dataKey="cum" stroke="#00ff9d" strokeWidth={2}
                        fill="url(#cg)" dot={false}/>
                    </AreaChart>
                  </ResponsiveContainer>
                </div>
              )}

              {weekly.length>1 && (
                <div className="card">
                  <div className="lbl" style={{marginBottom:10}}>Jarak per Minggu</div>
                  <ResponsiveContainer width="100%" height={120}>
                    <BarChart data={weekly} margin={{top:4,right:4,left:-20,bottom:0}} barSize={7}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#1a2535"/>
                      <XAxis dataKey="k" tick={{fill:"#445",fontSize:9}} tickLine={false} interval={3}/>
                      <YAxis tick={{fill:"#445",fontSize:9}} tickLine={false} axisLine={false}/>
                      <Tooltip contentStyle={{background:"#0e1520",border:"1px solid #1a2535",
                        borderRadius:8,fontFamily:"'Barlow Condensed'",fontSize:12}}
                        formatter={v=>[`${v} km`,"Jarak"]}/>
                      <Bar dataKey="dist" fill="#7eb8f7" radius={[3,3,0,0]}/>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              )}

              <div className="card">
                <div className="lbl" style={{marginBottom:10}}>Kalender {year}</div>
                <HeatmapCalendar sessions={sessions} year={year}/>
              </div>
            </>}
          </div>
        )}

        {/* ═══ BULANAN ═══ */}
        {tab==="bulanan" && (
          <div style={{padding:"14px 16px",display:"flex",flexDirection:"column",gap:10}} className="fade">
            {!sessions.length && <EmptyState year={year}/>}

            {monthly.map((m,i)=>{
              const prev=monthly[i-1];
              const tr=trendLabel(m.dist,prev?.dist);
              const ms=sessions.filter(s=>s.month===i);
              const maxD=Math.max(...ms.map(s=>s.distance),1);
              return (
                <div key={i} className="card" style={{opacity:m.count===0?.4:1}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:6}}>
                    <div>
                      <div className="lbl">{MONTHS_ID[i]} {year}</div>
                      <div className="big" style={{fontSize:26}}>{m.dist||0}
                        <span style={{fontSize:13,color:"#556"}}> km</span>
                      </div>
                    </div>
                    <div style={{textAlign:"right"}}>
                      {m.count>0 && <div style={{fontSize:12,fontWeight:700,color:tr.color}}>{tr.text}</div>}
                      <div style={{fontSize:11,color:"#445"}}>{m.count} sesi</div>
                      {m.pace>0 && <div style={{fontSize:11,color:"#ffb347"}}>{fmtPace(m.pace)} /km</div>}
                    </div>
                  </div>
                  {ms.length>0 ? (
                    <div style={{display:"flex",gap:2,alignItems:"flex-end",height:28}}>
                      {ms.map((s,j)=>(
                        <div key={j} style={{flex:1,background:"#00ff9d",borderRadius:"2px 2px 0 0",
                          height:`${(s.distance/maxD)*100}%`,opacity:.6,minHeight:3}} title={`${s.distance}km`}/>
                      ))}
                    </div>
                  ):(
                    <div style={{fontSize:11,color:"#2a3545",textAlign:"center",paddingTop:2}}>Tidak ada sesi</div>
                  )}
                </div>
              );
            })}

            {monthly.filter(m=>m.pace>0).length>1 && (
              <div className="card">
                <div className="lbl" style={{marginBottom:10}}>Tren Pace Bulanan</div>
                <ResponsiveContainer width="100%" height={150}>
                  <LineChart data={monthly.filter(m=>m.pace>0)} margin={{top:4,right:4,left:-20,bottom:0}}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#1a2535"/>
                    <XAxis dataKey="name" tick={{fill:"#445",fontSize:9}} tickLine={false}/>
                    <YAxis domain={["auto","auto"]} tick={{fill:"#445",fontSize:9}} tickLine={false}
                      axisLine={false} tickFormatter={v=>fmtPace(v)}/>
                    <Tooltip contentStyle={{background:"#0e1520",border:"1px solid #1a2535",
                      borderRadius:8,fontFamily:"'Barlow Condensed'",fontSize:12}}
                      formatter={v=>[fmtPace(v),"Pace"]}/>
                    <Line type="monotone" dataKey="pace" stroke="#ffb347" strokeWidth={2}
                      dot={{fill:"#ffb347",strokeWidth:0,r:3}} activeDot={{r:5}}/>
                  </LineChart>
                </ResponsiveContainer>
              </div>
            )}
          </div>
        )}

        {/* ═══ PERBANDINGAN ═══ */}
        {tab==="perbandingan" && (
          <div style={{padding:"14px 16px",display:"flex",flexDirection:"column",gap:12}} className="fade">
            {!sessions.length && <EmptyState year={year}/>}

            <div className="card">
              <div className="lbl" style={{marginBottom:8,color:"#00ff9d"}}>Bulan A</div>
              <div style={{display:"flex",flexWrap:"wrap",gap:5,marginBottom:14}}>
                {MONTHS_SHORT.map((m,i)=>(
                  <button key={i} className={`chip ${cmpA===i?"a":""}`} onClick={()=>setCmpA(i)}>{m}</button>
                ))}
              </div>
              <div className="lbl" style={{marginBottom:8,color:"#7eb8f7"}}>Bulan B</div>
              <div style={{display:"flex",flexWrap:"wrap",gap:5}}>
                {MONTHS_SHORT.map((m,i)=>(
                  <button key={i} className={`chip ${cmpB===i?"b":""}`} onClick={()=>setCmpB(i)}>{m}</button>
                ))}
              </div>
            </div>

            <div style={{display:"grid",gridTemplateColumns:"1fr 36px 1fr",gap:8,alignItems:"center"}}>
              <CmpCard m={mA} color="#00ff9d" label="A" year={year}/>
              <div style={{textAlign:"center",fontFamily:"'Barlow Condensed'",fontWeight:800,
                fontSize:16,color:"#2a3545"}}>VS</div>
              <CmpCard m={mB} color="#7eb8f7" label="B" year={year}/>
            </div>

            {(mA.count>0||mB.count>0) && (
              <div className="card">
                <div className="lbl" style={{marginBottom:10}}>Perbandingan Langsung</div>
                <ResponsiveContainer width="100%" height={170}>
                  <BarChart layout="vertical" margin={{top:4,right:8,left:55,bottom:0}}
                    data={[
                      {m:"Jarak (km)",a:mA.dist,b:mB.dist},
                      {m:"Sesi",a:mA.count,b:mB.count},
                      {m:"Pace ×10",a:+(mA.pace*10).toFixed(1),b:+(mB.pace*10).toFixed(1)},
                    ]}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#1a2535" horizontal={false}/>
                    <XAxis type="number" tick={{fill:"#445",fontSize:9}} tickLine={false} axisLine={false}/>
                    <YAxis type="category" dataKey="m" tick={{fill:"#889",fontSize:11,
                      fontFamily:"'Barlow Condensed'",fontWeight:700}} tickLine={false} axisLine={false}/>
                    <Tooltip contentStyle={{background:"#0e1520",border:"1px solid #1a2535",
                      borderRadius:8,fontFamily:"'Barlow Condensed'",fontSize:12}}/>
                    <Bar dataKey="a" name={MONTHS_SHORT[cmpA]} fill="#00ff9d" radius={[0,4,4,0]} barSize={14}/>
                    <Bar dataKey="b" name={MONTHS_SHORT[cmpB]} fill="#7eb8f7" radius={[0,4,4,0]} barSize={14}/>
                  </BarChart>
                </ResponsiveContainer>
              </div>
            )}

            <div className="card">
              <div className="lbl" style={{marginBottom:10}}>Kesimpulan</div>
              {[
                {k:"Jarak",a:mA.dist,b:mB.dist,u:" km",hi:true,f:null},
                {k:"Sesi",a:mA.count,b:mB.count,u:"x",hi:true,f:null},
                {k:"Pace",a:mA.pace,b:mB.pace,u:"",hi:false,f:fmtPace},
              ].map(({k,a,b,u,hi,f})=>{
                const win=hi?(a>b?"A":a<b?"B":"="):(a&&b?(a<b?"A":a>b?"B":"="):"=");
                const d=a&&b?Math.abs(parseFloat(pct(a,b))):null;
                return (
                  <div key={k} style={{display:"flex",alignItems:"center",gap:8,
                    padding:"9px 0",borderBottom:"1px solid #1a2535"}}>
                    <div style={{width:48,fontSize:11,fontWeight:700,letterSpacing:1,color:"#556"}}>{k}</div>
                    <div style={{fontFamily:"'Barlow Condensed'",fontWeight:800,fontSize:17,
                      color:win==="A"?"#00ff9d":"#889",minWidth:48}}>{f?f(a):(a||0)+u}</div>
                    <div style={{flex:1,height:3,background:"#1a2535",borderRadius:2,overflow:"hidden",position:"relative"}}>
                      {a&&b&&<div style={{position:"absolute",left:0,top:0,height:"100%",
                        width:`${Math.min((a/(a+b))*100,100)}%`,background:"#00ff9d",borderRadius:2}}/>}
                    </div>
                    <div style={{fontFamily:"'Barlow Condensed'",fontWeight:800,fontSize:17,
                      color:win==="B"?"#7eb8f7":"#889",minWidth:48,textAlign:"right"}}>{f?f(b):(b||0)+u}</div>
                    <div style={{width:40,fontSize:10,color:"#445",textAlign:"right",fontWeight:700}}>
                      {win!=="="?<><span style={{color:win==="A"?"#00ff9d":"#7eb8f7"}}>{win}</span>
                        {d!==null&&<span style={{display:"block",fontSize:9}}>{d.toFixed(1)}%</span>}</>:"—"}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* ═══ LOG ═══ */}
        {tab==="log" && (
          <div style={{padding:"14px 16px",display:"flex",flexDirection:"column",gap:8}} className="fade">
            <div className="lbl">{sessions.length} SESI — {year}</div>
            {!sessions.length && <EmptyState year={year}/>}
            {[...sessions].reverse().map(s=>(
              <div key={s.id} className="card" style={{display:"flex",alignItems:"center",gap:8}}>
                <div style={{flex:1,minWidth:0}}>
                  <div className="lbl">{fmtDate(s.date)} · {DAYS_ID[new Date(s.date).getDay()]}</div>
                  <div style={{display:"flex",gap:12,alignItems:"baseline",marginTop:2}}>
                    <span style={{fontFamily:"'Barlow Condensed'",fontWeight:800,fontSize:20,color:"#00ff9d"}}>
                      {s.distance}<span style={{fontSize:11,color:"#556"}}> km</span>
                    </span>
                    <span style={{fontFamily:"'Barlow Condensed'",fontWeight:700,fontSize:14,color:"#ffb347"}}>
                      {fmtPace(s.pace)}<span style={{fontSize:10,color:"#445"}}>/km</span>
                    </span>
                    <span style={{fontSize:11,color:"#445"}}>{s.duration}min</span>
                  </div>
                </div>
                <button className="edit-btn" onClick={()=>startEdit(s)}>EDIT</button>
                {deleteConfirm===s.id ? (
                  <div style={{display:"flex",gap:4}}>
                    <button onClick={()=>deleteSession(s.id)} style={{background:"#ff4d6d",color:"#fff",
                      border:"none",borderRadius:6,padding:"4px 8px",fontSize:11,cursor:"pointer",fontWeight:700}}>Hapus</button>
                    <button onClick={()=>setDeleteConfirm(null)} style={{background:"#1a2535",color:"#889",
                      border:"none",borderRadius:6,padding:"4px 8px",fontSize:11,cursor:"pointer"}}>Batal</button>
                  </div>
                ):(
                  <button className="del-btn" onClick={()=>setDeleteConfirm(s.id)}>×</button>
                )}
              </div>
            ))}
          </div>
        )}
      </div>

      {/* ── BOTTOM NAV ── */}
      <nav style={{position:"fixed",bottom:0,left:"50%",transform:"translateX(-50%)",
        width:"100%",maxWidth:480,background:"#0a0f18",borderTop:"1px solid #1a2535",
        display:"flex",zIndex:100}}>
        {[
          {id:"dashboard",icon:"⚡",label:"Dashboard"},
          {id:"bulanan",icon:"📅",label:"Bulanan"},
          {id:"perbandingan",icon:"⚖️",label:"Bandingkan"},
          {id:"log",icon:"📋",label:"Log"},
        ].map(t=>(
          <button key={t.id} onClick={()=>setTab(t.id)} style={{flex:1,background:"transparent",
            border:"none",padding:"10px 0 12px",cursor:"pointer",display:"flex",
            flexDirection:"column",alignItems:"center",gap:2}}>
            <span style={{fontSize:18,lineHeight:1}}>{t.icon}</span>
            <span style={{fontFamily:"'Barlow Condensed'",fontWeight:700,fontSize:10,
              letterSpacing:1,textTransform:"uppercase",
              color:tab===t.id?"#00ff9d":"#445",transition:"color .2s"}}>{t.label}</span>
            {tab===t.id&&<div style={{width:18,height:2,background:"#00ff9d",borderRadius:1,marginTop:1}}/>}
          </button>
        ))}
      </nav>

      {/* ── ADD YEAR MODAL ── */}
      {showAddYear && (
        <div className="modal-bg" onClick={()=>setShowAddYear(false)}>
          <div className="modal fade" onClick={e=>e.stopPropagation()}>
            <div style={{fontSize:11,fontWeight:700,letterSpacing:2,color:"#00ff9d",marginBottom:14}}>
              TAMBAH TAHUN BARU
            </div>
            <div className="lbl" style={{marginBottom:6}}>Masukkan Tahun</div>
            <input type="number" className="inp" placeholder="contoh: 2026"
              value={newYearInput} onChange={e=>setNewYearInput(e.target.value)}
              onKeyDown={e=>e.key==="Enter"&&addYear()}
              style={{marginBottom:12}} autoFocus/>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
              <button className="btn-ghost" onClick={()=>{setShowAddYear(false);setNewYearInput("");}}>Batal</button>
              <button className="btn-g" onClick={addYear}>Tambah</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

/* ─── sub-components ─── */
function EmptyState({year}) {
  return (
    <div style={{background:"#0e1520",border:"1px solid #1a2535",borderRadius:12,
      padding:"36px 20px",textAlign:"center",marginTop:8}}>
      <div style={{fontSize:40,marginBottom:10}}>🏃</div>
      <div style={{fontFamily:"'Barlow Condensed'",fontWeight:800,fontSize:20,marginBottom:6}}>
        Belum ada sesi di {year}
      </div>
      <div style={{fontSize:13,color:"#445",lineHeight:1.6}}>
        Tekan <span style={{color:"#00ff9d",fontWeight:700}}>+</span> di atas<br/>untuk mulai catat larimu
      </div>
    </div>
  );
}

function CmpCard({m,color,label,year}) {
  return (
    <div style={{background:"#0e1520",border:`1px solid ${color}33`,borderRadius:12,
      padding:"14px 12px",textAlign:"center"}}>
      <div style={{fontSize:9,fontWeight:700,letterSpacing:3,color,marginBottom:2}}>BULAN {label}</div>
      <div style={{fontFamily:"'Barlow Condensed'",fontWeight:800,fontSize:13,color:"#889",marginBottom:4}}>
        {m.name} {year}
      </div>
      <div style={{fontFamily:"'Barlow Condensed'",fontWeight:800,fontSize:30,color,lineHeight:1}}>
        {m.dist}<span style={{fontSize:13,color:"#556"}}> km</span>
      </div>
      <div style={{fontSize:11,color:"#445",marginTop:4}}>{m.count} sesi</div>
      {m.pace>0&&<div style={{fontSize:11,color:"#ffb347",marginTop:2}}>{(() => {
        const min=Math.floor(m.pace),sec=Math.round((m.pace-min)*60);
        return `${min}:${sec.toString().padStart(2,"0")} /km`;
      })()}</div>}
    </div>
  );
}

function HeatmapCalendar({sessions,year}) {
  const map = useMemo(()=>{
    const m={};
    sessions.forEach(s=>{m[s.date]=(m[s.date]||0)+s.distance});
    return m;
  },[sessions]);

  const cells = useMemo(()=>{
    const arr=[];
    for(let mo=0;mo<12;mo++){
      const days=new Date(year,mo+1,0).getDate();
      for(let d=1;d<=days;d++){
        const dt=`${year}-${String(mo+1).padStart(2,"0")}-${String(d).padStart(2,"0")}`;
        arr.push({date:dt,dist:map[dt]||0,dow:new Date(dt).getDay()});
      }
    }
    return arr;
  },[map,year]);

  const weeks=useMemo(()=>{
    const w=[];let wk=[];
    cells.forEach((c,i)=>{
      if(i===0){for(let j=0;j<c.dow;j++)wk.push(null);}
      wk.push(c);
      if(c.dow===6){w.push(wk);wk=[];}
    });
    if(wk.length)w.push(wk);
    return w;
  },[cells]);

  const color=(d)=>!d?"#1a2535":d<5?"#00ff9d2a":d<10?"#00ff9d66":d<15?"#00ff9daa":"#00ff9d";

  return (
    <div style={{overflowX:"auto"}}>
      <div style={{display:"flex",gap:2,paddingBottom:4,minWidth:"max-content"}}>
        {weeks.map((wk,wi)=>(
          <div key={wi} style={{display:"flex",flexDirection:"column",gap:2}}>
            {Array.from({length:7}).map((_,di)=>{
              const c=wk[di];
              return c?(
                <div key={di} className="hcell" style={{background:color(c.dist)}}
                  title={`${c.date}${c.dist?" — "+c.dist+"km":""}`}/>
              ):<div key={di} style={{width:11,height:11}}/>;
            })}
          </div>
        ))}
      </div>
      <div style={{display:"flex",alignItems:"center",gap:5,marginTop:8}}>
        <span style={{fontSize:10,color:"#445"}}>Tidak lari</span>
        {[0.16,0.4,0.67,1].map((o,i)=>(
          <div key={i} style={{width:10,height:10,borderRadius:2,background:`rgba(0,255,157,${o})`}}/>
        ))}
        <span style={{fontSize:10,color:"#445"}}>15km+</span>
      </div>
    </div>
  );
}
