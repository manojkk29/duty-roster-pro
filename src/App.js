import React, { useState, useEffect, useMemo } from 'react';
import { Users, Settings, PlaneTakeoff, Play, Plus, Trash2, Printer, Search, Upload, FileDown, Calendar, ClipboardList, BarChart3, RefreshCcw, AlertCircle } from 'lucide-react';
import * as XLSX from 'xlsx';

const RosterApp = () => {
  const [activeTab, setActiveTab] = useState('staff');
  const [searchTerm, setSearchTerm] = useState('');
  
  const [staff, setStaff] = useState(() => {
    const saved = localStorage.getItem('roster_staff');
    return saved ? JSON.parse(saved) : [{ name: 'Sample Staff', designation: 'AI Lead', role: 'Senior' }];
  });

  const [holidays, setHolidays] = useState(() => {
    const saved = localStorage.getItem('roster_holidays');
    return saved ? JSON.parse(saved) : [{ date: '2026-05-01', name: 'Maharashtra Day' }];
  });

  const [offDays, setOffDays] = useState([]);
  const [config, setConfig] = useState({ start: '2026-04-01', end: '2026-04-30' });
  const [generatedRoster, setGeneratedRoster] = useState(null);

  useEffect(() => {
    localStorage.setItem('roster_staff', JSON.stringify(staff));
    localStorage.setItem('roster_holidays', JSON.stringify(holidays));
  }, [staff, holidays]);

  const downloadTemplate = () => {
    const wb = XLSX.utils.book_new();
    const wsStaff = XLSX.utils.json_to_sheet([{ Name: 'John Doe', Designation: 'AI Specialist', Role: 'Senior' }]);
    const wsHolidays = XLSX.utils.json_to_sheet([{ Date: '2026-05-01', Name: 'Holiday Name' }]);
    XLSX.utils.book_append_sheet(wb, wsStaff, "Staff");
    XLSX.utils.book_append_sheet(wb, wsHolidays, "Holiday");
    XLSX.writeFile(wb, "Roster_Template.xlsx");
  };

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      if (wb.SheetNames.includes('Staff')) {
        const data = XLSX.utils.sheet_to_json(wb.Sheets['Staff']);
        setStaff(data.map(item => ({ 
            name: item.Name || item.name || '', 
            designation: item.Designation || item.designation || '',
            role: item.Role || item.role || 'Junior' 
        })));
      }
    };
    reader.readAsBinaryString(file);
  };

  const shiftStats = useMemo(() => {
    if (!generatedRoster) return [];
    const stats = {};
    staff.forEach(s => stats[s.name] = { morning: 0, evening: 0, night: 0, total: 0, role: s.role, designation: s.designation });
    generatedRoster.forEach(row => {
      ['morning', 'evening', 'night'].forEach(s => { if (row[s] !== 'OFF' && stats[row[s]]) { stats[row[s]][s]++; stats[row[s]].total++; } });
    });
    return Object.entries(stats).map(([name, data]) => ({ name, ...data }));
  }, [generatedRoster, staff]);

  const generateRoster = () => {
    let roster = [];
    let curr = new Date(config.start);
    const end = new Date(config.end);
    let jIdx = 0, sIdx = 0;
    const juniors = staff.filter(s => s.role === 'Junior').map(s => s.name);
    const seniors = staff.filter(s => s.role === 'Senior').map(s => s.name);
    
    const offLookup = {};
    offDays.forEach(o => {
        if(o.date && o.name) {
            const d = o.date;
            if(o.m) offLookup[`${d}-${o.name}-morning`] = true;
            if(o.e) offLookup[`${d}-${o.name}-evening`] = true;
            if(o.n) offLookup[`${d}-${o.name}-night`] = true;
        }
    });

    while (curr <= end) {
      const dStr = curr.toISOString().split('T')[0];
      const isExtra = holidays.some(h => h.date === dStr) || curr.getDay() === 0;
      let row = { date: dStr, day: curr.toLocaleDateString('en-US', { weekday: 'short' }), morning: 'OFF', evening: 'OFF', night: 'OFF' };
      
      const getAvailable = (list, idx, shiftType) => {
          if(!list.length) return { name: 'OFF', newIdx: idx };
          for(let i=0; i < list.length; i++) {
              let p = list[(idx + i) % list.length];
              if(!offLookup[`${dStr}-${p}-${shiftType}`]) return { name: p, newIdx: idx + i + 1 };
          }
          return { name: 'SHORTAGE', newIdx: idx };
      };

      let nRes = getAvailable(juniors, jIdx, 'night'); row.night = nRes.name; jIdx = nRes.newIdx;
      let eRes = getAvailable(seniors, sIdx, 'evening'); row.evening = eRes.name; sIdx = eRes.newIdx;
      if (isExtra) {
          let mRes = getAvailable(seniors, sIdx, 'morning'); row.morning = mRes.name; sIdx = mRes.newIdx;
      }
      roster.push(row);
      curr.setDate(curr.getDate() + 1);
    }
    setGeneratedRoster(roster);
    setActiveTab('roster');
  };

  const addLeaveEntry = () => {
    setOffDays([...offDays, { name: '', date: '', m: true, e: true, n: true }]);
  };

  const updateLeave = (index, field, value) => {
    const updated = [...offDays];
    const entry = { ...updated[index], [field]: value };
    
    // Validation: Check for duplicates when name or date changes
    if (field === 'name' || field === 'date') {
        const isDuplicate = offDays.some((o, i) => i !== index && o.name === (field === 'name' ? value : o.name) && o.date === (field === 'date' ? value : o.date) && o.name !== '' && o.date !== '');
        if (isDuplicate) {
            alert("This staff member already has a leave entry for this date.");
            return;
        }
    }
    
    updated[index] = entry;
    setOffDays(updated);
  };

  return (
    <div className="max-w-6xl mx-auto p-4 md:p-8 bg-slate-50 min-h-screen font-sans">
      <header className="flex justify-between items-center mb-6 no-print">
        <h1 className="text-2xl font-bold text-indigo-900">Duty Roster Pro</h1>
        <div className="flex gap-2">
          <button onClick={downloadTemplate} className="bg-slate-100 p-2 rounded text-xs font-bold flex items-center gap-1"><FileDown size={14}/> Template</button>
          <label className="bg-white border p-2 rounded text-xs font-bold flex items-center gap-1 cursor-pointer"><Upload size={14}/> Import<input type="file" className="hidden" onChange={handleFileUpload} /></label>
          <button onClick={generateRoster} className="bg-indigo-600 text-white p-2 px-4 rounded text-xs font-bold flex items-center gap-1"><Play size={14}/> Generate</button>
        </div>
      </header>

      <div className="bg-white rounded-xl shadow-sm border overflow-hidden">
        <nav className="flex bg-slate-100 border-b no-print">
          <TabBtn active={activeTab === 'staff'} onClick={() => setActiveTab('staff')} label="Staff" icon={<Users size={16}/>}/>
          <TabBtn active={activeTab === 'holidays'} onClick={() => setActiveTab('holidays')} label="Holidays" icon={<Calendar size={16}/>}/>
          <TabBtn active={activeTab === 'off'} onClick={() => setActiveTab('off')} label="Leave" icon={<PlaneTakeoff size={16}/>}/>
          <TabBtn active={activeTab === 'config'} onClick={() => setActiveTab('config')} label="Dates" icon={<Settings size={16}/>}/>
          {generatedRoster && <TabBtn active={activeTab === 'roster'} onClick={() => setActiveTab('roster')} label="Roster" icon={<ClipboardList size={16}/>} highlight/>}
          {generatedRoster && <TabBtn active={activeTab === 'summary'} onClick={() => setActiveTab('summary')} label="Summary" icon={<BarChart3 size={16}/>} highlight/>}
        </nav>

        <div className="p-4">
          {activeTab === 'off' && (
            <div className="max-h-96 overflow-auto border rounded">
              <table className="w-full text-left text-xs border-collapse">
                <thead className="bg-slate-50 sticky top-0 border-b">
                  <tr>
                    <th className="p-2 w-10 text-center">SN</th>
                    <th className="p-2 w-48">Staff Name</th>
                    <th className="p-2 w-40">Date</th>
                    <th className="p-2">Shifts to mark as OFF</th>
                    <th className="p-2 w-10"></th>
                  </tr>
                </thead>
                <tbody>
                  {offDays.map((o, i) => (
                    <tr key={i} className="border-b bg-white">
                      <td className="p-1 text-center text-slate-400 font-mono">{i + 1}</td>
                      <td className="p-1">
                        <select className="border w-full p-1 rounded font-bold" value={o.name} onChange={e => updateLeave(i, 'name', e.target.value)}>
                          <option value="">Select Staff</option>
                          {staff.map(s => <option key={s.name} value={s.name}>{s.name}</option>)}
                        </select>
                      </td>
                      <td className="p-1"><input type="date" className="border w-full p-1 rounded font-bold" value={o.date} onChange={e => updateLeave(i, 'date', e.target.value)} /></td>
                      <td className="p-1">
                        <div className="flex gap-4 px-2 py-1 text-[10px] font-black">
                          <label className="flex items-center gap-1 cursor-pointer"><input type="checkbox" checked={o.m} onChange={e => updateLeave(i, 'm', e.target.checked)}/> MOR</label>
                          <label className="flex items-center gap-1 cursor-pointer"><input type="checkbox" checked={o.e} onChange={e => updateLeave(i, 'e', e.target.checked)}/> EVE</label>
                          <label className="flex items-center gap-1 cursor-pointer"><input type="checkbox" checked={o.n} onChange={e => updateLeave(i, 'n', e.target.checked)}/> NIT</label>
                        </div>
                      </td>
                      <td className="p-1 text-center"><button onClick={() => setOffDays(offDays.filter((_, idx) => idx !== i))} className="text-rose-500 hover:bg-rose-50 p-1 rounded-full"><Trash2 size={14}/></button></td>
                    </tr>
                  ))}
                </tbody>
              </table>
              <button onClick={addLeaveEntry} className="m-2 text-indigo-600 font-bold text-xs flex items-center gap-1"><Plus size={14}/> Add Leave Entry</button>
            </div>
          )}

          {activeTab === 'staff' && (
            <div className="max-h-96 overflow-auto border rounded">
              <table className="w-full text-left text-xs border-collapse">
                <thead className="bg-slate-50 sticky top-0 border-b">
                  <tr><th className="p-2 w-10 text-center">SN</th><th className="p-2">Name</th><th className="p-2">Designation</th><th className="p-2 w-24">Role</th><th className="p-2 w-10"></th></tr>
                </thead>
                <tbody>
                  {staff.map((s, i) => (
                    <tr key={i} className="border-b bg-white">
                      <td className="p-1 text-center text-slate-400 font-mono">{i + 1}</td>
                      <td className="p-1"><input className="border w-full p-1 rounded" value={s.name} onChange={e => {const n = [...staff]; n[i].name = e.target.value; setStaff(n)}}/></td>
                      <td className="p-1"><input className="border w-full p-1 rounded" value={s.designation} onChange={e => {const n = [...staff]; n[i].designation = e.target.value; setStaff(n)}}/></td>
                      <td className="p-1"><select className="border w-full p-1 rounded font-bold" value={s.role} onChange={e => {const n = [...staff]; n[i].role = e.target.value; setStaff(n)}}><option>Senior</option><option>Junior</option></select></td>
                      <td className="p-1 text-center"><button onClick={() => setStaff(staff.filter((_, idx) => idx !== i))} className="text-rose-500"><Trash2 size={14}/></button></td>
                    </tr>
                  ))}
                </tbody>
              </table>
              <button onClick={() => setStaff([...staff, {name: '', designation: '', role: 'Junior'}])} className="m-2 text-indigo-600 font-bold text-xs flex items-center gap-1"><Plus size={14}/> Add New Staff</button>
            </div>
          )}

          {activeTab === 'summary' && (
            <div className="border rounded overflow-hidden">
              <table className="w-full text-left text-[11px]" style={{tableLayout:'fixed'}}>
                <thead className="bg-slate-900 text-white font-black uppercase text-[9px]">
                  <tr>
                    <th style={{padding: '8px', width: '6%'}} className="text-center">SN</th>
                    <th style={{padding: '8px', width: '22%'}}>Name</th>
                    <th style={{padding: '8px', width: '15%'}}>Desig</th>
                    <th style={{padding: '8px', textAlign: 'center', width: '7%'}}>R</th>
                    <th style={{padding: '8px', textAlign: 'center', width: '10%'}}>M</th>
                    <th style={{padding: '8px', textAlign: 'center', width: '10%'}}>E</th>
                    <th style={{padding: '8px', textAlign: 'center', width: '10%'}}>N</th>
                    <th style={{padding: '8px', textAlign: 'center', width: '20%', backgroundColor: '#312e81'}}>Total</th>
                  </tr>
                </thead>
                <tbody className="divide-y">
                  {shiftStats.sort((a,b) => b.total - a.total).map((s, i) => (
                    <tr key={i} style={{backgroundColor: i % 2 === 0 ? '#fff' : '#fafafa'}}>
                      <td style={{padding: '5px 8px', textAlign: 'center', color: '#94a3b8'}} className="font-mono">{i + 1}</td>
                      <td style={{padding: '5px 8px', fontWeight: 'bold'}}>{s.name}</td>
                      <td style={{padding: '5px 8px', color: '#64748b', fontSize: '10px'}}>{s.designation}</td>
                      <td style={{padding: '5px 8px', textAlign: 'center'}}><span style={{backgroundColor: s.role === 'Senior' ? '#e0e7ff' : '#f1f5f9', padding: '1px 4px', borderRadius: '3px'}}>{s.role[0]}</span></td>
                      <td style={{padding: '5px 8px', textAlign: 'center'}}>{s.morning}</td>
                      <td style={{padding: '5px 8px', textAlign: 'center'}}>{s.evening}</td>
                      <td style={{padding: '5px 8px', textAlign: 'center'}}>{s.night}</td>
                      <td style={{padding: '5px 8px', textAlign: 'center', fontWeight: '900', backgroundColor: '#eef2ff', color: '#4338ca'}}>{s.total}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {activeTab === 'holidays' && (
            <div className="max-h-96 overflow-auto border rounded">
              <table className="w-full text-left text-xs border-collapse">
                <thead><tr className="bg-slate-50 border-b"><th className="p-2 w-10 text-center">SN</th><th className="p-2 w-32">Date</th><th className="p-2">Holiday Name</th><th className="p-2 w-10"></th></tr></thead>
                <tbody>{holidays.map((h, i) => (<tr key={i} className="border-b"><td className="p-1 text-center text-slate-400 font-mono">{i + 1}</td><td className="p-1"><input type="date" className="border p-1 rounded" value={h.date} onChange={e => {const n = [...holidays]; n[i].date = e.target.value; setHolidays(n)}}/></td><td className="p-1"><input className="border w-full p-1 rounded" value={h.name} onChange={e => {const n = [...holidays]; n[i].name = e.target.value; setHolidays(n)}}/></td><td className="p-1 text-center"><button onClick={() => setHolidays(holidays.filter((_, idx) => idx !== i))} className="text-rose-500"><Trash2 size={14}/></button></td></tr>))}</tbody>
              </table>
              <button onClick={() => setHolidays([...holidays, {date: '', name: ''}])} className="m-2 text-indigo-600 font-bold text-xs flex items-center gap-1"><Plus size={14}/> Add Holiday</button>
            </div>
          )}

          {activeTab === 'config' && (
            <div className="flex flex-col gap-4">
                <div className="flex gap-4 p-4 bg-indigo-50 border border-indigo-100 rounded w-fit">
                    <div><label className="text-[10px] font-bold block mb-1">Start Date</label><input type="date" className="border p-1 rounded font-bold" value={config.start} onChange={e => setConfig({...config, start: e.target.value})}/></div>
                    <div><label className="text-[10px] font-bold block mb-1">End Date</label><input type="date" className="border p-1 rounded font-bold" value={config.end} onChange={e => setConfig({...config, end: e.target.value})}/></div>
                </div>
                <button onClick={() => {if(window.confirm("Clear all data?")) {setStaff([]); setHolidays([]); setOffDays([]); setGeneratedRoster(null); localStorage.clear();}}} className="flex items-center gap-2 text-rose-600 text-xs font-bold bg-rose-50 p-3 rounded border border-rose-100 w-fit hover:bg-rose-100"><RefreshCcw size={14}/> Reset All Local Data</button>
            </div>
          )}

          {activeTab === 'roster' && generatedRoster && (
            <div className="border rounded overflow-hidden">
              <table className="w-full text-left text-xs border-collapse">
                <thead className="bg-slate-900 text-white font-bold sticky top-0">
                  <tr><th className="p-2 border">Date</th><th className="p-2 border">Morning (S)</th><th className="p-2 border">Evening (S)</th><th className="p-2 border">Night (J)</th></tr>
                </thead>
                <tbody>{generatedRoster.map((r, i) => (<tr key={i} className={i % 2 === 0 ? 'bg-white' : 'bg-slate-50'}><td className="p-2 border font-bold">{r.date} ({r.day})</td><td className={`p-2 border font-bold ${r.morning === 'OFF' ? 'text-slate-200' : 'text-indigo-600'}`}>{r.morning}</td><td className="p-2 border text-emerald-600 font-bold">{r.evening}</td><td className="p-2 border font-black">{r.night}</td></tr>))}</tbody>
              </table>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

const TabBtn = ({ active, onClick, icon, label, highlight }) => (
  <button onClick={onClick} className={`px-4 py-3 flex items-center gap-2 font-bold transition-all border-b-2 text-xs ${active ? 'bg-white border-indigo-600 text-indigo-600 shadow-[inset_0_-2px_0_rgba(79,70,229,1)]' : 'text-slate-400 border-transparent hover:bg-slate-50'} ${highlight && !active ? 'text-indigo-500' : ''}`}>
    {icon} {label}
  </button>
);

export default RosterApp;