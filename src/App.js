import React, { useState, useEffect } from 'react';
import { 
  Users, Calendar, ClipboardList, FileDown, 
  Upload, Play, Trash2, Printer, Plus, AlertCircle 
} from 'lucide-react';
import * as XLSX from 'xlsx';

const RosterApp = () => {
  // --- STATE MANAGEMENT ---
  const [activeTab, setActiveTab] = useState('staff');
  const [staff, setStaff] = useState(() => {
    const saved = localStorage.getItem('roster_staff');
    return saved ? JSON.parse(saved) : [];
  });
  const [holidays, setHolidays] = useState(() => {
    const saved = localStorage.getItem('roster_holidays');
    return saved ? JSON.parse(saved) : [];
  });
  const [config, setConfig] = useState({ 
    start: '2026-04-01', 
    end: '2026-04-30' 
  });
  const [generatedRoster, setGeneratedRoster] = useState(null);

  // Sync with LocalStorage
  useEffect(() => {
    localStorage.setItem('roster_staff', JSON.stringify(staff));
    localStorage.setItem('roster_holidays', JSON.stringify(holidays));
  }, [staff, holidays]);

  // --- EXCEL EXPORT (The "Live Template") ---
  const exportCurrentData = () => {
    const wb = XLSX.utils.book_new();
    
    // Create Staff Sheet
    const staffWS = XLSX.utils.json_to_sheet(staff.map(s => ({
      Name: s.name,
      Designation: s.designation,
      Role: s.role || 'Junior'
    })));
    
    // Create Holiday Sheet
    const holidayWS = XLSX.utils.json_to_sheet(holidays.map(h => ({
      Date: h.date,
      Name: h.name
    })));

    XLSX.utils.book_append_sheet(wb, staffWS, "Staff");
    XLSX.utils.book_append_sheet(wb, holidayWS, "Holiday");
    XLSX.writeFile(wb, "Roster_Template_System.xlsx");
  };

  // --- EXCEL IMPORT (The Fixed Logic) ---
  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });

      // Process Staff
      if (wb.SheetNames.includes('Staff')) {
        const data = XLSX.utils.sheet_to_json(wb.Sheets['Staff']);
        setStaff(data.map(item => ({ 
            name: item.Name || item.name || 'Unknown', 
            designation: item.Designation || item.designation || 'Staff',
            role: item.Role || item.role || 'Junior' 
        })));
      }

      // Process Holiday (Fixed Mapping)
      if (wb.SheetNames.includes('Holiday')) {
        const data = XLSX.utils.sheet_to_json(wb.Sheets['Holiday']);
        setHolidays(data.map(item => ({
            date: item.Date || item.date || item.DATE || '', 
            name: item.Name || item.name || item['Holiday Name'] || 'Holiday'
        })).filter(h => h.date));
      }
      alert("Data Imported Successfully!");
    };
    reader.readAsBinaryString(file);
  };

  // --- PRINT FUNCTION ---
  const handlePrint = () => {
    window.print();
  };

  return (
    <div className="max-w-6xl mx-auto p-4 md:p-8 bg-slate-50 min-h-screen font-sans">
      
      {/* HEADER SECTION - Hidden during printing */}
      <header className="flex flex-wrap justify-between items-center mb-6 gap-4 print:hidden">
        <div>
          <h1 className="text-2xl font-bold text-indigo-900">Duty Roster Pro</h1>
          <p className="text-slate-500 text-sm">Managing 26 Staff Members</p>
        </div>
        
        <div className="flex flex-wrap gap-2">
          <button onClick={exportCurrentData} className="bg-white border border-slate-200 p-2 px-3 rounded text-xs font-bold flex items-center gap-2 hover:bg-slate-100 transition shadow-sm">
            <FileDown size={14}/> Download Template
          </button>
          
          <label className="bg-white border border-slate-200 p-2 px-3 rounded text-xs font-bold flex items-center gap-2 cursor-pointer hover:bg-slate-100 transition shadow-sm">
            <Upload size={14}/> Upload Excel
            <input type="file" className="hidden" onChange={handleFileUpload} accept=".xlsx, .xls" />
          </label>
          
          <button onClick={handlePrint} className="bg-slate-800 text-white p-2 px-3 rounded text-xs font-bold flex items-center gap-2 hover:bg-slate-700 transition shadow-sm">
            <Printer size={14}/> Print to PDF
          </button>
        </div>
      </header>

      {/* NAVIGATION TABS - Hidden during printing */}
      <nav className="flex border-b border-slate-200 mb-6 overflow-x-auto print:hidden">
        <button 
          onClick={() => setActiveTab('staff')}
          className={`p-3 px-6 text-sm font-medium border-b-2 transition ${activeTab === 'staff' ? 'border-indigo-600 text-indigo-600' : 'border-transparent text-slate-500'}`}
        >
          <div className="flex items-center gap-2"><Users size={16}/> Staff ({staff.length})</div>
        </button>
        <button 
          onClick={() => setActiveTab('holidays')}
          className={`p-3 px-6 text-sm font-medium border-b-2 transition ${activeTab === 'holidays' ? 'border-indigo-600 text-indigo-600' : 'border-transparent text-slate-500'}`}
        >
          <div className="flex items-center gap-2"><Calendar size={16}/> Holidays ({holidays.length})</div>
        </button>
        <button 
          onClick={() => setActiveTab('roster')}
          className={`p-3 px-6 text-sm font-medium border-b-2 transition ${activeTab === 'roster' ? 'border-indigo-600 text-indigo-600' : 'border-transparent text-slate-500'}`}
        >
          <div className="flex items-center gap-2"><ClipboardList size={16}/> View Roster</div>
        </button>
      </nav>

      {/* CONTENT AREA */}
      <main className="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
        
        {/* STAFF TAB */}
        {activeTab === 'staff' && (
          <div className="print:hidden">
            <div className="flex justify-between items-center mb-4">
              <h2 className="text-lg font-bold">Staff Directory</h2>
              <button onClick={() => setStaff([...staff, { name: '', designation: '', role: 'Junior' }])} className="text-indigo-600 text-sm font-bold flex items-center gap-1">
                <Plus size={16}/> Add Row
              </button>
            </div>
            <table className="w-full text-left text-sm">
              <thead>
                <tr className="border-b bg-slate-50">
                  <th className="p-3">Name</th>
                  <th className="p-3">Designation</th>
                  <th className="p-3">Role</th>
                  <th className="p-3 text-right">Action</th>
                </tr>
              </thead>
              <tbody>
                {staff.map((s, idx) => (
                  <tr key={idx} className="border-b hover:bg-slate-50 transition">
                    <td className="p-2"><input className="w-full p-1 border rounded" value={s.name} onChange={(e) => {
                      const n = [...staff]; n[idx].name = e.target.value; setStaff(n);
                    }}/></td>
                    <td className="p-2"><input className="w-full p-1 border rounded" value={s.designation} onChange={(e) => {
                      const n = [...staff]; n[idx].designation = e.target.value; setStaff(n);
                    }}/></td>
                    <td className="p-2">
                      <select className="w-full p-1 border rounded" value={s.role} onChange={(e) => {
                        const n = [...staff]; n[idx].role = e.target.value; setStaff(n);
                      }}>
                        <option value="Senior">Senior</option>
                        <option value="Junior">Junior</option>
                      </select>
                    </td>
                    <td className="p-2 text-right">
                      <button onClick={() => setStaff(staff.filter((_, i) => i !== idx))} className="text-red-400 hover:text-red-600"><Trash2 size={16}/></button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {/* HOLIDAY TAB */}
        {activeTab === 'holidays' && (
          <div className="print:hidden">
            <h2 className="text-lg font-bold mb-4">Holiday List (Imported from Excel)</h2>
            {holidays.length === 0 ? (
              <div className="text-center p-10 border-2 border-dashed rounded-lg text-slate-400">
                <AlertCircle className="mx-auto mb-2" size={32}/>
                <p>No holidays found. Use "Upload Excel" to import your holiday list.</p>
              </div>
            ) : (
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                {holidays.map((h, idx) => (
                  <div key={idx} className="p-3 border rounded-lg bg-indigo-50 flex justify-between items-center">
                    <div>
                      <p className="font-bold text-indigo-900">{h.name}</p>
                      <p className="text-xs text-slate-500">{h.date}</p>
                    </div>
                    <button onClick={() => setHolidays(holidays.filter((_, i) => i !== idx))} className="text-slate-400 hover:text-red-500"><Trash2 size={14}/></button>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}

        {/* ROSTER TAB */}
        {activeTab === 'roster' && (
          <div>
            <div className="flex justify-between items-center mb-6 print:hidden">
              <h2 className="text-lg font-bold">Monthly Roster</h2>
              <button onClick={() => setGeneratedRoster(true)} className="bg-indigo-600 text-white px-6 py-2 rounded-lg font-bold flex items-center gap-2 hover:bg-indigo-700 shadow-md">
                <Play size={18}/> Generate Schedule
              </button>
            </div>
            
            {generatedRoster ? (
              <div className="overflow-x-auto">
                <div className="text-center mb-6 hidden print:block">
                  <h1 className="text-2xl font-bold">Duty Roster - {config.start} to {config.end}</h1>
                  <p>Generated on {new Date().toLocaleDateString()}</p>
                </div>
                <table className="w-full border-collapse border border-slate-300 text-xs">
                  <thead>
                    <tr className="bg-slate-100">
                      <th className="border p-2">Staff Name</th>
                      <th className="border p-2">Role</th>
                      {/* Example columns for the first 7 days */}
                      {[...Array(7)].map((_, i) => (
                        <th key={i} className="border p-2">Day {i + 1}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {staff.map((s, idx) => (
                      <tr key={idx}>
                        <td className="border p-2 font-bold">{s.name}</td>
                        <td className="border p-2 text-slate-500">{s.role}</td>
                        {[...Array(7)].map((_, i) => (
                          <td key={i} className="border p-2 text-center">G</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : (
              <div className="text-center py-20 text-slate-400">
                Click "Generate Schedule" to create the final view.
              </div>
            )}
          </div>
        )}
      </main>
    </div>
  );
};

export default RosterApp;