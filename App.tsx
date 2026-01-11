import React, { useState, useEffect, useRef } from 'react';

import { 
  Plus, 
  Camera, 
  ChevronRight, 
  LogOut, 
  X, 
  CheckCircle2, 
  Clock, 
  BarChart3,
  ClipboardList,
  FileText,
  Layers,
  Image as ImageIcon,
  ChevronLeft,
  Cloud,
  RefreshCw,
  Settings,
  Database,
  ChevronDown,
  ChevronUp,
  AlertTriangle,
  Copy,
  Check,
  Maximize2,
  Download,
  Users,
  TrendingUp,
  Table2,
  Grid
} from 'lucide-react';
// @ts-ignore
import ExcelJS from 'exceljs';
import { Defect, UserSession, AppRole, ZoneId, ZoneResponse, ManpowerEntry, ManagerAnalysis, TsprEntry } from './types';
import { ZONES, ROLE_CONFIG } from './constants';

const LOCAL_STORAGE_KEY = 'autoline_defects_v1';
const CLOUD_URL_KEY = 'autoline_cloud_url_v1';
const TSPR_STORAGE_KEY = 'autoline_tspr_v1';
const SESSION_KEY = 'autoline_session_v1';

const TIME_SLOTS = [
  '08:00 - 09:00',
  '09:00 - 10:00',
  '10:00 - 11:45',
  '12:00 - 13:00',
  '13:00 - 14:00',
  '14:00 - 15:00',
  '15:00 - 16:45'
];

const INITIAL_TSPR_DATA: TsprEntry[] = TIME_SLOTS.map(slot => ({ slot, total: 0, ng: 0 }));

// --- HELPER FUNCTIONS ---

const processImageForExcel = (base64Url: string): Promise<{ base64: string, extension: 'png' | 'jpeg' | 'gif', width: number, height: number }> => {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.crossOrigin = 'Anonymous';
    img.onload = () => {
      const mime = base64Url.match(/data:image\/(\w+);base64/)?.[1]?.toLowerCase();
      if (mime === 'png' || mime === 'jpeg' || mime === 'jpg' || mime === 'gif') {
         resolve({ 
           base64: base64Url.split(',')[1], 
           extension: mime === 'jpg' ? 'jpeg' : mime as 'png' | 'jpeg' | 'gif',
           width: img.width,
           height: img.height
         });
      } else {
         const canvas = document.createElement('canvas');
         canvas.width = img.width;
         canvas.height = img.height;
         const ctx = canvas.getContext('2d');
         if(ctx) {
            ctx.drawImage(img, 0, 0);
            const newDataUrl = canvas.toDataURL('image/png');
            resolve({
               base64: newDataUrl.split(',')[1],
               extension: 'png',
               width: img.width,
               height: img.height
            });
         } else {
            resolve({ 
               base64: base64Url.split(',')[1], 
               extension: 'png',
               width: img.width,
               height: img.height
            }); 
         }
      }
    };
    img.onerror = () => {
        resolve({ base64: base64Url.split(',')[1], extension: 'png', width: 100, height: 100 });
    };
    img.src = base64Url;
  });
};

const downloadReport = async (defects: Defect[]) => {
  if (defects.length === 0) {
    alert("No data to export.");
    return;
  }
  
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Inspection Report');

  worksheet.columns = [
    { header: 'Report ID', key: 'id', width: 12 },
    { header: 'Date', key: 'date', width: 12 },
    { header: 'Time', key: 'time', width: 12 },
    { header: 'Vehicle No', key: 'frameNo', width: 18 },
    { header: 'Model', key: 'model', width: 12 },
    { header: 'Defect Type', key: 'defectType', width: 20 },
    { header: 'Defect Details', key: 'defectDetails', width: 35 },
    { header: 'Targeted Zones', key: 'targetedZones', width: 18 },
    { header: 'Zone Analysis & Findings', key: 'analysis', width: 50 },
    { header: '4M Manager Analysis', key: 'manager', width: 40 },
    { header: 'Status', key: 'status', width: 12 },
    { header: 'Visual Evidence', key: 'image', width: 65 },
  ];

  const headerRow = worksheet.getRow(1);
  headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12 };
  headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F46E5' } };
  headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
  headerRow.height = 30;

  for (let i = 0; i < defects.length; i++) {
    const d = defects[i];
    const dateObj = new Date(d.timestamp);
    const m = d.managerAnalysis || { man: '', machine: '', material: '', method: '', rootCauseSummary: '', status: 'OPEN' };

    let analysisText = '';
    const allZoneIds = Array.from(new Set([...d.targetedZones, ...Object.keys(d.responses)]));
    
    allZoneIds.forEach(z => {
       const r = d.responses[z];
       if (!r) return;
       analysisText += `[Zone ${z}]: ${r.involved}\n`;
       if (r.involved === 'YES') {
          analysisText += `Reason: ${r.reason || 'N/A'}\nAction: ${r.actionTaken || 'N/A'}\n`;
          if (r.manpower && r.manpower.length > 0) {
             const mpStr = r.manpower.map(m => `${m.name} (${m.ein})`).join(', ');
             analysisText += `Manpower: ${mpStr}\n`;
          }
       }
       analysisText += '\n';
    });

    let managerText = '';
    if (m.man) managerText += `Man: ${m.man}\n`;
    if (m.machine) managerText += `Machine: ${m.machine}\n`;
    if (m.material) managerText += `Material: ${m.material}\n`;
    if (m.method) managerText += `Method: ${m.method}\n`;
    if (m.rootCauseSummary) managerText += `\nSummary: ${m.rootCauseSummary}`;

    const row = worksheet.addRow({
      id: d.id,
      date: dateObj.toLocaleDateString(),
      time: dateObj.toLocaleTimeString(),
      frameNo: d.frameNo,
      model: d.modelName,
      defectType: d.defectName,
      defectDetails: d.description,
      targetedZones: d.targetedZones.join(', '),
      analysis: analysisText.trim(),
      manager: managerText.trim(),
      status: m.status
    });

    row.height = 220;
    row.eachCell((cell: any, colNumber: number) => {
      cell.alignment = { vertical: 'middle', wrapText: true, horizontal: 'center' };
      const colKey = worksheet.getColumn(colNumber).key;
      if (['defectDetails', 'analysis', 'manager'].includes(colKey as string)) {
          cell.alignment = { vertical: 'top', wrapText: true, horizontal: 'left' };
      }
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
      };
    });
    
    const statusCell = row.getCell('status');
    statusCell.font = { bold: true, color: { argb: m.status === 'CLOSED' ? 'FF16A34A' : 'FFDC2626' } };

    if (d.photoUrl && d.photoUrl.startsWith('data:image')) {
       try {
         const processedImage = await processImageForExcel(d.photoUrl);
         const imageId = workbook.addImage({ base64: processedImage.base64, extension: processedImage.extension });
         const cellWidthPx = 450;
         const cellHeightPx = 280;
         const cellAspectRatio = cellWidthPx / cellHeightPx;
         const imgAspectRatio = processedImage.width / processedImage.height;
         const PADDING = 0.05;
         let wRatio = 1 - (PADDING * 2);
         let hRatio = 1 - (PADDING * 2);
         if (imgAspectRatio > cellAspectRatio) {
           hRatio = (1 - PADDING * 2) * (cellAspectRatio / imgAspectRatio);
         } else {
           wRatio = (1 - PADDING * 2) * (imgAspectRatio / cellAspectRatio);
         }
         const colOff = PADDING + (1 - PADDING * 2 - wRatio) / 2;
         const rowOff = PADDING + (1 - PADDING * 2 - hRatio) / 2;
         const imageColIndex = 11;
         worksheet.addImage(imageId, {
           tl: { col: imageColIndex + colOff, row: row.number - 1 + rowOff } as any,
           br: { col: imageColIndex + colOff + wRatio, row: row.number - 1 + rowOff + hRatio } as any,
           editAs: 'oneCell' 
         });
       } catch (e) {
         row.getCell('image').value = "[Image Error]";
       }
    } else {
       row.getCell('image').value = "No Image";
    }
  }

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const filename = `AUTOLINE_REPORT_${new Date().toISOString().replace(/[:.]/g, '-')}.xlsx`;
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
};

const SCRIPT_CODE = `function doPost(e) { return handleRequest(e); }
function doGet(e) { return handleRequest(e); }
function handleRequest(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (e.postData && e.postData.contents) {
    var json = JSON.parse(e.postData.contents);
    var defects = json.defects || [];
    var tspr = json.tspr || [];
    var storageSheet = ss.getSheetByName("Storage_DB");
    if (!storageSheet) { storageSheet = ss.insertSheet("Storage_DB"); }
    storageSheet.getRange(1, 1).setValue(JSON.stringify({ defects: defects, tspr: tspr }));
    var sheet = ss.getSheetByName("Defects");
    if (!sheet) { sheet = ss.insertSheet("Defects"); }
    var headers = ['ID', 'Timestamp', 'Frame No', 'Model', 'Defect Name', 'Description', 'Targeted Zones', 'Involved Zones', 'Detailed Findings', 'Manpower', '4M Man', '4M Machine', '4M Material', '4M Method', 'Root Cause Summary', 'Status'];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e0e0e0");
      sheet.setFrozenRows(1);
    }
    var rows = [];
    defects.forEach(function(d) {
      var involved = [], findings = [], manpower = [];
      var allZones = (d.targetedZones || []).concat(Object.keys(d.responses || {})).filter((v, i, a) => a.indexOf(v) === i);
      allZones.forEach(function(z) {
          var r = d.responses ? d.responses[z] : null;
          if (r && r.involved === 'YES') {
            involved.push(z);
            findings.push(z + ": " + (r.reason || '-') + " -> " + (r.actionTaken || '-'));
            if (r.manpower) r.manpower.forEach(function(m) { manpower.push(m.name + " (" + m.ein + ")"); });
          } else if (r && r.involved === 'NO') findings.push(z + ": Not Involved");
          else if (d.targetedZones && d.targetedZones.indexOf(z) > -1) findings.push(z + ": Pending");
      });
      var m = d.managerAnalysis || {};
      rows.push([d.id, new Date(d.timestamp).toLocaleString(), d.frameNo, d.modelName, d.defectName, d.description, d.targetedZones ? d.targetedZones.join(', ') : '', involved.join(', '), findings.join('\\n'), manpower.join('\\n'), m.man || '', m.machine || '', m.material || '', m.method || '', m.rootCauseSummary || '', m.status || 'OPEN']);
    });
    if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn()).clearContent();
    if (rows.length > 0) {
      var range = sheet.getRange(2, 1, rows.length, rows[0].length);
      range.setValues(rows);
      range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      range.setVerticalAlignment("top");
    }
    var tsprSheet = ss.getSheetByName("TSPR");
    if (!tsprSheet) { tsprSheet = ss.insertSheet("TSPR"); }
    var tsprHeaders = ["Time Slot", "Total Vehicles", "NG Vehicles"];
    if (tsprSheet.getLastRow() === 0) { tsprSheet.appendRow(tsprHeaders); tsprSheet.getRange(1, 1, 1, 3).setFontWeight("bold"); }
    if (tsprSheet.getLastRow() > 1) tsprSheet.getRange(2, 1, tsprSheet.getLastRow(), 3).clearContent();
    var tsprRows = tspr.map(function(t) { return [t.slot, t.total, t.ng]; });
    if (tsprRows.length > 0) tsprSheet.getRange(2, 1, tsprRows.length, 3).setValues(tsprRows);
    return ContentService.createTextOutput(JSON.stringify({status: 'success', data: defects, tspr: tspr})).setMimeType(ContentService.MimeType.JSON);
  } else {
    var storageSheet = ss.getSheetByName("Storage_DB");
    if (!storageSheet) return ContentService.createTextOutput(JSON.stringify({status: 'empty', data: [], tspr: []})).setMimeType(ContentService.MimeType.JSON);
    var data = storageSheet.getRange(1, 1).getValue();
    if (!data) return ContentService.createTextOutput(JSON.stringify({status: 'empty', data: [], tspr: []})).setMimeType(ContentService.MimeType.JSON);
    var parsed = JSON.parse(data);
    return ContentService.createTextOutput(JSON.stringify({status: 'success', data: parsed.defects || [], tspr: parsed.tspr || []})).setMimeType(ContentService.MimeType.JSON);
  }
}`;

// --- SUB-COMPONENTS ---

const TsprGraph: React.FC<{ data: TsprEntry[] }> = ({ data }) => {
  const width = 800;
  const height = 300;
  const padding = 40;
  const graphWidth = width - padding * 2;
  const graphHeight = height - padding * 2;

  const points = data.map((d, i) => {
    if (d.total === 0) return null;
    const x = padding + (i / (data.length - 1)) * graphWidth;
    const percentage = ((d.total - d.ng) / d.total) * 100;
    const y = height - padding - (percentage / 100) * graphHeight;
    return { x, y, percentage, slot: d.slot };
  }).filter(p => p !== null) as {x: number, y: number, percentage: number, slot: string}[];

  const pathD = points.length > 0 ? `M ${points.map(p => `${p.x},${p.y}`).join(' L ')}` : '';
  const fillD = points.length > 1 ? `${pathD} L ${points[points.length-1].x},${height-padding} L ${points[0].x},${height-padding} Z` : '';

  return (
    <div className="w-full overflow-x-auto">
      <div className="min-w-[600px] bg-slate-900 rounded-[2rem] p-6 shadow-2xl">
        <h4 className="text-white font-black uppercase tracking-widest text-xs mb-6 flex items-center gap-2">
          <TrendingUp className="text-green-400" size={16}/> TSPR Hourly Trend
        </h4>
        <svg viewBox={`0 0 ${width} ${height}`} className="w-full h-auto">
          {[0, 25, 50, 75, 100].map(tick => {
            const yPos = height - padding - (tick / 100) * graphHeight;
            return (
              <g key={tick}>
                <line x1={padding} y1={yPos} x2={width - padding} y2={yPos} stroke="#334155" strokeWidth="1" strokeDasharray="4 4" />
                <text x={padding - 10} y={yPos + 4} fill="#94a3b8" fontSize="10" textAnchor="end" fontWeight="bold">{tick}%</text>
              </g>
            );
          })}
          {data.map((d, i) => {
              const x = padding + (i / (data.length - 1)) * graphWidth;
              return <text key={i} x={x} y={height - padding + 20} fill="#64748b" fontSize="10" textAnchor="middle" fontWeight="bold" style={{textTransform: 'uppercase'}}>{d.slot.split(' ')[0]}</text>;
          })}
          {points.length > 1 && (
             <>
               <path d={pathD} fill="none" stroke="#4ade80" strokeWidth="3" />
               <path d={fillD} fill="url(#gradient)" opacity="0.2" />
             </>
          )}
          <defs>
            <linearGradient id="gradient" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="#4ade80" /><stop offset="100%" stopColor="#4ade80" stopOpacity="0" /></linearGradient>
          </defs>
          {points.map((p, i) => (
            <g key={i}>
              <circle cx={p.x} cy={p.y} r="4" fill="#ffffff" stroke="#4ade80" strokeWidth="2" />
              <text x={p.x} y={p.y - 12} fill="#ffffff" fontSize="12" textAnchor="middle" fontWeight="black">{p.percentage.toFixed(1)}%</text>
            </g>
          ))}
          {points.length === 0 && <text x={width/2} y={height/2} fill="#64748b" textAnchor="middle" fontSize="14" fontWeight="bold">Waiting for production data...</text>}
        </svg>
      </div>
    </div>
  );
};

const TsprSheet: React.FC<{ data: TsprEntry[], onUpdate: (data: TsprEntry[]) => void }> = ({ data, onUpdate }) => {
  const handleChange = (index: number, field: 'total' | 'ng', value: string) => {
    const newData = [...data];
    const numVal = parseInt(value) || 0;
    newData[index] = { ...newData[index], [field]: numVal };
    onUpdate(newData);
  };
  const getPercentage = (total: number, ng: number) => total === 0 ? '100.0' : (((total - ng) / total) * 100).toFixed(1);

  return (
    <div className="space-y-6 animate-in fade-in slide-in-from-right-4">
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <div className="bg-white p-6 rounded-[2.5rem] shadow-xl border border-slate-100">
           <div className="flex items-center justify-between mb-6">
              <h3 className="font-black text-slate-800 text-lg tracking-tight">Data Entry</h3>
              <div className="px-3 py-1 bg-indigo-50 text-indigo-600 rounded-lg text-[10px] font-black uppercase tracking-widest border border-indigo-100">Live Input</div>
           </div>
           <div className="overflow-hidden rounded-2xl border border-slate-200">
             <table className="w-full text-sm">
               <thead className="bg-slate-50 text-slate-500 font-black text-[10px] uppercase tracking-widest">
                 <tr><th className="py-4 px-4 text-left">Time Slot</th><th className="py-4 px-4 text-center">Total</th><th className="py-4 px-4 text-center">NG</th><th className="py-4 px-4 text-right">Ratio</th></tr>
               </thead>
               <tbody className="divide-y divide-slate-100">
                 {data.map((row, i) => {
                   const pct = parseFloat(getPercentage(row.total, row.ng));
                   let color = 'text-green-600';
                   if(pct < 90) color = 'text-orange-500';
                   if(pct < 80) color = 'text-red-500';
                   return (
                     <tr key={row.slot} className="hover:bg-slate-50/50 transition-colors">
                       <td className="py-3 px-4 font-bold text-slate-700">{row.slot}</td>
                       <td className="py-3 px-4"><input type="number" className="w-full text-center bg-slate-100 border-transparent focus:bg-white focus:border-indigo-500 rounded-xl py-2 font-bold text-slate-800 outline-none" value={row.total || ''} onChange={(e) => handleChange(i, 'total', e.target.value)} /></td>
                       <td className="py-3 px-4"><input type="number" className="w-full text-center bg-slate-100 border-transparent focus:bg-white focus:border-indigo-500 rounded-xl py-2 font-bold text-slate-800 outline-none" value={row.ng || ''} onChange={(e) => handleChange(i, 'ng', e.target.value)} /></td>
                       <td className={`py-3 px-4 text-right font-black ${color}`}>{pct}%</td>
                     </tr>
                   );
                 })}
               </tbody>
             </table>
           </div>
        </div>
        <div className="flex flex-col gap-6">
           <div className="bg-indigo-600 p-6 rounded-[2.5rem] text-white shadow-xl shadow-indigo-500/20 relative overflow-hidden">
              <div className="relative z-10">
                 <h4 className="opacity-80 font-bold uppercase tracking-widest text-xs mb-1">Current Shift TSPR</h4>
                 <div className="text-5xl font-black tracking-tighter mb-4">{getPercentage(data.reduce((a,c)=>a+c.total,0), data.reduce((a,c)=>a+c.ng,0))}%</div>
                 <div className="flex gap-4 text-xs font-bold opacity-80">
                    <span>Total: {data.reduce((a,c)=>a+c.total,0)}</span><span>Defective: {data.reduce((a,c)=>a+c.ng,0)}</span>
                 </div>
              </div>
              <div className="absolute -right-6 -bottom-6 opacity-20 transform rotate-12"><TrendingUp size={120} /></div>
           </div>
           <TsprGraph data={data} />
        </div>
      </div>
    </div>
  );
};

const DefectCard: React.FC<{ defect: Defect }> = ({ defect }) => {
  return (
    <div className="bg-white p-5 rounded-[2rem] border border-slate-100 shadow-sm flex gap-4">
      <div className="w-24 h-24 bg-slate-100 rounded-2xl flex-shrink-0 overflow-hidden">
        {defect.photoUrl ? <img src={defect.photoUrl} className="w-full h-full object-cover" /> : <div className="flex items-center justify-center w-full h-full text-slate-300"><ImageIcon size={20} /></div>}
      </div>
      <div className="flex-1 min-w-0">
        <div className="flex justify-between items-start">
          <div>
            <span className="text-[10px] font-black uppercase tracking-widest text-slate-400">{defect.frameNo} • {defect.modelName}</span>
            <h4 className="font-bold text-slate-800 text-sm truncate">{defect.defectName}</h4>
          </div>
          <div className={`text-[10px] font-black px-3 py-1 rounded-lg uppercase tracking-widest ${defect.managerAnalysis?.status === 'CLOSED' ? 'bg-green-100 text-green-600' : 'bg-red-100 text-red-600'}`}>
            {defect.managerAnalysis?.status || 'OPEN'}
          </div>
        </div>
        <div className="mt-4 flex flex-wrap gap-1">
          {defect.targetedZones.map(z => {
             const r = defect.responses[z];
             let color = 'bg-slate-50 text-slate-400';
             if (r?.involved === 'YES') color = 'bg-red-50 text-red-500 border-red-100';
             if (r?.involved === 'NO') color = 'bg-green-50 text-green-500 border-green-100';
             return <span key={z} className={`text-[9px] font-black px-2 py-0.5 rounded border ${color}`}>{z}</span>;
          })}
        </div>
      </div>
    </div>
  );
};

const ManagerDefectThumbnail: React.FC<{ defect: Defect, onClick: () => void }> = ({ defect, onClick }) => {
  const isClosed = defect.managerAnalysis?.status === 'CLOSED';
  return (
    <button onClick={onClick} className="group relative aspect-square bg-slate-100 rounded-[2rem] overflow-hidden border border-slate-200 transition-all hover:scale-95 active:scale-90 shadow-sm">
      {defect.photoUrl ? <img src={defect.photoUrl} className="w-full h-full object-cover" /> : <div className="w-full h-full flex items-center justify-center"><ImageIcon className="text-slate-300" size={32} /></div>}
      <div className={`absolute top-3 right-3 p-1.5 rounded-xl backdrop-blur-md border ${isClosed ? 'bg-green-500/20 border-green-500/30 text-green-500' : 'bg-red-500/20 border-red-500/30 text-red-500'}`}>
        {isClosed ? <CheckCircle2 size={16} /> : <AlertTriangle size={16} />}
      </div>
      <div className="absolute inset-0 bg-gradient-to-t from-slate-900/80 via-transparent to-transparent opacity-0 group-hover:opacity-100 transition-opacity flex flex-col justify-end p-4 text-left">
        <span className="text-[9px] font-black text-white/60 uppercase tracking-widest">{defect.frameNo}</span>
        <h5 className="text-white font-bold text-xs truncate">{defect.defectName}</h5>
      </div>
    </button>
  );
};

const LeaderDefectTask: React.FC<{ defect: Defect, zoneId: ZoneId, onSave: (resp: ZoneResponse) => void }> = ({ defect, zoneId, onSave }) => {
  const response = defect.responses[zoneId] as ZoneResponse | undefined;
  const isDone = response && response.involved !== 'PENDING';
  const [isEditing, setIsEditing] = useState(!isDone);
  const [involved, setInvolved] = useState<'YES' | 'NO' | 'PENDING'>(response?.involved || 'PENDING');
  const [reason, setReason] = useState(response?.reason || '');
  const [actionTaken, setActionTaken] = useState(response?.actionTaken || '');
  const [manpower, setManpower] = useState<ManpowerEntry[]>(response?.manpower || []);

  const handleSave = () => {
    onSave({ involved, reason, actionTaken, manpower, updatedAt: new Date().toISOString() });
    setIsEditing(false);
  };

  if (!isEditing) {
    return (
      <div className="bg-white rounded-[2.5rem] p-6 shadow-sm border border-slate-100 flex flex-col md:flex-row gap-6">
        <div className="w-full md:w-48 h-48 rounded-[2rem] overflow-hidden flex-shrink-0">
          {defect.photoUrl ? <img src={defect.photoUrl} className="w-full h-full object-cover" /> : <div className="w-full h-full bg-slate-100 flex items-center justify-center text-slate-300"><ImageIcon size={32} /></div>}
        </div>
        <div className="flex-1">
          <div className="flex justify-between items-start mb-4">
            <div>
               <div className="flex items-center gap-2 mb-1">
                 <span className="px-2 py-0.5 bg-slate-100 rounded text-slate-500 text-[10px] font-bold">{defect.frameNo}</span>
                 <span className={`px-2 py-0.5 rounded text-[10px] font-black uppercase ${involved === 'YES' ? 'bg-red-50 text-red-500' : 'bg-green-50 text-green-500'}`}>{involved === 'YES' ? 'Involved' : 'Clear'}</span>
               </div>
               <h4 className="text-xl font-black text-slate-800 tracking-tight">{defect.defectName}</h4>
            </div>
            <button onClick={() => setIsEditing(true)} className="p-2 hover:bg-slate-50 rounded-xl transition-colors"><RefreshCw size={18} className="text-slate-400" /></button>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div className="bg-slate-50 p-4 rounded-2xl"><span className="block text-[9px] font-black text-slate-400 uppercase mb-1">Root Cause</span><p className="text-xs font-medium text-slate-600">{reason || 'No reason provided'}</p></div>
            <div className="bg-slate-50 p-4 rounded-2xl"><span className="block text-[9px] font-black text-slate-400 uppercase mb-1">Containment</span><p className="text-xs font-medium text-slate-600">{actionTaken || 'No action recorded'}</p></div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="bg-white rounded-[2.5rem] p-8 shadow-xl border-2 border-orange-500/20 animate-in zoom-in-95 duration-300">
      <div className="flex flex-col lg:flex-row gap-10">
        <div className="lg:w-1/3">
           <div className="relative group rounded-[2.5rem] overflow-hidden shadow-2xl">
              {defect.photoUrl ? <img src={defect.photoUrl} className="w-full aspect-square object-cover" /> : <div className="w-full aspect-square bg-slate-100 flex items-center justify-center text-slate-300"><ImageIcon size={64} /></div>}
           </div>
        </div>
        <div className="lg:w-2/3 space-y-8">
           <div>
              <span className="text-[10px] font-black text-orange-600 uppercase tracking-[0.2em] mb-2 block">Pending Analysis</span>
              <h3 className="text-3xl font-black text-slate-800 tracking-tighter">{defect.defectName}</h3>
              <p className="text-slate-500 text-sm font-medium mt-2">{defect.description}</p>
           </div>
           <div className="flex gap-4">
              <button onClick={() => setInvolved('YES')} className={`flex-1 py-4 rounded-2xl font-black uppercase tracking-widest text-xs transition-all border-2 ${involved === 'YES' ? 'bg-red-500 text-white border-red-500' : 'bg-white text-slate-400 border-slate-100 hover:border-red-200'}`}>Yes, Involved</button>
              <button onClick={() => setInvolved('NO')} className={`flex-1 py-4 rounded-2xl font-black uppercase tracking-widest text-xs transition-all border-2 ${involved === 'NO' ? 'bg-green-500 text-white border-green-500' : 'bg-white text-slate-400 border-slate-100 hover:border-green-200'}`}>Not Involved</button>
           </div>
           {involved === 'YES' && (
              <div className="space-y-6 animate-in slide-in-from-top-4">
                 <div className="space-y-2"><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest px-2">Reason (Why?)</label><textarea className="w-full px-5 py-4 rounded-[1.5rem] bg-slate-50 border border-slate-200 focus:bg-white focus:ring-4 focus:ring-orange-500/10 focus:border-orange-500 transition-all outline-none text-sm font-medium" placeholder="Why did this happen?" rows={2} value={reason} onChange={e => setReason(e.target.value)} /></div>
                 <div className="space-y-2"><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest px-2">Containment Action</label><textarea className="w-full px-5 py-4 rounded-[1.5rem] bg-slate-50 border border-slate-200 focus:bg-white focus:ring-4 focus:ring-orange-500/10 focus:border-orange-500 transition-all outline-none text-sm font-medium" placeholder="Immediate fix applied..." rows={2} value={actionTaken} onChange={e => setActionTaken(e.target.value)} /></div>
                 <div className="space-y-3">
                    <div className="flex justify-between items-center px-2">
                       <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Responsible Manpower</label>
                       <button onClick={() => setManpower([...manpower, { name: '', ein: '' }])} className="w-6 h-6 rounded-full bg-slate-100 flex items-center justify-center text-slate-600 hover:bg-orange-100 hover:text-orange-600 transition-colors"><Plus size={14} /></button>
                    </div>
                    {manpower.map((m, i) => (
                       <div key={i} className="flex gap-2 animate-in slide-in-from-left-2">
                          <input className="flex-1 px-4 py-3 bg-slate-50 rounded-xl text-xs font-bold border-none" placeholder="Name" value={m.name} onChange={e => { const nm = [...manpower]; nm[i].name = e.target.value; setManpower(nm); }} />
                          <input className="w-32 px-4 py-3 bg-slate-50 rounded-xl text-xs font-bold border-none" placeholder="EIN" value={m.ein} onChange={e => { const nm = [...manpower]; nm[i].ein = e.target.value; setManpower(nm); }} />
                       </div>
                    ))}
                 </div>
              </div>
           )}
           <button onClick={handleSave} disabled={involved === 'PENDING'} className="w-full bg-slate-900 text-white py-5 rounded-[1.5rem] font-black uppercase tracking-widest text-sm shadow-xl shadow-slate-900/10 active:scale-95 transition-all disabled:opacity-20">Submit Response</button>
        </div>
      </div>
    </div>
  );
};

const InspectorView: React.FC<{ defects: Defect[], onAddDefect: (d: Defect) => void }> = ({ defects, onAddDefect }) => {
  const [showForm, setShowForm] = useState(false);
  const [frameNo, setFrameNo] = useState('');
  const [modelName, setModelName] = useState('VERNA');
  const [defectName, setDefectName] = useState('');
  const [description, setDescription] = useState('');
  const [selectedZones, setSelectedZones] = useState<ZoneId[]>([]);
  const [photoUrl, setPhotoUrl] = useState<string | null>(null);
  const [isCapturing, setIsCapturing] = useState(false);
  const videoRef = useRef<HTMLVideoElement>(null);

  const startCamera = async () => {
    setIsCapturing(true);
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: 'environment' } });
      if (videoRef.current) videoRef.current.srcObject = stream;
    } catch (e) {
      alert("Camera failed. Check permissions.");
      setIsCapturing(false);
    }
  };

  const capturePhoto = () => {
    if (videoRef.current) {
      const canvas = document.createElement('canvas');
      canvas.width = videoRef.current.videoWidth;
      canvas.height = videoRef.current.videoHeight;
      canvas.getContext('2d')?.drawImage(videoRef.current, 0, 0);
      setPhotoUrl(canvas.toDataURL('image/jpeg', 0.8));
      const stream = videoRef.current.srcObject as MediaStream;
      stream.getTracks().forEach(track => track.stop());
      setIsCapturing(false);
    }
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (selectedZones.length === 0) { alert("Target at least one zone."); return; }
    const newDefect: Defect = {
      id: Math.random().toString(36).substr(2, 9).toUpperCase(),
      timestamp: new Date().toISOString(),
      frameNo, modelName, defectName, description,
      photoUrl: photoUrl || '',
      targetedZones: selectedZones,
      responses: {}
    };
    onAddDefect(newDefect);
    setFrameNo(''); setDefectName(''); setDescription(''); setSelectedZones([]); setPhotoUrl(null); setShowForm(false);
  };

  if (showForm) {
    return (
      <div className="bg-white rounded-[3rem] shadow-2xl overflow-hidden border border-slate-100 animate-in slide-in-from-bottom-8 duration-500">
        <div className="p-8 bg-slate-900 text-white flex justify-between items-center">
           <h2 className="text-3xl font-black tracking-tighter">Report Defect</h2>
           <button onClick={() => setShowForm(false)} className="p-2 bg-white/10 rounded-2xl hover:bg-white/20 transition-all"><X /></button>
        </div>
        <form onSubmit={handleSubmit} className="p-8 space-y-10">
           <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
              <div className="space-y-6">
                 <div className="space-y-2"><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest px-2">Vehicle Identity</label><div className="flex gap-2"><input required className="flex-1 px-6 py-4 bg-slate-50 rounded-2xl border-2 border-transparent focus:border-blue-600 focus:bg-white transition-all font-bold outline-none" placeholder="FRAME NO (e.g. 1234)" value={frameNo} onChange={e => setFrameNo(e.target.value.toUpperCase())} /><select className="px-6 py-4 bg-slate-50 rounded-2xl font-bold border-none outline-none" value={modelName} onChange={e => setModelName(e.target.value)}><option>VERNA</option><option>CRETA</option><option>ALCAZAR</option><option>EXTERNAL</option></select></div></div>
                 <div className="space-y-2"><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest px-2">Issue Type</label><input required className="w-full px-6 py-4 bg-slate-50 rounded-2xl border-2 border-transparent focus:border-blue-600 focus:bg-white transition-all font-bold outline-none" placeholder="Defect Name (e.g. Paint Scratch)" value={defectName} onChange={e => setDefectName(e.target.value)} /></div>
                 <div className="space-y-2"><label className="text-[10px] font-black text-slate-400 uppercase tracking-widest px-2">Details</label><textarea className="w-full px-6 py-4 bg-slate-50 rounded-2xl border-2 border-transparent focus:border-blue-600 focus:bg-white transition-all font-bold outline-none" placeholder="Describe the findings..." rows={3} value={description} onChange={e => setDescription(e.target.value)} /></div>
              </div>
              <div className="space-y-4">
                 <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest px-2">Visual Evidence</label>
                 <div className="aspect-square bg-slate-100 rounded-[2.5rem] overflow-hidden relative group border-4 border-dashed border-slate-200">
                    {photoUrl ? <><img src={photoUrl} className="w-full h-full object-cover" /><button type="button" onClick={() => setPhotoUrl(null)} className="absolute top-4 right-4 bg-red-500 text-white p-3 rounded-2xl shadow-xl hover:bg-red-600 transition-all"><X size={20}/></button></> : isCapturing ? <><video ref={videoRef} autoPlay playsInline className="w-full h-full object-cover" /><button type="button" onClick={capturePhoto} className="absolute bottom-6 left-1/2 -translate-x-1/2 w-16 h-16 bg-white rounded-full border-4 border-blue-600 shadow-2xl active:scale-90 transition-all" /></> : <button type="button" onClick={startCamera} className="w-full h-full flex flex-col items-center justify-center text-slate-400 hover:text-blue-600 hover:bg-slate-50 transition-all gap-4"><Camera size={48} /><span className="text-xs font-black uppercase tracking-widest">Tap to Launch Camera</span></button>}
                 </div>
              </div>
           </div>
           <div className="space-y-4">
              <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest px-2">Target Responsibility Zones</label>
              <div className="grid grid-cols-3 md:grid-cols-9 gap-2">
                 {ZONES.map(z => <button key={z} type="button" onClick={() => setSelectedZones(p => p.includes(z) ? p.filter(x => x !== z) : [...p, z])} className={`py-4 rounded-xl text-xs font-black transition-all border-2 ${selectedZones.includes(z) ? 'bg-blue-600 text-white border-blue-600 shadow-lg shadow-blue-500/20' : 'bg-white text-slate-400 border-slate-100 hover:border-blue-200'}`}>{z}</button>)}
              </div>
           </div>
           <button type="submit" className="w-full bg-blue-600 text-white py-6 rounded-[2rem] font-black uppercase tracking-widest text-lg shadow-2xl shadow-blue-500/30 hover:bg-blue-700 active:scale-95 transition-all">Report to Production Line</button>
        </form>
      </div>
    );
  }

  return (
    <div className="space-y-10 animate-in fade-in duration-700">
      <div className="flex justify-between items-end">
        <div><h2 className="text-4xl font-black text-slate-800 tracking-tighter">Line Quality</h2><p className="text-xs font-bold text-slate-400 uppercase tracking-widest">Active Assembly Line Track</p></div>
        <button onClick={() => setShowForm(true)} className="bg-blue-600 text-white p-4 rounded-3xl font-black flex items-center shadow-2xl shadow-blue-500/20 active:scale-90 transition-all"><Plus size={24} /></button>
      </div>
      <div className="space-y-6">
        {defects.length === 0 ? (
          <div className="text-center py-40 bg-white rounded-[3rem] border-4 border-dashed border-slate-100 shadow-inner">
            <div className="bg-slate-50 w-20 h-20 rounded-full flex items-center justify-center mx-auto mb-6"><ClipboardList className="text-slate-200" size={40} /></div>
            <p className="text-slate-400 font-bold uppercase tracking-widest text-[10px]">No active defects reported</p>
          </div>
        ) : ( defects.map(d => <DefectCard key={d.id} defect={d} />) )}
      </div>
    </div>
  );
};

const LeaderZoneView: React.FC<{ zoneId: ZoneId, defects: Defect[], onSave: (d: Defect) => void }> = ({ zoneId, defects, onSave }) => {
  const zoneDefects = defects.filter(d => d.targetedZones.includes(zoneId) || d.responses[zoneId]);
  return (
    <div className="space-y-10 animate-in fade-in duration-700">
      <div className="flex justify-between items-end">
        <div><h2 className="text-4xl font-black text-slate-800 tracking-tighter">Zone {zoneId} Panel</h2><p className="text-xs font-bold text-slate-400 uppercase tracking-widest">Correction & Verification Tasks</p></div>
        <div className="px-6 py-3 bg-orange-100 text-orange-600 rounded-2xl font-black uppercase tracking-widest text-xs flex items-center gap-2"><Clock size={16} />{zoneDefects.filter(d => !d.responses[zoneId]).length} Pending</div>
      </div>
      <div className="space-y-8">
        {zoneDefects.length === 0 ? (
           <div className="text-center py-40 bg-white rounded-[3rem] border-4 border-dashed border-slate-100 shadow-inner">
              <div className="bg-slate-50 w-20 h-20 rounded-full flex items-center justify-center mx-auto mb-6"><CheckCircle2 className="text-slate-200" size={40} /></div>
              <p className="text-slate-400 font-bold uppercase tracking-widest text-[10px]">Zone is clear</p>
           </div>
        ) : (
          zoneDefects.map(d => (
            <LeaderDefectTask key={d.id} defect={d} zoneId={zoneId} onSave={(resp) => onSave({ ...d, responses: { ...d.responses, [zoneId]: resp } })} />
          ))
        )}
      </div>
    </div>
  );
};

const ManagerDefectModal: React.FC<{ defect: Defect, onClose: () => void, onSave: (d: Defect) => void }> = ({ defect, onClose, onSave }) => {
  const [analysis, setAnalysis] = useState<ManagerAnalysis>(defect.managerAnalysis || { man: '', machine: '', material: '', method: '', rootCauseSummary: '', status: 'OPEN' });
  const handleSave = () => { onSave({ ...defect, managerAnalysis: { ...analysis, closedAt: analysis.status === 'CLOSED' ? new Date().toISOString() : undefined } }); onClose(); };

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/90 backdrop-blur-sm overflow-y-auto">
      <div className="min-h-screen flex items-center justify-center p-4">
        <div className="bg-white w-full max-w-4xl rounded-[2rem] shadow-2xl overflow-hidden animate-in zoom-in-95 duration-200">
          <div className="relative h-96 bg-slate-100 group">
            {defect.photoUrl ? <img src={defect.photoUrl} className="w-full h-full object-cover" /> : <div className="w-full h-full flex items-center justify-center text-slate-200"><ImageIcon size={64} /></div>}
            <button onClick={onClose} className="absolute top-6 right-6 p-4 bg-white/20 backdrop-blur-xl border border-white/30 text-white rounded-full hover:bg-white/40 active:scale-90 transition-all"><X /></button>
            <div className="absolute bottom-0 left-0 right-0 p-10 bg-gradient-to-t from-slate-900 via-slate-900/40 to-transparent">
              <span className="text-[10px] font-black text-indigo-400 uppercase tracking-widest mb-2 block">{defect.frameNo} • {defect.modelName}</span>
              <h3 className="text-4xl font-black text-white tracking-tighter">{defect.defectName}</h3>
            </div>
          </div>
          <div className="p-10">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-8 mb-10">
              {['man', 'machine', 'material', 'method'].map(f => (
                <div key={f} className="space-y-2">
                   <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest px-2">4M Analysis: {f.toUpperCase()}</label>
                   <input className="w-full px-5 py-4 bg-slate-50 border border-slate-100 rounded-xl text-sm font-bold focus:bg-white focus:ring-2 focus:ring-indigo-500/20 outline-none transition-all" value={(analysis as any)[f]} onChange={e => setAnalysis({...analysis, [f]: e.target.value})} placeholder={`Observation for ${f}...`} />
                </div>
              ))}
            </div>
            <div className="grid grid-cols-1 md:grid-cols-3 gap-8 items-end">
               <div className="md:col-span-2 space-y-2">
                  <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest px-2">Manager Root Cause Summary</label>
                  <textarea className="w-full px-5 py-4 bg-slate-50 border border-indigo-100 rounded-xl p-4 text-sm font-medium focus:bg-white focus:ring-2 focus:ring-indigo-500/20 outline-none transition-all" rows={4} placeholder="Final conclusion..." value={analysis.rootCauseSummary} onChange={e => setAnalysis({...analysis, rootCauseSummary: e.target.value})} />
               </div>
               <div className="space-y-2">
                  <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest px-2">Defect Status</label>
                  <select className="w-full px-5 py-4 bg-slate-50 border border-slate-200 rounded-xl font-black text-sm uppercase tracking-widest" value={analysis.status} onChange={e => setAnalysis({...analysis, status: e.target.value as any})}>
                     <option value="OPEN">Open (Pending)</option>
                     <option value="CLOSED">Closed (Verified)</option>
                  </select>
               </div>
            </div>
            <div className="pt-6 mt-6 border-t border-slate-100"><button onClick={handleSave} className="w-full bg-slate-900 text-white py-4 rounded-xl font-black text-sm uppercase tracking-widest shadow-xl hover:bg-slate-800 active:scale-95 transition-all">Save Analysis Sheet</button></div>
          </div>
        </div>
      </div>
    </div>
  );
};

const SettingsModal: React.FC<{ currentUrl: string, onSave: (url: string) => void, onClose: () => void }> = ({ currentUrl, onSave, onClose }) => {
  const [url, setUrl] = useState(currentUrl);
  const [copied, setCopied] = useState(false);
  const copyScript = () => { navigator.clipboard.writeText(SCRIPT_CODE); setCopied(true); setTimeout(() => setCopied(false), 2000); };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-sm animate-in fade-in">
      <div className="bg-white p-6 rounded-[2rem] shadow-2xl w-full max-w-2xl max-h-[90vh] overflow-y-auto">
        <div className="flex justify-between items-center mb-6">
           <h3 className="text-xl font-black tracking-tight flex items-center gap-2"><Settings size={20} /> Cloud Connection</h3>
           <button onClick={onClose} className="p-2 hover:bg-slate-100 rounded-full transition-colors"><X /></button>
        </div>
        <div className="space-y-6">
           <div className="bg-blue-50 p-6 rounded-2xl border border-blue-100">
              <div className="flex items-center gap-3 mb-2 text-blue-700 font-bold text-sm"><Database size={18} /> Step 1: Apps Script URL</div>
              <p className="text-xs text-blue-600 mb-4 font-medium leading-relaxed">Enter your Google Apps Script Web App URL below to sync data to Google Sheets.</p>
              <input className="w-full px-4 py-3 rounded-xl border-2 border-blue-200 focus:border-blue-500 outline-none text-xs font-mono font-bold" value={url} onChange={e => setUrl(e.target.value)} placeholder="https://script.google.com/macros/s/.../exec" />
           </div>
           <div className="bg-slate-50 p-6 rounded-2xl border border-slate-200">
              <div className="flex items-center justify-between mb-4">
                 <div className="flex items-center gap-3 text-slate-700 font-bold text-sm"><FileText size={18} /> Step 2: Google Sheets Code</div>
                 <span className="px-2 py-1 bg-slate-100 text-slate-500 text-[10px] font-bold rounded border border-slate-200">Storage_DB (Hidden)</span>
              </div>
              <div className="relative group">
                 <div className="absolute top-2 right-2"><button onClick={copyScript} className={`flex items-center gap-2 px-3 py-1.5 rounded-lg text-[10px] font-black uppercase tracking-widest transition-all ${copied ? 'bg-green-500 text-white' : 'bg-white text-slate-600 shadow-sm hover:bg-slate-50'}`}>{copied ? <Check size={12} /> : <Copy size={12} />}{copied ? 'Copied!' : 'Copy Code'}</button></div>
                 <pre className="p-4 bg-slate-900 text-slate-300 rounded-xl text-[10px] font-mono overflow-x-auto max-h-40">{SCRIPT_CODE}</pre>
              </div>
           </div>
           <button onClick={() => onSave(url)} className="w-full bg-slate-900 text-white py-4 rounded-xl font-black text-sm uppercase tracking-widest">Connect and Save</button>
        </div>
      </div>
    </div>
  );
};

// --- MAIN APP COMPONENT ---

const App: React.FC = () => {
  const [defects, setDefects] = useState<Defect[]>(() => {
    const saved = localStorage.getItem(LOCAL_STORAGE_KEY);
    return saved ? JSON.parse(saved) : [];
  });
  const [tsprData, setTsprData] = useState<TsprEntry[]>(() => {
    const saved = localStorage.getItem(TSPR_STORAGE_KEY);
    return saved ? JSON.parse(saved) : INITIAL_TSPR_DATA;
  });
  const [session, setSession] = useState<UserSession | null>(() => {
    const saved = localStorage.getItem(SESSION_KEY);
    return saved ? JSON.parse(saved) : null;
  });
  const [cloudUrl, setCloudUrl] = useState<string>(() => localStorage.getItem(CLOUD_URL_KEY) || '');
  const [showSettings, setShowSettings] = useState(false);
  const [isSyncing, setIsSyncing] = useState(false);
  const [selectedDefect, setSelectedDefect] = useState<Defect | null>(null);
  const [managerTab, setManagerTab] = useState<'DEFECTS' | 'TSPR'>('DEFECTS');

  useEffect(() => { localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(defects)); }, [defects]);
  useEffect(() => { localStorage.setItem(TSPR_STORAGE_KEY, JSON.stringify(tsprData)); }, [tsprData]);

  // --- AUTOMATIC REFRESH HEARTBEAT (Every 15s) ---
  useEffect(() => {
    if (!cloudUrl) return;
    loadInitialData(); 
    
    const interval = setInterval(() => {
      handleRefresh(); 
    }, 15000);

    return () => clearInterval(interval);
  }, [cloudUrl]);

  // --- ROBUST SYNC LOGIC (Merged Push) ---
  const robustSync = async (updatedDefects: Defect[], updatedTspr: TsprEntry[]) => {
    if (!cloudUrl) return;
    setIsSyncing(true);
    try {
      const response = await fetch(cloudUrl);
      const cloudJson = await response.json();
      const serverDefects: Defect[] = cloudJson.data || [];

      // Merge: Local IDs overwrite cloud matches, missing IDs from cloud are added
      const map: Record<string, Defect> = {};
      serverDefects.forEach(d => (map[d.id] = d));
      updatedDefects.forEach(d => (map[d.id] = d));
      const mergedDefects = Object.values(map);

      setDefects(mergedDefects);
      setTsprData(updatedTspr);

      await fetch(cloudUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ defects: mergedDefects, tspr: updatedTspr })
      });
    } catch (err) {
      console.error("Auto-sync failed:", err);
    } finally {
      setIsSyncing(false);
    }
  };

  const loadInitialData = async () => {
    if (!cloudUrl) return;
    setIsSyncing(true);
    try {
        const response = await fetch(cloudUrl);
        const json = await response.json();
        if (json.status === 'success') {
           if (json.data) setDefects(json.data);
           if (json.tspr) setTsprData(json.tspr);
        }
    } catch (e) { console.error("Load failed", e); }
    finally { setIsSyncing(false); }
  };

  const handleRefresh = async () => {
    if (!cloudUrl) return;
    setIsSyncing(true);
    try {
        const response = await fetch(cloudUrl);
        const json = await response.json();
        if (json.status === 'success') {
           if (json.data) setDefects(json.data);
           if (json.tspr) setTsprData(json.tspr);
        }
    } catch (e) { console.error("Refresh failed", e); }
    finally { setIsSyncing(false); }
  };

  const handleLogin = (role: AppRole, zoneId?: ZoneId) => {
    const newSession: UserSession = { role, zoneId, name: role === 'GROUP_LEADER' ? `Leader ${zoneId}` : role === 'INSPECTOR' ? 'Inspector' : 'Manager' };
    setSession(newSession);
    localStorage.setItem(SESSION_KEY, JSON.stringify(newSession));
  };

  const handleLogout = () => { setSession(null); localStorage.removeItem(SESSION_KEY); };

  if (!session) {
    return (
      <div className="min-h-screen bg-slate-50 flex items-center justify-center p-4">
        <div className="max-w-md w-full space-y-8">
          <div className="text-center">
            <div className="inline-flex p-4 bg-indigo-600 rounded-[2rem] text-white shadow-xl shadow-indigo-500/30 mb-6"><Layers size={40} /></div>
            <h1 className="text-4xl font-black text-slate-800 tracking-tighter mb-2">AutoLine</h1>
            <p className="text-slate-400 font-bold uppercase tracking-widest text-[10px]">Real-time Quality Management</p>
          </div>
          <div className="space-y-4">
            {(['INSPECTOR', 'GROUP_LEADER', 'MANAGER'] as AppRole[]).map(role => (
              <div key={role} className="space-y-2">
                {role === 'GROUP_LEADER' ? (
                  <div className="grid grid-cols-5 gap-2">
                    <div className="col-span-5 text-[10px] font-black text-slate-400 uppercase tracking-widest px-2 mb-1">Group Leaders</div>
                    {ZONES.map(z => <button key={z} onClick={() => handleLogin('GROUP_LEADER', z)} className="py-3 bg-white border border-slate-200 rounded-xl text-[10px] font-black hover:border-orange-500 hover:text-orange-500 transition-all shadow-sm">{z}</button>)}
                  </div>
                ) : (
                  <button onClick={() => handleLogin(role)} className={`w-full p-6 bg-white border border-slate-200 rounded-[2rem] flex items-center gap-6 shadow-sm hover:shadow-xl hover:border-indigo-500 transition-all group`}>
                    <div className={`p-4 rounded-2xl ${ROLE_CONFIG[role].color} text-white group-hover:scale-110 transition-transform`}>{ROLE_CONFIG[role].icon}</div>
                    <div className="text-left"><h3 className="font-black text-slate-800 uppercase tracking-widest text-sm">{ROLE_CONFIG[role].title}</h3><p className="text-xs font-medium text-slate-400">Tap to authorize access</p></div>
                    <ChevronRight className="ml-auto text-slate-300 group-hover:text-indigo-500 group-hover:translate-x-1 transition-all" />
                  </button>
                )}
              </div>
            ))}
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 pb-20">
      <nav className="bg-white/80 backdrop-blur-md sticky top-0 z-40 border-b border-slate-100">
        <div className="max-w-7xl mx-auto px-6 h-20 flex items-center justify-between">
          <div className="flex items-center gap-4">
             <div className="bg-slate-900 text-white p-2.5 rounded-2xl"><Layers size={20} /></div>
             <div className="w-px h-6 bg-slate-200" />
             <div className="flex flex-col"><span className="text-[10px] font-black text-slate-400 uppercase tracking-widest leading-none mb-1">Live Environment</span><span className="text-sm font-black text-slate-800 tracking-tight">{ROLE_CONFIG[session.role].title} {session.zoneId ? `• Zone ${session.zoneId}` : ''}</span></div>
          </div>
          <div className="flex items-center gap-2">
            <button onClick={handleRefresh} disabled={isSyncing} className="p-3 rounded-2xl bg-white text-slate-600 hover:bg-slate-50 hover:text-blue-600 transition-colors shadow-sm border border-slate-100"><RefreshCw size={20} className={isSyncing ? 'animate-spin' : ''} /></button>
            <button onClick={() => setShowSettings(true)} className="p-3 rounded-2xl bg-slate-100 text-slate-600 hover:bg-slate-200 transition-colors"><Settings size={20} /></button>
            <div className="w-px h-8 bg-slate-200 mx-2" />
            <button onClick={handleLogout} className="p-3 rounded-2xl bg-red-50 text-red-500 hover:bg-red-100 transition-colors"><LogOut size={20} /></button>
          </div>
        </div>
      </nav>

      <div className="max-w-7xl mx-auto px-6 pt-10">
         {session.role === 'INSPECTOR' && <InspectorView defects={defects} onAddDefect={(newDefect) => { const upd = [newDefect, ...defects]; setDefects(upd); robustSync(upd, tsprData); }} />}
         {session.role === 'GROUP_LEADER' && <LeaderZoneView zoneId={session.zoneId!} defects={defects} onSave={(updated) => { const upd = defects.map(d => d.id === updated.id ? updated : d); setDefects(upd); robustSync(upd, tsprData); }} />}
         {session.role === 'MANAGER' && (
            <div className="space-y-10 animate-in fade-in duration-700">
               <div className="flex flex-col md:flex-row justify-between items-start md:items-end gap-6">
                  <div><h2 className="text-4xl font-black text-slate-800 tracking-tighter">Line Management</h2><p className="text-xs font-bold text-slate-400 uppercase tracking-widest">Global Overview & Analytical Control</p></div>
                  <div className="flex gap-2 bg-white p-1.5 rounded-[1.5rem] shadow-sm border border-slate-100">
                     <button onClick={() => setManagerTab('DEFECTS')} className={`px-6 py-3 rounded-xl text-xs font-black uppercase tracking-widest transition-all ${managerTab === 'DEFECTS' ? 'bg-indigo-600 text-white shadow-lg' : 'text-slate-400 hover:bg-slate-50'}`}>Analysis</button>
                     <button onClick={() => setManagerTab('TSPR')} className={`px-6 py-3 rounded-xl text-xs font-black uppercase tracking-widest transition-all ${managerTab === 'TSPR' ? 'bg-indigo-600 text-white shadow-lg' : 'text-slate-400 hover:bg-slate-50'}`}>Production</button>
                  </div>
               </div>
               {managerTab === 'DEFECTS' && (
                  <div className="space-y-10">
                     <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                        <div className="bg-white p-6 rounded-[2rem] border border-slate-100 shadow-sm"><span className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-1 block">Total Defects</span><div className="text-3xl font-black text-slate-800">{defects.length}</div></div>
                        <div className="bg-white p-6 rounded-[2rem] border border-slate-100 shadow-sm"><span className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-1 block">Closed Cases</span><div className="text-3xl font-black text-green-500">{defects.filter(d => d.managerAnalysis?.status === 'CLOSED').length}</div></div>
                        <div className="bg-white p-6 rounded-[2rem] border border-slate-100 shadow-sm"><span className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-1 block">Avg TSPR</span><div className="text-3xl font-black text-indigo-600">92%</div></div>
                        <div className="bg-indigo-600 p-6 rounded-[2rem] text-white flex flex-col justify-center items-center gap-2 cursor-pointer hover:bg-indigo-700 transition-all" onClick={() => downloadReport(defects)}><Download size={24} /><span className="text-[10px] font-black uppercase tracking-widest">Export Excel</span></div>
                     </div>
                     <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-6 gap-6">
                        {defects.length === 0 ? <div className="col-span-full py-20 text-center font-bold text-slate-300 uppercase tracking-widest">No defect entries found</div> : defects.map(d => <ManagerDefectThumbnail key={d.id} defect={d} onClick={() => setSelectedDefect(d)} />)}
                     </div>
                  </div>
               )}
               {managerTab === 'TSPR' && (
                  <div className="animate-in fade-in slide-in-from-bottom-4">
                      <div className="mb-6"><h2 className="text-3xl font-black text-slate-800 tracking-tighter">TSPR Control Sheet</h2><p className="text-xs font-bold text-slate-400 uppercase tracking-widest">Production Data Entry & Trend Analysis</p></div>
                      <TsprSheet data={tsprData} onUpdate={(newData) => { setTsprData(newData); robustSync(defects, newData); }} />
                  </div>
               )}
            </div>
         )}
      </div>

      {showSettings && <SettingsModal currentUrl={cloudUrl} onSave={(url) => { setCloudUrl(url); localStorage.setItem(CLOUD_URL_KEY, url); setShowSettings(false); }} onClose={() => setShowSettings(false)} />}
      {selectedDefect && <ManagerDefectModal defect={selectedDefect} onClose={() => setSelectedDefect(null)} onSave={(updated) => { const upd = defects.map(d => d.id === updated.id ? updated : d); setDefects(upd); robustSync(upd, tsprData); }} />}
    </div>
  );
};

export default App;
