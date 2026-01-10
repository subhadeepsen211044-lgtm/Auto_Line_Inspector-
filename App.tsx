import React, { useState, useEffect, useRef } from 'react';


const syncToCloud = async (updatedDefects: Defect[], tspr: TsprEntry[]) => {
  if (!cloudUrl) return;

  try {
    // 1️⃣ Fetch latest from server
    const latestRes = await fetch(cloudUrl);
    const latest = await latestRes.json();

    const serverDefects: Defect[] = latest.data || [];

    // 2️⃣ Merge (local wins on conflict)
    const map: Record<string, Defect> = {};
    serverDefects.forEach(d => (map[d.id] = d));
    updatedDefects.forEach(d => (map[d.id] = d));

    const merged = Object.values(map);

    // 3️⃣ POST merged data
    await fetch(cloudUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ defects: merged, tspr })
    });
  } catch (err) {
    console.error("Cloud sync failed:", err);
  }
};

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

// Helper to process images: converts unsupported formats (like webp) to PNG and returns dimensions
const processImageForExcel = (base64Url: string): Promise<{ base64: string, extension: 'png' | 'jpeg' | 'gif', width: number, height: number }> => {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.crossOrigin = 'Anonymous';
    img.onload = () => {
      const mime = base64Url.match(/data:image\/(\w+);base64/)?.[1]?.toLowerCase();
      // ExcelJS natively supports png, jpeg, gif. Others (like webp) need conversion.
      if (mime === 'png' || mime === 'jpeg' || mime === 'jpg' || mime === 'gif') {
         resolve({ 
           base64: base64Url.split(',')[1], 
           extension: mime === 'jpg' ? 'jpeg' : mime as 'png' | 'jpeg' | 'gif',
           width: img.width,
           height: img.height
         });
      } else {
         // Convert unsupported formats to PNG via Canvas
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
               extension: 'png', // Fallback
               width: img.width,
               height: img.height
            }); 
         }
      }
    };
    img.onerror = () => {
        // Fallback if loading fails, just try to pass raw data
        resolve({ 
           base64: base64Url.split(',')[1], 
           extension: 'png', 
           width: 100, 
           height: 100 
        });
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

  // Define Columns
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
    { header: 'Visual Evidence', key: 'image', width: 65 },  // MOVED TO END & WIDENED
  ];

  // Style the Header Row
  const headerRow = worksheet.getRow(1);
  headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12 };
  headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4F46E5' } }; // Indigo-600
  headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
  headerRow.height = 30;

  // Iterate over defects and add rows
  for (let i = 0; i < defects.length; i++) {
    const d = defects[i];
    const dateObj = new Date(d.timestamp);
    const m = d.managerAnalysis || { man: '', machine: '', material: '', method: '', rootCauseSummary: '', status: 'OPEN' };

    // Construct Zone Analysis Text
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

    // Style the Row
    row.height = 220; // Increased for better image visibility
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

    // --- EMBED IMAGE ---
    if (d.photoUrl && d.photoUrl.startsWith('data:image')) {
       try {
         const processedImage = await processImageForExcel(d.photoUrl);

         const imageId = workbook.addImage({
           base64: processedImage.base64,
           extension: processedImage.extension,
         });

         // IMPROVED SIZING - Fills cell with 5% padding
         const cellWidthPx = 450;  // Approx 65 Excel width units
         const cellHeightPx = 280;  // Approx 220 points height
         const cellAspectRatio = cellWidthPx / cellHeightPx;
         const imgAspectRatio = processedImage.width / processedImage.height;

         const PADDING = 0.05; // 5% padding on each side
         let wRatio = 1 - (PADDING * 2);
         let hRatio = 1 - (PADDING * 2);

         // Scale to fit while maintaining aspect ratio
         if (imgAspectRatio > cellAspectRatio) {
           hRatio = (1 - PADDING * 2) * (cellAspectRatio / imgAspectRatio);
         } else {
           wRatio = (1 - PADDING * 2) * (imgAspectRatio / cellAspectRatio);
         }

         // Center the image
         const colOff = PADDING + (1 - PADDING * 2 - wRatio) / 2;
         const rowOff = PADDING + (1 - PADDING * 2 - hRatio) / 2;
         
         const imageColIndex = 11; // Visual Evidence is at index 11 (12th column)

         worksheet.addImage(imageId, {
           tl: { col: imageColIndex + colOff, row: row.number - 1 + rowOff } as any,
           br: { col: imageColIndex + colOff + wRatio, row: row.number - 1 + rowOff + hRatio } as any,
           editAs: 'oneCell' 
         });
       } catch (e) {
         console.error("Error adding image to Excel", e);
         row.getCell('image').value = "[Image Error]";
       }
    } else {
       row.getCell('image').value = "No Image";
    }
  }

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  
  // Use a unique filename with seconds to prevent browser cache/overwrite issues
  const filename = `AUTOLINE_REPORT_${new Date().toISOString().replace(/[:.]/g, '-')}.xlsx`;
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
};

const downloadJSON = (defects: Defect[]) => {
  const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(defects, null, 2));
  const link = document.createElement("a");
  link.setAttribute("href", dataStr);
  link.setAttribute("download", `autoline_backup_${new Date().toISOString().slice(0,10)}.json`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
};

const SCRIPT_CODE = `function doPost(e) {
  return handleRequest(e);
}

function doGet(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Handle POST data (Saving)
  if (e.postData && e.postData.contents) {
    var json = JSON.parse(e.postData.contents);
    var defects = json.defects || [];
    var tspr = json.tspr || [];

    // --- 1. SAVE DEFECTS to Storage_DB (The "State" for the App) ---
    var storageSheet = ss.getSheetByName("Storage_DB");
    if (!storageSheet) { storageSheet = ss.insertSheet("Storage_DB"); }
    // Save the entire state as a JSON string in cell A1
    storageSheet.getRange(1, 1).setValue(JSON.stringify({ defects: defects, tspr: tspr }));

    // --- 2. SAVE READABLE DEFECTS (For Humans/Reporting) ---
    var sheet = ss.getSheetByName("Defects");
    if (!sheet) { sheet = ss.insertSheet("Defects"); }
    
    // Headers
    var headers = [
      'ID', 'Timestamp', 'Frame No', 'Model', 'Defect Name', 'Description', 
      'Targeted Zones', 'Involved Zones', 'Detailed Findings', 'Manpower', 
      '4M Man', '4M Machine', '4M Material', '4M Method', 'Root Cause Summary', 'Status'
    ];
    
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e0e0e0");
      sheet.setFrozenRows(1);
    }
    
    // Transform Data
    var rows = [];
    defects.forEach(function(d) {
      var involved = [];
      var findings = [];
      var manpower = [];
      
      // Collect responses from all zones
      var allZones = [];
      if (d.targetedZones) allZones = allZones.concat(d.targetedZones);
      if (d.responses) allZones = allZones.concat(Object.keys(d.responses));
      // Unique zones
      allZones = allZones.filter(function(item, pos) { return allZones.indexOf(item) == pos; });

      allZones.forEach(function(z) {
          var r = d.responses ? d.responses[z] : null;
          if (r && r.involved === 'YES') {
            involved.push(z);
            findings.push(z + ": " + (r.reason || '-') + " -> " + (r.actionTaken || '-'));
            if (r.manpower) {
               r.manpower.forEach(function(m) { manpower.push(m.name + " (" + m.ein + ")"); });
            }
          } else if (r && r.involved === 'NO') {
            findings.push(z + ": Not Involved");
          } else if (d.targetedZones && d.targetedZones.indexOf(z) > -1) {
            findings.push(z + ": Pending");
          }
      });

      var m = d.managerAnalysis || {};
      rows.push([
        d.id, new Date(d.timestamp).toLocaleString(), d.frameNo, d.modelName, d.defectName, d.description,
        d.targetedZones ? d.targetedZones.join(', ') : '', involved.join(', '), findings.join('\\n'), manpower.join('\\n'),
        m.man || '', m.machine || '', m.material || '', m.method || '', m.rootCauseSummary || '', m.status || 'OPEN'
      ]);
    });

    // Overwrite Defects Sheet with fresh data
    if (sheet.getLastRow() > 1) {
      // Clear all content below header
      sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn()).clearContent();
    }
    if (rows.length > 0) {
      var range = sheet.getRange(2, 1, rows.length, rows[0].length);
      range.setValues(rows);
      range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      range.setVerticalAlignment("top");
    }

    // --- 3. SAVE TSPR ---
    var tsprSheet = ss.getSheetByName("TSPR");
    if (!tsprSheet) { tsprSheet = ss.insertSheet("TSPR"); }
    
    var tsprHeaders = ["Time Slot", "Total Vehicles", "NG Vehicles"];
    if (tsprSheet.getLastRow() === 0) {
        tsprSheet.appendRow(tsprHeaders);
        tsprSheet.getRange(1, 1, 1, 3).setFontWeight("bold");
    }
    
    if (tsprSheet.getLastRow() > 1) {
        tsprSheet.getRange(2, 1, tsprSheet.getLastRow(), 3).clearContent();
    }
    
    var tsprRows = tspr.map(function(t) { return [t.slot, t.total, t.ng]; });
    if (tsprRows.length > 0) {
        tsprSheet.getRange(2, 1, tsprRows.length, 3).setValues(tsprRows);
    }
    
    // Return Success
    return ContentService.createTextOutput(JSON.stringify({
       status: 'success', 
       data: defects,
       tspr: tspr
    })).setMimeType(ContentService.MimeType.JSON);
  } 
  
  // Handle GET (Fetching Data for Sync)
  else {
    var storageSheet = ss.getSheetByName("Storage_DB");
    if (!storageSheet) { 
       // If no DB yet, return empty structure
       return ContentService.createTextOutput(JSON.stringify({status: 'empty', data: [], tspr: []})).setMimeType(ContentService.MimeType.JSON);
    }
    
    var data = storageSheet.getRange(1, 1).getValue();
    if (!data) {
       return ContentService.createTextOutput(JSON.stringify({status: 'empty', data: [], tspr: []})).setMimeType(ContentService.MimeType.JSON);
    }

    var parsed = JSON.parse(data);
    
    return ContentService.createTextOutput(JSON.stringify({
       status: 'success', 
       data: parsed.defects || [],
       tspr: parsed.tspr || []
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
`;

// --- COMPONENTS ---

// TSPR Graph Component
const TsprGraph: React.FC<{ data: TsprEntry[] }> = ({ data }) => {
  const width = 800;
  const height = 300;
  const padding = 40;
  const graphWidth = width - padding * 2;
  const graphHeight = height - padding * 2;

  // Map active data to points (only where total > 0)
  const points = data.map((d, i) => {
    if (d.total === 0) return null; // Skip empty slots
    const x = padding + (i / (data.length - 1)) * graphWidth;
    // Formula: (Total - NG) / Total * 100
    const percentage = ((d.total - d.ng) / d.total) * 100;
    const y = height - padding - (percentage / 100) * graphHeight;
    return { x, y, percentage, slot: d.slot };
  }).filter(p => p !== null) as {x: number, y: number, percentage: number, slot: string}[];

  const pathD = points.length > 0 ? `M ${points.map(p => `${p.x},${p.y}`).join(' L ')}` : '';

  // Background gradient path needs at least 2 points to make sense, or use vertical lines for single points
  const fillD = points.length > 1 ? 
    `${pathD} L ${points[points.length-1].x},${height-padding} L ${points[0].x},${height-padding} Z` : '';

  return (
    <div className="w-full overflow-x-auto">
      <div className="min-w-[600px] bg-slate-900 rounded-[2rem] p-6 shadow-2xl">
        <h4 className="text-white font-black uppercase tracking-widest text-xs mb-6 flex items-center gap-2">
          <TrendingUp className="text-green-400" size={16}/> TSPR Hourly Trend
        </h4>
        <svg viewBox={`0 0 ${width} ${height}`} className="w-full h-auto">
          {/* Grid Lines */}
          {[0, 25, 50, 75, 100].map(tick => {
            const y = height - padding - (tick / 100) * graphHeight;
            const yPos = height - padding - (tick / 100) * graphHeight;
            return (
              <g key={tick}>
                <line x1={padding} y1={yPos} x2={width - padding} y2={yPos} stroke="#334155" strokeWidth="1" strokeDasharray="4 4" />
                <text x={padding - 10} y={yPos + 4} fill="#94a3b8" fontSize="10" textAnchor="end" fontWeight="bold">{tick}%</text>
              </g>
            );
          })}

          {/* Time Labels (All Slots) */}
          {data.map((d, i) => {
              const x = padding + (i / (data.length - 1)) * graphWidth;
              return (
                 <text key={i} x={x} y={height - padding + 20} fill="#64748b" fontSize="10" textAnchor="middle" fontWeight="bold" style={{textTransform: 'uppercase'}}>
                    {d.slot.split(' ')[0]}
                 </text>
              );
          })}

          {/* Path */}
          {points.length > 1 && (
             <>
               <path d={pathD} fill="none" stroke="#4ade80" strokeWidth="3" />
               <path d={fillD} fill="url(#gradient)" opacity="0.2" />
             </>
          )}

          <defs>
            <linearGradient id="gradient" x1="0" y1="0" x2="0" y2="1">
              <stop offset="0%" stopColor="#4ade80" />
              <stop offset="100%" stopColor="#4ade80" stopOpacity="0" />
            </linearGradient>
          </defs>

          {/* Points & Labels */}
          {points.map((p, i) => (
            <g key={i}>
              <circle cx={p.x} cy={p.y} r="4" fill="#ffffff" stroke="#4ade80" strokeWidth="2" />
              <text x={p.x} y={p.y - 12} fill="#ffffff" fontSize="12" textAnchor="middle" fontWeight="black">{p.percentage.toFixed(1)}%</text>
            </g>
          ))}
          
          {points.length === 0 && (
             <text x={width/2} y={height/2} fill="#64748b" textAnchor="middle" fontSize="14" fontWeight="bold">Waiting for production data...</text>
          )}
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

  const getPercentage = (total: number, ng: number) => {
    if (total === 0) return '100.0';
    return (((total - ng) / total) * 100).toFixed(1);
  };

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
                 <tr>
                   <th className="py-4 px-4 text-left">Time Slot</th>
                   <th className="py-4 px-4 text-center">Total Vehicles</th>
                   <th className="py-4 px-4 text-center">NG Vehicles</th>
                   <th className="py-4 px-4 text-right">Ratio (%)</th>
                 </tr>
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
                       <td className="py-3 px-4">
                         <input 
                           type="number" 
                           className="w-full text-center bg-slate-100 border-transparent focus:bg-white focus:border-indigo-500 focus:ring-2 focus:ring-indigo-500/20 rounded-xl py-2 font-bold text-slate-800 outline-none transition-all"
                           value={row.total || ''}
                           onChange={(e) => handleChange(i, 'total', e.target.value)}
                           placeholder="0"
                         />
                       </td>
                       <td className="py-3 px-4">
                         <input 
                           type="number" 
                           className="w-full text-center bg-slate-100 border-transparent focus:bg-white focus:border-indigo-500 focus:ring-2 focus:ring-indigo-500/20 rounded-xl py-2 font-bold text-slate-800 outline-none transition-all"
                           value={row.ng || ''}
                           onChange={(e) => handleChange(i, 'ng', e.target.value)}
                           placeholder="0"
                         />
                       </td>
                       <td className={`py-3 px-4 text-right font-black ${color}`}>
                         {pct}%
                       </td>
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
                 <h4 className="opacity-80 font-bold uppercase tracking-widest text-xs mb-1">Current Shift TSPR (Average)</h4>
                 <div className="text-5xl font-black mb-2">
                    {(() => {
                       const total = data.reduce((acc, curr) => acc + curr.total, 0);
                       const ng = data.reduce((acc, curr) => acc + curr.ng, 0);
                       if (total === 0) return '100.0%';
                       return (((total - ng) / total) * 100).toFixed(1) + '%';
                    })()}
                 </div>
                 <div className="flex gap-4 text-xs font-bold opacity-80">
                    <span>Total Produced: {data.reduce((acc, curr) => acc + curr.total, 0)}</span>
                    <span>•</span>
                    <span>Total Defective: {data.reduce((acc, curr) => acc + curr.ng, 0)}</span>
                 </div>
              </div>
              <div className="absolute -right-6 -bottom-6 opacity-20 transform rotate-12">
                 <TrendingUp size={120} />
              </div>
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
            {defect.photoUrl ? (
                <img src={defect.photoUrl} className="w-full h-full object-cover" />
            ) : (
                <div className="flex items-center justify-center w-full h-full text-slate-300"><ImageIcon size={20} /></div>
            )}
        </div>
        <div className="flex-1 min-w-0">
            <div className="flex justify-between items-start">
                <div>
                    <span className="text-[10px] font-black uppercase tracking-widest text-slate-400">{defect.frameNo} • {defect.modelName}</span>
                    <h4 className="font-bold text-slate-800 text-sm truncate">{defect.defectName}</h4>
                </div>
                <div className="text-[10px] font-bold text-slate-300">{new Date(defect.timestamp).toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'})}</div>
            </div>
            <p className="text-xs text-slate-500 line-clamp-2 mt-1 mb-2">{defect.description}</p>
            <div className="flex gap-1 overflow-x-auto pb-1">
                {defect.targetedZones.map(z => {
                    const status = defect.responses[z]?.involved;
                    let color = 'bg-slate-100 text-slate-400 border-slate-200';
                    if (status === 'YES') color = 'bg-red-50 text-red-600 border-red-100';
                    if (status === 'NO') color = 'bg-green-50 text-green-600 border-green-100';
                    return (
                        <span key={z} className={`px-2 py-0.5 rounded-lg border text-[9px] font-black ${color}`}>
                            {z}
                        </span>
                    );
                })}
            </div>
        </div>
      </div>
    );
};

// Manager View Card (Grid Item)
const ManagerDefectThumbnail: React.FC<{ defect: Defect, onClick: () => void }> = ({ defect, onClick }) => (
    <div onClick={onClick} className="group cursor-pointer bg-white rounded-3xl p-3 shadow-md hover:shadow-xl hover:-translate-y-1 transition-all border border-slate-100 h-full flex flex-col">
        <div className="relative aspect-square rounded-2xl overflow-hidden bg-slate-100 mb-3">
            {defect.photoUrl ? (
                <img src={defect.photoUrl} className="w-full h-full object-cover group-hover:scale-105 transition-transform duration-500" />
            ) : (
                <div className="w-full h-full flex items-center justify-center text-slate-300"><ImageIcon size={32} /></div>
            )}
            <div className="absolute top-2 right-2 px-2 py-1 rounded-lg bg-white/90 backdrop-blur text-[10px] font-black uppercase tracking-widest shadow-sm">
                {defect.frameNo}
            </div>
            <div className={`absolute bottom-0 left-0 right-0 p-2 text-[10px] font-black uppercase tracking-widest text-center text-white ${defect.managerAnalysis?.status === 'CLOSED' ? 'bg-green-500' : 'bg-red-500'}`}>
                {defect.managerAnalysis?.status || 'OPEN'}
            </div>
        </div>
        <div className="px-2 pb-2 flex-1">
            <h4 className="font-bold text-slate-800 text-sm truncate">{defect.defectName}</h4>
            <p className="text-xs text-slate-400 font-medium truncate">{defect.modelName}</p>
        </div>
    </div>
);

// Detailed 4M Analysis Sheet Modal
const ManagerDefectModal: React.FC<{
  defect: Defect;
  onClose: () => void;
  onSave: (d: Defect) => void;
}> = ({ defect, onClose, onSave }) => {
  // State for form
  const [analysis, setAnalysis] = useState<ManagerAnalysis>(defect.managerAnalysis || {
    man: '', machine: '', material: '', method: '', rootCauseSummary: '', status: 'OPEN'
  });

  const handleSave = () => {
    const updated: Defect = {
        ...defect,
        managerAnalysis: {
            ...analysis,
            closedAt: analysis.status === 'CLOSED' ? new Date().toISOString() : undefined
        }
    };
    onSave(updated);
    onClose();
  };

  return (
     <div className="fixed inset-0 z-50 bg-slate-900/90 backdrop-blur-sm overflow-y-auto">
        <div className="min-h-screen flex items-center justify-center p-4">
            <div className="bg-white w-full max-w-4xl rounded-[2rem] shadow-2xl overflow-hidden animate-in zoom-in-95 duration-200">
                {/* Header / Image Area */}
                <div className="relative h-96 bg-slate-100 group">
                    {defect.photoUrl ? (
                        <img src={defect.photoUrl} className="w-full h-full object-contain bg-black/5" />
                    ) : (
                        <div className="w-full h-full flex items-center justify-center text-slate-400">
                            <div className="text-center">
                                <ImageIcon size={48} className="mx-auto mb-2 opacity-50"/>
                                <span className="font-bold uppercase tracking-widest text-xs">No Evidence</span>
                            </div>
                        </div>
                    )}
                    
                    {/* Overlay Actions */}
                    <div className="absolute top-0 left-0 right-0 p-6 flex justify-between items-start bg-gradient-to-b from-black/50 to-transparent">
                        <div className="text-white">
                            <h2 className="text-3xl font-black tracking-tighter">{defect.defectName}</h2>
                            <p className="font-medium opacity-90">{defect.frameNo} • {defect.modelName}</p>
                        </div>
                        <button onClick={onClose} className="bg-white/20 hover:bg-white/40 backdrop-blur-md p-2 rounded-full text-white transition-all">
                            <X size={24} />
                        </button>
                    </div>

                    {defect.photoUrl && (
                        <a 
                            href={defect.photoUrl} 
                            download={`Defect_${defect.frameNo}_${defect.id}.jpg`}
                            className="absolute bottom-6 right-6 bg-indigo-600 hover:bg-indigo-500 text-white px-6 py-3 rounded-xl font-black text-xs uppercase tracking-widest shadow-xl flex items-center gap-2 transition-all active:scale-95 z-20"
                            onClick={(e) => e.stopPropagation()}
                        >
                            <Download size={16} /> Download Evidence
                        </a>
                    )}
                </div>

                <div className="grid grid-cols-1 lg:grid-cols-2">
                    {/* Left Column: Info & Leader Reports */}
                    <div className="p-8 bg-slate-50 border-r border-slate-100 space-y-8">
                        <div>
                            <h4 className="flex items-center gap-2 font-black text-slate-800 text-sm uppercase tracking-widest mb-4">
                                <FileText size={16} className="text-indigo-600"/> Defect Details
                            </h4>
                            <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 space-y-4">
                                <div>
                                    <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Description</label>
                                    <p className="text-slate-700 font-medium leading-relaxed mt-1">{defect.description}</p>
                                </div>
                                <div className="flex gap-6">
                                    <div>
                                        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Logged At</label>
                                        <p className="text-slate-700 font-bold mt-1">{new Date(defect.timestamp).toLocaleString()}</p>
                                    </div>
                                    <div>
                                        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Inspector</label>
                                        <p className="text-slate-700 font-bold mt-1">Final Line Check</p>
                                    </div>
                                </div>
                            </div>
                        </div>

                        <div>
                            <h4 className="flex items-center gap-2 font-black text-slate-800 text-sm uppercase tracking-widest mb-4">
                                <Users size={16} className="text-orange-600"/> Group Leader Findings
                            </h4>
                            <div className="space-y-3">
                                {Object.keys(defect.responses).length === 0 ? (
                                    <div className="text-slate-400 text-sm italic p-4 text-center">No zone responses yet.</div>
                                ) : (
                                    (Object.entries(defect.responses) as [string, ZoneResponse][]).map(([zone, resp]) => (
                                        <div key={zone} className={`p-5 rounded-2xl border-l-4 shadow-sm ${resp.involved === 'YES' ? 'bg-red-50 border-red-500' : 'bg-white border-green-500'}`}>
                                            <div className="flex justify-between items-center mb-2">
                                                <span className="font-black text-slate-700">Zone {zone}</span>
                                                <span className={`px-2 py-1 rounded text-[10px] font-black uppercase tracking-widest ${resp.involved === 'YES' ? 'bg-red-100 text-red-700' : 'bg-green-100 text-green-700'}`}>
                                                    {resp.involved === 'YES' ? 'Involved' : 'Not Involved'}
                                                </span>
                                            </div>
                                            {resp.involved === 'YES' && (
                                                <div className="space-y-2 mt-3 text-sm">
                                                    <div><span className="font-bold text-slate-500 text-xs uppercase">Root Cause:</span> <span className="text-slate-800">{resp.reason}</span></div>
                                                    <div><span className="font-bold text-slate-500 text-xs uppercase">Action:</span> <span className="text-slate-800">{resp.actionTaken}</span></div>
                                                    {resp.manpower && resp.manpower.length > 0 && (
                                                        <div className="pt-2 border-t border-red-100 mt-2">
                                                            <span className="font-bold text-slate-500 text-xs uppercase block mb-1">Manpower:</span>
                                                            <div className="flex flex-wrap gap-1">
                                                                {resp.manpower.map((m, i) => (
                                                                    <span key={i} className="bg-white px-2 py-1 rounded border border-red-100 text-xs font-bold text-red-600">{m.name} ({m.ein})</span>
                                                                ))}
                                                            </div>
                                                        </div>
                                                    )}
                                                </div>
                                            )}
                                        </div>
                                    ))
                                )}
                            </div>
                        </div>
                    </div>

                    {/* Right Column: 4M Analysis Form */}
                    <div className="p-8 bg-white flex flex-col h-full">
                        <div className="flex items-center justify-between mb-6">
                            <h4 className="flex items-center gap-2 font-black text-slate-800 text-sm uppercase tracking-widest">
                                <ClipboardList size={16} className="text-indigo-600"/> 4M Analysis Sheet
                            </h4>
                            <div className="flex items-center gap-2 bg-slate-100 p-1 rounded-lg">
                                <button 
                                    onClick={() => setAnalysis({...analysis, status: 'OPEN'})}
                                    className={`px-3 py-1.5 rounded-md text-[10px] font-black uppercase tracking-widest transition-all ${analysis.status === 'OPEN' ? 'bg-white text-red-600 shadow-sm' : 'text-slate-400 hover:text-slate-600'}`}
                                >Open</button>
                                <button 
                                    onClick={() => setAnalysis({...analysis, status: 'CLOSED'})}
                                    className={`px-3 py-1.5 rounded-md text-[10px] font-black uppercase tracking-widest transition-all ${analysis.status === 'CLOSED' ? 'bg-green-500 text-white shadow-sm' : 'text-slate-400 hover:text-slate-600'}`}
                                >Closed</button>
                            </div>
                        </div>

                        <div className="flex-1 space-y-6 overflow-y-auto pr-2 custom-scrollbar">
                            <div className="grid grid-cols-2 gap-4">
                                <div className="space-y-1">
                                    <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest px-1">Man</label>
                                    <textarea 
                                        value={analysis.man}
                                        onChange={e => setAnalysis({...analysis, man: e.target.value})}
                                        className="w-full bg-slate-50 border border-slate-200 rounded-xl p-3 text-sm font-medium focus:bg-white focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 outline-none transition-all"
                                        rows={3}
                                        placeholder="Human factors..."
                                    />
                                </div>
                                <div className="space-y-1">
                                    <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest px-1">Machine</label>
                                    <textarea 
                                        value={analysis.machine}
                                        onChange={e => setAnalysis({...analysis, machine: e.target.value})}
                                        className="w-full bg-slate-50 border border-slate-200 rounded-xl p-3 text-sm font-medium focus:bg-white focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 outline-none transition-all"
                                        rows={3}
                                        placeholder="Equipment factors..."
                                    />
                                </div>
                                <div className="space-y-1">
                                    <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest px-1">Material</label>
                                    <textarea 
                                        value={analysis.material}
                                        onChange={e => setAnalysis({...analysis, material: e.target.value})}
                                        className="w-full bg-slate-50 border border-slate-200 rounded-xl p-3 text-sm font-medium focus:bg-white focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 outline-none transition-all"
                                        rows={3}
                                        placeholder="Part/Component factors..."
                                    />
                                </div>
                                <div className="space-y-1">
                                    <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest px-1">Method</label>
                                    <textarea 
                                        value={analysis.method}
                                        onChange={e => setAnalysis({...analysis, method: e.target.value})}
                                        className="w-full bg-slate-50 border border-slate-200 rounded-xl p-3 text-sm font-medium focus:bg-white focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 outline-none transition-all"
                                        rows={3}
                                        placeholder="Process/Standard factors..."
                                    />
                                </div>
                            </div>

                            <div className="space-y-1">
                                <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest px-1">Root Cause Summary & Conclusion</label>
                                <textarea 
                                    value={analysis.rootCauseSummary}
                                    onChange={e => setAnalysis({...analysis, rootCauseSummary: e.target.value})}
                                    className="w-full bg-indigo-50 border border-indigo-100 rounded-xl p-4 text-sm font-medium focus:bg-white focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 outline-none transition-all"
                                    rows={4}
                                    placeholder="Final conclusion on the defect origin and countermeasure validation..."
                                />
                            </div>
                        </div>

                        <div className="pt-6 mt-6 border-t border-slate-100">
                            <button 
                                onClick={handleSave}
                                className="w-full bg-slate-900 text-white py-4 rounded-xl font-black text-sm uppercase tracking-widest shadow-xl hover:bg-slate-800 active:scale-95 transition-all"
                            >
                                Save Analysis Sheet
                            </button>
                        </div>
                    </div>
                </div>
            </div>
        </div>
     </div>
  );
};

const SettingsModal: React.FC<{ currentUrl: string, onSave: (url: string) => void, onClose: () => void }> = ({ currentUrl, onSave, onClose }) => {
    const [url, setUrl] = useState(currentUrl);
    const [copied, setCopied] = useState(false);

    const copyScript = () => {
      navigator.clipboard.writeText(SCRIPT_CODE);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    };

    return (
      <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/50 backdrop-blur-sm animate-in fade-in">
        <div className="bg-white p-6 rounded-[2rem] shadow-2xl w-full max-w-2xl max-h-[90vh] overflow-y-auto">
           <div className="flex justify-between items-center mb-6">
             <h3 className="text-xl font-black text-slate-800 tracking-tight">System Configuration</h3>
             <button onClick={onClose} className="p-2 bg-slate-100 rounded-full"><X size={20} className="text-slate-400" /></button>
           </div>
           
           <div className="space-y-8">
             <div className="p-4 bg-blue-50 border border-blue-100 rounded-2xl">
                 <h4 className="font-bold text-blue-900 text-sm mb-2">How to Enable Multi-User Sync</h4>
                 <ol className="list-decimal list-inside text-xs text-blue-800 space-y-2">
                    <li>Copy the code below.</li>
                    <li>Go to <a href="https://script.google.com" target="_blank" className="underline font-bold">script.google.com</a> and create a new project.</li>
                    <li>Paste the code into <b>Code.gs</b> and save.</li>
                    <li>Click <b>Deploy</b> {'>'} <b>New Deployment</b>.</li>
                    <li>Select type: <b>Web app</b>.</li>
                    <li>Set <b>Who has access</b> to <b>"Anyone"</b> (Important for the app to connect).</li>
                    <li>Click Deploy and copy the <b>Web App URL</b>.</li>
                    <li>Paste the URL in the box below and save. Share this URL with all users!</li>
                 </ol>
             </div>

             <div className="space-y-4">
               <div className="p-4 bg-indigo-50 border border-indigo-100 rounded-2xl flex gap-4">
                  <div className="bg-indigo-100 p-3 rounded-xl h-fit"><Database className="text-indigo-600" size={20} /></div>
                  <div>
                    <h4 className="font-bold text-indigo-900 text-sm mb-1">Live Database Connection</h4>
                    <p className="text-xs text-indigo-700/80 leading-relaxed">Connect to a Google Sheet to enable real-time dashboard updates and persistent storage.</p>
                  </div>
               </div>
               <div className="space-y-2">
                 <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest px-1">Web App URL</label>
                 <input value={url} onChange={e => setUrl(e.target.value)} placeholder="https://script.google.com/macros/s/..." className="w-full p-4 bg-slate-50 border border-slate-200 rounded-2xl text-xs font-bold focus:ring-4 focus:ring-indigo-500/10 focus:border-indigo-500 outline-none"/>
               </div>
             </div>

             <div className="border-t border-slate-100 pt-6">
                <h4 className="font-black text-slate-800 text-sm mb-2">Google Apps Script Configuration</h4>
                
                <div className="mb-4">
                   <h5 className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-2">Sheets Created:</h5>
                   <div className="flex flex-wrap gap-2">
                        <span className="px-2 py-1 bg-green-100 text-green-700 text-[10px] font-bold rounded border border-green-200">Defects</span>
                        <span className="px-2 py-1 bg-blue-100 text-blue-700 text-[10px] font-bold rounded border border-blue-200">TSPR</span>
                        <span className="px-2 py-1 bg-slate-100 text-slate-500 text-[10px] font-bold rounded border border-slate-200">Storage_DB (Hidden)</span>
                   </div>
                </div>

                <div className="relative group">
                  <div className="absolute top-2 right-2">
                    <button onClick={copyScript} className={`flex items-center gap-2 px-3 py-1.5 rounded-lg text-[10px] font-black uppercase tracking-widest transition-all ${copied ? 'bg-green-500 text-white' : 'bg-white text-slate-600 shadow-sm hover:bg-slate-50'}`}>
                      {copied ? <Check size={12} /> : <Copy size={12} />}
                      {copied ? 'Copied!' : 'Copy Code'}
                    </button>
                  </div>
                  <pre className="bg-slate-900 text-slate-300 p-4 rounded-2xl text-[10px] font-mono overflow-x-auto h-48 border border-slate-800">
                    {SCRIPT_CODE}
                  </pre>
                </div>
             </div>
           </div>

           <div className="flex gap-3 mt-8 pt-4 border-t border-slate-100">
             <button onClick={() => onSave(url)} className="flex-1 bg-slate-900 text-white py-4 rounded-2xl font-black text-xs uppercase tracking-widest shadow-xl active:scale-95 transition-all">Save Connection</button>
           </div>
        </div>
      </div>
    );
};

const InspectorView: React.FC<{ defects: Defect[], onAddDefect: (d: Defect) => void }> = ({ defects, onAddDefect }) => {
    // ... existing InspectorView logic ...
    const [showForm, setShowForm] = useState(false);
    const [formData, setFormData] = useState({
      frameNo: '',
      modelName: '',
      defectName: '',
      description: '',
      targetedZones: [] as ZoneId[],
      photoUrl: ''
    });
  
    const handleSubmit = (e: React.FormEvent) => {
      e.preventDefault();
      if (formData.targetedZones.length === 0) {
        alert("Please select at least one possible zone");
        return;
      }
      const newDefect: Defect = {
        id: Date.now().toString(),
        ...formData,
        timestamp: new Date().toISOString(),
        responses: formData.targetedZones.reduce((acc, zone) => ({
          ...acc,
          [zone]: { zoneId: zone, involved: 'PENDING', manpower: [] }
        }), {})
      };
      onAddDefect(newDefect);
      setFormData({ frameNo: '', modelName: '', defectName: '', description: '', targetedZones: [], photoUrl: '' });
      setShowForm(false);
    };
  
    const toggleZone = (zone: ZoneId) => {
      setFormData(prev => ({
        ...prev,
        targetedZones: prev.targetedZones.includes(zone) 
          ? prev.targetedZones.filter(z => z !== zone)
          : [...prev.targetedZones, zone]
      }));
    };
  
    const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      if (file) {
        const reader = new FileReader();
        reader.onloadend = () => {
          setFormData(prev => ({ ...prev, photoUrl: reader.result as string }));
        };
        reader.readAsDataURL(file);
      }
    };
  
    if (showForm) {
      return (
        <div className="animate-in fade-in slide-in-from-bottom-4 duration-300 max-w-2xl mx-auto">
          <div className="flex items-center justify-between mb-6 px-1">
            <div>
              <h2 className="text-2xl font-black text-slate-800 tracking-tighter">Reporting Tool</h2>
              <p className="text-xs font-bold text-slate-400 uppercase tracking-widest">Identify & Dispatch</p>
            </div>
            <button onClick={() => setShowForm(false)} className="bg-slate-200/50 p-3 rounded-2xl text-slate-400 active:scale-90 transition-all"><X size={20} /></button>
          </div>
          
          <form onSubmit={handleSubmit} className="space-y-6 bg-white p-6 rounded-[2.5rem] shadow-2xl shadow-blue-500/5 border border-slate-100">
            <div className="grid grid-cols-2 gap-4">
              <div className="space-y-1">
                <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest px-2">Vehicle Frame No.</label>
                <input required className="w-full px-5 py-4 rounded-3xl bg-slate-50 border border-slate-200 focus:bg-white focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500 transition-all outline-none font-bold" placeholder="VF-0000" value={formData.frameNo} onChange={e => setFormData({...formData, frameNo: e.target.value})} />
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest px-2">Model Name</label>
                <input required className="w-full px-5 py-4 rounded-3xl bg-slate-50 border border-slate-200 focus:bg-white focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500 transition-all outline-none font-bold" placeholder="SUV-X" value={formData.modelName} onChange={e => setFormData({...formData, modelName: e.target.value})} />
              </div>
            </div>
  
            <div className="space-y-1">
              <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest px-2">Defect Category</label>
              <input required className="w-full px-5 py-4 rounded-3xl bg-slate-50 border border-slate-200 focus:bg-white focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500 transition-all outline-none font-bold" placeholder="E.g. Paint Scratch, Missing Bolt" value={formData.defectName} onChange={e => setFormData({...formData, defectName: e.target.value})} />
            </div>
  
            <div className="space-y-1">
              <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest px-2">Detailed Notes</label>
              <textarea className="w-full px-5 py-4 rounded-3xl bg-slate-50 border border-slate-200 focus:bg-white focus:ring-4 focus:ring-blue-500/10 focus:border-blue-500 transition-all outline-none font-medium text-sm" rows={3} placeholder="Describe the discrepancy precisely..." value={formData.description} onChange={e => setFormData({...formData, description: e.target.value})} />
            </div>
  
            <div className="space-y-2">
              <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest px-2">Visual Evidence</label>
              <div className="grid grid-cols-2 gap-4">
                 <label className="flex flex-col items-center justify-center border-2 border-dashed border-slate-200 rounded-3xl p-6 cursor-pointer hover:bg-blue-50 hover:border-blue-200 transition-all bg-slate-50/50 group">
                    <Camera className="text-blue-500 mb-2 group-hover:scale-110 transition-transform" size={28} />
                    <span className="text-[10px] font-black text-slate-500 uppercase tracking-widest">Take Photo</span>
                    <input type="file" accept="image/*" capture="environment" className="hidden" onChange={handleFileUpload} />
                 </label>
                 <label className="flex flex-col items-center justify-center border-2 border-dashed border-slate-200 rounded-3xl p-6 cursor-pointer hover:bg-indigo-50 hover:border-indigo-200 transition-all bg-slate-50/50 group">
                    <ImageIcon className="text-indigo-500 mb-2 group-hover:scale-110 transition-transform" size={28} />
                    <span className="text-[10px] font-black text-slate-500 uppercase tracking-widest">Upload File</span>
                    <input type="file" accept="image/*" className="hidden" onChange={handleFileUpload} />
                 </label>
              </div>
              {formData.photoUrl && (
                <div className="relative w-full aspect-video rounded-[2rem] overflow-hidden border-4 border-white shadow-2xl group mt-2">
                  <img src={formData.photoUrl} className="w-full h-full object-cover" />
                  <button type="button" onClick={() => setFormData({...formData, photoUrl: ''})} className="absolute top-4 right-4 bg-red-500 text-white p-3 rounded-full shadow-lg opacity-0 group-hover:opacity-100 transition-all hover:scale-110"><X size={20} /></button>
                </div>
              )}
            </div>
  
            <div className="space-y-3">
              <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest px-2">Probable Error Source (Target Zones)</label>
              <div className="grid grid-cols-3 md:grid-cols-5 gap-2">
                {ZONES.map(z => (
                  <button
                    key={z}
                    type="button"
                    onClick={() => toggleZone(z)}
                    className={`py-4 text-sm font-black rounded-2xl border-2 transition-all ${formData.targetedZones.includes(z) ? 'bg-blue-600 border-blue-600 text-white shadow-xl shadow-blue-500/20' : 'border-slate-100 bg-slate-50 text-slate-400 active:scale-95'}`}
                  >
                    {z}
                  </button>
                ))}
              </div>
            </div>
  
            <button type="submit" className="w-full bg-slate-900 text-white py-6 rounded-3xl font-black text-lg shadow-2xl active:scale-[0.98] transition-all uppercase tracking-tighter">
              DISPATCH REPORT
            </button>
          </form>
        </div>
      );
    }
  
    return (
      <div className="max-w-2xl mx-auto">
        <div className="flex items-center justify-between mb-8 px-1">
          <div>
            <h2 className="text-3xl font-black text-slate-800 tracking-tighter">Inspection Log</h2>
            <p className="text-xs font-bold text-slate-400 uppercase tracking-widest">Active Assembly Line Track</p>
          </div>
          <button onClick={() => setShowForm(true)} className="bg-blue-600 text-white p-4 rounded-3xl font-black flex items-center shadow-2xl shadow-blue-500/20 active:scale-90 transition-all">
            <Plus size={24} />
          </button>
        </div>
  
        <div className="space-y-6">
          {defects.length === 0 ? (
            <div className="text-center py-40 bg-white rounded-[3rem] border-4 border-dashed border-slate-100 shadow-inner">
              <div className="bg-slate-50 w-20 h-20 rounded-full flex items-center justify-center mx-auto mb-6">
                 <ClipboardList className="text-slate-200" size={40} />
              </div>
              <p className="text-slate-400 font-bold uppercase tracking-widest text-[10px]">No active defects reported</p>
            </div>
          ) : (
            defects.map(d => (
              <DefectCard key={d.id} defect={d} />
            ))
          )}
        </div>
      </div>
    );
};

const LeaderDefectTask: React.FC<{ 
  defect: Defect, 
  zoneId: ZoneId, 
  onSave: (resp: ZoneResponse) => void 
}> = ({ defect, zoneId, onSave }) => {
  const response = defect.responses[zoneId] as ZoneResponse | undefined;
  const isDone = response && response.involved !== 'PENDING';
  const [isEditing, setIsEditing] = useState(!isDone);
  const [involved, setInvolved] = useState<'YES' | 'NO' | 'PENDING'>(response?.involved || 'PENDING');
  const [reason, setReason] = useState(response?.reason || '');
  const [actionTaken, setActionTaken] = useState(response?.actionTaken || '');
  const [manpower, setManpower] = useState<ManpowerEntry[]>(response?.manpower?.length ? response.manpower : [{ name: '', ein: '' }]);
  const [isExpanded, setIsExpanded] = useState(false);

  const handleSave = () => {
    if (involved === 'YES') {
      if (!reason || !actionTaken || manpower.some(m => !m.name || !m.ein)) {
        alert("Incomplete analysis. All fields are required for 4M documentation.");
        return;
      }
    }
    
    onSave({
      zoneId,
      involved,
      reason: involved === 'YES' ? reason : 'N/A',
      actionTaken: involved === 'YES' ? actionTaken : 'N/A',
      manpower: involved === 'YES' ? manpower : [],
      timestamp: new Date().toISOString()
    });
    setIsEditing(false);
  };

  const markAsNA = () => {
    if (confirm("Confirm that Zone " + zoneId + " is NOT involved in this defect?")) {
        setInvolved('NO');
        onSave({
            zoneId,
            involved: 'NO',
            reason: 'N/A',
            actionTaken: 'N/A',
            manpower: [],
            timestamp: new Date().toISOString()
        });
        setIsEditing(false);
    }
  };

  return (
    <>
    {isExpanded && defect.photoUrl && (
        <div className="fixed inset-0 z-50 bg-black/95 flex items-center justify-center p-4 animate-in fade-in duration-200 cursor-zoom-out" onClick={() => setIsExpanded(false)}>
           <button onClick={() => setIsExpanded(false)} className="absolute top-6 right-6 text-white/50 hover:text-white p-2 transition-colors">
             <X size={32} />
           </button>
           <img src={defect.photoUrl} className="max-w-full max-h-full object-contain rounded-lg shadow-2xl" />
        </div>
    )}
    <div className="bg-white rounded-[2.5rem] shadow-xl shadow-slate-200/50 border border-slate-100 overflow-hidden relative">
      <div className="relative h-[500px] bg-slate-100 group cursor-pointer" onClick={() => defect.photoUrl && setIsExpanded(true)}>
         {defect.photoUrl ? (
            <>
                <img src={defect.photoUrl} className="w-full h-full object-cover transition-transform duration-700 group-hover:scale-105" />
                <div className="absolute inset-0 flex items-center justify-center bg-black/10 opacity-0 group-hover:opacity-100 transition-opacity">
                    <div className="bg-white/20 backdrop-blur-md p-4 rounded-full border border-white/30 text-white shadow-lg">
                        <Maximize2 size={32} />
                    </div>
                </div>
            </>
         ) : (
            <div className="w-full h-full flex items-center justify-center bg-slate-200">
                <ImageIcon className="text-slate-400 w-16 h-16" />
            </div>
         )}
      </div>

      <div className="px-6 pt-6">
          <div className="flex items-center justify-between mb-2">
              <div className="flex items-center gap-2">
                  <span className="px-3 py-1 bg-slate-100 rounded-lg text-slate-500 text-[10px] font-black uppercase tracking-widest border border-slate-200">{defect.frameNo}</span>
                  <span className="px-3 py-1 bg-orange-500 rounded-lg text-white text-[10px] font-black uppercase tracking-widest">{defect.modelName}</span>
              </div>
              <div className="flex items-center gap-1 text-slate-400">
                  <Clock size={14} />
                  <span className="text-[10px] font-bold uppercase tracking-widest">
                    {new Date(defect.timestamp).toLocaleString([], { month: 'short', day: 'numeric', hour: '2-digit', minute:'2-digit' })}
                  </span>
              </div>
          </div>
          <h3 className="text-2xl font-black text-slate-800 leading-tight mb-1">{defect.defectName}</h3>
          <p className="text-slate-500 text-sm font-medium line-clamp-2">"{defect.description}"</p>
          
          <div className="flex flex-wrap gap-2 mt-3">
              <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest mr-1 self-center">Targeted:</span>
              {defect.targetedZones.map(z => (
                  <span key={z} className={`px-2 py-1 rounded text-[10px] font-black border ${
                      z === zoneId 
                      ? 'bg-red-600 border-red-500 text-white shadow-lg shadow-red-500/30' 
                      : 'bg-slate-100 border-slate-200 text-slate-400'
                  }`}>
                      {z}
                  </span>
              ))}
          </div>
      </div>

      <div className="p-6 pt-4">
        {!isEditing ? (
            <div className="flex flex-col gap-4">
               <div className={`p-6 rounded-[2rem] border-2 flex flex-col items-center justify-center text-center gap-2 ${response?.involved === 'YES' ? 'bg-red-50 border-red-100 text-red-600' : 'bg-green-50 border-green-100 text-green-600'}`}>
                  {response?.involved === 'YES' ? <AlertTriangle size={32} /> : <CheckCircle2 size={32} />}
                  <h4 className="font-black uppercase tracking-widest text-sm">Status: {response?.involved === 'YES' ? 'Root Cause Identified' : 'Not Involved'}</h4>
               </div>
               <button onClick={() => setIsEditing(true)} className="w-full py-4 text-slate-400 font-bold text-xs uppercase tracking-widest hover:text-slate-600 transition-colors">Edit Analysis</button>
            </div>
        ) : (
          <div className="space-y-6 animate-in fade-in duration-500">
             
             {(involved === 'PENDING' || involved === 'NO') && (
                 <div className="grid grid-cols-2 gap-4">
                    <button onClick={() => setInvolved('YES')} className="py-8 rounded-[2rem] bg-slate-50 border-2 border-slate-100 hover:border-red-500 hover:bg-red-50 transition-all group">
                        <AlertTriangle className="mx-auto mb-2 text-slate-300 group-hover:text-red-500" size={32} />
                        <span className="block text-center font-black text-slate-500 group-hover:text-red-600 text-sm uppercase tracking-widest">Involved</span>
                    </button>
                    <button onClick={markAsNA} className="py-8 rounded-[2rem] bg-slate-50 border-2 border-slate-100 hover:border-green-500 hover:bg-green-50 transition-all group">
                        <CheckCircle2 className="mx-auto mb-2 text-slate-300 group-hover:text-green-500" size={32} />
                        <span className="block text-center font-black text-slate-500 group-hover:text-green-600 text-sm uppercase tracking-widest">Not Involved (NA)</span>
                    </button>
                 </div>
             )}

             {involved === 'YES' && (
               <div className="space-y-6 animate-in slide-in-from-bottom-4">
                  <div className="flex items-center justify-between">
                     <h4 className="font-black text-slate-800 uppercase tracking-widest text-xs">4M Root Cause Analysis</h4>
                     <button onClick={() => setInvolved('PENDING')} className="text-slate-400 text-xs font-bold hover:text-slate-600">Back</button>
                  </div>
                  
                  <div className="space-y-2">
                    <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest px-2">Root Cause Reason</label>
                    <textarea 
                        className="w-full px-5 py-4 rounded-[1.5rem] bg-slate-50 border border-slate-200 focus:bg-white focus:ring-4 focus:ring-orange-500/10 focus:border-orange-500 transition-all outline-none text-sm font-medium" 
                        placeholder="Why did this happen?" 
                        rows={2}
                        value={reason} 
                        onChange={e => setReason(e.target.value)} 
                    />
                  </div>

                  <div className="space-y-2">
                    <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest px-2">Containment Action</label>
                    <textarea 
                        className="w-full px-5 py-4 rounded-[1.5rem] bg-slate-50 border border-slate-200 focus:bg-white focus:ring-4 focus:ring-orange-500/10 focus:border-orange-500 transition-all outline-none text-sm font-medium" 
                        placeholder="Immediate fix applied..." 
                        rows={2}
                        value={actionTaken} 
                        onChange={e => setActionTaken(e.target.value)} 
                    />
                  </div>

                  <div className="space-y-3">
                     <div className="flex justify-between items-center px-2">
                        <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Responsible Manpower</label>
                        <button onClick={() => setManpower([...manpower, { name: '', ein: '' }])} className="w-6 h-6 rounded-full bg-slate-100 flex items-center justify-center text-slate-600 hover:bg-orange-100 hover:text-orange-600 transition-colors"><Plus size={14} /></button>
                     </div>
                     {manpower.map((m, idx) => (
                        <div key={idx} className="flex gap-2">
                           <input className="flex-1 px-4 py-3 rounded-xl bg-slate-50 border border-slate-200 text-xs font-bold outline-none focus:bg-white" placeholder="Name" value={m.name} onChange={e => {
                               const newM = manpower.map((item, i) => i === idx ? { ...item, name: e.target.value } : item); 
                               setManpower(newM);
                           }}/>
                           <input className="w-24 px-4 py-3 rounded-xl bg-slate-50 border border-slate-200 text-xs font-bold outline-none focus:bg-white" placeholder="EIN" value={m.ein} onChange={e => {
                               const newM = manpower.map((item, i) => i === idx ? { ...item, ein: e.target.value } : item);
                               setManpower(newM);
                           }}/>
                        </div>
                     ))}
                  </div>

                  <button onClick={handleSave} className="w-full bg-slate-900 text-white py-5 rounded-[1.5rem] font-black text-xs uppercase tracking-widest shadow-xl active:scale-[0.98] transition-all">
                    Submit Analysis
                  </button>
               </div>
             )}
          </div>
        )}
      </div>
    </div>
    </>
  );
};

const App: React.FC = () => {
  const [session, setSession] = useState<UserSession | null>(() => {
    const saved = localStorage.getItem(SESSION_KEY);
    return saved ? JSON.parse(saved) : null;
  });

  const [defects, setDefects] = useState<Defect[]>(() => {
    const saved = localStorage.getItem(LOCAL_STORAGE_KEY);
    return saved ? JSON.parse(saved) : [];
  });

  const [tsprData, setTsprData] = useState<TsprEntry[]>(() => {
    const saved = localStorage.getItem(TSPR_STORAGE_KEY);
    return saved ? JSON.parse(saved) : INITIAL_TSPR_DATA;
  });

  const [cloudUrl, setCloudUrl] = useState(() => localStorage.getItem(CLOUD_URL_KEY) || '');
  const [showSettings, setShowSettings] = useState(false);
  const [isSyncing, setIsSyncing] = useState(false);
  const [selectedDefect, setSelectedDefect] = useState<Defect | null>(null);
  const [managerTab, setManagerTab] = useState<'DEFECTS' | 'TSPR'>('DEFECTS');

  useEffect(() => {
    localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(defects));
  }, [defects]);

  useEffect(() => {
    localStorage.setItem(TSPR_STORAGE_KEY, JSON.stringify(tsprData));
  }, [tsprData]);

  // Fetch data on mount if URL is present (Basic Multi-user support)
  useEffect(() => {
    if (cloudUrl && session) {
      handleRefresh();
    }
  }, [cloudUrl, session]);

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
    } catch (e) {
      console.error("Failed to load data", e);
    } finally {
      setIsSyncing(false);
    }
  };

  const handleSync = async () => {
    if (!cloudUrl) {
      setShowSettings(true);
      return;
    }
    setIsSyncing(true);
    try {
      await fetch(cloudUrl, {
        method: 'POST',
        body: JSON.stringify({ defects, tspr: tsprData })
      });
      alert("Synced successfully!");
    } catch (e) {
      alert("Sync failed. Check connection.");
    } finally {
      setIsSyncing(false);
    }
  };

  const handleLogin = (role: AppRole, zoneId?: ZoneId) => {
    const newSession: UserSession = {
      role,
      zoneId,
      name: role === 'GROUP_LEADER' ? `Leader ${zoneId}` : role === 'INSPECTOR' ? 'Inspector' : 'Manager'
    };
    setSession(newSession);
    localStorage.setItem(SESSION_KEY, JSON.stringify(newSession));
  };

  const handleLogout = () => {
    setSession(null);
    localStorage.removeItem(SESSION_KEY);
  };

  if (!session) {
    return (
      <div className="min-h-screen bg-slate-50 flex items-center justify-center p-4">
        <div className="max-w-md w-full space-y-8">
          <div className="text-center">
            <div className="inline-flex p-4 bg-indigo-600 rounded-[2rem] text-white shadow-xl shadow-indigo-500/30 mb-6">
              <Layers size={40} />
            </div>
            <h1 className="text-4xl font-black text-slate-800 tracking-tighter mb-2">AutoLine <span className="text-indigo-600">Pro</span></h1>
            <p className="text-xs font-bold text-slate-400 uppercase tracking-widest">Quality Management System</p>
          </div>

          <div className="space-y-4">
            {(Object.entries(ROLE_CONFIG) as [AppRole, typeof ROLE_CONFIG[AppRole]][]).map(([role, config]) => (
              <div 
                key={role}
                className="group bg-white p-2 rounded-[2.5rem] border-2 border-slate-100 hover:border-indigo-500 transition-all shadow-lg hover:shadow-xl hover:shadow-indigo-500/10 overflow-hidden"
              >
                <div 
                  className="flex items-center gap-4 p-4 cursor-pointer"
                  onClick={() => role !== 'GROUP_LEADER' && handleLogin(role)}
                >
                  <div className={`p-4 rounded-2xl ${config.color} text-white shadow-lg group-hover:scale-110 transition-transform`}>
                    {config.icon}
                  </div>
                  <div className="flex-1">
                    <h3 className="font-black text-slate-800 text-lg">{config.title}</h3>
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Select Access Level</p>
                  </div>
                  {role !== 'GROUP_LEADER' && <ChevronRight className="text-slate-300 group-hover:text-indigo-500" />}
                </div>

                {role === 'GROUP_LEADER' && (
                  <div className="px-4 pb-4 pt-2 border-t border-slate-50">
                     <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-3 ml-2">Select Your Zone</p>
                     <div className="grid grid-cols-4 gap-2">
                        {ZONES.map(z => (
                          <button
                            key={z}
                            onClick={() => handleLogin('GROUP_LEADER', z)}
                            className="py-2 rounded-xl bg-slate-50 text-slate-600 text-xs font-black hover:bg-orange-500 hover:text-white transition-colors"
                          >
                            {z}
                          </button>
                        ))}
                     </div>
                  </div>
                )}
              </div>
            ))}
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50/50 pb-20">
      <div className="sticky top-0 z-40 bg-white/80 backdrop-blur-md border-b border-slate-200 px-6 py-4 mb-8">
         <div className="max-w-7xl mx-auto flex items-center justify-between">
            <div className="flex items-center gap-4">
               <div className={`p-2.5 rounded-xl ${ROLE_CONFIG[session.role].color} text-white shadow-lg shadow-indigo-500/20`}>
                  {ROLE_CONFIG[session.role].icon}
               </div>
               <div>
                  <h1 className="font-black text-slate-800 text-lg leading-none mb-1">{ROLE_CONFIG[session.role].title}</h1>
                  <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest flex items-center gap-2">
                    {session.name} <span className="w-1 h-1 rounded-full bg-green-500"></span> Online
                  </p>
               </div>
            </div>

            <div className="flex items-center gap-2">
               <button onClick={handleRefresh} disabled={isSyncing} className="p-3 rounded-2xl bg-white text-slate-600 hover:bg-slate-50 hover:text-blue-600 transition-colors shadow-sm relative group border border-slate-100">
                  <RefreshCw size={20} className={isSyncing ? 'animate-spin' : ''} />
               </button>
               <button onClick={handleSync} disabled={isSyncing} className="p-3 rounded-2xl bg-slate-100 text-slate-600 hover:bg-indigo-50 hover:text-indigo-600 transition-colors relative group">
                  <Cloud size={20} className={isSyncing ? 'animate-bounce' : ''} />
               </button>
               <button onClick={() => setShowSettings(true)} className="p-3 rounded-2xl bg-slate-100 text-slate-600 hover:bg-slate-200 transition-colors">
                  <Settings size={20} />
               </button>
               <div className="w-px h-8 bg-slate-200 mx-2"></div>
               <button onClick={handleLogout} className="p-3 rounded-2xl bg-red-50 text-red-600 hover:bg-red-500 hover:text-white transition-colors shadow-sm">
                  <LogOut size={20} />
               </button>
            </div>
         </div>
      </div>

      <div className="max-w-7xl mx-auto px-4 sm:px-6">
         {session.role === 'INSPECTOR' && (
            <InspectorView defects={defects} onAddDefect={d => setDefects([d, ...defects])} />
         )}

         {session.role === 'GROUP_LEADER' && session.zoneId && (
            <div className="space-y-10">
               <div>
                  <div className="flex items-center justify-between mb-6">
                     <div>
                       <h2 className="text-2xl font-black text-slate-800 tracking-tighter">Zone Operations</h2>
                       <p className="text-xs font-bold text-slate-400 uppercase tracking-widest">Global Defect List • Logged as {session.zoneId}</p>
                     </div>
                  </div>
                  <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-6">
                     {defects.map(d => (
                        <LeaderDefectTask 
                          key={d.id} 
                          defect={d} 
                          zoneId={session.zoneId!}
                          onSave={(resp) => {
                             setDefects(defects.map(curr => curr.id === d.id ? { ...curr, responses: { ...curr.responses, [session.zoneId!]: resp } } : curr));
                          }}
                        />
                     ))}
                     {defects.length === 0 && (
                        <div className="col-span-full py-20 text-center border-4 border-dashed border-slate-100 rounded-[3rem] bg-white">
                           <CheckCircle2 className="mx-auto text-green-500 mb-4" size={64} />
                           <h3 className="font-black text-slate-800 text-xl">All Clear!</h3>
                           <p className="text-slate-400 font-bold uppercase tracking-widest text-xs mt-2">No defects reported in the system</p>
                        </div>
                     )}
                  </div>
               </div>
            </div>
         )}

         {session.role === 'MANAGER' && (
            <div className="space-y-8">
               <div className="flex p-1 bg-slate-200/50 rounded-2xl w-fit mx-auto">
                  <button 
                    onClick={() => setManagerTab('DEFECTS')}
                    className={`px-8 py-3 rounded-xl font-black text-xs uppercase tracking-widest transition-all ${managerTab === 'DEFECTS' ? 'bg-white text-indigo-600 shadow-lg' : 'text-slate-500 hover:text-slate-700'}`}
                  >
                    Defect Sheet
                  </button>
                  <button 
                    onClick={() => setManagerTab('TSPR')}
                    className={`px-8 py-3 rounded-xl font-black text-xs uppercase tracking-widest transition-all ${managerTab === 'TSPR' ? 'bg-white text-indigo-600 shadow-lg' : 'text-slate-500 hover:text-slate-700'}`}
                  >
                    TSPR Sheet
                  </button>
               </div>

               {managerTab === 'DEFECTS' && (
                  <div className="animate-in fade-in slide-in-from-bottom-4">
                     <div className="flex justify-between items-end mb-6">
                        <div>
                            <h2 className="text-3xl font-black text-slate-800 tracking-tighter">Master Defect Sheet</h2>
                            <p className="text-xs font-bold text-slate-400 uppercase tracking-widest">Real-time Visualization & Analysis</p>
                        </div>
                        <div className="flex gap-2">
                           <button onClick={() => downloadReport(defects)} className="py-2 px-4 bg-indigo-600 text-white rounded-xl font-black uppercase tracking-widest text-[10px] shadow-lg shadow-indigo-500/20 hover:bg-indigo-700 active:scale-95 transition-all flex items-center gap-2">
                              <Download size={14} /> Excel
                           </button>
                           <button onClick={() => downloadJSON(defects)} className="py-2 px-4 bg-slate-800 text-white rounded-xl font-black uppercase tracking-widest text-[10px] hover:bg-slate-900 active:scale-95 transition-all">
                              Backup JSON
                           </button>
                        </div>
                     </div>
                     
                     <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 xl:grid-cols-5 gap-6">
                        {defects.length === 0 ? (
                            <div className="col-span-full py-20 text-center text-slate-400 font-bold uppercase tracking-widest text-xs border-4 border-dashed border-slate-100 rounded-[2rem]">Database Empty</div>
                        ) : (
                            defects.map(d => (
                                <ManagerDefectThumbnail 
                                    key={d.id} 
                                    defect={d} 
                                    onClick={() => setSelectedDefect(d)}
                                />
                            ))
                        )}
                     </div>
                  </div>
               )}

               {managerTab === 'TSPR' && (
                  <div className="animate-in fade-in slide-in-from-bottom-4">
                      <div className="mb-6">
                            <h2 className="text-3xl font-black text-slate-800 tracking-tighter">TSPR Control Sheet</h2>
                            <p className="text-xs font-bold text-slate-400 uppercase tracking-widest">Production Data Entry & Trend Analysis</p>
                      </div>
                      <TsprSheet data={tsprData} onUpdate={setTsprData} />
                  </div>
               )}
            </div>
         )}
      </div>

      {showSettings && (
        <SettingsModal 
          currentUrl={cloudUrl} 
          onSave={(url) => { setCloudUrl(url); localStorage.setItem(CLOUD_URL_KEY, url); setShowSettings(false); }}
          onClose={() => setShowSettings(false)}
        />
      )}

      {selectedDefect && (
        <ManagerDefectModal 
            defect={selectedDefect} 
            onClose={() => setSelectedDefect(null)}
            onSave={(updated) => setDefects(defects.map(d => d.id === updated.id ? updated : d))}
        />
      )}
    </div>
  );
};

export default App;
