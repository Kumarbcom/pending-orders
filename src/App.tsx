import React, { useState, useMemo, useRef, useEffect, ChangeEvent } from 'react';
import { 
  Package, 
  Users, 
  LayoutDashboard, 
  ClipboardList, 
  FileText, 
  ChevronRight, 
  Search,
  Download,
  Filter,
  ArrowUpRight,
  TrendingUp,
  AlertCircle,
  CheckCircle2,
  PackageCheck,
  Upload,
  ChevronLeft,
  ChevronDown,
  ArrowUpDown,
  RefreshCw,
  XCircle,
  X,
  Menu,
  FileCheck
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { 
  PieChart, 
  Pie, 
  Cell, 
  ResponsiveContainer, 
  BarChart, 
  Bar, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  Legend,
  Bar as ReBar,
  LabelList
} from 'recharts';
import { motion, AnimatePresence } from 'motion/react';
import { ALL_SO as INITIAL_SO, ALL_INVOICE, META } from './data';
import { SalesOrder, Invoice, PurchaseOrder, StockItem, MaterialMasterItem, CustomerMasterItem } from './types';
import { cn } from './lib/utils';
import logo from './logo.jpg';
import { db } from './firebase';
import { doc, getDoc, setDoc } from 'firebase/firestore';

// --- Utils ---
const fmtCur = (v: number) => {
  if (v == null || isNaN(v)) return '—';
  const abs = Math.abs(v);
  if (abs >= 1e7) return '₹' + (v / 1e7).toFixed(2) + ' Cr';
  if (abs >= 1e5) return '₹' + (v / 1e5).toFixed(2) + ' L';
  return '₹' + v.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
};

const fmtNum = (v: number) => {
  if (v == null) return '—';
  return v.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
};

const fmtDate = (d: any) => {
  if (!d) return '—';
  let dt: Date;
  
  if (d instanceof Date) {
    dt = d;
  } else if (typeof d === 'number') {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    dt = new Date(excelEpoch.getTime() + d * 86400000);
  } else {
    const str = String(d);
    dt = new Date(str);
    if (isNaN(dt.getTime())) {
      const parts = str.split(/[\/\-\.]/);
      if (parts.length === 3) {
        // Handle DD.MM.YYYY or YYYY.MM.DD
        if (parts[2].length === 4) dt = new Date(`${parts[2]}-${parts[1]}-${parts[0]}`);
        else if (parts[0].length === 4) dt = new Date(`${parts[0]}-${parts[1]}-${parts[2]}`);
      }
    }
  }

  if (isNaN(dt.getTime())) return String(d);
  
  const day = dt.getDate().toString().padStart(2, '0');
  const month = dt.toLocaleString('en-GB', { month: 'short' });
  const year = dt.getFullYear().toString().slice(-2);
  return `${day}-${month}-${year}`;
};

const parseDateObj = (d: any): Date | null => {
  if (!d) return null;
  if (d instanceof Date) return d;
  try {
    if (typeof d === 'number') {
       const excelEpoch = new Date(Date.UTC(1899, 11, 30));
       return new Date(excelEpoch.getTime() + d * 86400000);
    }
    const str = String(d);
    const dt = new Date(str);
    if (!isNaN(dt.getTime())) return dt;
    
    const parts = str.split(/[\/\-\.]/);
    if (parts.length === 3) {
      if (parts[2].length === 4) return new Date(`${parts[2]}-${parts[1]}-${parts[0]}`);
      if (parts[0].length === 4) return new Date(`${parts[0]}-${parts[1]}-${parts[2]}`);
    }
  } catch(e) { return null; }
  return null;
};

// --- Error Boundary ---
class ErrorBoundary extends React.Component<{children: React.ReactNode}, {hasError: boolean}> {
  constructor(props: any) {
    super(props);
    this.state = { hasError: false };
  }
  static getDerivedStateFromError() { return { hasError: true }; }
  render() {
    if (this.state.hasError) {
      return (
        <div className="h-screen w-full flex flex-col items-center justify-center bg-bg text-text-main p-10 text-center">
          <h2 className="text-2xl font-black mb-4 uppercase tracking-tighter">Something went wrong</h2>
          <p className="text-text-muted mb-6">The application encountered an unexpected error. Please try refreshing or resetting the data.</p>
          <button onClick={() => window.location.reload()} className="bg-primary text-white px-6 py-3 rounded-xl font-bold uppercase tracking-widest text-xs">Refresh Application</button>
        </div>
      );
    }
    return this.props.children;
  }
}

// --- Components ---

const StatCard = ({ title, value, subValue, type, details }: any) => (
  <div className="bg-surface border border-border-custom rounded-2xl p-4 shadow-sm flex flex-col relative transition-all hover:shadow-md">
    <div className="flex justify-between items-start mb-2">
      <span className="text-[10px] font-bold uppercase tracking-wider text-text-muted">
        {title}
      </span>
      {type === 'due' && <AlertCircle className="w-4 h-4 text-due" />}
      {type === 'sched' && <TrendingUp className="w-4 h-4 text-sched" />}
    </div>
    <div className="text-[22px] leading-none font-bold text-text-main mb-1 tracking-tight">
      {value}
    </div>
    {subValue && (
      <div className="text-xs text-text-muted font-medium flex items-center gap-1.5 mt-1">
        {subValue}
      </div>
    )}
    
    {details && (
      <div className="grid grid-cols-2 gap-y-4 gap-x-2 mt-5 pt-5 border-t border-border-custom">
        {details.map((d: any, i: number) => (
          <div key={i}>
            <div className="text-[8px] font-black text-text-muted uppercase mb-1 tracking-widest leading-tight h-5 flex items-end">{d.label}</div>
            <div className={cn("font-black text-[11px] leading-tight break-all", d.color)}>{d.value}</div>
          </div>
        ))}
      </div>
    )}
  </div>
);

const Badge = ({ children, className }: any) => (
  <span className={cn("px-2 py-0.5 rounded-full text-[10px] font-bold uppercase tracking-wider whitespace-nowrap", className)}>
    {children}
  </span>
);

const SortIcon = ({ active, direction }: { active: boolean; direction: 'asc' | 'desc' }) => (
  <span className="inline-flex flex-col ml-1 opacity-60">
    <ArrowUpDown className={cn("w-2.5 h-2.5", active ? "text-primary" : "text-text-muted")} />
  </span>
);

const Th = ({ children, onSort, sortKey, activeField, direction, className }: any) => (
  <th 
    onClick={() => onSort?.(sortKey)}
    className={cn("cursor-pointer hover:bg-slate-100 transition-colors group relative", className)}
  >
    <div className="flex items-center justify-between gap-1">
      <span>{children}</span>
      {sortKey && <SortIcon active={activeField === sortKey} direction={direction} />}
    </div>
  </th>
);

const exportToExcel = (data: any[], fileName: string) => {
  try {
    if (!data || !data.length) {
      alert("No data available to export.");
      return;
    }
    const cleanName = (fileName || 'Report').replace(/[^a-z0-9]/gi, '_').slice(0, 50);
    const headers = Object.keys(data[0]);

    // Create styled HTML table
    let tableRows = `<tr>${headers.map(h => `<th style="background-color: #E0E0E0; border: 1px solid #000000; font-family: 'Cambria', serif; font-size: 10pt; text-align: left; padding: 5px;">${h}</th>`).join('')}</tr>`;
    
    data.forEach(row => {
      tableRows += '<tr>';
      headers.forEach(h => {
        let val = row[h];
        if (val === null || val === undefined) val = '';
        
        const isNum = /rate|value|price|amount|balance|total|ordered|qty/i.test(h);
        const isDate = /date|due/i.test(h);
        
        let style = "border: 1px solid #000000; font-family: 'Cambria', serif; font-size: 10pt; padding: 5px;";
        let msoFormat = "";
        
        if (isNum) {
          style += " text-align: right;";
          msoFormat = 'style="mso-number-format:\'\\#\\,\\#\\#0\\.00\'"';
        } else if (isDate) {
          msoFormat = 'style="mso-number-format:\'dd\\-mmm\\-yy\'"';
        } else {
          msoFormat = 'style="mso-number-format:\'@\'"'; // Format as text
        }
        
        tableRows += `<td ${msoFormat} style="${style}">${val}</td>`;
      });
      tableRows += '</tr>';
    });

    const html = `
      <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
      <head><meta charset="utf-8"></head>
      <body>
        <table border="1">${tableRows}</table>
      </body>
      </html>
    `;

    const blob = new Blob([html], { type: 'application/vnd.ms-excel' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${cleanName}.xls`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
  } catch (err) {
    console.error("Export Error:", err);
    alert("Export failed: " + err);
  }
};

function MainApp() {
  const [activeTab, setActiveTab] = useState<'dashboard' | 'pending-so' | 'pending-po' | 'stock' | 'material-master' | 'customer-master' | 'invoices' | 'customers'>('dashboard');
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  
  // Data States
  const [dynamicSO, setDynamicSO] = useState<SalesOrder[]>([]);
  const [dynamicPO, setDynamicPO] = useState<PurchaseOrder[]>([]);
  const [dynamicStock, setDynamicStock] = useState<StockItem[]>([]);
  const [dynamicMaterialMaster, setDynamicMaterialMaster] = useState<MaterialMasterItem[]>([]);
  const [dynamicCustomerMaster, setDynamicCustomerMaster] = useState<CustomerMasterItem[]>([]);
  const [dynamicInvoices, setDynamicInvoices] = useState<Invoice[]>([]);
  const [showUploadMenu, setShowUploadMenu] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [soSearch, setSoSearch] = useState('');
  const [poSearch, setPoSearch] = useState('');
  const [stockSearch, setStockSearch] = useState('');
  const [materialSearch, setMaterialSearch] = useState('');
  const [customerMasterSearch, setCustomerMasterSearch] = useState('');
  const [invoiceSearch, setInvoiceSearch] = useState('');
  const [popupSearch, setPopupSearch] = useState('');
  const [isSyncing, setIsSyncing] = useState(false);
  const uploadMenuRef = useRef<HTMLDivElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const fileInputPORef = useRef<HTMLInputElement>(null);
  const fileInputStockRef = useRef<HTMLInputElement>(null);
  const fileInputMaterialRef = useRef<HTMLInputElement>(null);
  const fileInputCustomerRef = useRef<HTMLInputElement>(null);
  const fileInputInvoiceRef = useRef<HTMLInputElement>(null);

  // Sorting State
  const [sortField, setSortField] = useState<string | null>(null);
  const [sortDirection, setSortDirection] = useState<'asc' | 'desc'>('asc');

  // Dashboard Slicers
  const [dMake, setDMake] = useState('');
  const [dGroup, setDGroup] = useState('');
  const [dCGroup, setDCGroup] = useState('');
  const [dOrderType, setDOrderType] = useState('');

  // Pending SO Slicers
  const [soType, setSoType] = useState('');
  const [soMake, setSoMake] = useState('');
  const [soGroup, setSoGroup] = useState('');
  const [soCust, setSoCust] = useState('');
  const [soStatus, setSoStatus] = useState('');

  // Customers Slicers
  const [cGroup, setCGroup] = useState('');
  const [cCGroup, setCCGroup] = useState('');
  const [cCust, setCCust] = useState('');
  const [cSearch, setCSearch] = useState('');

  // DETAILS POPUP STATE
  const [showSOPopup, setShowSOPopup] = useState<string | null>(null);
  const [showInvPopup, setShowInvPopup] = useState<string | null>(null);

   // --- Firebase Persistence Logic ---
   const saveToFirebase = async (type: string, data: any) => {
     setIsSyncing(true);
     try {
       await setDoc(doc(db, "app_data", type), { data, updatedAt: new Date().toISOString() });
     } catch (error) {
       console.error(`Error saving ${type}:`, error);
     } finally {
       setIsSyncing(false);
     }
   };

   useEffect(() => {
     const loadData = async () => {
       setIsSyncing(true);
       try {
         const datasets = [
           { type: 'dynamicSO', setter: setDynamicSO },
           { type: 'dynamicPO', setter: setDynamicPO },
           { type: 'dynamicStock', setter: setDynamicStock },
           { type: 'dynamicMaterialMaster', setter: setDynamicMaterialMaster },
           { type: 'dynamicCustomerMaster', setter: setDynamicCustomerMaster },
           { type: 'dynamicInvoices', setter: setDynamicInvoices }
         ];

         for (const ds of datasets) {
           const snap = await getDoc(doc(db, "app_data", ds.type));
           if (snap.exists()) {
             ds.setter(snap.data().data || []);
           }
         }
       } catch (error) {
         console.error("Error loading data from Firebase:", error);
       } finally {
         setIsSyncing(false);
       }
     };
     loadData();
   }, []);

  const handleSort = (field: string) => {
    if (sortField === field) {
      setSortDirection(sortDirection === 'asc' ? 'desc' : 'asc');
    } else {
      setSortField(field);
      setSortDirection('asc');
    }
  };

  const handleReset = (type: 'so' | 'po' | 'stock' | 'material' | 'customer' | 'invoice' | 'all') => {
    setSearchTerm('');
    setPopupSearch('');
    if (type === 'so' || type === 'all') { 
      setSoType(''); setSoMake(''); setSoGroup(''); setSoCust(''); setSoStatus(''); 
      setDMake(''); setDGroup(''); setDCGroup(''); setDOrderType('');
    }
    if (type === 'po' || type === 'all') { setPoSearch(''); }
    if (type === 'stock' || type === 'all') { setStockSearch(''); }
    if (type === 'material' || type === 'all') { setMaterialSearch(''); }
    if (type === 'customer' || type === 'all') { setCustomerMasterSearch(''); setCSearch(''); setCGroup(''); setCCGroup(''); }
    if (type === 'invoice' || type === 'all') { setInvoiceSearch(''); }
  };

  const handleWipeData = async () => {
    if (!confirm("ARE YOU SURE? This will permanently delete ALL uploaded data from the cloud database.")) return;
    setIsSyncing(true);
    try {
      const types = ['dynamicSO', 'dynamicPO', 'dynamicStock', 'dynamicMaterialMaster', 'dynamicCustomerMaster', 'dynamicInvoices'];
      for (const t of types) {
        await saveToFirebase(t, []);
      }
      setDynamicSO([]);
      setDynamicPO([]);
      setDynamicStock([]);
      setDynamicMaterialMaster([]);
      setDynamicCustomerMaster([]);
      setDynamicInvoices([]);
      handleReset('all');
      alert("Database wiped successfully.");
    } catch (err) {
      alert("Error wiping database: " + err);
    } finally {
      setIsSyncing(false);
    }
  };

  // Close upload menu on outside click
  React.useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (uploadMenuRef.current && !uploadMenuRef.current.contains(event.target as Node)) {
        setShowUploadMenu(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const handleFileUpload = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const rawData = (XLSX.utils.sheet_to_json(ws, { header: 1 }) || []) as any[][];
      if (!rawData || rawData.length === 0) return;
      
      let headerIdx = 0;
      for (let i = 0; i < Math.min(rawData.length, 20); i++) {
        const row = rawData[i];
        if (row && row.some(cell => typeof cell === 'string' && (cell.includes("Party") || cell.includes("Name of Item") || cell.includes("Due") || cell.includes("Customer")))) {
          headerIdx = i;
          break;
        }
      }

      const headers = rawData[headerIdx] || [];
      const rows = rawData.slice(headerIdx + 1);

      const parsed: SalesOrder[] = rows.map((rowArr: any[]) => {
        const row: any = {};
        headers.forEach((h, idx) => {
          if (h && typeof h === 'string') {
             const cleanHeader = h.replace(/\r?\n|\r/g, ' ').trim();
             row[cleanHeader] = rowArr[idx];
          }
        });

        const extractNum = (val: any) => {
          if (typeof val === 'number') return val;
          if (!val) return 0;
          const match = String(val).match(/[\d.]+/);
          return match ? parseFloat(match[0]) : 0;
        };

        const party = String(row["Party's Name"] || row['Party Name'] || row['End-Customer'] || row['Customer'] || '').trim();
        const itemName = String(row['Name of Item'] || row['Item Name'] || row['Description'] || '').trim();
        const orderQty = extractNum(row['Ordered'] || row['Order'] || row['Ordered Qty'] || row['Qty']);
        const balQty = extractNum(row['Balance'] || row['Balar'] || row['Balance Qty']);
        const val = extractNum(row['Value'] || row['Amount']);

        if (!party && !itemName) return null; // Skip empty rows

        return {
          Date: row['Date'] || row['Due'] || '',
          Order: row['Order'] || row['Ref No'] || row['Order No'] || row['Voucher No'] || row['Part No'] || '',
          PartyName: party,
          NameOfItem: itemName,
          MaterialCode: row['Material Code'] || '',
          PartNo: row['Part No'] || '',
          Ordered: orderQty,
          Balance: balQty,
          Rate: extractNum(row['Rate'] || row['R'] || row['Price']),
          Discount: extractNum(row['Discount'] || row['Discou']),
          Value: val,
          DueOn: row['Due on'] || row['Due'] || null,
          DueSerial: null,
          Make: '',
          MaterialGroup: '',
          Group: '',
          CustomerGroup: '',
          OrderType: 'Due',
          StockAllocated: 0,
          StockShortfall: 0,
          StockStatus: 'Need to Place Order',
          POStatus: '',
          ExpDelivery: ''
        };
      }).filter(Boolean) as SalesOrder[];

      if (parsed.length > 0) {
        setDynamicSO(parsed);
        saveToFirebase('dynamicSO', parsed);
        alert(`Successfully uploaded ${parsed.length} Sales Orders.`);
      } else {
        alert("No Sales Order data could be parsed from this file. Check headers.");
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = '';
  };

  const handlePOUpload = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      
      const rawData = (XLSX.utils.sheet_to_json(ws, { header: 1 }) || []) as any[][];
      if (!rawData || rawData.length === 0) return;
      
      let headerIdx = 0;
      for (let i = 0; i < Math.min(rawData.length, 20); i++) {
        const row = rawData[i];
        if (row && row.some(cell => typeof cell === 'string' && (cell.includes("Party") || cell.includes("Name of Item") || cell.includes("Due")))) {
          headerIdx = i;
          break;
        }
      }

      const headers = rawData[headerIdx] || [];
      const rows = rawData.slice(headerIdx + 1);

      const parsed: PurchaseOrder[] = rows.map((rowArr: any[]) => {
        const row: any = {};
        headers.forEach((h, idx) => {
          if (h && typeof h === 'string') {
             // Remove any newlines or weird spaces from headers
             const cleanHeader = h.replace(/\r?\n|\r/g, ' ').trim();
             row[cleanHeader] = rowArr[idx];
          }
        });

        const extractNum = (val: any) => {
          if (typeof val === 'number') return val;
          if (!val) return 0;
          const match = String(val).match(/[\d.]+/);
          return match ? parseFloat(match[0]) : 0;
        };

        const party = String(row["Party's Name"] || row['Party Name'] || row['Supplier'] || '').trim();
        const itemName = String(row['Name of Item'] || row['Item Name'] || row['Description'] || '').trim();
        const orderQty = extractNum(row['Ordered'] || row['Ordered Qty'] || row['Qty']);
        const balQty = extractNum(row['Balance'] || row['Balance Qty']);
        const val = extractNum(row['Value'] || row['Amount']);

        if (!party && !itemName) return null; // Skip empty rows

        return {
          Date: row['Date'] || '',
          Order: row['Order'] || row['Ref No'] || row['Order No'] || row['Voucher No'] || '',
          PartyName: party,
          NameOfItem: itemName,
          MaterialCode: row['Material Code'] || '',
          PartNo: row['Part No'] || '',
          Ordered: orderQty,
          Balance: balQty,
          Rate: extractNum(row['Rate'] || row['Price']),
          Discount: extractNum(row['Discount']),
          Value: val,
          DueOn: row['Due on'] || row['Due'] || null
        };
      }).filter(Boolean) as PurchaseOrder[];

      if (parsed.length > 0) {
        setDynamicPO(parsed);
        saveToFirebase('dynamicPO', parsed);
        alert(`Successfully uploaded ${parsed.length} PO records.`);
      } else {
        alert("No Purchase Order data could be parsed from this file. Check headers.");
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = '';
  };

  const handleStockUpload = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws);

      const parsed: StockItem[] = data.map((row: any) => {
        const extractNum = (val: any) => {
          if (typeof val === 'number') return val;
          if (!val) return 0;
          const match = String(val).match(/[\d.]+/);
          return match ? parseFloat(match[0]) : 0;
        };

        return {
          Particulars: row['Particulars'] || '',
          Quantity: extractNum(row['Quantity']),
          Rate: extractNum(row['Rate']),
          Value: extractNum(row['Value'])
        };
      });

      if (parsed.length > 0) {
        setDynamicStock(parsed);
        saveToFirebase('dynamicStock', parsed);
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleMaterialUpload = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws);

      const parsed: MaterialMasterItem[] = data.map((row: any) => {
        return {
          Description: row['Description'] || '',
          PartNo: row['Part No'] || '',
          Make: row['Make'] || '',
          MaterialGroup: row['Material Group'] || row['Material Gro'] || ''
        };
      });

      if (parsed.length > 0) {
        setDynamicMaterialMaster(parsed);
        saveToFirebase('dynamicMaterialMaster', parsed);
        alert(`Successfully uploaded ${parsed.length} Materials.`);
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleCustomerUpload = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws);

      const parsed: CustomerMasterItem[] = data.map((row: any) => {
        return {
          CustomerName: row['Customer Name'] || '',
          Group: row['Group'] || '',
          SalesRep: row['Sales Rep'] || '',
          Status: row['Status'] || '',
          CustomerGroup: row['Customer Group'] || ''
        };
      });

      if (parsed.length > 0) {
        setDynamicCustomerMaster(parsed);
        saveToFirebase('dynamicCustomerMaster', parsed);
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleInvoiceUpload = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      
      const rawData = (XLSX.utils.sheet_to_json(ws, { header: 1 }) || []) as any[][];
      if (!rawData || rawData.length === 0) return;
      
      let headerIdx = 0;
      for (let i = 0; i < Math.min(rawData.length, 20); i++) {
        const row = rawData[i];
        if (row && row.some(cell => typeof cell === 'string' && (cell.includes("Buyer") || cell.includes("Voucher No") || cell.includes("Particulars") || cell.includes("Description")))) {
          headerIdx = i;
          break;
        }
      }

      const headers = rawData[headerIdx] || [];
      const rows = rawData.slice(headerIdx + 1);

      const invoices: Invoice[] = [];
      
      // Tracking variables for forward-fill
      let lastDate: any = '';
      let lastBuyer: any = '';
      let lastConsignee: any = '';
      let lastVType = '';
      let lastVNo = '';
      let lastVRef = '';
      let lastGSTIN = '';

      const extractNum = (val: any) => {
        if (typeof val === 'number') return val;
        if (!val) return 0;
        const match = String(val).match(/[\d.]+/);
        return match ? parseFloat(match[0]) : 0;
      };

      rows.forEach((rowArr: any[]) => {
        const row: any = {};
        headers.forEach((h, idx) => {
          if (h && typeof h === 'string') {
             const cleanHeader = h.replace(/\r?\n|\r/g, ' ').trim();
             row[cleanHeader] = rowArr[idx];
          }
        });

        const date = row['Date'];
        const buyer = row['Buyer'];
        const particulars = row['Particulars'] || row['Description'] || '';
        
        if (!particulars && !buyer && !date) return; // Skip empty rows

        // Forward fill logic
        if (date) lastDate = date;
        if (buyer) lastBuyer = String(buyer).trim();
        if (row['Consignee']) lastConsignee = String(row['Consignee']).trim();
        else if (buyer) lastConsignee = String(buyer).trim();

        if (row['Voucher Type']) lastVType = String(row['Voucher Type']).trim();
        if (row['Voucher No.'] || row['Voucher No']) lastVNo = String(row['Voucher No.'] || row['Voucher No']).trim();
        if (row['Voucher Ref. No.'] || row['Voucher Ref No']) lastVRef = String(row['Voucher Ref. No.'] || row['Voucher Ref No']).trim();
        if (row['GSTIN/UIN'] || row['GSTIN']) lastGSTIN = String(row['GSTIN/UIN'] || row['GSTIN']).trim();

        if (particulars) {
          const qty = extractNum(row['Quantity'] || row['Qty']);
          const val = extractNum(row['Value'] || row['Amount']);
          if (qty <= 0 && val <= 0) return; 

          let pText = String(particulars).trim();
          if (!pText || pText.toLowerCase() === 'particulars') return;

          // Skip if particulars look like common non-material accounts
          const upperP = pText.toUpperCase();
          if (upperP.includes('CGST') || upperP.includes('SGST') || upperP.includes('IGST') || upperP.includes('ROUNDING') || upperP.includes('DISCOUNT')) {
             return;
          }

          // Remove customer name if it exists in particulars
          if (lastBuyer) {
            const escapedBuyer = lastBuyer.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
            const buyerRegex = new RegExp(`^${escapedBuyer}|${escapedBuyer}$`, 'gi');
            pText = pText.replace(buyerRegex, '').replace(/^[\s\W]+|[\s\W]+$/g, '').trim();
          }

          if (!pText || pText.length < 3 || pText.toLowerCase() === 'particulars') return; 
          
          invoices.push({
            Date: lastDate,
            Particulars: pText,
            Buyer: lastBuyer,
            Consignee: lastConsignee,
            VoucherType: lastVType,
            VoucherNo: lastVNo,
            VoucherRef: lastVRef,
            GSTIN: lastGSTIN,
            Quantity: qty,
            Value: val
          });
        }
      });

      if (invoices.length > 0) {
        setDynamicInvoices(invoices);
        saveToFirebase('dynamicInvoices', invoices);
        alert(`Successfully uploaded ${invoices.length} line items.`);
      } else {
        alert("No Invoice data could be parsed from this file. Check headers.");
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = '';
  };
  
   // Stock Slicers
   // Material Master Slicers
   // Customer Master Slicers
   // Invoice Master Slicers

  const selectedCustomerData = useMemo(() => {
    if (!showSOPopup && !showInvPopup) return null;
    const name = showSOPopup || showInvPopup;
    return dynamicCustomerMaster.find(c => c.CustomerName === name) || null;
  }, [showSOPopup, showInvPopup, dynamicCustomerMaster]);

  // FIFO PROCESSING LOGIC
  const processedSO = useMemo(() => {
    const now = new Date();
    const CUTOFF_DATE = new Date(now.getFullYear(), now.getMonth() + 1, 0);
    
    // 1. Prepare Stock Map for FIFO
    const stockMap: Record<string, number> = {};
    dynamicStock.forEach(s => {
      stockMap[s.Particulars] = (stockMap[s.Particulars] || 0) + s.Quantity;
    });

    // 2. Prepare PO Map
    const poMap: Record<string, number> = {};
    dynamicPO.forEach(p => {
      poMap[p.NameOfItem] = (poMap[p.NameOfItem] || 0) + p.Balance;
    });

    // 3. Enrich and Classify
    const sortedRaw = [...dynamicSO].sort((a, b) => {
      const dtA = parseDateObj(a.DueOn);
      const dtB = parseDateObj(b.DueOn);
      const dbA = dtA ? dtA.getTime() : 0;
      const dbB = dtB ? dtB.getTime() : 0;
      return dbA - dbB;
    });

    return sortedRaw.map(so => {
      const dueDate = parseDateObj(so.DueOn);
      const orderType = (dueDate && dueDate <= CUTOFF_DATE) ? 'Due' : 'Schedule';
      
      // Stock Allocation
      const available = stockMap[so.NameOfItem] || 0;
      const allocated = Math.min(available, so.Balance);
      stockMap[so.NameOfItem] = Math.max(0, available - allocated);
      
      const shortfall = so.Balance - allocated;
      
      // PO Lookup
      let status: "Available" | "PO Exist - Expedite" | "Need to Place Order" = "Available";
      if (shortfall > 0) {
        const poAvailable = poMap[so.NameOfItem] || 0;
        if (poAvailable > 0) {
          status = "PO Exist - Expedite";
          poMap[so.NameOfItem] = Math.max(0, poAvailable - shortfall);
        } else {
          status = "Need to Place Order";
        }
      }

      const material = dynamicMaterialMaster.find(m => m.Description === so.NameOfItem);
      const customer = dynamicCustomerMaster.find(c => c.CustomerName === so.PartyName);

      return {
        ...so,
        OrderType: orderType,
        StockAllocated: allocated,
        StockShortfall: shortfall,
        StockStatus: status,
        Make: material?.Make || so.Make,
        MaterialGroup: material?.MaterialGroup || so.MaterialGroup,
        Group: customer?.Group || so.Group,
        CustomerGroup: customer?.CustomerGroup || so.CustomerGroup
      } as SalesOrder;
    });
  }, [dynamicSO, dynamicStock, dynamicPO, dynamicMaterialMaster, dynamicCustomerMaster]);

  const [selectedCustomer, setSelectedCustomer] = useState<string | null>(null);
  const [selectedInvoiceCust, setSelectedInvoiceCust] = useState<string | null>(null);

  // --- Data Logic ---
  
  const filteredDashboardSO = useMemo(() => {
    return processedSO.filter(r => {
      if (dMake && r.Make !== dMake) return false;
      if (dGroup && r.Group !== dGroup) return false;
      if (dCGroup && r.CustomerGroup !== dCGroup) return false;
      if (dOrderType && r.OrderType !== dOrderType) return false;
      
      if (searchTerm) {
        const s = searchTerm.toLowerCase();
        return (
          r.PartyName.toLowerCase().includes(s) ||
          r.NameOfItem.toLowerCase().includes(s) ||
          r.Order.toLowerCase().includes(s)
        );
      }
      return true;
    });
  }, [processedSO, dMake, dGroup, dCGroup, dOrderType, searchTerm]);

  const dashboardStats = useMemo(() => {
    const total = filteredDashboardSO.reduce((s, r) => s + r.Value, 0);
    const uniqueOrders = new Set(filteredDashboardSO.map(r => r.Order)).size;
    const uniqueCustomers = new Set(filteredDashboardSO.map(r => r.PartyName)).size;

    const due = filteredDashboardSO.filter(r => r.OrderType === 'Due');
    const sched = filteredDashboardSO.filter(r => r.OrderType === 'Schedule');
    
    const dueVal = due.reduce((s, r) => s + r.Value, 0);
    const dueAvail = due.filter(r => r.StockStatus === 'Available').reduce((s, r) => s + r.Value, 0);
    const dueArr = dueVal - dueAvail;

    const schedVal = sched.reduce((s, r) => s + r.Value, 0);
    const schedAvail = sched.filter(r => r.StockStatus === 'Available').reduce((s, r) => s + r.Value, 0);
    const schedArr = schedVal - schedAvail;

    const totalPO = dynamicPO.reduce((s, p) => s + p.Value, 0);
    const poCount = dynamicPO.length;
    const uniquePO = new Set(dynamicPO.map(p => p.Order)).size;
    const uniqueSuppliers = new Set(dynamicPO.map(p => p.PartyName)).size;

    return { 
      total, 
      count: filteredDashboardSO.length, 
      uniqueOrders, 
      uniqueCustomers, 
      dueVal, 
      dueAvail, 
      dueArr, 
      schedVal, 
      schedAvail, 
      schedArr, 
      totalPO, 
      poCount,
      uniquePO,
      uniqueSuppliers
    };
  }, [filteredDashboardSO, dynamicPO]);

  const filteredSO = useMemo(() => {
    let list = processedSO.filter(r => {
      if (soType && r.OrderType !== soType) return false;
      if (soMake && r.Make !== soMake) return false;
      if (soGroup && r.Group !== soGroup) return false;
      if (soCust) {
        const s = soCust.toLowerCase();
        if (
          !(r.PartyName?.toLowerCase() || "").includes(s) && 
          !(r.NameOfItem?.toLowerCase() || "").includes(s) && 
          !(r.Order?.toLowerCase() || "").includes(s)
        ) return false;
      }
      if (soStatus && r.StockStatus !== soStatus) return false;

      if (searchTerm) {
        const s = searchTerm.toLowerCase();
        return (
          (r.PartyName?.toLowerCase() || "").includes(s) ||
          (r.NameOfItem?.toLowerCase() || "").includes(s) ||
          (r.Order?.toLowerCase() || "").includes(s)
        );
      }
      return true;
    });
    if (sortField) {
      list = [...list].sort((a: any, b: any) => {
        const fieldA = a[sortField];
        const fieldB = b[sortField];
        if (typeof fieldA === 'number' && typeof fieldB === 'number') return sortDirection === 'asc' ? fieldA - fieldB : fieldB - fieldA;
        if (sortField.toLowerCase().includes('date')) {
          const dA = fieldA ? new Date(fieldA).getTime() : 0;
          const dB = fieldB ? new Date(fieldB).getTime() : 0;
          return sortDirection === 'asc' ? dA - dB : dB - dA;
        }
        return sortDirection === 'asc' ? String(fieldA).localeCompare(String(fieldB)) : String(fieldB).localeCompare(String(fieldA));
      });
    }
    return list;
  }, [processedSO, soType, soMake, soGroup, soCust, soStatus, sortField, sortDirection]);

  const customFilterSOTab = useMemo(() => {
    const total = filteredSO.reduce((s, r) => s + r.Value, 0);
    const due = filteredSO.filter(r => r.OrderType === 'Due');
    const sched = filteredSO.filter(r => r.OrderType === 'Schedule');
    const dueVal = due.reduce((s, r) => s + r.Value, 0);
    const dueAvail = due.filter(r => r.StockStatus === 'Available').reduce((s, r) => s + r.Value, 0);
    const schedVal = sched.reduce((s, r) => s + r.Value, 0);
    const schedAvail = sched.filter(r => r.StockStatus === 'Available').reduce((s, r) => s + r.Value, 0);

    return { total, count: filteredSO.length, dueVal, dueAvail, schedVal, schedAvail };
  }, [filteredSO]);

  const customersList = useMemo(() => {
    const custMap: Record<string, any> = {};
    const s = cSearch.toLowerCase();

    // 1. Process Sales Orders
    processedSO.forEach(r => {
      if (!r?.PartyName) return;
      if (cGroup && r.Group !== cGroup) return;
      if (cCGroup && r.CustomerGroup !== cCGroup) return;

      const matchesSearch = !cSearch || 
        r.PartyName.toLowerCase().includes(s) || 
        (r.Order || "").toLowerCase().includes(s) ||
        (r.NameOfItem || "").toLowerCase().includes(s);

      if (!matchesSearch) return;

      if (!custMap[r.PartyName]) {
        custMap[r.PartyName] = { name: r.PartyName, total: 0, dueVal: 0, schedVal: 0, group: r.Group || 'N/A', cgroup: r.CustomerGroup || 'N/A', invCount: 0, invVal: 0 };
      }
      const c = custMap[r.PartyName];
      c.total += r.Value || 0;
      if (r.OrderType === 'Due') c.dueVal += r.Value || 0;
      else if (r.OrderType === 'Schedule') c.schedVal += r.Value || 0;
    });

    // 2. Process Invoices
    dynamicInvoices.forEach(inv => {
      if (!inv?.Buyer) return;
      const name = inv.Buyer;
      
      const matchesSearch = !cSearch || 
        name.toLowerCase().includes(s) || 
        (inv.VoucherNo || "").toLowerCase().includes(s) || 
        (inv.VoucherRef || "").toLowerCase().includes(s) ||
        (inv.Particulars || "").toLowerCase().includes(s);

      if (!matchesSearch) return;

      if (!custMap[name]) {
        custMap[name] = { name, total: 0, dueVal: 0, schedVal: 0, group: 'N/A', cgroup: 'N/A', invCount: 0, invVal: 0 };
      }
      const c = custMap[name];
      c.invCount++;
      c.invVal += inv.Value || 0;
    });

    return Object.values(custMap).sort((a: any, b: any) => (b.total + b.invVal) - (a.total + a.invVal));
  }, [processedSO, dynamicInvoices, cGroup, cCGroup, cSearch]);

  const dashboardChartsData = useMemo(() => {
    // Total Pending Value Difference Chart (Horizontal Bar)
    const pendingTypeData = [
      { name: 'Due Orders', value: dashboardStats.dueVal, fill: '#ff4d4f' },
      { name: 'Schedule Orders', value: dashboardStats.schedVal, fill: '#1890ff' }
    ];

    // Make Distribution Breakdown (Due vs Schedule)
    const makeMap: Record<string, { name: string; due: number; schedule: number }> = {};
    filteredDashboardSO.forEach(r => {
      const m = r.Make || 'Other';
      if (!makeMap[m]) makeMap[m] = { name: m, due: 0, schedule: 0 };
      if (r.OrderType === 'Due') makeMap[m].due += r.Value;
      else makeMap[m].schedule += r.Value;
    });
    const makeStackedData = Object.values(makeMap).sort((a, b) => (b.due + b.schedule) - (a.due + a.schedule)).slice(0, 8);

    // Top 10 Customers Detailed Breakdown
    const custMap: Record<string, any> = {};
    filteredDashboardSO.forEach(r => {
      const k = r.PartyName;
      if (!custMap[k]) custMap[k] = { 
        name: k, 
        group: r.Group, 
        cgroup: r.CustomerGroup, 
        total: 0, 
        dueAvail: 0, 
        dueArr: 0, 
        sched: 0 
      };
      
      custMap[k].total += r.Value;
      if (r.OrderType === 'Due') {
        if (r.StockStatus === 'Available') custMap[k].dueAvail += r.Value;
        else custMap[k].dueArr += r.Value;
      } else {
        custMap[k].sched += r.Value;
      }
    });

    const top10Detailed = Object.values(custMap)
      .sort((a, b) => {
        if (sortField && a[sortField] !== undefined) {
          const fieldA = a[sortField];
          const fieldB = b[sortField];
          if (typeof fieldA === 'number' && typeof fieldB === 'number') return sortDirection === 'asc' ? fieldA - fieldB : fieldB - fieldA;
          return sortDirection === 'asc' ? String(fieldA).localeCompare(String(fieldB)) : String(fieldB).localeCompare(String(fieldA));
        }
        return b.total - a.total;
      })
      .slice(0, 10);

    return { pendingTypeData, makeStackedData, top10Detailed };
  }, [filteredDashboardSO, dashboardStats, sortField, sortDirection]);

  const filteredPOList = useMemo(() => {
    let list = dynamicPO.filter(r => {
      const s = poSearch.toLowerCase();
      if (poSearch && 
          !(r.PartyName?.toLowerCase() || "").includes(s) && 
          !(r.NameOfItem?.toLowerCase() || "").includes(s) && 
          !(r.Order?.toLowerCase() || "").includes(s)
      ) return false;
      
      if (searchTerm) {
        const s = searchTerm.toLowerCase();
        return (
          (r.PartyName?.toLowerCase() || "").includes(s) ||
          (r.NameOfItem?.toLowerCase() || "").includes(s) ||
          (r.Order?.toLowerCase() || "").includes(s)
        );
      }
      return true;
    });
    if (sortField) {
      list = [...list].sort((a: any, b: any) => {
        const fieldA = a[sortField];
        const fieldB = b[sortField];
        if (typeof fieldA === 'number' && typeof fieldB === 'number') return sortDirection === 'asc' ? fieldA - fieldB : fieldB - fieldA;
        return sortDirection === 'asc' ? String(fieldA).localeCompare(String(fieldB)) : String(fieldB).localeCompare(String(fieldA));
      });
    }
    return list;
  }, [dynamicPO, poSearch, searchTerm, sortField, sortDirection]);

  const poStats = useMemo(() => {
    const total = filteredPOList.reduce((s, r) => s + r.Value, 0);
    const uniqueOrders = new Set(filteredPOList.map(r => r.Order)).size;
    const uniqueSuppliers = new Set(filteredPOList.map(r => r.PartyName)).size;
    return { total, count: filteredPOList.length, uniqueOrders, uniqueSuppliers };
  }, [filteredPOList]);

  const filteredStockList = useMemo(() => {
    let list = dynamicStock.filter(s => {
      if (stockSearch && !(s.Particulars?.toLowerCase() || "").includes(stockSearch.toLowerCase())) return false;
      
      if (searchTerm) {
        const srch = searchTerm.toLowerCase();
        return s.Particulars.toLowerCase().includes(srch);
      }
      return true;
    });
    if (sortField) {
      list = [...list].sort((a: any, b: any) => {
        const fieldA = a[sortField];
        const fieldB = b[sortField];
        if (typeof fieldA === 'number' && typeof fieldB === 'number') return sortDirection === 'asc' ? fieldA - fieldB : fieldB - fieldA;
        return sortDirection === 'asc' ? String(fieldA).localeCompare(String(fieldB)) : String(fieldB).localeCompare(String(fieldA));
      });
    }
    return list;
  }, [dynamicStock, stockSearch, searchTerm, sortField, sortDirection]);

  const filteredMaterialList = useMemo(() => {
    let list = dynamicMaterialMaster.filter(m => {
      if (materialSearch && 
          !m.Description.toLowerCase().includes(materialSearch.toLowerCase()) && 
          !m.PartNo.toLowerCase().includes(materialSearch.toLowerCase())) return false;
      return true;
    });
    if (sortField) {
      list = [...list].sort((a: any, b: any) => {
        const fieldA = a[sortField];
        const fieldB = b[sortField];
        if (typeof fieldA === 'number' && typeof fieldB === 'number') return sortDirection === 'asc' ? fieldA - fieldB : fieldB - fieldA;
        return sortDirection === 'asc' ? String(fieldA).localeCompare(String(fieldB)) : String(fieldB).localeCompare(String(fieldA));
      });
    }
    return list;
  }, [dynamicMaterialMaster, materialSearch, sortField, sortDirection]);

  const filteredCustomerMasterList = useMemo(() => {
    let list = dynamicCustomerMaster.filter(c => {
      if (customerMasterSearch && !(c.CustomerName?.toLowerCase() || "").includes(customerMasterSearch.toLowerCase())) return false;
      return true;
    });
    if (sortField) {
      list = [...list].sort((a: any, b: any) => {
        const fieldA = a[sortField];
        const fieldB = b[sortField];
        if (typeof fieldA === 'number' && typeof fieldB === 'number') return sortDirection === 'asc' ? fieldA - fieldB : fieldB - fieldA;
        return sortDirection === 'asc' ? String(fieldA).localeCompare(String(fieldB)) : String(fieldB).localeCompare(String(fieldA));
      });
    }
    return list;
  }, [dynamicCustomerMaster, customerMasterSearch, sortField, sortDirection]);

  const filteredInvoiceList = useMemo(() => {
    let list = dynamicInvoices.filter(i => {
      const s = invoiceSearch.toLowerCase();
      if (invoiceSearch && 
          !(i.Buyer?.toLowerCase() || "").includes(s) && 
          !(i.VoucherNo?.toLowerCase() || "").includes(s) && 
          !(i.VoucherRef?.toLowerCase() || "").includes(s) &&
          !(i.Particulars?.toLowerCase() || "").includes(s)
      ) return false;
      return true;
    });
    if (sortField) {
      list = [...list].sort((a: any, b: any) => {
        const fieldA = a[sortField];
        const fieldB = b[sortField];
        if (typeof fieldA === 'number' && typeof fieldB === 'number') return sortDirection === 'asc' ? fieldA - fieldB : fieldB - fieldA;
        if (sortField.toLowerCase().includes('date')) {
          const dA = fieldA ? new Date(fieldA).getTime() : 0;
          const dB = fieldB ? new Date(fieldB).getTime() : 0;
          return sortDirection === 'asc' ? dA - dB : dB - dA;
        }
        return sortDirection === 'asc' ? String(fieldA).localeCompare(String(fieldB)) : String(fieldB).localeCompare(String(fieldA));
      });
    }
    return list;
  }, [dynamicInvoices, invoiceSearch, sortField, sortDirection]);

  return (
    <div className="flex h-screen w-full bg-bg text-text-main overflow-hidden">
      {/* SIDEBAR */}
      <aside className={cn(
        "bg-surface border-r border-border-custom flex flex-col p-4 shrink-0 relative z-20 transition-all duration-300",
        sidebarCollapsed ? "w-20" : "w-[260px]"
      )}>
        <div className="flex items-center justify-between mb-8 px-2">
          {!sidebarCollapsed && (
            <div className="flex items-center gap-3">
              <div className="w-12 h-12 rounded-xl bg-white flex items-center justify-center shadow-sm border border-border-custom overflow-hidden p-1">
                <img src={logo} alt="Logo" className="w-full h-full object-contain" />
              </div>
              <h1 className="text-[12px] font-black tracking-tight text-text-main leading-tight">
                SIDDHI KABEL CORPORATION<br/>
                <span className="text-[10px] text-primary opacity-80 uppercase tracking-widest">Pending PO Review</span>
              </h1>
            </div>
          )}
          <button 
            onClick={() => setSidebarCollapsed(!sidebarCollapsed)}
            className="w-8 h-8 rounded-full hover:bg-surface2 flex items-center justify-center border border-border-custom mx-auto"
          >
            {sidebarCollapsed ? <Menu className="w-4 h-4" /> : <ChevronLeft className="w-4 h-4" />}
          </button>
        </div>

        <nav className="flex flex-col gap-1 flex-1 overflow-y-auto scrollbar-none">
          {[
            { id: 'dashboard', label: 'Overview', icon: LayoutDashboard },
            { id: 'pending-so', label: 'Pending SO', icon: ClipboardList },
            { id: 'pending-po', label: 'Pending PO', icon: PackageCheck },
            { id: 'stock', label: 'Current Stock', icon: Package },
            { id: 'material-master', label: 'Material Master', icon: FileText },
            { id: 'customer-master', label: 'Customer Master', icon: Users },
            { id: 'invoices', label: 'Sales Invoices', icon: FileText },
            { id: 'customers', label: 'Customer Analysis', icon: TrendingUp },
          ].map(tab => (
            <button
              key={tab.id}
              onClick={() => setActiveTab(tab.id as any)}
              className={cn(
                "flex items-center gap-3 px-3 py-2.5 rounded-xl text-[13px] font-semibold transition-all duration-200 text-left relative group",
                activeTab === tab.id 
                  ? "bg-primary text-white shadow-lg shadow-primary/10" 
                  : "text-text-muted hover:bg-surface2 hover:text-text-main"
              )}
              title={sidebarCollapsed ? tab.label : ''}
            >
              <tab.icon className={cn("w-5 h-5 shrink-0", activeTab === tab.id ? "text-white" : "text-text-muted group-hover:text-text-main")} />
              {!sidebarCollapsed && <span>{tab.label}</span>}
            </button>
          ))}
        </nav>

        <div className="mt-auto space-y-1">
           {!sidebarCollapsed && (
             <div className="p-4 bg-surface2 rounded-2xl mb-4 border border-border-custom">
                <div className="text-[10px] font-bold text-text-muted uppercase mb-2 tracking-widest">Database Sync</div>
                <div className="flex items-center gap-2">
                   <div className={cn("w-2 h-2 rounded-full", isSyncing ? "bg-primary animate-pulse" : "bg-avail")} />
                   <span className="text-xs font-bold font-mono">{isSyncing ? "SYNCING..." : "LIVE_FEED_01"}</span>
                </div>
             </div>
           )}
           <button className="flex items-center gap-3 px-3 py-2.5 w-full rounded-xl text-[13px] font-semibold text-text-muted hover:bg-surface2 transition-all">
             <RefreshCw className="w-4 h-4" />
             {!sidebarCollapsed && <span>Sync Data</span>}
           </button>
        </div>
      </aside>

      {/* MAIN CONTENT AREA */}
      <div className="flex-1 flex flex-col min-w-0 overflow-hidden relative">
        <header className="h-20 bg-surface border-b border-border-custom flex items-center justify-between px-10 shrink-0">
          <div className="relative group">
            <Search className="absolute left-4 top-1/2 -translate-y-1/2 w-4 h-4 text-text-muted transition-colors group-focus-within:text-primary" />
            <input 
              placeholder="Search across portfolio..." 
              value={searchTerm}
              onChange={e => setSearchTerm(e.target.value)}
              className="bg-surface2 rounded-full pl-11 pr-6 py-2.5 text-sm w-[380px] border border-transparent focus:border-primary focus:bg-white outline-none transition-all shadow-sm focus:shadow-md"
            />
          </div>

          <div className="flex items-center gap-6">
             <div className="relative" ref={uploadMenuRef}>
               <button 
                 onClick={() => setShowUploadMenu(!showUploadMenu)}
                 className="flex items-center gap-2 bg-primary text-white px-5 py-2.5 rounded-xl text-xs font-bold shadow-lg shadow-primary/20 transition-all hover:bg-primary/90 active:scale-95"
               >
                 <Upload className="w-4 h-4" /> UPLOAD EXCEL
               </button>
               
               <AnimatePresence>
                 {showUploadMenu && (
                   <motion.div 
                     initial={{ opacity: 0, y: 10, scale: 0.95 }}
                     animate={{ opacity: 1, y: 0, scale: 1 }}
                     exit={{ opacity: 0, y: 10, scale: 0.95 }}
                     className="absolute top-full right-0 mt-3 w-64 bg-white border border-border-custom rounded-2xl shadow-2xl z-50 py-3 overflow-hidden origin-top-right"
                   >
                     <div className="px-4 py-2 mb-1">
                        <div className="text-[10px] font-black text-text-muted uppercase tracking-widest">Select Data Type</div>
                     </div>
                     <button 
                       onClick={() => { fileInputRef.current?.click(); setShowUploadMenu(false); }}
                       className="w-full px-4 py-2.5 text-left text-[13px] font-bold text-text-main hover:bg-slate-50 flex items-center gap-3 transition-colors border-l-4 border-l-transparent hover:border-l-primary"
                     >
                       <ClipboardList className="w-4 h-4 text-primary" />
                       <span>Sales Orders (SO)</span>
                     </button>
                     <button 
                       onClick={() => { fileInputPORef.current?.click(); setShowUploadMenu(false); }}
                       className="w-full px-4 py-2.5 text-left text-[13px] font-bold text-text-main hover:bg-slate-50 flex items-center gap-3 transition-colors border-l-4 border-l-transparent hover:border-l-primary"
                     >
                       <PackageCheck className="w-4 h-4 text-primary" />
                       <span>Purchase Orders (PO)</span>
                     </button>
                     <button 
                       onClick={() => { fileInputStockRef.current?.click(); setShowUploadMenu(false); }}
                       className="w-full px-4 py-2.5 text-left text-[13px] font-bold text-text-main hover:bg-slate-50 flex items-center gap-3 transition-colors border-l-4 border-l-transparent hover:border-l-primary"
                     >
                       <Package className="w-4 h-4 text-primary" />
                       <span>Stock Inventory</span>
                     </button>
                     <div className="h-px bg-border-custom my-2 mx-4" />
                     <button 
                       onClick={() => { fileInputMaterialRef.current?.click(); setShowUploadMenu(false); }}
                       className="w-full px-4 py-2.5 text-left text-[13px] font-bold text-text-main hover:bg-slate-50 flex items-center gap-3 transition-colors border-l-4 border-l-transparent hover:border-l-text-muted"
                     >
                       <FileText className="w-4 h-4 text-text-muted" />
                       <span>Material Master</span>
                     </button>
                     <button 
                       onClick={() => { fileInputCustomerRef.current?.click(); setShowUploadMenu(false); }}
                       className="w-full px-4 py-2.5 text-left text-[13px] font-bold text-text-main hover:bg-slate-50 flex items-center gap-3 transition-colors border-l-4 border-l-transparent hover:border-l-text-muted"
                     >
                       <Users className="w-4 h-4 text-text-muted" />
                       <span>Customer Master</span>
                     </button>
                     <button 
                       onClick={() => { fileInputInvoiceRef.current?.click(); setShowUploadMenu(false); }}
                       className="w-full px-4 py-2.5 text-left text-[13px] font-bold text-text-main hover:bg-slate-50 flex items-center gap-3 transition-colors border-l-4 border-l-transparent hover:border-l-text-muted"
                     >
                       <FileText className="w-4 h-4 text-text-muted" />
                       <span>Sales Invoices</span>
                     </button>
                   </motion.div>
                 )}
               </AnimatePresence>
             </div>

             <div className="text-right hidden sm:block">
               <div className="text-[14px] font-bold text-text-main tracking-tight">Admin User</div>
               <div className="text-[11px] font-bold text-text-muted uppercase tracking-widest leading-none mt-0.5">Control Access</div>
             </div>
             <div className="w-10 h-10 rounded-full bg-slate-200 border-2 border-surface shadow-sm object-cover overflow-hidden">
               <img src="https://picsum.photos/seed/user/100/100" alt="Avatar" />
             </div>
          </div>
        </header>

        <main className="flex-1 overflow-y-auto scrollbar-custom bg-bg p-10">
          <AnimatePresence mode="wait">
            {activeTab === 'dashboard' && (
              <motion.div 
                key="dashboard"
                initial={{ opacity: 0, x: -10 }} 
                animate={{ opacity: 1, x: 0 }} 
                exit={{ opacity: 0, x: 10 }}
                className="space-y-8"
              >
                {/* FILTERS */}
                <div className="bg-surface border border-border-custom rounded-2xl p-5 flex flex-wrap items-center gap-4 shadow-sm">
                  <div className="flex flex-col gap-1.5">
                    <label className="text-[10px] font-bold text-text-muted uppercase px-1">Make</label>
                    <select 
                      value={dMake} 
                      onChange={e => setDMake(e.target.value)}
                      className="bg-surface2 border border-border-custom rounded-xl px-4 py-2 outline-none focus:border-primary text-[13px] min-w-[140px] font-medium"
                    >
                      <option value="">All Makes</option>
                      {META.makes.map(m => <option key={m} value={m}>{m}</option>)}
                    </select>
                  </div>

                  <div className="flex flex-col gap-1.5">
                    <label className="text-[10px] font-bold text-text-muted uppercase px-1">Group</label>
                    <select 
                      value={dGroup} 
                      onChange={e => setDGroup(e.target.value)}
                      className="bg-surface2 border border-border-custom rounded-xl px-4 py-2 outline-none focus:border-primary text-[13px] min-w-[180px] font-medium"
                    >
                      <option value="">All Regions</option>
                      {META.groups.map(g => <option key={g} value={g}>{g}</option>)}
                    </select>
                  </div>

                  <div className="flex flex-col gap-1.5">
                    <label className="text-[10px] font-bold text-text-muted uppercase px-1">Order Type</label>
                    <select 
                      value={dOrderType} 
                      onChange={e => setDOrderType(e.target.value)}
                      className="bg-surface2 border border-border-custom rounded-xl px-4 py-2 outline-none focus:border-primary text-[13px] min-w-[140px] font-medium"
                    >
                      <option value="">All Orders</option>
                      <option value="Due">Due Only</option>
                      <option value="Schedule">Schedule Only</option>
                    </select>
                  </div>

                  <div className="flex items-center gap-2 ml-auto">
                    <button 
                      onClick={() => { setDMake(''); setDGroup(''); setDOrderType(''); setDCGroup(''); setSearchTerm(''); }}
                      className="flex items-center gap-1.5 bg-white border border-border-custom text-text-muted px-3 py-1.5 rounded-lg text-[10px] font-black uppercase tracking-wider hover:bg-slate-50 active:scale-95 transition-all shadow-sm"
                    >
                      <RefreshCw className="w-3.5 h-3.5" /> RESET FILTERS
                    </button>
                  </div>
                </div>

                {/* KPI ROW */}
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
                  <StatCard 
                    title="📦 SO Portfolio Summary" 
                    value={fmtCur(dashboardStats.total)} 
                    details={[
                      { label: 'Total Value', value: fmtCur(dashboardStats.total), color: 'text-text-main' },
                      { label: 'No of lines pending', value: String(dashboardStats.count), color: 'text-text-muted' },
                      { label: 'Total SO Pending (Uniq)', value: String(dashboardStats.uniqueOrders), color: 'text-primary' },
                      { label: 'Uniq Customers', value: String(dashboardStats.uniqueCustomers), color: 'text-avail' },
                    ]}
                  />
                  <StatCard 
                    title="🚚 Open PO Summary" 
                    value={fmtCur(dashboardStats.totalPO)} 
                    details={[
                      { label: 'Total Value', value: fmtCur(dashboardStats.totalPO), color: 'text-text-main' },
                      { label: 'No of lines pending', value: String(dashboardStats.poCount), color: 'text-text-muted' },
                      { label: 'Uniq PO Pending', value: String(dashboardStats.uniquePO), color: 'text-primary' },
                      { label: 'Uniq Suppliers', value: String(dashboardStats.uniqueSuppliers), color: 'text-avail' },
                    ]}
                  />
                  <StatCard 
                    title="🔴 Due (<= 30.04.2026)" 
                    value={fmtCur(dashboardStats.dueVal)}
                    type="due"
                    details={[
                      { label: 'Stock Available', value: fmtCur(dashboardStats.dueAvail), color: 'text-avail' },
                      { label: 'Need to Arrange', value: fmtCur(dashboardStats.dueArr), color: 'text-danger' },
                    ]}
                  />
                  <StatCard 
                    title="🔵 Schedule (> 30.04.2026)" 
                    value={fmtCur(dashboardStats.schedVal)} 
                    type="sched"
                    details={[
                      { label: 'Stock Available', value: fmtCur(dashboardStats.schedAvail), color: 'text-primary' },
                      { label: 'Need to Arrange', value: fmtCur(dashboardStats.schedArr), color: 'text-text-muted' },
                    ]}
                  />
                </div>

                {/* CHARTS CONTAINER */}
                <div className="grid grid-cols-1 xl:grid-cols-2 gap-6">
                   <div className="bg-white border border-border-custom rounded-2xl p-8 shadow-sm">
                      <h3 className="text-sm font-black text-text-main uppercase tracking-tight mb-8">Pending Value Split</h3>
                      <div className="h-[300px]">
                         <ResponsiveContainer width="100%" height="100%">
                            <BarChart data={dashboardChartsData.pendingTypeData} layout="vertical">
                               <XAxis type="number" hide />
                               <YAxis dataKey="name" type="category" stroke="#64748b" fontSize={11} width={100} />
                               <Tooltip formatter={(v: any) => fmtCur(v)} cursor={{fill: 'transparent'}} />
                               <Bar dataKey="value" radius={[0, 4, 4, 0]}>
                                 <LabelList dataKey="value" position="right" formatter={(v: number) => fmtCur(v)} style={{ fontSize: '9px', fontWeight: 'bold', fill: '#64748b' }} />
                                 {dashboardChartsData.pendingTypeData.map((entry: any, index: number) => (
                                   <Cell key={`cell-${index}`} fill={index === 0 ? '#ff4d4f' : '#1890ff'} />
                                 ))}
                               </Bar>
                            </BarChart>
                         </ResponsiveContainer>
                      </div>
                   </div>

                   <div className="bg-white border border-border-custom rounded-2xl p-8 shadow-sm">
                      <h3 className="text-sm font-black text-text-main uppercase tracking-tight mb-8">Make Wise Breakdown</h3>
                      <div className="h-[300px]">
                         <ResponsiveContainer width="100%" height="100%">
                            <BarChart data={dashboardChartsData.makeStackedData}>
                               <CartesianGrid strokeDasharray="3 3" vertical={false} opacity={0.3} />
                               <XAxis dataKey="name" stroke="#64748b" fontSize={11} dy={10} />
                               <YAxis stroke="#64748b" fontSize={11} tickFormatter={v => fmtCur(v)} />
                               <Tooltip formatter={(v: any) => fmtCur(v)} contentStyle={{borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)'}} />
                               <Legend verticalAlign="top" align="right" height={36} />
                               <Bar dataKey="due" name="Due Value" fill="#ff4d4f" stackId="make">
                                   <LabelList dataKey="due" position="center" formatter={(v: number) => v > 0 ? fmtCur(v) : ''} style={{ fontSize: '8px', fontWeight: 'bold', fill: '#fff' }} />
                                </Bar>
                                <Bar dataKey="schedule" name="Schedule Value" fill="#1890ff" stackId="make">
                                   <LabelList dataKey="schedule" position="center" formatter={(v: number) => v > 0 ? fmtCur(v) : ''} style={{ fontSize: '8px', fontWeight: 'bold', fill: '#fff' }} />
                                </Bar>
                            </BarChart>
                         </ResponsiveContainer>
                      </div>
                   </div>
                </div>

                {/* RECENT TOP CLIENTS */}
                <div className="bg-surface border border-border-custom shadow-sm overflow-hidden">
                  <div className="px-4 py-3 border-b border-border-custom bg-surface2/30 flex justify-between items-center text-[Cambria]">
                    <h3 className="text-[12px] font-black text-text-main uppercase tracking-tight">Top 10 Customers Pending SO Breakdown</h3>
                    <button 
                      onClick={() => exportToExcel(dashboardChartsData.top10Detailed, 'Top_10_Customers_Breakdown')}
                      className="flex items-center gap-1.5 bg-text-main text-white px-3 py-1.5 rounded-lg text-[9px] font-black uppercase tracking-wider shadow-md hover:bg-primary active:scale-95 transition-all"
                    >
                      <Download className="w-3 h-3" /> EXPORT REPORT
                    </button>
                  </div>
                  <div className="overflow-x-auto scrollbar-custom">
                    <table className="excel-table">
                      <thead>
                        <tr>
                          <Th sortKey="group" onSort={handleSort} activeField={sortField} direction={sortDirection}>Group / Area</Th>
                          <Th sortKey="cgroup" onSort={handleSort} activeField={sortField} direction={sortDirection}>Customer Group</Th>
                          <Th sortKey="name" onSort={handleSort} activeField={sortField} direction={sortDirection}>Customer Name</Th>
                          <th className="text-right">Due (Avail)</th>
                          <th className="text-right">Due (Arrange)</th>
                          <th className="text-right">Schedule</th>
                          <Th sortKey="total" onSort={handleSort} activeField={sortField} direction={sortDirection} className="text-right">Total Outstanding</Th>
                        </tr>
                      </thead>
                      <tbody>
                        {(dashboardChartsData.top10Detailed || []).map((c, i) => (
                          <tr key={i} className="group cursor-pointer" onClick={() => setShowSOPopup(c.name)}>
                            <td className="font-bold text-text-muted">{c.group}</td>
                            <td className="font-bold text-text-muted">{c.cgroup}</td>
                            <td className="font-black text-text-main uppercase">{c.name}</td>
                            <td className="text-right font-bold text-avail">{fmtCur(c.dueAvail)}</td>
                            <td className="text-right font-bold text-danger">{fmtCur(c.dueArr)}</td>
                            <td className="text-right font-bold text-primary">{fmtCur(c.sched)}</td>
                            <td className="text-right font-black text-text-main bg-slate-50/50">{fmtCur(c.total)}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </motion.div>
            )}

            {activeTab === 'pending-so' && (
              <motion.div 
                key="pending-so"
                initial={{ opacity: 0, x: -10 }} 
                animate={{ opacity: 1, x: 0 }} 
                exit={{ opacity: 0, x: 10 }}
                className="space-y-8"
              >
                  {/* SUMMARY LABELS */}
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                    <div className="bg-white border border-border-custom p-6 rounded-2xl shadow-sm">
                      <div className="text-[10px] font-black text-text-muted uppercase tracking-widest mb-1">Total Pending SO</div>
                      <div className="text-2xl font-black text-text-main">{fmtCur(customFilterSOTab.total)}</div>
                    </div>
                    <div className="bg-white border border-border-custom p-6 rounded-2xl shadow-sm border-l-4 border-l-danger">
                      <div className="text-[10px] font-black text-danger uppercase tracking-widest mb-1">Due Orders</div>
                      <div className="flex items-baseline gap-3">
                         <div className="text-2xl font-black text-text-main">{fmtCur(customFilterSOTab.dueVal)}</div>
                         <div className="text-[11px] font-bold text-avail">Avail: {fmtCur(customFilterSOTab.dueAvail)}</div>
                      </div>
                    </div>
                    <div className="bg-white border border-border-custom p-6 rounded-2xl shadow-sm border-l-4 border-l-primary">
                      <div className="text-[10px] font-black text-primary uppercase tracking-widest mb-1">Schedule Orders</div>
                      <div className="flex items-baseline gap-3">
                         <div className="text-2xl font-black text-text-main">{fmtCur(customFilterSOTab.schedVal)}</div>
                         <div className="text-[11px] font-bold text-primary opacity-70">Avail: {fmtCur(customFilterSOTab.schedAvail)}</div>
                      </div>
                    </div>
                  </div>

                  {/* FILTERS & ACTIONS */}
                  <div className="bg-surface border border-border-custom rounded-2xl p-6 flex flex-wrap items-end gap-5 shadow-sm">
                    <div className="flex flex-col gap-1.5 min-w-[120px]">
                      <label className="text-[10px] font-bold text-text-muted uppercase px-1">Order Class</label>
                      <select value={soType} onChange={e => setSoType(e.target.value)} className="bg-surface2 border border-border-custom rounded-xl px-4 py-2.5 text-xs font-semibold outline-none focus:border-primary">
                        <option value="">All Classes</option>
                        <option value="Due">Due Orders</option>
                        <option value="Schedule">Schedule Orders</option>
                      </select>
                    </div>

                    <div className="flex flex-col gap-1.5 min-w-[150px]">
                      <label className="text-[10px] font-bold text-text-muted uppercase px-1">Allocation Status</label>
                      <select value={soStatus} onChange={e => setSoStatus(e.target.value)} className="bg-surface2 border border-border-custom rounded-xl px-4 py-2.5 text-xs font-semibold outline-none focus:border-primary">
                        <option value="">All Status</option>
                        <option value="Available">Available</option>
                        <option value="PO Exist - Expedite">PO Exist - Expedite</option>
                        <option value="Need to Place Order">Need to Place Order</option>
                      </select>
                    </div>

                    <div className="flex flex-col gap-1.5 min-w-[120px]">
                      <label className="text-[10px] font-bold text-text-muted uppercase px-1">Make</label>
                      <select value={soMake} onChange={e => setSoMake(e.target.value)} className="bg-surface2 border border-border-custom rounded-xl px-4 py-2.5 text-xs font-semibold outline-none focus:border-primary">
                        <option value="">All Makes</option>
                        {META.makes.map(m => <option key={m} value={m}>{m}</option>)}
                      </select>
                    </div>

                    <div className="flex flex-col gap-1.5 min-w-[120px]">
                      <label className="text-[10px] font-bold text-text-muted uppercase px-1">Group</label>
                      <select value={soGroup} onChange={e => setSoGroup(e.target.value)} className="bg-surface2 border border-border-custom rounded-xl px-4 py-2.5 text-xs font-semibold outline-none focus:border-primary">
                        <option value="">All Groups</option>
                        {META.groups.map(g => <option key={g} value={g}>{g}</option>)}
                      </select>
                    </div>

                    <div className="flex flex-col gap-1.5 flex-1 min-w-[240px]">
                      <label className="text-[10px] font-bold text-text-muted uppercase px-1">Universal Search</label>
                      <div className="relative">
                        <Search className="absolute left-3.5 top-2.5 w-4 h-4 text-text-muted" />
                        <input 
                           placeholder="Search Customer / Item / Order No..." 
                           value={soCust} 
                           onChange={e => setSoCust(e.target.value)}
                           className="pl-10 w-full bg-surface2 border border-border-custom rounded-xl px-4 py-2.5 text-xs font-semibold focus:border-primary outline-none"
                         />
                      </div>
                    </div>

                    <div className="flex items-center gap-2 ml-auto self-start">
                      <button 
                        onClick={() => fileInputRef.current?.click()}
                        className="flex items-center gap-1.5 bg-avail text-white px-3 py-1.5 rounded-lg text-[10px] font-black uppercase tracking-wider shadow-md hover:bg-avail/90 active:scale-95 transition-all"
                      >
                        <Upload className="w-3.5 h-3.5" /> UPLOAD
                      </button>
                      <button 
                        onClick={() => { if(confirm('Reset SO data?')) handleReset('so') }}
                        className="flex items-center gap-1.5 bg-white border border-border-custom text-text-muted px-3 py-1.5 rounded-lg text-[10px] font-black uppercase tracking-wider hover:bg-surface2 active:scale-95 transition-all shadow-sm"
                      >
                        <RefreshCw className="w-3.5 h-3.5" /> RESET
                      </button>
                      <button 
                        onClick={() => {
                          const exportData = filteredSO.map(r => ({
                            "Due Date": fmtDate(r.DueOn || r.Date),
                            "Voucher No": r.Order,
                            "Customer Name": r.PartyName,
                            "Item Description": r.NameOfItem,
                            "Material Code": r.MaterialCode,
                            "Pending Qty": r.Balance,
                            "Rate": r.Rate,
                            "Value": r.Value,
                            "Allocation Status": r.StockStatus
                          }));
                          exportToExcel(exportData, 'Pending_SO_Report');
                        }}
                        className="flex items-center gap-1.5 bg-text-main text-white px-3 py-1.5 rounded-lg text-[10px] font-black uppercase tracking-wider shadow-md hover:bg-primary active:scale-95 transition-all"
                      >
                        <Download className="w-3.5 h-3.5" /> EXTRACT
                      </button>
                    </div>
                    <input type="file" ref={fileInputRef} onChange={handleFileUpload} className="hidden" accept=".xlsx, .xls, .csv" />
                  </div>

                  <div className="bg-white border-x border-t border-border-custom px-6 py-4 flex justify-end gap-12 shadow-[0_-2px_10px_-4px_rgba(0,0,0,0.05)] rounded-t-2xl relative z-10">
                      <div className="flex flex-col items-end">
                         <div className="text-[9px] font-black text-text-muted uppercase tracking-widest mb-0.5">Total Bal Qty</div>
                         <div className="text-sm font-black text-text-main">{fmtNum(filteredSO.reduce((s, r) => s + (r.Balance || 0), 0))}</div>
                      </div>
                      <div className="flex flex-col items-end">
                         <div className="text-[9px] font-black text-primary uppercase tracking-widest mb-0.5">Total Market Value</div>
                         <div className="text-sm font-black text-primary">{fmtCur(filteredSO.reduce((s, r) => s + (r.Value || 0), 0))}</div>
                      </div>
                  </div>

                  {/* DATA GRID */}
                  <div className="bg-white border border-border-custom shadow-sm overflow-hidden">
                     <div className="overflow-x-auto scrollbar-custom max-h-[calc(100vh-450px)]">
                        <table className="excel-table">
                           <thead>
                              <tr>
                                 <Th sortKey="DueOn" onSort={handleSort} activeField={sortField} direction={sortDirection} className="whitespace-nowrap">Due Date</Th>
                                 <Th sortKey="Order" onSort={handleSort} activeField={sortField} direction={sortDirection}>Voucher No</Th>
                                 <Th sortKey="PartyName" onSort={handleSort} activeField={sortField} direction={sortDirection}>Customer Name</Th>
                                 <Th sortKey="NameOfItem" onSort={handleSort} activeField={sortField} direction={sortDirection}>Item Name / Description</Th>
                                 <Th sortKey="Balance" onSort={handleSort} activeField={sortField} direction={sortDirection} className="text-right">Bal Qty</Th>
                                 <Th sortKey="Value" onSort={handleSort} activeField={sortField} direction={sortDirection} className="text-right">Market Value</Th>
                                 <Th sortKey="StockStatus" onSort={handleSort} activeField={sortField} direction={sortDirection} className="text-center">Allocation status</Th>
                              </tr>
                           </thead>
                           <tbody className="bg-white">
                              {(filteredSO || []).map((r, idx) => (
                                <tr key={idx} className="hover:bg-slate-50 transition-colors group">
                                  <td className={cn("font-bold whitespace-nowrap", r.OrderType === 'Due' ? "text-danger" : "text-primary")}>
                                     {fmtDate(r.DueOn || r.Date)}
                                  </td>
                                  <td className="font-mono text-text-muted uppercase">{r.Order}</td>
                                  <td className="font-black text-text-main uppercase group-hover:text-primary transition-colors cursor-pointer" onClick={() => setShowSOPopup(r.PartyName)}>
                                    {r.PartyName}
                                  </td>
                                  <td className="font-bold text-text-muted whitespace-normal flex-wrap leading-tight max-w-[300px]">
                                    {r.NameOfItem}
                                  </td>
                                  <td className="text-right font-black text-text-main">{fmtNum(r.Balance)}</td>
                                  <td className="text-right font-black text-primary">{fmtCur(r.Value)}</td>
                                  <td className="text-center">
                                     <div className={cn(
                                       "px-2 py-0.5 rounded text-[9px] font-black uppercase inline-block border",
                                       r.StockStatus === 'Available' ? "bg-avail/10 text-avail border-avail/10" :
                                       r.StockStatus.includes('PO') ? "bg-primary/10 text-primary border-primary/10" :
                                       "bg-danger/10 text-danger border-danger/10"
                                     )}>
                                        {r.StockStatus}
                                     </div>
                                  </td>
                                </tr>
                              ))}
                           </tbody>
                        </table>
                     </div>
                  </div>
              </motion.div>
            )}

            {activeTab === 'pending-po' && (
              <motion.div 
                key="pending-po"
                initial={{ opacity: 0, x: -10 }} 
                animate={{ opacity: 1, x: 0 }} 
                exit={{ opacity: 0, x: 10 }}
                className="space-y-8"
              >
                  <div className="grid grid-cols-1 md:grid-cols-4 gap-6">
                    <div className="bg-white border border-border-custom p-6 rounded-2xl shadow-sm">
                      <div className="text-[10px] font-black text-text-muted uppercase tracking-widest mb-1">Total PO Value</div>
                      <div className="text-2xl font-black text-text-main">{fmtCur(poStats.total)}</div>
                    </div>
                    <div className="bg-white border border-border-custom p-6 rounded-2xl shadow-sm">
                      <div className="text-[10px] font-black text-text-muted uppercase tracking-widest mb-1">No of lines pending</div>
                      <div className="text-2xl font-black text-text-main">{poStats.count}</div>
                    </div>
                    <div className="bg-white border border-border-custom p-6 rounded-2xl shadow-sm">
                      <div className="text-[10px] font-black text-primary uppercase tracking-widest mb-1">Uniq PO Pending</div>
                      <div className="text-2xl font-black text-primary">{poStats.uniqueOrders}</div>
                    </div>
                    <div className="bg-white border border-border-custom p-6 rounded-2xl shadow-sm">
                      <div className="text-[10px] font-black text-avail uppercase tracking-widest mb-1">Uniq Suppliers</div>
                      <div className="text-2xl font-black text-avail">{poStats.uniqueSuppliers}</div>
                    </div>
                  </div>

                  <div className="bg-surface border border-border-custom shadow-sm p-6 flex flex-wrap items-end gap-5">
                    <div className="flex flex-col gap-1.5 flex-1 min-w-[240px]">
                      <label className="text-[10px] font-bold text-text-muted uppercase px-1">Global PO Search</label>
                      <div className="relative">
                        <Search className="absolute left-3.5 top-2.5 w-4 h-4 text-text-muted" />
                        <input 
                          placeholder="Search Party / Item / PO No..." 
                          value={poSearch} 
                          onChange={e => setPoSearch(e.target.value)}
                          className="pl-10 w-full bg-surface2 border border-border-custom rounded-xl px-4 py-2.5 text-xs font-semibold focus:border-primary outline-none focus:bg-white transition-all"
                        />
                      </div>
                    </div>

                    <div className="flex items-center gap-2 ml-auto self-start">
                      <button 
                        onClick={() => fileInputPORef.current?.click()}
                        className="flex items-center gap-1.5 bg-avail text-white px-3 py-1.5 rounded-lg text-[10px] font-black uppercase tracking-wider shadow-md hover:bg-avail/90 active:scale-95 transition-all"
                      >
                        <Upload className="w-3.5 h-3.5" /> UPLOAD PO
                      </button>
                      <button 
                         onClick={() => { if(confirm('Reset PO data?')) handleReset('po') }}
                         className="flex items-center gap-1.5 bg-white border border-border-custom text-text-muted px-3 py-1.5 rounded-lg text-[10px] font-black uppercase tracking-wider hover:bg-slate-50 active:scale-95 transition-all shadow-sm"
                      >
                         <RefreshCw className="w-3.5 h-3.5" /> RESET
                      </button>
                      <button 
                        onClick={() => {
                          const exportData = filteredPOList.map(r => ({
                            "Date": r.Date,
                            "PO No": r.Order,
                            "Supplier Name": r.PartyName,
                            "Item Description": r.NameOfItem,
                            "Pending Qty": r.Balance,
                            "Rate": r.Rate,
                            "Value": r.Value
                          }));
                          exportToExcel(exportData, 'Pending_PO_Report');
                        }}
                        className="flex items-center gap-1.5 bg-text-main text-white px-3 py-1.5 rounded-lg text-[10px] font-black uppercase tracking-wider shadow-md hover:bg-primary active:scale-95 transition-all"
                      >
                        <Download className="w-3.5 h-3.5" /> EXPORT
                      </button>
                    </div>
                    <input type="file" ref={fileInputPORef} onChange={handlePOUpload} className="hidden" accept=".xlsx, .xls, .csv" />
                  </div>

                  <div className="bg-surface border border-border-custom shadow-sm overflow-hidden">
                     <div className="overflow-x-auto scrollbar-custom max-h-[calc(100vh-320px)]">
                        <table className="excel-table">
                           <thead>
                              <tr>
                                 <Th sortKey="Date" onSort={handleSort} activeField={sortField} direction={sortDirection} className="whitespace-nowrap">Date</Th>
                                 <Th sortKey="Order" onSort={handleSort} activeField={sortField} direction={sortDirection}>Order #</Th>
                                 <Th sortKey="PartyName" onSort={handleSort} activeField={sortField} direction={sortDirection}>Supplier/Party</Th>
                                 <Th sortKey="NameOfItem" onSort={handleSort} activeField={sortField} direction={sortDirection}>Item Details</Th>
                                 <Th sortKey="Ordered" onSort={handleSort} activeField={sortField} direction={sortDirection} className="text-right">Qty</Th>
                                 <Th sortKey="Balance" onSort={handleSort} activeField={sortField} direction={sortDirection} className="text-right">Balance</Th>
                                 <Th sortKey="Rate" onSort={handleSort} activeField={sortField} direction={sortDirection} className="text-right">Rate</Th>
                                 <Th sortKey="Value" onSort={handleSort} activeField={sortField} direction={sortDirection} className="text-right">Value</Th>
                                 <Th sortKey="DueOn" onSort={handleSort} activeField={sortField} direction={sortDirection} className="text-center whitespace-nowrap">Due On</Th>
                              </tr>
                           </thead>
                           <tbody className="bg-white">
                              {(filteredPOList || []).map((p, idx) => (
                                <tr key={idx} className="hover:bg-slate-50 transition-colors group">
                                  <td className="whitespace-nowrap">{fmtDate(p.Date)}</td>
                                  <td className="font-mono text-text-muted uppercase">{p.Order}</td>
                                  <td className="font-black text-text-main uppercase">{p.PartyName}</td>
                                  <td className="whitespace-normal leading-tight max-w-[200px]">
                                     <div className="font-bold text-text-main">{p.NameOfItem}</div>
                                     <div className="text-[9px] text-text-muted font-mono">{p.PartNo}</div>
                                  </td>
                                  <td className="text-right font-medium opacity-60">{fmtNum(p.Ordered)}</td>
                                  <td className="text-right font-black text-text-main">{fmtNum(p.Balance)}</td>
                                  <td className="text-right font-black text-text-muted">{fmtCur(p.Rate)}</td>
                                  <td className="text-right font-black text-primary">{fmtCur(p.Value)}</td>
                                  <td className="text-center font-bold text-text-muted whitespace-nowrap">
                                     {fmtDate(p.DueOn)}
                                  </td>
                                </tr>
                              ))}
                           </tbody>
                        </table>
                     </div>
                  </div>
              </motion.div>
            )}

            {activeTab === 'stock' && (
              <motion.div 
                key="stock"
                initial={{ opacity: 0, x: -10 }} 
                animate={{ opacity: 1, x: 0 }} 
                exit={{ opacity: 0, x: 10 }}
                className="space-y-8"
              >
                  <div className="bg-surface border border-border-custom shadow-sm p-6 flex flex-wrap items-end gap-5">
                    <div className="flex flex-col gap-1.5 flex-1 min-w-[240px]">
                      <label className="text-[10px] font-bold text-text-muted uppercase px-1">Inventory Item Search</label>
                      <div className="relative">
                        <Search className="absolute left-3.5 top-2.5 w-4 h-4 text-text-muted" />
                        <input 
                          placeholder="Search Particulars..." 
                          value={stockSearch} 
                          onChange={e => setStockSearch(e.target.value)}
                          className="pl-10 w-full bg-surface2 border border-border-custom rounded-xl px-4 py-2.5 text-xs font-semibold focus:border-primary outline-none focus:bg-white transition-all"
                        />
                      </div>
                    </div>

                    <div className="flex items-center gap-2 ml-auto self-start">
                      <button 
                        onClick={() => fileInputStockRef.current?.click()}
                        className="flex items-center gap-1.5 bg-avail text-white px-3 py-1.5 rounded-lg text-[10px] font-black uppercase tracking-wider shadow-md hover:bg-avail/90 active:scale-95 transition-all"
                      >
                        <Upload className="w-3.5 h-3.5" /> UPLOAD STOCK
                      </button>
                      <button 
                         onClick={() => { if(confirm('Reset Inventory data?')) handleReset('stock') }}
                         className="flex items-center gap-1.5 bg-white border border-border-custom text-text-muted px-3 py-1.5 rounded-lg text-[10px] font-black uppercase tracking-wider hover:bg-slate-50 active:scale-95 transition-all shadow-sm"
                      >
                         <RefreshCw className="w-3.5 h-3.5" /> RESET
                      </button>
                      <button 
                        onClick={() => {
                          const exportData = filteredStockList.map(r => ({
                            "Item Description": r.Particulars,
                            "Available Stock": r.Quantity
                          }));
                          exportToExcel(exportData, 'Stock_Inventory_Report');
                        }}
                        className="flex items-center gap-1.5 bg-text-main text-white px-3 py-1.5 rounded-lg text-[10px] font-black uppercase tracking-wider shadow-md hover:bg-primary active:scale-95 transition-all"
                      >
                        <Download className="w-3.5 h-3.5" /> EXPORT
                      </button>
                    </div>
                    <input type="file" ref={fileInputStockRef} onChange={handleStockUpload} className="hidden" accept=".xlsx, .xls, .csv" />
                  </div>

                  <div className="bg-surface border border-border-custom shadow-sm overflow-hidden">
                     <div className="overflow-x-auto scrollbar-custom max-h-[calc(100vh-340px)]">
                        <table className="excel-table">
                           <thead>
                              <tr>
                                 <Th sortKey="Particulars" onSort={handleSort} activeField={sortField} direction={sortDirection}>Description / Particulars</Th>
                                 <Th sortKey="Quantity" onSort={handleSort} activeField={sortField} direction={sortDirection} className="text-right">Quantity</Th>
                                 <Th sortKey="Rate" onSort={handleSort} activeField={sortField} direction={sortDirection} className="text-right">Rate</Th>
                                 <Th sortKey="Value" onSort={handleSort} activeField={sortField} direction={sortDirection} className="text-right">Total Valuation</Th>
                              </tr>
                           </thead>
                           <tbody className="bg-white">
                              {(filteredStockList || []).map((s, idx) => (
                                <tr key={idx} className="hover:bg-slate-50 transition-colors group">
                                  <td className="font-black text-text-main uppercase whitespace-normal leading-tight max-w-[400px]">{s.Particulars}</td>
                                  <td className="text-right font-black text-text-main">{fmtNum(s.Quantity)}</td>
                                  <td className="text-right font-black text-text-muted">{fmtCur(s.Rate)}</td>
                                  <td className="text-right font-black text-primary bg-slate-50/30">{fmtCur(s.Value)}</td>
                                </tr>
                              ))}
                           </tbody>
                           {filteredStockList.length > 0 && (
                             <tfoot className="sticky bottom-0 bg-white border-t-2 border-grid">
                               <tr className="font-black text-text-main">
                                  <td className="text-right uppercase tracking-wider text-[10px] text-text-muted">Consolidated Inventory Value</td>
                                  <td></td>
                                  <td></td>
                                  <td className="text-right text-[14px] text-primary bg-primary/5 border-l border-grid">
                                    {fmtCur(filteredStockList.reduce((sum, item) => sum + item.Value, 0))}
                                  </td>
                               </tr>
                             </tfoot>
                           )}
                        </table>
                     </div>
                  </div>
              </motion.div>
            )}

            {activeTab === 'material-master' && (
              <motion.div 
                key="material-master"
                initial={{ opacity: 0, x: -10 }} 
                animate={{ opacity: 1, x: 0 }} 
                exit={{ opacity: 0, x: 10 }}
                className="space-y-8"
              >
                  <div className="bg-surface border border-border-custom shadow-sm p-6 flex flex-wrap items-end gap-5">
                    <div className="flex flex-col gap-1.5 flex-1 min-w-[240px]">
                      <label className="text-[10px] font-bold text-text-muted uppercase px-1">Material Analytics Search</label>
                      <div className="relative">
                        <Search className="absolute left-3.5 top-2.5 w-4 h-4 text-text-muted" />
                        <input 
                          placeholder="Search Description or Part No..." 
                          value={materialSearch} 
                          onChange={e => setMaterialSearch(e.target.value)}
                          className="pl-10 w-full bg-surface2 border border-border-custom rounded-xl px-4 py-2.5 text-xs font-semibold focus:border-primary outline-none focus:bg-white transition-all shadow-inner"
                        />
                      </div>
                    </div>

                    <div className="flex items-center gap-2 ml-auto self-start">
                      <button 
                        onClick={() => fileInputMaterialRef.current?.click()}
                        className="flex items-center gap-1.5 bg-avail text-white px-3 py-1.5 rounded-lg text-[10px] font-black uppercase tracking-wider shadow-md hover:bg-avail/90 active:scale-95 transition-all"
                      >
                        <Upload className="w-3.5 h-3.5" /> UPLOAD MASTER
                      </button>
                      <button 
                         onClick={() => { if(confirm('Reset Material Master?')) handleReset('material') }}
                         className="flex items-center gap-1.5 bg-white border border-border-custom text-text-muted px-3 py-1.5 rounded-lg text-[10px] font-black uppercase tracking-wider hover:bg-slate-50 active:scale-95 transition-all shadow-sm"
                      >
                         <RefreshCw className="w-3.5 h-3.5" /> RESET
                      </button>
                      <button 
                        onClick={() => {
                          const exportData = filteredMaterialList.map(r => ({
                            "Material Code": r.MaterialCode,
                            "Material Name": r.MaterialName,
                            "Group": r.Group,
                            "Base Unit": r.BaseUnit
                          }));
                          exportToExcel(exportData, 'Material_Master_Report');
                        }}
                        className="flex items-center gap-1.5 bg-text-main text-white px-3 py-1.5 rounded-lg text-[10px] font-black uppercase tracking-wider shadow-md hover:bg-primary active:scale-95 transition-all"
                      >
                        <Download className="w-3.5 h-3.5" /> EXPORT
                      </button>
                    </div>
                    <input type="file" ref={fileInputMaterialRef} onChange={handleMaterialUpload} className="hidden" accept=".xlsx, .xls, .csv" />
                  </div>

                  <div className="bg-surface border border-border-custom shadow-sm overflow-hidden">
                     <div className="overflow-x-auto scrollbar-custom max-h-[calc(100vh-320px)]">
                        <table className="excel-table">
                           <thead>
                              <tr>
                                 <Th sortKey="Description" onSort={handleSort} activeField={sortField} direction={sortDirection}>Item Description</Th>
                                 <Th sortKey="PartNo" onSort={handleSort} activeField={sortField} direction={sortDirection}>Part Number</Th>
                                 <Th sortKey="Make" onSort={handleSort} activeField={sortField} direction={sortDirection}>Make / Brand</Th>
                                 <Th sortKey="MaterialGroup" onSort={handleSort} activeField={sortField} direction={sortDirection}>Material Group</Th>
                              </tr>
                           </thead>
                           <tbody className="bg-white">
                              {(filteredMaterialList || []).map((m, idx) => (
                                <tr key={idx} className="hover:bg-slate-50 transition-colors">
                                  <td className="font-black text-text-main uppercase whitespace-normal leading-tight max-w-[400px]">{m.Description}</td>
                                  <td className="font-mono text-text-muted">{m.PartNo}</td>
                                  <td className="font-bold text-primary italic uppercase">{m.Make}</td>
                                  <td className="font-bold text-text-muted italic">{m.MaterialGroup}</td>
                                </tr>
                              ))}
                           </tbody>
                        </table>
                     </div>
                  </div>
               </motion.div>
             )}
             {activeTab === 'customer-master' && (
              <motion.div 
                key="customer-master"
                initial={{ opacity: 0, x: -10 }} 
                animate={{ opacity: 1, x: 0 }} 
                exit={{ opacity: 0, x: 10 }}
                className="space-y-8"
              >
                  <div className="bg-surface border border-border-custom shadow-sm p-6 flex flex-wrap items-end gap-5">
                    <div className="flex flex-col gap-1.5 flex-1 min-w-[240px]">
                      <label className="text-[10px] font-bold text-text-muted uppercase px-1">Customer CRM Search</label>
                      <div className="relative">
                        <Search className="absolute left-3.5 top-2.5 w-4 h-4 text-text-muted" />
                        <input 
                          placeholder="Search Customer Name..." 
                          value={customerMasterSearch} 
                          onChange={e => setCustomerMasterSearch(e.target.value)}
                          className="pl-10 w-full bg-surface2 border border-border-custom rounded-xl px-4 py-2.5 text-xs font-semibold focus:border-primary outline-none focus:bg-white transition-all shadow-inner"
                        />
                      </div>
                    </div>

                    <div className="flex items-center gap-2 ml-auto self-start">
                      <button 
                        onClick={() => fileInputCustomerRef.current?.click()}
                        className="flex items-center gap-1.5 bg-avail text-white px-3 py-1.5 rounded-lg text-[10px] font-black uppercase tracking-wider shadow-md hover:bg-avail/90 active:scale-95 transition-all"
                      >
                        <Upload className="w-3.5 h-3.5" /> UPLOAD CSV
                      </button>
                      <button 
                         onClick={() => { if(confirm('Reset Customer Master?')) handleReset('customer') }}
                         className="flex items-center gap-1.5 bg-white border border-border-custom text-text-muted px-3 py-1.5 rounded-lg text-[10px] font-black uppercase tracking-wider hover:bg-slate-50 active:scale-95 transition-all shadow-sm"
                      >
                         <RefreshCw className="w-3.5 h-3.5" /> RESET
                      </button>
                      <button 
                        onClick={() => {
                          const exportData = filteredCustomerMasterList.map(r => ({
                            "Customer Name": r.CustomerName,
                            "Group": r.Group,
                            "Status": r.Status,
                            "Segment": r.CustomerGroup
                          }));
                          exportToExcel(exportData, 'Customer_Master_Report');
                        }}
                        className="flex items-center gap-1.5 bg-text-main text-white px-3 py-1.5 rounded-lg text-[10px] font-black uppercase tracking-wider shadow-md hover:bg-primary active:scale-95 transition-all"
                      >
                        <Download className="w-3.5 h-3.5" /> EXPORT
                      </button>
                    </div>
                    <input type="file" ref={fileInputCustomerRef} onChange={handleCustomerUpload} className="hidden" accept=".xlsx, .xls, .csv" />
                  </div>

                  <div className="bg-surface border border-border-custom shadow-sm overflow-hidden">
                     <div className="overflow-x-auto scrollbar-custom max-h-[calc(100vh-320px)]">
                        <table className="excel-table">
                           <thead>
                              <tr>
                                 <Th sortKey="CustomerName" onSort={handleSort} activeField={sortField} direction={sortDirection}>Complete Customer Name</Th>
                                 <Th sortKey="Group" onSort={handleSort} activeField={sortField} direction={sortDirection}>Group / Area</Th>
                                 <Th sortKey="Status" onSort={handleSort} activeField={sortField} direction={sortDirection}>Interaction Status</Th>
                                 <Th sortKey="CustomerGroup" onSort={handleSort} activeField={sortField} direction={sortDirection}>Segment</Th>
                              </tr>
                           </thead>
                           <tbody className="bg-white">
                              {(filteredCustomerMasterList || []).map((c, idx) => (
                                <tr key={idx} className="hover:bg-slate-50 transition-colors group cursor-pointer" onClick={() => setShowSOPopup(c.CustomerName)}>
                                  <td className="font-black text-text-main uppercase whitespace-normal leading-tight max-w-[300px]">{c.CustomerName}</td>
                                  <td className="font-bold text-text-muted">{c.Group}</td>
                                  <td className="text-center">
                                      <span className={cn(
                                        "px-2 py-0.5 rounded text-[9px] font-black uppercase border",
                                        c.Status?.includes('REPETED') ? "bg-avail/10 text-avail border-avail/10" : 
                                        c.Status?.includes('LOST') ? "bg-danger/10 text-danger border-danger/10" : 
                                        "bg-primary/10 text-primary border-primary/10"
                                      )}>
                                        {c.Status}
                                      </span>
                                  </td>
                                  <td className="text-text-muted italic font-bold">{c.CustomerGroup || 'N/A'}</td>
                                </tr>
                              ))}
                           </tbody>
                        </table>
                     </div>
                  </div>
              </motion.div>
            )}

            {activeTab === 'invoices' && (
              <motion.div 
                key="invoices"
                initial={{ opacity: 0, x: -10 }} 
                animate={{ opacity: 1, x: 0 }} 
                exit={{ opacity: 0, x: 10 }}
                className="space-y-8"
              >
                  <div className="bg-surface border border-border-custom rounded-2xl p-6 flex flex-wrap items-end gap-5 shadow-sm">
                    <div className="flex flex-col gap-1.5 flex-1 min-w-[240px]">
                      <label className="text-[10px] font-bold text-text-muted uppercase px-1">Invoice Search</label>
                      <div className="relative">
                        <Search className="absolute left-3.5 top-2.5 w-4 h-4 text-text-muted" />
                        <input 
                          placeholder="Search Buyer or Voucher No..." 
                          value={invoiceSearch} 
                          onChange={e => setInvoiceSearch(e.target.value)}
                          className="pl-10 w-full bg-surface2 border border-border-custom rounded-xl px-4 py-2.5 text-xs font-semibold focus:border-primary outline-none"
                        />
                      </div>
                    </div>

                    <div className="flex items-center gap-2 ml-auto self-start">
                      <button 
                        onClick={() => fileInputInvoiceRef.current?.click()}
                        className="flex items-center gap-1.5 bg-avail text-white px-3 py-1.5 rounded-lg text-[10px] font-black uppercase tracking-wider shadow-md hover:bg-avail/90 active:scale-95 transition-all"
                      >
                        <Upload className="w-3.5 h-3.5" /> UPLOAD INVOICE
                      </button>
                      <button 
                        onClick={() => { if(confirm('Reset Invoice data?')) handleReset('invoice') }}
                        className="flex items-center gap-1.5 bg-white border border-border-custom text-text-muted px-3 py-1.5 rounded-lg text-[10px] font-black uppercase tracking-wider hover:bg-surface2 active:scale-95 transition-all shadow-sm"
                      >
                        <RefreshCw className="w-3.5 h-3.5" /> RESET
                      </button>
                      <button 
                        onClick={() => {
                          const exportData = filteredInvoiceList.map(r => ({
                            "Date": r.Date,
                            "Particulars": r.Particulars,
                            "Buyer": r.Buyer,
                            "Consignee": r.Consignee,
                            "Voucher Type": r.VoucherType,
                            "Voucher No": r.VoucherNo,
                            "Voucher Ref": r.VoucherRef,
                            "GSTIN": r.GSTIN,
                            "Quantity": r.Quantity,
                            "Value": r.Value
                          }));
                          exportToExcel(exportData, 'Sales_Invoice_Report');
                        }}
                        className="flex items-center gap-1.5 bg-text-main text-white px-3 py-1.5 rounded-lg text-[10px] font-black uppercase tracking-wider shadow-md hover:bg-primary active:scale-95 transition-all"
                      >
                        <Download className="w-3.5 h-3.5" /> EXPORT
                      </button>
                    </div>
                    <input 
                      type="file" 
                      ref={fileInputInvoiceRef} 
                      onChange={handleInvoiceUpload} 
                      className="hidden" 
                      accept=".xlsx, .xls, .csv" 
                    />
                  </div>

                  <div className="bg-surface border border-border-custom rounded-2xl overflow-hidden shadow-sm">
                     <div className="overflow-x-auto scrollbar-custom max-h-[calc(100vh-320px)]">
                        <table className="excel-table">
                           <thead>
                              <tr>
                                 <th className="whitespace-nowrap">Date</th>
                                 <th className="min-w-[200px]">Particulars</th>
                                 <th>Buyer</th>
                                 <th>Consignee</th>
                                 <th className="whitespace-nowrap">Voucher Type</th>
                                 <th className="whitespace-nowrap">Voucher No.</th>
                                 <th className="whitespace-nowrap">Voucher Ref. No.</th>
                                 <th className="whitespace-nowrap">GSTIN/UIN</th>
                                 <th className="text-right">Quantity</th>
                                 <th className="text-right">Value</th>
                              </tr>
                           </thead>
                           <tbody className="bg-white">
                              {(filteredInvoiceList || []).map((inv, idx) => (
                                <tr key={idx} className="hover:bg-slate-50 transition-colors group">
                                  <td className="whitespace-nowrap text-text-muted">{fmtDate(inv.Date)}</td>
                                  <td className="font-bold text-text-main">{inv.Particulars}</td>
                                  <td className="uppercase text-[11px] font-bold">{inv.Buyer}</td>
                                  <td className="uppercase text-[11px] font-bold text-text-muted">{inv.Consignee}</td>
                                  <td className="text-center">
                                    <span className="px-1.5 py-0.5 rounded bg-slate-100 text-text-muted text-[9px] font-black uppercase border border-border-custom">
                                      {inv.VoucherType}
                                    </span>
                                  </td>
                                  <td className="font-mono text-primary font-black uppercase">{inv.VoucherNo}</td>
                                  <td className="font-mono text-text-muted text-[10px]">{inv.VoucherRef}</td>
                                  <td className="font-mono text-text-muted text-[10px] uppercase">{inv.GSTIN}</td>
                                  <td className="text-right font-black text-text-main">{fmtNum(inv.Quantity)}</td>
                                  <td className="text-right font-black text-primary">{fmtCur(inv.Value)}</td>
                                </tr>
                              ))}
                              {filteredInvoiceList.length === 0 && (
                                 <tr>
                                    <td colSpan={10} className="py-20 text-center text-text-muted italic font-bold">
                                       No invoices uploaded. Upload your sales register to track billing history.
                                    </td>
                                 </tr>
                              )}
                           </tbody>
                        </table>
                     </div>
                  </div>
              </motion.div>
            )}

            {activeTab === 'customers' && (
              <motion.div 
                key="customers"
                initial={{ opacity: 0, scale: 0.98 }} 
                animate={{ opacity: 1, scale: 1 }} 
                exit={{ opacity: 0, scale: 1.02 }}
                className="space-y-8"
              >
                  <div className="flex justify-between items-center bg-surface border border-border-custom rounded-2xl p-6 shadow-sm">
                     <div className="flex gap-4 flex-1 max-w-3xl items-center">
                        <div className="relative flex-1">
                          <Search className="absolute left-4 top-1/2 -translate-y-1/2 w-4 h-4 text-text-muted" />
                          <input 
                            placeholder="Search Customer / Order ID / Invoice No / PO No..." 
                            className="bg-surface2 rounded-2xl pl-12 pr-6 py-3 text-sm w-full outline-none focus:ring-2 focus:ring-primary/20 transition-all font-medium border border-transparent focus:border-primary"
                            value={cSearch}
                            onChange={e => setCSearch(e.target.value)}
                          />
                        </div>
                        <select 
                          value={cGroup} 
                          onChange={e => setCGroup(e.target.value)}
                          className="bg-surface2 border border-border-custom rounded-2xl px-6 py-3 text-sm font-bold text-text-main"
                        >
                          <option value="">All Regions</option>
                          {META.groups.map(g => <option key={g} value={g}>{g}</option>)}
                        </select>
                        <button 
                          onClick={() => { setCSearch(''); setCGroup(''); setCCGroup(''); setSearchTerm(''); }}
                          className="flex items-center gap-1.5 bg-white border border-border-custom text-text-muted px-4 py-3 rounded-2xl text-[10px] font-black uppercase tracking-wider hover:bg-slate-50 active:scale-95 transition-all shadow-sm h-full"
                        >
                          <RefreshCw className="w-3.5 h-3.5" /> RESET
                        </button>
                     </div>
                      <button 
                         onClick={() => {
                           const exportData = customersList.map(c => ({
                             "Customer Name": c.name,
                             "Total Pending SO": c.total,
                             "Due Orders": c.dueVal,
                             "Schedule Orders": c.schedVal,
                             "Total Invoiced": c.invVal
                           }));
                           exportToExcel(exportData, 'Customer_Analysis_Report');
                         }}
                         className="flex items-center gap-2 bg-text-main text-white px-6 py-3 rounded-2xl text-xs font-bold shadow-lg hover:shadow-primary/20 transition-all hover:bg-primary"
                       >
                         <Download className="w-4 h-4" /> EXPORT REPORT
                       </button>
                  </div>

                  <div className="bg-surface border border-border-custom rounded-2xl overflow-hidden shadow-sm">
                     <div className="overflow-x-auto scrollbar-custom max-h-[calc(100vh-320px)]">
                        <table className="excel-table">
                           <thead>
                              <tr>
                                 <th>Customer Name</th>
                                 <th className="text-right">Total Pending SO</th>
                                 <th className="text-right">Due Orders</th>
                                 <th className="text-right">Schedule Orders</th>
                                 <th className="text-right">Total Invoiced</th>
                                 <th className="text-center">Action</th>
                              </tr>
                           </thead>
                           <tbody className="bg-white">
                              {(customersList || []).map((c, idx) => (
                                <tr key={idx} className="hover:bg-slate-50 transition-colors group">
                                  <td className="font-black text-text-main uppercase text-[13px]">{c.name}</td>
                                  <td className="text-right font-black text-primary">
                                     <button 
                                      onClick={() => setShowSOPopup(c.name)}
                                      className="hover:underline hover:text-primary/80 transition-all"
                                     >
                                       {fmtCur(c.total)}
                                     </button>
                                  </td>
                                  <td className="text-right font-bold text-danger">{fmtCur(c.dueVal)}</td>
                                  <td className="text-right font-bold text-text-muted">{fmtCur(c.schedVal)}</td>
                                  <td className="text-right font-black text-avail">
                                     <button 
                                      onClick={() => setShowInvPopup(c.name)}
                                      className="hover:underline hover:text-avail/80 transition-all"
                                     >
                                       {fmtCur(c.invVal)}
                                     </button>
                                  </td>
                                  <td className="text-center">
                                    <div className="flex justify-center gap-2">
                                      <button onClick={() => setShowSOPopup(c.name)} className="p-1.5 rounded-lg bg-primary/10 text-primary hover:bg-primary hover:text-white transition-all" title="View SO Report">
                                        <ClipboardList className="w-4 h-4" />
                                      </button>
                                      <button onClick={() => setShowInvPopup(c.name)} className="p-1.5 rounded-lg bg-avail/10 text-avail hover:bg-avail hover:text-white transition-all" title="View Invoice Report">
                                        <FileText className="w-4 h-4" />
                                      </button>
                                    </div>
                                  </td>
                                </tr>
                              ))}
                              {customersList.length === 0 && (
                                 <tr>
                                    <td colSpan={6} className="py-20 text-center text-text-muted italic font-bold">
                                       No customer data located within specified filters.
                                    </td>
                                 </tr>
                              )}
                           </tbody>
                        </table>
                     </div>
                  </div>
              </motion.div>
            )}
          </AnimatePresence>
        </main>
      </div>

      {/* --- MODALS --- */}

      {/* CUSTOMER SO MODAL */}
      <AnimatePresence>
        {showSOPopup && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center p-4">
            <motion.div 
              initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }}
              className="absolute inset-0 bg-text-main/40 backdrop-blur-sm" 
              onClick={() => setShowSOPopup(null)}
            />
            <motion.div 
              initial={{ scale: 0.95, opacity: 0, y: 20 }} animate={{ scale: 1, opacity: 1, y: 0 }} exit={{ scale: 0.95, opacity: 0, y: 20 }}
              className="bg-surface w-[98vw] h-[98vh] overflow-hidden rounded-2xl shadow-2xl relative flex flex-col border border-border-custom"
            >
              <div className="p-6 border-b border-border-custom bg-surface2/30 flex justify-between items-center shrink-0">
                <div className="flex items-center gap-6">
                  <div className="w-14 h-14 bg-primary rounded-2xl flex items-center justify-center shadow-lg shadow-primary/20 shrink-0">
                    <ClipboardList className="w-7 h-7 text-white" />
                  </div>
                  <div>
                    <h2 className="text-2xl font-black text-text-main uppercase tracking-tight leading-none mb-1">{showSOPopup}</h2>
                    <div className="flex items-center gap-3">
                       <span className="text-[10px] font-black bg-surface2 px-2 py-0.5 rounded text-text-muted border border-border-custom uppercase">Group: {selectedCustomerData?.Group || 'N/A'}</span>
                       <span className={cn("text-[10px] font-black px-2 py-0.5 rounded border uppercase", 
                         selectedCustomerData?.Status === 'REPETED CUSTOMER' ? "bg-avail/10 text-avail border-avail/20" : 
                         selectedCustomerData?.Status === 'LOST CUSTOMER' ? "bg-danger/10 text-danger border-danger/20" : "bg-primary/10 text-primary border-primary/20")}>
                         {selectedCustomerData?.Status || 'ACTIVE'}
                       </span>
                    </div>
                  </div>
                  <div className="ml-10 relative">
                     <Search className="absolute left-3.5 top-1/2 -translate-y-1/2 w-4 h-4 text-text-muted" />
                     <input 
                      placeholder="Search items / Ref No..."
                      value={popupSearch}
                      onChange={e => setPopupSearch(e.target.value)}
                      className="bg-white border border-border-custom rounded-xl pl-10 pr-4 py-2 text-xs font-bold outline-none focus:border-primary w-64 shadow-sm"
                     />
                   </div>
                </div>
                <div className="flex items-center gap-4">
                   <button 
                    onClick={() => {
                      const s = cSearch.toLowerCase();
                      const exportData = processedSO
                        .filter(r => r.PartyName === showSOPopup)
                        .filter(r => activeTab !== 'customers' || !cSearch || (r.Order || "").toLowerCase().includes(s) || (r.NameOfItem || "").toLowerCase().includes(s))
                        .map(r => ({
                          "Due Date": fmtDate(r.DueOn || r.Date),
                          "Ref No": r.Order,
                          "Item Description": r.NameOfItem,
                          "Qty": r.Balance,
                          "Value": r.Value,
                          "Allocation Status": r.StockStatus
                        }));
                      exportToExcel(exportData, `${showSOPopup}_Pending_SO`);
                    }}
                    className="flex items-center gap-1.5 bg-white border border-border-custom text-text-muted px-4 py-2.5 rounded-xl text-[10px] font-black uppercase tracking-wider hover:bg-slate-50 transition-all shadow-sm"
                   >
                     <Download className="w-3.5 h-3.5" /> EXPORT EXCEL
                   </button>
                   <button 
                    onClick={() => setShowSOPopup(null)}
                    className="w-10 h-10 rounded-full hover:bg-danger/10 text-danger flex items-center justify-center transition-all border border-danger/20 bg-white shadow-sm hover:scale-110"
                   >
                    <X className="w-5 h-5" />
                   </button>
                </div>
              </div>

              <div className="flex-1 overflow-y-auto scrollbar-custom p-8 bg-slate-50/20">
                <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-8">
                   {(() => {
                      const s = cSearch.toLowerCase();
                      const items = processedSO
                         .filter(r => r.PartyName === showSOPopup)
                         .filter(r => activeTab !== 'customers' || !cSearch || (r.Order || "").toLowerCase().includes(s) || (r.NameOfItem || "").toLowerCase().includes(s));
                      const total = items.reduce((s, r) => s + r.Value, 0);
                      const dueVal = items.filter(r => r.OrderType === 'Due').reduce((s, r) => s + r.Value, 0);
                      const availVal = items.filter(r => r.StockStatus === 'Available').reduce((s, r) => s + r.Value, 0);
                      const schedVal = items.filter(r => r.OrderType === 'Schedule').reduce((s, r) => s + r.Value, 0);

                      return [
                        { label: 'Total Outstanding', val: fmtCur(total), color: 'text-text-main' },
                        { label: 'Due On-Hand', val: fmtCur(availVal), color: 'text-avail' },
                        { label: 'Due Procure', val: fmtCur(dueVal - availVal), color: 'text-danger' },
                        { label: 'Future Sched', val: fmtCur(schedVal), color: 'text-primary' },
                      ];
                   })().map((st, i) => (
                      <div key={i} className="bg-white p-5 rounded-2xl border border-border-custom shadow-sm">
                         <div className="text-[9px] font-black text-text-muted uppercase tracking-widest mb-1">{st.label}</div>
                         <div className={cn("text-xl font-black", st.color)}>{st.val}</div>
                      </div>
                   ))}
                </div>

                <div className="bg-white border border-border-custom rounded-2xl overflow-hidden shadow-sm">
                   <div className="px-6 py-4 border-b border-border-custom bg-slate-50/50">
                      <h4 className="text-[11px] font-black text-text-main uppercase tracking-widest">Specific Pending Line Items</h4>
                   </div>
                   <table className="w-full text-left text-[12px] border-separate border-spacing-0">
                      <thead className="sticky top-0 bg-white z-10 font-black uppercase text-[10px] text-text-muted">
                         <tr>
                            <th className="px-6 py-4 border-b border-border-custom text-avail">Due Date</th>
                            <th className="px-4 py-4 border-b border-border-custom">Ref No</th>
                            <th className="px-4 py-4 border-b border-border-custom">Item Description</th>
                            <th className="px-4 py-4 border-b border-border-custom text-right">Qty</th>
                            <th className="px-4 py-4 border-b border-border-custom text-right">Value</th>
                            <th className="px-6 py-4 border-b border-border-custom text-center">Allocation status</th>
                         </tr>
                      </thead>
                      <tbody className="divide-y divide-border-custom">
                         {(() => {
                            const s = cSearch.toLowerCase();
                            return processedSO
                              .filter(r => r.PartyName === showSOPopup)
                              .filter(r => activeTab !== 'customers' || !cSearch || (r.Order || "").toLowerCase().includes(s) || (r.NameOfItem || "").toLowerCase().includes(s))
                              .filter(r => !popupSearch || (r.NameOfItem || "").toLowerCase().includes(popupSearch.toLowerCase()) || (r.Order || "").toLowerCase().includes(popupSearch.toLowerCase()))
                              .sort((a,b) => new Date(a.DueOn || 0).getTime() - new Date(b.DueOn || 0).getTime())
                              .map((r, i) => (
                            <tr key={i} className="hover:bg-slate-50 transition-colors">
                               <td className="px-6 py-4 font-bold text-text-muted">{fmtDate(r.DueOn || r.Date)}</td>
                               <td className="px-4 py-4 font-mono font-black text-text-muted uppercase">#{r.Order.slice(-10)}</td>
                               <td className="px-4 py-4 font-bold text-text-main max-w-[200px] truncate">{r.NameOfItem}</td>
                               <td className="px-4 py-4 text-right font-black">{fmtNum(r.Balance)}</td>
                               <td className="px-4 py-4 text-right font-black text-primary">{fmtCur(r.Value)}</td>
                               <td className="px-6 py-4 text-center">
                                  <div className={cn(
                                    "px-2 py-1 rounded text-[9px] font-black uppercase inline-block border",
                                    r.StockStatus === 'Available' ? "bg-avail/10 text-avail border-avail/10" :
                                    r.StockStatus.includes('PO') ? "bg-primary/10 text-primary border-primary/10" : "bg-danger/10 text-danger border-danger/20"
                                  )}>
                                     {r.StockStatus}
                                  </div>
                               </td>
                            </tr>
                         ))})()}
                      </tbody>
                   </table>
                </div>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* CUSTOMER INVOICE MODAL */}
      <AnimatePresence>
        {showInvPopup && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center p-4">
            <motion.div 
              initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }}
              className="absolute inset-0 bg-text-main/40 backdrop-blur-sm" 
              onClick={() => setShowInvPopup(null)}
            />
            <motion.div 
              initial={{ scale: 0.95, opacity: 0, x: 30 }} animate={{ scale: 1, opacity: 1, x: 0 }} exit={{ scale: 0.95, opacity: 0, x: 30 }}
              className="bg-surface w-[98vw] h-[98vh] overflow-hidden rounded-2xl shadow-2xl relative flex flex-col border border-border-custom"
            >
              <div className="p-8 border-b border-border-custom flex justify-between items-center bg-surface2/30">
                <div className="flex items-center gap-5">
                   <div className="w-14 h-14 bg-avail rounded-2xl flex items-center justify-center shadow-lg shadow-avail/20 shrink-0">
                    <FileText className="w-7 h-7 text-white" />
                  </div>
                  <div>
                    <h2 className="text-2xl font-black text-text-main uppercase tracking-tight leading-none mb-1">{showInvPopup}</h2>
                    <p className="text-[10px] text-text-muted font-bold uppercase tracking-widest opacity-80">Historical Billing & Ledger Detailed View</p>
                  </div>
                  <div className="ml-10 relative">
                     <Search className="absolute left-3.5 top-1/2 -translate-y-1/2 w-4 h-4 text-text-muted" />
                     <input 
                      placeholder="Search particulars / Voucher No..."
                      value={popupSearch}
                      onChange={e => setPopupSearch(e.target.value)}
                      className="bg-white border border-border-custom rounded-xl pl-10 pr-4 py-2 text-xs font-bold outline-none focus:border-primary w-64 shadow-sm"
                     />
                   </div>
                </div>
                <div className="flex items-center gap-4">
                   <button 
                     onClick={() => {
                       const s = cSearch.toLowerCase();
                       const filteredInvoices = dynamicInvoices.filter(inv => (inv.Buyer || "").toLowerCase().trim() === showInvPopup?.toLowerCase().trim());
                       const exportData = filteredInvoices
                          .filter(inv => activeTab !== 'customers' || !cSearch || (inv.VoucherNo || "").toLowerCase().includes(s) || (inv.VoucherRef || "").toLowerCase().includes(s) || (inv.Particulars || "").toLowerCase().includes(s))
                          .map(inv => ({
                         "Date": fmtDate(inv.Date),
                         "Invoice No": inv.VoucherNo,
                         "Customer PO No": inv.VoucherRef,
                         "Item Description": inv.Particulars,
                         "Quantity": inv.Quantity,
                         "Rate": inv.Value / (inv.Quantity || 1),
                         "Total Value": inv.Value
                       }));
                       exportToExcel(exportData, `${showInvPopup}_Invoices`);
                     }}
                     className="flex items-center gap-1.5 bg-white border border-border-custom text-text-muted px-4 py-2.5 rounded-xl text-[10px] font-black uppercase tracking-wider hover:bg-slate-50 transition-all shadow-sm"
                    >
                      <Download className="w-3.5 h-3.5" /> EXPORT EXCEL
                   </button>
                   <button 
                    onClick={() => setShowInvPopup(null)}
                    className="w-10 h-10 rounded-full hover:bg-danger/10 text-danger flex items-center justify-center transition-all border border-danger/20 bg-white shadow-sm hover:scale-110"
                   >
                    <X className="w-5 h-5" />
                   </button>
                 </div>
               </div>
               <div className="flex-1 overflow-y-auto scrollbar-custom p-8 bg-slate-50/20">
                  <div className="bg-white border border-border-custom rounded-2xl overflow-hidden shadow-sm">
                    <table className="w-full text-left text-[11px] border-separate border-spacing-0">
                       <thead className="bg-slate-50/50 sticky top-0 z-10 font-black text-text-muted uppercase text-[10px] border-b border-border-custom">
                          <tr>
                             <th className="px-6 py-4 border-b border-border-custom">Date</th>
                             <th className="px-4 py-4 border-b border-border-custom">Invoice No</th>
                             <th className="px-4 py-4 border-b border-border-custom">Customer PO No</th>
                             <th className="px-4 py-4 border-b border-border-custom">Description</th>
                             <th className="px-4 py-4 border-b border-border-custom text-right">Qty</th>
                             <th className="px-4 py-4 border-b border-border-custom text-right">Rate</th>
                             <th className="px-6 py-4 border-b border-border-custom text-right">Value</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {(() => {
                            const s = cSearch.toLowerCase();
                            const flatItems = dynamicInvoices
                              .filter(inv => (inv.Buyer || "").toLowerCase().trim() === showInvPopup?.toLowerCase().trim())
                              .filter(inv => activeTab !== 'customers' || !cSearch || (inv.VoucherNo || "").toLowerCase().includes(s) || (inv.VoucherRef || "").toLowerCase().includes(s) || (inv.Particulars || "").toLowerCase().includes(s))
                              .filter(inv => !popupSearch || (inv.Particulars || "").toLowerCase().includes(popupSearch.toLowerCase()) || (inv.VoucherNo || "").toLowerCase().includes(popupSearch.toLowerCase()) || (inv.VoucherRef || "").toLowerCase().includes(popupSearch.toLowerCase()));

                            if (flatItems.length === 0) {
                              return (
                                <tr>
                                  <td colSpan={7} className="py-20 text-center opacity-40 italic font-bold text-sm">No verified invoicing records located for this specific buyer profile.</td>
                                </tr>
                              );
                            }

                            const totalQty = flatItems.reduce((s, it) => s + (it.Quantity || 0), 0);
                            const totalVal = flatItems.reduce((s, it) => s + (it.Value || 0), 0);

                            return (
                              <>
                                <tr className="bg-slate-50/80 border-b-2 border-slate-200">
                                   <td colSpan={4} className="px-6 py-3 text-[10px] font-black text-text-muted uppercase tracking-widest text-right">Filtered Totals:</td>
                                   <td className="px-4 py-3 text-right font-black text-text-main text-sm">{fmtNum(totalQty)}</td>
                                   <td className="px-4 py-3"></td>
                                   <td className="px-6 py-3 text-right font-black text-primary text-sm">{fmtCur(totalVal)}</td>
                                </tr>
                                {flatItems.map((it, idx) => (
                                  <tr key={idx} className="hover:bg-slate-50/50 transition-colors">
                                    <td className="px-6 py-4 font-bold text-text-muted whitespace-nowrap">{fmtDate(it.Date)}</td>
                                    <td className="px-4 py-4 font-mono font-black text-text-muted uppercase whitespace-nowrap">#{it.VoucherNo}</td>
                                    <td className="px-4 py-4 font-bold text-text-muted uppercase whitespace-nowrap text-[10px]">{it.VoucherRef}</td>
                                    <td className="px-4 py-4 font-bold text-text-main max-w-[300px] truncate">{it.Particulars}</td>
                                    <td className="px-4 py-4 text-right font-black">{it.Quantity}</td>
                                    <td className="px-4 py-4 text-right font-bold text-text-muted">{fmtCur(it.Value / (it.Quantity || 1))}</td>
                                    <td className="px-6 py-4 text-right font-black text-avail">{fmtCur(it.Value)}</td>
                                  </tr>
                                ))}
                              </>
                            );
                          })()}
                        </tbody>
                    </table>
                  </div>
               </div>

            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}

export default function App() {
  return (
    <ErrorBoundary>
      <MainApp />
    </ErrorBoundary>
  );
}
