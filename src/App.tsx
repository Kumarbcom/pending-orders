import React, { useState, useMemo, useRef, ChangeEvent } from 'react';
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
  Menu
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
  Bar as ReBar 
} from 'recharts';
import { motion, AnimatePresence } from 'motion/react';
import { ALL_SO as INITIAL_SO, ALL_INVOICE, META } from './data';
import { SalesOrder, Invoice, PurchaseOrder, StockItem, MaterialMasterItem, CustomerMasterItem } from './types';
import { cn } from './lib/utils';

// --- Utils ---
const fmtCur = (v: number) => {
  if (v == null || isNaN(v)) return '—';
  const abs = Math.abs(v);
  if (abs >= 1e7) return '₹' + (v / 1e7).toFixed(2) + ' Cr';
  if (abs >= 1e5) return '₹' + (v / 1e5).toFixed(2) + ' L';
  return '₹' + v.toLocaleString('en-IN', { maximumFractionDigits: 0 });
};

const fmtNum = (v: number) => {
  if (v == null) return '—';
  return v.toLocaleString('en-IN', { maximumFractionDigits: 2 });
};

const fmtDate = (d: string | null) => {
  if (!d) return '—';
  const parts = d.split('-');
  if (parts.length !== 3) return d;
  return `${parts[2]}/${parts[1]}/${parts[0]}`;
};

// --- Components ---

const StatCard = ({ title, value, subValue, type, details }: any) => (
  <div className="bg-surface border border-border-custom rounded-2xl p-6 shadow-sm flex flex-col relative transition-all hover:shadow-md">
    <div className="flex justify-between items-start mb-3">
      <span className="text-[11px] font-bold uppercase tracking-wider text-text-muted">
        {title}
      </span>
      {type === 'due' && <AlertCircle className="w-4 h-4 text-due" />}
      {type === 'sched' && <TrendingUp className="w-4 h-4 text-sched" />}
    </div>
    <div className="text-3xl font-bold text-text-main mb-1 tracking-tight">
      {value}
    </div>
    {subValue && (
      <div className="text-xs text-text-muted font-medium flex items-center gap-1.5 mt-1">
        {subValue}
      </div>
    )}
    
    {details && (
      <div className="flex gap-4 mt-5 pt-5 border-t border-border-custom">
        {details.map((d: any, i: number) => (
          <div key={i} className="flex-1">
            <div className="text-[9px] font-bold text-text-muted uppercase mb-1 tracking-wide">{d.label}</div>
            <div className={cn("font-bold text-[13px] leading-none", d.color)}>{d.value}</div>
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

export default function App() {
  const [activeTab, setActiveTab] = useState<'dashboard' | 'pending-so' | 'pending-po' | 'stock' | 'material-master' | 'customer-master' | 'invoices' | 'customers'>('dashboard');
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
  
  // Data States
  const [dynamicSO, setDynamicSO] = useState<SalesOrder[]>(INITIAL_SO);
  const [dynamicPO, setDynamicPO] = useState<PurchaseOrder[]>([]);
  const [dynamicStock, setDynamicStock] = useState<StockItem[]>([]);
  const [dynamicMaterialMaster, setDynamicMaterialMaster] = useState<MaterialMasterItem[]>([]);
  const [dynamicCustomerMaster, setDynamicCustomerMaster] = useState<CustomerMasterItem[]>([]);
  const [dynamicInvoices, setDynamicInvoices] = useState<Invoice[]>([]);

  // Sorting State
  const [sortField, setSortField] = useState<string | null>(null);
  const [sortDirection, setSortDirection] = useState<'asc' | 'desc'>('asc');

  const handleSort = (field: string) => {
    if (sortField === field) {
      setSortDirection(sortDirection === 'asc' ? 'desc' : 'asc');
    } else {
      setSortField(field);
      setSortDirection('asc');
    }
  };

  const handleReset = (type: 'so' | 'po' | 'stock' | 'material' | 'customer' | 'invoice' | 'all') => {
    if (type === 'so' || type === 'all') setDynamicSO(INITIAL_SO);
    if (type === 'po' || type === 'all') setDynamicPO([]);
    if (type === 'stock' || type === 'all') setDynamicStock([]);
    if (type === 'material' || type === 'all') setDynamicMaterialMaster([]);
    if (type === 'customer' || type === 'all') setDynamicCustomerMaster([]);
    if (type === 'invoice' || type === 'all') setDynamicInvoices([]);
  };

  const fileInputRef = useRef<HTMLInputElement>(null);
  const fileInputPORef = useRef<HTMLInputElement>(null);
  const fileInputStockRef = useRef<HTMLInputElement>(null);
  const fileInputMaterialRef = useRef<HTMLInputElement>(null);
  const fileInputCustomerRef = useRef<HTMLInputElement>(null);
  const fileInputInvoiceRef = useRef<HTMLInputElement>(null);

  const handleFileUpload = (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws);

      // Clean and map data based on the screenshot provided
      const parsed: SalesOrder[] = data.map((row: any) => {
        // Simple numeric extraction for fields like "105 ST"
        const extractNum = (val: any) => {
          if (typeof val === 'number') return val;
          if (!val) return 0;
          const match = String(val).match(/[\d.]+/);
          return match ? parseFloat(match[0]) : 0;
        };

        return {
          Date: row['Due'] || new Date().toISOString().split('T')[0],
          Order: row['Part No'] || 'UPLOADED',
          PartyName: row['End-Customer'] || row['Customer'] || 'Imported Client',
          NameOfItem: row['Name of Item'] || row['Description'] || 'Imported Item',
          MaterialCode: row['Material Code'] || '',
          PartNo: row['Part No'] || '',
          Ordered: extractNum(row['Order']),
          Balance: extractNum(row['Balar']) || extractNum(row['Balance']),
          Rate: extractNum(row['R']) || extractNum(row['Rate']),
          Discount: extractNum(row['Discou']) || 0,
          Value: extractNum(row['Value']),
          DueOn: row['Due'] || null,
          DueSerial: null,
          Make: 'IMPORTED',
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
      });

      if (parsed.length > 0) {
        setDynamicSO(parsed);
      }
    };
    reader.readAsBinaryString(file);
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
      const data = XLSX.utils.sheet_to_json(ws);

      const parsed: PurchaseOrder[] = data.map((row: any) => {
        const extractNum = (val: any) => {
          if (typeof val === 'number') return val;
          if (!val) return 0;
          const match = String(val).match(/[\d.]+/);
          return match ? parseFloat(match[0]) : 0;
        };

        return {
          Date: row['Date'] || '',
          Order: row['Order'] || '',
          PartyName: row["Party's Name"] || row['Party Name'] || '',
          NameOfItem: row['Name of Item'] || '',
          MaterialCode: row['Material Code'] || '',
          PartNo: row['Part No'] || '',
          Ordered: extractNum(row['Ordered']),
          Balance: extractNum(row['Balance']),
          Rate: extractNum(row['Rate']),
          Discount: extractNum(row['Discount']),
          Value: extractNum(row['Value']),
          DueOn: row['Due on'] || row['Due'] || null
        };
      });

      if (parsed.length > 0) {
        setDynamicPO(parsed);
      }
    };
    reader.readAsBinaryString(file);
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
      const data: any[] = XLSX.utils.sheet_to_json(ws);

      const invoices: Invoice[] = [];
      let currentInvoice: Invoice | null = null;

      const extractNum = (val: any) => {
        if (typeof val === 'number') return val;
        if (!val) return 0;
        const match = String(val).match(/[\d.]+/);
        return match ? parseFloat(match[0]) : 0;
      };

      data.forEach((row) => {
        const date = row['Date'];
        const buyer = row['Buyer'];
        
        if (date && buyer) {
          // Start of a new invoice
          currentInvoice = {
            Date: String(date),
            Buyer: String(buyer),
            Consignee: String(row['Consignee'] || buyer),
            VoucherNo: String(row['Voucher No.'] || ''),
            VoucherRef: String(row['Voucher Ref. No.'] || ''),
            Quantity: extractNum(row['Quantity']),
            Value: extractNum(row['Value']),
            Items: []
          };
          invoices.push(currentInvoice);
        } else if (currentInvoice && row['Particulars']) {
          // Add item to current invoice
          currentInvoice.Items.push({
            Particulars: String(row['Particulars']),
            Quantity: extractNum(row['Quantity']),
            Value: extractNum(row['Value'])
          });
        }
      });

      if (invoices.length > 0) {
        setDynamicInvoices(invoices);
      }
    };
    reader.readAsBinaryString(file);
  };
  
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

  // Pending PO Slicers
  const [poSearch, setPoSearch] = useState('');

  // Stock Slicers
  const [stockSearch, setStockSearch] = useState('');

  // Material Master Slicers
  const [materialSearch, setMaterialSearch] = useState('');

  // Customer Master Slicers
  const [customerMasterSearch, setCustomerMasterSearch] = useState('');

  // Invoice Master Slicers
  const [invoiceSearch, setInvoiceSearch] = useState('');

  // DETAILS POPUP STATE
  const [showSOPopup, setShowSOPopup] = useState<string | null>(null);
  const [showInvPopup, setShowInvPopup] = useState<string | null>(null);

  const selectedCustomerData = useMemo(() => {
    if (!showSOPopup && !showInvPopup) return null;
    const name = showSOPopup || showInvPopup;
    return dynamicCustomerMaster.find(c => c.CustomerName === name) || null;
  }, [showSOPopup, showInvPopup, dynamicCustomerMaster]);

  // FIFO PROCESSING LOGIC
  const processedSO = useMemo(() => {
    const CUTOFF_DATE = new Date('2026-04-30');
    
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
      const dbA = a.DueOn ? new Date(a.DueOn).getTime() : 0;
      const dbB = b.DueOn ? new Date(b.DueOn).getTime() : 0;
      return dbA - dbB;
    });

    return sortedRaw.map(so => {
      const dueDate = so.DueOn ? new Date(so.DueOn) : null;
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
      return true;
    });
  }, [processedSO, dMake, dGroup, dCGroup, dOrderType]);

  const dashboardStats = useMemo(() => {
    const total = filteredDashboardSO.reduce((s, r) => s + r.Value, 0);
    const due = filteredDashboardSO.filter(r => r.OrderType === 'Due');
    const sched = filteredDashboardSO.filter(r => r.OrderType === 'Schedule');
    
    const dueVal = due.reduce((s, r) => s + r.Value, 0);
    const dueAvail = due.filter(r => r.StockStatus === 'Available').reduce((s, r) => s + r.Value, 0);
    const dueArr = dueVal - dueAvail;

    const schedVal = sched.reduce((s, r) => s + r.Value, 0);
    const schedAvail = sched.filter(r => r.StockStatus === 'Available').reduce((s, r) => s + r.Value, 0);
    const schedArr = schedVal - schedAvail;

    return { total, count: filteredDashboardSO.length, dueVal, dueAvail, dueArr, schedVal, schedAvail, schedArr };
  }, [filteredDashboardSO]);

  const filteredSO = useMemo(() => {
    let list = processedSO.filter(r => {
      if (soType && r.OrderType !== soType) return false;
      if (soMake && r.Make !== soMake) return false;
      if (soGroup && r.Group !== soGroup) return false;
      if (soCust && r.PartyName !== soCust) return false;
      if (soStatus && r.StockStatus !== soStatus) return false;
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
    processedSO.forEach(r => {
      if (cGroup && r.Group !== cGroup) return;
      if (cCGroup && r.CustomerGroup !== cCGroup) return;
      if (cCust && r.PartyName !== cCust) return;
      if (cSearch && !r.PartyName.toLowerCase().includes(cSearch.toLowerCase())) return;
      
      const k = r.PartyName;
      if (!custMap[k]) {
        custMap[k] = { name: k, group: r.Group, cgroup: r.CustomerGroup, dueVal: 0, schedVal: 0, total: 0 };
      }
      const c = custMap[k];
      c.total += r.Value;
      if (r.OrderType === 'Due') c.dueVal += r.Value; else c.schedVal += r.Value;
    });

    const invTotals: Record<string, { val: number; count: number }> = {};
    dynamicInvoices.forEach(i => {
      const k = i.Buyer.toLowerCase().trim();
      if (!invTotals[k]) invTotals[k] = { val: 0, count: 0 };
      invTotals[k].val += i.Value;
      invTotals[k].count += 1;
    });

    return Object.values(custMap)
      .map(c => ({
        ...c,
        invVal: invTotals[c.name.toLowerCase().trim()]?.val || 0,
        invCount: invTotals[c.name.toLowerCase().trim()]?.count || 0,
      }))
      .sort((a, b) => b.total - a.total);
  }, [processedSO, dynamicInvoices, cGroup, cCGroup, cCust, cSearch]);

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
      .sort((a, b) => b.total - a.total)
      .slice(0, 10);

    return { pendingTypeData, makeStackedData, top10Detailed };
  }, [filteredDashboardSO, dashboardStats]);

  const filteredPOList = useMemo(() => {
    let list = dynamicPO.filter(p => {
      if (poSearch && !p.PartyName.toLowerCase().includes(poSearch.toLowerCase()) && !p.NameOfItem.toLowerCase().includes(poSearch.toLowerCase())) return false;
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
  }, [dynamicPO, poSearch, sortField, sortDirection]);

  const filteredStockList = useMemo(() => {
    let list = dynamicStock.filter(s => {
      if (stockSearch && !s.Particulars.toLowerCase().includes(stockSearch.toLowerCase())) return false;
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
  }, [dynamicStock, stockSearch, sortField, sortDirection]);

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
      if (customerMasterSearch && !c.CustomerName.toLowerCase().includes(customerMasterSearch.toLowerCase())) return false;
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
      if (invoiceSearch && !i.Buyer.toLowerCase().includes(invoiceSearch.toLowerCase()) && !i.VoucherNo.toLowerCase().includes(invoiceSearch.toLowerCase())) return false;
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
              <div className="w-8 h-8 rounded-lg bg-primary flex items-center justify-center shadow-lg shadow-primary/20">
                <Package className="w-5 h-5 text-white" />
              </div>
              <h1 className="text-lg font-bold tracking-tight text-text-main uppercase">Portfolio</h1>
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
            { id: 'customers', label: 'Dashboard Analysis', icon: TrendingUp },
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
                   <div className="w-2 h-2 rounded-full bg-avail animate-pulse" />
                   <span className="text-xs font-bold font-mono">LIVE_FEED_01</span>
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
              className="bg-surface2 rounded-full pl-11 pr-6 py-2.5 text-sm w-[380px] border border-transparent focus:border-primary focus:bg-surface outline-none transition-all"
            />
          </div>

          <div className="flex items-center gap-6">
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

                  <button 
                    onClick={() => { setDMake(''); setDGroup(''); setDOrderType(''); }}
                    className="ml-auto bg-surface2 hover:bg-border-custom/50 p-2.5 rounded-xl transition-colors text-text-muted"
                  >
                    <Search className="w-4 h-4 rotate-45" />
                  </button>
                </div>

                {/* KPI ROW */}
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                  <StatCard 
                    title="Total Pending SO" 
                    value={fmtCur(dashboardStats.total)} 
                    subValue={<><ClipboardList className="w-3.5 h-3.5" /> {dashboardStats.count} Total Orders</>}
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
                               <Bar dataKey="value" radius={[0, 6, 6, 0]} barSize={40} />
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
                               <Bar dataKey="due" name="Due Value" fill="#ff4d4f" stackId="make" />
                               <Bar dataKey="schedule" name="Schedule Value" fill="#1890ff" stackId="make" />
                            </BarChart>
                         </ResponsiveContainer>
                      </div>
                   </div>
                </div>

                {/* RECENT TOP CLIENTS */}
                <div className="bg-surface border border-border-custom shadow-sm overflow-hidden">
                  <div className="px-4 py-3 border-b border-border-custom bg-surface2/30 flex justify-between items-center text-[Cambria]">
                    <h3 className="text-[12px] font-black text-text-main uppercase tracking-tight">Top 10 Customers Pending SO Breakdown</h3>
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
                        {dashboardChartsData.top10Detailed.map((c, i) => (
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

                    <div className="flex flex-col gap-1.5 flex-1 min-w-[240px]">
                      <label className="text-[10px] font-bold text-text-muted uppercase px-1">Customer Search</label>
                      <div className="relative">
                        <Search className="absolute left-3.5 top-2.5 w-4 h-4 text-text-muted" />
                        <input 
                          placeholder="Search customer name..." 
                          value={soCust} 
                          onChange={e => setSoCust(e.target.value)}
                          className="pl-10 w-full bg-surface2 border border-border-custom rounded-xl px-4 py-2.5 text-xs font-semibold focus:border-primary outline-none"
                        />
                      </div>
                    </div>

                    <div className="flex items-center gap-3">
                      <button 
                        onClick={() => fileInputRef.current?.click()}
                        className="flex items-center gap-2 bg-avail text-white px-5 py-2.5 rounded-xl text-xs font-bold shadow-lg hover:shadow-primary/20 transition-all hover:bg-avail/90"
                      >
                        <Upload className="w-4 h-4" /> UPLOAD
                      </button>
                      <button 
                        onClick={() => { if(confirm('Reset SO data?')) handleReset('so') }}
                        className="flex items-center gap-2 bg-white border border-border-custom text-text-muted px-5 py-2.5 rounded-xl text-xs font-bold hover:bg-surface2 transition-all shadow-sm"
                      >
                        <RefreshCw className="w-4 h-4" /> RESET
                      </button>
                      <button className="flex items-center gap-2 bg-text-main text-white px-5 py-2.5 rounded-xl text-xs font-bold shadow-lg hover:shadow-primary/20 transition-all hover:bg-primary">
                        <Download className="w-4 h-4" /> EXTRACT
                      </button>
                    </div>
                    <input type="file" ref={fileInputRef} onChange={handleFileUpload} className="hidden" accept=".xlsx, .xls, .csv" />
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
                              {filteredSO.map((r, idx) => (
                                <tr key={idx} className="hover:bg-slate-50 transition-colors group">
                                  <td className={cn("font-bold whitespace-nowrap", r.OrderType === 'Due' ? "text-danger" : "text-primary")}>
                                     {fmtDate(r.DueOn || r.Date)}
                                  </td>
                                  <td className="font-mono text-text-muted uppercase">#{r.Order.slice(-8)}</td>
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
                  <div className="bg-surface border border-border-custom shadow-sm p-6 flex flex-wrap items-end gap-5">
                    <div className="flex flex-col gap-1.5 flex-1 min-w-[240px]">
                      <label className="text-[10px] font-bold text-text-muted uppercase px-1">PO Analytical Search</label>
                      <div className="relative">
                        <Search className="absolute left-3.5 top-2.5 w-4 h-4 text-text-muted" />
                        <input 
                          placeholder="Search Party or Item..." 
                          value={poSearch} 
                          onChange={e => setPoSearch(e.target.value)}
                          className="pl-10 w-full bg-surface2 border border-border-custom rounded-xl px-4 py-2.5 text-xs font-semibold focus:border-primary outline-none focus:bg-white transition-all"
                        />
                      </div>
                    </div>

                    <div className="flex items-center gap-3">
                      <button 
                        onClick={() => fileInputPORef.current?.click()}
                        className="flex items-center gap-2 bg-avail text-white px-5 py-2.5 rounded-xl text-xs font-bold shadow-lg hover:shadow-primary/20 transition-all hover:bg-avail/90"
                      >
                        <Upload className="w-4 h-4" /> UPLOAD PO
                      </button>
                      <button 
                         onClick={() => { if(confirm('Reset PO data?')) handleReset('po') }}
                         className="flex items-center gap-2 bg-white border border-border-custom text-text-muted px-5 py-2.5 rounded-xl text-xs font-bold hover:bg-slate-50 transition-all shadow-sm"
                      >
                         <RefreshCw className="w-4 h-4" /> RESET
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
                              {filteredPOList.map((p, idx) => (
                                <tr key={idx} className="hover:bg-slate-50 transition-colors group">
                                  <td className="whitespace-nowrap">{fmtDate(p.Date)}</td>
                                  <td className="font-mono text-text-muted uppercase">#{p.Order.slice(-8)}</td>
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

                    <div className="flex items-center gap-3">
                      <button 
                        onClick={() => fileInputStockRef.current?.click()}
                        className="flex items-center gap-2 bg-avail text-white px-5 py-2.5 rounded-xl text-xs font-bold shadow-lg hover:shadow-primary/20 transition-all hover:bg-avail/90"
                      >
                        <Upload className="w-4 h-4" /> UPLOAD STOCK
                      </button>
                      <button 
                         onClick={() => { if(confirm('Reset Inventory data?')) handleReset('stock') }}
                         className="flex items-center gap-2 bg-white border border-border-custom text-text-muted px-5 py-2.5 rounded-xl text-xs font-bold hover:bg-slate-50 transition-all shadow-sm"
                      >
                         <RefreshCw className="w-4 h-4" /> RESET
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
                              {filteredStockList.map((s, idx) => (
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

                    <div className="flex items-center gap-3">
                      <button 
                        onClick={() => fileInputMaterialRef.current?.click()}
                        className="flex items-center gap-2 bg-avail text-white px-5 py-2.5 rounded-xl text-xs font-bold shadow-lg hover:shadow-primary/20 transition-all hover:bg-avail/90"
                      >
                        <Upload className="w-4 h-4" /> UPLOAD MASTER
                      </button>
                      <button 
                         onClick={() => { if(confirm('Reset Material Master?')) handleReset('material') }}
                         className="flex items-center gap-2 bg-white border border-border-custom text-text-muted px-5 py-2.5 rounded-xl text-xs font-bold hover:bg-slate-50 transition-all shadow-sm"
                      >
                         <RefreshCw className="w-4 h-4" /> RESET
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
                              {filteredMaterialList.map((m, idx) => (
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

                    <div className="flex items-center gap-3">
                      <button 
                        onClick={() => fileInputCustomerRef.current?.click()}
                        className="flex items-center gap-2 bg-avail text-white px-5 py-2.5 rounded-xl text-xs font-bold shadow-lg hover:shadow-primary/20 transition-all hover:bg-avail/90"
                      >
                        <Upload className="w-4 h-4" /> UPLOAD CSV
                      </button>
                      <button 
                         onClick={() => { if(confirm('Reset Customer Master?')) handleReset('customer') }}
                         className="flex items-center gap-2 bg-white border border-border-custom text-text-muted px-5 py-2.5 rounded-xl text-xs font-bold hover:bg-slate-50 transition-all shadow-sm"
                      >
                         <RefreshCw className="w-4 h-4" /> RESET
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
                              {filteredCustomerMasterList.map((c, idx) => (
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

                    <button 
                      onClick={() => fileInputInvoiceRef.current?.click()}
                      className="flex items-center gap-2 bg-avail text-white px-5 py-2.5 rounded-xl text-xs font-bold shadow-lg hover:shadow-primary/20 transition-all hover:bg-avail/90"
                    >
                      <Upload className="w-4 h-4" /> UPLOAD INVOICE EXCEL
                    </button>
                    <input 
                      type="file" 
                      ref={fileInputInvoiceRef} 
                      onChange={handleInvoiceUpload} 
                      className="hidden" 
                      accept=".xlsx, .xls, .csv" 
                    />

                    <button className="flex items-center gap-2 bg-text-main text-white px-5 py-2.5 rounded-xl text-xs font-bold shadow-lg hover:shadow-primary/20 transition-all hover:bg-primary">
                      <Download className="w-4 h-4" /> EXPORT
                    </button>
                  </div>

                  <div className="bg-surface border border-border-custom rounded-2xl overflow-hidden shadow-sm">
                     <div className="overflow-x-auto scrollbar-custom max-h-[calc(100vh-320px)]">
                        <table className="w-full text-[12px] border-separate border-spacing-0">
                           <thead className="sticky top-0 bg-white z-10 font-bold uppercase tracking-wider text-text-muted border-b border-border-custom">
                              <tr>
                                 <th className="px-6 py-4 bg-white border-b border-border-custom text-left">Date & Voucher</th>
                                 <th className="px-4 py-4 bg-white border-b border-border-custom text-left">Buyer / Party Details</th>
                                 <th className="px-4 py-4 bg-white border-b border-border-custom text-right">Qty</th>
                                 <th className="px-6 py-4 bg-white border-b border-border-custom text-right text-primary font-black">Total Value</th>
                              </tr>
                           </thead>
                           <tbody className="bg-white">
                              {filteredInvoiceList.map((inv, idx) => (
                                <React.Fragment key={idx}>
                                  <tr className="bg-slate-50/50 hover:bg-slate-100/50 transition-colors">
                                    <td className="px-6 py-4">
                                      <div className="font-bold text-text-main">{inv.Date}</div>
                                      <div className="text-[10px] text-primary font-black uppercase tracking-tighter">{inv.VoucherNo}</div>
                                    </td>
                                    <td className="px-4 py-4">
                                      <div className="font-black text-text-main uppercase text-[13px]">{inv.Buyer}</div>
                                      <div className="text-[10px] text-text-muted font-mono truncate max-w-sm">{inv.VoucherRef}</div>
                                    </td>
                                    <td className="px-4 py-4 text-right font-black text-text-main">{fmtNum(inv.Quantity)}</td>
                                    <td className="px-6 py-4 text-right font-black text-primary text-sm">{fmtCur(inv.Value)}</td>
                                  </tr>
                                  {inv.Items.map((item, idy) => (
                                    <tr key={`${idx}-${idy}`} className="group border-b border-border-custom/30 last:border-b-0">
                                      <td className="px-6 py-2.5"></td>
                                      <td className="px-4 py-2.5 text-text-muted text-[11px] font-medium pl-10 border-l-2 border-primary/20 bg-slate-50/20">
                                        {item.Particulars}
                                      </td>
                                      <td className="px-4 py-2.5 text-right text-text-muted text-[11px] opacity-70">
                                        {fmtNum(item.Quantity)}
                                      </td>
                                      <td className="px-6 py-2.5 text-right text-text-muted text-[11px] italic">
                                        {fmtCur(item.Value)}
                                      </td>
                                    </tr>
                                  ))}
                                </React.Fragment>
                              ))}
                              {filteredInvoiceList.length === 0 && (
                                 <tr>
                                    <td colSpan={4} className="py-20 text-center text-text-muted italic font-bold">
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
                     <div className="flex gap-4 flex-1 max-w-2xl">
                        <div className="relative flex-1">
                          <Search className="absolute left-4 top-1/2 -translate-y-1/2 w-4 h-4 text-text-muted" />
                          <input 
                            placeholder="Find specific customer by name..." 
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
                          <option value="">Global Regions</option>
                          {META.groups.map(g => <option key={g} value={g}>{g}</option>)}
                        </select>
                     </div>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 2xl:grid-cols-4 gap-6">
                    {customersList.map(c => (
                      <div 
                        key={c.name}
                        onClick={() => setSelectedCustomer(c.name)}
                        className="bg-surface border border-border-custom rounded-3xl p-8 flex flex-col justify-between hover:shadow-2xl hover:shadow-slate-200 transition-all duration-300 transform hover:-translate-y-1 group relative overflow-hidden"
                      >
                         <div className="absolute -right-4 -bottom-4 w-32 h-32 bg-slate-50 rounded-full group-hover:scale-150 transition-transform duration-500 -z-10" />
                         
                         <div className="space-y-2">
                           <div className="w-10 h-10 rounded-2xl bg-surface2 flex items-center justify-center mb-4 group-hover:bg-primary group-hover:text-white transition-colors">
                              <Users className="w-5 h-5 text-text-muted group-hover:text-white" />
                           </div>
                           <h4 className="text-base font-black text-text-main leading-tight group-hover:text-primary transition-colors min-h-[48px] line-clamp-2 uppercase tracking-tight">{c.name}</h4>
                           <div className="flex items-center gap-2">
                              <div className="w-1.5 h-1.5 rounded-full bg-primary" />
                              <span className="text-[10px] font-black uppercase text-text-muted tracking-widest">{c.cgroup || c.group}</span>
                           </div>
                         </div>

                         <div className="mt-8 pt-8 border-t border-border-custom/50 space-y-4">
                            <div>
                               <div className="text-[10px] font-black text-text-muted uppercase tracking-wider mb-1 flex justify-between">
                                 Portfolio Value
                                 <span className="text-text-main font-mono">{fmtCur(c.total)}</span>
                               </div>
                               <div className="w-full h-1.5 bg-surface2 rounded-full overflow-hidden flex">
                                  <div className="h-full bg-due" style={{ width: `${(c.dueVal/c.total)*100}%` }} />
                                  <div className="h-full bg-sched" style={{ width: `${(c.schedVal/c.total)*100}%` }} />
                               </div>
                            </div>
                            
                            <div className="flex justify-between items-end">
                               <div 
                                onClick={(e) => {
                                  if (c.invVal > 0) {
                                    e.stopPropagation();
                                    setSelectedInvoiceCust(c.name);
                                  }
                                }}
                                className="text-xs font-bold text-primary hover:underline cursor-pointer flex items-center gap-1"
                               >
                                  {c.invCount} Records <ArrowUpRight className="w-3 h-3" />
                               </div>
                               <div className="text-right">
                                  <div className="text-[10px] font-bold text-text-muted uppercase leading-none mb-1">Invoiced</div>
                                  <div className="text-md font-black text-text-main leading-none">{c.invVal > 0 ? fmtCur(c.invVal) : '₹0.00'}</div>
                               </div>
                            </div>
                         </div>
                      </div>
                    ))}
                  </div>

                  {customersList.length === 0 && (
                     <div className="bg-surface2/30 rounded-3xl border-2 border-dashed border-border-custom py-32 text-center text-text-muted font-bold italic tracking-wider">
                        No clients located within specified search parameters.
                     </div>
                  )}
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
              className="bg-surface w-full max-w-6xl max-h-[90vh] overflow-hidden rounded-[32px] shadow-2xl relative flex flex-col border border-border-custom"
            >
              <div className="p-6 border-b border-border-custom bg-surface2/30 flex justify-between items-center shrink-0">
                <div className="flex items-center gap-6">
                  <div className="w-14 h-14 bg-primary rounded-2xl flex items-center justify-center shadow-lg shadow-primary/20">
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
                </div>
                <button 
                  onClick={() => setShowSOPopup(null)}
                  className="w-10 h-10 rounded-full hover:bg-slate-200 flex items-center justify-center transition-colors border border-border-custom bg-white shadow-sm"
                >
                  <Search className="w-5 h-5 text-text-muted rotate-45" />
                </button>
              </div>

              <div className="flex-1 overflow-y-auto scrollbar-custom p-8 bg-slate-50/20">
                <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-8">
                   {(() => {
                      const items = processedSO.filter(r => r.PartyName === showSOPopup);
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
                            <th className="px-6 py-4 border-b border-border-custom">Date</th>
                            <th className="px-4 py-4 border-b border-border-custom">Ref No</th>
                            <th className="px-4 py-4 border-b border-border-custom">Item Description</th>
                            <th className="px-4 py-4 border-b border-border-custom text-right">Qty</th>
                            <th className="px-4 py-4 border-b border-border-custom text-right">Value</th>
                            <th className="px-6 py-4 border-b border-border-custom text-center">Allocation status</th>
                         </tr>
                      </thead>
                      <tbody className="divide-y divide-border-custom">
                         {processedSO.filter(r => r.PartyName === showSOPopup).sort((a,b) => new Date(a.DueOn || 0).getTime() - new Date(b.DueOn || 0).getTime()).map((r, i) => (
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
                         ))}
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
              className="bg-surface w-full max-w-4xl max-h-[85vh] overflow-hidden rounded-[32px] shadow-2xl relative flex flex-col border border-border-custom"
            >
              <div className="p-8 border-b border-border-custom flex justify-between items-center bg-surface2/30">
                <div className="flex items-center gap-5">
                   <div className="w-12 h-12 bg-white border border-border-custom rounded-2xl flex items-center justify-center shadow-sm">
                      <FileText className="w-6 h-6 text-primary" />
                   </div>
                   <div>
                    <h2 className="text-xl font-black text-text-main uppercase tracking-tight leading-none mb-1">{showInvPopup}</h2>
                    <p className="text-[10px] text-text-muted font-bold uppercase tracking-widest opacity-80">Historical Billing & Ledger</p>
                   </div>
                </div>
                <button 
                  onClick={() => setShowInvPopup(null)}
                  className="w-10 h-10 rounded-full hover:bg-slate-200 flex items-center justify-center transition-colors border border-border-custom bg-white shadow-sm"
                >
                  <Search className="w-5 h-5 text-text-muted rotate-45" />
                </button>
              </div>

              <div className="flex-1 overflow-y-auto scrollbar-custom p-8 space-y-6 bg-slate-50/20">
                 {dynamicInvoices.filter(inv => inv.Buyer.toLowerCase().trim() === showInvPopup?.toLowerCase().trim()).map((inv, idx) => (
                    <div key={idx} className="bg-white border border-border-custom rounded-2xl overflow-hidden shadow-sm hover:shadow-md transition-all">
                       <div className="p-5 flex justify-between items-center bg-slate-50/50 border-b border-border-custom">
                          <div className="flex gap-8 items-center">
                             <div>
                               <div className="text-[9px] font-black text-text-muted uppercase tracking-widest mb-0.5 opacity-70">Invoice ID</div>
                               <div className="text-sm font-black font-mono text-text-main">{inv.VoucherNo}</div>
                             </div>
                             <div>
                               <div className="text-[9px] font-black text-text-muted uppercase tracking-widest mb-0.5 opacity-70">Date</div>
                               <div className="text-sm font-bold text-text-main">{fmtDate(inv.Date)}</div>
                             </div>
                          </div>
                          <div className="text-right">
                            <div className="text-[10px] font-black text-text-muted uppercase tracking-widest mb-0.5 opacity-70">Total Value</div>
                            <div className="text-lg font-black text-primary">{fmtCur(inv.Value)}</div>
                          </div>
                       </div>
                       <div className="overflow-x-auto">
                          <table className="w-full text-left text-[11px] border-separate border-spacing-0">
                             <thead className="bg-slate-50/30 font-black text-text-muted uppercase text-[10px] border-b border-border-custom">
                                <tr>
                                   <th className="px-6 py-3">Description</th>
                                   <th className="px-4 py-3 text-right">Quantity</th>
                                   <th className="px-4 py-3 text-right">Rate</th>
                                   <th className="px-6 py-3 text-right">Line Total</th>
                                </tr>
                             </thead>
                             <tbody className="divide-y divide-slate-100">
                               {inv.Items.map((it, j) => (
                                 <tr key={j} className="hover:bg-slate-50/50">
                                   <td className="px-6 py-3 text-text-main font-bold truncate max-w-[300px]">{it.Particulars}</td>
                                   <td className="px-4 py-3 text-right font-mono font-bold text-text-muted">{it.Quantity}</td>
                                   <td className="px-4 py-3 text-right font-mono font-bold text-text-muted">{fmtCur(it.Value / (it.Quantity || 1))}</td>
                                   <td className="px-6 py-3 text-right font-mono font-black text-text-main">{fmtCur(it.Value)}</td>
                                 </tr>
                               ))}
                             </tbody>
                             <tfoot>
                               <tr className="bg-slate-100/50 font-black text-text-main border-t border-border-custom">
                                  <td colSpan={3} className="px-6 py-4 text-right uppercase tracking-wider text-[10px] text-text-muted">Invoice Final Value</td>
                                  <td className="px-6 py-4 text-right text-base text-primary">{fmtCur(inv.Value)}</td>
                                </tr>
                             </tfoot>
                          </table>
                       </div>
                    </div>
                 ))}
                 {dynamicInvoices.filter(inv => inv.Buyer.toLowerCase().trim() === showInvPopup?.toLowerCase().trim()).length === 0 && (
                    <div className="py-20 text-center opacity-40 italic font-bold">No verified invoicing records located for this specific buyer profile.</div>
                 )}
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}
