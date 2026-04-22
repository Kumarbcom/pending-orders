export interface SalesOrder {
  Date: string;
  Order: string;
  PartyName: string;
  NameOfItem: string;
  MaterialCode: string;
  PartNo: string;
  Ordered: number;
  Balance: number;
  Rate: number;
  Discount: number;
  Value: number;
  DueOn: string | null;
  DueSerial: number | null;
  Make: string;
  MaterialGroup: string;
  Group: string;
  CustomerGroup: string;
  OrderType: "Due" | "Schedule";
  StockAllocated: number;
  StockShortfall: number;
  StockStatus: "Available" | "PO Exist - Expedite" | "Need to Place Order";
  POStatus: string;
  ExpDelivery: string;
}

export interface PurchaseOrder {
  Date: string;
  Order: string;
  PartyName: string;
  NameOfItem: string;
  MaterialCode: string;
  PartNo: string;
  Ordered: number;
  Balance: number;
  Rate: number;
  Discount: number;
  Value: number;
  DueOn: string | null;
}

export interface StockItem {
  Particulars: string;
  Quantity: number;
  Rate: number;
  Value: number;
}

export interface MaterialMasterItem {
  Description: string;
  PartNo: string;
  Make: string;
  MaterialGroup: string;
}

export interface CustomerMasterItem {
  CustomerName: string;
  Group: string;
  SalesRep: string;
  Status: string;
  CustomerGroup: string;
}

export interface InvoiceItem {
  Particulars: string;
  Quantity: number;
  Value: number;
}

export interface Invoice {
  Date: string;
  Buyer: string;
  Consignee: string;
  VoucherNo: string;
  VoucherRef: string;
  Quantity: number;
  Value: number;
  Items: InvoiceItem[];
}
