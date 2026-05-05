export interface Product {
  id?: string;
  name: string;
  price: number;
  purchasePrice?: number;
  supplier?: string;
  brand?: string;
  category: string;
  stock: number;
  imageUrl?: string;
  createdAt: string;
}

export interface Category {
  id: string;
  name: string;
}

export interface Brand {
  id: string;
  name: string;
}

export interface Client {
  id?: string;
  clientNumber?: number;
  name: string;
  email?: string;
  phone?: string;
  address?: string;
  salonName?: string;
  password?: string;
  createdAt: string;
}

export interface CartItem extends Product {
  quantity: number;
}

export interface PaymentMethod {
  type: 'Efectivo' | 'Tarjeta' | 'Transferencia' | 'Otro';
  amount: number;
}

export interface Sale {
  id?: string;
  items: {
    productId: string;
    name: string;
    price: number;
    quantity: number;
  }[];
  total: number;
  timestamp: string;
  userId: string;
  clientId?: string;
  folio?: number;
  paymentMethods: PaymentMethod[];
}

export enum View {
  POS = 'pos',
  INVENTORY = 'inventory',
  PRODUCTS = 'products_list',
  DASHBOARD = 'dashboard',
  SETTINGS = 'settings',
  CLIENTS = 'clients',
  SALES_HISTORY = 'sales_history',
  WEB_STORE = 'web_store',
  REPORTS = 'reports'
}

export interface TicketSettings {
  logoUrl?: string;
  companyName: string;
  fiscalName?: string;
  address?: string;
  phone?: string;
  email?: string;
  footerMessage?: string;
  secondaryFooterMessage?: string;
  showLogo: boolean;
  showFiscalName: boolean;
  showAddress: boolean;
  showPhone: boolean;
  showEmail: boolean;
}

export interface AppSettings {
  nextFolio: number;
  nextClientNumber: number;
  showWebStore?: boolean;
  ticket: TicketSettings;
}
