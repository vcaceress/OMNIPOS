import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import { 
  LayoutGrid, 
  ShoppingCart, 
  Package, 
  BarChart3, 
  LogOut, 
  Plus, 
  Minus, 
  Trash2, 
  Search, 
  ChevronRight,
  ChevronDown,
  ChevronUp,
  ChevronsUpDown,
  Menu,
  X,
  User,
  History,
  Settings,
  Users,
  Edit,
  Globe,
  Tag,
  List,
  FileText,
  Calendar,
  Clock,
  Filter,
  Download,
  Printer,
  MessageSquare,
  CheckCircle2,
  Upload,
  Save,
  Eye,
  Lock,
  Mail,
  ArrowLeft,
  ArrowRight,
  PieChart,
  TrendingUp,
  Image as ImageIcon
} from 'lucide-react';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import { motion, AnimatePresence } from 'motion/react';
import { 
  ResponsiveContainer, 
  BarChart, 
  Bar, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  Legend, 
  LineChart, 
  Line,
  Cell
} from 'recharts';
import { auth, loginWithGoogle, logout, db } from './firebase';
import { onAuthStateChanged, User as FirebaseUser } from 'firebase/auth';
import { 
  collection, 
  onSnapshot, 
  query, 
  orderBy, 
  addDoc, 
  updateDoc, 
  deleteDoc, 
  doc,
  serverTimestamp,
  increment,
  setDoc,
  getDocFromServer,
  writeBatch
} from 'firebase/firestore';
import { Product, Category, Brand, CartItem, Sale, View, AppSettings, TicketSettings, Client, PaymentMethod } from './types';
import { handleFirestoreError, OperationType } from './services/firestoreService';

// --- Components ---

const Highlight = ({ text, query }: { text: string, query: string }) => {
  if (!query.trim()) return <>{text}</>;
  const parts = text.split(new RegExp(`(${query})`, 'gi'));
  return (
    <>
      {parts.map((part, i) => 
        part.toLowerCase() === query.toLowerCase() 
          ? <span key={i} className="bg-amber-100 text-amber-900 rounded-sm px-0.5">{part}</span> 
          : part
      )}
    </>
  );
};

const SidebarItem = ({ icon: Icon, label, active, onClick }: { icon: any, label: string, active: boolean, onClick: () => void }) => (
  <button
    onClick={onClick}
    className={`w-full flex items-center gap-3 px-4 py-2.5 rounded-lg transition-all duration-300 group ${
      active 
        ? 'bg-zinc-900 text-white shadow-md shadow-zinc-200/50' 
        : 'text-zinc-500 hover:bg-zinc-100 hover:text-zinc-900'
    }`}
  >
    <Icon size={18} className={`transition-colors ${active ? 'text-white' : 'text-zinc-400 group-hover:text-zinc-900'}`} />
    <span className="text-sm font-semibold tracking-tight">{label}</span>
    {active && (
      <motion.div 
        layoutId="active-nav-indicator"
        className="ml-auto w-1.5 h-1.5 rounded-full bg-white/40"
      />
    )}
  </button>
);

const Login = ({ settings }: { settings: AppSettings }) => (
  <div className="min-h-screen flex items-center justify-center bg-zinc-50 p-6">
    <motion.div 
      initial={{ opacity: 0, y: 30 }}
      animate={{ opacity: 1, y: 0 }}
      transition={{ duration: 0.6, ease: [0.22, 1, 0.36, 1] }}
      className="max-w-md w-full bg-white p-10 rounded-[2.5rem] shadow-[0_32px_64px_-16px_rgba(0,0,0,0.08)] border border-zinc-100 text-center"
    >
      <div className="w-20 h-20 bg-zinc-900 rounded-3xl flex items-center justify-center mx-auto mb-8 shadow-2xl shadow-zinc-900/20 rotate-6 overflow-hidden">
        {settings.ticket?.logoUrl ? (
          <img src={settings.ticket.logoUrl} alt="Logo" className="w-full h-full object-contain p-2 -rotate-6" referrerPolicy="no-referrer" />
        ) : (
          <ShoppingCart className="text-white w-10 h-10 -rotate-6" />
        )}
      </div>
      <h1 className="text-4xl font-extrabold text-zinc-900 mb-2 tracking-tighter uppercase">{settings.ticket?.companyName || 'OmniPOS'}</h1>
      <p className="text-zinc-500 mb-10 font-medium">{settings.ticket?.companyName ? 'Gestión inteligente para tu negocio' : 'Gestión inteligente para tu negocio'}</p>
      
      <div className="space-y-4">
        <button
          onClick={loginWithGoogle}
          className="w-full flex items-center justify-center gap-3 bg-zinc-900 text-white px-6 py-5 rounded-[1.25rem] font-bold hover:bg-zinc-800 transition-all duration-300 shadow-xl shadow-zinc-900/10 active:scale-[0.98]"
        >
          <img src="https://www.google.com/favicon.ico" alt="Google" className="w-5 h-5 brightness-0 invert" referrerPolicy="no-referrer" />
          Acceso Administrador
        </button>

        <div className="relative py-4 flex items-center">
          <div className="flex-grow border-t border-zinc-100"></div>
          <span className="flex-shrink mx-4 text-[10px] font-black text-zinc-300 uppercase tracking-widest">O accede como</span>
          <div className="flex-grow border-t border-zinc-100"></div>
        </div>

        <button
          onClick={() => window.location.hash = '#tienda'}
          className="w-full flex items-center justify-center gap-3 bg-white border border-zinc-200 text-zinc-900 px-6 py-5 rounded-[1.25rem] font-bold hover:bg-zinc-50 transition-all duration-300 active:scale-[0.98] shadow-sm"
        >
          <Globe size={20} className="text-zinc-400" />
          Entrar a la Tienda Web
        </button>
      </div>

      <p className="mt-10 text-[10px] text-zinc-300 font-bold uppercase tracking-widest leading-relaxed">
        OmniPOS &copy; 2026<br/>
        Professional Retail Experience
      </p>
    </motion.div>
  </div>
);

// --- Ticket Component ---

const Ticket = ({ sale, clients, settings, isPreview = false, overrideSettings }: { sale: Sale, clients: Client[], settings: AppSettings, isPreview?: boolean, overrideSettings?: TicketSettings }) => {
  const client = clients.find(c => c.id === sale.clientId);
  const ticket = overrideSettings || settings.ticket;
  
  return (
    <div className={`bg-white ${isPreview ? 'p-8 shadow-sm border border-zinc-200 w-[80mm] scale-110 origin-top' : 'p-4 w-[80mm]'} font-mono text-[10px] text-black leading-tight mx-auto`}>
      {ticket.showLogo && ticket.logoUrl && (
        <div className="text-center mb-4">
          <img src={ticket.logoUrl} alt="Logo" className="max-w-[40mm] mx-auto" referrerPolicy="no-referrer" />
        </div>
      )}
      
      <div className="text-center mb-4">
        <h1 className="text-sm font-bold uppercase leading-tight mb-1">{ticket.companyName}</h1>
        {ticket.showFiscalName && ticket.fiscalName && <p className="text-[8px] font-bold">{ticket.fiscalName}</p>}
        {ticket.showAddress && ticket.address && <p className="text-[8px] leading-tight">{ticket.address}</p>}
        {ticket.showPhone && ticket.phone && <p className="text-[8px]">Tel: {ticket.phone}</p>}
        {ticket.showEmail && ticket.email && <p className="text-[8px]">{ticket.email}</p>}
      </div>

      <div className="text-center border-y border-dashed border-black py-2 mb-4">
        <h2 className="font-bold tracking-[0.2em] uppercase text-[9px]">Recibo de Venta</h2>
      </div>

      <div className="space-y-1 mb-4">
        <div className="flex justify-between">
          <span className="font-bold uppercase">Folio:</span>
          <span>#{sale.folio}</span>
        </div>
        <div className="flex justify-between">
          <span className="font-bold uppercase">Fecha:</span>
          <span>{new Date(sale.timestamp).toLocaleDateString()}</span>
        </div>
        <div className="flex justify-between">
          <span className="font-bold uppercase">Hora:</span>
          <span>{new Date(sale.timestamp).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', hour12: true })}</span>
        </div>
        <div className="border-t border-dotted border-black pt-1 mt-1">
          <p className="font-bold uppercase text-[8px] mb-0.5">Cliente:</p>
          <p className="uppercase">{client?.name || 'Consumidor Final'}</p>
        </div>
      </div>

      <table className="w-full mb-2">
        <thead>
          <tr className="border-y border-dashed border-black">
            <th className="text-left font-bold py-1 uppercase text-[8px] w-1/2">Producto</th>
            <th className="text-center font-bold py-1 uppercase text-[8px]">Cant</th>
            <th className="text-right font-bold py-1 uppercase text-[8px]">Total</th>
          </tr>
        </thead>
        <tbody className="divide-y divide-dotted divide-black">
          {sale.items.map((item, i) => (
            <tr key={i}>
              <td className="py-1.5 leading-tight pr-2">
                <div className="font-medium text-[9px] uppercase">{item.name}</div>
                <div className="text-[7px] text-zinc-600">${item.price.toFixed(2)} c/u</div>
              </td>
              <td className="text-center py-1.5 align-top font-bold">{item.quantity}</td>
              <td className="text-right py-1.5 align-top font-bold">${(item.price * item.quantity).toFixed(2)}</td>
            </tr>
          ))}
        </tbody>
      </table>

      <div className="border-t border-dashed border-black pt-2 mt-2 space-y-1">
        <div className="flex justify-between text-[8px]">
          <span>CANTIDAD DE ARTICULOS:</span>
          <span>{sale.items.reduce((acc, item) => acc + item.quantity, 0)}</span>
        </div>
        <div className="flex justify-between text-[11px] font-black py-2 border-y border-double border-black mt-1">
          <span>TOTAL:</span>
          <span>${sale.total.toFixed(2)}</span>
        </div>
      </div>

      <div className="mt-8 text-center text-[8px] space-y-2">
        {ticket.footerMessage ? (
          <>
            <p className="font-bold uppercase">{ticket.footerMessage}</p>
            {ticket.secondaryFooterMessage && <p>{ticket.secondaryFooterMessage}</p>}
          </>
        ) : (
          <>
            <p className="font-bold italic uppercase">¡Gracias por su preferencia!</p>
            <p className="uppercase">Es un placer atenderle</p>
          </>
        )}
        <div className="pt-4 opacity-50">
          <p>*** FIN DEL TICKET ***</p>
        </div>
      </div>
    </div>
  );
};


export default function App() {
  const [user, setUser] = useState<FirebaseUser | null>(null);
  const [loading, setLoading] = useState(true);
  const [hash, setHash] = useState(window.location.hash);
  const [currentView, setCurrentView] = useState<View>(View.POS);
  const [products, setProducts] = useState<Product[]>([]);
  const [categories, setCategories] = useState<Category[]>([]);
  const [brands, setBrands] = useState<Brand[]>([]);
  const [cart, setCart] = useState<CartItem[]>([]);
  const [searchQuery, setSearchQuery] = useState('');
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);
  const [sales, setSales] = useState<Sale[]>([]);
  const [clients, setClients] = useState<Client[]>([]);
  const [selectedClientId, setSelectedClientId] = useState<string>('');
  const [selectedCategory, setSelectedCategory] = useState<string>('Todas');
  const [posLayout, setPosLayout] = useState<'grid' | 'list'>('grid');
  const [isPaymentModalOpen, setIsPaymentModalOpen] = useState(false);
  const [isReceiptModalOpen, setIsReceiptModalOpen] = useState(false);
  const [showClientAlert, setShowClientAlert] = useState(false);
  const [lastSale, setLastSale] = useState<Sale | null>(null);
  const [paymentMethods, setPaymentMethods] = useState<PaymentMethod[]>([]);
  const [saleToPrint, setSaleToPrint] = useState<Sale | null>(null);
  const [printSettingsOverride, setPrintSettingsOverride] = useState<TicketSettings | null>(null);
  const [settings, setSettings] = useState<AppSettings>({ 
    nextFolio: 1, 
    nextClientNumber: 1,
    showWebStore: true,
    ticket: {
      companyName: 'OmniPOS',
      showLogo: false,
      showFiscalName: false,
      showAddress: false,
      showPhone: false,
      showEmail: false
    }
  });

  useEffect(() => {
    const handleHashChange = () => setHash(window.location.hash);
    window.addEventListener('hashchange', handleHashChange);
    return () => window.removeEventListener('hashchange', handleHashChange);
  }, []);

  useEffect(() => {
    async function testConnection() {
      try {
        await getDocFromServer(doc(db, 'test', 'connection'));
      } catch (error) {
        if(error instanceof Error && error.message.includes('the client is offline')) {
          console.error("Please check your Firebase configuration.");
        }
      }
    }
    testConnection();

    const unsubscribe = onAuthStateChanged(auth, (user) => {
      setUser(user);
      setLoading(false);
    });
    return () => unsubscribe();
  }, []);

  useEffect(() => {
    if (!user) return;

    const productsQuery = query(collection(db, 'products'), orderBy('createdAt', 'desc'));
    const categoriesQuery = query(collection(db, 'categories'), orderBy('name', 'asc'));
    const brandsQuery = query(collection(db, 'brands'), orderBy('name', 'asc'));
    const salesQuery = query(collection(db, 'sales'), orderBy('timestamp', 'desc'));

    const unsubProducts = onSnapshot(productsQuery, (snapshot) => {
      setProducts(snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as Product)));
    }, (err) => handleFirestoreError(err, OperationType.LIST, 'products'));

    const unsubCategories = onSnapshot(categoriesQuery, (snapshot) => {
      setCategories(snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as Category)));
    }, (err) => handleFirestoreError(err, OperationType.LIST, 'categories'));

    const unsubBrands = onSnapshot(brandsQuery, (snapshot) => {
      setBrands(snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as Brand)));
    }, (err) => handleFirestoreError(err, OperationType.LIST, 'brands'));

    const unsubSales = onSnapshot(salesQuery, (snapshot) => {
      setSales(snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as Sale)));
    }, (err) => handleFirestoreError(err, OperationType.LIST, 'sales'));

    const unsubClients = onSnapshot(query(collection(db, 'clients'), orderBy('name', 'asc')), (snapshot) => {
      const clientsData = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() } as Client));
      setClients(clientsData);
      
      // Migration: Assign numbers to existing clients that don't have one
      const clientsWithoutNumber = snapshot.docs.filter(d => !d.data().clientNumber);
      if (clientsWithoutNumber.length > 0) {
        const currentNext = settings.nextClientNumber || 1;
        let count = 0;
        clientsWithoutNumber.forEach(async (clientDoc) => {
          const num = currentNext + count;
          await updateDoc(doc(db, 'clients', clientDoc.id), { clientNumber: num });
          count++;
        });
        updateDoc(doc(db, 'settings', 'global'), { nextClientNumber: currentNext + clientsWithoutNumber.length });
      }
    }, (err) => handleFirestoreError(err, OperationType.LIST, 'clients'));

    const unsubSettings = onSnapshot(doc(db, 'settings', 'global'), (snapshot) => {
      if (snapshot.exists()) {
        const data = snapshot.data();
        setSettings({
          ...data,
          nextFolio: data.nextFolio || 1,
          nextClientNumber: data.nextClientNumber || 1,
          showWebStore: data.showWebStore !== undefined ? data.showWebStore : true,
          ticket: data.ticket || {
            companyName: 'OmniPOS',
            showLogo: false,
            showFiscalName: false,
            showAddress: false,
            showPhone: false,
            showEmail: false
          }
        } as AppSettings);
      } else {
        setDoc(doc(db, 'settings', 'global'), { 
          nextFolio: 1, 
          nextClientNumber: 1,
          showWebStore: true,
          ticket: {
            companyName: 'OmniPOS',
            showLogo: false,
            showFiscalName: false,
            showAddress: false,
            showPhone: false,
            showEmail: false
          }
        });
      }
    }, (err) => handleFirestoreError(err, OperationType.GET, 'settings'));

    return () => {
      unsubProducts();
      unsubCategories();
      unsubBrands();
      unsubSales();
      unsubClients();
      unsubSettings();
    };
  }, [user]);

  // Efecto temporal para importar productos de la marca Organic
  useEffect(() => {
    const importOrganic = async () => {
      if (!user || loading) return;
      
      const alreadyImported = localStorage.getItem('organic_imported_v1');
      if (alreadyImported) return;

      try {
        const brandName = 'Organic';
        
        // Asegurar que la marca existe
        if (!brands.some(b => b.name.toLowerCase() === brandName.toLowerCase())) {
          await addDoc(collection(db, 'brands'), { name: brandName });
        }

        const organicProducts = [
          { name: 'Shampoo Orgánico', price: 15, purchasePrice: 8, stock: 20, category: 'Cuidado Capilar', brand: brandName },
          { name: 'Acondicionador Orgánico', price: 15, purchasePrice: 8, stock: 20, category: 'Cuidado Capilar', brand: brandName },
          { name: 'Crema Facial Orgánica', price: 25, purchasePrice: 12, stock: 15, category: 'Cuidado Facial', brand: brandName },
          { name: 'Jabón Líquido Orgánico', price: 12, purchasePrice: 6, stock: 25, category: 'Cuidado Corporal', brand: brandName },
          { name: 'Jabón de Manos Orgánico', price: 8, purchasePrice: 4, stock: 30, category: 'Cuidado Corporal', brand: brandName }
        ];

        // Asegurar que las categorías existen
        const neededCats = ['Cuidado Capilar', 'Cuidado Facial', 'Cuidado Corporal'];
        for (const catName of neededCats) {
          if (!categories.some(c => c.name === catName)) {
            await addDoc(collection(db, 'categories'), { name: catName });
          }
        }

        // Añadir los productos si no existen por nombre
        for (const p of organicProducts) {
          if (!products.some(existing => existing.name === p.name)) {
            await addDoc(collection(db, 'products'), {
              ...p,
              createdAt: new Date().toISOString()
            });
          }
        }
        
        localStorage.setItem('organic_imported_v1', 'true');
      } catch (error) {
        console.error('Error al importar productos Organic:', error);
      }
    };

    importOrganic();
  }, [user, loading, products, brands, categories]);

  const addToCart = (product: Product) => {
    setCart(prev => {
      const existing = prev.find(item => item.id === product.id);
      if (existing) {
        return prev.map(item => item.id === product.id ? { ...item, quantity: item.quantity + 1 } : item);
      }
      return [...prev, { ...product, quantity: 1 }];
    });
  };

  const removeFromCart = (productId: string) => {
    setCart(prev => prev.filter(item => item.id !== productId));
  };

  const updateQuantity = (productId: string, delta: number) => {
    setCart(prev => prev.map(item => {
      if (item.id === productId) {
        const newQty = Math.max(1, item.quantity + delta);
        return { ...item, quantity: newQty };
      }
      return item;
    }));
  };

  const cartTotal = cart.reduce((sum, item) => sum + (item.price * item.quantity), 0);

  const checkout = () => {
    if (cart.length === 0 || !user) return;
    if (!selectedClientId) {
      setShowClientAlert(true);
      return;
    }
    setPaymentMethods([{ type: 'Efectivo', amount: cartTotal }]);
    setIsPaymentModalOpen(true);
  };

  const processSale = async () => {
    const totalPaid = paymentMethods.reduce((sum, pm) => sum + pm.amount, 0);
    if (Math.abs(totalPaid - cartTotal) > 0.01) {
      alert('El monto total de los métodos de pago debe coincidir con el total de la venta.');
      return;
    }

    try {
      const currentFolio = settings.nextFolio;
      const saleData: Sale = {
        items: cart.map(item => ({
          productId: item.id!,
          name: item.name,
          price: item.price,
          quantity: item.quantity
        })),
        total: cartTotal,
        timestamp: new Date().toISOString(),
        userId: user.uid,
        clientId: selectedClientId || null,
        folio: currentFolio,
        paymentMethods: paymentMethods
      };

      await addDoc(collection(db, 'sales'), saleData);
      
      // Update folio
      await updateDoc(doc(db, 'settings', 'global'), {
        nextFolio: currentFolio + 1
      });

      // Update stock
      for (const item of cart) {
        const productRef = doc(db, 'products', item.id!);
        await updateDoc(productRef, {
          stock: increment(-item.quantity)
        });
      }

      setCart([]);
      setSelectedClientId('');
      setIsPaymentModalOpen(false);
      setLastSale(saleData);
      setIsReceiptModalOpen(true);
    } catch (error) {
      handleFirestoreError(error, OperationType.WRITE, 'sales');
    }
  };

  const addPaymentMethod = () => {
    if (paymentMethods.length >= 2) return;
    const currentTotal = paymentMethods.reduce((sum, pm) => sum + pm.amount, 0);
    const remaining = Math.max(0, cartTotal - currentTotal);
    setPaymentMethods([...paymentMethods, { type: 'Tarjeta', amount: remaining }]);
  };

  const removePaymentMethod = (index: number) => {
    const newMethods = paymentMethods.filter((_, i) => i !== index);
    if (newMethods.length === 1) {
      newMethods[0].amount = cartTotal;
    }
    setPaymentMethods(newMethods);
  };

  const updatePaymentMethod = (index: number, field: keyof PaymentMethod, value: any) => {
    setPaymentMethods(paymentMethods.map((pm, i) => i === index ? { ...pm, [field]: value } : pm));
  };

  const printSaleReceipt = (sale: Sale, overrideSettings?: TicketSettings) => {
    setSaleToPrint(sale);
    setPrintSettingsOverride(overrideSettings || null);
    // Give enough time for the state to update and the component to render
    setTimeout(() => {
      window.print();
      // Reset after a delay to allow the print dialog to close
      setTimeout(() => {
        setSaleToPrint(null);
        setPrintSettingsOverride(null);
      }, 1000);
    }, 800);
  };

  if (loading) return (
    <div className="min-h-screen flex items-center justify-center bg-zinc-50">
      <motion.div 
        animate={{ rotate: 360 }}
        transition={{ repeat: Infinity, duration: 1, ease: "linear" }}
        className="w-12 h-12 border-4 border-zinc-200 border-t-zinc-900 rounded-full"
      />
    </div>
  );

  // Allow access to Web Store even if not logged in as admin
  const isWebStoreMode = hash === '#tienda';
  
  if (!user && !isWebStoreMode) return <Login settings={settings} />;

  if (isWebStoreMode) {
    return (
      <div className="min-h-screen flex flex-col bg-zinc-50">
        <WebStoreView products={products} clients={clients} onSaleComplete={() => {}} />
      </div>
    );
  }

  const filteredProducts = searchQuery.trim() === '' ? [] : products.filter(p => {
    const matchesSearch = p.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
                         p.category.toLowerCase().includes(searchQuery.toLowerCase()) ||
                         (p.brand && p.brand.toLowerCase().includes(searchQuery.toLowerCase()));
    const matchesCategory = selectedCategory === 'Todas' || p.category === selectedCategory;
    return matchesSearch && matchesCategory;
  });

  return (
    <div className={`min-h-screen bg-zinc-50 flex flex-col md:flex-row font-sans text-zinc-900 ${saleToPrint ? 'is-printing-ticket' : ''}`}>
      {/* --- Desktop Sidebar --- */}
      <aside className="hidden md:flex flex-col w-64 bg-zinc-50 border-r border-zinc-200/60 p-6 no-print">
        <div className="flex items-center gap-3 mb-12 px-2">
          <div className="w-9 h-9 bg-zinc-900 rounded-xl flex items-center justify-center shadow-lg rotate-3 group-hover:rotate-0 transition-transform overflow-hidden">
            {settings.ticket?.logoUrl ? (
              <img src={settings.ticket.logoUrl} alt="Logo" className="w-full h-full object-contain p-1 -rotate-3" referrerPolicy="no-referrer" />
            ) : (
              <ShoppingCart className="text-white w-4.5 h-4.5 -rotate-3 group-hover:rotate-0 transition-transform" />
            )}
          </div>
          <h1 className="text-lg font-black tracking-tighter uppercase truncate">{settings.ticket?.companyName || 'OmniPOS'}</h1>
        </div>

        <nav className="flex-1 space-y-1">
          <div className="px-2 mb-4">
            <p className="text-[10px] font-black text-zinc-400 uppercase tracking-[0.2em] mb-4">Principal</p>
            <div className="space-y-1">
              <SidebarItem icon={LayoutGrid} label="Ventas" active={currentView === View.POS} onClick={() => setCurrentView(View.POS)} />
              <SidebarItem icon={List} label="Productos" active={currentView === View.PRODUCTS} onClick={() => setCurrentView(View.PRODUCTS)} />
              <SidebarItem icon={Package} label="Inventario" active={currentView === View.INVENTORY} onClick={() => setCurrentView(View.INVENTORY)} />
              <SidebarItem icon={Users} label="Clientes" active={currentView === View.CLIENTS} onClick={() => setCurrentView(View.CLIENTS)} />
            </div>
          </div>

          <div className="px-2 mt-8">
            <p className="text-[10px] font-black text-zinc-400 uppercase tracking-[0.2em] mb-4">Reportes & Web</p>
            <div className="space-y-1">
              <SidebarItem icon={History} label="Historial" active={currentView === View.SALES_HISTORY} onClick={() => setCurrentView(View.SALES_HISTORY)} />
              {settings.showWebStore && <SidebarItem icon={Globe} label="Tienda Web" active={currentView === View.WEB_STORE} onClick={() => setCurrentView(View.WEB_STORE)} />}
              <SidebarItem icon={BarChart3} label="Panel" active={currentView === View.DASHBOARD} onClick={() => setCurrentView(View.DASHBOARD)} />
              <SidebarItem icon={TrendingUp} label="Reportes" active={currentView === View.REPORTS} onClick={() => setCurrentView(View.REPORTS)} />
            </div>
          </div>
        </nav>

        <div className="mt-auto pt-6">
          <div className="px-2 mb-6">
            <SidebarItem icon={Settings} label="Configuración" active={currentView === View.SETTINGS} onClick={() => setCurrentView(View.SETTINGS)} />
          </div>
          
          <div className="bg-white p-4 rounded-2xl border border-zinc-200/60 shadow-sm">
            <div className="flex items-center gap-3 mb-4">
              <div className="w-10 h-10 bg-zinc-100 rounded-xl flex items-center justify-center overflow-hidden border border-zinc-200">
                {user.photoURL ? (
                  <img src={user.photoURL} alt={user.displayName || ''} className="w-full h-full object-cover" referrerPolicy="no-referrer" />
                ) : (
                  <User size={18} className="text-zinc-400" />
                )}
              </div>
              <div className="flex-1 min-w-0">
                <p className="text-xs font-bold truncate text-zinc-900">{user.displayName}</p>
                <p className="text-[10px] text-zinc-400 font-medium truncate">{user.email}</p>
              </div>
            </div>
            <button
              onClick={logout}
              className="w-full flex items-center justify-center gap-2 py-2 rounded-xl text-zinc-400 hover:text-red-500 hover:bg-red-50 transition-all duration-300 text-xs font-bold border border-transparent hover:border-red-100"
            >
              <LogOut size={14} />
              <span>Cerrar Sesión</span>
            </button>
          </div>
        </div>
      </aside>

      {/* --- Mobile Header --- */}
      <header className="md:hidden bg-white border-b border-zinc-100 px-4 py-3 flex items-center justify-between sticky top-0 z-50 no-print h-[60px] shrink-0">
        <div className="flex items-center gap-2">
          <div className="w-8 h-8 bg-zinc-900 rounded-lg flex items-center justify-center overflow-hidden">
            {settings.ticket?.logoUrl ? (
              <img src={settings.ticket.logoUrl} alt="Logo" className="w-full h-full object-contain p-1" referrerPolicy="no-referrer" />
            ) : (
              <ShoppingCart className="text-white w-4 h-4" />
            )}
          </div>
          <span className="font-bold tracking-tight uppercase truncate max-w-[150px]">{settings.ticket?.companyName || 'OmniPOS'}</span>
        </div>
        <button onClick={() => setIsSidebarOpen(true)} className="p-2 hover:bg-zinc-100 rounded-lg">
          <Menu size={24} />
        </button>
      </header>

      {/* --- Mobile Sidebar Overlay --- */}
      <AnimatePresence>
        {isSidebarOpen && (
          <>
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setIsSidebarOpen(false)}
              className="fixed inset-0 bg-black/20 backdrop-blur-sm z-[60] md:hidden"
            />
            <motion.aside
              initial={{ x: '100%' }}
              animate={{ x: 0 }}
              exit={{ x: '100%' }}
              className="fixed right-0 top-0 bottom-0 w-72 bg-white z-[70] p-6 flex flex-col shadow-2xl md:hidden"
            >
              <div className="flex justify-between items-center mb-10">
                <span className="font-bold">Menu</span>
                <button onClick={() => setIsSidebarOpen(false)} className="p-2 hover:bg-zinc-100 rounded-lg">
                  <X size={24} />
                </button>
              </div>
              <nav className="flex-1 space-y-2">
                <SidebarItem icon={LayoutGrid} label="Ventas" active={currentView === View.POS} onClick={() => { setCurrentView(View.POS); setIsSidebarOpen(false); }} />
                <SidebarItem icon={List} label="Productos" active={currentView === View.PRODUCTS} onClick={() => { setCurrentView(View.PRODUCTS); setIsSidebarOpen(false); }} />
                <SidebarItem icon={Package} label="Inventario" active={currentView === View.INVENTORY} onClick={() => { setCurrentView(View.INVENTORY); setIsSidebarOpen(false); }} />
                <SidebarItem icon={Users} label="Clientes" active={currentView === View.CLIENTS} onClick={() => { setCurrentView(View.CLIENTS); setIsSidebarOpen(false); }} />
                <SidebarItem icon={History} label="Historial" active={currentView === View.SALES_HISTORY} onClick={() => { setCurrentView(View.SALES_HISTORY); setIsSidebarOpen(false); }} />
                {settings.showWebStore && <SidebarItem icon={Globe} label="Tienda Web" active={currentView === View.WEB_STORE} onClick={() => { setCurrentView(View.WEB_STORE); setIsSidebarOpen(false); }} />}
                <SidebarItem icon={BarChart3} label="Panel" active={currentView === View.DASHBOARD} onClick={() => { setCurrentView(View.DASHBOARD); setIsSidebarOpen(false); }} />
                <SidebarItem icon={TrendingUp} label="Reportes" active={currentView === View.REPORTS} onClick={() => { setCurrentView(View.REPORTS); setIsSidebarOpen(false); }} />
                <SidebarItem icon={Settings} label="Configuración" active={currentView === View.SETTINGS} onClick={() => { setCurrentView(View.SETTINGS); setIsSidebarOpen(false); }} />
              </nav>
              <button onClick={logout} className="mt-auto flex items-center gap-3 p-4 text-red-500 font-bold">
                <LogOut size={20} /> Cerrar Sesión
              </button>
            </motion.aside>
          </>
        )}
      </AnimatePresence>

      {/* --- Main Content --- */}
      <main className={`flex-1 flex flex-col h-[calc(100vh-64px)] md:h-screen overflow-hidden ${saleToPrint ? 'no-print' : ''} bg-white md:m-3 md:rounded-[2rem] border border-zinc-200/50 shadow-sm relative overflow-hidden`}>
        <AnimatePresence mode="wait">
          <motion.div
            key={currentView}
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: -20 }}
            transition={{ duration: 0.4, ease: [0.22, 1, 0.36, 1] }}
            className="flex-1 flex flex-col overflow-hidden"
          >
            {currentView === View.POS && (
          <div className="flex-1 flex flex-col lg:flex-row overflow-hidden bg-zinc-50/50">
            {/* Product Selection Area */}
            <div className="flex-1 flex flex-col p-4 md:p-6 overflow-hidden">
              <div className="flex flex-col gap-6 mb-6">
                <div className="flex flex-col gap-4">
                  <div className="flex items-center gap-3">
                    <h2 className="text-2xl md:text-3xl font-black tracking-tight text-zinc-900">Nueva Venta</h2>
                    <div className="flex items-center gap-2">
                      <span className="bg-zinc-900 text-white px-3 py-1 rounded-lg text-[10px] font-black uppercase tracking-widest shadow-sm">
                        Folio: #{settings.nextFolio}
                      </span>
                      <span className="bg-white border border-zinc-200 text-zinc-500 px-3 py-1 rounded-lg text-[10px] font-bold uppercase tracking-widest shadow-sm flex items-center gap-2">
                        <span>{new Date().toLocaleDateString('es-ES', { day: '2-digit', month: 'short', year: 'numeric' })}</span>
                        <span className="border-l border-zinc-200 pl-2 text-zinc-900">{new Date().toLocaleTimeString('es-ES', { hour: '2-digit', minute: '2-digit', hour12: true })}</span>
                      </span>
                    </div>
                  </div>

                  <div className="flex flex-col md:flex-row md:items-center gap-4">
                    <div className="flex items-end gap-2 w-full md:w-80">
                      <div className="flex-1">
                        <label className="text-[10px] font-bold text-zinc-400 uppercase tracking-widest block mb-1.5 ml-1">Seleccionar Cliente</label>
                        <select 
                          value={selectedClientId}
                          onChange={(e) => setSelectedClientId(e.target.value)}
                          className="w-full bg-white border border-zinc-200 rounded-xl p-2.5 text-sm outline-none focus:ring-2 focus:ring-zinc-900 font-medium shadow-sm transition-all"
                        >
                          <option value="">Consumidor Final</option>
                          {clients.map(client => (
                            <option key={client.id} value={client.id}>{client.name} {client.salonName ? `(${client.salonName})` : ''}</option>
                          ))}
                        </select>
                      </div>
                      <button 
                        onClick={() => setCurrentView(View.CLIENTS)}
                        className="bg-zinc-100 text-zinc-600 p-2.5 rounded-xl hover:bg-zinc-200 transition-all shadow-sm group"
                        title="Gestionar Clientes"
                      >
                        <Plus size={18} className="group-hover:rotate-90 transition-transform" />
                      </button>
                    </div>
                  </div>
                  
                  <div className="relative w-full max-w-2xl">
                    <Search className="absolute left-4 top-1/2 -translate-y-1/2 text-zinc-400" size={18} />
                    <input
                      type="text"
                      placeholder="¿Qué producto buscas hoy?..."
                      value={searchQuery}
                      onChange={(e) => setSearchQuery(e.target.value)}
                      className="w-full bg-white border border-zinc-200 rounded-2xl py-4 pl-12 pr-12 focus:ring-2 focus:ring-zinc-900 focus:border-transparent transition-all outline-none font-medium shadow-sm text-lg"
                    />
                    <AnimatePresence>
                      {searchQuery && (
                        <motion.button
                          initial={{ opacity: 0, scale: 0.8 }}
                          animate={{ opacity: 1, scale: 1 }}
                          exit={{ opacity: 0, scale: 0.8 }}
                          onClick={() => setSearchQuery('')}
                          className="absolute right-4 top-1/2 -translate-y-1/2 p-1 hover:bg-zinc-100 rounded-full text-zinc-400 hover:text-zinc-900 transition-colors"
                        >
                          <X size={16} />
                        </motion.button>
                      )}
                    </AnimatePresence>
                  </div>
                </div>
              </div>

              <div className="flex-1 overflow-y-auto pr-2 custom-scrollbar">
                {searchQuery.trim() !== '' && (
                  <>
                    {posLayout === 'grid' ? (
                      <div className="grid grid-cols-2 sm:grid-cols-3 xl:grid-cols-4 2xl:grid-cols-5 gap-4">
                        <AnimatePresence mode="popLayout">
                          {filteredProducts.map((product) => (
                            <motion.button
                              layout
                              initial={{ opacity: 0, scale: 0.9 }}
                              animate={{ opacity: 1, scale: 1 }}
                              exit={{ opacity: 0, scale: 0.9 }}
                              key={product.id}
                              whileTap={{ scale: 0.97 }}
                              onClick={() => addToCart(product)}
                              disabled={product.stock <= 0}
                              className={`group relative bg-white rounded-3xl border border-zinc-100 shadow-sm hover:shadow-xl hover:border-zinc-300 transition-all text-left flex flex-col overflow-hidden ${product.stock <= 0 ? 'opacity-50 grayscale cursor-not-allowed' : ''}`}
                            >
                              <div className="aspect-square bg-zinc-50 relative overflow-hidden">
                                {product.imageUrl ? (
                                  <img src={product.imageUrl} alt={product.name} className="w-full h-full object-cover group-hover:scale-110 transition-transform duration-500" referrerPolicy="no-referrer" />
                                ) : (
                                  <div className="w-full h-full flex items-center justify-center">
                                    <Package size={40} className="text-zinc-200" />
                                  </div>
                                )}
                                <div className="absolute top-3 right-3 bg-white/90 backdrop-blur-md px-2 py-1 rounded-lg text-[10px] font-black text-zinc-900 border border-zinc-100">
                                  {product.brand || 'S/M'}
                                </div>
                                {product.stock <= 5 && product.stock > 0 && (
                                  <div className="absolute bottom-3 left-3 bg-amber-500 text-white px-2 py-1 rounded-lg text-[10px] font-black uppercase tracking-wider">
                                    Stock Bajo: {product.stock}
                                  </div>
                                )}
                                {product.stock <= 0 && (
                                  <div className="absolute inset-0 bg-black/40 flex items-center justify-center">
                                    <span className="bg-white text-zinc-900 px-3 py-1 rounded-full text-xs font-black uppercase">Agotado</span>
                                  </div>
                                )}
                              </div>
                              <div className="p-4 flex-1 flex flex-col justify-between gap-2">
                                <div>
                                  <p className="text-[10px] font-bold text-zinc-400 uppercase tracking-widest mb-1">{product.category}</p>
                                  <h3 className="font-bold text-zinc-900 leading-tight line-clamp-2 text-sm">
                                    <Highlight text={product.name} query={searchQuery} />
                                  </h3>
                                </div>
                                <div className="flex items-center justify-between mt-2">
                                  <p className="text-lg font-black text-zinc-900">${product.price.toFixed(2)}</p>
                                  <div className="bg-zinc-100 text-zinc-400 p-2 rounded-xl group-hover:bg-zinc-900 group-hover:text-white transition-all duration-300">
                                    <Plus size={18} />
                                  </div>
                                </div>
                              </div>
                            </motion.button>
                          ))}
                        </AnimatePresence>
                      </div>
                    ) : (
                      <div className="flex flex-col gap-2">
                        <AnimatePresence mode="popLayout">
                          {filteredProducts.map((product) => (
                            <motion.button
                              layout
                              initial={{ opacity: 0, x: -20 }}
                              animate={{ opacity: 1, x: 0 }}
                              exit={{ opacity: 0, scale: 0.95 }}
                              key={product.id}
                              whileTap={{ scale: 0.99 }}
                              onClick={() => addToCart(product)}
                              disabled={product.stock <= 0}
                              className={`group relative bg-white p-3 rounded-2xl border border-zinc-100 shadow-sm hover:shadow-md hover:border-zinc-300 transition-all text-left flex items-center gap-4 ${product.stock <= 0 ? 'opacity-50 grayscale cursor-not-allowed' : ''}`}
                            >
                              <div className="w-12 h-12 bg-zinc-50 rounded-xl flex-shrink-0 overflow-hidden">
                                {product.imageUrl ? (
                                  <img src={product.imageUrl} alt={product.name} className="w-full h-full object-cover" referrerPolicy="no-referrer" />
                                ) : (
                                  <div className="w-full h-full flex items-center justify-center">
                                    <Package size={20} className="text-zinc-200" />
                                  </div>
                                )}
                              </div>
                              <div className="flex-1 min-w-0">
                                <div className="flex items-center gap-2 mb-0.5">
                                  <span className="text-[9px] font-black text-zinc-400 uppercase tracking-widest bg-zinc-50 px-1.5 py-0.5 rounded border border-zinc-100">{product.brand || 'S/M'}</span>
                                  <span className="text-[9px] font-bold text-zinc-400 uppercase tracking-widest">{product.category}</span>
                                </div>
                                <h3 className="font-bold text-zinc-900 text-sm truncate">
                                  <Highlight text={product.name} query={searchQuery} />
                                </h3>
                              </div>
                              <div className="text-right flex-shrink-0 px-4">
                                <p className="text-xs font-bold text-zinc-400 uppercase tracking-widest mb-0.5">Stock</p>
                                <p className={`text-sm font-black ${product.stock <= 5 ? 'text-amber-500' : 'text-zinc-900'}`}>{product.stock}</p>
                              </div>
                              <div className="flex items-center gap-4">
                                <p className="text-lg font-black text-zinc-900 whitespace-nowrap">${product.price.toFixed(2)}</p>
                                <div className="bg-zinc-100 text-zinc-400 p-2 rounded-xl group-hover:bg-zinc-900 group-hover:text-white transition-all duration-300">
                                  <Plus size={18} />
                                </div>
                              </div>
                            </motion.button>
                          ))}
                        </AnimatePresence>
                      </div>
                    )}
                    
                    {filteredProducts.length === 0 && (
                      <div className="col-span-full py-20 text-center">
                        <div className="w-20 h-20 bg-zinc-100 rounded-full flex items-center justify-center mx-auto mb-4">
                          <Search className="text-zinc-300" size={40} />
                        </div>
                        <p className="text-zinc-500 font-bold text-lg">No se encontraron productos</p>
                        <p className="text-zinc-400 text-sm">Prueba ajustando la búsqueda</p>
                      </div>
                    )}
                  </>
                )}
              </div>
            </div>

            {/* Cart Sidebar */}
            <div className="w-full lg:w-[420px] bg-white border-t lg:border-t-0 lg:border-l border-zinc-100 flex flex-col shadow-2xl lg:shadow-none z-10">
              <div className="p-6 border-b border-zinc-50 flex items-center justify-between">
                <div className="flex items-center gap-3">
                  <ShoppingCart size={20} className="text-zinc-900" />
                  <h2 className="text-xl font-bold tracking-tight">Pedido Actual</h2>
                </div>
                <span className="bg-zinc-100 text-zinc-900 px-3 py-1 rounded-full text-xs font-bold">
                  {cart.reduce((s, i) => s + i.quantity, 0)} artículos
                </span>
              </div>

              <div className="flex-1 overflow-y-auto p-6 space-y-4 custom-scrollbar">
                {cart.length === 0 ? (
                  <div className="h-full flex flex-col items-center justify-center text-zinc-400 opacity-50">
                    <ShoppingCart size={48} className="mb-4" />
                    <p className="font-bold">Tu carrito está vacío</p>
                  </div>
                ) : (
                  cart.map((item) => (
                    <motion.div 
                      layout
                      initial={{ opacity: 0, x: 20 }}
                      animate={{ opacity: 1, x: 0 }}
                      key={item.id} 
                      className="flex items-center gap-4 group"
                    >
                      <div className="w-16 h-16 bg-zinc-50 rounded-2xl flex-shrink-0 overflow-hidden">
                        {item.imageUrl ? (
                          <img src={item.imageUrl} alt={item.name} className="w-full h-full object-cover" referrerPolicy="no-referrer" />
                        ) : (
                          <div className="w-full h-full flex items-center justify-center"><Package size={20} className="text-zinc-200" /></div>
                        )}
                      </div>
                      <div className="flex-1 min-w-0">
                        <h4 className="font-bold text-sm truncate">{item.name}</h4>
                        <p className="text-zinc-500 text-xs font-medium">${item.price.toFixed(2)} c/u</p>
                      </div>
                      <div className="flex items-center gap-2 bg-zinc-50 p-1 rounded-xl border border-zinc-100">
                        <button onClick={() => updateQuantity(item.id!, -1)} className="p-1 hover:bg-white hover:shadow-sm rounded-lg transition-all"><Minus size={14} /></button>
                        <span className="w-6 text-center text-sm font-bold">{item.quantity}</span>
                        <button onClick={() => updateQuantity(item.id!, 1)} className="p-1 hover:bg-white hover:shadow-sm rounded-lg transition-all"><Plus size={14} /></button>
                      </div>
                      <button onClick={() => removeFromCart(item.id!)} className="p-2 text-zinc-300 hover:text-red-500 transition-colors"><Trash2 size={18} /></button>
                    </motion.div>
                  ))
                )}
              </div>

              <div className="p-6 bg-zinc-50/50 border-t border-zinc-100 space-y-4">
                <div className="flex justify-between items-center">
                  <span className="text-zinc-500 font-bold">Subtotal</span>
                  <span className="font-bold text-lg">${cartTotal.toFixed(2)}</span>
                </div>
                <div className="flex justify-between items-center">
                  <span className="text-zinc-500 font-bold">Impuestos (0%)</span>
                  <span className="font-bold text-lg">$0.00</span>
                </div>
                <div className="pt-4 border-t border-zinc-200 flex justify-between items-center">
                  <span className="text-xl font-black">Total</span>
                  <span className="text-3xl font-black tracking-tighter">${cartTotal.toFixed(2)}</span>
                </div>
                <button
                  onClick={checkout}
                  disabled={cart.length === 0}
                  className="w-full bg-zinc-900 text-white py-5 rounded-2xl font-bold text-lg shadow-xl shadow-zinc-200 hover:scale-[1.02] active:scale-95 disabled:opacity-50 disabled:hover:scale-100 transition-all duration-300 flex items-center justify-center gap-3"
                >
                  Completar Pago
                  <ChevronRight size={20} />
                </button>
              </div>
            </div>
          </div>
        )}

        <AnimatePresence>
          {isPaymentModalOpen && (
            <div className="fixed inset-0 bg-black/40 backdrop-blur-sm z-[150] flex items-center justify-center p-4 no-print">
              <motion.div
                initial={{ opacity: 0, scale: 0.9, y: 20 }}
                animate={{ opacity: 1, scale: 1, y: 0 }}
                exit={{ opacity: 0, scale: 0.9, y: 20 }}
                className="bg-white w-full max-w-md rounded-3xl shadow-2xl overflow-hidden"
              >
                <div className="p-6 border-b border-zinc-50 flex items-center justify-between">
                  <h3 className="text-xl font-bold">Método de Pago</h3>
                  <button onClick={() => setIsPaymentModalOpen(false)} className="p-2 hover:bg-zinc-100 rounded-xl transition-colors">
                    <X size={20} />
                  </button>
                </div>
                
                <div className="p-6 space-y-6">
                  <div className="bg-zinc-50 p-4 rounded-2xl flex justify-between items-center">
                    <span className="text-zinc-500 font-bold">Total a Pagar</span>
                    <span className="text-2xl font-black">${cartTotal.toFixed(2)}</span>
                  </div>

                  <div className="space-y-4">
                    {paymentMethods.map((pm, index) => (
                      <div key={index} className="space-y-3 p-4 border border-zinc-100 rounded-2xl bg-white shadow-sm relative">
                        <div className="flex justify-between items-center">
                          <label className="text-[10px] font-bold text-zinc-400 uppercase tracking-widest">Método {index + 1}</label>
                          {paymentMethods.length > 1 && (
                            <button onClick={() => removePaymentMethod(index)} className="text-red-500 hover:bg-red-50 p-1 rounded-lg transition-colors">
                              <Trash2 size={14} />
                            </button>
                          )}
                        </div>
                        <div className="grid grid-cols-2 gap-3">
                          <select
                            value={pm.type}
                            onChange={(e) => updatePaymentMethod(index, 'type', e.target.value)}
                            className="bg-zinc-50 border border-zinc-100 rounded-xl p-2.5 text-sm outline-none focus:ring-2 focus:ring-zinc-900 font-medium"
                          >
                            <option value="Efectivo">Efectivo</option>
                            <option value="Tarjeta">Tarjeta</option>
                            <option value="Transferencia">Transferencia</option>
                            <option value="Otro">Otro</option>
                          </select>
                          <input
                            type="number"
                            value={pm.amount}
                            onChange={(e) => updatePaymentMethod(index, 'amount', parseFloat(e.target.value) || 0)}
                            className="bg-zinc-50 border border-zinc-100 rounded-xl p-2.5 text-sm outline-none focus:ring-2 focus:ring-zinc-900 font-bold text-right"
                          />
                        </div>
                      </div>
                    ))}
                  </div>

                  {paymentMethods.length < 2 && (
                    <button
                      onClick={addPaymentMethod}
                      className="w-full py-3 border-2 border-dashed border-zinc-200 rounded-2xl text-zinc-400 font-bold text-sm hover:border-zinc-900 hover:text-zinc-900 transition-all flex items-center justify-center gap-2"
                    >
                      <Plus size={16} />
                      Dividir Pago
                    </button>
                  )}

                  <div className="pt-4 space-y-3">
                    <div className="flex justify-between items-center px-2">
                      <span className="text-xs font-bold text-zinc-400 uppercase tracking-widest">Total Ingresado</span>
                      <span className={`font-bold ${Math.abs(paymentMethods.reduce((s, p) => s + p.amount, 0) - cartTotal) < 0.01 ? 'text-green-600' : 'text-red-600'}`}>
                        ${paymentMethods.reduce((s, p) => s + p.amount, 0).toFixed(2)}
                      </span>
                    </div>
                    <button
                      onClick={processSale}
                      className="w-full bg-zinc-900 text-white py-4 rounded-2xl font-bold text-lg shadow-xl shadow-zinc-200 hover:scale-[1.02] active:scale-95 transition-all"
                    >
                      Confirmar y Finalizar
                    </button>
                  </div>
                </div>
              </motion.div>
            </div>
          )}
        </AnimatePresence>

        <AnimatePresence>
          {isReceiptModalOpen && lastSale && (
            <div className="fixed inset-0 bg-black/40 backdrop-blur-sm z-[200] flex items-center justify-center p-4 no-print">
              <motion.div
                initial={{ opacity: 0, scale: 0.9, y: 20 }}
                animate={{ opacity: 1, scale: 1, y: 0 }}
                exit={{ opacity: 0, scale: 0.9, y: 20 }}
                className="bg-white w-full max-w-md rounded-3xl shadow-2xl overflow-hidden"
              >
                <div className="p-8 text-center">
                  <div className="w-20 h-20 bg-green-100 text-green-600 rounded-full flex items-center justify-center mx-auto mb-6">
                    <CheckCircle2 size={40} />
                  </div>
                  <h3 className="text-2xl font-black text-zinc-900 mb-2">¡Venta Exitosa!</h3>
                  <p className="text-zinc-500 font-medium mb-1">La transacción se ha registrado con el folio <span className="text-zinc-900 font-bold">#{lastSale.folio}</span></p>
                  <p className="text-[10px] text-zinc-400 font-bold uppercase tracking-widest mb-8">
                    {new Date(lastSale.timestamp).toLocaleDateString()} — {new Date(lastSale.timestamp).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', hour12: true })}
                  </p>
                  
                  <div className="grid grid-cols-1 gap-3">
                    <button
                      onClick={() => printSaleReceipt(lastSale)}
                      className="w-full flex items-center justify-center gap-3 bg-zinc-900 text-white py-4 rounded-2xl font-bold hover:scale-[1.02] active:scale-95 transition-all shadow-lg shadow-zinc-200"
                    >
                      <Printer size={20} />
                      Imprimir Recibo
                    </button>
                    
                    <button
                      onClick={() => {
                        const client = clients.find(c => c.id === lastSale.clientId);
                        const phone = client?.phone || '';
                        const message = `Hola ${client?.name || 'cliente'}, gracias por tu compra en OmniPOS. Folio: #${lastSale.folio}. Total: $${lastSale.total.toFixed(2)}.`;
                        window.open(`https://wa.me/${phone.replace(/\D/g, '')}?text=${encodeURIComponent(message)}`, '_blank');
                      }}
                      className="w-full flex items-center justify-center gap-3 bg-white border-2 border-zinc-900 text-zinc-900 py-4 rounded-2xl font-bold hover:bg-zinc-50 active:scale-95 transition-all"
                    >
                      <MessageSquare size={20} />
                      Enviar por WhatsApp / SMS
                    </button>

                    <button
                      onClick={() => setIsReceiptModalOpen(false)}
                      className="w-full py-4 text-zinc-400 font-bold hover:text-zinc-900 transition-colors"
                    >
                      Cerrar
                    </button>
                  </div>
                </div>
              </motion.div>
            </div>
          )}
        </AnimatePresence>

        <AnimatePresence>
          {showClientAlert && (
            <div className="fixed inset-0 bg-black/40 backdrop-blur-sm z-[200] flex items-center justify-center p-4 no-print">
              <motion.div
                initial={{ opacity: 0, scale: 0.9, y: 20 }}
                animate={{ opacity: 1, scale: 1, y: 0 }}
                exit={{ opacity: 0, scale: 0.9, y: 20 }}
                className="bg-white w-full max-w-sm rounded-3xl shadow-2xl overflow-hidden p-8 text-center"
              >
                <div className="w-20 h-20 bg-amber-50 text-amber-500 rounded-full flex items-center justify-center mx-auto mb-6">
                  <User size={40} />
                </div>
                <h3 className="text-2xl font-black text-zinc-900 mb-2">Cliente Requerido</h3>
                <p className="text-zinc-500 font-medium mb-8">
                  Para continuar con la venta, es necesario seleccionar un cliente de la lista para asignar la compra.
                </p>
                <div className="flex flex-col gap-3">
                  <button
                    onClick={() => setShowClientAlert(false)}
                    className="w-full bg-zinc-900 text-white py-4 rounded-2xl font-bold hover:scale-[1.02] active:scale-95 transition-all shadow-lg shadow-zinc-200"
                  >
                    Seleccionar de la lista
                  </button>
                  <button
                    onClick={() => {
                      setShowClientAlert(false);
                      setCurrentView(View.CLIENTS);
                    }}
                    className="w-full bg-zinc-100 text-zinc-900 py-4 rounded-2xl font-bold hover:bg-zinc-200 transition-all"
                  >
                    Registrar Nuevo Cliente
                  </button>
                </div>
              </motion.div>
            </div>
          )}
        </AnimatePresence>

        {currentView === View.INVENTORY && <InventoryView products={products} categories={categories} brands={brands} onSetView={setCurrentView} />}
        {currentView === View.PRODUCTS && <SimpleProductsView products={products} categories={categories} brands={brands} onSetView={setCurrentView} />}
        {currentView === View.CLIENTS && <ClientsView clients={clients} settings={settings} />}
        {currentView === View.SALES_HISTORY && <SalesHistoryView sales={sales} clients={clients} onPrint={printSaleReceipt} onView={setSaleToPrint} />}
        {currentView === View.WEB_STORE && (
          <WebStoreView 
            products={products} 
            clients={clients} 
            onSaleComplete={(sale) => {
              setSales([sale, ...sales]);
              setSettings({ ...settings, nextFolio: settings.nextFolio + 1 });
              printSaleReceipt(sale);
            }} 
          />
        )}
        {currentView === View.DASHBOARD && <DashboardView sales={sales} products={products} clients={clients} />}
        {currentView === View.REPORTS && <ReportsView sales={sales} products={products} clients={clients} categories={categories} />}
        {currentView === View.SETTINGS && (
          <SettingsView 
            settings={settings} 
            categories={categories} 
            brands={brands} 
            onPreviewTicket={(mockSale, override) => {
              setSaleToPrint(mockSale);
              setPrintSettingsOverride(override);
            }}
          />
        )}
          </motion.div>
        </AnimatePresence>
      </main>

      {/* --- Ticket Visualization Overlay & Printable Area --- */}
      <AnimatePresence>
        {saleToPrint && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 z-[100] flex items-center justify-center p-4 md:p-8"
          >
            {/* Backdrop */}
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setSaleToPrint(null)}
              className="absolute inset-0 bg-zinc-900/60 backdrop-blur-sm no-print"
            />
            
            {/* Ticket Container */}
            <motion.div 
              initial={{ scale: 0.9, opacity: 0, y: 20 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              exit={{ scale: 0.9, opacity: 0, y: 20 }}
              className="relative w-full max-w-md bg-white rounded-[2.5rem] shadow-2xl overflow-hidden flex flex-col max-h-full no-print"
            >
              <div className="p-6 border-b border-zinc-100 flex items-center justify-between bg-zinc-50/50">
                <div>
                  <h3 className="text-xl font-black tracking-tighter uppercase">Vista Previa de Ticket</h3>
                  <p className="text-xs font-bold text-zinc-400">Folio: #{saleToPrint.folio}</p>
                </div>
                <button 
                  onClick={() => setSaleToPrint(null)}
                  className="p-2 text-zinc-400 hover:text-zinc-900 hover:bg-white rounded-xl transition-all shadow-sm"
                >
                  <X size={24} />
                </button>
              </div>

              <div className="flex-1 overflow-y-auto p-8 bg-zinc-100/30 flex justify-center custom-scrollbar">
                <Ticket sale={saleToPrint} clients={clients} settings={settings} isPreview={true} overrideSettings={printSettingsOverride || undefined} />
              </div>

              <div className="p-6 bg-white border-t border-zinc-100 flex gap-3">
                <button 
                  onClick={() => { setSaleToPrint(null); setPrintSettingsOverride(null); }}
                  className="flex-1 bg-zinc-100 text-zinc-900 px-6 py-3 rounded-2xl font-bold hover:bg-zinc-200 transition-all active:scale-95"
                >
                  Cerrar
                </button>
                <button 
                  onClick={() => window.print()}
                  className="flex-1 bg-zinc-900 text-white px-6 py-3 rounded-2xl font-bold hover:scale-105 active:scale-95 transition-all shadow-lg shadow-zinc-200 flex items-center justify-center gap-2"
                >
                  <Printer size={20} />
                  Imprimir
                </button>
              </div>
            </motion.div>

            {/* Hidden Printable Content (Optimized for Printer) */}
            <div className="print-only fixed inset-0 bg-white z-[9999]">
              <Ticket sale={saleToPrint} clients={clients} settings={settings} overrideSettings={printSettingsOverride || undefined} />
            </div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}

// --- Sub-Views ---

const SalesHistoryView = ({ sales, clients, onPrint, onView }: { sales: Sale[], clients: Client[], onPrint: (sale: Sale) => void, onView: (sale: Sale) => void }) => {
  const [filterType, setFilterType] = useState<'all' | 'daily' | 'monthly' | 'client' | 'range'>('all');
  const [selectedClientId, setSelectedClientId] = useState<string>('');
  const [startDate, setStartDate] = useState<string>('');
  const [endDate, setEndDate] = useState<string>('');
  const [selectedMonth, setSelectedMonth] = useState<string>(new Date().toISOString().slice(0, 7));

  const filteredSales = sales.filter(sale => {
    const saleDate = new Date(sale.timestamp);
    const today = new Date();
    
    if (filterType === 'daily') {
      return saleDate.toDateString() === today.toDateString();
    }
    if (filterType === 'monthly') {
      return sale.timestamp.startsWith(selectedMonth);
    }
    if (filterType === 'client') {
      return sale.clientId === selectedClientId;
    }
    if (filterType === 'range') {
      if (!startDate || !endDate) return true;
      const start = new Date(startDate);
      const end = new Date(endDate);
      end.setHours(23, 59, 59, 999);
      return saleDate >= start && saleDate <= end;
    }
    return true;
  }).sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());

    const exportToPDF = () => {
      const doc = new jsPDF() as any;
      
      doc.setFontSize(20);
      doc.text('Historial de Ventas - OmniPOS', 14, 22);
      
      doc.setFontSize(10);
      doc.setTextColor(100);
      doc.text(`Generado el: ${new Date().toLocaleString()}`, 14, 30);
      doc.text(`Filtro: ${filterType}`, 14, 35);
      
      const tableData = filteredSales.map(sale => [
        sale.folio ? `#${sale.folio}` : '-',
        new Date(sale.timestamp).toLocaleString(),
        clients.find(c => c.id === sale.clientId)?.name || 'Consumidor Final',
        (sale.paymentMethods || []).map(pm => `${pm.type}: $${pm.amount.toFixed(2)}`).join(', '),
        `$${sale.total.toFixed(2)}`
      ]);

      autoTable(doc, {
        startY: 45,
        head: [['Folio', 'Fecha', 'Cliente', 'Métodos de Pago', 'Total']],
        body: tableData,
        theme: 'striped',
        headStyles: { fillColor: [24, 24, 27] },
        styles: { fontSize: 8 }
      });

      const finalY = (doc as any).lastAutoTable.finalY || 45;
      const totalAmount = filteredSales.reduce((sum, s) => sum + s.total, 0);
      
      // Calculate totals by payment method
      const totalsByMethod: Record<string, number> = {};
      filteredSales.forEach(sale => {
        (sale.paymentMethods || []).forEach(pm => {
          totalsByMethod[pm.type] = (totalsByMethod[pm.type] || 0) + pm.amount;
        });
      });

      doc.setFontSize(12);
      doc.setTextColor(0);
      doc.setFont(undefined, 'bold');
      doc.text(`Resumen de Totales:`, 14, finalY + 15);
      
      doc.setFontSize(10);
      doc.setFont(undefined, 'normal');
      let currentY = finalY + 22;
      
      Object.entries(totalsByMethod).forEach(([method, amount]) => {
        doc.text(`${method}: $${amount.toFixed(2)}`, 14, currentY);
        currentY += 6;
      });
      
      doc.setFontSize(12);
      doc.setFont(undefined, 'bold');
      doc.text(`Total General: $${totalAmount.toFixed(2)}`, 14, currentY + 4);
      
      doc.save(`ventas_${new Date().getTime()}.pdf`);
    };

  return (
    <div className="flex-1 p-4 md:p-8 overflow-y-auto custom-scrollbar">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-6 mb-10">
        <div>
          <h2 className="text-3xl font-bold tracking-tight mb-1">Historial de Ventas</h2>
          <p className="text-zinc-500 font-medium">Consulta y exporta tus registros de transacciones</p>
        </div>
        <button 
          onClick={exportToPDF}
          disabled={filteredSales.length === 0}
          className="flex items-center justify-center gap-2 bg-zinc-900 text-white px-6 py-3 rounded-2xl font-bold hover:scale-105 active:scale-95 transition-all shadow-lg shadow-zinc-200 disabled:opacity-50 disabled:hover:scale-100"
        >
          <Download size={20} />
          Exportar PDF
        </button>
      </div>

      <div className="bg-white rounded-3xl border border-zinc-100 shadow-sm p-6 mb-8">
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
          <div className="space-y-2">
            <label className="text-[10px] font-bold text-zinc-400 uppercase tracking-widest flex items-center gap-2">
              <Filter size={12} /> Tipo de Filtro
            </label>
            <select 
              value={filterType}
              onChange={(e) => setFilterType(e.target.value as any)}
              className="w-full bg-zinc-50 border border-zinc-100 rounded-xl p-2.5 text-sm outline-none focus:ring-2 focus:ring-zinc-900 font-medium"
            >
              <option value="all">Todas las Ventas</option>
              <option value="daily">Ventas del Día</option>
              <option value="monthly">Ventas del Mes</option>
              <option value="client">Por Cliente</option>
              <option value="range">Rango de Fechas</option>
            </select>
          </div>

          {filterType === 'client' && (
            <div className="space-y-2">
              <label className="text-[10px] font-bold text-zinc-400 uppercase tracking-widest flex items-center gap-2">
                <Users size={12} /> Seleccionar Cliente
              </label>
              <select 
                value={selectedClientId}
                onChange={(e) => setSelectedClientId(e.target.value)}
                className="w-full bg-zinc-50 border border-zinc-100 rounded-xl p-2.5 text-sm outline-none focus:ring-2 focus:ring-zinc-900 font-medium"
              >
                <option value="">Seleccionar...</option>
                {clients.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
              </select>
            </div>
          )}

          {filterType === 'monthly' && (
            <div className="space-y-2">
              <label className="text-[10px] font-bold text-zinc-400 uppercase tracking-widest flex items-center gap-2">
                <Calendar size={12} /> Seleccionar Mes
              </label>
              <input 
                type="month"
                value={selectedMonth}
                onChange={(e) => setSelectedMonth(e.target.value)}
                className="w-full bg-zinc-50 border border-zinc-100 rounded-xl p-2.5 text-sm outline-none focus:ring-2 focus:ring-zinc-900 font-medium"
              />
            </div>
          )}

          {filterType === 'range' && (
            <>
              <div className="space-y-2">
                <label className="text-[10px] font-bold text-zinc-400 uppercase tracking-widest">Desde</label>
                <input 
                  type="date"
                  value={startDate}
                  onChange={(e) => setStartDate(e.target.value)}
                  className="w-full bg-zinc-50 border border-zinc-100 rounded-xl p-2.5 text-sm outline-none focus:ring-2 focus:ring-zinc-900 font-medium"
                />
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-bold text-zinc-400 uppercase tracking-widest">Hasta</label>
                <input 
                  type="date"
                  value={endDate}
                  onChange={(e) => setEndDate(e.target.value)}
                  className="w-full bg-zinc-50 border border-zinc-100 rounded-xl p-2.5 text-sm outline-none focus:ring-2 focus:ring-zinc-900 font-medium"
                />
              </div>
            </>
          )}
        </div>
      </div>

      <div className="bg-white rounded-3xl border border-zinc-100 shadow-sm overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="border-b border-zinc-50">
                <th className="p-6 text-xs font-bold text-zinc-400 uppercase tracking-widest">Folio</th>
                <th className="p-6 text-xs font-bold text-zinc-400 uppercase tracking-widest">Fecha / Hora</th>
                <th className="p-6 text-xs font-bold text-zinc-400 uppercase tracking-widest">Cliente</th>
                <th className="p-6 text-xs font-bold text-zinc-400 uppercase tracking-widest">Método de Pago</th>
                <th className="p-6 text-xs font-bold text-zinc-400 uppercase tracking-widest text-right">Total</th>
                <th className="p-6 text-xs font-bold text-zinc-400 uppercase tracking-widest text-right">Acciones</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-zinc-50">
              {filteredSales.map((sale) => (
                <tr key={sale.id} className="hover:bg-zinc-50/50 transition-colors group">
                  <td className="p-6">
                    <span className="bg-zinc-100 text-zinc-600 px-3 py-1 rounded-full text-xs font-black tracking-tighter">
                      #{sale.folio || '-'}
                    </span>
                  </td>
                  <td className="p-6">
                    <p className="font-bold text-zinc-900 text-sm">{new Date(sale.timestamp).toLocaleDateString()}</p>
                    <p className="text-xs text-zinc-400 font-medium">{new Date(sale.timestamp).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', hour12: true })}</p>
                  </td>
                  <td className="p-6">
                    <p className="font-bold text-zinc-900 text-sm">
                      {clients.find(c => c.id === sale.clientId)?.name || 'Consumidor Final'}
                    </p>
                  </td>
                  <td className="p-6">
                    <div className="flex flex-wrap gap-1">
                      {(sale.paymentMethods || []).map((pm, i) => (
                        <span key={i} className="bg-zinc-50 border border-zinc-100 text-zinc-500 px-2 py-0.5 rounded text-[10px] font-bold">
                          {pm.type}: ${pm.amount.toFixed(2)}
                        </span>
                      ))}
                    </div>
                  </td>
                  <td className="p-6 text-right">
                    <span className="text-lg font-black text-zinc-900">${sale.total.toFixed(2)}</span>
                  </td>
                  <td className="p-6 text-right">
                    <div className="flex justify-end gap-1">
                      <button
                        onClick={() => onView(sale)}
                        className="p-2 text-zinc-400 hover:text-zinc-900 hover:bg-zinc-100 rounded-xl transition-all"
                        title="Visualizar Ticket"
                      >
                        <Eye size={20} />
                      </button>
                      <button
                        onClick={() => onPrint(sale)}
                        className="p-2 text-zinc-400 hover:text-zinc-900 hover:bg-zinc-100 rounded-xl transition-all"
                        title="Reimprimir Ticket"
                      >
                        <Printer size={20} />
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
              {filteredSales.length === 0 && (
                <tr>
                  <td colSpan={6} className="p-12 text-center text-zinc-400 font-bold">No se encontraron ventas con los filtros seleccionados</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

const SimpleProductsView = ({ products, categories, brands, onSetView }: { products: Product[], categories: Category[], brands: Brand[], onSetView: (view: View) => void }) => {
  const [searchQuery, setSearchQuery] = useState('');
  const [isAdding, setIsAdding] = useState(false);
  const [editingProduct, setEditingProduct] = useState<Product | null>(null);
  const [newProduct, setNewProduct] = useState<Partial<Product>>({
    name: '',
    price: 0,
    purchasePrice: 0,
    supplier: '',
    brand: brands[0]?.name || '',
    category: categories[0]?.name || 'General',
    stock: 0,
    imageUrl: ''
  });

  const handleAdd = async (e: React.FormEvent) => {
    e.preventDefault();
    try {
      await addDoc(collection(db, 'products'), {
        ...newProduct,
        createdAt: new Date().toISOString()
      });
      setIsAdding(false);
      setNewProduct({ name: '', price: 0, purchasePrice: 0, supplier: '', brand: brands[0]?.name || '', category: categories[0]?.name || 'General', stock: 0, imageUrl: '' });
      alert('¡Producto guardado exitosamente!');
    } catch (error) {
      alert('Error al guardar el producto: ' + (error instanceof Error ? error.message : 'Error desconocido'));
      handleFirestoreError(error, OperationType.WRITE, 'products');
    }
  };

  const handleUpdate = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingProduct?.id) return;
    try {
      const { id, ...data } = editingProduct;
      await updateDoc(doc(db, 'products', id), data);
      setEditingProduct(null);
      alert('¡Producto actualizado exitosamente!');
    } catch (error) {
      alert('Error al actualizar el producto: ' + (error instanceof Error ? error.message : 'Error desconocido'));
      handleFirestoreError(error, OperationType.WRITE, 'products');
    }
  };

  const handleDelete = async (id: string) => {
    if (!confirm('¿Estás seguro de que quieres eliminar este producto?')) return;
    try {
      await deleteDoc(doc(db, 'products', id));
      alert('¡Producto eliminado exitosamente!');
    } catch (error) {
      alert('Error al eliminar el producto: ' + (error instanceof Error ? error.message : 'Error desconocido'));
      handleFirestoreError(error, OperationType.DELETE, 'products');
    }
  };

  return (
    <div className="flex-1 p-3 md:p-8 flex flex-col h-full overflow-hidden">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-3 mb-4 shrink-0">
        <div>
          <h2 className="text-2xl md:text-3xl font-bold tracking-tight mb-0.5">Lista de Productos</h2>
          <p className="text-zinc-500 font-medium text-xs md:text-base">Vista rápida de precios y disponibilidad</p>
        </div>
        <div className="flex flex-col sm:flex-row gap-2">
          <div className="relative">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-zinc-400" size={16} />
            <input 
              type="text"
              placeholder="Buscar producto..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="pl-9 pr-4 py-2 md:py-3 bg-white border border-zinc-200 rounded-xl w-full sm:w-64 outline-none focus:ring-2 focus:ring-zinc-900 shadow-sm transition-all text-sm font-medium"
            />
          </div>
          <button 
            onClick={() => window.print()}
            className="flex items-center justify-center gap-2 bg-zinc-900 text-white px-4 py-2 md:px-6 md:py-3 rounded-xl md:rounded-2xl font-bold hover:scale-105 active:scale-95 transition-all shadow-lg shadow-zinc-200 text-sm"
          >
            <Printer size={16} />
            Imprimir
          </button>
          <button 
            onClick={() => setIsAdding(true)}
            className="flex items-center justify-center gap-2 bg-zinc-900 text-white px-4 py-2 md:px-6 md:py-3 rounded-xl md:rounded-2xl font-bold hover:scale-105 active:scale-95 transition-all shadow-lg shadow-zinc-200 text-sm"
          >
            + Agregar
          </button>
        </div>
      </div>

      <div className="bg-white rounded-2xl md:rounded-3xl border border-zinc-100 shadow-sm overflow-hidden flex-1 flex flex-col min-h-0">
        <div className="divide-y divide-zinc-50 overflow-y-auto custom-scrollbar flex-1">
          {products
            .filter(p => p.name.toLowerCase().includes(searchQuery.toLowerCase()))
            .map((product) => (
            <div key={product.id} className="p-4 md:p-6 flex flex-col md:flex-row md:items-center justify-between hover:bg-zinc-50/50 transition-colors gap-3">
              <div className="flex flex-col">
                <span className={`font-bold text-base md:text-lg ${product.stock <= 0 ? 'text-red-600' : 'text-zinc-900'}`}>
                  {product.name}
                </span>
                <div className="flex items-center gap-2 mt-0.5">
                  <span className="text-[10px] font-bold px-1.5 py-0.5 bg-zinc-100 text-zinc-500 rounded-full capitalize">
                    {product.category}
                  </span>
                  {product.brand && (
                    <span className="text-[10px] font-bold px-1.5 py-0.5 bg-zinc-100 text-zinc-500 rounded-full capitalize">
                      {product.brand}
                    </span>
                  )}
                  <span className={`text-[10px] font-bold ${product.stock <= 5 ? 'text-amber-600' : 'text-zinc-400'}`}>
                    Stock: {product.stock}
                  </span>
                </div>
              </div>
              <div className="flex items-center justify-between md:justify-end gap-4 md:gap-6">
                <span className="font-black text-xl md:text-2xl text-zinc-900">
                  ${product.price.toFixed(2)}
                </span>
                <div className="flex items-center gap-1 md:gap-2">
                  <button 
                    onClick={() => setEditingProduct(product)}
                    className="p-2 md:p-2.5 text-zinc-400 hover:text-zinc-900 hover:bg-zinc-100 rounded-lg md:rounded-xl transition-all"
                    title="Editar producto"
                  >
                    <Edit size={18} />
                  </button>
                  <button 
                    onClick={() => handleDelete(product.id!)}
                    className="p-2 md:p-2.5 text-zinc-400 hover:text-red-600 hover:bg-red-50 rounded-lg md:rounded-xl transition-all"
                    title="Eliminar producto"
                  >
                    <Trash2 size={18} />
                  </button>
                </div>
              </div>
            </div>
          ))}
          {products.filter(p => p.name.toLowerCase().includes(searchQuery.toLowerCase())).length === 0 && (
            <div className="p-8 md:p-12 text-center text-zinc-400 font-bold">No hay productos que coincidan</div>
          )}
        </div>
      </div>

      <AnimatePresence>
        {(isAdding || editingProduct) && (
          <div className="fixed inset-0 bg-black/40 backdrop-blur-sm z-[100] flex items-center justify-center p-2 md:p-4">
            <motion.div 
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              className="bg-white w-full max-w-lg rounded-2xl md:rounded-3xl shadow-2xl overflow-hidden max-h-[95vh] flex flex-col"
            >
              <div className="p-4 md:p-6 border-b border-zinc-100 flex justify-between items-center shrink-0">
                <h3 className="text-lg md:text-xl font-bold">{editingProduct ? 'Editar Producto' : 'Nuevo Producto'}</h3>
                <button onClick={() => { setIsAdding(false); setEditingProduct(null); }} className="p-2 hover:bg-zinc-100 rounded-xl transition-all"><X size={20} /></button>
              </div>
              <form onSubmit={editingProduct ? handleUpdate : handleAdd} className="p-4 md:p-6 space-y-3 md:space-y-4 overflow-y-auto custom-scrollbar">
                <div className="space-y-2">
                  <label className="text-xs font-bold text-zinc-400 uppercase tracking-widest">Nombre del Producto</label>
                  <input 
                    required
                    type="text" 
                    value={editingProduct ? editingProduct.name : newProduct.name}
                    onChange={e => editingProduct 
                      ? setEditingProduct({...editingProduct, name: e.target.value})
                      : setNewProduct({...newProduct, name: e.target.value})
                    }
                    className="w-full bg-zinc-50 border border-zinc-100 rounded-xl p-3 outline-none focus:ring-2 focus:ring-zinc-900 transition-all font-medium" 
                  />
                </div>
                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-2">
                    <label className="text-xs font-bold text-zinc-400 uppercase tracking-widest">Precio Compra ($)</label>
                    <input 
                      type="number" 
                      step="0.01"
                      value={editingProduct ? (editingProduct.purchasePrice || 0) : (newProduct.purchasePrice || 0)}
                      onChange={e => editingProduct
                        ? setEditingProduct({...editingProduct, purchasePrice: parseFloat(e.target.value)})
                        : setNewProduct({...newProduct, purchasePrice: parseFloat(e.target.value)})
                      }
                      className="w-full bg-zinc-50 border border-zinc-100 rounded-xl p-3 outline-none focus:ring-2 focus:ring-zinc-900 transition-all font-medium" 
                    />
                  </div>
                  <div className="space-y-2">
                    <label className="text-xs font-bold text-zinc-400 uppercase tracking-widest">Precio Venta ($)</label>
                    <input 
                      required
                      type="number" 
                      step="0.01"
                      value={editingProduct ? editingProduct.price : newProduct.price}
                      onChange={e => editingProduct
                        ? setEditingProduct({...editingProduct, price: parseFloat(e.target.value)})
                        : setNewProduct({...newProduct, price: parseFloat(e.target.value)})
                      }
                      className="w-full bg-zinc-50 border border-zinc-100 rounded-xl p-3 outline-none focus:ring-2 focus:ring-zinc-900 transition-all font-medium" 
                    />
                  </div>
                </div>
                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-2">
                    <label className="text-xs font-bold text-zinc-400 uppercase tracking-widest">Proveedor</label>
                    <input 
                      type="text" 
                      value={editingProduct ? (editingProduct.supplier || '') : (newProduct.supplier || '')}
                      onChange={e => editingProduct
                        ? setEditingProduct({...editingProduct, supplier: e.target.value})
                        : setNewProduct({...newProduct, supplier: e.target.value})
                      }
                      placeholder="Nombre del proveedor"
                      className="w-full bg-zinc-50 border border-zinc-100 rounded-xl p-3 outline-none focus:ring-2 focus:ring-zinc-900 transition-all font-medium" 
                    />
                  </div>
                  <div className="space-y-2">
                    <label className="text-xs font-bold text-zinc-400 uppercase tracking-widest">Stock</label>
                    <input 
                      required
                      type="number" 
                      value={editingProduct ? editingProduct.stock : newProduct.stock}
                      onChange={e => editingProduct
                        ? setEditingProduct({...editingProduct, stock: parseInt(e.target.value)})
                        : setNewProduct({...newProduct, stock: parseInt(e.target.value)})
                      }
                      className="w-full bg-zinc-50 border border-zinc-100 rounded-xl p-3 outline-none focus:ring-2 focus:ring-zinc-900 transition-all font-medium" 
                    />
                  </div>
                </div>
                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-2">
                    <div className="flex justify-between items-center px-1">
                      <label className="text-xs font-bold text-zinc-400 uppercase tracking-widest">Categoría</label>
                      <button 
                        type="button"
                        onClick={() => { onSetView(View.SETTINGS); setIsAdding(false); setEditingProduct(null); }}
                        className="text-[10px] font-bold text-zinc-900 bg-zinc-100 px-2 py-0.5 rounded-md hover:bg-zinc-200 transition-all underline underline-offset-2"
                      >
                        Gestionar
                      </button>
                    </div>
                    <select 
                      value={editingProduct ? editingProduct.category : newProduct.category}
                      onChange={e => editingProduct
                        ? setEditingProduct({...editingProduct, category: e.target.value})
                        : setNewProduct({...newProduct, category: e.target.value})
                      }
                      className="w-full bg-zinc-50 border border-zinc-100 rounded-xl p-3 outline-none focus:ring-2 focus:ring-zinc-900 transition-all font-medium"
                    >
                      {categories.map(cat => (
                        <option key={cat.id} value={cat.name}>{cat.name}</option>
                      ))}
                      {categories.length === 0 && <option value="General">General</option>}
                    </select>
                  </div>
                  <div className="space-y-2">
                    <div className="flex justify-between items-center px-1">
                      <label className="text-xs font-bold text-zinc-400 uppercase tracking-widest">Marca</label>
                      <button 
                        type="button"
                        onClick={() => { onSetView(View.SETTINGS); setIsAdding(false); setEditingProduct(null); }}
                        className="text-[10px] font-bold text-zinc-900 bg-zinc-100 px-2 py-0.5 rounded-md hover:bg-zinc-200 transition-all underline underline-offset-2"
                      >
                        Gestionar
                      </button>
                    </div>
                    <select 
                      value={editingProduct ? (editingProduct.brand || '') : (newProduct.brand || '')}
                      onChange={e => editingProduct
                        ? setEditingProduct({...editingProduct, brand: e.target.value})
                        : setNewProduct({...newProduct, brand: e.target.value})
                      }
                      className="w-full bg-zinc-50 border border-zinc-100 rounded-xl p-3 outline-none focus:ring-2 focus:ring-zinc-900 transition-all font-medium"
                    >
                      <option value="">Sin marca</option>
                      {brands.map(brand => (
                        <option key={brand.id} value={brand.name}>{brand.name}</option>
                      ))}
                    </select>
                  </div>
                </div>
                <button type="submit" className="w-full bg-zinc-900 text-white py-4 rounded-2xl font-bold text-lg mt-4 shadow-xl shadow-zinc-200 hover:scale-[1.02] active:scale-95 transition-all">
                  {editingProduct ? 'Actualizar Producto' : 'Guardar Producto'}
                </button>
              </form>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
};

const ClientsView = ({ clients, settings }: { clients: Client[], settings: AppSettings }) => {
  const [isAdding, setIsAdding] = useState(false);
  const [editingClient, setEditingClient] = useState<Client | null>(null);
  const [newClient, setNewClient] = useState<Partial<Client>>({
    name: '',
    email: '',
    phone: '',
    address: '',
    salonName: '',
    password: ''
  });

  const handleAdd = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newClient.name?.trim()) return;

    // Normalize values for comparison
    const normalizedName = newClient.name.trim().toLowerCase();
    const normalizedEmail = newClient.email?.trim().toLowerCase();
    const normalizedPhone = newClient.phone?.trim().replace(/\D/g, '');

    const nameExists = clients.some(c => c.name.toLowerCase() === normalizedName);
    if (nameExists) {
      alert('Ya existe un cliente con este nombre.');
      return;
    }

    if (normalizedEmail) {
      const emailExists = clients.some(c => c.email?.toLowerCase() === normalizedEmail);
      if (emailExists) {
        alert('Ya existe un cliente con este correo electrónico.');
        return;
      }
    }

    if (normalizedPhone) {
      const phoneExists = clients.some(c => c.phone?.replace(/\D/g, '') === normalizedPhone);
      if (phoneExists) {
        alert('Ya existe un cliente con este número de teléfono.');
        return;
      }
    }

    try {
      const currentClientNumber = settings.nextClientNumber || 1;
      await addDoc(collection(db, 'clients'), {
        ...newClient,
        name: newClient.name.trim(),
        clientNumber: currentClientNumber,
        createdAt: new Date().toISOString()
      });
      
      // Update next client number
      await updateDoc(doc(db, 'settings', 'global'), {
        nextClientNumber: currentClientNumber + 1
      });

      // Close modal and show list
      setIsAdding(false);
      setNewClient({ name: '', email: '', phone: '', address: '', salonName: '', password: '' });
      alert('¡Cliente guardado exitosamente!');
    } catch (error) {
      alert('Error al guardar el cliente: ' + (error instanceof Error ? error.message : 'Error desconocido'));
      handleFirestoreError(error, OperationType.WRITE, 'clients');
    }
  };

  const handleUpdate = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingClient?.id || !editingClient.name.trim()) return;

    // Normalize values for comparison
    const normalizedName = editingClient.name.trim().toLowerCase();
    const normalizedEmail = editingClient.email?.trim().toLowerCase();
    const normalizedPhone = editingClient.phone?.trim().replace(/\D/g, '');

    const nameExists = clients.some(c => 
      c.name.toLowerCase() === normalizedName && c.id !== editingClient.id
    );
    if (nameExists) {
      alert('Ya existe otro cliente con este nombre.');
      return;
    }

    if (normalizedEmail) {
      const emailExists = clients.some(c => 
        c.email?.toLowerCase() === normalizedEmail && c.id !== editingClient.id
      );
      if (emailExists) {
        alert('Ya existe otro cliente con este correo electrónico.');
        return;
      }
    }

    if (normalizedPhone) {
      const phoneExists = clients.some(c => 
        c.phone?.replace(/\D/g, '') === normalizedPhone && c.id !== editingClient.id
      );
      if (phoneExists) {
        alert('Ya existe otro cliente con este número de teléfono.');
        return;
      }
    }

    try {
      const { id, ...data } = editingClient;
      await updateDoc(doc(db, 'clients', id), {
        ...data,
        name: data.name.trim()
      });
      // Close modal and show updated list
      setEditingClient(null);
      setIsAdding(false);
      alert('¡Cliente actualizado exitosamente!');
    } catch (error) {
      alert('Error al actualizar el cliente: ' + (error instanceof Error ? error.message : 'Error desconocido'));
      handleFirestoreError(error, OperationType.WRITE, 'clients');
    }
  };

  const handleDelete = async (id: string) => {
    if (!confirm('¿Estás seguro de que quieres eliminar este cliente?')) return;
    try {
      await deleteDoc(doc(db, 'clients', id));
    } catch (error) {
      handleFirestoreError(error, OperationType.DELETE, 'clients');
    }
  };

  const handleDeleteAll = async () => {
    if (!confirm('¿Estás COMPLETAMENTE SEGURO de que quieres eliminar TODOS los clientes? Esta acción no se puede deshacer.')) return;
    
    try {
      const batch = writeBatch(db);
      clients.forEach((client) => {
        if (client.id) {
          batch.delete(doc(db, 'clients', client.id));
        }
      });
      
      // Reset client counter in settings
      batch.update(doc(db, 'settings', 'global'), {
        nextClientNumber: 1
      });
      
      await batch.commit();
      alert(`Se han eliminado ${clients.length} clientes correctamente.`);
    } catch (error) {
      handleFirestoreError(error, OperationType.DELETE, 'clients');
    }
  };

  return (
    <div className="flex-1 p-3 md:p-8 flex flex-col h-full overflow-hidden">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-3 mb-4 shrink-0">
        <div>
          <h2 className="text-2xl md:text-3xl font-bold tracking-tight mb-0.5">Clientes</h2>
          <p className="text-zinc-500 font-medium text-xs md:text-base">Gestiona tu directorio de clientes</p>
        </div>
        <div className="flex gap-2">
          {clients.length > 0 && (
            <button 
              onClick={handleDeleteAll}
              className="flex items-center justify-center gap-2 bg-red-50 text-red-600 px-4 py-2 rounded-xl font-bold hover:bg-red-100 transition-all text-sm"
            >
              <Trash2 size={18} />
              Vaciar
            </button>
          )}
          <button 
            onClick={() => window.print()}
            className="flex items-center justify-center gap-2 bg-zinc-900 text-white px-4 py-2 rounded-xl font-bold hover:scale-105 active:scale-95 transition-all shadow-lg shadow-zinc-200 text-sm"
          >
            <Printer size={16} />
            Imprimir
          </button>
          <button 
            onClick={() => setIsAdding(true)}
            className="flex items-center justify-center gap-2 bg-zinc-900 text-white px-4 py-2 rounded-xl font-bold hover:scale-105 active:scale-95 transition-all shadow-lg shadow-zinc-200 text-sm"
          >
            + Cliente
          </button>
        </div>
      </div>

      <div className="bg-white rounded-2xl md:rounded-3xl border border-zinc-100 shadow-sm overflow-hidden flex-1 flex flex-col min-h-0">
        <div className="overflow-x-auto flex-1 custom-scrollbar">
          <table className="w-full text-left border-collapse min-w-[600px]">
            <thead>
              <tr className="border-b border-zinc-50">
                <th className="p-6 text-xs font-bold text-zinc-400 uppercase tracking-widest"># Cliente</th>
                <th className="p-6 text-xs font-bold text-zinc-400 uppercase tracking-widest">Nombre / Salón</th>
                <th className="p-6 text-xs font-bold text-zinc-400 uppercase tracking-widest">Contacto</th>
                <th className="p-6 text-xs font-bold text-zinc-400 uppercase tracking-widest">Dirección</th>
                <th className="p-6 text-xs font-bold text-zinc-400 uppercase tracking-widest text-right">Acciones</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-zinc-50">
              {clients.map((client) => (
                <tr key={client.id} className="hover:bg-zinc-50/50 transition-colors group">
                  <td className="p-6">
                    <span className="bg-zinc-100 text-zinc-600 px-3 py-1 rounded-full text-xs font-black tracking-tighter">
                      #{client.clientNumber || '-'}
                    </span>
                  </td>
                  <td className="p-6">
                    <div className="flex items-center gap-3">
                      <div className="w-10 h-10 bg-zinc-100 rounded-full flex items-center justify-center text-zinc-500 font-bold">
                        {client.name.charAt(0).toUpperCase()}
                      </div>
                      <div>
                        <p className="font-bold text-zinc-900">{client.name}</p>
                        {client.salonName && <p className="text-xs text-zinc-500 font-medium">{client.salonName}</p>}
                      </div>
                    </div>
                  </td>
                  <td className="p-6">
                    <p className="text-sm font-medium text-zinc-900">{client.email || '-'}</p>
                    <p className="text-xs text-zinc-500 font-medium">{client.phone || '-'}</p>
                  </td>
                  <td className="p-6">
                    <p className="text-sm text-zinc-500 font-medium max-w-[200px] truncate">{client.address || '-'}</p>
                  </td>
                  <td className="p-6 text-right">
                    <div className="flex justify-end gap-2">
                      <button 
                        onClick={() => setEditingClient(client)}
                        className="p-2 text-zinc-300 hover:text-zinc-900 transition-colors"
                      >
                        <Edit size={18} />
                      </button>
                      <button 
                        onClick={() => handleDelete(client.id!)}
                        className="p-2 text-zinc-300 hover:text-red-500 transition-colors"
                      >
                        <Trash2 size={18} />
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
              {clients.length === 0 && (
                <tr>
                  <td colSpan={5} className="p-20 text-center text-zinc-400 font-bold">
                    No hay clientes registrados
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      <AnimatePresence>
        {(isAdding || editingClient) && (
          <div className="fixed inset-0 bg-black/20 backdrop-blur-sm z-[100] flex items-center justify-center p-2 md:p-4">
            <motion.div 
              initial={{ opacity: 0, scale: 0.9, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.9, y: 20 }}
              className="bg-white w-full max-w-md rounded-2xl md:rounded-3xl shadow-2xl overflow-hidden max-h-[95vh] flex flex-col"
            >
              <div className="p-4 md:p-6 border-b border-zinc-50 flex items-center justify-between shrink-0">
                <h3 className="text-lg md:text-xl font-bold">{editingClient ? 'Editar Cliente' : 'Nuevo Cliente'}</h3>
                <button onClick={() => { setIsAdding(false); setEditingClient(null); }} className="p-2 hover:bg-zinc-100 rounded-xl transition-colors">
                  <X size={20} />
                </button>
              </div>
              <form onSubmit={editingClient ? handleUpdate : handleAdd} className="p-4 md:p-6 space-y-3 md:space-y-4 overflow-y-auto custom-scrollbar">
                <div className="flex items-center justify-between mb-4">
                  <span className="text-xs font-bold text-zinc-400 uppercase tracking-widest">Número de Cliente</span>
                  <div className="w-12 h-12 bg-zinc-100 text-zinc-900 rounded-full flex items-center justify-center text-sm font-black shadow-inner border border-zinc-200/50">
                    #{editingClient ? editingClient.clientNumber : settings.nextClientNumber}
                  </div>
                </div>
                <div className="space-y-2">
                  <label className="text-xs font-bold text-zinc-400 uppercase tracking-widest">Nombre Completo</label>
                  <input 
                    required
                    type="text" 
                    value={editingClient ? editingClient.name : newClient.name}
                    onChange={e => editingClient 
                      ? setEditingClient({...editingClient, name: e.target.value})
                      : setNewClient({...newClient, name: e.target.value})
                    }
                    className="w-full bg-zinc-50 border border-zinc-100 rounded-xl p-3 outline-none focus:ring-2 focus:ring-zinc-900 transition-all font-medium" 
                  />
                </div>
                <div className="space-y-2">
                  <label className="text-xs font-bold text-zinc-400 uppercase tracking-widest">Nombre del Salón (Opcional)</label>
                  <input 
                    type="text" 
                    value={editingClient ? editingClient.salonName : newClient.salonName}
                    onChange={e => editingClient
                      ? setEditingClient({...editingClient, salonName: e.target.value})
                      : setNewClient({...newClient, salonName: e.target.value})
                    }
                    placeholder="Ej. Nails & Beauty"
                    className="w-full bg-zinc-50 border border-zinc-100 rounded-xl p-3 outline-none focus:ring-2 focus:ring-zinc-900 transition-all font-medium" 
                  />
                </div>
                <div className="space-y-2">
                  <label className="text-xs font-bold text-zinc-400 uppercase tracking-widest">Correo Electrónico</label>
                  <input 
                    type="email" 
                    value={editingClient ? editingClient.email : newClient.email}
                    onChange={e => editingClient
                      ? setEditingClient({...editingClient, email: e.target.value})
                      : setNewClient({...newClient, email: e.target.value})
                    }
                    className="w-full bg-zinc-50 border border-zinc-100 rounded-xl p-3 outline-none focus:ring-2 focus:ring-zinc-900 transition-all font-medium" 
                  />
                </div>
                <div className="space-y-2">
                  <label className="text-xs font-bold text-zinc-400 uppercase tracking-widest">Celular</label>
                  <input 
                    type="tel" 
                    value={editingClient ? editingClient.phone : newClient.phone}
                    onChange={e => editingClient
                      ? setEditingClient({...editingClient, phone: e.target.value})
                      : setNewClient({...newClient, phone: e.target.value})
                    }
                    className="w-full bg-zinc-50 border border-zinc-100 rounded-xl p-3 outline-none focus:ring-2 focus:ring-zinc-900 transition-all font-medium" 
                  />
                </div>
                <div className="space-y-2">
                  <label className="text-xs font-bold text-zinc-400 uppercase tracking-widest">Dirección</label>
                  <textarea 
                    value={editingClient ? editingClient.address : newClient.address}
                    onChange={e => editingClient
                      ? setEditingClient({...editingClient, address: e.target.value})
                      : setNewClient({...newClient, address: e.target.value})
                    }
                    className="w-full bg-zinc-50 border border-zinc-100 rounded-xl p-3 outline-none focus:ring-2 focus:ring-zinc-900 transition-all font-medium h-24 resize-none" 
                  />
                </div>
                <div className="space-y-2">
                  <label className="text-xs font-bold text-zinc-400 uppercase tracking-widest">Contraseña Tienda Web</label>
                  <input 
                    type="text" 
                    value={editingClient ? (editingClient.password || '') : (newClient.password || '')}
                    onChange={e => editingClient
                      ? setEditingClient({...editingClient, password: e.target.value})
                      : setNewClient({...newClient, password: e.target.value})
                    }
                    placeholder="Contraseña para acceso web"
                    className="w-full bg-zinc-50 border border-zinc-100 rounded-xl p-3 outline-none focus:ring-2 focus:ring-zinc-900 transition-all font-medium" 
                  />
                </div>
                <button type="submit" className="w-full bg-zinc-900 text-white py-4 rounded-2xl font-bold text-lg mt-4 shadow-xl shadow-zinc-200 hover:scale-[1.02] active:scale-95 transition-all">
                  {editingClient ? 'Actualizar Cliente' : 'Guardar Cliente'}
                </button>
              </form>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
};

const InventoryView = ({ products, categories, brands, onSetView }: { products: Product[], categories: Category[], brands: Brand[], onSetView: (view: View) => void }) => {
  const [isAdding, setIsAdding] = useState(false);
  const [editingProduct, setEditingProduct] = useState<Product | null>(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [sortConfig, setSortConfig] = useState<{ key: keyof Product, direction: 'asc' | 'desc' } | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [newProduct, setNewProduct] = useState<Partial<Product>>({
    name: '',
    price: 0,
    purchasePrice: 0,
    supplier: '',
    brand: brands[0]?.name || '',
    category: categories[0]?.name || 'General',
    stock: 0,
    imageUrl: ''
  });

  const handleAdd = async (e: React.FormEvent) => {
    e.preventDefault();
    try {
      await addDoc(collection(db, 'products'), {
        ...newProduct,
        createdAt: new Date().toISOString()
      });
      setIsAdding(false);
      setNewProduct({ name: '', price: 0, purchasePrice: 0, supplier: '', brand: brands[0]?.name || '', category: categories[0]?.name || 'General', stock: 0, imageUrl: '' });
    } catch (error) {
      handleFirestoreError(error, OperationType.WRITE, 'products');
    }
  };

  const handleUpdate = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingProduct?.id) return;
    try {
      const { id, ...data } = editingProduct;
      await updateDoc(doc(db, 'products', id), data);
      setEditingProduct(null);
    } catch (error) {
      handleFirestoreError(error, OperationType.WRITE, 'products');
    }
  };

  const handleDelete = async (id: string) => {
    if (!confirm('¿Estás seguro de que quieres eliminar este producto?')) return;
    try {
      await deleteDoc(doc(db, 'products', id));
    } catch (error) {
      handleFirestoreError(error, OperationType.DELETE, 'products');
    }
  };

  const downloadTemplate = () => {
    const templateData = [
      {
        'Nombre': 'Ejemplo Producto',
        'Precio': 100.00,
        'Precio Compra': 50.00,
        'Proveedor': 'Proveedor SA',
        'Marca': 'Marca Ejemplo',
        'Categoría': 'General',
        'Stock': 10,
        'URL Imagen': 'https://link-a-imagen.com'
      }
    ];

    const ws = XLSX.utils.json_to_sheet(templateData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Plantilla Importacion");
    XLSX.writeFile(wb, "Plantilla_Inventario_OmniPOS.xlsx");
  };

  const handleExcelImport = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async (event) => {
      try {
        const bstr = event.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws) as any[];

        let importedCount = 0;
        let updatedCount = 0;

        for (const row of data) {
          const productName = row.Nombre || row.name || row.Producto || row.name;
          if (!productName) continue;

          const productData: any = {
            name: productName.toString(),
            price: Math.max(0.01, Number(row.Precio || row.price || 0.01)),
            category: row.Categoría || row.category || 'General',
            stock: Math.max(0, Number(row.Stock || row.stock || 0)),
            createdAt: new Date().toISOString()
          };

          const purchasePrice = Number(row['Precio Compra'] || row.purchasePrice || 0);
          if (purchasePrice > 0) productData.purchasePrice = purchasePrice;

          const supplier = row.Proveedor || row.supplier || '';
          if (supplier) productData.supplier = supplier;

          const brand = row.Marca || row.brand || '';
          if (brand) productData.brand = brand;

          const imageUrl = row['URL Imagen'] || row.imageUrl || '';
          if (imageUrl && (imageUrl.startsWith('http://') || imageUrl.startsWith('https://'))) {
            productData.imageUrl = imageUrl;
          }

          const existingProduct = products.find(p => p.name.toLowerCase() === productName.toString().toLowerCase());

          if (existingProduct?.id) {
            const { createdAt, ...updateData } = productData;
            await updateDoc(doc(db, 'products', existingProduct.id), updateData);
            updatedCount++;
          } else {
            await addDoc(collection(db, 'products'), productData);
            importedCount++;
          }
        }

        alert(`Importación finalizada: ${importedCount} nuevos, ${updatedCount} actualizados.`);
        if (fileInputRef.current) fileInputRef.current.value = '';
      } catch (error) {
        console.error("Error al importar Excel:", error);
        alert("Hubo un error al procesar el archivo Excel. Verifica el formato.");
      }
    };
    reader.readAsBinaryString(file);
  };

  const filteredProducts = products.filter(product => 
    product.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
    product.brand?.toLowerCase().includes(searchTerm.toLowerCase()) ||
    product.category.toLowerCase().includes(searchTerm.toLowerCase()) ||
    product.supplier?.toLowerCase().includes(searchTerm.toLowerCase())
  );

  const sortedProducts = [...filteredProducts].sort((a, b) => {
    if (!sortConfig) return 0;
    const { key, direction } = sortConfig;
    
    const aValue = a[key];
    const bValue = b[key];

    if (aValue === undefined || bValue === undefined) return 0;

    if (aValue < bValue) return direction === 'asc' ? -1 : 1;
    if (aValue > bValue) return direction === 'asc' ? 1 : -1;
    return 0;
  });

  const handleSort = (key: keyof Product) => {
    let direction: 'asc' | 'desc' = 'asc';
    if (sortConfig && sortConfig.key === key && sortConfig.direction === 'asc') {
      direction = 'desc';
    }
    setSortConfig({ key, direction });
  };

  const SortIcon = ({ columnKey }: { columnKey: keyof Product }) => {
    if (!sortConfig || sortConfig.key !== columnKey) return <ChevronsUpDown size={14} className="text-zinc-300" />;
    return sortConfig.direction === 'asc' ? <ChevronUp size={14} className="text-zinc-900" /> : <ChevronDown size={14} className="text-zinc-900" />;
  };

  return (
    <div className="flex-1 p-3 md:p-8 flex flex-col h-full overflow-hidden">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-3 mb-4 shrink-0">
        <div>
          <h2 className="text-2xl md:text-3xl font-bold tracking-tight mb-0.5">Inventario</h2>
          <p className="text-zinc-500 font-medium text-xs md:text-base">Gestiona tus productos y niveles de stock</p>
        </div>
        <div className="flex flex-wrap gap-2">
          <input 
            type="file" 
            ref={fileInputRef} 
            onChange={handleExcelImport} 
            accept=".xlsx, .xls, .csv" 
            className="hidden" 
          />
          <button 
            onClick={downloadTemplate}
            className="flex items-center justify-center gap-2 bg-zinc-100 text-zinc-600 px-4 py-2 rounded-xl font-bold hover:bg-zinc-200 active:scale-95 transition-all text-sm"
          >
            <FileText size={16} />
            Plantilla
          </button>
          <button 
            onClick={() => fileInputRef.current?.click()}
            className="flex items-center justify-center gap-2 bg-emerald-50 text-emerald-600 px-4 py-2 rounded-xl font-bold hover:bg-emerald-100 active:scale-95 transition-all text-sm"
          >
            <Download size={16} />
            Importar
          </button>
          <button 
            onClick={() => window.print()}
            className="flex items-center justify-center gap-2 bg-zinc-900 text-white px-4 py-2 rounded-xl font-bold hover:scale-105 active:scale-95 transition-all shadow-lg shadow-zinc-200 text-sm"
          >
            <Printer size={16} />
            Imprimir
          </button>
          <button 
            onClick={() => setIsAdding(true)}
            className="flex items-center justify-center gap-2 bg-zinc-900 text-white px-4 py-2 rounded-xl font-bold hover:scale-105 active:scale-95 transition-all shadow-lg shadow-zinc-200 text-sm"
          >
            + Producto
          </button>
        </div>
      </div>

      <div className="relative w-full md:w-96 group mb-4 shrink-0">
        <Search className="absolute left-4 top-1/2 -translate-y-1/2 text-zinc-400 group-focus-within:text-zinc-900 transition-colors" size={18} />
        <input 
          type="text"
          placeholder="Buscar..."
          value={searchTerm}
          onChange={(e) => setSearchTerm(e.target.value)}
          className="w-full bg-white border border-zinc-200 rounded-xl py-2 pl-12 pr-4 text-sm font-medium focus:outline-none focus:ring-2 focus:ring-zinc-900/5 focus:border-zinc-900 transition-all shadow-sm"
        />
      </div>

      <div className="bg-white rounded-2xl md:rounded-3xl border border-zinc-100 shadow-sm overflow-hidden flex-1 flex flex-col min-h-0">
        <div className="overflow-x-auto flex-1 custom-scrollbar pr-1">
          <table className="w-full text-left border-collapse min-w-[800px]">
            <thead>
              <tr className="bg-zinc-50/50 border-b border-zinc-100">
                <th 
                  className="px-6 py-4 text-xs font-bold text-zinc-400 uppercase tracking-widest cursor-pointer hover:bg-zinc-100/50 transition-colors"
                  onClick={() => handleSort('name')}
                >
                  <div className="flex items-center gap-2">
                    Producto <SortIcon columnKey="name" />
                  </div>
                </th>
                <th 
                  className="px-6 py-4 text-xs font-bold text-zinc-400 uppercase tracking-widest cursor-pointer hover:bg-zinc-100/50 transition-colors"
                  onClick={() => handleSort('brand')}
                >
                  <div className="flex items-center gap-2">
                    Marca <SortIcon columnKey="brand" />
                  </div>
                </th>
                <th 
                  className="px-6 py-4 text-xs font-bold text-zinc-400 uppercase tracking-widest cursor-pointer hover:bg-zinc-100/50 transition-colors"
                  onClick={() => handleSort('category')}
                >
                  <div className="flex items-center gap-2">
                    Categoría <SortIcon columnKey="category" />
                  </div>
                </th>
                <th 
                  className="px-6 py-4 text-xs font-bold text-zinc-400 uppercase tracking-widest cursor-pointer hover:bg-zinc-100/50 transition-colors"
                  onClick={() => handleSort('price')}
                >
                  <div className="flex items-center gap-2">
                    Precio <SortIcon columnKey="price" />
                  </div>
                </th>
                <th 
                  className="px-6 py-4 text-xs font-bold text-zinc-400 uppercase tracking-widest cursor-pointer hover:bg-zinc-100/50 transition-colors"
                  onClick={() => handleSort('supplier')}
                >
                  <div className="flex items-center gap-2">
                    Proveedor <SortIcon columnKey="supplier" />
                  </div>
                </th>
                <th 
                  className="px-6 py-4 text-xs font-bold text-zinc-400 uppercase tracking-widest cursor-pointer hover:bg-zinc-100/50 transition-colors"
                  onClick={() => handleSort('stock')}
                >
                  <div className="flex items-center gap-2">
                    Stock <SortIcon columnKey="stock" />
                  </div>
                </th>
                <th className="px-6 py-4 text-xs font-bold text-zinc-400 uppercase tracking-widest text-right">Acciones</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-zinc-50">
              {sortedProducts.map((product) => (
                <tr key={product.id} className="hover:bg-zinc-50/50 transition-colors group">
                  <td className="px-6 py-4">
                    <span className="font-bold text-zinc-900">{product.name}</span>
                  </td>
                  <td className="px-6 py-4">
                    <span className="text-zinc-500 font-medium">{product.brand || '-'}</span>
                  </td>
                  <td className="px-6 py-4">
                    <span className="bg-zinc-100 text-zinc-600 px-3 py-1 rounded-full text-xs font-bold">{product.category}</span>
                  </td>
                  <td className="px-6 py-4 font-bold text-zinc-900">${product.price.toFixed(2)}</td>
                  <td className="px-6 py-4 text-zinc-500 text-sm font-medium">{product.supplier || '-'}</td>
                  <td className="px-6 py-4">
                    <div className="flex items-center gap-2">
                      <div className={`w-2 h-2 rounded-full ${product.stock > 10 ? 'bg-green-500' : 'bg-amber-500'}`} />
                      <span className="font-bold">{product.stock} unidades</span>
                    </div>
                  </td>
                  <td className="px-6 py-4 text-right">
                    <div className="flex justify-end gap-2">
                      <button 
                        onClick={() => setEditingProduct(product)}
                        className="p-2 text-zinc-300 hover:text-zinc-900 transition-colors"
                      >
                        <Edit size={18} />
                      </button>
                      <button 
                        onClick={() => handleDelete(product.id!)}
                        className="p-2 text-zinc-300 hover:text-red-500 transition-colors"
                      >
                        <Trash2 size={18} />
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      <AnimatePresence>
        {(isAdding || editingProduct) && (
          <div className="fixed inset-0 bg-black/40 backdrop-blur-sm z-[100] flex items-center justify-center p-4">
            <motion.div 
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              className="bg-white w-full max-w-lg rounded-3xl shadow-2xl overflow-hidden"
            >
              <div className="p-6 border-b border-zinc-100 flex justify-between items-center">
                <h3 className="text-xl font-bold">{editingProduct ? 'Editar Producto' : 'Nuevo Producto'}</h3>
                <button onClick={() => { setIsAdding(false); setEditingProduct(null); }} className="p-2 hover:bg-zinc-100 rounded-xl transition-all"><X size={20} /></button>
              </div>
              <form onSubmit={editingProduct ? handleUpdate : handleAdd} className="p-6 space-y-4">
                <div className="space-y-2">
                  <label className="text-xs font-bold text-zinc-400 uppercase tracking-widest">Nombre del Producto</label>
                  <input 
                    required
                    type="text" 
                    value={editingProduct ? editingProduct.name : newProduct.name}
                    onChange={e => editingProduct 
                      ? setEditingProduct({...editingProduct, name: e.target.value})
                      : setNewProduct({...newProduct, name: e.target.value})
                    }
                    className="w-full bg-zinc-50 border border-zinc-100 rounded-xl p-3 outline-none focus:ring-2 focus:ring-zinc-900 transition-all font-medium" 
                  />
                </div>
                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-2">
                    <label className="text-xs font-bold text-zinc-400 uppercase tracking-widest">Precio Compra ($)</label>
                    <input 
                      type="number" 
                      step="0.01"
                      value={editingProduct ? (editingProduct.purchasePrice || 0) : (newProduct.purchasePrice || 0)}
                      onChange={e => editingProduct
                        ? setEditingProduct({...editingProduct, purchasePrice: parseFloat(e.target.value)})
                        : setNewProduct({...newProduct, purchasePrice: parseFloat(e.target.value)})
                      }
                      className="w-full bg-zinc-50 border border-zinc-100 rounded-xl p-3 outline-none focus:ring-2 focus:ring-zinc-900 transition-all font-medium" 
                    />
                  </div>
                  <div className="space-y-2">
                    <label className="text-xs font-bold text-zinc-400 uppercase tracking-widest">Precio Venta ($)</label>
                    <input 
                      required
                      type="number" 
                      step="0.01"
                      value={editingProduct ? editingProduct.price : newProduct.price}
                      onChange={e => editingProduct
                        ? setEditingProduct({...editingProduct, price: parseFloat(e.target.value)})
                        : setNewProduct({...newProduct, price: parseFloat(e.target.value)})
                      }
                      className="w-full bg-zinc-50 border border-zinc-100 rounded-xl p-3 outline-none focus:ring-2 focus:ring-zinc-900 transition-all font-medium" 
                    />
                  </div>
                </div>
                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-2">
                    <label className="text-xs font-bold text-zinc-400 uppercase tracking-widest">Proveedor</label>
                    <input 
                      type="text" 
                      value={editingProduct ? (editingProduct.supplier || '') : (newProduct.supplier || '')}
                      onChange={e => editingProduct
                        ? setEditingProduct({...editingProduct, supplier: e.target.value})
                        : setNewProduct({...newProduct, supplier: e.target.value})
                      }
                      placeholder="Nombre del proveedor"
                      className="w-full bg-zinc-50 border border-zinc-100 rounded-xl p-3 outline-none focus:ring-2 focus:ring-zinc-900 transition-all font-medium" 
                    />
                  </div>
                  <div className="space-y-2">
                    <label className="text-xs font-bold text-zinc-400 uppercase tracking-widest">Stock</label>
                    <input 
                      required
                      type="number" 
                      value={editingProduct ? editingProduct.stock : newProduct.stock}
                      onChange={e => editingProduct
                        ? setEditingProduct({...editingProduct, stock: parseInt(e.target.value)})
                        : setNewProduct({...newProduct, stock: parseInt(e.target.value)})
                      }
                      className="w-full bg-zinc-50 border border-zinc-100 rounded-xl p-3 outline-none focus:ring-2 focus:ring-zinc-900 transition-all font-medium" 
                    />
                  </div>
                </div>
                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-2">
                    <div className="flex justify-between items-center px-1">
                      <label className="text-xs font-bold text-zinc-400 uppercase tracking-widest">Categoría</label>
                      <button 
                        type="button"
                        onClick={() => { onSetView(View.SETTINGS); setIsAdding(false); setEditingProduct(null); }}
                        className="text-[10px] font-bold text-zinc-900 bg-zinc-100 px-2 py-0.5 rounded-md hover:bg-zinc-200 transition-all underline underline-offset-2"
                      >
                        Gestionar
                      </button>
                    </div>
                    <select 
                      value={editingProduct ? editingProduct.category : newProduct.category}
                      onChange={e => editingProduct
                        ? setEditingProduct({...editingProduct, category: e.target.value})
                        : setNewProduct({...newProduct, category: e.target.value})
                      }
                      className="w-full bg-zinc-50 border border-zinc-100 rounded-xl p-3 outline-none focus:ring-2 focus:ring-zinc-900 transition-all font-medium"
                    >
                      {categories.map(cat => (
                        <option key={cat.id} value={cat.name}>{cat.name}</option>
                      ))}
                      {categories.length === 0 && <option value="General">General</option>}
                    </select>
                  </div>
                  <div className="space-y-2">
                    <div className="flex justify-between items-center px-1">
                      <label className="text-xs font-bold text-zinc-400 uppercase tracking-widest">Marca</label>
                      <button 
                        type="button"
                        onClick={() => { onSetView(View.SETTINGS); setIsAdding(false); setEditingProduct(null); }}
                        className="text-[10px] font-bold text-zinc-900 bg-zinc-100 px-2 py-0.5 rounded-md hover:bg-zinc-200 transition-all underline underline-offset-2"
                      >
                        Gestionar
                      </button>
                    </div>
                    <select 
                      value={editingProduct ? (editingProduct.brand || '') : (newProduct.brand || '')}
                      onChange={e => editingProduct
                        ? setEditingProduct({...editingProduct, brand: e.target.value})
                        : setNewProduct({...newProduct, brand: e.target.value})
                      }
                      className="w-full bg-zinc-50 border border-zinc-100 rounded-xl p-3 outline-none focus:ring-2 focus:ring-zinc-900 transition-all font-medium"
                    >
                      <option value="">Sin marca</option>
                      {brands.map(brand => (
                        <option key={brand.id} value={brand.name}>{brand.name}</option>
                      ))}
                    </select>
                  </div>
                </div>
                <div className="space-y-2">
                  <label className="text-xs font-bold text-zinc-400 uppercase tracking-widest">URL de Imagen (Opcional)</label>
                  <input 
                    type="url" 
                    value={editingProduct ? editingProduct.imageUrl : newProduct.imageUrl}
                    onChange={e => editingProduct
                      ? setEditingProduct({...editingProduct, imageUrl: e.target.value})
                      : setNewProduct({...newProduct, imageUrl: e.target.value})
                    }
                    placeholder="https://images.unsplash.com/..."
                    className="w-full bg-zinc-50 border border-zinc-100 rounded-xl p-3 outline-none focus:ring-2 focus:ring-zinc-900 transition-all font-medium" 
                  />
                </div>
                <button type="submit" className="w-full bg-zinc-900 text-white py-4 rounded-2xl font-bold text-lg mt-4 shadow-xl shadow-zinc-200 hover:scale-[1.02] active:scale-95 transition-all">
                  {editingProduct ? 'Actualizar Producto' : 'Guardar Producto'}
                </button>
              </form>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
};

const SettingsView = ({ settings, categories, brands, onPreviewTicket }: { settings: AppSettings, categories: Category[], brands: Brand[], onPreviewTicket: (sale: Sale, ticketSettings: TicketSettings) => void }) => {
  const [nextFolio, setNextFolio] = useState(settings.nextFolio);
  const [nextClientNumber, setNextClientNumber] = useState(settings.nextClientNumber || 1);
  const [ticketSettings, setTicketSettings] = useState(settings.ticket || {
    companyName: 'OmniPOS',
    showLogo: false,
    showFiscalName: false,
    showAddress: false,
    showPhone: false,
    showEmail: false
  });
  const [newCategoryName, setNewCategoryName] = useState('');
  const [editingCategory, setEditingCategory] = useState<Category | null>(null);
  const [newBrandName, setNewBrandName] = useState('');
  const [editingBrand, setEditingBrand] = useState<Brand | null>(null);
  const [catSearch, setCatSearch] = useState('');
  const [brandSearch, setBrandSearch] = useState('');

  const filteredCategories = categories.filter(c => c.name.toLowerCase().includes(catSearch.toLowerCase()));
  const filteredBrands = brands.filter(b => b.name.toLowerCase().includes(brandSearch.toLowerCase()));

  const updateFolio = async () => {
    try {
      await updateDoc(doc(db, 'settings', 'global'), {
        nextFolio: Number(nextFolio)
      });
      alert('Folio actualizado con éxito');
    } catch (error) {
      handleFirestoreError(error, OperationType.WRITE, 'settings');
    }
  };

  const updateClientNumber = async () => {
    try {
      await updateDoc(doc(db, 'settings', 'global'), {
        nextClientNumber: Number(nextClientNumber)
      });
      alert('Número de cliente actualizado con éxito');
    } catch (error) {
      handleFirestoreError(error, OperationType.WRITE, 'settings');
    }
  };

  const updateTicketSettings = async () => {
    try {
      await updateDoc(doc(db, 'settings', 'global'), {
        ticket: ticketSettings
      });
      alert('Configuración de ticket actualizada con éxito');
    } catch (error) {
      handleFirestoreError(error, OperationType.WRITE, 'settings');
    }
  };

  const previewTicket = () => {
    const mockSale: Sale = {
      folio: 0,
      timestamp: new Date().toISOString(),
      items: [
        { productId: '1', name: 'PRODUCTO DE EJEMPLO 1', price: 25.50, quantity: 2 },
        { productId: '2', name: 'PRODUCTO DE EJEMPLO 2', price: 12.00, quantity: 1 }
      ],
      total: 63.00,
      userId: auth.currentUser?.uid || 'demo',
      paymentMethods: [{ type: 'Efectivo', amount: 63.00 }]
    };
    onPreviewTicket(mockSale, ticketSettings);
  };

  const handleLogoUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      if (file.size > 1024 * 1024) { // 1MB limit for Firestore doc size
        alert('La imagen es demasiado grande. Por favor usa una imagen de menos de 1MB.');
        return;
      }
      const reader = new FileReader();
      reader.onloadend = () => {
        setTicketSettings({ ...ticketSettings, logoUrl: reader.result as string });
      };
      reader.readAsDataURL(file);
    }
  };

  const handleAddCategory = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newCategoryName.trim()) return;
    try {
      await addDoc(collection(db, 'categories'), {
        name: newCategoryName.trim()
      });
      setNewCategoryName('');
    } catch (error) {
      handleFirestoreError(error, OperationType.WRITE, 'categories');
    }
  };

  const handleUpdateCategory = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingCategory || !editingCategory.name.trim()) return;
    try {
      await updateDoc(doc(db, 'categories', editingCategory.id), {
        name: editingCategory.name.trim()
      });
      setEditingCategory(null);
    } catch (error) {
      handleFirestoreError(error, OperationType.WRITE, 'categories');
    }
  };

  const handleDeleteCategory = async (id: string) => {
    if (!confirm('¿Estás seguro de que quieres eliminar esta categoría?')) return;
    try {
      await deleteDoc(doc(db, 'categories', id));
    } catch (error) {
      handleFirestoreError(error, OperationType.DELETE, 'categories');
    }
  };

  const handleAddBrand = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newBrandName.trim()) return;
    try {
      await addDoc(collection(db, 'brands'), {
        name: newBrandName.trim()
      });
      setNewBrandName('');
    } catch (error) {
      handleFirestoreError(error, OperationType.WRITE, 'brands');
    }
  };

  const handleUpdateBrand = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingBrand || !editingBrand.name.trim()) return;
    try {
      await updateDoc(doc(db, 'brands', editingBrand.id), {
        name: editingBrand.name.trim()
      });
      setEditingBrand(null);
    } catch (error) {
      handleFirestoreError(error, OperationType.WRITE, 'brands');
    }
  };

  const handleDeleteBrand = async (id: string) => {
    if (!confirm('¿Estás seguro de que quieres eliminar esta marca?')) return;
    try {
      await deleteDoc(doc(db, 'brands', id));
    } catch (error) {
      handleFirestoreError(error, OperationType.DELETE, 'brands');
    }
  };

  useEffect(() => {
    setNextFolio(settings.nextFolio);
    setNextClientNumber(settings.nextClientNumber || 1);
    if (settings.ticket) {
      setTicketSettings(settings.ticket);
    }
  }, [settings.nextFolio, settings.nextClientNumber, settings.ticket]);

  return (
    <div className="flex-1 p-4 md:p-8 overflow-y-auto custom-scrollbar">
      <div className="mb-10">
        <h2 className="text-3xl font-bold tracking-tight mb-1">Configuración</h2>
        <p className="text-zinc-500 font-medium">Personaliza tu experiencia y gestiona tu cuenta</p>
      </div>

      <div className="max-w-2xl space-y-6">
        <div className="bg-white p-8 rounded-3xl border border-zinc-100 shadow-sm">
          <h3 className="font-bold mb-6 flex items-center gap-2">
            <User size={20} />
            Perfil de Usuario
          </h3>
          <div className="space-y-4">
            <div className="flex items-center gap-4 p-4 bg-zinc-50 rounded-2xl">
              <div className="w-16 h-16 bg-zinc-200 rounded-full overflow-hidden">
                {auth.currentUser?.photoURL && <img src={auth.currentUser.photoURL} alt="" className="w-full h-full object-cover" referrerPolicy="no-referrer" />}
              </div>
              <div>
                <p className="font-bold text-lg">{auth.currentUser?.displayName || 'Usuario'}</p>
                <p className="text-zinc-500 font-medium">{auth.currentUser?.email}</p>
              </div>
            </div>
          </div>
        </div>

        <div className="bg-white p-8 rounded-3xl border border-zinc-100 shadow-sm">
          <h3 className="font-bold mb-6 flex items-center gap-2">
            <Globe size={20} />
            Tienda Web
          </h3>
          <div className="flex items-center justify-between p-4 bg-zinc-50 rounded-2xl">
            <div>
              <p className="font-bold">Habilitar Tienda Web</p>
              <p className="text-xs text-zinc-500">Muestra u oculta el acceso a la tienda web en el menú</p>
            </div>
            <button 
              onClick={async () => {
                try {
                  await updateDoc(doc(db, 'settings', 'global'), {
                    showWebStore: !settings.showWebStore
                  });
                } catch (error) {
                  handleFirestoreError(error, OperationType.WRITE, 'settings');
                }
              }}
              className={`w-14 h-8 rounded-full transition-all relative ${settings.showWebStore ? 'bg-zinc-900' : 'bg-zinc-200'}`}
            >
              <div className={`absolute top-1 w-6 h-6 bg-white rounded-full transition-all ${settings.showWebStore ? 'left-7' : 'left-1'}`} />
            </button>
          </div>
          {settings.showWebStore && (
            <div className="mt-4 p-4 bg-zinc-900 text-white rounded-2xl flex items-center justify-between">
              <div>
                <p className="text-xs font-bold text-zinc-400 uppercase tracking-widest mb-1">Enlace de la Tienda</p>
                <p className="font-mono text-xs truncate max-w-[200px]">{window.location.origin}/#tienda</p>
              </div>
              <button 
                onClick={() => window.open(`${window.location.origin}/#tienda`, '_blank')}
                className="bg-white text-zinc-900 px-4 py-2 rounded-xl font-bold text-xs hover:bg-zinc-100 transition-all"
              >
                Abrir Tienda
              </button>
            </div>
          )}
        </div>

        <div className="bg-white p-8 rounded-3xl border border-zinc-100 shadow-sm">
          <h3 className="font-bold mb-6 flex items-center gap-2">
            <Printer size={20} />
            Configuración de Ticket (80mm)
          </h3>
          <div className="space-y-6">
            <div className="p-4 bg-amber-50 border border-amber-100 rounded-2xl flex gap-3 text-amber-900 mb-2">
              <Tag size={20} className="shrink-0 mt-0.5" />
              <div className="text-xs">
                <p className="font-bold mb-1 uppercase tracking-wider">Personalización Visual</p>
                <p className="font-medium opacity-80 leading-relaxed">
                  El <strong>Nombre Comercial</strong> y el <strong>Logo</strong> configurados aquí se aplicarán automáticamente a tu ticket y a la interfaz del sistema.
                </p>
              </div>
            </div>

            <div className="flex items-center justify-between p-4 bg-zinc-50 rounded-2xl border border-zinc-100/50">
              <div className="flex items-center gap-3">
                <div className={`w-10 h-10 rounded-xl flex items-center justify-center transition-all ${ticketSettings.showLogo ? 'bg-amber-100 text-amber-600' : 'bg-zinc-100 text-zinc-400'}`}>
                  <ImageIcon size={20} />
                </div>
                <div>
                  <p className="font-bold text-sm">Mostrar Logo en Ticket</p>
                  <p className="text-[10px] text-zinc-500 font-medium">Habilita o deshabilita la imagen del logo en la impresión</p>
                </div>
              </div>
              <button 
                onClick={() => setTicketSettings({ ...ticketSettings, showLogo: !ticketSettings.showLogo })}
                className={`w-12 h-6 rounded-full transition-all relative ${ticketSettings.showLogo ? 'bg-amber-500' : 'bg-zinc-200'}`}
              >
                <div className={`absolute top-1 w-4 h-4 bg-white rounded-full transition-all ${ticketSettings.showLogo ? 'left-7' : 'left-1'}`} />
              </button>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div className="space-y-2">
                <label className="text-[10px] font-bold text-zinc-400 uppercase tracking-widest block ml-1">Nombre Comercial</label>
                <input 
                  type="text" 
                  value={ticketSettings.companyName}
                  onChange={(e) => setTicketSettings({ ...ticketSettings, companyName: e.target.value })}
                  placeholder="Ej. Mi Negocio"
                  className="w-full bg-zinc-50 border border-zinc-100 rounded-xl py-3 px-4 focus:ring-2 focus:ring-zinc-900 outline-none font-medium"
                />
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-bold text-zinc-400 uppercase tracking-widest block ml-1">Nombre Fiscal (Opcional)</label>
                <div className="flex gap-2">
                  <input 
                    type="text" 
                    value={ticketSettings.fiscalName || ''}
                    onChange={(e) => setTicketSettings({ ...ticketSettings, fiscalName: e.target.value })}
                    placeholder="Ej. Razón Social S.A."
                    className="flex-1 bg-zinc-50 border border-zinc-100 rounded-xl py-3 px-4 focus:ring-2 focus:ring-zinc-900 outline-none font-medium"
                  />
                  <button 
                    onClick={() => setTicketSettings({ ...ticketSettings, showFiscalName: !ticketSettings.showFiscalName })}
                    className={`px-3 rounded-xl border transition-all ${ticketSettings.showFiscalName ? 'bg-zinc-900 text-white border-zinc-900' : 'bg-white text-zinc-400 border-zinc-200'}`}
                  >
                    {ticketSettings.showFiscalName ? 'Ver' : 'Ocultar'}
                  </button>
                </div>
              </div>
            </div>

            <div className="space-y-2">
              <label className="text-[10px] font-bold text-zinc-400 uppercase tracking-widest block ml-1">Dirección</label>
              <div className="flex gap-2">
                <input 
                  type="text" 
                  value={ticketSettings.address || ''}
                  onChange={(e) => setTicketSettings({ ...ticketSettings, address: e.target.value })}
                  placeholder="Calle, Número, Colonia, Ciudad"
                  className="flex-1 bg-zinc-50 border border-zinc-100 rounded-xl py-3 px-4 focus:ring-2 focus:ring-zinc-900 outline-none font-medium"
                />
                <button 
                  onClick={() => setTicketSettings({ ...ticketSettings, showAddress: !ticketSettings.showAddress })}
                  className={`px-3 rounded-xl border transition-all ${ticketSettings.showAddress ? 'bg-zinc-900 text-white border-zinc-900' : 'bg-white text-zinc-400 border-zinc-200'}`}
                >
                  {ticketSettings.showAddress ? 'Ver' : 'Ocultar'}
                </button>
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div className="space-y-2">
                <label className="text-[10px] font-bold text-zinc-400 uppercase tracking-widest block ml-1">Celular</label>
                <div className="flex gap-2">
                  <input 
                    type="text" 
                    value={ticketSettings.phone || ''}
                    onChange={(e) => setTicketSettings({ ...ticketSettings, phone: e.target.value })}
                    placeholder="Ej. 123 456 7890"
                    className="flex-1 bg-zinc-50 border border-zinc-100 rounded-xl py-3 px-4 focus:ring-2 focus:ring-zinc-900 outline-none font-medium"
                  />
                  <button 
                    onClick={() => setTicketSettings({ ...ticketSettings, showPhone: !ticketSettings.showPhone })}
                    className={`px-3 rounded-xl border transition-all ${ticketSettings.showPhone ? 'bg-zinc-900 text-white border-zinc-900' : 'bg-white text-zinc-400 border-zinc-200'}`}
                  >
                    {ticketSettings.showPhone ? 'Ver' : 'Ocultar'}
                  </button>
                </div>
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-bold text-zinc-400 uppercase tracking-widest block ml-1">Correo Electrónico</label>
                <div className="flex gap-2">
                  <input 
                    type="email" 
                    value={ticketSettings.email || ''}
                    onChange={(e) => setTicketSettings({ ...ticketSettings, email: e.target.value })}
                    placeholder="contacto@empresa.com"
                    className="flex-1 bg-zinc-50 border border-zinc-100 rounded-xl py-3 px-4 focus:ring-2 focus:ring-zinc-900 outline-none font-medium"
                  />
                  <button 
                    onClick={() => setTicketSettings({ ...ticketSettings, showEmail: !ticketSettings.showEmail })}
                    className={`px-3 rounded-xl border transition-all ${ticketSettings.showEmail ? 'bg-zinc-900 text-white border-zinc-900' : 'bg-white text-zinc-400 border-zinc-200'}`}
                  >
                    {ticketSettings.showEmail ? 'Ver' : 'Ocultar'}
                  </button>
                </div>
              </div>
            </div>

            <div className="space-y-4">
              <label className="text-[10px] font-bold text-zinc-400 uppercase tracking-widest block ml-1">Logotipo del Negocio</label>
              <div className="flex flex-col md:flex-row gap-6 p-6 bg-zinc-50 rounded-2xl border border-zinc-100 items-start md:items-center">
                <div className="w-24 h-24 bg-white border-2 border-dashed border-zinc-200 rounded-2xl flex items-center justify-center overflow-hidden shrink-0 group relative">
                  {ticketSettings.logoUrl ? (
                    <>
                      <img src={ticketSettings.logoUrl} alt="Logo" className="w-full h-full object-contain p-2" />
                      <button 
                        onClick={() => setTicketSettings({ ...ticketSettings, logoUrl: '' })}
                        className="absolute inset-0 bg-black/40 opacity-0 group-hover:opacity-100 transition-opacity flex items-center justify-center text-white"
                      >
                        <Trash2 size={20} />
                      </button>
                    </>
                  ) : (
                    <ImageIcon className="text-zinc-300" size={32} />
                  )}
                </div>
                <div className="flex-1 space-y-3">
                  <div className="flex flex-wrap gap-2">
                    <label className="bg-zinc-900 text-white px-6 py-3 rounded-xl font-bold text-sm cursor-pointer hover:scale-105 transition-all inline-flex items-center gap-2">
                      <Upload size={18} />
                      Subir Logotipo
                      <input type="file" className="hidden" accept="image/*" onChange={handleLogoUpload} />
                    </label>
                    <div className={`px-4 py-3 rounded-xl border font-bold text-xs flex items-center gap-2 ${ticketSettings.showLogo ? 'bg-emerald-50 text-emerald-600 border-emerald-100' : 'bg-red-50 text-red-600 border-red-100'}`}>
                      <div className={`w-2 h-2 rounded-full ${ticketSettings.showLogo ? 'bg-emerald-500 animate-pulse' : 'bg-red-500'}`} />
                      {ticketSettings.showLogo ? 'Logo Activado' : 'Logo Desactivado'}
                    </div>
                  </div>
                  <p className="text-[10px] text-zinc-400 font-medium italic">Recomendado: Imagen cuadrada, blanco y negro (PNG/JPG). Máx: 1MB.</p>
                </div>
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div className="space-y-2">
                <label className="text-[10px] font-bold text-zinc-400 uppercase tracking-widest block ml-1">Mensaje Pie de Página 1</label>
                <input 
                  type="text" 
                  value={ticketSettings.footerMessage || ''}
                  onChange={(e) => setTicketSettings({ ...ticketSettings, footerMessage: e.target.value })}
                  placeholder="Ej. ¡Gracias por su compra!"
                  className="w-full bg-zinc-50 border border-zinc-100 rounded-xl py-3 px-4 focus:ring-2 focus:ring-zinc-900 outline-none font-medium"
                />
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-bold text-zinc-400 uppercase tracking-widest block ml-1">Mensaje Pie de Página 2</label>
                <input 
                  type="text" 
                  value={ticketSettings.secondaryFooterMessage || ''}
                  onChange={(e) => setTicketSettings({ ...ticketSettings, secondaryFooterMessage: e.target.value })}
                  placeholder="Ej. Vuelva pronto"
                  className="w-full bg-zinc-50 border border-zinc-100 rounded-xl py-3 px-4 focus:ring-2 focus:ring-zinc-900 outline-none font-medium"
                />
              </div>
            </div>

            <div className="flex flex-col sm:flex-row gap-3">
              <button 
                onClick={previewTicket}
                className="flex-1 bg-white border-2 border-zinc-900 text-zinc-900 py-4 rounded-2xl font-bold text-lg hover:bg-zinc-50 active:scale-95 transition-all flex items-center justify-center gap-3"
              >
                <Eye size={20} />
                Vista Previa
              </button>
              <button 
                onClick={updateTicketSettings}
                className="flex-[2] bg-zinc-900 text-white py-4 rounded-2xl font-bold text-lg shadow-xl shadow-zinc-200 hover:scale-[1.02] active:scale-95 transition-all flex items-center justify-center gap-3"
              >
                <Save size={20} />
                Guardar Configuración de Ticket
              </button>
            </div>
          </div>
        </div>

        <div className="bg-white p-8 rounded-3xl border border-zinc-100 shadow-sm">
          <h3 className="font-bold mb-6 flex items-center gap-2">
            <History size={20} />
            Configuración de Folios
          </h3>
          <div className="space-y-4">
            <div className="p-4 bg-zinc-50 rounded-2xl">
              <label className="block text-sm font-bold text-zinc-500 mb-2 uppercase tracking-widest">Siguiente Número de Folio</label>
              <div className="flex gap-3">
                <input 
                  type="number" 
                  value={nextFolio}
                  onChange={(e) => setNextFolio(Number(e.target.value))}
                  className="flex-1 bg-white border border-zinc-200 rounded-xl py-3 px-4 focus:ring-2 focus:ring-zinc-900 outline-none font-bold"
                />
                <button 
                  onClick={updateFolio}
                  className="bg-zinc-900 text-white px-6 py-3 rounded-xl font-bold hover:scale-105 active:scale-95 transition-all"
                >
                  Guardar
                </button>
              </div>
              <p className="mt-2 text-xs text-zinc-400 font-medium">Este número se asignará a la próxima venta y se incrementará automáticamente.</p>
            </div>

            <div className="p-4 bg-zinc-50 rounded-2xl">
              <label className="block text-sm font-bold text-zinc-500 mb-2 uppercase tracking-widest">Siguiente Número de Cliente</label>
              <div className="flex gap-3">
                <input 
                  type="number" 
                  value={nextClientNumber}
                  onChange={(e) => setNextClientNumber(Number(e.target.value))}
                  className="flex-1 bg-white border border-zinc-200 rounded-xl py-3 px-4 focus:ring-2 focus:ring-zinc-900 outline-none font-bold"
                />
                <button 
                  onClick={updateClientNumber}
                  className="bg-zinc-900 text-white px-6 py-3 rounded-xl font-bold hover:scale-105 active:scale-95 transition-all"
                >
                  Guardar
                </button>
              </div>
              <p className="mt-2 text-xs text-zinc-400 font-medium">Este número se asignará al próximo cliente registrado y se incrementará automáticamente.</p>
            </div>
          </div>
        </div>

        <div className="bg-white p-8 rounded-3xl border border-zinc-100 shadow-sm">
          <h3 className="font-bold mb-6 flex items-center gap-2">
            <LayoutGrid size={20} />
            Gestión de Categorías
          </h3>
          
          <div className="flex gap-3 mb-6">
            <div className="flex-1 relative group">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-zinc-400 group-focus-within:text-zinc-900 transition-colors" size={16} />
              <input 
                type="text" 
                placeholder="Filtrar categorías..."
                value={catSearch}
                onChange={(e) => setCatSearch(e.target.value)}
                className="w-full bg-zinc-50 border border-zinc-100 rounded-xl py-2 pl-10 pr-4 focus:ring-2 focus:ring-zinc-900 outline-none text-sm transition-all"
              />
            </div>
            <form onSubmit={handleAddCategory} className="flex gap-2">
              <input 
                type="text" 
                placeholder="Nueva..."
                value={newCategoryName}
                onChange={(e) => setNewCategoryName(e.target.value)}
                className="w-32 bg-zinc-50 border border-zinc-100 rounded-xl py-2 px-3 focus:ring-2 focus:ring-zinc-900 outline-none text-sm font-medium"
              />
              <button 
                type="submit"
                className="bg-zinc-900 text-white px-4 py-2 rounded-xl font-bold hover:scale-105 active:scale-95 transition-all flex items-center gap-2 text-sm shadow-sm"
              >
                <Plus size={16} />
                Añadir
              </button>
            </form>
          </div>

          <div className="space-y-3 max-h-[400px] overflow-y-auto pr-1 custom-scrollbar">
            {filteredCategories.map((category) => (
              <div key={category.id} className="flex items-center justify-between p-4 bg-zinc-50 rounded-2xl group">
                {editingCategory?.id === category.id ? (
                  <form onSubmit={handleUpdateCategory} className="flex-1 flex gap-2">
                    <input 
                      type="text" 
                      value={editingCategory.name}
                      onChange={(e) => setEditingCategory({ ...editingCategory, name: e.target.value })}
                      className="flex-1 bg-white border border-zinc-200 rounded-lg px-3 py-1 outline-none focus:ring-2 focus:ring-zinc-900 font-medium"
                      autoFocus
                    />
                    <button type="submit" className="text-green-600 font-bold text-sm px-2">Guardar</button>
                    <button type="button" onClick={() => setEditingCategory(null)} className="text-zinc-400 font-bold text-sm px-2">Cancelar</button>
                  </form>
                ) : (
                  <>
                    <span className="font-bold text-zinc-900">{category.name}</span>
                    <div className="flex gap-1">
                      <button 
                        onClick={() => setEditingCategory(category)}
                        className="p-2 text-zinc-500 hover:text-zinc-900 hover:bg-white rounded-lg transition-all"
                        title="Editar"
                      >
                        <Edit size={18} />
                      </button>
                      <button 
                        onClick={() => handleDeleteCategory(category.id)}
                        className="p-2 text-zinc-500 hover:text-red-600 hover:bg-red-50 rounded-lg transition-all"
                        title="Eliminar"
                      >
                        <Trash2 size={18} />
                      </button>
                    </div>
                  </>
                )}
              </div>
            ))}
            {categories.length === 0 && (
              <p className="text-center py-4 text-zinc-400 font-medium">No hay categorías registradas</p>
            )}
          </div>
        </div>

        <div className="bg-white p-8 rounded-3xl border border-zinc-100 shadow-sm">
          <h3 className="font-bold mb-6 flex items-center gap-2">
            <Tag size={20} />
            Gestión de Marcas
          </h3>
          
          <div className="flex gap-3 mb-6">
            <div className="flex-1 relative group">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-zinc-400 group-focus-within:text-zinc-900 transition-colors" size={16} />
              <input 
                type="text" 
                placeholder="Filtrar marcas..."
                value={brandSearch}
                onChange={(e) => setBrandSearch(e.target.value)}
                className="w-full bg-zinc-50 border border-zinc-100 rounded-xl py-2 pl-10 pr-4 focus:ring-2 focus:ring-zinc-900 outline-none text-sm transition-all"
              />
            </div>
            <form onSubmit={handleAddBrand} className="flex gap-2">
              <input 
                type="text" 
                placeholder="Nueva..."
                value={newBrandName}
                onChange={(e) => setNewBrandName(e.target.value)}
                className="w-32 bg-zinc-50 border border-zinc-100 rounded-xl py-2 px-3 focus:ring-2 focus:ring-zinc-900 outline-none text-sm font-medium"
              />
              <button 
                type="submit"
                className="bg-zinc-900 text-white px-4 py-2 rounded-xl font-bold hover:scale-105 active:scale-95 transition-all flex items-center gap-2 text-sm shadow-sm"
              >
                <Plus size={16} />
                Añadir
              </button>
            </form>
          </div>

          <div className="space-y-3 max-h-[400px] overflow-y-auto pr-1 custom-scrollbar">
            {filteredBrands.map((brand) => (
              <div key={brand.id} className="flex items-center justify-between p-4 bg-zinc-50 rounded-2xl group">
                {editingBrand?.id === brand.id ? (
                  <form onSubmit={handleUpdateBrand} className="flex-1 flex gap-2">
                    <input 
                      type="text" 
                      value={editingBrand.name}
                      onChange={(e) => setEditingBrand({ ...editingBrand, name: e.target.value })}
                      className="flex-1 bg-white border border-zinc-200 rounded-lg px-3 py-1 outline-none focus:ring-2 focus:ring-zinc-900 font-medium"
                      autoFocus
                    />
                    <button type="submit" className="text-green-600 font-bold text-sm px-2">Guardar</button>
                    <button type="button" onClick={() => setEditingBrand(null)} className="text-zinc-400 font-bold text-sm px-2">Cancelar</button>
                  </form>
                ) : (
                  <>
                    <span className="font-bold text-zinc-900">{brand.name}</span>
                    <div className="flex gap-1">
                      <button 
                        onClick={() => setEditingBrand(brand)}
                        className="p-2 text-zinc-500 hover:text-zinc-900 hover:bg-white rounded-lg transition-all"
                        title="Editar"
                      >
                        <Edit size={18} />
                      </button>
                      <button 
                        onClick={() => handleDeleteBrand(brand.id)}
                        className="p-2 text-zinc-500 hover:text-red-600 hover:bg-red-50 rounded-lg transition-all"
                        title="Eliminar"
                      >
                        <Trash2 size={18} />
                      </button>
                    </div>
                  </>
                )}
              </div>
            ))}
            {brands.length === 0 && (
              <p className="text-center py-4 text-zinc-400 font-medium">No hay marcas registradas</p>
            )}
          </div>
        </div>

        <div className="bg-white p-8 rounded-3xl border border-zinc-100 shadow-sm">
          <h3 className="font-bold mb-6 flex items-center gap-2">
            <LayoutGrid size={20} />
            Preferencias del Sistema
          </h3>
          <div className="space-y-4">
            <div className="flex items-center justify-between p-4 hover:bg-zinc-50 rounded-2xl transition-colors cursor-pointer">
              <div>
                <p className="font-bold">Modo Oscuro</p>
                <p className="text-xs text-zinc-500 font-medium">Próximamente</p>
              </div>
              <div className="w-12 h-6 bg-zinc-200 rounded-full relative">
                <div className="absolute left-1 top-1 w-4 h-4 bg-white rounded-full shadow-sm" />
              </div>
            </div>
            <div className="flex items-center justify-between p-4 hover:bg-zinc-50 rounded-2xl transition-colors cursor-pointer">
              <div>
                <p className="font-bold">Notificaciones de Stock</p>
                <p className="text-xs text-zinc-500 font-medium">Recibe alertas cuando el stock sea bajo</p>
              </div>
              <div className="w-12 h-6 bg-green-500 rounded-full relative">
                <div className="absolute right-1 top-1 w-4 h-4 bg-white rounded-full shadow-sm" />
              </div>
            </div>
          </div>
        </div>

        <div className="bg-white p-8 rounded-3xl border border-zinc-100 shadow-sm">
          <h3 className="font-bold mb-6 text-red-600">Zona de Peligro</h3>
          <button 
            onClick={() => logout()}
            className="w-full flex items-center justify-center gap-2 bg-red-50 text-red-600 py-4 rounded-2xl font-bold hover:bg-red-100 transition-all"
          >
            <LogOut size={20} />
            Cerrar Sesión de la Cuenta
          </button>
        </div>
      </div>
    </div>
  );
};

const WebStoreView = ({ products, clients, onSaleComplete }: { products: Product[], clients: Client[], onSaleComplete: (sale: Sale) => void }) => {
  const [isLoggedIn, setIsLoggedIn] = useState(false);
  const [currentClient, setCurrentClient] = useState<Client | null>(null);
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [cart, setCart] = useState<CartItem[]>([]);
  const [isCheckingOut, setIsCheckingOut] = useState(false);

  const handleLogin = (e: React.FormEvent) => {
    e.preventDefault();
    const client = clients.find(c => c.email === email && c.password === password);
    if (client) {
      setIsLoggedIn(true);
      setCurrentClient(client);
      setError('');
    } else {
      setError('Credenciales incorrectas');
    }
  };

  const addToCart = (product: Product) => {
    if (product.stock <= 0) return;
    setCart(prev => {
      const existing = prev.find(item => item.id === product.id);
      if (existing) {
        return prev.map(item => item.id === product.id ? { ...item, quantity: item.quantity + 1 } : item);
      }
      return [...prev, { ...product, quantity: 1 }];
    });
  };

  const removeFromCart = (id: string) => {
    setCart(prev => prev.filter(item => item.id !== id));
  };

  const updateQuantity = (id: string, delta: number) => {
    setCart(prev => prev.map(item => {
      if (item.id === id) {
        const newQty = Math.max(1, item.quantity + delta);
        const product = products.find(p => p.id === id);
        if (product && newQty > product.stock) return item;
        return { ...item, quantity: newQty };
      }
      return item;
    }));
  };

  const cartTotal = cart.reduce((sum, item) => sum + (item.price * item.quantity), 0);

  const handleCheckout = async () => {
    if (!currentClient || cart.length === 0) return;
    setIsCheckingOut(true);
    try {
      const saleData: Sale = {
        items: cart.map(item => ({
          productId: item.id!,
          name: item.name,
          price: item.price,
          quantity: item.quantity
        })),
        total: cartTotal,
        timestamp: new Date().toISOString(),
        userId: 'webstore',
        clientId: currentClient.id,
        paymentMethods: [{ type: 'Otro', amount: cartTotal }]
      };

      const docRef = await addDoc(collection(db, 'sales'), saleData);
      
      for (const item of cart) {
        const productRef = doc(db, 'products', item.id!);
        await updateDoc(productRef, {
          stock: increment(-item.quantity)
        });
      }

      onSaleComplete({ ...saleData, id: docRef.id });
      setCart([]);
      setIsCheckingOut(false);
      alert('¡Compra realizada con éxito!');
    } catch (err) {
      console.error(err);
      setIsCheckingOut(false);
      alert('Error al procesar la compra');
    }
  };

  if (!isLoggedIn) {
    return (
      <div className="flex-1 flex items-center justify-center bg-zinc-50 p-4">
        <motion.div 
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          className="bg-white p-8 rounded-3xl border border-zinc-100 shadow-xl max-w-md w-full"
        >
          <div className="text-center mb-8">
            <div className="w-16 h-16 bg-zinc-900 text-white rounded-2xl flex items-center justify-center mx-auto mb-4 shadow-lg">
              <Globe size={32} />
            </div>
            <h2 className="text-2xl font-bold">Tienda Web OmniPOS</h2>
            <p className="text-zinc-500">Inicia sesión con tu cuenta de cliente</p>
          </div>

          <form onSubmit={handleLogin} className="space-y-4">
            <div>
              <label className="block text-xs font-bold text-zinc-400 uppercase tracking-widest mb-2">Email</label>
              <div className="relative">
                <Mail className="absolute left-4 top-1/2 -translate-y-1/2 text-zinc-400" size={18} />
                <input 
                  type="email"
                  value={email}
                  onChange={(e) => setEmail(e.target.value)}
                  className="w-full bg-zinc-50 border-none rounded-xl py-3 pl-12 pr-4 focus:ring-2 focus:ring-zinc-900 transition-all"
                  placeholder="tu@email.com"
                  required
                />
              </div>
            </div>
            <div>
              <label className="block text-xs font-bold text-zinc-400 uppercase tracking-widest mb-2">Contraseña</label>
              <div className="relative">
                <Lock className="absolute left-4 top-1/2 -translate-y-1/2 text-zinc-400" size={18} />
                <input 
                  type="password"
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  className="w-full bg-zinc-50 border-none rounded-xl py-3 pl-12 pr-4 focus:ring-2 focus:ring-zinc-900 transition-all"
                  placeholder="••••••••"
                  required
                />
              </div>
            </div>
            {error && <p className="text-red-500 text-sm font-bold text-center">{error}</p>}
            <button 
              type="submit"
              className="w-full bg-zinc-900 text-white py-4 rounded-2xl font-bold hover:scale-[1.02] active:scale-95 transition-all shadow-lg"
            >
              Entrar
            </button>
          </form>

          <button 
            onClick={() => window.location.hash = ''}
            className="w-full mt-6 flex items-center justify-center gap-2 text-zinc-400 font-bold text-sm hover:text-zinc-900 transition-colors"
          >
            <ArrowLeft size={16} />
            Volver al Inicio
          </button>
        </motion.div>
      </div>
    );
  }

  return (
    <div className="flex-1 flex flex-col h-full overflow-hidden bg-zinc-50/50">
      <header className="bg-white border-b border-zinc-100 p-6 flex items-center justify-between">
        <div className="flex items-center gap-4">
          <div className="w-10 h-10 bg-zinc-900 text-white rounded-xl flex items-center justify-center shadow-md">
            <Globe size={20} />
          </div>
          <div>
            <h2 className="font-bold">Bienvenido, {currentClient?.name}</h2>
            <p className="text-xs text-zinc-500">{currentClient?.salonName || 'Cliente Web'}</p>
          </div>
        </div>
        <button 
          onClick={() => setIsLoggedIn(false)}
          className="text-zinc-400 font-bold text-sm hover:text-zinc-900 transition-colors"
        >
          Cerrar Sesión
        </button>
      </header>

      <div className="flex-1 flex overflow-hidden">
        <div className="flex-1 overflow-y-auto p-6 custom-scrollbar">
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-6">
            {products.filter(p => p.stock > 0).map(product => (
              <motion.div 
                key={product.id}
                whileHover={{ y: -4 }}
                className="bg-white rounded-3xl border border-zinc-100 shadow-sm overflow-hidden group"
              >
                <div className="aspect-square bg-zinc-100 relative overflow-hidden">
                  {product.imageUrl ? (
                    <img src={product.imageUrl} alt={product.name} className="w-full h-full object-cover group-hover:scale-110 transition-transform duration-500" referrerPolicy="no-referrer" />
                  ) : (
                    <div className="w-full h-full flex items-center justify-center text-zinc-300">
                      <Package size={48} />
                    </div>
                  )}
                  <div className="absolute top-4 right-4 bg-white/90 backdrop-blur-md px-3 py-1 rounded-full text-xs font-black shadow-sm">
                    ${product.price.toFixed(2)}
                  </div>
                </div>
                <div className="p-4">
                  <p className="text-[10px] font-bold text-zinc-400 uppercase tracking-widest mb-1">{product.category}</p>
                  <h4 className="font-bold text-zinc-900 mb-4 line-clamp-1">{product.name}</h4>
                  <button 
                    onClick={() => addToCart(product)}
                    className="w-full bg-zinc-900 text-white py-3 rounded-xl font-bold text-sm hover:bg-zinc-800 transition-all flex items-center justify-center gap-2"
                  >
                    <Plus size={16} />
                    Agregar
                  </button>
                </div>
              </motion.div>
            ))}
          </div>
        </div>

        <div className="w-96 bg-white border-l border-zinc-100 flex flex-col hidden xl:flex">
          <div className="p-6 border-b border-zinc-100">
            <h3 className="font-bold flex items-center gap-2">
              <ShoppingCart size={20} />
              Tu Carrito
            </h3>
          </div>
          <div className="flex-1 overflow-y-auto p-6 space-y-4 custom-scrollbar">
            {cart.map(item => (
              <div key={item.id} className="flex items-center gap-4">
                <div className="w-12 h-12 bg-zinc-100 rounded-lg overflow-hidden flex-shrink-0">
                  {item.imageUrl && <img src={item.imageUrl} alt={item.name} className="w-full h-full object-cover" referrerPolicy="no-referrer" />}
                </div>
                <div className="flex-1 min-w-0">
                  <p className="font-bold text-sm truncate">{item.name}</p>
                  <p className="text-xs text-zinc-400">${item.price.toFixed(2)} x {item.quantity}</p>
                </div>
                <div className="flex items-center gap-2">
                  <button onClick={() => updateQuantity(item.id!, -1)} className="p-1 hover:bg-zinc-100 rounded-md"><Minus size={14} /></button>
                  <span className="text-sm font-bold w-4 text-center">{item.quantity}</span>
                  <button onClick={() => updateQuantity(item.id!, 1)} className="p-1 hover:bg-zinc-100 rounded-md"><Plus size={14} /></button>
                  <button onClick={() => removeFromCart(item.id!)} className="p-1 text-red-500 hover:bg-red-50 rounded-md ml-2"><Trash2 size={14} /></button>
                </div>
              </div>
            ))}
            {cart.length === 0 && (
              <div className="h-full flex flex-col items-center justify-center text-zinc-300 gap-4">
                <ShoppingCart size={48} />
                <p className="font-bold">Carrito vacío</p>
              </div>
            )}
          </div>
          <div className="p-6 bg-zinc-50 border-t border-zinc-100">
            <div className="flex justify-between items-center mb-6">
              <span className="font-bold text-zinc-400 uppercase tracking-widest text-xs">Total</span>
              <span className="text-2xl font-black tracking-tighter">${cartTotal.toFixed(2)}</span>
            </div>
            <button 
              onClick={handleCheckout}
              disabled={cart.length === 0 || isCheckingOut}
              className="w-full bg-zinc-900 text-white py-4 rounded-2xl font-bold hover:scale-[1.02] active:scale-95 transition-all shadow-lg disabled:opacity-50 disabled:scale-100"
            >
              {isCheckingOut ? 'Procesando...' : 'Finalizar Compra'}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};

const DashboardView = ({ sales, products, clients }: { sales: Sale[], products: Product[], clients: Client[] }) => {
  const totalRevenue = sales.reduce((sum, sale) => sum + sale.total, 0);
  const totalSalesCount = sales.length;
  const lowStockProducts = products.filter(p => p.stock < 5);

  // Helper to filter sales by date range
  const getSalesInRange = (startDate: Date, endDate: Date) => {
    return sales.filter(sale => {
      const saleDate = new Date(sale.timestamp);
      return saleDate >= startDate && saleDate <= endDate;
    });
  };

  const now = new Date();
  const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  const endOfMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59);
  const startOfYear = new Date(now.getFullYear(), 0, 1);
  const endOfYear = new Date(now.getFullYear(), 11, 31, 23, 59, 59);

  const monthSales = getSalesInRange(startOfMonth, endOfMonth);
  const yearSales = getSalesInRange(startOfYear, endOfYear);

  // Helper to get client ranking (best or worst)
  const getClientRanking = (periodSales: Sale[], order: 'asc' | 'desc') => {
    const clientSpending: { [key: string]: number } = {};
    
    // Initialize all clients with 0 spending
    clients.forEach(client => {
      clientSpending[client.id!] = 0;
    });

    // Add spending from sales
    periodSales.forEach(sale => {
      if (sale.clientId && clientSpending[sale.clientId] !== undefined) {
        clientSpending[sale.clientId] += sale.total;
      }
    });

    const ranking = Object.entries(clientSpending)
      .map(([id, total]) => ({
        id,
        name: clients.find(c => c.id === id)?.name || 'Cliente Desconocido',
        total
      }));

    if (order === 'asc') {
      return ranking.sort((a, b) => a.total - b.total).slice(0, 5);
    } else {
      return ranking.sort((a, b) => b.total - a.total).slice(0, 5);
    }
  };

  const worstClientsMonth = getClientRanking(monthSales, 'asc');
  const worstClientsYear = getClientRanking(yearSales, 'asc');
  const bestClientsMonth = getClientRanking(monthSales, 'desc');
  const bestClientsYear = getClientRanking(yearSales, 'desc');

  // Process data for Category Chart
  const salesByCategory = sales.reduce((acc: { [key: string]: number }, sale) => {
    sale.items.forEach(item => {
      const product = products.find(p => p.id === item.productId);
      const category = product?.category || 'General';
      acc[category] = (acc[category] || 0) + (item.price * item.quantity);
    });
    return acc;
  }, {});

  const categoryData = Object.keys(salesByCategory).map(name => ({
    name,
    value: salesByCategory[name]
  })).sort((a, b) => b.value - a.value);

  // Process data for Daily Sales Chart
  const salesByDate = sales.reduce((acc: { [key: string]: number }, sale) => {
    const date = new Date(sale.timestamp).toLocaleDateString();
    acc[date] = (acc[date] || 0) + sale.total;
    return acc;
  }, {});

  const dailyData = Object.keys(salesByDate).map(date => ({
    date,
    total: salesByDate[date]
  })).sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());

  // Top Products
  const productSales = sales.reduce((acc: { [key: string]: { name: string, total: number, quantity: number } }, sale) => {
    sale.items.forEach(item => {
      if (!acc[item.productId]) {
        acc[item.productId] = { name: item.name, total: 0, quantity: 0 };
      }
      acc[item.productId].total += item.price * item.quantity;
      acc[item.productId].quantity += item.quantity;
    });
    return acc;
  }, {});

  const topProducts = Object.values(productSales)
    .sort((a, b) => b.total - a.total)
    .slice(0, 5);

  const COLORS = ['#18181b', '#3f3f46', '#71717a', '#a1a1aa', '#d4d4d8'];

  return (
    <div className="flex-1 p-6 md:p-10 lg:p-12 overflow-y-auto custom-scrollbar bg-zinc-50/10">
      <header className="flex flex-col md:flex-row md:items-end justify-between gap-8 mb-12">
        <div className="space-y-1">
          <h2 className="text-4xl font-black tracking-tighter text-zinc-900 uppercase">Panel de Control</h2>
          <div className="flex items-center gap-2">
            <span className="w-2 h-2 rounded-full bg-green-500 animate-pulse" />
            <p className="text-zinc-400 text-[10px] font-black uppercase tracking-[0.2em]">En tiempo real — {new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}</p>
          </div>
        </div>
        
        <div className="flex items-center gap-3 no-print">
          <button 
            onClick={() => window.print()}
            className="flex items-center gap-2 bg-white border border-zinc-200 text-zinc-900 px-6 py-3.5 rounded-2xl text-[10px] font-black uppercase tracking-widest hover:bg-zinc-50 transition-all shadow-sm active:scale-95"
          >
            <Printer size={16} />
            Exportar Informe
          </button>
        </div>
      </header>

      {/* Stats Grid */}
      <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-6 mb-12">
        <div className="group bg-white p-8 rounded-[2rem] border border-zinc-100 shadow-sm hover:shadow-2xl hover:shadow-zinc-200/50 transition-all duration-500">
          <div className="flex items-center justify-between mb-8">
            <div className="w-14 h-14 bg-zinc-900 text-white rounded-3xl flex items-center justify-center rotate-3 group-hover:rotate-0 transition-transform">
              <BarChart3 size={24} className="-rotate-3 group-hover:rotate-0 transition-transform" />
            </div>
          </div>
          <p className="text-[10px] font-black text-zinc-300 uppercase tracking-[0.2em] mb-1">Ingresos Totales</p>
          <h3 className="text-5xl font-black tracking-tighter text-zinc-900">${totalRevenue.toLocaleString('es-ES', { minimumFractionDigits: 2 })}</h3>
        </div>

        <div className="group bg-white p-8 rounded-[2rem] border border-zinc-100 shadow-sm hover:shadow-2xl hover:shadow-zinc-200/50 transition-all duration-500">
          <div className="flex items-center justify-between mb-8">
            <div className="w-14 h-14 bg-white border border-zinc-100 rounded-3xl flex items-center justify-center -rotate-3 group-hover:rotate-0 transition-transform shadow-sm">
              <ShoppingCart size={24} className="text-zinc-900 rotate-3 group-hover:rotate-0 transition-transform" />
            </div>
          </div>
          <p className="text-[10px] font-black text-zinc-300 uppercase tracking-[0.2em] mb-1">Pedidos Totales</p>
          <h3 className="text-5xl font-black tracking-tighter text-zinc-900">{totalSalesCount}</h3>
        </div>

        <div className="group bg-white p-8 rounded-[2rem] border border-zinc-100 shadow-sm hover:shadow-2xl hover:shadow-zinc-200/50 transition-all duration-500">
          <div className="flex items-center justify-between mb-8">
            <div className={`w-14 h-14 rounded-3xl flex items-center justify-center transition-all ${lowStockProducts.length > 0 ? 'bg-amber-500 text-white shadow-lg shadow-amber-500/20' : 'bg-zinc-50 text-zinc-300'}`}>
              <Package size={24} />
            </div>
          </div>
          <p className="text-[10px] font-black text-zinc-300 uppercase tracking-[0.2em] mb-1">Stock Bajo</p>
          <h3 className="text-5xl font-black tracking-tighter text-zinc-900">{lowStockProducts.length}</h3>
        </div>
      </div>

      {/* Charts Section */}
      <div className="grid grid-cols-1 xl:grid-cols-2 gap-8 mb-12">
        <section className="bg-white p-8 rounded-[2rem] border border-zinc-100 shadow-sm hover:shadow-xl transition-shadow">
          <header className="flex items-center justify-between mb-10">
            <div>
              <h3 className="text-lg font-black tracking-tighter uppercase mb-1">Ventas Diarias</h3>
              <p className="text-[10px] font-bold text-zinc-400 uppercase tracking-widest">Tendencia de ingresos</p>
            </div>
            <div className="p-3 bg-zinc-50 rounded-2xl text-zinc-400">
              <TrendingUp size={20} />
            </div>
          </header>
          
          <div className="h-[320px] w-full">
            <ResponsiveContainer width="100%" height="100%">
              <LineChart data={dailyData}>
                <CartesianGrid strokeDasharray="4 4" vertical={false} stroke="#f4f4f5" />
                <XAxis 
                  dataKey="date" 
                  axisLine={false} 
                  tickLine={false} 
                  tick={{ fill: '#a1a1aa', fontSize: 9, fontWeight: 700 }}
                  dy={10}
                />
                <YAxis 
                  axisLine={false} 
                  tickLine={false} 
                  tick={{ fill: '#a1a1aa', fontSize: 9, fontWeight: 700 }}
                  tickFormatter={(v) => `$${v}`}
                />
                <Tooltip 
                  contentStyle={{ 
                    borderRadius: '24px', 
                    border: 'none', 
                    boxShadow: '0 20px 40px -10px rgba(0,0,0,0.12)',
                    padding: '16px',
                    fontFamily: 'Inter',
                    fontSize: '12px'
                  }}
                />
                <Line 
                  type="monotone" 
                  dataKey="total" 
                  stroke="#09090b" 
                  strokeWidth={4} 
                  dot={{ fill: '#09090b', strokeWidth: 2, r: 4, stroke: '#fff' }}
                  activeDot={{ r: 8, fill: '#09090b' }}
                />
              </LineChart>
            </ResponsiveContainer>
          </div>
        </section>

        <section className="bg-white p-8 rounded-[2rem] border border-zinc-100 shadow-sm hover:shadow-xl transition-shadow">
          <header className="flex items-center justify-between mb-10">
            <div>
              <h3 className="text-lg font-black tracking-tighter uppercase mb-1">Categorías Top</h3>
              <p className="text-[10px] font-bold text-zinc-400 uppercase tracking-widest">Distribución por volumen</p>
            </div>
            <div className="p-3 bg-zinc-50 rounded-2xl text-zinc-400">
              <PieChart size={20} />
            </div>
          </header>

          <div className="h-[320px] w-full">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={categoryData} layout="vertical">
                <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#f4f4f5" />
                <XAxis type="number" hide />
                <YAxis 
                  dataKey="name" 
                  type="category" 
                  axisLine={false} 
                  tickLine={false} 
                  tick={{ fill: '#18181b', fontSize: 11, fontWeight: 800 }}
                  width={100}
                />
                <Tooltip 
                  cursor={{ fill: '#f8fafc' }}
                  contentStyle={{ borderRadius: '24px', border: 'none', boxShadow: '0 20px 40px -10px rgba(0,0,0,0.12)' }}
                />
                <Bar dataKey="value" radius={[0, 12, 12, 0]} barSize={24}>
                  {categoryData.map((entry, index) => (
                    <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                  ))}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </section>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-8 mb-12">
        <section className="bg-white p-8 rounded-[2rem] border border-zinc-100 shadow-sm">
          <h3 className="text-lg font-black tracking-tighter uppercase mb-8">Mejores Clientes (Mes)</h3>
          <div className="space-y-4">
            {bestClientsMonth.map((c, idx) => (
              <div key={idx} className="flex items-center justify-between group p-3 hover:bg-zinc-50 rounded-[1.5rem] transition-colors">
                <div className="flex items-center gap-4">
                  <div className="w-10 h-10 bg-zinc-50 text-zinc-400 rounded-xl flex items-center justify-center border border-zinc-100 group-hover:bg-zinc-900 group-hover:text-white group-hover:border-zinc-900 transition-all">
                    <User size={18} />
                  </div>
                  <div className="min-w-0">
                    <p className="font-bold text-zinc-900 truncate uppercase tracking-tight text-sm">{c.name}</p>
                    <p className="text-[10px] text-zinc-500 font-black tracking-widest">Inversión mensual</p>
                  </div>
                </div>
                <div className="text-right">
                  <p className="font-black text-zinc-900">${c.total.toLocaleString('es-ES', { minimumFractionDigits: 2 })}</p>
                </div>
              </div>
            ))}
          </div>
        </section>

        <section className="bg-white p-8 rounded-[2rem] border border-zinc-100 shadow-sm">
          <h3 className="text-lg font-black tracking-tighter uppercase mb-8">Productos más vendidos</h3>
          <div className="space-y-4">
            {topProducts.map((p, idx) => (
              <div key={idx} className="flex items-center justify-between group p-3 hover:bg-zinc-50 rounded-[1.5rem] transition-colors">
                <div className="flex items-center gap-4">
                  <div className="w-10 h-10 bg-zinc-900 text-white rounded-xl flex items-center justify-center text-xs font-black">
                    0{idx + 1}
                  </div>
                  <div className="min-w-0">
                    <p className="font-bold text-zinc-900 truncate uppercase tracking-tight text-sm">{p.name}</p>
                    <p className="text-[10px] text-zinc-400 font-black uppercase tracking-widest">{p.quantity} unidades vendidas</p>
                  </div>
                </div>
                <p className="font-black text-zinc-900">${p.total.toLocaleString('es-ES', { minimumFractionDigits: 2 })}</p>
              </div>
            ))}
          </div>
        </section>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-8 mb-12">
        <section className="bg-white p-8 rounded-[2rem] border border-zinc-100 shadow-sm">
          <h3 className="text-lg font-black tracking-tighter uppercase mb-8 text-red-600">Peores Clientes (Mes)</h3>
          <div className="space-y-4">
            {worstClientsMonth.map((c, idx) => (
              <div key={idx} className="flex items-center justify-between group p-3 hover:bg-zinc-50 rounded-[1.5rem] transition-colors">
                <div className="flex items-center gap-4">
                  <div className="w-10 h-10 bg-zinc-50 text-zinc-300 rounded-xl flex items-center justify-center border border-zinc-100 group-hover:bg-zinc-900 group-hover:text-white group-hover:border-zinc-900 transition-all">
                    <User size={18} />
                  </div>
                  <div className="min-w-0">
                    <p className="font-bold text-zinc-900 truncate uppercase tracking-tight text-sm">{c.name}</p>
                    <p className="text-[10px] text-zinc-400 font-black tracking-widest">Inversión mínima</p>
                  </div>
                </div>
                <div className="text-right">
                  <p className="font-black text-zinc-900">${c.total.toLocaleString('es-ES', { minimumFractionDigits: 2 })}</p>
                </div>
              </div>
            ))}
          </div>
        </section>

        <section className="bg-white p-8 rounded-[2rem] border border-zinc-100 shadow-sm">
          <h3 className="text-lg font-black tracking-tighter uppercase mb-8">Transacciones Recientes</h3>
          <div className="space-y-4">
            {sales.slice(0, 5).map((sale) => (
              <div key={sale.id} className="flex items-center justify-between group p-3 hover:bg-zinc-50 rounded-[1.5rem] transition-colors">
                <div className="flex items-center gap-4">
                  <div className="w-10 h-10 bg-zinc-50 text-zinc-400 rounded-xl flex items-center justify-center border border-zinc-100">
                    <History size={18} />
                  </div>
                  <div className="min-w-0">
                    <p className="font-bold text-zinc-900 truncate uppercase tracking-tight text-sm">{sale.items.length} artículos</p>
                    <p className="text-[10px] text-zinc-400 font-bold tracking-widest">{new Date(sale.timestamp).toLocaleString()}</p>
                  </div>
                </div>
                <div className="text-right">
                  <p className="font-black text-zinc-900">${sale.total.toLocaleString('es-ES', { minimumFractionDigits: 2 })}</p>
                </div>
              </div>
            ))}
          </div>
        </section>
      </div>
    </div>
  );
};

const ReportsView = ({ sales, products, clients, categories }: { sales: Sale[], products: Product[], clients: Client[], categories: Category[] }) => {
  const [reportDate, setReportDate] = useState<string>(new Date().toISOString().slice(0, 7)); // Monthly view by default
  const [reportType, setReportType] = useState<'sales' | 'products' | 'clients' | 'payments'>('sales');

  const filteredSales = sales.filter(sale => sale.timestamp.startsWith(reportDate));

  // --- Data Processing for Reports ---

  // 1. Sales by Day
  const salesByDay = filteredSales.reduce((acc: { [key: string]: number }, sale) => {
    const day = sale.timestamp.split('T')[0];
    acc[day] = (acc[day] || 0) + sale.total;
    return acc;
  }, {});

  const dailyData = Object.entries(salesByDay).map(([date, total]) => ({
    date: date.slice(-2),
    total
  })).sort((a, b) => a.date.localeCompare(b.date));

  // 2. Sales by Payment Method
  const paymentMethodData = filteredSales.reduce((acc: { [key: string]: number }, sale) => {
    (sale.paymentMethods || []).forEach(pm => {
      acc[pm.type] = (acc[pm.type] || 0) + pm.amount;
    });
    return acc;
  }, {});

  const paymentData = Object.entries(paymentMethodData).map(([name, value]) => ({ name, value }));

  // 3. Top Products (Quantity)
  const productQuantities = filteredSales.reduce((acc: { [key: string]: { name: string, quantity: number, total: number } }, sale) => {
    sale.items.forEach(item => {
      if (!acc[item.productId]) {
        acc[item.productId] = { name: item.name, quantity: 0, total: 0 };
      }
      acc[item.productId].quantity += item.quantity;
      acc[item.productId].total += item.price * item.quantity;
    });
    return acc;
  }, {});

  const topProductsSorted = Object.values(productQuantities).sort((a, b) => b.quantity - a.quantity).slice(0, 10);

  // 4. Client Spending
  const clientSpending = filteredSales.reduce((acc: { [key: string]: { name: string, total: number } }, sale) => {
    const clientId = sale.clientId || 'default';
    const clientName = clients.find(c => c.id === clientId)?.name || 'Consumidor Final';
    if (!acc[clientId]) {
      acc[clientId] = { name: clientName, total: 0 };
    }
    acc[clientId].total += sale.total;
    return acc;
  }, {});

  const topClients = Object.values(clientSpending).sort((a, b) => b.total - a.total).slice(0, 5);

  const exportToExcel = () => {
    const data = filteredSales.map(sale => ({
      Folio: sale.folio || '-',
      Fecha: new Date(sale.timestamp).toLocaleString(),
      Cliente: clients.find(c => c.id === sale.clientId)?.name || 'Consumidor Final',
      Total: sale.total,
      Metodos: sale.paymentMethods.map(pm => `${pm.type}: $${pm.amount}`).join(', ')
    }));

    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Reporte_Ventas");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(topProductsSorted), "Top_Productos");
    
    XLSX.writeFile(wb, `Reporte_OmniPOS_${reportDate}.xlsx`);
  };

  const COLORS = ['#18181b', '#3f3f46', '#71717a', '#a1a1aa', '#d4d4d8', '#27272a', '#52525b'];

  return (
    <div className="flex-1 p-3 md:p-8 flex flex-col h-full overflow-hidden">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-3 mb-4 shrink-0">
        <div>
          <h2 className="text-2xl md:text-3xl font-bold tracking-tight mb-0.5">Reportes</h2>
          <p className="text-zinc-500 font-medium text-xs md:text-base">Análisis detallado de tu negocio</p>
        </div>
        <div className="flex flex-wrap items-center gap-2 no-print">
          <input 
            type="month"
            value={reportDate}
            onChange={(e) => setReportDate(e.target.value)}
            className="bg-white border border-zinc-200 rounded-xl px-3 py-2 text-sm font-bold outline-none focus:ring-2 focus:ring-zinc-900 shadow-sm"
          />
          <button 
            onClick={exportToExcel}
            className="flex items-center gap-2 bg-white text-zinc-900 border border-zinc-200 px-3 py-2 rounded-xl text-xs font-bold hover:bg-zinc-50 transition-all shadow-sm"
          >
            <Download size={16} />
            Excel
          </button>
          <button 
            onClick={() => window.print()}
            className="flex items-center gap-2 bg-zinc-900 text-white px-3 py-2 rounded-xl text-xs font-bold hover:scale-105 transition-all shadow-lg"
          >
            <Printer size={16} />
            Imprimir
          </button>
        </div>
      </div>

      {/* Report Type Tabs */}
      <div className="flex overflow-x-auto gap-2 mb-4 no-scrollbar no-print shrink-0">
        {[
          { id: 'sales', label: 'Resumen', icon: TrendingUp },
          { id: 'products', label: 'Top Prod.', icon: Package },
          { id: 'clients', label: 'Clientes', icon: Users },
          { id: 'payments', label: 'Pagos', icon: PieChart },
        ].map((tab) => (
          <button
            key={tab.id}
            onClick={() => setReportType(tab.id as any)}
            className={`flex items-center gap-2 px-4 py-2 rounded-xl font-bold text-xs whitespace-nowrap transition-all ${
              reportType === tab.id 
                ? 'bg-zinc-900 text-white shadow-lg' 
                : 'bg-white text-zinc-500 border border-zinc-100 hover:border-zinc-300'
            }`}
          >
            <tab.icon size={16} />
            {tab.label}
          </button>
        ))}
      </div>

      <div className="flex-1 overflow-y-auto custom-scrollbar pr-1">

      <AnimatePresence mode="wait">
        <motion.div
          key={reportType}
          initial={{ opacity: 0, y: 10 }}
          animate={{ opacity: 1, y: 0 }}
          exit={{ opacity: 0, y: -10 }}
          className="space-y-8"
        >
          {reportType === 'sales' && (
            <>
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
                <div className="bg-white p-6 rounded-3xl border border-zinc-100 shadow-sm">
                  <p className="text-[10px] font-bold text-zinc-400 uppercase tracking-widest mb-1">Ventas Totales</p>
                  <h4 className="text-2xl font-black">${filteredSales.reduce((acc, s) => acc + s.total, 0).toFixed(2)}</h4>
                </div>
                <div className="bg-white p-6 rounded-3xl border border-zinc-100 shadow-sm">
                  <p className="text-[10px] font-bold text-zinc-400 uppercase tracking-widest mb-1">Cantidad de Pedidos</p>
                  <h4 className="text-2xl font-black">{filteredSales.length}</h4>
                </div>
                <div className="bg-white p-6 rounded-3xl border border-zinc-100 shadow-sm">
                  <p className="text-[10px] font-bold text-zinc-400 uppercase tracking-widest mb-1">Ticket Promedio</p>
                  <h4 className="text-2xl font-black">
                    ${filteredSales.length > 0 ? (filteredSales.reduce((acc, s) => acc + s.total, 0) / filteredSales.length).toFixed(2) : '0.00'}
                  </h4>
                </div>
              </div>
              <div className="bg-white p-6 rounded-3xl border border-zinc-100 shadow-sm">
                <h3 className="font-bold mb-6">Tendencia diaria del mes ({reportDate})</h3>
                <div className="h-[300px] w-full">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={dailyData}>
                      <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f4f4f5" />
                      <XAxis dataKey="date" axisLine={false} tickLine={false} />
                      <YAxis axisLine={false} tickLine={false} tickFormatter={(v) => `$${v}`} />
                      <Tooltip contentStyle={{ borderRadius: '16px', border: 'none' }} />
                      <Bar dataKey="total" fill="#18181b" radius={[4, 4, 0, 0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </>
          )}

          {reportType === 'products' && (
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
              <div className="bg-white p-6 rounded-3xl border border-zinc-100 shadow-sm">
                <h3 className="font-bold mb-6">Unidades Vendidas (Top 10)</h3>
                <div className="h-[400px]">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart layout="vertical" data={topProductsSorted}>
                      <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#f4f4f5" />
                      <XAxis type="number" hide />
                      <YAxis type="category" dataKey="name" width={150} axisLine={false} tickLine={false} tick={{ fontSize: 10 }} />
                      <Tooltip cursor={{ fill: '#f4f4f5' }} />
                      <Bar dataKey="quantity" fill="#18181b" radius={[0, 4, 4, 0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
              <div className="bg-white rounded-3xl border border-zinc-100 shadow-sm overflow-hidden">
                <div className="p-6 border-b border-zinc-50 font-bold">Resumen Detallado</div>
                <div className="overflow-x-auto">
                  <table className="w-full text-left">
                    <thead className="bg-zinc-50 border-b border-zinc-100">
                      <tr>
                        <th className="p-4 text-[10px] font-bold text-zinc-400 uppercase tracking-widest">Producto</th>
                        <th className="p-4 text-[10px] font-bold text-zinc-400 uppercase tracking-widest text-right">Cant.</th>
                        <th className="p-4 text-[10px] font-bold text-zinc-400 uppercase tracking-widest text-right">Total</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-zinc-50">
                      {topProductsSorted.map((p, i) => (
                        <tr key={i} className="hover:bg-zinc-50/50">
                          <td className="p-4 text-sm font-bold">{p.name}</td>
                          <td className="p-4 text-sm text-right">{p.quantity}</td>
                          <td className="p-4 text-sm font-black text-right">${p.total.toFixed(2)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {reportType === 'clients' && (
            <div className="bg-white rounded-3xl border border-zinc-100 shadow-sm overflow-hidden">
              <div className="p-6 border-b border-zinc-50 font-bold">Clientes con mejores compras en el periodo</div>
              <div className="divide-y divide-zinc-50">
                {topClients.map((c, i) => (
                  <div key={i} className="p-6 flex items-center justify-between">
                    <div className="flex items-center gap-4">
                      <div className="w-10 h-10 bg-zinc-900 text-white rounded-full flex items-center justify-center font-bold">
                        {i + 1}
                      </div>
                      <div>
                        <p className="font-bold">{c.name}</p>
                        <p className="text-[10px] font-bold text-zinc-400 uppercase tracking-widest">Cliente frecuente</p>
                      </div>
                    </div>
                    <div className="text-right">
                      <p className="text-lg font-black">${c.total.toFixed(2)}</p>
                      <p className="text-[10px] font-bold text-green-600 uppercase tracking-widest">Facturado</p>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}

          {reportType === 'payments' && (
            <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
              <div className="bg-white p-6 rounded-3xl border border-zinc-100 shadow-sm flex flex-col items-center">
                <h3 className="font-bold mb-6 w-full">Distribución por Tipo</h3>
                <div className="h-[300px] w-full">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={paymentData}>
                      <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f4f4f5" />
                      <XAxis dataKey="name" axisLine={false} tickLine={false} />
                      <YAxis axisLine={false} tickLine={false} />
                      <Tooltip />
                      <Bar dataKey="value" fill="#18181b" radius={[4, 4, 0, 0]}>
                        {paymentData.map((_, i) => <Cell key={i} fill={COLORS[i % COLORS.length]} />)}
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
              <div className="bg-white p-6 rounded-3xl border border-zinc-100 shadow-sm">
                <h3 className="font-bold mb-6">Desglose de Montos</h3>
                <div className="space-y-4">
                  {paymentData.map((p, i) => (
                    <div key={i} className="flex items-center justify-between p-4 bg-zinc-50 rounded-2xl">
                      <div className="flex items-center gap-3">
                        <div className="w-3 h-3 rounded-full" style={{ backgroundColor: COLORS[i % COLORS.length] }}></div>
                        <span className="font-bold">{p.name}</span>
                      </div>
                      <span className="font-black">${p.value.toFixed(2)}</span>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          )}
        </motion.div>
      </AnimatePresence>
      </div>
    </div>
  );
};
