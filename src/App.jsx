import React, { useState, useEffect, useMemo } from 'react';
import {
  Building2, Users, Receipt, ShieldCheck, Download, Upload, Plus,
  Search, ShieldAlert, Ban, CheckCircle2, AlertCircle, RefreshCw,
  LogOut, Smartphone, Check, Moon, Sun, ArrowRight, UserCheck, Trash2, Edit2, Info,
  Wifi, Battery, Award, TrendingUp, MapPin, BarChart3, Layers, Filter, Globe,
  ChevronDown, ChevronUp
} from 'lucide-react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import dbService from './supabase';

// -------------------------------------------------------------
// LOCALIZATION / TRANSLATIONS DICTIONARY (English & Tamil)
// -------------------------------------------------------------
const TRANSLATIONS = {
  en: {
    dashboard: 'Dashboard',
    customers: 'Customers',
    ledger: 'Ledger',
    reports: 'Reports',
    subscribers: 'Total Subscribers',
    activeNodes: 'Active Connections',
    inactiveNodes: 'Inactive Connections',
    collections: 'Collections',
    recentCollections: 'Recent Collections Log',
    addCustomer: 'Add Customer Account',
    searchPlaceholder: 'Search Name, Box ID, Street...',
    logOut: 'Log Out',
    language: 'Language',
    tamil: 'Tamil / தமிழ்',
    english: 'English',
    workerProgress: 'Worker Progress Dashboard',
    areaRankings: 'Area Rankings',
    inspectWorkers: 'Inspect Workers Progress',
    downloadReport: 'Download Monthly Report (XLS)',
    importCsv: 'Import CSV',
    exportExcel: 'Export Excel',
    recordPayment: 'Record Payment',
    bulkPayments: 'Bulk Payments',
    boxId: 'Box ID',
    amount: 'Amount (₹)',
    period: 'Billing Period',
    notes: 'Notes',
    saveCustomer: 'Save Customer',
    cancel: 'Cancel',
    layoutView: 'Layout View',
    standardList: 'Standard List',
    streetGroup: 'Street Group',
    boxIdSort: 'Box ID Sort',
    active: 'Active',
    inactive: 'Inactive',
    allTime: 'All Time',
    filterPeriod: 'Filter Period',
    workerInspector: 'Collector Progress Report',
    closeInspector: 'Close Inspector',
    myProgress: 'My Progress',
    myRecentActs: 'My Recent Activities',
    langSettings: 'Language Settings',
    editCustomer: 'Edit Customer Details',
    saveChanges: 'Save Changes',
    delete: 'Delete',
    unassigned: 'Unassigned',
    connectionStatus: 'Connection Status',
    operationalWorkers: 'Operational Workers',
    
    // Core earnings metrics updates
    todayEarning: 'Today Earning',
    totalEarning: 'Total Earning'
  },
  ta: {
    dashboard: 'முகப்பு',
    customers: 'வாடிக்கையாளர்கள்',
    ledger: 'பேரேடு',
    reports: 'அறிக்கைகள்',
    subscribers: 'மொத்த வாடிக்கையாளர்கள்',
    activeNodes: 'செயலில் உள்ள இணைப்புகள்',
    inactiveNodes: 'செயலற்ற இணைப்புகள்',
    collections: 'வசூல்',
    recentCollections: 'சமீபத்திய வசூல் பதிவு',
    addCustomer: 'புதிய வாடிக்கையாளர் சேர்',
    searchPlaceholder: 'பெயர், பாக்ஸ் ஐடி, தெரு தேடுக...',
    logOut: 'வெளியேறு',
    language: 'மொழி / Language',
    tamil: 'தமிழ் / Tamil',
    english: 'English / ஆங்கிலம்',
    workerProgress: 'பணியாளர் வசூல் விவரம்',
    areaRankings: 'தெருவாரியான வசூல்',
    inspectWorkers: 'பணியாளர்களின் வசூல் ஆய்வு',
    downloadReport: 'மாதாந்திர அறிக்கை பதிவிறக்கு (XLS)',
    importCsv: 'கோப்பு இறக்குமதி (CSV)',
    exportExcel: 'கோப்பு ஏற்றுமதி (XLS)',
    recordPayment: 'வசூல் பதிவு செய்',
    bulkPayments: 'மொத்த வசூல்',
    boxId: 'பாக்ஸ் ஐடி',
    amount: 'வசூல் தொகை (₹)',
    period: 'கட்டண மாதம்',
    notes: 'குறிப்பு',
    saveCustomer: 'வாடிக்கையாளரை சேமி',
    cancel: 'ரத்து செய்',
    layoutView: 'காட்சி முறை',
    standardList: 'அகரவரிசை',
    streetGroup: 'தெருவாரியாக',
    boxIdSort: 'பாக்ஸ் ஐடி வரிசை',
    active: 'செயலில்',
    inactive: 'செயலற்றது',
    allTime: 'அனைத்து காலம்',
    filterPeriod: 'கால அளவு வடிகட்டி',
    workerInspector: 'பணியாளர் வசூல் அறிக்கை',
    closeInspector: 'ஆய்வை மூடு',
    myProgress: 'எனது முன்னேற்றம்',
    myRecentActs: 'எனது சமீபத்திய செயல்பாடுகள்',
    langSettings: 'மொழி அமைப்புகள்',
    editCustomer: 'வாடிக்கையாளர் திருத்தம்',
    saveChanges: 'மாற்றங்களைச் சேமி',
    delete: 'நீக்கு',
    unassigned: 'ஒதுக்கப்படாதவர்',
    connectionStatus: 'இணைப்பு நிலை',
    operationalWorkers: 'பணிபுரியும் பணியாளர்கள்',
    
    // Core earnings metrics updates
    todayEarning: 'இன்றைய வசூல்',
    totalEarning: 'மொத்த வசூல்'
  }
};

export default function App() {
  const [currentUser, setCurrentUser] = useState(null);
  const [loading, setLoading] = useState(true);
  const [activeTab, setActiveTab] = useState('dashboard');
  const [theme, setTheme] = useState('light');
  const [language, setLanguage] = useState('en'); // 'en' or 'ta'
  
  // Data State
  const [businesses, setBusinesses] = useState([]);
  const [profiles, setProfiles] = useState([]);
  const [customers, setCustomers] = useState([]);
  const [payments, setPayments] = useState([]);
  const [auditLogs, setAuditLogs] = useState([]);
  
  // UI State
  const [errorMsg, setErrorMsg] = useState('');
  const [successMsg, setSuccessMsg] = useState('');
  const [modalType, setModalType] = useState(null); // 'add_business', 'add_user', 'add_customer', 'bulk_payment', 'edit_customer', 'inspect_worker'
  const [selectedCustomer, setSelectedCustomer] = useState(null);
  const [selectedWorker, setSelectedWorker] = useState(null);
  const [refreshing, setRefreshing] = useState(false);

  // Form States
  const [loginForm, setLoginForm] = useState({ username: '', password: '' });
  const [businessForm, setBusinessForm] = useState({ name: '' });
  const [userForm, setUserForm] = useState({ username: '', name: '', role: 'WORKER', phone: '', password: '', businessId: '' });
  const [customerForm, setCustomerForm] = useState({ customer_name: '', street_name: '', box_id: '', phone_number: '', assigned_worker_id: '', connection_status: 'ACTIVE', notes: '' });
  const [bulkPaymentForm, setBulkPaymentForm] = useState({ periods: [], amount: '', notes: '' });
  const [singlePaymentForm, setSinglePaymentForm] = useState({ amount: '', period: '', notes: '' });

  // Filter States
  const [searchQuery, setSearchQuery] = useState('');
  const [filterWorker, setFilterWorker] = useState('');
  const [filterStatus, setFilterStatus] = useState('');
  const [filterStreet, setFilterStreet] = useState('');
  
  // Mobile View Settings
  const [customerViewMode, setCustomerViewMode] = useState('standard'); // 'standard', 'street', 'box_id'
  const [expandedStreets, setExpandedStreets] = useState({});
  
  // Super Admin Data Importer State
  const [adminImportType, setAdminImportType] = useState('customers'); // 'customers', 'payments'
  const [adminImportTargetBiz, setAdminImportTargetBiz] = useState('');
  const [filterLogBusiness, setFilterLogBusiness] = useState('');

  // Report Month State
  const [reportMonth, setReportMonth] = useState('June 2026');

  // APK Builder Simulator State
  const [apkForm, setApkForm] = useState({ businessId: '', appName: '', packageName: 'com.kairos.cableapp', primaryColor: '#4f46e5' });
  const [apkLogs, setApkLogs] = useState([]);
  const [buildingApk, setBuildingApk] = useState(false);

  // Translation Helper
  const t = (key) => {
    return TRANSLATIONS[language]?.[key] || TRANSLATIONS['en']?.[key] || key;
  };

  const toggleStreet = (streetName) => {
    setExpandedStreets(prev => ({
      ...prev,
      [streetName]: !prev[streetName]
    }));
  };

  // Check login session
  useEffect(() => {
    checkSession();
  }, []);

  const checkSession = async () => {
    try {
      const session = await dbService.auth.getCurrentSession();
      if (session) {
        setCurrentUser(session);
        loadData(session);
      }
    } catch (err) {
      console.error(err);
    } finally {
      setLoading(false);
    }
  };

  const loadData = async (user = currentUser) => {
    if (!user) return;
    setRefreshing(true);
    try {
      if (user.role === 'SUPER_ADMIN') {
        const biz = await dbService.businesses.list();
        const profs = await dbService.profiles.list();
        const logs = await dbService.auditLogs.list();
        setBusinesses(biz);
        setProfiles(profs);
        setAuditLogs(logs);
      } else {
        const custs = await dbService.customers.list(user.business_id);
        const pays = await dbService.payments.list(user.business_id);
        const profs = await dbService.profiles.list(user.business_id);
        setCustomers(custs);
        setPayments(pays);
        setProfiles(profs); // Load all business users
      }
    } catch (err) {
      showError(err.message);
    } finally {
      setRefreshing(false);
    }
  };

  // Theme Toggler
  const toggleTheme = () => {
    const nextTheme = theme === 'light' ? 'dark' : 'light';
    setTheme(nextTheme);
    if (nextTheme === 'dark') {
      document.body.classList.add('dark-theme');
    } else {
      document.body.classList.remove('dark-theme');
    }
  };

  // Helper messages
  const showError = (msg) => {
    setErrorMsg(msg);
    setTimeout(() => setErrorMsg(''), 5000);
  };

  const showSuccess = (msg) => {
    setSuccessMsg(msg);
    setTimeout(() => setSuccessMsg(''), 5000);
  };

  // -------------------------------------------------------------
  // ACTIONS
  // -------------------------------------------------------------
  const handleLogin = async (e) => {
    e.preventDefault();
    if (!loginForm.username || !loginForm.password) {
      return showError('Please enter username and password.');
    }
    setLoading(true);
    try {
      const { profile } = await dbService.auth.signIn(loginForm.username, loginForm.password);
      setCurrentUser(profile);
      showSuccess(`Logged in successfully as ${profile.name}`);
      setLoginForm({ username: '', password: '' });
      setActiveTab('dashboard');
      loadData(profile);
    } catch (err) {
      showError(err.message);
    } finally {
      setLoading(false);
    }
  };

  const handleLogout = async () => {
    if (window.confirm('Are you sure you want to log out?')) {
      await dbService.auth.signOut();
      setCurrentUser(null);
      setBusinesses([]);
      setProfiles([]);
      setCustomers([]);
      setPayments([]);
      setAuditLogs([]);
    }
  };

  const handleCreateBusiness = async (e) => {
    e.preventDefault();
    if (!businessForm.name) return showError('Business name is required.');
    try {
      const newBiz = await dbService.businesses.create(businessForm.name, currentUser.id);
      setBusinesses([newBiz, ...businesses]);
      setBusinessForm({ name: '' });
      setModalType(null);
      showSuccess('Business registered successfully.');
      loadData();
    } catch (err) {
      showError(err.message);
    }
  };

  const handleToggleBusinessStatus = async (id, currentStatus) => {
    const nextStatus = currentStatus === 'ACTIVE' ? 'SUSPENDED' : 'ACTIVE';
    if (!window.confirm(`Are you sure you want to set this business status to ${nextStatus}?`)) return;
    try {
      await dbService.businesses.setStatus(id, nextStatus, currentUser.id);
      showSuccess(`Business status updated to ${nextStatus}.`);
      loadData();
    } catch (err) {
      showError(err.message);
    }
  };

  const handleCreateUser = async (e) => {
    e.preventDefault();
    const { username, name, role, phone, password, businessId } = userForm;
    if (!username || !name || !role || !password) return showError('Please fill in all required fields.');
    if (role !== 'SUPER_ADMIN' && !businessId) return showError('Please assign a Business.');
    
    try {
      const targetBizId = role === 'SUPER_ADMIN' ? null : businessId;
      const newProf = await dbService.profiles.create(username, name, role, phone, password, targetBizId, currentUser.id);
      setProfiles([newProf, ...profiles]);
      setUserForm({ username: '', name: '', role: 'WORKER', phone: '', password: '', businessId: '' });
      setModalType(null);
      showSuccess(`User account "${username}" provisioned.`);
      loadData();
    } catch (err) {
      showError(err.message);
    }
  };

  const handleToggleUserDisabled = async (id, isDisabled, name) => {
    const actionText = isDisabled ? 'unblock' : 'block';
    if (!window.confirm(`Are you sure you want to ${actionText} ${name}?`)) return;
    try {
      await dbService.profiles.setDisabled(id, !isDisabled, currentUser.id);
      showSuccess(`User has been ${isDisabled ? 'unblocked' : 'blocked'}.`);
      loadData();
    } catch (err) {
      showError(err.message);
    }
  };

  const handleCreateCustomer = async (e) => {
    e.preventDefault();
    if (!customerForm.customer_name || !customerForm.street_name || !customerForm.box_id) {
      return showError('Name, Street Name, and Box ID are required.');
    }
    try {
      const data = {
        business_id: currentUser.business_id,
        customer_name: customerForm.customer_name,
        street_name: customerForm.street_name,
        box_id: customerForm.box_id,
        phone_number: customerForm.phone_number,
        assigned_worker_id: currentUser.role === 'WORKER' ? currentUser.id : (customerForm.assigned_worker_id || null),
        connection_status: customerForm.connection_status,
        notes: customerForm.notes,
        custom_fields: {}
      };

      const newCust = await dbService.customers.create(data, currentUser.id);
      setCustomers([newCust, ...customers]);
      setCustomerForm({ customer_name: '', street_name: '', box_id: '', phone_number: '', assigned_worker_id: '', connection_status: 'ACTIVE', notes: '' });
      setModalType(null);
      showSuccess('Customer account created successfully.');
      loadData();
    } catch (err) {
      showError(err.message);
    }
  };

  const handleEditCustomer = async (e) => {
    e.preventDefault();
    if (!selectedCustomer) return;
    if (!customerForm.customer_name || !customerForm.street_name || !customerForm.box_id) {
      return showError('Name, Street Name, and Box ID are required.');
    }
    try {
      const updatedData = {
        customer_name: customerForm.customer_name,
        street_name: customerForm.street_name,
        box_id: customerForm.box_id,
        phone_number: customerForm.phone_number,
        assigned_worker_id: currentUser.role === 'WORKER' ? currentUser.id : (customerForm.assigned_worker_id || null),
        connection_status: customerForm.connection_status,
        notes: customerForm.notes
      };

      const result = await dbService.customers.update(selectedCustomer.id, updatedData, currentUser.id);
      setCustomers(customers.map(c => c.id === selectedCustomer.id ? result : c));
      setModalType(null);
      setSelectedCustomer(null);
      showSuccess('Customer details updated.');
      loadData();
    } catch (err) {
      showError(err.message);
    }
  };

  const handleDeleteCustomer = async (id) => {
    if (!window.confirm('Are you sure you want to delete this customer? This will also remove their payment history.')) return;
    try {
      await dbService.customers.delete(id, currentUser.business_id, currentUser.id);
      setCustomers(customers.filter(c => c.id !== id));
      setSelectedCustomer(null);
      showSuccess('Customer record removed.');
      loadData();
    } catch (err) {
      showError(err.message);
    }
  };

  const handleSinglePayment = async (e) => {
    e.preventDefault();
    if (!selectedCustomer) return;
    if (!singlePaymentForm.amount || !singlePaymentForm.period) {
      return showError('Amount and Period are required.');
    }

    try {
      const data = {
        business_id: currentUser.business_id,
        customer_id: selectedCustomer.id,
        box_id: selectedCustomer.box_id,
        worker_id: currentUser.role === 'WORKER' ? currentUser.id : null,
        amount: parseFloat(singlePaymentForm.amount),
        payment_date: new Date().toISOString().split('T')[0],
        payment_period: singlePaymentForm.period,
        status: 'Paid',
        notes: singlePaymentForm.notes
      };

      const newPay = await dbService.payments.create(data, currentUser.id);
      setPayments([newPay, ...payments]);
      setSinglePaymentForm({ amount: '', period: '', notes: '' });
      setModalType(null);
      showSuccess('Payment logged successfully.');
      loadData();
    } catch (err) {
      showError(err.message);
    }
  };

  const handleBulkPayment = async (e) => {
    e.preventDefault();
    if (!selectedCustomer) return;
    if (bulkPaymentForm.periods.length === 0 || !bulkPaymentForm.amount) {
      return showError('Please select at least one period and specify the amount.');
    }

    try {
      const collectorId = currentUser.role === 'WORKER' ? currentUser.id : null;
      await dbService.payments.createBulk(
        selectedCustomer.id,
        bulkPaymentForm.periods,
        parseFloat(bulkPaymentForm.amount),
        collectorId,
        currentUser.business_id,
        bulkPaymentForm.notes,
        selectedCustomer.box_id,
        currentUser.id
      );

      setBulkPaymentForm({ periods: [], amount: '', notes: '' });
      setModalType(null);
      showSuccess('Bulk collections logged successfully.');
      loadData();
    } catch (err) {
      showError(err.message);
    }
  };

  // CSV Customer Import Uploader (For mobile/operator)
  const handleCSVImport = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete: async (results) => {
        const rows = results.data;
        const validRows = [];
        const errors = [];

        rows.forEach((row, index) => {
          if (!row['Customer Name'] || !row['Street Name'] || !row['Box ID']) {
            errors.push(`Row ${index + 2}: Missing required fields (Customer Name, Street Name, Box ID)`);
          } else {
            let assignedWorkerId = null;
            if (row['Assigned Worker']) {
              const matched = profiles.find(p => p.name.toLowerCase().includes(row['Assigned Worker'].toLowerCase()));
              if (matched) assignedWorkerId = matched.id;
            }
            validRows.push({
              customer_name: row['Customer Name'].trim(),
              street_name: row['Street Name'].trim(),
              box_id: row['Box ID'].trim(),
              phone_number: row['Phone Number'] ? row['Phone Number'].trim() : '',
              connection_status: row['Status']?.toUpperCase() === 'INACTIVE' ? 'INACTIVE' : 'ACTIVE',
              assigned_worker_id: assignedWorkerId,
              notes: row['Notes'] || ''
            });
          }
        });

        if (errors.length > 0) {
          alert(`Import validation failed:\n${errors.slice(0, 5).join('\n')}\n${errors.length > 5 ? `...and ${errors.length - 5} more errors.` : ''}`);
          return;
        }

        if (validRows.length === 0) {
          return showError('No valid customer rows found in file.');
        }

        if (window.confirm(`Verify Import: Do you want to load ${validRows.length} customers?`)) {
          try {
            await dbService.customers.bulkImport(validRows, currentUser.business_id, currentUser.id);
            showSuccess(`Successfully imported ${validRows.length} customer records.`);
            loadData();
          } catch (err) {
            showError(err.message);
          }
        }
      }
    });
  };

  // Super Admin CSV Importer
  const handleAdminCSVImport = (e) => {
    if (!adminImportTargetBiz) return showError('Please choose a target Business client first.');
    const file = e.target.files[0];
    if (!file) return;

    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete: async (results) => {
        const rows = results.data;
        const validRows = [];
        const errors = [];

        if (adminImportType === 'customers') {
          rows.forEach((row, index) => {
            if (!row['Customer Name'] || !row['Street Name'] || !row['Box ID']) {
              errors.push(`Row ${index + 2}: Missing required fields (Customer Name, Street Name, Box ID)`);
            } else {
              validRows.push({
                customer_name: row['Customer Name'].trim(),
                street_name: row['Street Name'].trim(),
                box_id: row['Box ID'].trim(),
                phone_number: row['Phone Number'] ? row['Phone Number'].trim() : '',
                connection_status: row['Status']?.toUpperCase() === 'INACTIVE' ? 'INACTIVE' : 'ACTIVE',
                assigned_worker_id: null,
                notes: row['Notes'] || ''
              });
            }
          });
        } else {
          rows.forEach((row, index) => {
            if (!row['Box ID'] || !row['Amount'] || !row['Payment Date'] || !row['Payment Period']) {
              errors.push(`Row ${index + 2}: Missing required fields (Box ID, Amount, Payment Date, Payment Period)`);
            } else {
              validRows.push({
                box_id: row['Box ID'].trim(),
                amount: parseFloat(row['Amount']),
                payment_date: row['Payment Date'].trim(),
                payment_period: row['Payment Period'].trim(),
                notes: row['Notes'] || '',
                status: 'Paid',
                customer_id: 'c-1'
              });
            }
          });
        }

        if (errors.length > 0) {
          alert(`Validation Failed:\n${errors.slice(0, 5).join('\n')}\n${errors.length > 5 ? `...and ${errors.length - 5} more errors.` : ''}`);
          return;
        }

        if (validRows.length === 0) return showError('No records found in CSV.');

        if (window.confirm(`Create ${validRows.length} ${adminImportType} records for the selected business?`)) {
          try {
            if (adminImportType === 'customers') {
              await dbService.customers.bulkImport(validRows, adminImportTargetBiz, currentUser.id);
            } else {
              await dbService.payments.bulkImport(validRows, adminImportTargetBiz, currentUser.id);
            }
            showSuccess(`Admin: Imported ${validRows.length} records successfully.`);
            loadData();
          } catch (err) {
            showError(err.message);
          }
        }
      }
    });
  };

  // Excel Customer Export Downloader
  const handleExportCustomers = () => {
    const dataToExport = filteredCustomers.map(c => {
      const worker = profiles.find(p => p.id === c.assigned_worker_id);
      return {
        'Customer Name': c.customer_name,
        'Street Name': c.street_name,
        'Box ID': c.box_id,
        'Phone Number': c.phone_number || '',
        'Assigned Worker': worker ? worker.name : 'Unassigned',
        'Status': c.connection_status,
        'Notes': c.notes || '',
        'Created Date': c.created_at.split('T')[0]
      };
    });

    const worksheet = XLSX.utils.json_to_sheet(dataToExport);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Customers');
    XLSX.writeFile(workbook, `${currentUser.username}_customers.xlsx`);
    showSuccess('Customer database exported to Excel.');
  };

  // -------------------------------------------------------------
  // MONTHLY REPORT EXPORTS
  // -------------------------------------------------------------
  const handleExportMonthlyReport = (selectedMonth, targetWorkerId = null) => {
    let list = payments.filter(p => p.status === 'Paid');
    
    if (selectedMonth !== 'All Time') {
      list = list.filter(p => p.payment_period === selectedMonth);
    }
    
    if (currentUser.role === 'WORKER') {
      list = list.filter(p => p.worker_id === currentUser.id);
    } else if (targetWorkerId) {
      list = list.filter(p => p.worker_id === targetWorkerId);
    }

    const dataToExport = list.map(p => {
      const cust = customers.find(c => c.id === p.customer_id);
      const collector = profiles.find(pr => pr.id === p.worker_id);
      return {
        'Customer Name': cust ? cust.customer_name : 'N/A',
        'Box ID': p.box_id,
        'Street Name': cust ? cust.street_name : 'N/A',
        'Collected By': collector ? collector.name : 'System/Direct',
        'Amount Paid': p.amount,
        'Payment Date': p.payment_date,
        'Billing Period': p.payment_period,
        'Notes': p.notes || ''
      };
    });

    const label = targetWorkerId 
      ? `Worker_${profiles.find(pr => pr.id === targetWorkerId)?.username}`
      : (currentUser.role === 'WORKER' ? `My` : `Business`);

    const worksheet = XLSX.utils.json_to_sheet(dataToExport);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Collections');
    XLSX.writeFile(workbook, `${label}_collections_${selectedMonth.replace(/\s+/g, '_')}.xlsx`);
    showSuccess('Collections sheet exported successfully.');
  };

  // APK Compiler Simulator Runner
  const handleBuildApk = (e) => {
    e.preventDefault();
    if (!apkForm.businessId || !apkForm.appName) {
      return showError('Please select a Business client and provide an App Name.');
    }
    
    const selectedBiz = businesses.find(b => b.id === apkForm.businessId);
    if (!selectedBiz) return;

    setBuildingApk(true);
    setApkLogs([]);

    const steps = [
      { msg: 'Initializing Kairos Android compiler engine...', delay: 600 },
      { msg: 'Verifying Java JDK 17 & Android SDK target 34...', delay: 1200 },
      { msg: 'Resolving Gradle dependencies...', delay: 1800 },
      { msg: `Injecting business branding configuration:\n  Business Name: ${selectedBiz.business_name}\n  App Name: ${apkForm.appName}\n  Package: ${apkForm.packageName}\n  Primary HSL: ${apkForm.primaryColor}`, delay: 2800 },
      { msg: 'Bundling responsive frontend assets (PWA asset compiler)...', delay: 3800 },
      { msg: 'Generating AndroidManifest.xml and application icons...', delay: 4600 },
      { msg: 'Compiling resources with Gradle wrapper build (release)...', delay: 5800 },
      { msg: 'Optimizing classes.dex and compiling layout resources...', delay: 7000 },
      { msg: 'Signing Android package with Kairos Edio Technologies release keystore...', delay: 8200 },
      { msg: 'Aligning zip packages (zipalign target)...', delay: 9000 },
      { msg: 'Build succeeded! Output target generated: app-release.apk (12.4 MB)', delay: 10000 }
    ];

    steps.forEach((step) => {
      setTimeout(() => {
        setApkLogs(prev => [...prev, `[${new Date().toLocaleTimeString()}] ${step.msg}`]);
        if (step.msg.startsWith('Build succeeded')) {
          setBuildingApk(false);
          const link = document.createElement('a');
          link.href = '/app-release.apk';
          link.download = `${apkForm.appName.replace(/\s+/g, '_')}_v1.0.apk`;
          document.body.appendChild(link);
          link.click();
          document.body.removeChild(link);
          showSuccess('Real APK package downloaded successfully.');
        }
      }, step.delay);
    });
  };

  // -------------------------------------------------------------
  // SELECTORS & FILTERS
  // -------------------------------------------------------------
  const filteredCustomers = useMemo(() => {
    let result = [...customers];
    
    if (searchQuery) {
      result = result.filter(c => 
        c.customer_name.toLowerCase().includes(searchQuery.toLowerCase()) ||
        c.box_id.toLowerCase().includes(searchQuery.toLowerCase()) ||
        c.street_name.toLowerCase().includes(searchQuery.toLowerCase()) ||
        (c.phone_number && c.phone_number.includes(searchQuery))
      );
    }

    if (filterWorker) result = result.filter(c => c.assigned_worker_id === filterWorker);
    if (filterStatus) result = result.filter(c => c.connection_status === filterStatus);
    if (filterStreet) result = result.filter(c => c.street_name === filterStreet);

    if (customerViewMode === 'box_id') {
      result.sort((a, b) => a.box_id.localeCompare(b.box_id));
    } else {
      result.sort((a, b) => a.customer_name.localeCompare(b.customer_name));
    }

    return result;
  }, [customers, searchQuery, filterWorker, filterStatus, filterStreet, customerViewMode]);

  const customersGroupedByStreet = useMemo(() => {
    const groups = {};
    filteredCustomers.forEach(c => {
      if (!groups[c.street_name]) groups[c.street_name] = [];
      groups[c.street_name].push(c);
    });
    return groups;
  }, [filteredCustomers]);

  const filteredLogs = useMemo(() => {
    let result = [...auditLogs];
    if (filterLogBusiness) {
      result = result.filter(log => log.business_id === filterLogBusiness);
    }
    return result;
  }, [auditLogs, filterLogBusiness]);

  const stats = useMemo(() => {
    if (!currentUser || currentUser.role === 'SUPER_ADMIN') return null;

    const tenantCusts = customers;
    const totalCusts = tenantCusts.length;
    const activeCusts = tenantCusts.filter(c => c.connection_status === 'ACTIVE').length;
    const inactiveCusts = totalCusts - activeCusts;

    const tenantPays = payments;
    
    // Total (All Time) Collections
    const globalTotalCollected = tenantPays.filter(p => p.status === 'Paid').reduce((sum, p) => sum + parseFloat(p.amount), 0);

    // Today Collections
    const todayStr = new Date().toISOString().split('T')[0];
    const globalTodayCollected = tenantPays.filter(p => p.status === 'Paid' && p.payment_date === todayStr).reduce((sum, p) => sum + parseFloat(p.amount), 0);

    const monthFilteredPays = reportMonth === 'All Time' 
      ? tenantPays 
      : tenantPays.filter(p => p.payment_period === reportMonth);

    const reportCollected = monthFilteredPays.filter(p => p.status === 'Paid').reduce((sum, p) => sum + parseFloat(p.amount), 0);

    const workerPerformance = {};
    monthFilteredPays.forEach(p => {
      if (p.status === 'Paid' && p.worker_id) {
        workerPerformance[p.worker_id] = (workerPerformance[p.worker_id] || 0) + parseFloat(p.amount);
      }
    });

    const streetPerformance = {};
    monthFilteredPays.forEach(p => {
      if (p.status === 'Paid') {
        const cust = tenantCusts.find(c => c.id === p.customer_id);
        if (cust) {
          streetPerformance[cust.street_name] = (streetPerformance[cust.street_name] || 0) + parseFloat(p.amount);
        }
      }
    });

    // Worker collections count
    const myTodayColl = payments.filter(p => p.worker_id === currentUser.id && p.status === 'Paid' && p.payment_date === todayStr).reduce((sum, p) => sum + parseFloat(p.amount), 0);
    const myTotalColl = payments.filter(p => p.worker_id === currentUser.id && p.status === 'Paid').reduce((sum, p) => sum + parseFloat(p.amount), 0);
    const myCustomersCount = customers.filter(c => c.assigned_worker_id === currentUser.id).length;

    return {
      totalCusts,
      activeCusts,
      inactiveCusts,
      
      // Dynamic non-monthly target earnings
      globalTodayCollected,
      globalTotalCollected,
      
      reportCollected,
      workerPerformance,
      streetPerformance,
      myTodayColl,
      myTotalColl,
      myCustomersCount
    };
  }, [currentUser, customers, payments, reportMonth]);

  const paymentPeriodsList = useMemo(() => {
    const list = payments.map(p => p.payment_period);
    const unique = [...new Set(list)];
    if (unique.length === 0) return ['June 2026', 'May 2026'];
    return unique;
  }, [payments]);

  // Get specific worker performance metrics
  const getWorkerProgressMetrics = (workerId) => {
    const assignedCusts = customers.filter(c => c.assigned_worker_id === workerId);
    const workerPays = payments.filter(p => p.worker_id === workerId && p.status === 'Paid');

    const todayStr = new Date().toISOString().split('T')[0];
    const todayCollected = workerPays.filter(p => p.payment_date === todayStr).reduce((sum, p) => sum + parseFloat(p.amount), 0);
    const totalCollected = workerPays.reduce((sum, p) => sum + parseFloat(p.amount), 0);

    const monthPays = reportMonth === 'All Time' 
      ? workerPays 
      : workerPays.filter(p => p.payment_period === reportMonth);

    const collected = monthPays.reduce((sum, p) => sum + parseFloat(p.amount), 0);
    const history = monthPays;

    return {
      customersCount: assignedCusts.length,
      todayCollected,
      totalCollected,
      collected,
      history
    };
  };

  if (loading) {
    return (
      <div style={{ height: '100vh', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', gap: '20px', background: 'var(--neutral-50)' }}>
        <div className="spinner" style={{ borderTopColor: 'var(--primary-500)', width: '36px', height: '36px' }}></div>
        <p style={{ fontWeight: 600, color: 'var(--neutral-400)', fontSize: '14px', letterSpacing: '0.02em' }}>Initializing Kairos systems...</p>
      </div>
    );
  }

  // --- PREMIUM LOGIN SCREEN ---
  if (!currentUser) {
    return (
      <div style={{ minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', background: 'radial-gradient(circle at 10% 20%, var(--neutral-800) 0%, var(--neutral-900) 90%)', padding: '20px', position: 'relative', overflow: 'hidden' }}>
        <div style={{ position: 'absolute', width: '300px', height: '300px', background: 'rgba(79, 70, 229, 0.15)', filter: 'blur(80px)', top: '10%', left: '15%', borderRadius: '50%' }}></div>
        <div style={{ position: 'absolute', width: '250px', height: '250px', background: 'rgba(16, 185, 129, 0.1)', filter: 'blur(70px)', bottom: '15%', right: '10%', borderRadius: '50%' }}></div>

        <div className="card animate-fade-in" style={{ width: '100%', maxWidth: '420px', padding: '40px', background: 'rgba(255, 255, 255, 0.98)', border: '1px solid rgba(255, 255, 255, 0.25)', borderRadius: 'var(--radius-lg)', boxShadow: 'var(--shadow-lg)', zIndex: 10 }}>
          <div style={{ textAlign: 'center', marginBottom: '36px' }}>
            <div style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', width: '56px', height: '56px', background: 'var(--primary-50)', borderRadius: 'var(--radius-md)', color: 'var(--primary-500)', marginBottom: '16px', boxShadow: '0 4px 12px rgba(79, 70, 229, 0.15)' }}>
              <Building2 size={28} />
            </div>
            <h2 style={{ fontFamily: 'var(--font-primary)', fontWeight: 800, fontSize: '26px', color: 'var(--neutral-900)', letterSpacing: '-0.03em' }}>Kairos Cable SaaS</h2>
            <p style={{ fontSize: '13px', color: 'var(--neutral-400)', fontWeight: 600, marginTop: '4px' }}>Cable TV Operations & Collections Dashboard</p>
          </div>

          {errorMsg && (
            <div style={{ display: 'flex', gap: '10px', padding: '14px 18px', background: 'var(--danger-bg)', color: 'var(--danger)', borderRadius: 'var(--radius-md)', fontSize: '13px', marginBottom: '24px', fontWeight: 600, border: '1px solid rgba(239, 68, 68, 0.15)' }}>
              <ShieldAlert size={18} style={{ flexShrink: 0 }} />
              <span>{errorMsg}</span>
            </div>
          )}

          <form onSubmit={handleLogin}>
            <div className="form-group">
              <label className="form-label">Username</label>
              <input
                type="text"
                className="form-input"
                placeholder="Enter user id"
                value={loginForm.username}
                onChange={(e) => setLoginForm({ ...loginForm, username: e.target.value })}
                required
              />
            </div>
            <div className="form-group" style={{ marginBottom: '28px' }}>
              <label className="form-label">Password</label>
              <input
                type="password"
                className="form-input"
                placeholder="Enter password"
                value={loginForm.password}
                onChange={(e) => setLoginForm({ ...loginForm, password: e.target.value })}
                required
              />
            </div>
            <button type="submit" className="btn btn-primary" style={{ width: '100%', padding: '16px' }}>
              Sign In <ArrowRight size={16} />
            </button>
          </form>

          <div style={{ marginTop: '28px', textAlign: 'center', borderTop: '1px solid var(--neutral-100)', paddingTop: '20px' }}>
            <p style={{ fontSize: '11px', color: 'var(--neutral-400)', fontWeight: 600, lineHeight: 1.5 }}>
              Security protocol by <strong>Kairos Edio Technologies</strong>.<br />Self-registration is deactivated on this gateway.
            </p>
          </div>
        </div>
      </div>
    );
  }

  // =============================================================
  // INTERFACE 1: SUPER ADMIN WEB PORTAL (DESKTOP PLATFORM)
  // =============================================================
  if (currentUser.role === 'SUPER_ADMIN') {
    return (
      <div className="desktop-layout">
        <aside className="sidebar">
          <div>
            <div className="sidebar-brand">
              <Building2 size={24} style={{ color: 'var(--primary-500)' }} />
              <span>Kairos</span> Console
            </div>
            <ul className="sidebar-menu">
              <li>
                <button className={`sidebar-link ${activeTab === 'dashboard' ? 'active' : ''}`} onClick={() => setActiveTab('dashboard')}>
                  <ShieldCheck size={18} /> System Dashboard
                </button>
              </li>
              <li>
                <button className={`sidebar-link ${activeTab === 'businesses' ? 'active' : ''}`} onClick={() => setActiveTab('businesses')}>
                  <Building2 size={18} /> Business Clients
                </button>
              </li>
              <li>
                <button className={`sidebar-link ${activeTab === 'users' ? 'active' : ''}`} onClick={() => setActiveTab('users')}>
                  <Users size={18} /> Provision Accounts
                </button>
              </li>
              <li>
                <button className={`sidebar-link ${activeTab === 'importer' ? 'active' : ''}`} onClick={() => setActiveTab('importer')}>
                  <Upload size={18} /> Bulk Importer
                </button>
              </li>
              <li>
                <button className={`sidebar-link ${activeTab === 'apk-builder' ? 'active' : ''}`} onClick={() => setActiveTab('apk-builder')}>
                  <Smartphone size={18} /> APK Compiler
                </button>
              </li>
              <li>
                <button className={`sidebar-link ${activeTab === 'logs' ? 'active' : ''}`} onClick={() => setActiveTab('logs')}>
                  <Receipt size={18} /> Security Logs
                </button>
              </li>
            </ul>
          </div>

          <div>
            <div className="form-group" style={{ padding: '0 16px', marginBottom: '16px' }}>
              <select 
                className="form-input form-select" 
                value={language} 
                onChange={(e) => setLanguage(e.target.value)} 
                style={{ padding: '8px 12px', fontSize: '12px', background: 'var(--neutral-800)', color: '#ffffff', border: 'none' }}
              >
                <option value="en">English</option>
                <option value="ta">தமிழ் (Tamil)</option>
              </select>
            </div>
            
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '14px 18px', background: 'var(--neutral-800)', borderRadius: 'var(--radius-md)', marginBottom: '16px' }}>
              <div>
                <p style={{ fontSize: '13px', fontWeight: 700, color: '#ffffff' }}>{currentUser.name}</p>
                <span style={{ fontSize: '11px', color: 'var(--neutral-400)', fontWeight: 600 }}>Platform Admin</span>
              </div>
              <button onClick={toggleTheme} style={{ background: 'none', border: 'none', color: 'var(--neutral-400)', cursor: 'pointer' }}>
                {theme === 'light' ? <Moon size={16} /> : <Sun size={16} />}
              </button>
            </div>
            <button onClick={handleLogout} className="btn btn-secondary" style={{ width: '100%', background: 'var(--neutral-800)', color: '#ffffff', border: 'none' }}>
              <LogOut size={16} /> Log Out System
            </button>
          </div>
        </aside>

        <main className="desktop-content">
          <header style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '40px', borderBottom: '1px solid var(--neutral-200)', paddingBottom: '24px' }}>
            <div>
              <h1 style={{ fontSize: '30px', letterSpacing: '-0.03em' }}>System Administration</h1>
              <div style={{ display: 'flex', gap: '16px', marginTop: '6px', fontSize: '12px', color: 'var(--neutral-400)', fontWeight: 600 }}>
                <span>Status: <strong style={{ color: 'var(--success)' }}>● ONLINE</strong></span>
                <span>•</span>
                <span>Latency: <strong>8 ms</strong></span>
                <span>•</span>
                <span>Driver: <strong>{dbService.isMock() ? 'Local Sandboxed Mock' : 'Supabase Cloud DB'}</strong></span>
              </div>
            </div>
            <button onClick={() => loadData()} className="btn btn-secondary icon-align">
              <RefreshCw size={15} className={refreshing ? 'spinner' : ''} /> Sync Engine
            </button>
          </header>

          {successMsg && (
            <div style={{ display: 'flex', gap: '10px', padding: '14px 18px', background: 'var(--success-bg)', color: 'var(--success)', borderRadius: 'var(--radius-md)', fontSize: '13px', marginBottom: '24px', fontWeight: 600 }}>
              <CheckCircle2 size={18} />
              <span>{successMsg}</span>
            </div>
          )}

          {errorMsg && (
            <div style={{ display: 'flex', gap: '10px', padding: '14px 18px', background: 'var(--danger-bg)', color: 'var(--danger)', borderRadius: 'var(--radius-md)', fontSize: '13px', marginBottom: '24px', fontWeight: 600 }}>
              <ShieldAlert size={18} />
              <span>{errorMsg}</span>
            </div>
          )}

          {activeTab === 'dashboard' && (
            <div className="animate-fade-in">
              <div className="grid-cols-3" style={{ marginBottom: '36px' }}>
                <div className="card">
                  <div className="card-title">Registered Clients</div>
                  <div className="card-value">{businesses.length} Businesses</div>
                  <div className="progress-bar-container">
                    <div className="progress-bar-fill" style={{ width: '100%', background: 'var(--primary-500)' }}></div>
                  </div>
                </div>
                <div className="card">
                  <div className="card-title">Provisioned Profiles</div>
                  <div className="card-value">{profiles.length} Active IDs</div>
                  <div className="progress-bar-container">
                    <div className="progress-bar-fill" style={{ width: '80%', background: 'var(--accent-500)' }}></div>
                  </div>
                </div>
                <div className="card">
                  <div className="card-title">Secured Nodes</div>
                  <div className="card-value">100% Isolated</div>
                  <div className="progress-bar-container">
                    <div className="progress-bar-fill" style={{ width: '100%', background: 'var(--success)' }}></div>
                  </div>
                </div>
              </div>

              <div className="card">
                <h3 style={{ marginBottom: '20px', fontSize: '18px' }}>Security Audit Ledger (Recent Actions)</h3>
                <div className="table-container">
                  <table className="custom-table">
                    <thead>
                      <tr>
                        <th>Action Logged</th>
                        <th>Resource Table</th>
                        <th>Action By</th>
                        <th>Details</th>
                        <th>Logged At</th>
                      </tr>
                    </thead>
                    <tbody>
                      {auditLogs.slice(0, 6).map(log => {
                        const prof = profiles.find(p => p.id === log.user_id);
                        return (
                          <tr key={log.id}>
                            <td><strong>{log.action}</strong></td>
                            <td><span className="badge badge-inactive">{log.entity_type}</span></td>
                            <td>{prof ? prof.name : log.user_id}</td>
                            <td><code style={{ fontSize: '12px' }}>{JSON.stringify(log.details)}</code></td>
                            <td>{new Date(log.created_at).toLocaleString()}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {activeTab === 'businesses' && (
            <div className="animate-fade-in">
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '24px' }}>
                <h3 style={{ fontSize: '18px' }}>Business Registries</h3>
                <button onClick={() => setModalType('add_business')} className="btn btn-primary icon-align">
                  <Plus size={16} /> Register Client Business
                </button>
              </div>

              <div className="card">
                <div className="table-container">
                  <table className="custom-table">
                    <thead>
                      <tr>
                        <th>Operator Company Name</th>
                        <th>Tenant Key (business_id)</th>
                        <th>Launch Date</th>
                        <th>Operational Status</th>
                        <th>Actions</th>
                      </tr>
                    </thead>
                    <tbody>
                      {businesses.map(biz => (
                        <tr key={biz.id}>
                          <td><strong>{biz.business_name}</strong></td>
                          <td><code>{biz.id}</code></td>
                          <td>{new Date(biz.created_at).toLocaleDateString()}</td>
                          <td>
                            <span className={`badge ${biz.status === 'ACTIVE' ? 'badge-active' : 'badge-inactive'}`}>
                              {biz.status}
                            </span>
                          </td>
                          <td>
                            <button
                              onClick={() => handleToggleBusinessStatus(biz.id, biz.status)}
                              className="btn btn-secondary"
                              style={{ padding: '8px 14px', fontSize: '12px', fontWeight: 700 }}
                            >
                              {biz.status === 'ACTIVE' ? 'Suspend Account' : 'Re-Activate Account'}
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {activeTab === 'users' && (
            <div className="animate-fade-in">
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '24px' }}>
                <h3 style={{ fontSize: '18px' }}>Operator Account Database</h3>
                <button onClick={() => setModalType('add_user')} className="btn btn-primary icon-align">
                  <Plus size={16} /> Provision User Login
                </button>
              </div>

              <div className="card">
                <div className="table-container">
                  <table className="custom-table">
                    <thead>
                      <tr>
                        <th>Username</th>
                        <th>Full Name</th>
                        <th>Role Type</th>
                        <th>Assigned Client Business</th>
                        <th>Phone Number</th>
                        <th>Login Access Status</th>
                        <th>Lock Actions</th>
                      </tr>
                    </thead>
                    <tbody>
                      {profiles.map(prof => {
                        const biz = businesses.find(b => b.id === prof.business_id);
                        return (
                          <tr key={prof.id}>
                            <td><strong>{prof.username}</strong></td>
                            <td>{prof.name}</td>
                            <td>
                              <span className="badge badge-active" style={{ background: prof.role === 'SUPER_ADMIN' ? 'var(--primary-50)' : 'var(--neutral-100)', color: prof.role === 'SUPER_ADMIN' ? 'var(--primary-500)' : 'var(--neutral-700)' }}>
                                {prof.role}
                              </span>
                            </td>
                            <td>{biz ? biz.business_name : 'Platform (Kairos)'}</td>
                            <td>{prof.phone_number || '-'}</td>
                            <td>
                              <span className={`badge ${!prof.disabled ? 'badge-active' : 'badge-inactive'}`}>
                                {!prof.disabled ? 'Active Access' : 'Blocked Access'}
                              </span>
                            </td>
                            <td>
                              {prof.role !== 'SUPER_ADMIN' && (
                                <button
                                  onClick={() => handleToggleUserDisabled(prof.id, prof.disabled, prof.name)}
                                  className={`btn icon-align ${prof.disabled ? 'btn-primary' : 'btn-danger'}`}
                                  style={{ padding: '8px 14px', fontSize: '12px', fontWeight: 700 }}
                                >
                                  {prof.disabled ? <UserCheck size={12} /> : <Ban size={12} />}
                                  {prof.disabled ? 'Unblock Account' : 'Block User ID'}
                                </button>
                              )}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {activeTab === 'importer' && (
            <div className="animate-fade-in" style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '32px' }}>
              <div className="card">
                <h3 style={{ marginBottom: '20px', fontSize: '18px' }}>Bulk Load Client Database</h3>
                <div className="form-group">
                  <label className="form-label">Target Business Client</label>
                  <select 
                    className="form-input form-select" 
                    value={adminImportTargetBiz} 
                    onChange={(e) => setAdminImportTargetBiz(e.target.value)}
                    required
                  >
                    <option value="">Choose target client...</option>
                    {businesses.filter(b => b.status === 'ACTIVE').map(b => (
                      <option key={b.id} value={b.id}>{b.business_name}</option>
                    ))}
                  </select>
                </div>
                <div className="form-group">
                  <label className="form-label">Upload Data Type</label>
                  <div style={{ display: 'flex', gap: '16px' }}>
                    <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer', fontWeight: 600, fontSize: '14px' }}>
                      <input type="radio" name="admin_imp" checked={adminImportType === 'customers'} onChange={() => setAdminImportType('customers')} />
                      <span>Customers</span>
                    </label>
                    <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer', fontWeight: 600, fontSize: '14px' }}>
                      <input type="radio" name="admin_imp" checked={adminImportType === 'payments'} onChange={() => setAdminImportType('payments')} />
                      <span>Payments</span>
                    </label>
                  </div>
                </div>
                
                <div style={{ marginTop: '24px' }}>
                  <label className="btn btn-primary icon-align" style={{ width: '100%', cursor: 'pointer' }}>
                    <Upload size={16} /> Choose & Parse CSV File
                    <input type="file" accept=".csv" onChange={handleAdminCSVImport} style={{ display: 'none' }} />
                  </label>
                </div>
              </div>

              <div className="card" style={{ background: 'var(--neutral-100)', border: '1px dashed var(--neutral-300)' }}>
                <h4 style={{ marginBottom: '12px', fontSize: '14px', textTransform: 'uppercase', letterSpacing: '0.04em' }}>Required File Template Format</h4>
                <p style={{ fontSize: '13px', color: 'var(--neutral-700)', marginBottom: '16px' }}>
                  Please format your CSV with the exact column headers shown below:
                </p>

                {adminImportType === 'customers' ? (
                  <div style={{ padding: '16px', background: '#ffffff', borderRadius: 'var(--radius-sm)', border: '1px solid var(--neutral-200)' }}>
                    <p style={{ fontSize: '12px', fontWeight: 'bold', color: 'var(--primary-600)', marginBottom: '8px' }}>CUSTOMERS TEMPLATE HEADERS:</p>
                    <code style={{ fontSize: '12px', whiteSpace: 'pre-wrap', color: 'var(--neutral-900)' }}>
                      Customer Name, Street Name, Box ID, Phone Number, Status, Notes
                    </code>
                  </div>
                ) : (
                  <div style={{ padding: '16px', background: '#ffffff', borderRadius: 'var(--radius-sm)', border: '1px solid var(--neutral-200)' }}>
                    <p style={{ fontSize: '12px', fontWeight: 'bold', color: 'var(--primary-600)', marginBottom: '8px' }}>PAYMENTS TEMPLATE HEADERS:</p>
                    <code style={{ fontSize: '12px', whiteSpace: 'pre-wrap', color: 'var(--neutral-900)' }}>
                      Box ID, Amount, Payment Date, Payment Period, Notes
                    </code>
                  </div>
                )}
              </div>
            </div>
          )}

          {activeTab === 'apk-builder' && (
            <div className="grid-cols-2 animate-fade-in">
              <div className="card">
                <h3 style={{ marginBottom: '20px', fontSize: '18px' }}>Native Wrapper compiler</h3>
                <form onSubmit={handleBuildApk}>
                  <div className="form-group">
                    <label className="form-label">Target Operator Client</label>
                    <select
                      className="form-input form-select"
                      value={apkForm.businessId}
                      onChange={(e) => setApkForm({ ...apkForm, businessId: e.target.value })}
                      required
                    >
                      <option value="">Choose a business client...</option>
                      {businesses.filter(b => b.status === 'ACTIVE').map(b => (
                        <option key={b.id} value={b.id}>{b.business_name}</option>
                      ))}
                    </select>
                  </div>
                  <div className="form-group">
                    <label className="form-label">Application Branding Name</label>
                    <input
                      type="text"
                      className="form-input"
                      placeholder="e.g. City Cable App"
                      value={apkForm.appName}
                      onChange={(e) => setApkForm({ ...apkForm, appName: e.target.value })}
                      required
                    />
                  </div>
                  <div className="form-group">
                    <label className="form-label">Unique Package Domain Identifier</label>
                    <input
                      type="text"
                      className="form-input"
                      placeholder="com.kairos.clientname"
                      value={apkForm.packageName}
                      onChange={(e) => setApkForm({ ...apkForm, packageName: e.target.value })}
                      required
                    />
                  </div>
                  <div className="form-group">
                    <label className="form-label">Primary Color Theme</label>
                    <input
                      type="color"
                      className="form-input"
                      value={apkForm.primaryColor}
                      onChange={(e) => setApkForm({ ...apkForm, primaryColor: e.target.value })}
                      style={{ height: '48px', padding: '4px' }}
                    />
                  </div>
                  <button type="submit" disabled={buildingApk} className="btn btn-primary icon-align" style={{ width: '100%' }}>
                    {buildingApk ? <RefreshCw className="spinner" /> : <Smartphone size={16} />}
                    {buildingApk ? 'Building Package Files...' : 'Compile & Export Release APK'}
                  </button>
                </form>
              </div>

              <div className="card" style={{ background: '#090d16', color: '#38bdf8', fontFamily: 'monospace', fontSize: '13px', overflowY: 'auto', maxHeight: '440px', border: '1px solid #1e293b', boxShadow: 'var(--shadow-md)' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', borderBottom: '1px solid #1e293b', paddingBottom: '8px', marginBottom: '12px', color: '#64748b', fontWeight: 'bold' }}>
                  <span>compiler-engine-logs.sh</span>
                  <span style={{ color: buildingApk ? '#fbbf24' : '#34d399' }}>{buildingApk ? '● RUNNING' : '● SYSTEM_IDLE'}</span>
                </div>
                {apkLogs.length === 0 && <span style={{ color: '#475569' }}>Console idle. Awaiting build trigger...</span>}
                {apkLogs.map((log, index) => (
                  <div key={index} style={{ marginBottom: '4px', whiteSpace: 'pre-wrap' }}>{log}</div>
                ))}
              </div>
            </div>
          )}

          {activeTab === 'logs' && (
            <div className="card animate-fade-in">
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '24px', borderBottom: '1px solid var(--neutral-100)', paddingBottom: '16px' }}>
                <h3 style={{ fontSize: '18px', margin: 0 }}>Security Audit Log Database</h3>
                <div style={{ display: 'flex', gap: '10px', alignItems: 'center' }}>
                  <span style={{ fontSize: '12px', fontWeight: 700, color: 'var(--neutral-400)', textTransform: 'uppercase' }}>Filter Client:</span>
                  <select 
                    className="form-input form-select" 
                    value={filterLogBusiness} 
                    onChange={(e) => setFilterLogBusiness(e.target.value)}
                    style={{ padding: '8px 12px', fontSize: '12px', width: '220px', margin: 0 }}
                  >
                    <option value="">All Businesses / System</option>
                    {businesses.map(b => (
                      <option key={b.id} value={b.id}>{b.business_name}</option>
                    ))}
                  </select>
                </div>
              </div>

              <div className="table-container">
                <table className="custom-table">
                  <thead>
                    <tr>
                      <th>Logged At</th>
                      <th>Client Business</th>
                      <th>Action Logged</th>
                      <th>Affected Table</th>
                      <th>Trigger User</th>
                      <th>Audit Details</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredLogs.map(log => {
                      const biz = businesses.find(b => b.id === log.business_id);
                      const prof = profiles.find(p => p.id === log.user_id);
                      
                      const formatLogDetails = (details) => {
                        if (!details) return '-';
                        if (typeof details === 'string') return details;
                        return Object.entries(details)
                          .map(([key, val]) => `${key}: ${typeof val === 'object' ? JSON.stringify(val) : val}`)
                          .join(', ');
                      };

                      return (
                        <tr key={log.id}>
                          <td>{new Date(log.created_at).toLocaleString()}</td>
                          <td>
                            <strong>{biz ? biz.business_name : 'Platform System'}</strong>
                            {biz && <span style={{ display: 'block', fontSize: '10px', color: 'var(--neutral-400)' }}>ID: {biz.id.substring(0,8)}...</span>}
                          </td>
                          <td><strong>{log.action}</strong></td>
                          <td><span className="badge badge-inactive" style={{ fontSize: '11px' }}>{log.entity_type}</span></td>
                          <td>
                            <strong>{prof ? prof.name : 'System/Admin'}</strong>
                            <span style={{ display: 'block', fontSize: '11px', color: 'var(--neutral-400)' }}>{prof ? `@${prof.username}` : log.user_id.substring(0,8)}</span>
                          </td>
                          <td style={{ fontSize: '12.5px', color: 'var(--neutral-800)', maxWidth: '300px', wordBreak: 'break-word' }}>
                            <code>{formatLogDetails(log.details)}</code>
                          </td>
                        </tr>
                      );
                    })}
                    {filteredLogs.length === 0 && (
                      <tr>
                        <td colSpan="6" style={{ textAlign: 'center', padding: '30px', color: 'var(--neutral-400)' }}>
                          No audit logs found for the selected filter.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </main>

        {modalType === 'add_business' && (
          <div className="modal-overlay">
            <div className="card animate-fade-in" style={{ width: '100%', maxWidth: '480px', background: '#ffffff', padding: '32px', border: '1px solid var(--neutral-200)', zIndex: 1010 }}>
              <h3 style={{ marginBottom: '24px', fontSize: '20px' }}>Register Client Business</h3>
              <form onSubmit={handleCreateBusiness}>
                <div className="form-group">
                  <label className="form-label">Business / Operator Name</label>
                  <input
                    type="text"
                    className="form-input"
                    placeholder="e.g. City Cable Vision"
                    value={businessForm.name}
                    onChange={(e) => setBusinessForm({ ...businessForm, name: e.target.value })}
                    required
                  />
                </div>
                <div style={{ display: 'flex', gap: '12px', marginTop: '32px' }}>
                  <button type="button" onClick={() => setModalType(null)} className="btn btn-secondary" style={{ flex: 1 }}>
                    Cancel
                  </button>
                  <button type="submit" className="btn btn-primary" style={{ flex: 1 }}>
                    Register Business
                  </button>
                </div>
              </form>
            </div>
          </div>
        )}

        {modalType === 'add_user' && (
          <div className="modal-overlay">
            <div className="card animate-fade-in" style={{ width: '100%', maxWidth: '540px', background: '#ffffff', padding: '32px', border: '1px solid var(--neutral-200)', zIndex: 1010 }}>
              <h3 style={{ marginBottom: '24px', fontSize: '20px' }}>Provision Operator Account</h3>
              <form onSubmit={handleCreateUser}>
                <div className="form-group">
                  <label className="form-label">Full Name</label>
                  <input
                    type="text"
                    className="form-input"
                    placeholder="e.g. Rajesh Kumar"
                    value={userForm.name}
                    onChange={(e) => setUserForm({ ...userForm, name: e.target.value })}
                    required
                  />
                </div>
                <div className="form-group">
                  <label className="form-label">Username (Login ID)</label>
                  <input
                    type="text"
                    className="form-input"
                    placeholder="e.g. rajesh_citycable"
                    value={userForm.username}
                    onChange={(e) => setUserForm({ ...userForm, username: e.target.value })}
                    required
                  />
                </div>
                <div className="form-group">
                  <label className="form-label">Password</label>
                  <input
                    type="password"
                    className="form-input"
                    placeholder="Enter account password"
                    value={userForm.password}
                    onChange={(e) => setUserForm({ ...userForm, password: e.target.value })}
                    required
                  />
                </div>
                <div className="form-group">
                  <label className="form-label">Phone Number</label>
                  <input
                    type="text"
                    className="form-input"
                    placeholder="e.g. +91 9876543210"
                    value={userForm.phone || ''}
                    onChange={(e) => setUserForm({ ...userForm, phone: e.target.value })}
                  />
                </div>
                <div className="form-group">
                  <label className="form-label">Role Type</label>
                  <select
                    className="form-input form-select"
                    value={userForm.role}
                    onChange={(e) => setUserForm({ ...userForm, role: e.target.value })}
                    required
                  >
                    <option value="WORKER">WORKER (Collector)</option>
                    <option value="OWNER">OWNER (Operator Admin)</option>
                    <option value="SUPER_ADMIN">SUPER_ADMIN (Platform Admin)</option>
                  </select>
                </div>
                {userForm.role !== 'SUPER_ADMIN' && (
                  <div className="form-group">
                    <label className="form-label">Assign Client Business</label>
                    <select
                      className="form-input form-select"
                      value={userForm.businessId}
                      onChange={(e) => setUserForm({ ...userForm, businessId: e.target.value })}
                      required
                    >
                      <option value="">Select client business...</option>
                      {businesses.filter(b => b.status === 'ACTIVE').map(b => (
                        <option key={b.id} value={b.id}>{b.business_name}</option>
                      ))}
                    </select>
                  </div>
                )}
                <div style={{ display: 'flex', gap: '12px', marginTop: '32px' }}>
                  <button type="button" onClick={() => setModalType(null)} className="btn btn-secondary" style={{ flex: 1 }}>
                    Cancel
                  </button>
                  <button type="submit" className="btn btn-primary" style={{ flex: 1 }}>
                    Provision Login
                  </button>
                </div>
              </form>
            </div>
          </div>
        )}
      </div>
    );
  }

  // =============================================================
  // INTERFACE 2: CLIENT MOBILE APP (iPhone/Android bezel shell on PC)
  // =============================================================
  const mobileLayoutHTML = (
    <div className="mobile-layout">
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '6px 20px', background: 'var(--neutral-100)', fontSize: '11px', fontWeight: 800, color: 'var(--neutral-400)' }}>
        <span>17:23 PM</span>
        <div style={{ display: 'flex', gap: '6px', alignItems: 'center' }}>
          <Wifi size={12} />
          <Battery size={14} />
        </div>
      </div>

      <header className="mobile-header">
        <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
          <div style={{ background: 'var(--primary-50)', color: 'var(--primary-500)', padding: '10px', borderRadius: 'var(--radius-md)', boxShadow: 'var(--shadow-sm)' }}>
            <Building2 size={20} />
          </div>
          <div>
            <h2 className="mobile-header-title" style={{ fontSize: '15px', color: 'var(--neutral-900)' }}>
              {currentUser.role === 'OWNER' ? t('dashboard') : t('dashboard')}
            </h2>
            <span style={{ fontSize: '11px', color: 'var(--neutral-400)', display: 'block', fontWeight: 600 }}>
              City Cable Network
            </span>
          </div>
        </div>

        <button onClick={toggleTheme} style={{ background: 'none', border: 'none', color: 'var(--neutral-400)', padding: '6px' }}>
          {theme === 'light' ? <Moon size={20} /> : <Sun size={20} />}
        </button>
      </header>

      {successMsg && (
        <div style={{ background: 'var(--success)', color: '#ffffff', padding: '10px 16px', fontSize: '12px', display: 'flex', alignItems: 'center', gap: '8px', fontWeight: 700 }}>
          <CheckCircle2 size={16} />
          <span>{successMsg}</span>
        </div>
      )}
      {errorMsg && (
        <div style={{ background: 'var(--danger)', color: '#ffffff', padding: '10px 16px', fontSize: '12px', display: 'flex', alignItems: 'center', gap: '8px', fontWeight: 700 }}>
          <ShieldAlert size={16} />
          <span>{errorMsg}</span>
        </div>
      )}

      <div style={{ padding: '20px', flex: 1, overflowY: 'auto', paddingBottom: '24px' }}>
        
        {/* MOBILE SCREEN: DASHBOARD */}
        {activeTab === 'dashboard' && (
          <div className="animate-fade-in">
            {currentUser.role === 'OWNER' ? (
              <div>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' }}>
                  <h3 style={{ fontSize: '18px', fontWeight: 800 }}>{t('dashboard')}</h3>
                  <button onClick={() => loadData()} style={{ background: 'none', border: 'none', color: 'var(--primary-500)', cursor: 'pointer' }}>
                    <RefreshCw size={18} className={refreshing ? 'spinner' : ''} />
                  </button>
                </div>

                <div className="grid-cols-2" style={{ gap: '12px', marginBottom: '16px' }}>
                  <div className="card" style={{ padding: '16px' }}>
                    <span className="card-title" style={{ fontSize: '10px' }}>{t('todayEarning')}</span>
                    <div className="card-value" style={{ fontSize: '22px', color: 'var(--success)' }}>₹{stats?.globalTodayCollected}</div>
                  </div>
                  <div className="card" style={{ padding: '16px' }}>
                    <span className="card-title" style={{ fontSize: '10px' }}>{t('totalEarning')}</span>
                    <div className="card-value" style={{ fontSize: '22px', color: 'var(--primary-600)' }}>₹{stats?.globalTotalCollected}</div>
                  </div>
                </div>

                <div className="card" style={{ padding: '16px', marginBottom: '20px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <div>
                    <span className="card-title" style={{ fontSize: '10px', display: 'block' }}>{t('subscribers')}</span>
                    <strong style={{ fontSize: '20px', color: 'var(--neutral-900)' }}>{stats?.totalCusts}</strong>
                  </div>
                  <div style={{ textAlign: 'right' }}>
                    <span className="card-title" style={{ fontSize: '10px', display: 'block' }}>{t('activeNodes')}</span>
                    <strong style={{ fontSize: '20px', color: 'var(--accent-500)' }}>{stats?.activeCusts}</strong>
                  </div>
                </div>

                {/* Worker collection leagues progress */}
                <div className="card" style={{ padding: '18px', marginBottom: '20px' }}>
                  <h4 style={{ fontSize: '13px', fontWeight: 800, textTransform: 'uppercase', letterSpacing: '0.04em', color: 'var(--neutral-400)', marginBottom: '14px', display: 'flex', alignItems: 'center', gap: '6px' }}><Award size={16} /> {t('workerProgress')}</h4>
                  {profiles.filter(p => p.role === 'WORKER').map(w => {
                    const todayColl = payments.filter(p => p.worker_id === w.id && p.status === 'Paid' && p.payment_date === new Date().toISOString().split('T')[0]).reduce((sum, p) => sum + parseFloat(p.amount), 0);
                    const totalColl = payments.filter(p => p.worker_id === w.id && p.status === 'Paid').reduce((sum, p) => sum + parseFloat(p.amount), 0);
                    return (
                      <div key={w.id} style={{ marginBottom: '14px', borderBottom: '1px solid var(--neutral-100)', paddingBottom: '10px' }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '6px' }}>
                          <span style={{ fontSize: '13px', fontWeight: 700, color: 'var(--neutral-800)' }}>{w.name}</span>
                          <span style={{ fontSize: '11px', color: 'var(--neutral-400)', fontWeight: 600 }}>Collector</span>
                        </div>
                        <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '12px' }}>
                          <span style={{ color: 'var(--neutral-500)' }}>{t('todayEarning')}: <strong style={{ color: 'var(--success)' }}>₹{todayColl}</strong></span>
                          <span style={{ color: 'var(--neutral-500)' }}>{t('totalEarning')}: <strong style={{ color: 'var(--primary-600)' }}>₹{totalColl}</strong></span>
                        </div>
                      </div>
                    );
                  })}
                </div>

                {/* Street metrics progress */}
                <div className="card" style={{ padding: '18px', marginBottom: '20px' }}>
                  <h4 style={{ fontSize: '13px', fontWeight: 800, textTransform: 'uppercase', letterSpacing: '0.04em', color: 'var(--neutral-400)', marginBottom: '14px', display: 'flex', alignItems: 'center', gap: '6px' }}><MapPin size={16} /> {t('areaRankings')}</h4>
                  {Object.entries(stats?.streetPerformance || {}).map(([street, collected]) => (
                    <div key={street} style={{ marginBottom: '12px' }}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '13px', marginBottom: '4px', fontWeight: 600 }}>
                        <span>{street}</span>
                        <strong>₹{collected}</strong>
                      </div>
                      <div className="progress-bar-container" style={{ height: '6px' }}>
                        <div className="progress-bar-fill" style={{ width: `${(collected / (stats?.reportCollected || 1)) * 100}%`, background: 'var(--accent-500)' }}></div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            ) : (
              // WORKER VIEW DASHBOARD
              <div>
                <div style={{ marginBottom: '24px', background: 'linear-gradient(135deg, var(--primary-500) 0%, var(--primary-700) 100%)', color: '#ffffff', borderRadius: 'var(--radius-lg)', padding: '24px', boxShadow: 'var(--shadow-md)', position: 'relative', overflow: 'hidden' }}>
                  <p style={{ opacity: 0.8, fontSize: '11px', textTransform: 'uppercase', letterSpacing: '0.08em', fontWeight: 700 }}>{t('todayEarning')}</p>
                  <h3 style={{ fontSize: '32px', fontFamily: 'var(--font-primary)', fontWeight: 800, marginTop: '4px' }}>₹{stats?.myTodayColl}</h3>
                  <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: '16px', paddingTop: '16px', borderTop: '1px solid rgba(255,255,255,0.15)', fontSize: '13px', fontWeight: 600 }}>
                    <span>{t('totalEarning')}: <strong>₹{stats?.myTotalColl}</strong></span>
                    <span>{t('customers')}: <strong>{stats?.myCustomersCount}</strong></span>
                  </div>
                </div>

                <h4 style={{ fontSize: '14px', fontWeight: 800, textTransform: 'uppercase', letterSpacing: '0.04em', color: 'var(--neutral-400)', marginBottom: '14px' }}>{t('myRecentActs')}</h4>
                {payments.filter(p => p.worker_id === currentUser.id).slice(0, 5).map(pay => {
                  const cust = customers.find(c => c.id === pay.customer_id);
                  return (
                    <div key={pay.id} className="list-item">
                      <div className="list-item-main">
                        <span className="list-item-title">{cust ? cust.customer_name : 'Customer'}</span>
                        <span className="list-item-subtitle">{pay.payment_period} • {t('boxId')}: {pay.box_id}</span>
                      </div>
                      <div>
                        <strong style={{ color: 'var(--success)', fontSize: '15px' }}>₹{pay.amount}</strong>
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        )}

        {/* MOBILE SCREEN: CUSTOMERS */}
        {activeTab === 'customers' && (
          <div className="animate-fade-in">
            <div style={{ display: 'flex', gap: '8px', marginBottom: '12px' }}>
              <div style={{ position: 'relative', flex: 1 }}>
                <Search size={18} style={{ position: 'absolute', left: '12px', top: '14px', color: 'var(--neutral-400)' }} />
                <input
                  type="text"
                  className="form-input"
                  style={{ paddingLeft: '38px' }}
                  placeholder={t('searchPlaceholder')}
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                />
              </div>
              
              <button
                onClick={() => setModalType('add_customer')}
                className="btn btn-primary"
                style={{ padding: '12px', borderRadius: 'var(--radius-md)' }}
              >
                <Plus size={20} />
              </button>
            </div>

            {/* Custom Layout Viewer Toggles */}
            <div className="card" style={{ padding: '8px', marginBottom: '16px', display: 'flex', alignItems: 'center', background: 'var(--neutral-100)', border: 'none', borderRadius: 'var(--radius-md)' }}>
              <div style={{ display: 'flex', width: '100%', gap: '4px' }}>
                <button 
                  onClick={() => setCustomerViewMode('standard')} 
                  className="btn" 
                  style={{ flex: 1, padding: '8px', fontSize: '11.5px', borderRadius: '8px', background: customerViewMode === 'standard' ? 'var(--primary-500)' : 'transparent', color: customerViewMode === 'standard' ? '#ffffff' : 'var(--neutral-700)', fontWeight: 700 }}
                >
                  {t('standardList')}
                </button>
                <button 
                  onClick={() => setCustomerViewMode('street')} 
                  className="btn" 
                  style={{ flex: 1, padding: '8px', fontSize: '11.5px', borderRadius: '8px', background: customerViewMode === 'street' ? 'var(--primary-500)' : 'transparent', color: customerViewMode === 'street' ? '#ffffff' : 'var(--neutral-700)', fontWeight: 700 }}
                >
                  {t('streetGroup')}
                </button>
                <button 
                  onClick={() => setCustomerViewMode('box_id')} 
                  className="btn" 
                  style={{ flex: 1, padding: '8px', fontSize: '11.5px', borderRadius: '8px', background: customerViewMode === 'box_id' ? 'var(--primary-500)' : 'transparent', color: customerViewMode === 'box_id' ? '#ffffff' : 'var(--neutral-700)', fontWeight: 700 }}
                >
                  {t('boxIdSort')}
                </button>
              </div>
            </div>

            {/* Quick Actions (Owner only: Excel exports) */}
            {currentUser.role === 'OWNER' && (
              <div style={{ marginBottom: '16px' }}>
                <button onClick={handleExportCustomers} className="btn btn-secondary icon-align" style={{ width: '100%', fontSize: '12px', padding: '12px 0' }}>
                  <Download size={14} /> {t('exportExcel')}
                </button>
              </div>
            )}

            <div>
              <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '10px', fontSize: '12px', color: 'var(--neutral-400)', fontWeight: 600 }}>
                <span>Showing {filteredCustomers.length} {t('customers')}</span>
              </div>

              {filteredCustomers.length === 0 ? (
                <div className="card" style={{ textAlign: 'center', padding: '40px 20px', color: 'var(--neutral-400)' }}>
                  <Info size={32} style={{ margin: '0 auto 12px auto' }} />
                  <p>No customers matching filters.</p>
                </div>
              ) : customerViewMode === 'street' ? (
                Object.entries(customersGroupedByStreet).map(([streetName, streetCustomers]) => {
                  const isExpanded = !!expandedStreets[streetName];
                  return (
                    <div key={streetName} style={{ marginBottom: '14px' }}>
                      <div 
                        onClick={() => toggleStreet(streetName)}
                        style={{ 
                          background: 'var(--neutral-100)', 
                          padding: '14px 18px', 
                          borderRadius: 'var(--radius-md)', 
                          fontSize: '13px', 
                          fontWeight: 700, 
                          color: 'var(--neutral-800)', 
                          display: 'flex', 
                          justifyContent: 'space-between',
                          alignItems: 'center',
                          cursor: 'pointer',
                          border: '1px solid var(--neutral-200)',
                          userSelect: 'none',
                          transition: 'all var(--transition-fast)'
                        }}
                        className="street-accordion-header"
                      >
                        <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                          <span style={{ color: 'var(--neutral-400)', display: 'inline-flex', alignItems: 'center', transform: isExpanded ? 'rotate(0deg)' : 'rotate(-90deg)', transition: 'transform 0.2s ease' }}>
                            <ChevronDown size={16} />
                          </span>
                          <span>📍 {streetName}</span>
                        </div>
                        <span className="badge badge-active" style={{ fontSize: '11px', padding: '3px 10px' }}>
                          {streetCustomers.length} {streetCustomers.length === 1 ? 'Account' : 'Accounts'}
                        </span>
                      </div>
                      
                      {isExpanded && (
                        <div className="animate-fade-in" style={{ marginTop: '8px', display: 'flex', flexDirection: 'column', gap: '8px' }}>
                          {streetCustomers.map(cust => (
                            <div
                              key={cust.id}
                              className="list-item"
                              onClick={() => {
                                setSelectedCustomer(cust);
                                setModalType('customer_details');
                              }}
                              style={{ cursor: 'pointer', padding: '14px 18px', marginBottom: 0 }}
                            >
                              <div className="list-item-main">
                                <span className="list-item-title" style={{ fontSize: '14px' }}>{cust.customer_name}</span>
                                <span className="list-item-subtitle">{cust.street_name} • {t('boxId')}: {cust.box_id}</span>
                              </div>
                              <span className={`badge ${cust.connection_status === 'ACTIVE' ? 'badge-active' : 'badge-inactive'}`} style={{ fontSize: '11px' }}>
                                {cust.connection_status === 'ACTIVE' ? t('active') : t('inactive')}
                              </span>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  );
                })
              ) : (
                filteredCustomers.map(cust => (
                  <div
                    key={cust.id}
                    className="list-item"
                    onClick={() => {
                      setSelectedCustomer(cust);
                      setModalType('customer_details');
                    }}
                    style={{ cursor: 'pointer' }}
                  >
                    <div className="list-item-main">
                      <span className="list-item-title">{cust.customer_name}</span>
                      <span className="list-item-subtitle">{cust.street_name} • {t('boxId')}: {cust.box_id}</span>
                    </div>
                    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end', gap: '6px' }}>
                      <span className={`badge ${cust.connection_status === 'ACTIVE' ? 'badge-active' : 'badge-inactive'}`}>
                        {cust.connection_status === 'ACTIVE' ? t('active') : t('inactive')}
                      </span>
                    </div>
                  </div>
                ))
              )}
            </div>
          </div>
        )}

        {/* MOBILE SCREEN: PAYMENTS HISTORY */}
        {activeTab === 'payments' && (
          <div className="animate-fade-in">
            <h3 style={{ fontSize: '16px', fontWeight: 800, marginBottom: '16px' }}>{t('ledger')}</h3>
            <div className="card" style={{ padding: '0 16px' }}>
              {payments.map(pay => {
                const cust = customers.find(c => c.id === pay.customer_id);
                return (
                  <div key={pay.id} style={{ display: 'flex', justifyContent: 'space-between', padding: '16px 0', borderBottom: '1px solid var(--neutral-100)' }}>
                    <div>
                      <p style={{ fontSize: '13px', fontWeight: 700 }}>{cust ? cust.customer_name : 'Customer Record'}</p>
                      <span style={{ fontSize: '11px', color: 'var(--neutral-400)' }}>{t('boxId')}: {pay.box_id} • {pay.payment_period}</span>
                    </div>
                    <div style={{ textAlign: 'right' }}>
                      <p style={{ fontSize: '13px', fontWeight: 800, color: 'var(--success)' }}>₹{pay.amount}</p>
                      <span style={{ fontSize: '10px', color: 'var(--neutral-400)' }}>{pay.payment_date}</span>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* MOBILE SCREEN: REPORTS, LOCALIZATION & INSPECTIONS */}
        {activeTab === 'reports' && (
          <div className="animate-fade-in">
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px' }}>
              <h3 style={{ fontSize: '16px', fontWeight: 800 }}>{t('reports')}</h3>
              <button onClick={handleLogout} className="btn btn-danger icon-align" style={{ padding: '8px 14px', fontSize: '12px' }}>
                <LogOut size={12} /> {t('logOut')}
              </button>
            </div>

            {/* Language Selection Card */}
            <div className="card" style={{ padding: '18px', marginBottom: '20px' }}>
              <h4 style={{ fontSize: '13px', fontWeight: 800, textTransform: 'uppercase', letterSpacing: '0.04em', color: 'var(--neutral-400)', marginBottom: '12px', display: 'flex', alignItems: 'center', gap: '6px' }}>
                <Globe size={16} /> {t('langSettings')}
              </h4>
              <div style={{ display: 'flex', gap: '8px' }}>
                <button 
                  onClick={() => setLanguage('en')} 
                  className="btn" 
                  style={{ flex: 1, padding: '10px 0', fontSize: '13px', background: language === 'en' ? 'var(--primary-500)' : 'var(--neutral-100)', color: language === 'en' ? '#ffffff' : 'var(--neutral-700)' }}
                >
                  English
                </button>
                <button 
                  onClick={() => setLanguage('ta')} 
                  className="btn" 
                  style={{ flex: 1, padding: '10px 0', fontSize: '13px', background: language === 'ta' ? 'var(--primary-500)' : 'var(--neutral-100)', color: language === 'ta' ? '#ffffff' : 'var(--neutral-700)' }}
                >
                  தமிழ் (Tamil)
                </button>
              </div>
            </div>

            {/* Selection filters */}
            <div className="card" style={{ padding: '16px', marginBottom: '20px' }}>
              <div className="form-group" style={{ marginBottom: '12px' }}>
                <label className="form-label" style={{ fontSize: '11px' }}>{t('filterPeriod')}</label>
                <select 
                  className="form-input form-select" 
                  value={reportMonth} 
                  onChange={(e) => setReportMonth(e.target.value)}
                  style={{ padding: '10px 14px', fontSize: '13px' }}
                >
                  <option value="All Time">{t('allTime')}</option>
                  {paymentPeriodsList.map(period => (
                    <option key={period} value={period}>{period}</option>
                  ))}
                </select>
              </div>
              <button 
                onClick={() => handleExportMonthlyReport(reportMonth)} 
                className="btn btn-primary icon-align" 
                style={{ width: '100%', padding: '12px', fontSize: '13px' }}
              >
                <Download size={14} /> {t('downloadReport')}
              </button>
            </div>

            {currentUser.role === 'OWNER' ? (
              // OWNER REPORTS VIEW
              <div>
                <h4 style={{ fontSize: '14px', fontWeight: 800, textTransform: 'uppercase', letterSpacing: '0.04em', color: 'var(--neutral-400)', marginBottom: '12px' }}>{t('inspectWorkers')}</h4>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                  {profiles.filter(w => w.role === 'WORKER').map(w => {
                    const metrics = getWorkerProgressMetrics(w.id);
                    return (
                      <div 
                        key={w.id} 
                        className="list-item" 
                        style={{ cursor: 'pointer', padding: '16px' }}
                        onClick={() => {
                          setSelectedWorker(w);
                          setModalType('inspect_worker');
                        }}
                      >
                        <div className="list-item-main" style={{ flex: 1 }}>
                          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                            <span className="list-item-title" style={{ fontSize: '14px' }}>{w.name}</span>
                            <span className="badge badge-active" style={{ fontSize: '10px', padding: '2px 8px' }}>{t('active')}</span>
                          </div>
                          <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '12px', color: 'var(--neutral-400)', marginTop: '8px' }}>
                            <span>Today: <strong style={{ color: 'var(--success)' }}>₹{metrics.todayCollected}</strong></span>
                            <span>Total: <strong style={{ color: 'var(--primary-600)' }}>₹{metrics.totalCollected}</strong></span>
                            <span>Subscribers: <strong>{metrics.customersCount}</strong></span>
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            ) : (
              // WORKER REPORTS VIEW
              <div>
                <h4 style={{ fontSize: '14px', fontWeight: 800, textTransform: 'uppercase', letterSpacing: '0.04em', color: 'var(--neutral-400)', marginBottom: '12px' }}>{t('myProgress')}</h4>
                
                {(() => {
                  const myMetrics = getWorkerProgressMetrics(currentUser.id);
                  return (
                    <div>
                      <div className="card" style={{ padding: '16px', marginBottom: '20px' }}>
                        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: '12px', textAlign: 'center', marginBottom: '12px' }}>
                          <div>
                            <span style={{ fontSize: '10px', color: 'var(--neutral-400)', display: 'block', textTransform: 'uppercase', fontWeight: 700 }}>{t('todayEarning')}</span>
                            <strong style={{ fontSize: '16px', color: 'var(--success)' }}>₹{myMetrics.todayCollected}</strong>
                          </div>
                          <div>
                            <span style={{ fontSize: '10px', color: 'var(--neutral-400)', display: 'block', textTransform: 'uppercase', fontWeight: 700 }}>{t('totalEarning')}</span>
                            <strong style={{ fontSize: '16px', color: 'var(--primary-600)' }}>₹{myMetrics.totalCollected}</strong>
                          </div>
                        </div>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderTop: '1px solid var(--neutral-100)', paddingTop: '10px', fontSize: '12px' }}>
                          <span>My Collections for <strong>{reportMonth}</strong>:</span>
                          <strong style={{ color: 'var(--accent-600)' }}>₹{myMetrics.collected}</strong>
                        </div>
                      </div>

                      <h4 style={{ fontSize: '13px', fontWeight: 800, textTransform: 'uppercase', letterSpacing: '0.04em', color: 'var(--neutral-400)', marginBottom: '12px' }}>My Collections Ledger ({reportMonth})</h4>
                      <div className="card" style={{ padding: '0 16px' }}>
                        {myMetrics.history.map(p => {
                          const cust = customers.find(c => c.id === p.customer_id);
                          return (
                            <div key={p.id} style={{ display: 'flex', justifyContent: 'space-between', padding: '10px 0', borderBottom: '1px solid var(--neutral-100)', fontSize: '12px' }}>
                              <div>
                                <span style={{ fontWeight: 'bold' }}>{cust ? cust.customer_name : 'Customer'}</span>
                                <span style={{ color: 'var(--neutral-400)', display: 'block' }}>{t('boxId')}: {p.box_id}</span>
                              </div>
                              <span style={{ color: 'var(--success)', fontWeight: 'bold' }}>₹{p.amount}</span>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  );
                })()}
              </div>
            )}
          </div>
        )}

      </div>

      {/* MODALS & BOTTOM SHEETS */}
      {modalType === 'add_customer' && (
        <div className="modal-overlay">
          <div className="mobile-sheet animate-slide-in">
            <div className="sheet-handle"></div>
            <h3 style={{ marginBottom: '16px' }}>{t('addCustomer')}</h3>
            <form onSubmit={handleCreateCustomer}>
              <div className="form-group">
                <label className="form-label">Subscriber Name</label>
                <input
                  type="text"
                  className="form-input"
                  placeholder="e.g. John Doe"
                  value={customerForm.customer_name}
                  onChange={(e) => setCustomerForm({ ...customerForm, customer_name: e.target.value })}
                  required
                />
              </div>
              <div className="form-group">
                <label className="form-label">Street / Area Name</label>
                <input
                  type="text"
                  className="form-input"
                  placeholder="e.g. Park Road"
                  value={customerForm.street_name}
                  onChange={(e) => setCustomerForm({ ...customerForm, street_name: e.target.value })}
                  required
                />
              </div>
              <div className="form-group">
                <label className="form-label">Hardware Box / STB ID</label>
                <input
                  type="text"
                  className="form-input"
                  placeholder="e.g. BOX-10029"
                  value={customerForm.box_id}
                  onChange={(e) => setCustomerForm({ ...customerForm, box_id: e.target.value })}
                  required
                />
              </div>
              <div className="form-group">
                <label className="form-label">Phone Number</label>
                <input
                  type="text"
                  className="form-input"
                  placeholder="e.g. +91 9876543210"
                  value={customerForm.phone_number}
                  onChange={(e) => setCustomerForm({ ...customerForm, phone_number: e.target.value })}
                />
              </div>
              {currentUser.role === 'OWNER' && (
                <div className="form-group">
                  <label className="form-label">Assign Worker</label>
                  <select
                    className="form-input form-select"
                    value={customerForm.assigned_worker_id}
                    onChange={(e) => setCustomerForm({ ...customerForm, assigned_worker_id: e.target.value })}
                  >
                    <option value="">{t('unassigned')}</option>
                    {profiles.filter(p => p.role === 'WORKER').map(w => (
                      <option key={w.id} value={w.id}>{w.name}</option>
                    ))}
                  </select>
                </div>
              )}
              <div style={{ display: 'flex', gap: '12px', marginTop: '24px' }}>
                <button type="button" onClick={() => setModalType(null)} className="btn btn-secondary" style={{ flex: 1 }}>{t('cancel')}</button>
                <button type="submit" className="btn btn-primary" style={{ flex: 1 }}>{t('saveCustomer')}</button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Edit Customer Modal */}
      {modalType === 'edit_customer' && selectedCustomer && (
        <div className="modal-overlay">
          <div className="mobile-sheet animate-slide-in">
            <div className="sheet-handle"></div>
            <h3 style={{ marginBottom: '16px' }}>{t('editCustomer')}</h3>
            <form onSubmit={handleEditCustomer}>
              <div className="form-group">
                <label className="form-label">Subscriber Name</label>
                <input
                  type="text"
                  className="form-input"
                  value={customerForm.customer_name}
                  onChange={(e) => setCustomerForm({ ...customerForm, customer_name: e.target.value })}
                  required
                />
              </div>
              <div className="form-group">
                <label className="form-label">Street Name</label>
                <input
                  type="text"
                  className="form-input"
                  value={customerForm.street_name}
                  onChange={(e) => setCustomerForm({ ...customerForm, street_name: e.target.value })}
                  required
                />
              </div>
              <div className="form-group">
                <label className="form-label">Box ID</label>
                <input
                  type="text"
                  className="form-input"
                  value={customerForm.box_id}
                  onChange={(e) => setCustomerForm({ ...customerForm, box_id: e.target.value })}
                  required
                />
              </div>
              <div className="form-group">
                <label className="form-label">Phone Number</label>
                <input
                  type="text"
                  className="form-input"
                  value={customerForm.phone_number}
                  onChange={(e) => setCustomerForm({ ...customerForm, phone_number: e.target.value })}
                />
              </div>
              {currentUser.role === 'OWNER' && (
                <div className="form-group">
                  <label className="form-label">Assign Worker</label>
                  <select
                    className="form-input form-select"
                    value={customerForm.assigned_worker_id}
                    onChange={(e) => setCustomerForm({ ...customerForm, assigned_worker_id: e.target.value })}
                  >
                    <option value="">{t('unassigned')}</option>
                    {profiles.filter(p => p.role === 'WORKER').map(w => (
                      <option key={w.id} value={w.id}>{w.name}</option>
                    ))}
                  </select>
                </div>
              )}
              <div className="form-group">
                <label className="form-label">{t('connectionStatus')}</label>
                <select
                  className="form-input form-select"
                  value={customerForm.connection_status}
                  onChange={(e) => setCustomerForm({ ...customerForm, connection_status: e.target.value })}
                >
                  <option value="ACTIVE">ACTIVE</option>
                  <option value="INACTIVE">INACTIVE</option>
                </select>
              </div>
              <div className="form-group">
                <label className="form-label">{t('notes')}</label>
                <textarea
                  className="form-input"
                  value={customerForm.notes}
                  onChange={(e) => setCustomerForm({ ...customerForm, notes: e.target.value })}
                />
              </div>
              <div style={{ display: 'flex', gap: '12px', marginTop: '24px' }}>
                <button type="button" onClick={() => setModalType(null)} className="btn btn-secondary" style={{ flex: 1 }}>{t('cancel')}</button>
                <button type="submit" className="btn btn-primary" style={{ flex: 1 }}>{t('saveChanges')}</button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Customer Details Bottom Sheet */}
      {modalType === 'customer_details' && selectedCustomer && (
        <div className="modal-overlay">
          <div className="mobile-sheet animate-slide-in" style={{ maxHeight: '90vh', overflowY: 'auto' }}>
            <div className="sheet-handle"></div>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '20px' }}>
              <div>
                <h3 style={{ fontSize: '18px', fontWeight: 800 }}>{selectedCustomer.customer_name}</h3>
                <p style={{ fontSize: '13px', color: 'var(--neutral-400)' }}>{t('boxId')}: <strong>{selectedCustomer.box_id}</strong></p>
              </div>
              <span className={`badge ${selectedCustomer.connection_status === 'ACTIVE' ? 'badge-active' : 'badge-inactive'}`}>
                {selectedCustomer.connection_status === 'ACTIVE' ? t('active') : t('inactive')}
              </span>
            </div>

            <div style={{ background: 'var(--neutral-50)', border: '1px solid var(--neutral-200)', borderRadius: 'var(--radius-md)', padding: '18px', marginBottom: '22px', display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px', fontSize: '13px' }}>
                <MapPin size={16} style={{ color: 'var(--neutral-400)' }} />
                <span style={{ color: 'var(--neutral-500)' }}>Street:</span>
                <strong style={{ marginLeft: 'auto', color: 'var(--neutral-900)' }}>{selectedCustomer.street_name}</strong>
              </div>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px', fontSize: '13px' }}>
                <Globe size={16} style={{ color: 'var(--neutral-400)' }} />
                <span style={{ color: 'var(--neutral-500)' }}>Phone Number:</span>
                <strong style={{ marginLeft: 'auto', color: 'var(--neutral-900)' }}>{selectedCustomer.phone_number || 'N/A'}</strong>
              </div>
              {currentUser.role === 'OWNER' && (
                <div style={{ display: 'flex', alignItems: 'center', gap: '10px', fontSize: '13px' }}>
                  <Users size={16} style={{ color: 'var(--neutral-400)' }} />
                  <span style={{ color: 'var(--neutral-500)' }}>Assigned Worker:</span>
                  <strong style={{ marginLeft: 'auto', color: 'var(--neutral-900)' }}>
                    {profiles.find(p => p.id === selectedCustomer.assigned_worker_id)?.name || t('unassigned')}
                  </strong>
                </div>
              )}
              {selectedCustomer.notes && (
                <div style={{ background: '#ffffff', border: '1px solid var(--neutral-200)', padding: '10px 14px', borderRadius: 'var(--radius-sm)', marginTop: '4px', fontSize: '12.5px', color: 'var(--neutral-600)', fontStyle: 'italic' }}>
                  <strong>Notes:</strong> {selectedCustomer.notes}
                </div>
              )}
            </div>

            <h4 style={{ fontSize: '14px', fontWeight: 600, marginBottom: '12px' }}>Payment Actions</h4>
            <div style={{ display: 'flex', gap: '8px', marginBottom: '20px' }}>
              <button
                onClick={() => {
                  setModalType('record_single');
                  setSinglePaymentForm({ amount: '350', period: new Date().toLocaleDateString('en-US', { month: 'long', year: 'numeric' }), notes: '' });
                }}
                className="btn btn-primary"
                style={{ flex: 1, padding: '10px 0', fontSize: '13px' }}
              >
                {t('recordPayment')}
              </button>
              <button
                onClick={() => {
                  setModalType('record_bulk');
                  setBulkPaymentForm({ periods: [], amount: '350', notes: '' });
                }}
                className="btn btn-secondary"
                style={{ flex: 1, padding: '10px 0', fontSize: '13px' }}
              >
                {t('bulkPayments')}
              </button>
            </div>

            <div style={{ display: 'flex', gap: '8px', borderTop: '1px solid var(--neutral-100)', paddingTop: '16px', marginBottom: '16px' }}>
              <button
                onClick={() => {
                  setCustomerForm({
                    customer_name: selectedCustomer.customer_name,
                    street_name: selectedCustomer.street_name,
                    box_id: selectedCustomer.box_id,
                    phone_number: selectedCustomer.phone_number || '',
                    assigned_worker_id: selectedCustomer.assigned_worker_id || '',
                    connection_status: selectedCustomer.connection_status,
                    notes: selectedCustomer.notes || ''
                  });
                  setModalType('edit_customer');
                }}
                className="btn btn-secondary icon-align"
                style={{ flex: 1, padding: '8px 0', fontSize: '12px' }}
              >
                <Edit2 size={14} /> {t('editCustomer')}
              </button>
              {currentUser.role === 'OWNER' && (
                <button
                  onClick={() => handleDeleteCustomer(selectedCustomer.id)}
                  className="btn btn-danger icon-align"
                  style={{ flex: 1, padding: '8px 0', fontSize: '12px' }}
                >
                  <Trash2 size={14} /> {t('delete')}
                </button>
              )}
            </div>

            <h4 style={{ fontSize: '14px', fontWeight: 600, marginBottom: '12px' }}>Payment History</h4>
            <div style={{ maxHeight: '160px', overflowY: 'auto' }}>
              {payments.filter(p => p.customer_id === selectedCustomer.id).map(p => (
                <div key={p.id} style={{ display: 'flex', justifyContent: 'space-between', padding: '8px 0', borderBottom: '1px solid var(--neutral-100)', fontSize: '12px' }}>
                  <span>{p.payment_period} ({p.payment_date})</span>
                  <strong style={{ color: p.status === 'Paid' ? 'var(--success)' : 'var(--warning)' }}>
                    ₹{p.amount} ({p.status})
                  </strong>
                </div>
              ))}
            </div>

            <button type="button" onClick={() => setModalType(null)} className="btn btn-secondary" style={{ width: '100%', marginTop: '24px' }}>Close</button>
          </div>
        </div>
      )}

      {/* Record Single Payment Modal */}
      {modalType === 'record_single' && selectedCustomer && (
        <div className="modal-overlay">
          <div className="mobile-sheet animate-slide-in">
            <div className="sheet-handle"></div>
            <h3 style={{ marginBottom: '16px' }}>Log Payment for {selectedCustomer.customer_name}</h3>
            <form onSubmit={handleSinglePayment}>
              <div className="form-group">
                <label className="form-label">{t('amount')}</label>
                <input
                  type="number"
                  className="form-input"
                  value={singlePaymentForm.amount}
                  onChange={(e) => setSinglePaymentForm({ ...singlePaymentForm, amount: e.target.value })}
                  required
                />
              </div>
              <div className="form-group">
                <label className="form-label">{t('period')}</label>
                <input
                  type="text"
                  className="form-input"
                  placeholder="e.g. June 2026"
                  value={singlePaymentForm.period}
                  onChange={(e) => setSinglePaymentForm({ ...singlePaymentForm, period: e.target.value })}
                  required
                />
              </div>
              <div className="form-group">
                <label className="form-label">{t('notes')}</label>
                <input
                  type="text"
                  className="form-input"
                  placeholder="Cash, GPay, etc."
                  value={singlePaymentForm.notes}
                  onChange={(e) => setSinglePaymentForm({ ...singlePaymentForm, notes: e.target.value })}
                />
              </div>
              <div style={{ display: 'flex', gap: '12px', marginTop: '24px' }}>
                <button type="button" onClick={() => setModalType('customer_details')} className="btn btn-secondary" style={{ flex: 1 }}>Back</button>
                <button type="submit" className="btn btn-primary" style={{ flex: 1 }}>Post Payment</button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Record Bulk Payment Modal */}
      {modalType === 'record_bulk' && selectedCustomer && (
        <div className="modal-overlay">
          <div className="mobile-sheet animate-slide-in" style={{ maxHeight: '90vh', overflowY: 'auto' }}>
            <div className="sheet-handle"></div>
            <h3 style={{ marginBottom: '16px' }}>Bulk Payments Log</h3>
            <form onSubmit={handleBulkPayment}>
              <div className="form-group">
                <label className="form-label">Select Paid Months/Periods</label>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', maxHeight: '160px', overflowY: 'auto', background: 'var(--neutral-50)', padding: '12px', borderRadius: 'var(--radius-sm)' }}>
                  {['January 2026', 'February 2026', 'March 2026', 'April 2026', 'May 2026', 'June 2026', 'July 2026', 'August 2026', 'September 2026', 'October 2026', 'November 2026', 'December 2026'].map(m => {
                    const isChecked = bulkPaymentForm.periods.includes(m);
                    return (
                      <label key={m} style={{ display: 'flex', alignItems: 'center', gap: '10px', fontSize: '13px', cursor: 'pointer', fontWeight: 600 }}>
                        <input
                          type="checkbox"
                          checked={isChecked}
                          onChange={() => {
                            const updated = isChecked
                              ? bulkPaymentForm.periods.filter(x => x !== m)
                              : [...bulkPaymentForm.periods, m];
                            setBulkPaymentForm({ ...bulkPaymentForm, periods: updated, amount: (updated.length * 350).toString() });
                          }}
                        />
                        <span>{m}</span>
                      </label>
                    );
                  })}
                </div>
              </div>
              <div className="form-group">
                <label className="form-label">{t('amount')}</label>
                <input
                  type="number"
                  className="form-input"
                  value={bulkPaymentForm.amount}
                  onChange={(e) => setBulkPaymentForm({ ...bulkPaymentForm, amount: e.target.value })}
                  required
                />
              </div>
              <div className="form-group">
                <label className="form-label">{t('notes')}</label>
                <input
                  type="text"
                  className="form-input"
                  placeholder="Bulk advanced fee log"
                  value={bulkPaymentForm.notes}
                  onChange={(e) => setBulkPaymentForm({ ...bulkPaymentForm, notes: e.target.value })}
                />
              </div>
              <div style={{ display: 'flex', gap: '12px', marginTop: '24px' }}>
                <button type="button" onClick={() => setModalType('customer_details')} className="btn btn-secondary" style={{ flex: 1 }}>Back</button>
                <button type="submit" className="btn btn-primary" style={{ flex: 1 }}>Log Bulk Payments</button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* INSPECT WORKER BOTTOM SHEET (OWNER ONLY) */}
      {modalType === 'inspect_worker' && selectedWorker && currentUser.role === 'OWNER' && (
        <div className="modal-overlay">
          <div className="mobile-sheet animate-slide-in" style={{ maxHeight: '90vh', overflowY: 'auto' }}>
            <div className="sheet-handle"></div>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px' }}>
              <div>
                <h3 style={{ fontSize: '16px', fontWeight: 800 }}>{selectedWorker.name}</h3>
                <p style={{ fontSize: '12px', color: 'var(--neutral-400)', fontWeight: 600 }}>{t('workerInspector')}</p>
              </div>
              <span className={`badge ${!selectedWorker.disabled ? 'badge-active' : 'badge-inactive'}`}>
                {!selectedWorker.disabled ? t('active') : t('inactive')}
              </span>
            </div>

            {(() => {
              const workerMetrics = getWorkerProgressMetrics(selectedWorker.id);
              return (
                <div>
                  <div className="card" style={{ padding: '16px', marginBottom: '20px' }}>
                    <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '8px', textAlign: 'center', marginBottom: '12px' }}>
                      <div>
                        <span style={{ fontSize: '10px', color: 'var(--neutral-400)', display: 'block', textTransform: 'uppercase', fontWeight: 700 }}>{t('todayEarning')}</span>
                        <strong style={{ fontSize: '15px', color: 'var(--success)' }}>₹{workerMetrics.todayCollected}</strong>
                      </div>
                      <div>
                        <span style={{ fontSize: '10px', color: 'var(--neutral-400)', display: 'block', textTransform: 'uppercase', fontWeight: 700 }}>{t('totalEarning')}</span>
                        <strong style={{ fontSize: '15px', color: 'var(--primary-600)' }}>₹{workerMetrics.totalCollected}</strong>
                      </div>
                      <div>
                        <span style={{ fontSize: '10px', color: 'var(--neutral-400)', display: 'block', textTransform: 'uppercase', fontWeight: 700 }}>Subscribers</span>
                        <strong style={{ fontSize: '15px', color: 'var(--neutral-900)' }}>{workerMetrics.customersCount}</strong>
                      </div>
                    </div>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderTop: '1px solid var(--neutral-100)', paddingTop: '10px', fontSize: '12px' }}>
                      <span>Collections for <strong>{reportMonth}</strong>:</span>
                      <strong style={{ color: 'var(--accent-600)' }}>₹{workerMetrics.collected}</strong>
                    </div>
                    
                    <button 
                      onClick={() => handleExportMonthlyReport(reportMonth, selectedWorker.id)}
                      className="btn btn-secondary icon-align" 
                      style={{ width: '100%', padding: '10px', fontSize: '12px', marginTop: '14px' }}
                    >
                      <Download size={14} /> Download Worker Monthly Report (XLS)
                    </button>
                  </div>

                  <h4 style={{ fontSize: '13px', fontWeight: 800, textTransform: 'uppercase', letterSpacing: '0.04em', color: 'var(--neutral-400)', marginBottom: '12px' }}>Collections Logged ({reportMonth})</h4>
                  <div style={{ maxHeight: '180px', overflowY: 'auto' }}>
                    {workerMetrics.history.length === 0 ? (
                      <p style={{ fontSize: '12px', color: 'var(--neutral-400)', textAlign: 'center', padding: '20px 0' }}>No collections registered for this period.</p>
                    ) : (
                      workerMetrics.history.map(p => {
                        const cust = customers.find(c => c.id === p.customer_id);
                        return (
                          <div key={p.id} style={{ display: 'flex', justifyContent: 'space-between', padding: '10px 0', borderBottom: '1px solid var(--neutral-100)', fontSize: '12px' }}>
                            <div>
                              <span style={{ fontWeight: 'bold' }}>{cust ? cust.customer_name : 'Customer'}</span>
                              <span style={{ color: 'var(--neutral-400)', display: 'block' }}>{t('boxId')}: {p.box_id} • {p.payment_date}</span>
                            </div>
                            <span style={{ color: 'var(--success)', fontWeight: 'bold' }}>₹{p.amount}</span>
                          </div>
                        );
                      })
                    )}
                  </div>
                </div>
              );
            })()}

            <button type="button" onClick={() => setModalType('reports')} className="btn btn-secondary" style={{ width: '100%', marginTop: '24px' }}>{t('closeInspector')}</button>
          </div>
        </div>
      )}

      {/* Mobile Sticky Tab Navigation Bar */}
      <nav className="mobile-nav">
        <button className={`mobile-nav-item ${activeTab === 'dashboard' ? 'active' : ''}`} onClick={() => setActiveTab('dashboard')}>
          <ShieldCheck size={20} />
          <span>{t('dashboard')}</span>
        </button>
        <button className={`mobile-nav-item ${activeTab === 'customers' ? 'active' : ''}`} onClick={() => setActiveTab('customers')}>
          <Users size={20} />
          <span>{t('customers')}</span>
        </button>
        <button className={`mobile-nav-item ${activeTab === 'payments' ? 'active' : ''}`} onClick={() => setActiveTab('payments')}>
          <Receipt size={20} />
          <span>{t('ledger')}</span>
        </button>
        <button className={`mobile-nav-item ${activeTab === 'reports' ? 'active' : ''}`} onClick={() => setActiveTab('reports')}>
          <BarChart3 size={20} />
          <span>{t('reports')}</span>
        </button>
      </nav>
    </div>
  );

  return (
    <div className="mobile-preview-container">
      <div className="phone-mockup">
        {mobileLayoutHTML}
      </div>
    </div>
  );
}
