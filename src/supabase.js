import { createClient } from '@supabase/supabase-js';

// Read Supabase environment variables from Vite
const supabaseUrl = import.meta.env?.VITE_SUPABASE_URL || '';
const supabaseAnonKey = import.meta.env?.VITE_SUPABASE_ANON_KEY || '';

// Determine if we are using the Mock Database
const isMock = !supabaseUrl || !supabaseAnonKey;

// Real Supabase Client
const realSupabase = isMock ? null : createClient(supabaseUrl, supabaseAnonKey);

// Helper: Map simple username to email formatting
export const mapUsernameToEmail = (username) => {
  if (username.includes('@')) return username;
  return `${username.trim().toLowerCase()}@kairos.com`;
};

// Helper: Generate UUID (with standard fallback)
export const generateUUID = () => {
  if (typeof window !== 'undefined' && window.crypto && window.crypto.randomUUID) {
    return window.crypto.randomUUID();
  }
  // Fallback UUID v4 generator
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
    var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
    return v.toString(16);
  });
};

// Helper: Hash password client-side using SHA-256 for secure database stashing
export const hashPassword = async (password, salt = 'kairos_cable_salt_2026') => {
  if (!password) return '';
  if (/^[a-f0-9]{64}$/i.test(password)) return password;
  
  const encoder = new TextEncoder();
  const data = encoder.encode(password + salt);
  
  // Try using crypto subtle (standard in modern browsers/environments)
  if (typeof window !== 'undefined' && window.crypto && window.crypto.subtle) {
    const hashBuffer = await window.crypto.subtle.digest('SHA-256', data);
    const hashArray = Array.from(new Uint8Array(hashBuffer));
    return hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
  }
  
  // Clean, warning-free JS hashing fallback
  let hash = 5381;
  for (let i = 0; i < password.length; i++) {
    hash = ((hash << 5) + hash) + password.charCodeAt(i);
  }
  const hashStr = Math.abs(hash).toString(16).padStart(16, 'f');
  return (hashStr + hashStr + hashStr + hashStr).substring(0, 64);
};

// -------------------------------------------------------------
// LOCAL STORAGE MOCK DATABASE SETUP
// -------------------------------------------------------------
const INITIAL_BUSINESSES = [];

const INITIAL_PROFILES = [
  { id: 'a1111111-1111-1111-1111-111111111111', business_id: null, username: 'admin_rhemi', role: 'SUPER_ADMIN', name: 'Rhemi (Super Admin)', phone_number: '+919999999999', disabled: false, created_at: new Date().toISOString() }
];

const INITIAL_CUSTOMERS = [];
const INITIAL_PAYMENTS = [];
const INITIAL_AUDITS = [];

const getLocalStorage = (key, fallback) => {
  const data = localStorage.getItem(key);
  if (!data) {
    localStorage.setItem(key, JSON.stringify(fallback));
    return fallback;
  }
  try {
    return JSON.parse(data);
  } catch {
    return fallback;
  }
};

const setLocalStorage = (key, data) => {
  localStorage.setItem(key, JSON.stringify(data));
};

// Initialize Mock Collections in Memory / Storage
const mockDB = {
  getBusinesses: () => getLocalStorage('kairos_businesses', INITIAL_BUSINESSES),
  setBusinesses: (data) => setLocalStorage('kairos_businesses', data),
  
  getProfiles: () => getLocalStorage('kairos_profiles', INITIAL_PROFILES),
  setProfiles: (data) => setLocalStorage('kairos_profiles', data),
  
  getCustomers: () => getLocalStorage('kairos_customers', INITIAL_CUSTOMERS),
  setCustomers: (data) => setLocalStorage('kairos_customers', data),
  
  getPayments: () => getLocalStorage('kairos_payments', INITIAL_PAYMENTS),
  setPayments: (data) => setLocalStorage('kairos_payments', data),
  
  getAudits: () => getLocalStorage('kairos_audits', INITIAL_AUDITS),
  setAudits: (data) => setLocalStorage('kairos_audits', data),
  
  getCurrentSession: () => {
    try {
      const data = localStorage.getItem('kairos_session');
      return data ? JSON.parse(data) : null;
    } catch {
      return null;
    }
  },
  setCurrentSession: (profile) => {
    if (profile) {
      localStorage.setItem('kairos_session', JSON.stringify(profile));
    } else {
      localStorage.removeItem('kairos_session');
    }
  }
};

// -------------------------------------------------------------
// UNIFIED API SERVICE (SUPABASE VS MOCK DUAL DRIVER)
// -------------------------------------------------------------
export const dbService = {
  isMock: () => isMock,

  // --- 1. AUTHENTICATION SERVICES ---
  auth: {
    signIn: async (username, password) => {
      const hashedPassword = await hashPassword(password);
      if (isMock) {
        const profiles = mockDB.getProfiles();
        const lookup = username.trim().toLowerCase().split('@')[0];
        const found = profiles.find(p => p.username === lookup);
        
        // Simple password checking
        if (!found) {
          throw new Error('User not found.');
        }
        if (found.disabled) {
          throw new Error('This user account has been disabled/blocked by Kairos Eido Technologies.');
        }
        
        const match = found.password === hashedPassword || 
                      found.password === password ||
                      (!found.password && (password === 'admin' || password === 'password123' || password === 'pasword123')); // fallback if empty
                      
        if (!match) {
          throw new Error('Incorrect password.');
        }

        // Auto-migrate to hashed password in mock db if not already hashed
        if (!/^[a-f0-9]{64}$/i.test(found.password || '')) {
          found.password = hashedPassword;
          mockDB.setProfiles(profiles);
        }

        // Mock login success
        mockDB.setCurrentSession(found);
        return { user: found, profile: found };
      } else {
        const cleanUsername = username.trim().toLowerCase().split('@')[0];
        const { data: profile, error } = await realSupabase
          .from('profiles')
          .select('*')
          .eq('username', cleanUsername)
          .maybeSingle();

        if (error) throw new Error('Database query error: ' + error.message);
        if (!profile) {
          throw new Error('Invalid username or password.');
        }

        // Check if the password matches the hash or the plaintext
        const isHashedMatch = profile.password === hashedPassword;
        const isPlaintextMatch = profile.password === password;

        if (!isHashedMatch && !isPlaintextMatch) {
          throw new Error('Invalid username or password.');
        }

        if (profile.disabled) {
          throw new Error('This account has been blocked. Contact Kairos Eido Technologies.');
        }

        // Auto-migrate plaintext password to hashed format in the database
        if (isPlaintextMatch && !isHashedMatch) {
          try {
            await realSupabase
              .from('profiles')
              .update({ password: hashedPassword })
              .eq('id', profile.id);
            profile.password = hashedPassword;
          } catch (migrateErr) {
            console.error('Failed to auto-migrate user password:', migrateErr);
          }
        }

        // Store session in local storage for database-level auth persistence
        localStorage.setItem('kairos_session', JSON.stringify(profile));
        return { user: profile, profile };
      }
    },

    signOut: async () => {
      if (isMock) {
        mockDB.setCurrentSession(null);
        return true;
      } else {
        localStorage.removeItem('kairos_session');
        return true;
      }
    },

    getCurrentSession: async () => {
      if (isMock) {
        return mockDB.getCurrentSession();
      } else {
        try {
          const sessionData = localStorage.getItem('kairos_session');
          if (!sessionData) return null;
          const parsed = JSON.parse(sessionData);

          // Re-verify status from database to handle blocks dynamically
          const { data: profile, error } = await realSupabase
            .from('profiles')
            .select('*')
            .eq('id', parsed.id)
            .maybeSingle();

          if (error || !profile || profile.disabled) {
            localStorage.removeItem('kairos_session');
            return null;
          }
          return profile;
        } catch {
          return null;
        }
      }
    }
  },

  // --- 2. BUSINESS SERVICES ---
  businesses: {
    list: async () => {
      if (isMock) {
        return mockDB.getBusinesses();
      } else {
        const { data, error } = await realSupabase.from('businesses').select('*').order('created_at', { ascending: false });
        if (error) throw error;
        return data;
      }
    },

    create: async (name, adminUserId) => {
      if (isMock) {
        const list = mockDB.getBusinesses();
        const exists = list.some(b => b.business_name.toLowerCase() === name.toLowerCase());
        if (exists) throw new Error('A business with this name already exists.');
        
        const newBiz = {
          id: 'b-' + Math.random().toString(36).substr(2, 9),
          business_name: name,
          status: 'ACTIVE',
          created_at: new Date().toISOString()
        };
        mockDB.setBusinesses([newBiz, ...list]);
        await dbService.auditLogs.create('CREATE_BUSINESS', 'businesses', newBiz.id, { name }, adminUserId);
        return newBiz;
      } else {
        const id = generateUUID();
        const { data, error } = await realSupabase.from('businesses').insert({ id, business_name: name }).select().single();
        if (error) throw error;
        await dbService.auditLogs.create('CREATE_BUSINESS', 'businesses', data.id, { name }, adminUserId);
        return data;
      }
    },

    setStatus: async (id, status, adminUserId) => {
      if (isMock) {
        const list = mockDB.getBusinesses();
        const idx = list.findIndex(b => b.id === id);
        if (idx === -1) throw new Error('Business not found.');
        
        list[idx].status = status;
        mockDB.setBusinesses([...list]);
        await dbService.auditLogs.create('SET_BUSINESS_STATUS', 'businesses', id, { status }, adminUserId);
        
        // Disable all profiles under this business if suspended
        if (status === 'SUSPENDED') {
          const profiles = mockDB.getProfiles();
          const updated = profiles.map(p => p.business_id === id ? { ...p, disabled: true } : p);
          mockDB.setProfiles(updated);
        }
        return list[idx];
      } else {
        const { data, error } = await realSupabase.from('businesses').update({ status }).eq('id', id).select().single();
        if (error) throw error;

        // Disable all profiles belonging to this business if suspended, or enable them if active
        if (status === 'SUSPENDED') {
          await realSupabase.from('profiles').update({ disabled: true }).eq('business_id', id);
        } else if (status === 'ACTIVE') {
          await realSupabase.from('profiles').update({ disabled: false }).eq('business_id', id);
        }

        await dbService.auditLogs.create('SET_BUSINESS_STATUS', 'businesses', id, { status }, adminUserId);
        return data;
      }
    }
  },

  // --- 3. PROFILE / USER SERVICES ---
  profiles: {
    list: async (businessId = null) => {
      if (isMock) {
        const profiles = mockDB.getProfiles();
        if (businessId) return profiles.filter(p => p.business_id === businessId);
        return profiles;
      } else {
        let query = realSupabase.from('profiles').select('*');
        if (businessId) query = query.eq('business_id', businessId);
        const { data, error } = await query.order('created_at', { ascending: false });
        if (error) throw error;
        return data;
      }
    },

    create: async (username, name, role, phone, password, businessId, adminUserId) => {
      const cleanUsername = username.trim().toLowerCase();
      const hashedPassword = await hashPassword(password);
      if (isMock) {
        const profiles = mockDB.getProfiles();
        if (profiles.some(p => p.username === cleanUsername)) {
          throw new Error('Username is already taken.');
        }

        const newProfile = {
          id: 'u-' + Math.random().toString(36).substr(2, 9),
          business_id: businessId,
          username: cleanUsername,
          role,
          name,
          phone_number: phone,
          password: hashedPassword,
          disabled: false,
          created_at: new Date().toISOString()
        };

        mockDB.setProfiles([newProfile, ...profiles]);
        await dbService.auditLogs.create('PROVISION_USER', 'profiles', newProfile.id, { username: cleanUsername, role, name }, adminUserId);
        return newProfile;
      } else {
        const newProfile = {
          id: generateUUID(),
          business_id: businessId || null,
          username: cleanUsername,
          role,
          name,
          phone_number: phone || null,
          password: hashedPassword,
          disabled: false
        };

        const { data, error } = await realSupabase
          .from('profiles')
          .insert(newProfile)
          .select()
          .single();

        if (error) throw error;
        await dbService.auditLogs.create('PROVISION_USER', 'profiles', data.id, { username: cleanUsername, role, name }, adminUserId);
        return data;
      }
    },

    setDisabled: async (id, disabled, adminUserId) => {
      if (isMock) {
        const profiles = mockDB.getProfiles();
        const idx = profiles.findIndex(p => p.id === id);
        if (idx === -1) throw new Error('Profile not found.');
        
        profiles[idx].disabled = disabled;
        mockDB.setProfiles([...profiles]);
        await dbService.auditLogs.create(disabled ? 'BLOCK_USER' : 'UNBLOCK_USER', 'profiles', id, { username: profiles[idx].username }, adminUserId);
        return profiles[idx];
      } else {
        const { data, error } = await realSupabase.from('profiles').update({ disabled }).eq('id', id).select().single();
        if (error) throw error;
        await dbService.auditLogs.create(disabled ? 'BLOCK_USER' : 'UNBLOCK_USER', 'profiles', id, {}, adminUserId);
        return data;
      }
    }
  },

  // --- 4. CUSTOMER SERVICES ---
  customers: {
    list: async (businessId) => {
      if (isMock) {
        const list = mockDB.getCustomers();
        return list.filter(c => c.business_id === businessId);
      } else {
        const { data, error } = await realSupabase.from('customers').select('*').eq('business_id', businessId).order('customer_name');
        if (error) throw error;
        return data;
      }
    },

    create: async (data, userId) => {
      if (isMock) {
        const list = mockDB.getCustomers();
        const newCust = {
          id: 'c-' + Math.random().toString(36).substr(2, 9),
          ...data,
          custom_fields: data.custom_fields || {},
          created_at: new Date().toISOString()
        };
        mockDB.setCustomers([newCust, ...list]);
        await dbService.auditLogs.create('CREATE_CUSTOMER', 'customers', newCust.id, { name: data.customer_name }, userId, data.business_id);
        return newCust;
      } else {
        const insertData = {
          id: generateUUID(),
          ...data
        };
        const { data: res, error } = await realSupabase.from('customers').insert(insertData).select().single();
        if (error) throw error;
        await dbService.auditLogs.create('CREATE_CUSTOMER', 'customers', res.id, { name: data.customer_name }, userId, data.business_id);
        return res;
      }
    },

    update: async (id, data, userId) => {
      if (isMock) {
        const list = mockDB.getCustomers();
        const idx = list.findIndex(c => c.id === id);
        if (idx === -1) throw new Error('Customer not found.');

        list[idx] = { ...list[idx], ...data };
        mockDB.setCustomers([...list]);
        await dbService.auditLogs.create('UPDATE_CUSTOMER', 'customers', id, { name: data.customer_name }, userId, list[idx].business_id);
        return list[idx];
      } else {
        const { data: res, error } = await realSupabase.from('customers').update(data).eq('id', id).select().single();
        if (error) throw error;
        await dbService.auditLogs.create('UPDATE_CUSTOMER', 'customers', id, { name: data.customer_name }, userId, data.business_id);
        return res;
      }
    },

    delete: async (id, businessId, userId) => {
      if (isMock) {
        const list = mockDB.getCustomers();
        const idx = list.findIndex(c => c.id === id && c.business_id === businessId);
        if (idx === -1) throw new Error('Customer not found.');

        const name = list[idx].customer_name;
        const updated = list.filter(c => c.id !== id);
        mockDB.setCustomers(updated);
        await dbService.auditLogs.create('DELETE_CUSTOMER', 'customers', id, { name }, userId, businessId);
        return true;
      } else {
        const { error } = await realSupabase.from('customers').delete().eq('id', id).eq('business_id', businessId);
        if (error) throw error;
        await dbService.auditLogs.create('DELETE_CUSTOMER', 'customers', id, {}, userId, businessId);
        return true;
      }
    },

    bulkImport: async (customersArray, businessId, userId) => {
      if (isMock) {
        const list = mockDB.getCustomers();
        const imported = customersArray.map(c => ({
          id: 'c-' + Math.random().toString(36).substr(2, 9),
          business_id: businessId,
          assigned_worker_id: c.assigned_worker_id || null,
          customer_name: c.customer_name,
          street_name: c.street_name,
          box_id: c.box_id,
          phone_number: c.phone_number || '',
          connection_status: c.connection_status || 'ACTIVE',
          notes: c.notes || '',
          custom_fields: c.custom_fields || {},
          created_at: new Date().toISOString()
        }));

        mockDB.setCustomers([...imported, ...list]);
        await dbService.auditLogs.create('BULK_IMPORT_CUSTOMERS', 'customers', null, { count: customersArray.length }, userId, businessId);
        return imported;
      } else {
        const rows = customersArray.map(c => ({
          id: generateUUID(),
          ...c,
          business_id: businessId
        }));
        const { data, error } = await realSupabase.from('customers').insert(rows).select();
        if (error) throw error;
        await dbService.auditLogs.create('BULK_IMPORT_CUSTOMERS', 'customers', null, { count: customersArray.length }, userId, businessId);
        return data;
      }
    }
  },

  // --- 5. PAYMENT / COLLECTION SERVICES ---
  payments: {
    list: async (businessId) => {
      if (isMock) {
        const payments = mockDB.getPayments();
        return payments.filter(p => p.business_id === businessId);
      } else {
        const { data, error } = await realSupabase.from('payments').select('*').eq('business_id', businessId).order('payment_date', { ascending: false });
        if (error) throw error;
        return data;
      }
    },

    create: async (data, userId) => {
      if (isMock) {
        const list = mockDB.getPayments();
        const newPay = {
          id: 'p-' + Math.random().toString(36).substr(2, 9),
          ...data,
          created_at: new Date().toISOString()
        };
        mockDB.setPayments([newPay, ...list]);
        await dbService.auditLogs.create('COLLECT_PAYMENT', 'payments', newPay.id, { amount: data.amount, period: data.payment_period }, userId, data.business_id);
        return newPay;
      } else {
        const insertData = {
          id: generateUUID(),
          ...data
        };
        const { data: res, error } = await realSupabase.from('payments').insert(insertData).select().single();
        if (error) throw error;
        await dbService.auditLogs.create('COLLECT_PAYMENT', 'payments', res.id, { amount: data.amount, period: data.payment_period }, userId, data.business_id);
        return res;
      }
    },

    update: async (id, data, userId) => {
      if (isMock) {
        const list = mockDB.getPayments();
        const idx = list.findIndex(p => p.id === id);
        if (idx === -1) throw new Error('Payment record not found.');

        list[idx] = { ...list[idx], ...data };
        mockDB.setPayments([...list]);
        await dbService.auditLogs.create('UPDATE_PAYMENT', 'payments', id, { amount: data.amount, period: data.payment_period }, userId, list[idx].business_id);
        return list[idx];
      } else {
        const { data: res, error } = await realSupabase.from('payments').update(data).eq('id', id).select().single();
        if (error) throw error;
        await dbService.auditLogs.create('UPDATE_PAYMENT', 'payments', id, { amount: data.amount, period: data.payment_period }, userId, data.business_id);
        return res;
      }
    },

    createBulk: async (customerId, periods, totalAmount, workerId, businessId, notes, boxId, userId) => {
      const perPeriodAmount = (totalAmount / periods.length).toFixed(2);
      if (isMock) {
        const list = mockDB.getPayments();
        const newPayments = periods.map(period => ({
          id: 'p-' + Math.random().toString(36).substr(2, 9),
          business_id: businessId,
          customer_id: customerId,
          box_id: boxId,
          worker_id: workerId,
          amount: parseFloat(perPeriodAmount),
          payment_date: new Date().toISOString().split('T')[0],
          payment_period: period,
          status: 'Paid',
          notes: notes || 'Bulk payment log',
          created_at: new Date().toISOString()
        }));

        mockDB.setPayments([...newPayments, ...list]);
        await dbService.auditLogs.create('BULK_COLLECT_PAYMENT', 'payments', null, { amount: totalAmount, periodsCount: periods.length }, userId, businessId);
        return newPayments;
      } else {
        const rows = periods.map(period => ({
          id: generateUUID(),
          business_id: businessId,
          customer_id: customerId,
          box_id: boxId,
          worker_id: workerId,
          amount: parseFloat(perPeriodAmount),
          payment_date: new Date().toISOString().split('T')[0],
          payment_period: period,
          status: 'Paid',
          notes: notes || 'Bulk payment log'
        }));

        const { data, error } = await realSupabase.from('payments').insert(rows).select();
        if (error) throw error;
        await dbService.auditLogs.create('BULK_COLLECT_PAYMENT', 'payments', null, { amount: totalAmount, periodsCount: periods.length }, userId, businessId);
        return data;
      }
    },

    bulkImport: async (paymentsArray, businessId, userId) => {
      if (isMock) {
        const list = mockDB.getPayments();
        const imported = paymentsArray.map(p => ({
          id: 'p-' + Math.random().toString(36).substr(2, 9),
          business_id: businessId,
          customer_id: p.customer_id,
          box_id: p.box_id,
          worker_id: p.worker_id || null,
          amount: parseFloat(p.amount),
          payment_date: p.payment_date,
          payment_period: p.payment_period,
          status: p.status || 'Paid',
          notes: p.notes || '',
          created_at: new Date().toISOString()
        }));

        mockDB.setPayments([...imported, ...list]);
        await dbService.auditLogs.create('BULK_IMPORT_PAYMENTS', 'payments', null, { count: paymentsArray.length }, userId, businessId);
        return imported;
      } else {
        const rows = paymentsArray.map(p => ({
          id: generateUUID(),
          ...p,
          business_id: businessId
        }));
        const { data, error } = await realSupabase.from('payments').insert(rows).select();
        if (error) throw error;
        await dbService.auditLogs.create('BULK_IMPORT_PAYMENTS', 'payments', null, { count: paymentsArray.length }, userId, businessId);
        return data;
      }
    }
  },

  // --- 6. AUDIT LOG SERVICES ---
  auditLogs: {
    list: async (businessId = null) => {
      if (isMock) {
        const audits = mockDB.getAudits();
        if (businessId) return audits.filter(a => a.business_id === businessId);
        return audits;
      } else {
        let query = realSupabase.from('audit_logs').select('*');
        if (businessId) query = query.eq('business_id', businessId);
        const { data, error } = await query.order('created_at', { ascending: false });
        if (error) throw error;
        return data;
      }
    },

    create: async (action, entityType, entityId, details, userId, businessId = null) => {
      const log = {
        business_id: businessId,
        user_id: userId,
        action,
        entity_type: entityType,
        entity_id: entityId,
        details: details || {}
      };

      if (isMock) {
        const list = mockDB.getAudits();
        const newLog = {
          id: 'l-' + Math.random().toString(36).substr(2, 9),
          ...log,
          created_at: new Date().toISOString()
        };
        mockDB.setAudits([newLog, ...list]);
        return newLog;
      } else {
        const insertLog = {
          id: generateUUID(),
          ...log
        };
        const { data, error } = await realSupabase.from('audit_logs').insert(insertLog).select().single();
        if (error) throw error;
        return data;
      }
    }
  }
};
export default dbService;
