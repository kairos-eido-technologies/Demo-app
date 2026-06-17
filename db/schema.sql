-- =============================================================
-- KAIROS CABLE TV SAAS - DATABASE-ONLY SCHEMA MIGRATION
-- =============================================================

-- 1. Create User Roles ENUM
DO $$
BEGIN
    IF NOT EXISTS (SELECT 1 FROM pg_type WHERE typname = 'user_role') THEN
        CREATE TYPE user_role AS ENUM ('SUPER_ADMIN', 'OWNER', 'WORKER');
    END IF;
END$$;

-- 2. Create Tables

-- A. Table: public.businesses
CREATE TABLE IF NOT EXISTS public.businesses (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  business_name TEXT UNIQUE NOT NULL,
  status TEXT NOT NULL DEFAULT 'ACTIVE',
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- B. Table: public.profiles
CREATE TABLE IF NOT EXISTS public.profiles (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  business_id UUID REFERENCES public.businesses(id) ON DELETE CASCADE,
  username TEXT UNIQUE NOT NULL,
  password TEXT NOT NULL, -- Direct password column for database-only login
  role user_role NOT NULL,
  name TEXT NOT NULL,
  phone_number TEXT,
  disabled BOOLEAN NOT NULL DEFAULT false,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- C. Table: public.customers
CREATE TABLE IF NOT EXISTS public.customers (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  business_id UUID REFERENCES public.businesses(id) ON DELETE CASCADE,
  assigned_worker_id UUID REFERENCES public.profiles(id) ON DELETE SET NULL,
  customer_name TEXT NOT NULL,
  street_name TEXT NOT NULL,
  box_id TEXT NOT NULL,
  phone_number TEXT,
  connection_status TEXT NOT NULL DEFAULT 'ACTIVE',
  notes TEXT,
  custom_fields JSONB DEFAULT '{}',
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- D. Table: public.payments
CREATE TABLE IF NOT EXISTS public.payments (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  business_id UUID REFERENCES public.businesses(id) ON DELETE CASCADE,
  customer_id UUID REFERENCES public.customers(id) ON DELETE CASCADE,
  box_id TEXT NOT NULL,
  worker_id UUID REFERENCES public.profiles(id) ON DELETE SET NULL,
  amount NUMERIC(12, 2) NOT NULL,
  payment_date DATE NOT NULL DEFAULT CURRENT_DATE,
  payment_period TEXT NOT NULL,
  status TEXT NOT NULL DEFAULT 'Paid',
  notes TEXT,
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- E. Table: public.audit_logs
CREATE TABLE IF NOT EXISTS public.audit_logs (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  business_id UUID REFERENCES public.businesses(id) ON DELETE CASCADE,
  user_id UUID REFERENCES public.profiles(id) ON DELETE SET NULL,
  action TEXT NOT NULL,
  entity_type TEXT NOT NULL,
  entity_id UUID,
  details JSONB DEFAULT '{}',
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- 3. Setup Indexes for High-Performance Queries
CREATE INDEX IF NOT EXISTS idx_customers_tenant_search 
  ON public.customers (business_id, customer_name, box_id, street_name);

CREATE INDEX IF NOT EXISTS idx_payments_tenant_search 
  ON public.payments (business_id, customer_id, payment_date);

CREATE INDEX IF NOT EXISTS idx_profiles_username 
  ON public.profiles (username);

-- Disable Row Level Security (RLS) on all tables for Multi-Tenancy (custom DB login)
ALTER TABLE public.businesses DISABLE ROW LEVEL SECURITY;
ALTER TABLE public.profiles DISABLE ROW LEVEL SECURITY;
ALTER TABLE public.customers DISABLE ROW LEVEL SECURITY;
ALTER TABLE public.payments DISABLE ROW LEVEL SECURITY;
ALTER TABLE public.audit_logs DISABLE ROW LEVEL SECURITY;


