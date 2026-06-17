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

-- Enable Row Level Security (RLS) on all tables for Multi-Tenancy
ALTER TABLE public.businesses ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.profiles ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.customers ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.payments ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.audit_logs ENABLE ROW LEVEL SECURITY;

-- 1. Policies for public.businesses
CREATE POLICY "Super Admins can do anything on businesses" ON public.businesses
  FOR ALL TO authenticated USING (
    (SELECT role FROM public.profiles WHERE id = auth.uid()) = 'SUPER_ADMIN'
  );

CREATE POLICY "Owners and workers can read their own business details" ON public.businesses
  FOR SELECT TO authenticated USING (
    id = (SELECT business_id FROM public.profiles WHERE id = auth.uid())
  );

-- 2. Policies for public.profiles
CREATE POLICY "Super Admins can manage all profiles" ON public.profiles
  FOR ALL TO authenticated USING (
    (SELECT role FROM public.profiles WHERE id = auth.uid()) = 'SUPER_ADMIN'
  );

CREATE POLICY "Users can read profiles from their own business" ON public.profiles
  FOR SELECT TO authenticated USING (
    business_id = (SELECT business_id FROM public.profiles WHERE id = auth.uid())
  );

CREATE POLICY "Users can update their own profile data" ON public.profiles
  FOR UPDATE TO authenticated USING (
    id = auth.uid()
  ) WITH CHECK (
    id = auth.uid()
  );

-- 3. Policies for public.customers
CREATE POLICY "Super Admins can manage all customers" ON public.customers
  FOR ALL TO authenticated USING (
    (SELECT role FROM public.profiles WHERE id = auth.uid()) = 'SUPER_ADMIN'
  );

CREATE POLICY "Owners can manage customers of their business" ON public.customers
  FOR ALL TO authenticated USING (
    business_id = (SELECT business_id FROM public.profiles WHERE id = auth.uid())
    AND (SELECT role FROM public.profiles WHERE id = auth.uid()) = 'OWNER'
  );

CREATE POLICY "Workers can read and write customers assigned to their business" ON public.customers
  FOR ALL TO authenticated USING (
    business_id = (SELECT business_id FROM public.profiles WHERE id = auth.uid())
    AND (SELECT role FROM public.profiles WHERE id = auth.uid()) = 'WORKER'
  );

-- 4. Policies for public.payments
CREATE POLICY "Super Admins can manage all payments" ON public.payments
  FOR ALL TO authenticated USING (
    (SELECT role FROM public.profiles WHERE id = auth.uid()) = 'SUPER_ADMIN'
  );

CREATE POLICY "Owners can manage payments of their business" ON public.payments
  FOR ALL TO authenticated USING (
    business_id = (SELECT business_id FROM public.profiles WHERE id = auth.uid())
    AND (SELECT role FROM public.profiles WHERE id = auth.uid()) = 'OWNER'
  );

CREATE POLICY "Workers can read and write payments of their business" ON public.payments
  FOR ALL TO authenticated USING (
    business_id = (SELECT business_id FROM public.profiles WHERE id = auth.uid())
    AND (SELECT role FROM public.profiles WHERE id = auth.uid()) = 'WORKER'
  );

-- 5. Policies for public.audit_logs
CREATE POLICY "Super Admins can manage all audit logs" ON public.audit_logs
  FOR ALL TO authenticated USING (
    (SELECT role FROM public.profiles WHERE id = auth.uid()) = 'SUPER_ADMIN'
  );

CREATE POLICY "Owners can read audit logs of their business" ON public.audit_logs
  FOR SELECT TO authenticated USING (
    business_id = (SELECT business_id FROM public.profiles WHERE id = auth.uid())
    AND (SELECT role FROM public.profiles WHERE id = auth.uid()) = 'OWNER'
  );
