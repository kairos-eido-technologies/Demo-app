-- =============================================================
-- KAIROS CABLE TV SAAS - SEED SYSTEM SUPER ADMIN ONLY
-- =============================================================

-- Insert corresponding Profile record in public.profiles table
-- Username: admin_rhemi (maps from admin_rhemi@kairos.com login)
-- Password: pasword123
INSERT INTO public.profiles (
  id,
  business_id,
  username,
  password, -- Direct password column
  role,
  name,
  phone_number,
  disabled,
  created_at
)
VALUES (
  'a1111111-1111-1111-1111-111111111111', -- Fixed UUID for Admin
  NULL, -- Super Admin has no business tenant limit
  'admin_rhemi',
  'pasword123',
  'SUPER_ADMIN',
  'Rhemi (Super Admin)',
  '+919999999999',
  false,
  NOW()
)
ON CONFLICT (username) DO UPDATE SET 
  password = EXCLUDED.password,
  disabled = EXCLUDED.disabled;
