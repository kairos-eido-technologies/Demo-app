# Kairos Cable SaaS — Customer & Collection Management System

Welcome to the **Kairos Cable SaaS** management platform, developed by **Kairos Edio Technologies**. This application is a multi-tenant operations console for Cable TV and subscription operators, facilitating secure customer directories, ledger payments, and worker progress tracking.

---

## 🚀 Getting Started

### 1. Prerequisites
Ensure you have [Node.js](https://nodejs.org/) (version 18 or higher) installed.

### 2. Installation
Install project dependencies by running:
```bash
npm install
```

### 3. Local Development (Sandboxed Mock Mode)
Start the Vite local development server:
```bash
npm run dev
```
Open [http://localhost:5173](http://localhost:5173) in your browser.
By default, the application runs in a **Sandboxed Mock Mode** using local storage simulation. You can test immediately using pre-seeded profiles:
* **Super Admin**: `kairos_admin` (password: `password123`)
* **Business Owner**: `owner_city` (password: `password123`)
* **Worker Collector**: `worker_city` (password: `password123`)

---

## ⚡ Connecting to Your Live Supabase Database

To connect the application to an actual Supabase cloud database instance:

### Step 1: Run SQL Migrations
1. Open your [Supabase Dashboard](https://supabase.com).
2. Go to the **SQL Editor** tab of your project.
3. Paste the contents of [db/schema.sql](file:///d:/KET%20APP/Cable%20New%20Version/db/schema.sql) and click **Run**.
4. This creates tables for `businesses`, `profiles`, `customers`, `payments`, and `audit_logs` with pre-configured search indexes and Row Level Security (RLS) policies.

### Step 2: Set Environment Variables
1. Copy the file `.env.example` to a new file named `.env` in the root directory:
   ```bash
   cp .env.example .env
   ```
2. Open `.env` and fill in your Supabase credentials:
   ```env
   VITE_SUPABASE_URL=https://your-project-id.supabase.co
   VITE_SUPABASE_ANON_KEY=your-anon-key-here
   ```
3. Restart the development server:
   ```bash
   npm run dev
   ```
4. The database driver wrapper (`src/supabase.js`) will automatically detect the variables and route all reads and writes to your live cloud database instead of the local storage mock!

---

## 💎 Features & UI Enhancements

- **Segmented Layout View**: The Customers layout view toggles (Standard List, Street Group, Box ID Sort) are styled as a professional, full-width segmented control bar.
- **Collapsible Street Group accordion**: Customers grouped by street names are collapsed by default. Tapping on a street banner animates the Chevron indicator and reveals the subscribers registered under that street, keeping directory scrolling compact.
- **Today vs. Total Collections Metrics**: Under Owner and Worker screens, you can track current daily earnings (`Today Earning`) and lifetime earnings (`Total Earning`) instantly.
- **Super Admin Portal**: Create tenants, suspend/activate client business accounts, provision user IDs, track security audit logs, and simulate Android APK compilation wrapper wrappers.
