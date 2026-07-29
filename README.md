# Farspeed Depot Manifest

Depot & storage tracking for Farspeed Contractors Ltd — arrivals, the 14-day
free storage clock, partial deliveries, duplicate detection, itemized
packages, job numbers, the Site & Account Directory, Employees, a Job Log,
Cancelled Jobs, printable Job Sheets, and Excel/PDF packing list import —
in English and Traditional Chinese, light and dark mode.

## Part 1 — do this now (no credit card needed)

Supabase's free tier and GitHub don't require a payment method at all.

### 1. Set up the database (Supabase)

1. Create a free project at [supabase.com](https://supabase.com).
2. Open **SQL Editor**, paste in `supabase/schema.sql`, click Run.
3. Go to **Settings → API**. Copy your **Project URL** and your
   **anon / publishable** key (if you only see "Publishable and secret API
   keys", click the "Legacy anon, service_role API keys" tab and use the
   `anon` `public` key from there instead).

### 2. Try it locally

```bash
npm install
cp .env.example .env
# open .env and paste in your Supabase Project URL and anon key
npm run dev
```

Open the printed `localhost` link. Add a test entry, refresh — if it's
still there, Supabase is wired up correctly.

### 3. Put it on GitHub

```bash
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/YOUR-USERNAME/YOUR-REPO.git
git branch -M main
git push -u origin main
```

## Part 2 — do this once you have your card

### 4. Deploy it for real (Vercel)

Import the GitHub repo at [vercel.com](https://vercel.com), add
`VITE_SUPABASE_URL` and `VITE_SUPABASE_ANON_KEY` as environment
variables, deploy. Vercel's free Hobby tier works technically but its
terms restrict it to non-commercial use — Pro ($20/month) is the
compliant option for a company tool, and that's what needs a card.

### 5. PDF packing list scanning (optional, its own card)

Unlike earlier versions of this project, this now works for EVERYONE
using the app automatically — nobody (including Irene or other staff)
needs to enter or know an API key. You set it up once, as the admin:

1. Get an API key at [console.anthropic.com](https://console.anthropic.com)
2. In your Vercel project: **Settings -> Environment Variables**
3. Add a variable named exactly `ANTHROPIC_API_KEY` with your key as the
   value. Do NOT prefix it with `VITE_` (that would expose it to the
   browser, which defeats the point).
4. Redeploy (Deployments tab -> "..." on the latest one -> Redeploy)

From then on, the "Import -> PDF Scan" button just works for everyone,
no key entry, no setup on their end. Everything else in the app works
fully without this step too, if you'd rather skip it for now.

## Important limitation: no login yet

Anyone with the app's URL can view and edit everything. Fine for a small
trusted team; add [Supabase Auth](https://supabase.com/docs/guides/auth)
later if you need real access control.
