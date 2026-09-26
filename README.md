# IT Dept Timetable Management System

Faculty Master, Subject Master, Subject-Faculty Allotment, Our Department /
Other Department timetables (versioned, with faculty-conflict checking), and
class-wise / individual-faculty PDF + Excel reports — backed by a real
Supabase (Postgres) database instead of browser storage.

The Individual Faculty Report shows three grids per faculty — Our
Department only, Other Department only, and Combined — plus a subject
allotment table with a workload total.

## 1. Supabase setup

The database is already live: project `ttms-app`
(`jwodlfukmfdbkpbdsvoz.supabase.co`), schema applied, all 6 tables in place
with row-level security. If you ever need to recreate it elsewhere, run
`supabase/schema.sql` once in the SQL editor of a fresh project.

To sign in, create your own account in **Authentication → Users** on that
project (email + password) — there's no self-signup, this is a
single-admin tool.

## 2. Local setup

```bash
npm install
npm run dev
```

`.env` is already filled in with this project's URL and anon key (see
`.env.example` for the format if you need to point at a different project).
Open the printed local URL and sign in with the account you created above.

## 3. Deploy

### GitHub
```bash
git init
git add -A
git commit -m "Initial commit: Timetable Management System"
git branch -M main
git remote add origin https://github.com/<you>/<repo>.git
git push -u origin main
```

### Vercel
1. Import the GitHub repo in Vercel.
2. Framework preset: **Vite**.
3. Add environment variables `VITE_SUPABASE_URL` and `VITE_SUPABASE_ANON_KEY`
   (same values as `.env`) under Project Settings → Environment Variables.
4. Deploy. Every push to `main` auto-deploys.

## Data model notes

- `faculty`, `subject_master`, `allotments` are normal relational tables.
- `our_versions` (Our Department) and `other_timetables` (Other Department)
  keep their `meta`, `subjects`, and `grid` as JSONB columns — each
  timetable's subject list and 54-box grid is naturally a nested structure,
  so it's stored as one document per timetable/version.
- Only one `our_versions` row per class can have `is_active = true` at a
  time (enforced by a partial unique index) — that's the version used for
  faculty-conflict checks, the individual faculty timetable, and reports.
  Older versions stay in the table for history.
- Row Level Security requires a signed-in Supabase Auth user for every
  read/write — no public/anonymous access.

## What's NOT included yet

- No self-service signup or password reset flow (single admin account,
  managed from the Supabase dashboard).
- No realtime sync between multiple browser tabs/devices — use the
  "Reload" button in the top bar if you've made changes elsewhere.
- No automated tests.
- Not yet pushed to GitHub or deployed to Vercel — that's the next step.
