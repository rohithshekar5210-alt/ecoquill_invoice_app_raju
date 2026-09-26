ECOQUILL — INVOICE GENERATOR ONLY (no dashboard, no purchase/stock entry)

WHAT CHANGED FROM YOUR OLD APP
- Removed: Dashboard, AI Assistant, Sales Register, Purchases, Live
  Inventory, Reports, Suppliers. There is no stock tracking any more,
  so there is nothing that can "reset" between visits.
- Kept, unchanged: the exact invoice PDF design (green/dark EcoQuill
  layout, logo, signature, colours), the "Invoice Design Studio"
  customisation panel (business info, typography, invoice body
  wording/colours, logo & signature upload), invoice numbering
  (financial-year based), Invoice History with PDF re-download, and
  a simple Product Master so your bag sizes/rates still appear in a
  dropdown when billing.
- As soon as you open the app you can fill customer + product rows
  and click "Generate Invoice" — no purchase entry required first.

FILES IN THIS FOLDER
  app.py                          the whole app (one file)
  business_settings.json          your saved company/invoice settings
  requirements.txt                Python packages needed
  .streamlit_secrets_example.toml optional cloud-persistence keys

RUN LOCALLY — TWO STEPS
Open PowerShell/Terminal inside this folder.

STEP 1
  pip install -r requirements.txt

STEP 2
  streamlit run app.py

Open http://localhost:8501 (Streamlit prints the exact URL).

If you already have a virtual environment like your old project, you
can instead run:
  & "..\.venv\Scripts\python.exe" -m pip install -r .\requirements.txt
  & "..\.venv\Scripts\python.exe" -m streamlit run .\app.py --server.port 8507

DEPLOY TO STREAMLIT COMMUNITY CLOUD — TWO STEPS
STEP 1 — Push this folder to a GitHub repo, then on
  https://share.streamlit.io click "New app", pick the repo/branch,
  and set the main file to app.py. Click Deploy.

STEP 2 (optional but recommended) — Prevent invoice history from
  being lost when the app sleeps or is redeployed:
    1. Create a free Supabase project (https://supabase.com).
    2. Create a PRIVATE Storage bucket exactly named:
       ecoquill-persistence
    3. In Streamlit Cloud > your app > Settings > Secrets, paste the
       values shown in .streamlit_secrets_example.toml, replacing the
       placeholders with your real Supabase URL and service-role key.
    4. Reboot the app. The sidebar footer should say:
       "Persistent cloud storage configured".
  Without this step the app still works perfectly for generating and
  downloading invoices in the moment — only the saved history/PDFs on
  Streamlit's free tier are at risk if the container is recycled.
  (Every invoice is also offered as an immediate "Download Invoice"
  button right after it's generated, so your client copy is safe
  either way.)

NOTES
- Logo & signature: upload them once from the "Invoice Design Studio"
  → Assets tab, then click "Save Invoice Design". They're reused on
  every invoice after that.
- Product Master page: add your bag sizes/rates once so they show up
  as a dropdown on the invoice screen — this is optional, you can
  also just type a "Custom / Other" product name directly on an
  invoice.
- Everything about the invoice's look, wording, colours, bank
  details and terms is editable from "Invoice Design Studio" in the
  sidebar — nothing there changed from your original app.
