# AWP Dashboard setup

## Local test

Copy `.streamlit/secrets.toml.example` to `.streamlit/secrets.toml` and set an
admin password. When Supabase settings are absent, edits are saved to
`data.xlsx`.

Run:

```powershell
streamlit run app.py
```

## Persistent online storage

1. Create a Supabase project.
2. Open **SQL Editor** and run `supabase_schema.sql`.
3. In the Streamlit app settings, open **Secrets** and add:

```toml
ADMIN_PASSWORD = "your-strong-admin-password"
SUPABASE_URL = "https://YOUR-PROJECT.supabase.co"
SUPABASE_KEY = "YOUR-SERVICE-ROLE-KEY"
```

The service-role key must only be stored in Streamlit Secrets. Never commit it
to GitHub.

4. Restart the app, open **Data Management**, sign in, and import `data.xlsx`.

For additional password protection, `ADMIN_PASSWORD_SHA256` can be used instead
of `ADMIN_PASSWORD`.
