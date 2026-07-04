# Resource Planning Demo

## Files
- `resource_planning_sample.xlsx`: sample input workbook
- `app.py`: Streamlit dashboard
- `requirements.txt`: Python dependencies
- `Dockerfile`: container deployment
- `docker-compose.yml`: local container test

## Local run
```bash
python -m venv .venv
source .venv/bin/activate  # Linux/macOS
pip install -r requirements.txt
streamlit run app.py
```

## Docker run
```bash
docker compose up --build
```

## Recommended internal deployment
1. Store the Excel workbook in SharePoint or a controlled shared folder.
2. Run the Streamlit container on an internal VM or managed app service.
3. Put Nginx or an internal application gateway in front for HTTPS.
4. Use Microsoft Entra ID / OIDC if you add authentication.
5. Grant read-only access to the service identity for the data file location.
