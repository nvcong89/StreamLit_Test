# Deployment Architecture

## Recommended flow
Users -> Internal HTTPS endpoint -> Reverse proxy (Nginx/App Gateway) -> Streamlit container -> Excel workbook in SharePoint/shared storage

## Why this fits this dataset
- Business users can keep editing Excel.
- The dashboard app reads the workbook and visualizes utilization.
- Infrastructure stays simple for an internal team dashboard.

## Minimal production controls
- Keep the app private behind VPN or internal network.
- Add SSO with Microsoft Entra ID (OIDC) if needed.
- Run with read-only access to the workbook source.
- Log file refresh failures and add a health check.
