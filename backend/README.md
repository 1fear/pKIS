# TakSklad Backend MVP

This directory contains the first VDS-ready backend skeleton for TakSklad.

Current scope:

- FastAPI application shell.
- PostgreSQL connection settings.
- Initial API contracts for health, active orders, scans, imports, and day reports.
- Initial PostgreSQL schema SQL.
- Docker image definition.

This is not a production release yet. The desktop app still works directly with Google Sheets until the backend is deployed, verified, and connected behind feature flags.

## Local Docker Run

From repository root:

```bash
cp deploy/vds/.env.example deploy/vds/.env
docker compose --env-file deploy/vds/.env -f deploy/vds/docker-compose.yml up -d --build
```

When Traefik is configured on VDS, check through the configured domain:

```bash
curl https://$TAKSKLAD_BACKEND_HOST/health
```

The compose file does not publish PostgreSQL or the backend API directly to the public internet. HTTP traffic should go through Traefik.

## API Status

Implemented now:

- `GET /health`
- `GET /api/v1/orders/active`
- `POST /api/v1/scans`
- `POST /api/v1/orders/{order_id}/complete`
- `POST /api/v1/imports`
- `GET /api/v1/imports`
- `GET /api/v1/reports/day`

No contract placeholders remain in the backend API MVP.

## Day Report

`GET /api/v1/reports/day?report_date=YYYY-MM-DD`

Builds a PostgreSQL-based day summary:

- orders for the selected order date;
- orders scanned on the selected date;
- planned/scanned/remaining blocks;
- payment groups;
- SkladBot request number if imported with the order.
