# Excel Order Importer

Next.js App Router + TypeScript implementation for multi-template Excel batch order import.

## Local

```bash
npm install
npm run dev
```

## Database

If `POSTGRES_URL` or `DATABASE_URL` is present, imported orders are stored in Postgres.
Otherwise the app falls back to a local `.data/orders.json` file for development.
