# BusinessDesk Backend (Scaffold)

Prereqs:
- Node.js (v16+)
- npm

Install:
- npm install

Migrate DB:
- npm run migrate

Seed (creates admin user):
- npm run seed

Start:
- npm start

Default admin (seed):
- email: admin@example.com
- password: admin123
Change password immediately in production.

Notes:
- Uses SQLite at ./data/dev.sqlite3 for development. Switch to Postgres by updating knexfile.js.
- Sequences: production sequences are in `sequences` table; sample sequences are isolated in `sample_sequences`.
- Use `/seq/peek/invoice?date=YYYY-MM-DD` to preview numbers without consuming sequences.
- Creating a sales/purchase/voucher will consume the production sequence for the chosen document date.
- Add HTTPS, rate-limiting, logging, backups for production.
