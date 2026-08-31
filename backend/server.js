// server.js
// BusinessDesk backend scaffold (Express + SQLite + Knex)

const express = require('express');
const cors = require('cors');
const knexLib = require('knex');
const { body, validationResult } = require('express-validator');
const bcrypt = require('bcrypt');
const jwt = require('jsonwebtoken');
const ExcelJS = require('exceljs');
const { create } = require('xmlbuilder2');
const { stringify } = require('csv-stringify/sync');
const { v4: uuidv4 } = require('uuid');
const path = require('path');

const ENV_SECRET = process.env.BD_SECRET || 'dev_secret_change_me';
const PORT = process.env.PORT || 4000;

const knex = knexLib(require('./knexfile').development);

// ----------------- Utilities & Sequences -----------------

function _ymFromDate(dateStr){
  if(!dateStr){
    const n = new Date();
    return `${n.getFullYear()}${String(n.getMonth()+1).padStart(2,'0')}`;
  }
  const d = new Date(dateStr);
  if(isNaN(d)){
    const n = new Date();
    return `${n.getFullYear()}${String(n.getMonth()+1).padStart(2,'0')}`;
  }
  return `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}`;
}

async function peekSeq(prefix, dateStr, opts={ sample: false }){
  const table = opts.sample ? 'sample_sequences' : 'sequences';
  const ym = _ymFromDate(dateStr);
  const key = `${prefix}-${ym}`;
  const row = await knex(table).where({ key }).first();
  const v = row ? row.value : 0;
  return `${prefix}-${ym}-${String(v + 1).padStart(4,'0')}`;
}

async function nextSeq(prefix, dateStr, opts={ sample: false }){
  const table = opts.sample ? 'sample_sequences' : 'sequences';
  const ym = _ymFromDate(dateStr);
  const key = `${prefix}-${ym}`;
  const trx = await knex.transaction();
  try {
    const row = await trx(table).where({ key }).first();
    if(row){
      await trx(table).where({ key }).update({ value: row.value + 1 });
      await trx.commit();
      return `${prefix}-${ym}-${String(row.value + 1).padStart(4,'0')}`;
    } else {
      await trx(table).insert({ key, value: 1 });
      await trx.commit();
      return `${prefix}-${ym}-0001`;
    }
  } catch(err){
    await trx.rollback();
    throw err;
  }
}

function todayLocal(){
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}

function computeGST(taxable, rate, type){
  const t = Number(taxable || 0), r = Number(rate || 0);
  if(!isFinite(t) || !isFinite(r)) return { cgst:0, sgst:0, igst:0, totalGST:0, total:+t.toFixed(2) };
  if(type === 'IGST'){
    const igst = +(t * r / 100).toFixed(2);
    return { cgst:0, sgst:0, igst, totalGST: igst, total: +(t + igst).toFixed(2) };
  } else {
    const half = r/2;
    const cgst = +(t * half / 100).toFixed(2);
    const sgst = +(t * half / 100).toFixed(2);
    const total = +(t + cgst + sgst).toFixed(2);
    return { cgst, sgst, igst:0, totalGST: +(cgst + sgst).toFixed(2), total };
  }
}

function computeOutstanding(total, paid){
  total = Number(total || 0); paid = Number(paid || 0);
  const diff = +(total - paid);
  if(diff >= 0) return { outstanding: +diff.toFixed(2), credit: 0 };
  return { outstanding: 0, credit: +Math.abs(diff).toFixed(2) };
}

// ----------------- Migrations & Seed -----------------

async function migrate(){
  // users
  if(!(await knex.schema.hasTable('users'))){
    await knex.schema.createTable('users', t=>{
      t.increments('id').primary();
      t.string('email').notNullable().unique();
      t.string('password').notNullable();
      t.string('role').notNullable().defaultTo('viewer');
      t.string('name').nullable();
      t.timestamp('created_at').defaultTo(knex.fn.now());
    });
  }

  if(!(await knex.schema.hasTable('parties'))){
    await knex.schema.createTable('parties', t=>{
      t.string('id').primary();
      t.string('name').notNullable();
      t.string('type').notNullable();
      t.string('gstin').nullable();
      t.string('phone').nullable();
      t.string('state').nullable();
      t.timestamp('created_at').defaultTo(knex.fn.now());
    });
  }

  if(!(await knex.schema.hasTable('sales'))){
    await knex.schema.createTable('sales', t=>{
      t.string('id').primary();
      t.string('invoice_no').notNullable().unique();
      t.date('date').notNullable();
      t.string('customer_id').references('id').inTable('parties').onDelete('SET NULL');
      t.decimal('taxable',14,2).notNullable().defaultTo(0);
      t.decimal('gst_rate',6,2).notNullable().defaultTo(0);
      t.string('gst_type').notNullable().defaultTo('CGST+SGST');
      t.decimal('cgst',14,2).defaultTo(0);
      t.decimal('sgst',14,2).defaultTo(0);
      t.decimal('igst',14,2).defaultTo(0);
      t.decimal('total_gst',14,2).defaultTo(0);
      t.decimal('total',14,2).notNullable().defaultTo(0);
      t.decimal('paid',14,2).defaultTo(0);
      t.decimal('outstanding',14,2).defaultTo(0);
      t.decimal('credit',14,2).defaultTo(0);
      t.timestamp('created_at').defaultTo(knex.fn.now());
    });
  }

  if(!(await knex.schema.hasTable('purchases'))){
    await knex.schema.createTable('purchases', t=>{
      t.string('id').primary();
      t.string('bill_no').notNullable().unique();
      t.date('date').notNullable();
      t.string('supplier_id').references('id').inTable('parties').onDelete('SET NULL');
      t.decimal('taxable',14,2).notNullable().defaultTo(0);
      t.decimal('gst_rate',6,2).notNullable().defaultTo(0);
      t.string('gst_type').notNullable().defaultTo('CGST+SGST');
      t.decimal('cgst',14,2).defaultTo(0);
      t.decimal('sgst',14,2).defaultTo(0);
      t.decimal('igst',14,2).defaultTo(0);
      t.decimal('total_gst',14,2).defaultTo(0);
      t.decimal('total',14,2).notNullable().defaultTo(0);
      t.decimal('paid',14,2).defaultTo(0);
      t.decimal('outstanding',14,2).defaultTo(0);
      t.decimal('credit',14,2).defaultTo(0);
      t.timestamp('created_at').defaultTo(knex.fn.now());
    });
  }

  if(!(await knex.schema.hasTable('vouchers'))){
    await knex.schema.createTable('vouchers', t=>{
      t.string('id').primary();
      t.string('voucher_no').notNullable().unique();
      t.date('date').notNullable();
      t.string('type').notNullable();
      t.string('debit').notNullable();
      t.string('credit').notNullable();
      t.decimal('amount',14,2).notNullable().defaultTo(0);
      t.timestamp('created_at').defaultTo(knex.fn.now());
    });
  }

  if(!(await knex.schema.hasTable('sequences'))){
    await knex.schema.createTable('sequences', t=>{
      t.string('key').primary();
      t.integer('value').notNullable().defaultTo(0);
    });
  }

  if(!(await knex.schema.hasTable('sample_sequences'))){
    await knex.schema.createTable('sample_sequences', t=>{
      t.string('key').primary();
      t.integer('value').notNullable().defaultTo(0);
    });
  }

  if(!(await knex.schema.hasTable('audit_logs'))){
    await knex.schema.createTable('audit_logs', t=>{
      t.increments('id').primary();
      t.string('entity').notNullable();
      t.string('entity_id').notNullable();
      t.string('action').notNullable();
      t.json('payload').nullable();
      t.string('user_id').nullable();
      t.timestamp('created_at').defaultTo(knex.fn.now());
    });
  }

  console.log('Migrations done');
}

async function seed(){
  const users = await knex('users').select();
  if(users.length === 0){
    const hashed = await bcrypt.hash('admin123', 10);
    await knex('users').insert({ email: 'admin@example.com', password: hashed, role: 'admin', name: 'Administrator' });
    console.log('Seeded admin: admin@example.com / admin123');
  }
}

// ----------------- Auth, audit -----------------

async function audit(entity, entityId, action, payload, userId=null){
  try {
    await knex('audit_logs').insert({ entity, entity_id: entityId, action, payload: JSON.stringify(payload||{}), user_id: userId });
  } catch(e) {
    console.error('Audit write failed', e.message);
  }
}

function authMiddleware(req, res, next){
  const auth = req.headers.authorization;
  if(!auth) return res.status(401).json({ error: 'Missing token' });
  const token = auth.split(' ')[1];
  if(!token) return res.status(401).json({ error: 'Malformed token' });
  try {
    const payload = jwt.verify(token, ENV_SECRET);
    req.user = payload;
    next();
  } catch(e) {
    return res.status(401).json({ error: 'Invalid token' });
  }
}

function roleGuard(roles = []){
  return (req, res, next) => {
    if(!req.user) return res.status(401).json({ error: 'Not authenticated' });
    if(roles.length === 0 || roles.includes(req.user.role)) return next();
    return res.status(403).json({ error: 'Forbidden' });
  };
}

// ----------------- App & Routes -----------------

async function createApp(){
  const app = express();
  app.use(cors());
  app.use(express.json());

  // Health
  app.get('/health', (req,res)=> res.json({ ok: true }));

  // Auth
  app.post('/auth/login', [
    body('email').isEmail(),
    body('password').isString()
  ], async (req,res) => {
    const errors = validationResult(req);
    if(!errors.isEmpty()) return res.status(400).json({ errors: errors.array() });
    const { email, password } = req.body;
    const user = await knex('users').where({ email }).first();
    if(!user) return res.status(401).json({ error: 'Invalid credentials' });
    const ok = await bcrypt.compare(password, user.password);
    if(!ok) return res.status(401).json({ error: 'Invalid credentials' });
    const token = jwt.sign({ id: user.id, role: user.role, email: user.email, name: user.name }, ENV_SECRET, { expiresIn: '8h' });
    return res.json({ token });
  });

  // USERS (admin)
  app.post('/users', authMiddleware, roleGuard(['admin']), [
    body('email').isEmail(),
    body('password').isLength({ min: 6 }),
    body('role').isIn(['admin','accountant','viewer'])
  ], async (req,res) => {
    const { email, password, role, name } = req.body;
    const hashed = await bcrypt.hash(password, 10);
    try {
      const [id] = await knex('users').insert({ email, password: hashed, role, name });
      res.json({ id });
    } catch(e){
      res.status(500).json({ error: 'failed' });
    }
  });

  // PARTIES
  app.get('/parties', authMiddleware, async (req,res) => {
    const rows = await knex('parties').select();
    res.json(rows);
  });
  app.post('/parties', authMiddleware, roleGuard(['admin','accountant']), [
    body('name').notEmpty(),
    body('type').isIn(['Customer','Supplier'])
  ], async (req,res) => {
    const { name, type, gstin, phone, state } = req.body;
    const id = uuidv4();
    await knex('parties').insert({ id, name, type, gstin, phone, state });
    await audit('party', id, 'create', req.body, req.user && req.user.id);
    res.json({ id });
  });
  app.put('/parties/:id', authMiddleware, roleGuard(['admin','accountant']), async (req,res)=>{
    const id = req.params.id;
    const { name, type, gstin, phone, state } = req.body;
    await knex('parties').where({ id }).update({ name, type, gstin, phone, state });
    await audit('party', id, 'update', req.body, req.user && req.user.id);
    res.json({ ok:true });
  });
  app.delete('/parties/:id', authMiddleware, roleGuard(['admin']), async (req,res)=>{
    const id = req.params.id;
    await knex('parties').where({ id }).del();
    await audit('party', id, 'delete', {}, req.user && req.user.id);
    res.json({ ok:true });
  });

  // Sales
  app.get('/sales', authMiddleware, async (req,res) => {
    const rows = await knex('sales').select();
    res.json(rows);
  });

  app.get('/seq/peek/invoice', authMiddleware, async (req,res) => {
    const date = req.query.date || todayLocal();
    const val = await peekSeq('INV', date, { sample:false });
    res.json({ invoiceNo: val });
  });

  app.post('/sales', authMiddleware, roleGuard(['admin','accountant']), [
    body('date').isISO8601(),
    body('customerId').notEmpty(),
    body('taxable').isFloat({ min: 0 }),
    body('gstRate').isFloat({ min: 0 }),
    body('gstType').isIn(['CGST+SGST','IGST'])
  ], async (req,res) => {
    const { date, customerId, taxable, gstRate, gstType, paid } = req.body;
    const comps = computeGST(taxable, gstRate, gstType);
    const invoiceNo = await nextSeq('INV', date, { sample:false });
    const { outstanding, credit } = computeOutstanding(comps.total, paid || 0);
    const id = uuidv4();
    await knex('sales').insert({
      id, invoice_no: invoiceNo, date, customer_id: customerId,
      taxable, gst_rate: gstRate, gst_type: gstType,
      cgst: comps.cgst, sgst: comps.sgst, igst: comps.igst, total_gst: comps.totalGST,
      total: comps.total, paid: paid || 0, outstanding, credit
    });
    await audit('sale', id, 'create', req.body, req.user && req.user.id);
    res.json({ id, invoiceNo });
  });

  // Purchases
  app.get('/purchases', authMiddleware, async (req,res) => {
    const rows = await knex('purchases').select();
    res.json(rows);
  });

  app.post('/purchases', authMiddleware, roleGuard(['admin','accountant']), [
    body('date').isISO8601(),
    body('supplierId').notEmpty(),
    body('taxable').isFloat({ min: 0 }),
    body('gstRate').isFloat({ min: 0 }),
    body('gstType').isIn(['CGST+SGST','IGST'])
  ], async (req,res) => {
    const { date, supplierId, taxable, gstRate, gstType, paid } = req.body;
    const comps = computeGST(taxable, gstRate, gstType);
    const billNo = await nextSeq('PB', date, { sample:false });
    const { outstanding, credit } = computeOutstanding(comps.total, paid || 0);
    const id = uuidv4();
    await knex('purchases').insert({
      id, bill_no: billNo, date, supplier_id: supplierId,
      taxable, gst_rate: gstRate, gst_type: gstType,
      cgst: comps.cgst, sgst: comps.sgst, igst: comps.igst, total_gst: comps.totalGST,
      total: comps.total, paid: paid || 0, outstanding, credit
    });
    await audit('purchase', id, 'create', req.body, req.user && req.user.id);
    res.json({ id, billNo });
  });

  // Vouchers
  app.get('/vouchers', authMiddleware, async (req,res) => {
    const rows = await knex('vouchers').select();
    res.json(rows);
  });

  app.post('/vouchers', authMiddleware, roleGuard(['admin','accountant']), [
    body('date').isISO8601(),
    body('type').notEmpty(),
    body('debit').notEmpty(),
    body('credit').notEmpty(),
    body('amount').isFloat({ min: 0.01 })
  ], async (req,res) => {
    const { date, type, debit, credit, amount } = req.body;
    const voucherNo = await nextSeq('VCH', date, { sample:false });
    const id = uuidv4();
    await knex('vouchers').insert({ id, voucher_no: voucherNo, date, type, debit, credit, amount });
    await audit('voucher', id, 'create', req.body, req.user && req.user.id);
    res.json({ id, voucherNo });
  });

  // Exports
  app.get('/export/sales/csv', authMiddleware, async (req,res) => {
    const rows = await knex('sales').select();
    const csv = stringify(rows, { header: true });
    res.setHeader('Content-Type', 'text/csv');
    res.setHeader('Content-Disposition', 'attachment; filename="sales.csv"');
    res.send(csv);
  });

  app.get('/export/sales/xlsx', authMiddleware, async (req,res) => {
    const rows = await knex('sales').select();
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sales');
    if(rows.length > 0){
      ws.columns = Object.keys(rows[0]).map(k => ({ header: k, key: k }));
      rows.forEach(r => ws.addRow(r));
    }
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="sales.xlsx"');
    await wb.xlsx.write(res);
    res.end();
  });

  app.get('/export/sales/tally', authMiddleware, async (req,res) => {
    const sales = await knex('sales').select();
    const doc = create({ version: '1.0' }).ele('ENVELOPE').ele('DATA');
    const all = doc.ele('TALLYMESSAGE');
    for(const s of sales){
      const inv = all.ele('VOUCHER');
      inv.ele('INVOICENO').txt(s.invoice_no);
      inv.ele('DATE').txt(String(s.date));
      inv.ele('PARTY').txt(String(s.customer_id));
      inv.ele('TAXABLE').txt(String(s.taxable));
      inv.ele('GST').txt(String(s.total_gst));
      inv.ele('TOTAL').txt(String(s.total));
    }
    const xml = doc.end({ prettyPrint: true });
    res.setHeader('Content-Type', 'application/xml');
    res.setHeader('Content-Disposition', 'attachment; filename="tally_sales.xml"');
    res.send(xml);
  });

  // Audit (admin)
  app.get('/audit', authMiddleware, roleGuard(['admin']), async (req,res)=>{
    const logs = await knex('audit_logs').orderBy('created_at','desc').limit(200);
    res.json(logs);
  });

  // run migration/seed helpers exposed for CLI only (not exposed as HTTP)
  return app;
}

// ----------------- CLI entry -----------------

(async () => {
  const arg = process.argv[2];
  if(arg === 'migrate'){
    await migrate();
    process.exit(0);
  } else if(arg === 'seed'){
    await seed();
    process.exit(0);
  } else {
    // auto-migrate & seed on normal start (safe for dev)
    await migrate();
    await seed();
    const app = await createApp();
    app.listen(PORT, ()=> console.log(`BusinessDesk backend running on ${PORT}`));
  }
})();
