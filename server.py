#!/usr/bin/env python3
import os
import sys
import re
import io
import json
import sqlite3
import zipfile
import urllib.parse
from http.server import HTTPServer, BaseHTTPRequestHandler
import xml.etree.ElementTree as ET

def normalize_inv(inv_str):
    if not inv_str:
        return ""
    return re.sub(r'[^A-Za-z0-9]', '', str(inv_str)).upper()

def is_tax_ledger(lname):
    lname_l = lname.lower().strip()
    # Exclude base purchase, sales, expense, income, account, round off ledgers
    if any(k in lname_l for k in ['purchase', 'sales', 'expense', 'income', 'a/c', 'account', 'round']):
        return False, None
    if 'igst' in lname_l:
        return True, 'igst'
    if 'cgst' in lname_l:
        return True, 'cgst'
    if 'sgst' in lname_l or 'utgst' in lname_l:
        return True, 'sgst'
    if 'gst' in lname_l or 'tax' in lname_l or 'duty' in lname_l or 'duties' in lname_l:
        return True, 'tax'
    return False, None

def parse_gstr2b_xlsx_bytes(file_bytes, filename=""):
    invoices = []
    file_period = ""
    if filename:
        m = re.search(r'(\d{2})(\d{4})', filename)
        if m:
            mm, yyyy = m.groups()
            months = {'01':'Jan', '02':'Feb', '03':'Mar', '04':'Apr', '05':'May', '06':'Jun', '07':'Jul', '08':'Aug', '09':'Sep', '10':'Oct', '11':'Nov', '12':'Dec'}
            file_period = f"{months.get(mm, mm)}-{yyyy}"

    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes), 'r') as z:
            ss = []
            if 'xl/sharedStrings.xml' in z.namelist():
                r = ET.fromstring(z.read('xl/sharedStrings.xml'))
                for elem in r.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si'):
                    text = ''.join([t.text or '' for t in elem.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')])
                    ss.append(text)
            
            wb_root = ET.fromstring(z.read('xl/workbook.xml'))
            ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            sheets = [(s.get('name'), s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')) for s in wb_root.findall('.//s:sheet', ns)]
            rels_root = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
            r_map = {rel.get('Id'): rel.get('Target') for rel in rels_root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship')}
            
            for name, rId in sheets:
                name_upper = name.upper()
                if name_upper in ['B2B', 'B2BA', 'B2B-CDNR', 'CDNR', 'B2B-CDNRA', 'CDNRA', 'ITC AVAILABLE']:
                    target = r_map.get(rId, '')
                    sheet_path = 'xl/' + target if not target.startswith('xl/') else target
                    if sheet_path in z.namelist():
                        s_root = ET.fromstring(z.read(sheet_path))
                        is_cdnr = 'CDNR' in name_upper
                        
                        for row in s_root.findall('.//s:row', ns):
                            r_vals = []
                            for c in row.findall('.//s:c', ns):
                                t = c.get('t')
                                v = c.find('s:v', ns)
                                val_str = v.text if v is not None else ''
                                if t == 's' and val_str.isdigit() and int(val_str) < len(ss):
                                    val_str = ss[int(val_str)]
                                r_vals.append(val_str)
                            
                            if len(r_vals) >= 12 and len(r_vals[0]) == 15 and r_vals[0][:2].isdigit():
                                doc_type = 'Invoice'
                                if is_cdnr:
                                    doc_type = r_vals[3].strip() if len(r_vals) > 3 and r_vals[3] else 'Credit/Debit Note'
                                    inum = r_vals[2].strip()
                                    idate = r_vals[5].strip() if len(r_vals) > 5 else ''
                                    val = float(r_vals[6]) if len(r_vals) > 6 and r_vals[6] else 0.0
                                    rcm_raw = r_vals[8].strip() if len(r_vals) > 8 else ''
                                    txval = float(r_vals[9]) if len(r_vals) > 9 and r_vals[9] else 0.0
                                    igst = float(r_vals[10]) if len(r_vals) > 10 and r_vals[10] else 0.0
                                    cgst = float(r_vals[11]) if len(r_vals) > 11 and r_vals[11] else 0.0
                                    sgst = float(r_vals[12]) if len(r_vals) > 12 and r_vals[12] else 0.0
                                    gstr1_period = r_vals[20].strip() if len(r_vals) > 20 else ''
                                    filing_date = r_vals[21].strip() if len(r_vals) > 21 else ''
                                else:
                                    inum = r_vals[2].strip()
                                    idate = r_vals[4].strip() if len(r_vals) > 4 else ''
                                    val = float(r_vals[5]) if len(r_vals) > 5 and r_vals[5] else 0.0
                                    rcm_raw = r_vals[7].strip() if len(r_vals) > 7 else ''
                                    txval = float(r_vals[8]) if len(r_vals) > 8 and r_vals[8] else 0.0
                                    igst = float(r_vals[9]) if len(r_vals) > 9 and r_vals[9] else 0.0
                                    cgst = float(r_vals[10]) if len(r_vals) > 10 and r_vals[10] else 0.0
                                    sgst = float(r_vals[11]) if len(r_vals) > 11 and r_vals[11] else 0.0
                                    gstr1_period = r_vals[13].strip() if len(r_vals) > 13 else ''
                                    filing_date = r_vals[14].strip() if len(r_vals) > 14 else ''

                                period = file_period or gstr1_period or '-'
                                is_rcm = 'Y' if rcm_raw.lower() in ['y', 'yes'] else 'N'

                                invoices.append({
                                    'ctin': r_vals[0].strip(),
                                    'cname': r_vals[1].strip(),
                                    'inum': inum,
                                    'norm_inum': normalize_inv(inum),
                                    'dt': idate,
                                    'doc_type': doc_type,
                                    'period_2b': period,
                                    'filing_date': filing_date,
                                    'val': val,
                                    'txval': txval,
                                    'igst': igst,
                                    'cgst': cgst,
                                    'sgst': sgst,
                                    'tax': igst + cgst + sgst,
                                    'rcm': is_rcm
                                })
    except Exception as e:
        print(f"Error parsing XLSX bytes: {e}")
    return invoices

class TallyDatabase:
    def __init__(self):
        self.db = sqlite3.connect(':memory:', check_same_thread=False)
        self.db.row_factory = sqlite3.Row
        self._init_schema()
        self.company_name = "Unknown Company"
        self.file_name = ""
        self.min_date = ""
        self.max_date = ""
        self.reco_results = None

    def _init_schema(self):
        c = self.db.cursor()
        c.execute("DROP TABLE IF EXISTS groups")
        c.execute("DROP TABLE IF EXISTS ledgers")
        c.execute("DROP TABLE IF EXISTS vouchers")
        c.execute("DROP TABLE IF EXISTS ledger_entries")
        
        c.execute('''
            CREATE TABLE groups (
                name TEXT PRIMARY KEY,
                parent TEXT
            )
        ''')
        c.execute('''
            CREATE TABLE ledgers (
                name TEXT PRIMARY KEY,
                parent TEXT,
                opening_balance REAL
            )
        ''')
        c.execute('''
            CREATE TABLE vouchers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT,
                vtype TEXT,
                vno TEXT,
                party TEXT,
                gstin TEXT,
                narration TEXT,
                guid TEXT,
                is_cancelled INTEGER DEFAULT 0
            )
        ''')
        c.execute('''
            CREATE TABLE ledger_entries (
                voucher_id INTEGER,
                ledger TEXT,
                amount REAL,
                is_debit INTEGER
            )
        ''')
        c.execute("CREATE INDEX idx_vouchers_date ON vouchers(date)")
        c.execute("CREATE INDEX idx_vouchers_vtype ON vouchers(vtype)")
        c.execute("CREATE INDEX idx_entries_voucher ON ledger_entries(voucher_id)")
        c.execute("CREATE INDEX idx_entries_ledger ON ledger_entries(ledger)")
        self.db.commit()

    def parse_file(self, file_path_or_bytes, filename="data.xml"):
        self._init_schema()
        self.file_name = filename
        
        xml_content = None
        if isinstance(file_path_or_bytes, str):
            if file_path_or_bytes.endswith('.zip') or zipfile.is_zipfile(file_path_or_bytes):
                with zipfile.ZipFile(file_path_or_bytes, 'r') as z:
                    for fname in z.namelist():
                        if fname.lower().endswith('.xml'):
                            xml_content = z.read(fname)
                            self.file_name = f"{filename} ({fname})"
                            break
            else:
                with open(file_path_or_bytes, 'rb') as f:
                    xml_content = f.read()
        else:
            if filename.lower().endswith('.zip'):
                with zipfile.ZipFile(io.BytesIO(file_path_or_bytes), 'r') as z:
                    for fname in z.namelist():
                        if fname.lower().endswith('.xml'):
                            xml_content = z.read(fname)
                            self.file_name = f"{filename} ({fname})"
                            break
            else:
                xml_content = file_path_or_bytes

        if not xml_content:
            raise ValueError("No XML content found in file.")

        text = ""
        for encoding in ['utf-16', 'utf-16-le', 'utf-8', 'utf-8-sig', 'latin-1']:
            try:
                text = xml_content.decode(encoding)
                if '<ENVELOPE>' in text.upper() or '<TALLYMESSAGE' in text.upper():
                    break
            except Exception:
                continue

        if not text:
            text = xml_content.decode('utf-8', errors='ignore')

        text = re.sub(r'&#\d+;', '', text)
        text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', text)
        text = re.sub(r'&(?!amp;|lt;|gt;|quot;|apos;|#\d+;|#x[0-9a-fA-F]+;)', '&amp;', text)

        root = ET.fromstring(text)
        
        comp_elem = root.find('.//SVCURRENTCOMPANY')
        if comp_elem is not None and comp_elem.text:
            self.company_name = comp_elem.text.strip()
        else:
            comp = root.find('.//COMPANY')
            if comp is not None:
                self.company_name = comp.get('NAME') or comp.findtext('NAME') or "Tally Company"

        c = self.db.cursor()
        
        for msg in root.findall('.//TALLYMESSAGE'):
            grp = msg.find('GROUP')
            if grp is not None:
                gname = grp.get('NAME') or grp.findtext('NAME')
                gparent = grp.findtext('PARENT') or ''
                if gname:
                    c.execute('INSERT OR REPLACE INTO groups VALUES (?, ?)', (gname.strip(), gparent.strip()))
            
            led = msg.find('LEDGER')
            if led is not None:
                lname = led.get('NAME') or led.findtext('NAME')
                lparent = led.findtext('PARENT') or ''
                op_bal_str = led.findtext('OPENINGBALANCE') or '0'
                try: op_bal = float(op_bal_str)
                except ValueError: op_bal = 0.0
                if lname:
                    c.execute('INSERT OR REPLACE INTO ledgers VALUES (?, ?, ?)', (lname.strip(), lparent.strip(), op_bal))
            
            vch = msg.find('VOUCHER')
            if vch is not None:
                is_cancelled = 1 if vch.findtext('ISCANCELLED') == 'Yes' else 0
                vdate = vch.findtext('DATE') or ''
                vtype = vch.findtext('VOUCHERTYPENAME') or vch.get('VCHTYPE') or ''
                vno = vch.findtext('VOUCHERNUMBER') or ''
                party = vch.findtext('PARTYNAME') or vch.findtext('PARTYLEDGERNAME') or ''
                gstin = vch.findtext('PARTYGSTIN') or ''
                narration = vch.findtext('NARRATION') or ''
                guid = vch.findtext('GUID') or vch.get('REMOTEID') or ''
                
                c.execute('''
                    INSERT INTO vouchers (date, vtype, vno, party, gstin, narration, guid, is_cancelled)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ''', (vdate.strip(), vtype.strip(), vno.strip(), party.strip(), gstin.strip(), narration.strip(), guid.strip(), is_cancelled))
                
                v_id = c.lastrowid
                
                entries = vch.findall('ALLLEDGERENTRIES.LIST') or vch.findall('LEDGERENTRIES.LIST')
                for entry in entries:
                    l_name = entry.findtext('LEDGERNAME') or ''
                    if not l_name: continue
                    amt_str = entry.findtext('AMOUNT') or '0'
                    is_pos = entry.findtext('ISDEEMEDPOSITIVE')
                    try: raw_amt = float(amt_str)
                    except ValueError: raw_amt = 0.0
                    
                    is_debit = 1 if is_pos == 'Yes' else 0
                    c.execute('INSERT INTO ledger_entries VALUES (?, ?, ?, ?)',
                              (v_id, l_name.strip(), abs(raw_amt), is_debit))

        self.db.commit()

        row = c.execute("SELECT MIN(date), MAX(date) FROM vouchers WHERE date != ''").fetchone()
        if row and row[0]:
            self.min_date = row[0]
            self.max_date = row[1]

    def get_info(self):
        c = self.db.cursor()
        v_count = c.execute("SELECT COUNT(*) FROM vouchers").fetchone()[0]
        l_count = c.execute("SELECT COUNT(*) FROM ledgers").fetchone()[0]
        g_count = c.execute("SELECT COUNT(*) FROM groups").fetchone()[0]
        return {
            "company_name": self.company_name,
            "file_name": self.file_name,
            "min_date": self.min_date,
            "max_date": self.max_date,
            "voucher_count": v_count,
            "ledger_count": l_count,
            "group_count": g_count
        }

    def get_summary(self):
        c = self.db.cursor()
        info = self.get_info()
        
        totals = c.execute('''
            SELECT 
                SUM(CASE WHEN is_debit = 1 THEN amount ELSE 0 END) as total_debit,
                SUM(CASE WHEN is_debit = 0 THEN amount ELSE 0 END) as total_credit
            FROM ledger_entries le
            JOIN vouchers v ON le.voucher_id = v.id
            WHERE v.is_cancelled = 0
        ''').fetchone()
        
        vtype_rows = c.execute('''
            SELECT vtype, COUNT(*) as cnt
            FROM vouchers
            GROUP BY vtype
            ORDER BY cnt DESC
        ''').fetchall()
        
        monthly_rows = c.execute('''
            SELECT SUBSTR(date, 1, 6) as month, COUNT(DISTINCT v.id) as cnt,
                   SUM(CASE WHEN le.is_debit = 1 THEN le.amount ELSE 0 END) as debit,
                   SUM(CASE WHEN le.is_debit = 0 THEN le.amount ELSE 0 END) as credit
            FROM vouchers v
            JOIN ledger_entries le ON v.id = le.voucher_id
            WHERE v.date != '' AND v.is_cancelled = 0
            GROUP BY month
            ORDER BY month ASC
        ''').fetchall()
        
        top_ledgers = c.execute('''
            SELECT le.ledger, COALESCE(l.parent, 'Other') as parent,
                   SUM(CASE WHEN le.is_debit = 1 THEN le.amount ELSE 0 END) as debit,
                   SUM(CASE WHEN le.is_debit = 0 THEN le.amount ELSE 0 END) as credit
            FROM ledger_entries le
            LEFT JOIN ledgers l ON le.ledger = l.name
            JOIN vouchers v ON le.voucher_id = v.id
            WHERE v.is_cancelled = 0
            GROUP BY le.ledger
            ORDER BY (debit + credit) DESC
            LIMIT 7
        ''').fetchall()

        return {
            "info": info,
            "total_debit": totals["total_debit"] or 0,
            "total_credit": totals["total_credit"] or 0,
            "voucher_types": [{"vtype": r["vtype"], "count": r["cnt"]} for r in vtype_rows],
            "monthly": [{"month": r["month"], "count": r["cnt"], "debit": r["debit"] or 0, "credit": r["credit"] or 0} for r in monthly_rows],
            "top_ledgers": [{"ledger": r["ledger"], "parent": r["parent"], "debit": r["debit"], "credit": r["credit"]} for r in top_ledgers]
        }

    def get_daybook(self, page=1, limit=50, search="", vtype="", party="", date_from="", date_to="", min_amt=None, max_amt=None, sort_by="date", sort_order="DESC"):
        c = self.db.cursor()
        
        conditions = ["1=1"]
        params = []
        
        if search:
            s_term = f"%{search}%"
            conditions.append("(v.narration LIKE ? OR v.party LIKE ? OR v.vno LIKE ? OR le.ledger LIKE ?)")
            params.extend([s_term, s_term, s_term, s_term])
            
        if vtype:
            conditions.append("v.vtype = ?")
            params.append(vtype)
            
        if party:
            conditions.append("v.party LIKE ?")
            params.append(f"%{party}%")
            
        if date_from:
            clean_df = date_from.replace('-', '')
            conditions.append("v.date >= ?")
            params.append(clean_df)
            
        if date_to:
            clean_dt = date_to.replace('-', '')
            conditions.append("v.date <= ?")
            params.append(clean_dt)

        where_clause = " AND ".join(conditions)
        
        count_sql = f'''
            SELECT COUNT(DISTINCT v.id)
            FROM vouchers v
            LEFT JOIN ledger_entries le ON v.id = le.voucher_id
            WHERE {where_clause}
        '''
        total = c.execute(count_sql, params).fetchone()[0]

        allowed_sort = {"date": "v.date", "vtype": "v.vtype", "vno": "v.vno", "party": "v.party", "id": "v.id"}
        sort_col = allowed_sort.get(sort_by, "v.date")
        order = "DESC" if sort_order.upper() == "DESC" else "ASC"
        
        offset = (page - 1) * limit

        query_sql = f'''
            SELECT DISTINCT v.id, v.date, v.vtype, v.vno, v.party, v.gstin, v.narration, v.is_cancelled,
                   (SELECT SUM(amount) FROM ledger_entries WHERE voucher_id = v.id AND is_debit = 1) as total_amount
            FROM vouchers v
            LEFT JOIN ledger_entries le ON v.id = le.voucher_id
            WHERE {where_clause}
            ORDER BY {sort_col} {order}, v.id {order}
            LIMIT ? OFFSET ?
        '''
        params.extend([limit, offset])
        
        v_rows = c.execute(query_sql, params).fetchall()
        
        vouchers = []
        for r in v_rows:
            v_id = r["id"]
            entries = c.execute('''
                SELECT ledger, amount, is_debit
                FROM ledger_entries
                WHERE voucher_id = ?
            ''', (v_id,)).fetchall()
            
            debit_ledgers = [e["ledger"] for e in entries if e["is_debit"] == 1]
            credit_ledgers = [e["ledger"] for e in entries if e["is_debit"] == 0]
            
            vouchers.append({
                "id": v_id,
                "date": r["date"],
                "vtype": r["vtype"],
                "vno": r["vno"],
                "party": r["party"],
                "gstin": r["gstin"],
                "narration": r["narration"],
                "is_cancelled": r["is_cancelled"],
                "total_amount": r["total_amount"] or 0.0,
                "debit_ledgers": debit_ledgers,
                "credit_ledgers": credit_ledgers,
                "entries": [{"ledger": e["ledger"], "amount": e["amount"], "is_debit": e["is_debit"]} for e in entries]
            })

        total_pages = (total + limit - 1) // limit if limit > 0 else 1

        return {
            "total": total,
            "page": page,
            "limit": limit,
            "pages": total_pages,
            "data": vouchers
        }

    def get_voucher_detail(self, v_id):
        c = self.db.cursor()
        v = c.execute("SELECT * FROM vouchers WHERE id = ?", (v_id,)).fetchone()
        if not v:
            return None
        entries = c.execute("SELECT * FROM ledger_entries WHERE voucher_id = ?", (v_id,)).fetchall()
        return {
            "voucher": dict(v),
            "entries": [dict(e) for e in entries]
        }

    def update_voucher(self, v_id, date, vtype, vno, party, narration, entries):
        c = self.db.cursor()
        c.execute('''
            UPDATE vouchers
            SET date = ?, vtype = ?, vno = ?, party = ?, narration = ?
            WHERE id = ?
        ''', (date, vtype, vno, party, narration, v_id))
        
        c.execute("DELETE FROM ledger_entries WHERE voucher_id = ?", (v_id,))
        for e in entries:
            l_name = e.get('ledger', '').strip()
            amt = float(e.get('amount', 0))
            is_debit = 1 if e.get('is_debit') else 0
            if l_name:
                c.execute("INSERT INTO ledger_entries VALUES (?, ?, ?, ?)", (v_id, l_name, amt, is_debit))
        self.db.commit()
        return self.get_voucher_detail(v_id)

    def create_voucher(self, date, vtype, vno, party, narration, entries):
        c = self.db.cursor()
        c.execute('''
            INSERT INTO vouchers (date, vtype, vno, party, narration)
            VALUES (?, ?, ?, ?, ?)
        ''', (date, vtype, vno, party, narration))
        v_id = c.lastrowid
        for e in entries:
            l_name = e.get('ledger', '').strip()
            amt = float(e.get('amount', 0))
            is_debit = 1 if e.get('is_debit') else 0
            if l_name:
                c.execute("INSERT INTO ledger_entries VALUES (?, ?, ?, ?)", (v_id, l_name, amt, is_debit))
        self.db.commit()
        return self.get_voucher_detail(v_id)

    def get_trial_balance(self):
        c = self.db.cursor()
        g_rows = c.execute("SELECT name, parent FROM groups").fetchall()
        groups = {r["name"]: r["parent"] or "Primary" for r in g_rows}
        
        query = '''
            SELECT l.name, COALESCE(l.parent, 'Primary') as parent, l.opening_balance,
                   COALESCE(SUM(CASE WHEN le.is_debit = 1 THEN le.amount ELSE 0 END), 0) as debit,
                   COALESCE(SUM(CASE WHEN le.is_debit = 0 THEN le.amount ELSE 0 END), 0) as credit
            FROM ledgers l
            LEFT JOIN ledger_entries le ON l.name = le.ledger
            LEFT JOIN vouchers v ON le.voucher_id = v.id AND v.is_cancelled = 0
            GROUP BY l.name
        '''
        l_rows = c.execute(query).fetchall()
        
        ledgers = []
        for r in l_rows:
            op = r["opening_balance"] or 0.0
            dr = r["debit"] or 0.0
            cr = r["credit"] or 0.0
            closing = op + dr - cr
            ledgers.append({
                "name": r["name"],
                "parent": r["parent"],
                "opening": op,
                "debit": dr,
                "credit": cr,
                "closing": closing
            })

        return {
            "groups": groups,
            "ledgers": ledgers
        }

    def get_ledgers_list(self):
        c = self.db.cursor()
        rows = c.execute('''
            SELECT 
                all_l.name,
                COALESCE(l.parent, 'Primary') as parent,
                COALESCE(l.opening_balance, 0.0) as opening_balance,
                COALESCE(SUM(CASE WHEN le.is_debit = 1 THEN le.amount ELSE 0 END), 0) as debit,
                COALESCE(SUM(CASE WHEN le.is_debit = 0 THEN le.amount ELSE 0 END), 0) as credit
            FROM (
                SELECT name FROM ledgers
                UNION
                SELECT DISTINCT ledger as name FROM ledger_entries WHERE ledger IS NOT NULL AND TRIM(ledger) != ''
            ) all_l
            LEFT JOIN ledgers l ON all_l.name = l.name
            LEFT JOIN ledger_entries le ON all_l.name = le.ledger
            GROUP BY all_l.name
            ORDER BY all_l.name ASC
        ''').fetchall()
        
        res = []
        for r in rows:
            op = r["opening_balance"] or 0.0
            dr = r["debit"] or 0.0
            cr = r["credit"] or 0.0
            res.append({
                "name": r["name"],
                "parent": r["parent"] or "Primary",
                "opening": op,
                "debit": dr,
                "credit": cr,
                "closing": op + dr - cr
            })
        return res

    def get_ledger_statement(self, ledger_name, date_from="", date_to=""):
        c = self.db.cursor()
        
        l_info = c.execute("SELECT * FROM ledgers WHERE name = ?", (ledger_name,)).fetchone()
        op_bal = l_info["opening_balance"] if l_info else 0.0
        parent = l_info["parent"] if l_info else "Primary"

        conditions = ["le.ledger = ?", "v.is_cancelled = 0"]
        params = [ledger_name]
        
        if date_from:
            conditions.append("v.date >= ?")
            params.append(date_from.replace('-', ''))
        if date_to:
            conditions.append("v.date <= ?")
            params.append(date_to.replace('-', ''))

        where_clause = " AND ".join(conditions)

        sql = f'''
            SELECT v.id, v.date, v.vtype, v.vno, v.party, v.narration, le.amount, le.is_debit
            FROM ledger_entries le
            JOIN vouchers v ON le.voucher_id = v.id
            WHERE {where_clause}
            ORDER BY v.date ASC, v.id ASC
        '''
        
        entries = c.execute(sql, params).fetchall()

        statement = []
        running_bal = op_bal
        total_dr = 0.0
        total_cr = 0.0

        for r in entries:
            amt = r["amount"]
            is_debit = r["is_debit"]
            if is_debit == 1:
                total_dr += amt
                running_bal += amt
            else:
                total_cr += amt
                running_bal -= amt
                
            statement.append({
                "voucher_id": r["id"],
                "date": r["date"],
                "vtype": r["vtype"],
                "vno": r["vno"],
                "party": r["party"],
                "narration": r["narration"],
                "amount": amt,
                "is_debit": is_debit,
                "running_balance": running_bal
            })

        return {
            "ledger_name": ledger_name,
            "parent": parent,
            "opening_balance": op_bal,
            "total_debit": total_dr,
            "total_credit": total_cr,
            "closing_balance": running_bal,
            "entries": statement
        }

    def reconcile_gstr2b(self, gstr2b_invoices):
        c = self.db.cursor()
        
        # 1. Fetch Tally Purchase, Debit Note, Credit Note, and GST-related Journal Vouchers
        query = '''
            SELECT DISTINCT v.id, v.date, v.vno, v.vtype, v.party, v.gstin
            FROM vouchers v
            JOIN ledger_entries le ON v.id = le.voucher_id
            WHERE (
                (v.vtype IN ('Purchase', 'Debit Note', 'Credit Note'))
                OR
                (v.vtype = 'Journal' AND (
                    (v.gstin IS NOT NULL AND TRIM(v.gstin) != '')
                    OR EXISTS (
                        SELECT 1 FROM ledger_entries le2
                        WHERE le2.voucher_id = v.id
                          AND (LOWER(le2.ledger) LIKE '%cgst%' OR LOWER(le2.ledger) LIKE '%sgst%' OR LOWER(le2.ledger) LIKE '%igst%' OR LOWER(le2.ledger) LIKE '%utgst%')
                    )
                ))
            ) AND v.is_cancelled = 0
        '''
        tally_vouchers = c.execute(query).fetchall()
        
        tally_map = {}
        for tv in tally_vouchers:
            v_id = tv["id"]
            entries = c.execute("SELECT ledger, amount, is_debit FROM ledger_entries WHERE voucher_id = ?", (v_id,)).fetchall()
            
            cgst, sgst, igst, other_tax, total_val = 0.0, 0.0, 0.0, 0.0, 0.0
            has_cgst, has_sgst = False, False
            
            for e in entries:
                is_tax, t_type = is_tax_ledger(e["ledger"])
                if is_tax:
                    if t_type == 'igst':
                        igst += e["amount"]
                    elif t_type == 'cgst':
                        cgst += e["amount"]
                        has_cgst = True
                    elif t_type == 'sgst':
                        sgst += e["amount"]
                        has_sgst = True
                    elif t_type == 'tax':
                        other_tax += e["amount"]
                elif e["is_debit"] == 1:
                    total_val += e["amount"]
            
            total_tax = round(cgst + sgst + igst + other_tax, 2)
            is_asymmetric_gst = (has_cgst and not has_sgst) or (has_sgst and not has_cgst)
            
            tally_map[str(v_id)] = {
                "id": str(v_id),
                "date": tv["date"],
                "vno": tv["vno"],
                "vtype": tv["vtype"],
                "norm_vno": normalize_inv(tv["vno"]),
                "party": tv["party"],
                "gstin": (tv["gstin"] or "").strip(),
                "total_tax": total_tax,
                "total_val": round(total_val, 2),
                "has_cgst": has_cgst,
                "has_sgst": has_sgst,
                "is_asymmetric_gst": is_asymmetric_gst,
                "matched": False
            }

        # Check for split vouchers to combine when they match a 2B invoice (e.g. Sri Mahalakshmi CB/0034/25-26)
        vno_groups = {}
        for t_id, tv in tally_map.items():
            key = (tv["norm_vno"], tv["date"], tv["gstin"] or tv["party"].lower())
            if key not in vno_groups: vno_groups[key] = []
            vno_groups[key].append(t_id)

        merged_tally_map = {}
        skip_t_ids = set()

        for key, t_ids in vno_groups.items():
            if len(t_ids) > 1:
                tot_tax = sum(tally_map[tid]["total_tax"] for tid in t_ids)
                tot_val = sum(tally_map[tid]["total_val"] for tid in t_ids)
                first_tv = tally_map[t_ids[0]]
                
                matching_2b = False
                for inv in gstr2b_invoices:
                    c_norm = inv.get("norm_inum", "") or normalize_inv(inv.get("inum", ""))
                    if c_norm == first_tv["norm_vno"]:
                        c_tax = inv.get("tax", 0.0)
                        c_val = inv.get("val", 0.0)
                        if abs(tot_tax - c_tax) <= 2.0 or abs(tot_val - c_val) <= 2.0:
                            matching_2b = True
                            break
                
                if matching_2b:
                    comb_id = ",".join(str(tid) for tid in t_ids)
                    merged_tally_map[comb_id] = {
                        "id": comb_id,
                        "date": first_tv["date"],
                        "vno": first_tv["vno"],
                        "vtype": first_tv["vtype"],
                        "norm_vno": first_tv["norm_vno"],
                        "party": first_tv["party"],
                        "gstin": first_tv["gstin"],
                        "total_tax": tot_tax,
                        "total_val": tot_val,
                        "matched": False
                    }
                    skip_t_ids.update(t_ids)

        final_tally_map = {}
        for comb_id, m_tv in merged_tally_map.items():
            final_tally_map[comb_id] = m_tv

        for t_id, tv in tally_map.items():
            if t_id not in skip_t_ids:
                final_tally_map[t_id] = tv

        tally_map = final_tally_map

        reco_map = {}
        matched_exact = 0
        matched_amount = 0
        mismatched_count = 0

        # PASS 1: Exact Invoice Number AND Amount Match (+/- 2.0)
        for idx, inv in enumerate(gstr2b_invoices):
            c_norm_vno = inv.get("norm_inum", "") or normalize_inv(inv.get("inum", ""))
            c_tax = inv.get("tax", 0.0)
            c_val = inv.get("val", 0.0)
            c_gstin = inv.get("ctin", "") or inv.get("gstin", "")

            if not c_norm_vno: continue

            for t_id, tv in final_tally_map.items():
                if tv["matched"]: continue
                if tv["norm_vno"] == c_norm_vno:
                    tax_diff = abs(tv["total_tax"] - c_tax)
                    val_diff = abs(tv["total_val"] - c_val)
                    if tax_diff <= 2.0 or val_diff <= 2.0:
                        final_tally_map[t_id]["matched"] = True
                        matched_exact += 1
                        reco_map[idx] = {
                            "status": "MATCHED",
                            "remarks": "Exact Match (Inv No & Amount)",
                            "gstin": c_gstin or tv["gstin"],
                            "supplier": inv.get("cname", "") or inv.get("name", "") or tv["party"],
                            "inv_no_2b": inv.get("inum", ""),
                            "inv_no_tally": tv["vno"],
                            "vtype_tally": tv["vtype"],
                            "date_2b": inv.get("dt", "") or inv.get("idate", ""),
                            "date_tally": tv["date"],
                            "period_2b": inv.get("period_2b", "-"),
                            "filing_date_2b": inv.get("filing_date", "-"),
                            "tax_2b": c_tax,
                            "tax_tally": tv["total_tax"],
                            "val_2b": c_val,
                            "val_tally": tv["total_val"],
                            "rcm": inv.get("rcm", "N"),
                            "tally_voucher_id": tv["id"]
                        }
                        break

        # PASS 2: GSTIN + Amount Match (+/- 2.0) for remaining items -> Flagged as SHIFTED_BILL / Bill No Mismatch!
        for idx, inv in enumerate(gstr2b_invoices):
            if idx in reco_map: continue
            c_gstin = inv.get("ctin", "") or inv.get("gstin", "")
            c_tax = inv.get("tax", 0.0)
            c_val = inv.get("val", 0.0)

            if c_gstin:
                for t_id, tv in final_tally_map.items():
                    if tv["matched"]: continue
                    if tv["gstin"] == c_gstin and (abs(tv["total_val"] - c_val) <= 2.0 or abs(tv["total_tax"] - c_tax) <= 2.0):
                        final_tally_map[t_id]["matched"] = True
                        matched_amount += 1
                        inv_2b_no = inv.get("inum", "")
                        inv_tally_no = tv["vno"]
                        remark_str = f"Amount Matched, but Bill No differs: Tally '{inv_tally_no}' vs 2B '{inv_2b_no}'" if inv_2b_no != inv_tally_no else "Amount Matched"
                        reco_map[idx] = {
                            "status": "SHIFTED_BILL",
                            "remarks": remark_str,
                            "gstin": c_gstin,
                            "supplier": inv.get("cname", "") or inv.get("name", "") or tv["party"],
                            "inv_no_2b": inv_2b_no,
                            "inv_no_tally": inv_tally_no,
                            "vtype_tally": tv["vtype"],
                            "date_2b": inv.get("dt", "") or inv.get("idate", ""),
                            "date_tally": tv["date"],
                            "period_2b": inv.get("period_2b", "-"),
                            "filing_date_2b": inv.get("filing_date", "-"),
                            "tax_2b": c_tax,
                            "tax_tally": tv["total_tax"],
                            "val_2b": c_val,
                            "val_tally": tv["total_val"],
                            "rcm": inv.get("rcm", "N"),
                            "tally_voucher_id": tv["id"]
                        }
                        break

        # PASS 3: Exact Invoice Number Match with Mismatched Amount (for remaining items)
        for idx, inv in enumerate(gstr2b_invoices):
            if idx in reco_map: continue
            c_norm_vno = inv.get("norm_inum", "") or normalize_inv(inv.get("inum", ""))
            c_tax = inv.get("tax", 0.0)
            c_val = inv.get("val", 0.0)
            c_gstin = inv.get("ctin", "") or inv.get("gstin", "")

            if not c_norm_vno: continue

            for t_id, tv in final_tally_map.items():
                if tv["matched"]: continue
                if tv["norm_vno"] == c_norm_vno:
                    final_tally_map[t_id]["matched"] = True
                    mismatched_count += 1
                    tax_diff = c_tax - tv["total_tax"]
                    reco_map[idx] = {
                        "status": "MISMATCHED_AMOUNT",
                        "remarks": f"Inv No matches, but Tax Diff = ₹{tax_diff:.2f}",
                        "gstin": c_gstin or tv["gstin"],
                        "supplier": inv.get("cname", "") or inv.get("name", "") or tv["party"],
                        "inv_no_2b": inv.get("inum", ""),
                        "inv_no_tally": tv["vno"],
                        "vtype_tally": tv["vtype"],
                        "date_2b": inv.get("dt", "") or inv.get("idate", ""),
                        "date_tally": tv["date"],
                        "period_2b": inv.get("period_2b", "-"),
                        "filing_date_2b": inv.get("filing_date", "-"),
                        "tax_2b": c_tax,
                        "tax_tally": tv["total_tax"],
                        "val_2b": c_val,
                        "val_tally": tv["total_val"],
                        "rcm": inv.get("rcm", "N"),
                        "tally_voucher_id": tv["id"]
                    }
                    break

        # PASS 4: Unmatched GSTR-2B Items
        missing_tally_count = 0
        for idx, inv in enumerate(gstr2b_invoices):
            if idx not in reco_map:
                missing_tally_count += 1
                c_gstin = inv.get("ctin", "") or inv.get("gstin", "")
                reco_map[idx] = {
                    "status": "MISSING_IN_TALLY",
                    "remarks": "In GSTR-2B, missing in Tally",
                    "gstin": c_gstin,
                    "supplier": inv.get("cname", "") or inv.get("name", ""),
                    "inv_no_2b": inv.get("inum", ""),
                    "inv_no_tally": "-",
                    "vtype_tally": "-",
                    "date_2b": inv.get("dt", "") or inv.get("idate", ""),
                    "date_tally": "-",
                    "period_2b": inv.get("period_2b", "-"),
                    "filing_date_2b": inv.get("filing_date", "-"),
                    "tax_2b": inv.get("tax", 0.0),
                    "tax_tally": 0.0,
                    "val_2b": inv.get("val", 0.0),
                    "val_tally": 0.0,
                    "rcm": inv.get("rcm", "N"),
                    "tally_voucher_id": None
                }

        # Build lookup for previous notes & isRectified flags if self.reco_results exists
        prev_state = {}
        if self.reco_results and "records" in self.reco_results:
            for pr in self.reco_results["records"]:
                key = (pr.get("gstin", ""), pr.get("inv_no_2b", ""), pr.get("inv_no_tally", ""))
                prev_state[key] = {
                    "note": pr.get("note", ""),
                    "isRectified": pr.get("isRectified", False)
                }

        reco_list = [reco_map[i] for i in sorted(reco_map.keys())]

        # PASS 5: Unmatched Tally Vouchers
        missing_2b_count = 0
        for t_id, tv in final_tally_map.items():
            if not tv["matched"]:
                missing_2b_count += 1
                reco_list.append({
                    "status": "MISSING_IN_2B",
                    "remarks": "In Tally, missing in GSTR-2B",
                    "gstin": tv["gstin"],
                    "supplier": tv["party"],
                    "inv_no_2b": "-",
                    "inv_no_tally": tv["vno"],
                    "vtype_tally": tv["vtype"],
                    "date_2b": "-",
                    "date_tally": tv["date"],
                    "period_2b": "-",
                    "filing_date_2b": "-",
                    "tax_2b": 0.0,
                    "tax_tally": tv["total_tax"],
                    "val_2b": 0.0,
                    "val_tally": tv["total_val"],
                    "rcm": "N",
                    "tally_voucher_id": tv["id"]
                })

        for idx, r in enumerate(reco_list):
            r["id"] = f"rec_{idx}"
            key = (r.get("gstin", ""), r.get("inv_no_2b", ""), r.get("inv_no_tally", ""))
            if key in prev_state:
                r["note"] = prev_state[key]["note"]
                r["isRectified"] = prev_state[key]["isRectified"]
            else:
                r["note"] = r.get("note", "")
                r["isRectified"] = r.get("isRectified", False)

        rectified_cnt = sum(1 for r in reco_list if r.get("isRectified"))
        rcm_cnt = sum(1 for r in reco_list if r.get("rcm") == "Y")

        summary = {
            "total_2b": len(gstr2b_invoices),
            "matched_count": matched_exact + matched_amount,
            "matched_exact": matched_exact,
            "matched_amount": matched_amount,
            "mismatched_count": mismatched_count,
            "missing_tally_count": missing_tally_count,
            "missing_2b_count": missing_2b_count,
            "rectified_count": rectified_cnt,
            "rcm_count": rcm_cnt
        }

        self.reco_results = {
            "summary": summary,
            "records": reco_list
        }
        self.recalculate_reco_summary()
        return self.reco_results

    def update_reco_record(self, record_id, note=None, is_rectified=None):
        if not self.reco_results or 'records' not in self.reco_results:
            return False
        for r in self.reco_results['records']:
            if r.get('id') == record_id:
                if note is not None:
                    r['note'] = note
                if is_rectified is not None:
                    r['isRectified'] = bool(is_rectified)
                self.recalculate_reco_summary()
                return True
        return False

    def recalculate_reco_summary(self):
        if not self.reco_results or 'records' not in self.reco_results:
            return
        records = self.reco_results['records']
        rectified_cnt = sum(1 for r in records if r.get('isRectified'))
        
        mismatched_cnt = sum(1 for r in records if r.get('status') == 'MISMATCHED_AMOUNT' and not r.get('isRectified'))
        missing_tally_cnt = sum(1 for r in records if r.get('status') == 'MISSING_IN_TALLY' and not r.get('isRectified'))
        missing_2b_cnt = sum(1 for r in records if r.get('status') == 'MISSING_IN_2B' and not r.get('isRectified'))
        matched_exact = sum(1 for r in records if r.get('status') == 'MATCHED' and not r.get('isRectified'))
        matched_amount = sum(1 for r in records if r.get('status') == 'SHIFTED_BILL' and not r.get('isRectified'))
        
        self.reco_results['summary']['rectified_count'] = rectified_cnt
        self.reco_results['summary']['mismatched_count'] = mismatched_cnt
        self.reco_results['summary']['missing_tally_count'] = missing_tally_cnt
        self.reco_results['summary']['missing_2b_count'] = missing_2b_cnt
        self.reco_results['summary']['matched_count'] = matched_exact + matched_amount
        self.reco_results['summary']['matched_exact'] = matched_exact
        self.reco_results['summary']['matched_amount'] = matched_amount

    def export_session(self):
        c = self.db.cursor()
        groups = [dict(r) for r in c.execute("SELECT name, parent FROM groups").fetchall()]
        ledgers = [dict(r) for r in c.execute("SELECT name, parent, opening_balance FROM ledgers").fetchall()]
        vouchers = [dict(r) for r in c.execute("SELECT id, date, vtype, vno, party, gstin, narration, guid, is_cancelled FROM vouchers").fetchall()]
        entries = [dict(r) for r in c.execute("SELECT voucher_id, ledger, amount, is_debit FROM ledger_entries").fetchall()]
        
        return {
            "version": "1.0",
            "file_type": "rex_session",
            "company_name": self.company_name,
            "file_name": self.file_name,
            "min_date": self.min_date,
            "max_date": self.max_date,
            "info": self.get_info(),
            "tally_data": {
                "groups": groups,
                "ledgers": ledgers,
                "vouchers": vouchers,
                "ledger_entries": entries
            },
            "reco_results": self.reco_results
        }

    def import_session(self, session_data):
        if not isinstance(session_data, dict):
            raise ValueError("Invalid session payload")
        
        self._init_schema()
        self.company_name = session_data.get("company_name") or session_data.get("info", {}).get("company_name", "Loaded Session")
        self.file_name = session_data.get("file_name") or session_data.get("info", {}).get("file_name", "reconciliation_progress.rex")
        self.min_date = session_data.get("min_date") or session_data.get("info", {}).get("min_date", "")
        self.max_date = session_data.get("max_date") or session_data.get("info", {}).get("max_date", "")
        
        tally_data = session_data.get("tally_data", {})
        c = self.db.cursor()
        
        for g in tally_data.get("groups", []):
            c.execute("INSERT OR REPLACE INTO groups VALUES (?, ?)", (g["name"], g.get("parent", "")))
            
        for l in tally_data.get("ledgers", []):
            c.execute("INSERT OR REPLACE INTO ledgers VALUES (?, ?, ?)", (l["name"], l.get("parent", ""), l.get("opening_balance", 0.0)))
            
        for v in tally_data.get("vouchers", []):
            c.execute('''
                INSERT OR REPLACE INTO vouchers (id, date, vtype, vno, party, gstin, narration, guid, is_cancelled)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (v.get("id"), v.get("date", ""), v.get("vtype", ""), v.get("vno", ""), v.get("party", ""), v.get("gstin", ""), v.get("narration", ""), v.get("guid", ""), v.get("is_cancelled", 0)))
            
        for e in tally_data.get("ledger_entries", []):
            c.execute("INSERT INTO ledger_entries VALUES (?, ?, ?, ?)", (e["voucher_id"], e["ledger"], e["amount"], e["is_debit"]))
            
        self.db.commit()
        
        row = c.execute("SELECT MIN(date), MAX(date) FROM vouchers WHERE date != ''").fetchone()
        if row and row[0]:
            self.min_date = row[0]
            self.max_date = row[1]

        reco = session_data.get("reco_results")
        if reco and isinstance(reco, dict) and "records" in reco:
            self.reco_results = reco
            self.recalculate_reco_summary()
            
        return self.get_info()

    def export_tally_xml(self):
        c = self.db.cursor()
        root = ET.Element('ENVELOPE')
        header = ET.SubElement(root, 'HEADER')
        ET.SubElement(header, 'TALLYREQUEST').text = 'Import Data'
        body = ET.SubElement(root, 'BODY')
        imp = ET.SubElement(body, 'IMPORTDATA')
        reqdesc = ET.SubElement(imp, 'REQUESTDESC')
        ET.SubElement(reqdesc, 'REPORTNAME').text = 'All Masters'
        statvar = ET.SubElement(reqdesc, 'STATICVARIABLES')
        ET.SubElement(statvar, 'SVCURRENTCOMPANY').text = self.company_name

        reqdata = ET.SubElement(imp, 'REQUESTDATA')

        for gname, gparent in c.execute("SELECT name, parent FROM groups"):
            tm = ET.SubElement(reqdata, 'TALLYMESSAGE', {'xmlns:UDF': 'TallyUDF'})
            grp = ET.SubElement(tm, 'GROUP', {'NAME': gname, 'ACTION': 'Create'})
            ET.SubElement(grp, 'NAME').text = gname
            if gparent: ET.SubElement(grp, 'PARENT').text = gparent

        for lname, lparent, op_bal in c.execute("SELECT name, parent, opening_balance FROM ledgers"):
            tm = ET.SubElement(reqdata, 'TALLYMESSAGE', {'xmlns:UDF': 'TallyUDF'})
            led = ET.SubElement(tm, 'LEDGER', {'NAME': lname, 'ACTION': 'Create'})
            ET.SubElement(led, 'NAME').text = lname
            if lparent: ET.SubElement(led, 'PARENT').text = lparent
            if op_bal: ET.SubElement(led, 'OPENINGBALANCE').text = str(op_bal)

        for v in c.execute("SELECT * FROM vouchers"):
            v_id = v["id"]
            guid_val = (v["guid"] or "").strip()
            
            tm = ET.SubElement(reqdata, 'TALLYMESSAGE', {'xmlns:UDF': 'TallyUDF'})
            vch_attrs = {'VCHTYPE': v["vtype"]}
            if guid_val:
                vch_attrs['ACTION'] = 'Alter'
                vch_attrs['REMOTEID'] = guid_val
            else:
                vch_attrs['ACTION'] = 'Create'
                
            vch = ET.SubElement(tm, 'VOUCHER', vch_attrs)
            if guid_val:
                ET.SubElement(vch, 'GUID').text = guid_val
                
            ET.SubElement(vch, 'DATE').text = v["date"]
            ET.SubElement(vch, 'VOUCHERTYPENAME').text = v["vtype"]
            ET.SubElement(vch, 'VOUCHERNUMBER').text = v["vno"]
            ET.SubElement(vch, 'PARTYNAME').text = v["party"]
            if v["gstin"]: ET.SubElement(vch, 'PARTYGSTIN').text = v["gstin"]
            ET.SubElement(vch, 'NARRATION').text = v["narration"]
            if v["is_cancelled"]: ET.SubElement(vch, 'ISCANCELLED').text = 'Yes'

            entries = c.execute("SELECT * FROM ledger_entries WHERE voucher_id = ?", (v_id,)).fetchall()
            for e in entries:
                le = ET.SubElement(vch, 'ALLLEDGERENTRIES.LIST')
                ET.SubElement(le, 'LEDGERNAME').text = e["ledger"]
                ET.SubElement(le, 'ISDEEMEDPOSITIVE').text = 'Yes' if e["is_debit"] else 'No'
                amt_val = f'-{e["amount"]:.2f}' if e["is_debit"] else f'{e["amount"]:.2f}'
                ET.SubElement(le, 'AMOUNT').text = amt_val

        return ET.tostring(root, encoding='utf-8', xml_declaration=True).decode('utf-8')

    def export_csv(self, report_type, **kwargs):
        output = io.StringIO()
        if report_type == 'daybook':
            db_data = self.get_daybook(limit=100000, **kwargs)
            output.write("ID,Date,Voucher Type,Voucher No,Party Name,Debit Ledgers,Credit Ledgers,Amount,Narration\n")
            for v in db_data["data"]:
                dr_str = "; ".join(v["debit_ledgers"]).replace('"', '""')
                cr_str = "; ".join(v["credit_ledgers"]).replace('"', '""')
                narr = (v["narration"] or "").replace('"', '""').replace('\n', ' ')
                party = (v["party"] or "").replace('"', '""')
                output.write(f'"{v["id"]}","{v["date"]}","{v["vtype"]}","{v["vno"]}","{party}","{dr_str}","{cr_str}","{v["total_amount"]}","{narr}"\n')
                
        elif report_type == 'ledger':
            l_name = kwargs.get('ledger_name', '')
            stmt = self.get_ledger_statement(l_name)
            output.write(f"Statement of Account: {l_name} (Parent: {stmt['parent']})\n")
            output.write(f"Opening Balance: {stmt['opening_balance']}\n\n")
            output.write("Date,Voucher Type,Voucher No,Particulars,Debit,Credit,Running Balance,Narration\n")
            for e in stmt["entries"]:
                dr = e["amount"] if e["is_debit"] == 1 else 0
                cr = e["amount"] if e["is_debit"] == 0 else 0
                party = (e["party"] or "").replace('"', '""')
                narr = (e["narration"] or "").replace('"', '""').replace('\n', ' ')
                output.write(f'"{e["date"]}","{e["vtype"]}","{e["vno"]}","{party}","{dr}","{cr}","{e["running_balance"]}","{narr}"\n')

        elif report_type == 'trial_balance':
            tb = self.get_trial_balance()
            output.write("Ledger Name,Parent Group,Opening Balance,Total Debit,Total Credit,Closing Balance\n")
            for l in tb["ledgers"]:
                output.write(f'"{l["name"].replace('"', '""')}","{l["parent"].replace('"', '""')}","{l["opening"]}","{l["debit"]}","{l["credit"]}","{l["closing"]}"\n')

        elif report_type == 'gstr2b' and self.reco_results:
            output.write("Status,Is Rectified,Supplier GSTIN,Supplier Name,Invoice No (2B),Voucher Type (Tally),Invoice No (Tally),Date (2B),Date (Tally),2B Return Period,GSTR-1 Filing Date,Tax Amount (2B),Tax Amount (Tally),Tax Diff (2B - Tally),Total Value (2B),Total Value (Tally),Audit Remarks,Notes\n")
            for r in self.reco_results["records"]:
                sup = (r["supplier"] or "").replace('"', '""')
                rem = (r.get("remarks", "")).replace('"', '""')
                note = (r.get("note", "")).replace('"', '""').replace('\n', ' ')
                is_rect = "YES" if r.get("isRectified") else "NO"
                tax_diff = (r["tax_2b"] or 0.0) - (r["tax_tally"] or 0.0)
                dt_tally = r.get("date_tally", "-")
                if dt_tally and dt_tally != "-" and len(dt_tally) == 8:
                    dt_tally = f"{dt_tally[6:8]}/{dt_tally[4:6]}/{dt_tally[0:4]}"
                vt_tally = r.get("vtype_tally", "-")
                output.write(f'"{r["status"]}","{is_rect}","{r["gstin"]}","{sup}","{r["inv_no_2b"]}","{vt_tally}","{r["inv_no_tally"]}","{r["date_2b"]}","{dt_tally}","{r.get("period_2b", "-")}","{r.get("filing_date_2b", "-")}","{r["tax_2b"]}","{r["tax_tally"]}","{tax_diff:.2f}","{r["val_2b"]}","{r["val_tally"]}","{rem}","{note}"\n')

        return output.getvalue()

def parse_gstr2b_json_bytes(file_bytes, filename=""):
    invoices = []
    try:
        gstr2b_json = json.loads(file_bytes.decode('utf-8', errors='ignore'))
        period_2b = "-"
        if filename:
            m = re.search(r'(\d{2})(\d{4})', filename)
            if m:
                month_num, year = m.group(1), m.group(2)
                months_map = {'01':'Jan', '02':'Feb', '03':'Mar', '04':'Apr', '05':'May', '06':'Jun', '07':'Jul', '08':'Aug', '09':'Sep', '10':'Oct', '11':'Nov', '12':'Dec'}
                period_2b = f"{months_map.get(month_num, month_num)}'{year[2:]}"
        
        if isinstance(gstr2b_json, dict):
            b2b_list = gstr2b_json.get('b2b', []) or gstr2b_json.get('data', {}).get('b2b', [])
            for supplier in b2b_list:
                ctin = supplier.get('ctin', '').strip()
                cname = supplier.get('cname', '').strip()
                for inv in supplier.get('inv', []):
                    inum = str(inv.get('inum', '')).strip()
                    idt = inv.get('dt', '') or inv.get('idt', '')
                    val = float(inv.get('val', 0))
                    cgst, sgst, igst, txval = 0.0, 0.0, 0.0, 0.0
                    for item in inv.get('items', []):
                        det = item.get('itm_det', {})
                        txval += float(det.get('txval', 0))
                        cgst += float(det.get('cgst', 0))
                        sgst += float(det.get('sgst', 0))
                        igst += float(det.get('igst', 0))
                    rcm_raw = str(inv.get('rev', '') or inv.get('reverse_charge', '') or inv.get('rcm', '')).strip()
                    is_rcm = 'Y' if rcm_raw.lower() in ['y', 'yes', 'true'] else 'N'
                    invoices.append({
                        "ctin": ctin, "cname": cname, "inum": inum,
                        "norm_inum": normalize_inv(inum), "dt": idt,
                        "val": val, "txval": txval, "cgst": cgst, "sgst": sgst, "igst": igst,
                        "tax": cgst + sgst + igst, "period_2b": period_2b, "filing_date": "-",
                        "rcm": is_rcm
                    })
    except Exception as e:
        print("JSON decode error:", e)
    return invoices


tally_db = TallyDatabase()

def auto_load_r2b_files():
    r2b_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'R2B')
    if not os.path.exists(r2b_dir):
        return []
    all_files = [os.path.join(r2b_dir, f) for f in os.listdir(r2b_dir) if f.lower().endswith('.xlsx') or f.lower().endswith('.json')]
    if not all_files:
        return []
    
    all_invoices = []
    for fpath in sorted(all_files):
        try:
            with open(fpath, 'rb') as f:
                content = f.read()
                if fpath.lower().endswith('.xlsx'):
                    invs = parse_gstr2b_xlsx_bytes(content, filename=os.path.basename(fpath))
                else:
                    invs = parse_gstr2b_json_bytes(content, filename=os.path.basename(fpath))
                all_invoices.extend(invs)
        except Exception as e:
            print(f"Error reading {fpath}: {e}")
    if all_invoices:
        print(f"Auto-reconciling {len(all_invoices)} GSTR-2B invoices from {len(all_files)} files...")
        tally_db.reconcile_gstr2b(all_invoices)
    return all_invoices

class RequestHandler(BaseHTTPRequestHandler):
    def send_json(self, data, status=200):
        body = json.dumps(data).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        params = urllib.parse.parse_qs(parsed.query)

        if path == '/' or path == '/index.html':
            html_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'index.html')
            if os.path.exists(html_file):
                with open(html_file, 'rb') as f:
                    content = f.read()
                self.send_response(200)
                self.send_header('Content-Type', 'text/html; charset=utf-8')
                self.send_header('Content-Length', str(len(content)))
                self.end_headers()
                self.wfile.write(content)
            else:
                self.send_json({"error": "index.html not found"}, 404)

        elif path == '/api/info':
            self.send_json(tally_db.get_info())

        elif path == '/api/summary':
            self.send_json(tally_db.get_summary())

        elif path == '/api/daybook':
            page = int(params.get('page', [1])[0])
            limit = int(params.get('limit', [50])[0])
            search = params.get('search', [''])[0]
            vtype = params.get('vtype', [''])[0]
            party = params.get('party', [''])[0]
            date_from = params.get('date_from', [''])[0]
            date_to = params.get('date_to', [''])[0]
            sort_by = params.get('sort_by', ['date'])[0]
            sort_order = params.get('sort_order', ['DESC'])[0]
            
            data = tally_db.get_daybook(
                page=page, limit=limit, search=search, vtype=vtype,
                party=party, date_from=date_from, date_to=date_to,
                sort_by=sort_by, sort_order=sort_order
            )
            self.send_json(data)

        elif path.startswith('/api/voucher/'):
            try:
                v_id = int(path.split('/')[-1])
                detail = tally_db.get_voucher_detail(v_id)
                if detail:
                    self.send_json(detail)
                else:
                    self.send_json({"error": "Voucher not found"}, 404)
            except ValueError:
                self.send_json({"error": "Invalid Voucher ID"}, 400)

        elif path == '/api/trial-balance':
            self.send_json(tally_db.get_trial_balance())

        elif path == '/api/ledgers':
            self.send_json(tally_db.get_ledgers_list())

        elif path == '/api/ledger-statement':
            ledger_name = params.get('name', [''])[0]
            date_from = params.get('date_from', [''])[0]
            date_to = params.get('date_to', [''])[0]
            if not ledger_name:
                self.send_json({"error": "Ledger name required"}, 400)
            else:
                self.send_json(tally_db.get_ledger_statement(ledger_name, date_from, date_to))

        elif path == '/api/gstr2b/results':
            if tally_db.reco_results is None:
                auto_load_r2b_files()
            self.send_json(tally_db.reco_results or {"summary": {}, "records": []})

        elif path == '/api/session/export':
            session_data = tally_db.export_session()
            body = json.dumps(session_data, indent=2).encode('utf-8')
            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.send_header('Content-Disposition', 'attachment; filename="reconciliation_progress.rex"')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        elif path == '/api/export/tally-xml':
            xml_data = tally_db.export_tally_xml()
            # Add subtle branding comment in XML
            if xml_data.startswith('<?xml'):
                first_nl = xml_data.find('?>')
                if first_nl != -1:
                    xml_data = xml_data[:first_nl+2] + '\n<!-- Generated by Rexinux\'s Tally Analyser -->\n' + xml_data[first_nl+2:]
            body = xml_data.encode('utf-8')
            self.send_response(200)
            self.send_header('Content-Type', 'application/xml; charset=utf-8')
            self.send_header('Content-Disposition', 'attachment; filename="Rexinux_Tally_Export.xml"')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        elif path == '/api/export':
            rtype = params.get('type', ['daybook'])[0]
            l_name = params.get('name', [''])[0]
            csv_data = tally_db.export_csv(rtype, ledger_name=l_name)
            
            body = csv_data.encode('utf-8')
            self.send_response(200)
            self.send_header('Content-Type', 'text/csv; charset=utf-8')
            self.send_header('Content-Disposition', f'attachment; filename="Rexinux_{rtype}_Export.csv"')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        elif path == '/api/gstr2b/clear':
            r2b_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'R2B')
            if os.path.exists(r2b_dir):
                for f in os.listdir(r2b_dir):
                    fp = os.path.join(r2b_dir, f)
                    if os.path.isfile(fp) and not f.startswith('.'):
                        try: os.remove(fp)
                        except Exception: pass
            tally_db.reco_results = None
            self.send_json({"success": True, "message": "GSTR-2B files cleared"})

        else:
            self.send_json({"error": "Endpoint not found"}, 404)

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        length = int(self.headers.get('Content-Length', 0))
        raw_body = self.rfile.read(length)

        if parsed.path == '/api/session/import':
            try:
                payload = json.loads(raw_body.decode('utf-8'))
                info = tally_db.import_session(payload)
                self.send_json({"success": True, "info": info, "reco_results": tally_db.reco_results})
            except Exception as e:
                self.send_json({"error": f"Failed to import session: {str(e)}"}, 400)

        elif parsed.path == '/api/gstr2b/update-record':
            try:
                payload = json.loads(raw_body.decode('utf-8'))
                updated = tally_db.update_reco_record(payload.get('id'), payload.get('note'), payload.get('isRectified'))
                self.send_json({"success": updated, "summary": tally_db.reco_results['summary'] if tally_db.reco_results else {}})
            except Exception as e:
                self.send_json({"error": f"Failed to update record: {str(e)}"}, 400)

        elif parsed.path == '/api/gstr2b/reconcile':
            try:
                content_type = self.headers.get('Content-Type', '')
                all_invoices = []

                if 'multipart/form-data' in content_type:
                    boundary_match = re.search(r'boundary=([^\s;]+)', content_type)
                    if boundary_match:
                        boundary = boundary_match.group(1).encode('utf-8')
                        if boundary.startswith(b'"') and boundary.endswith(b'"'):
                            boundary = boundary[1:-1]
                        parts = raw_body.split(b'--' + boundary)
                        for part in parts:
                            if b'filename=' in part:
                                header_end = part.find(b'\r\n\r\n')
                                if header_end != -1:
                                    headers_text = part[:header_end].decode('utf-8', errors='ignore')
                                    fn_match = re.search(r'filename="([^"]+)"', headers_text)
                                    filename = fn_match.group(1) if fn_match else "file.xlsx"
                                    file_bytes = part[header_end + 4:].rstrip(b'\r\n--')

                                    if filename.lower().endswith('.zip'):
                                        try:
                                            with zipfile.ZipFile(io.BytesIO(file_bytes), 'r') as z:
                                                for zname in z.namelist():
                                                    if zname.lower().endswith('.xlsx'):
                                                        invs = parse_gstr2b_xlsx_bytes(z.read(zname), filename=zname)
                                                        all_invoices.extend(invs)
                                                    elif zname.lower().endswith('.json'):
                                                        invs = parse_gstr2b_json_bytes(z.read(zname), filename=zname)
                                                        all_invoices.extend(invs)
                                        except Exception as e:
                                            print(f"Error reading zip part {filename}: {e}")
                                    elif filename.lower().endswith('.xlsx'):
                                        invs = parse_gstr2b_xlsx_bytes(file_bytes, filename=filename)
                                        all_invoices.extend(invs)
                                    elif filename.lower().endswith('.json'):
                                        invs = parse_gstr2b_json_bytes(file_bytes, filename=filename)
                                        all_invoices.extend(invs)
                elif 'application/json' in content_type:
                    invs = parse_gstr2b_json_bytes(raw_body, filename="raw_data.json")
                    all_invoices.extend(invs)
                else:
                    filename = self.headers.get('X-File-Name', 'file.xlsx')
                    if filename.lower().endswith('.xlsx'):
                        all_invoices.extend(parse_gstr2b_xlsx_bytes(raw_body, filename=filename))
                    elif filename.lower().endswith('.json'):
                        all_invoices.extend(parse_gstr2b_json_bytes(raw_body, filename=filename))

                reco = tally_db.reconcile_gstr2b(all_invoices)
                self.send_json(reco)
            except Exception as e:
                self.send_json({"error": f"Failed to reconcile GSTR-2B files: {str(e)}"}, 400)

        elif parsed.path == '/api/upload':
            content_type = self.headers.get('Content-Type', '')
            file_bytes = None
            filename = self.headers.get('X-File-Name', 'uploaded_file.zip')
            
            if 'multipart/form-data' in content_type:
                boundary = content_type.split('boundary=')[-1].encode('utf-8')
                parts = raw_body.split(b'--' + boundary)
                for part in parts:
                    if b'filename=' in part:
                        header_end = part.find(b'\r\n\r\n')
                        if header_end != -1:
                            headers_text = part[:header_end].decode('utf-8', errors='ignore')
                            fn_match = re.search(r'filename="([^"]+)"', headers_text)
                            if fn_match:
                                filename = fn_match.group(1)
                            file_bytes = part[header_end + 4:].rstrip(b'\r\n--')
                            break
            else:
                file_bytes = raw_body
                
            if file_bytes:
                try:
                    is_rex = filename.lower().endswith('.rex')
                    if not is_rex and file_bytes.strip().startswith(b'{'):
                        try:
                            parsed_json = json.loads(file_bytes.decode('utf-8', errors='ignore'))
                            if isinstance(parsed_json, dict) and ('tally_data' in parsed_json or 'reco_results' in parsed_json or parsed_json.get('file_type') == 'rex_session'):
                                is_rex = True
                                file_bytes = parsed_json
                        except Exception:
                            pass
                    
                    if is_rex:
                        session_json = file_bytes if isinstance(file_bytes, dict) else json.loads(file_bytes.decode('utf-8'))
                        info = tally_db.import_session(session_json)
                        self.send_json({"success": True, "info": info, "reco_results": tally_db.reco_results, "is_session": True})
                    else:
                        tally_db.parse_file(file_bytes, filename=filename)
                        auto_load_r2b_files()
                        self.send_json({"success": True, "info": tally_db.get_info()})
                except Exception as e:
                    self.send_json({"error": f"Failed to parse file: {str(e)}"}, 400)
            else:
                self.send_json({"error": "No file content received"}, 400)
        else:
            self.send_json({"error": "Method not allowed"}, 405)


def start_server(port=8000):
    base_dir = os.path.dirname(os.path.abspath(__file__))
    xml_path = os.path.join(base_dir, 'AQUAXMLFY26.xml')
    zip_path = os.path.join(base_dir, 'AivarFY26.zip')
    
    target_file = None
    if os.path.exists(xml_path):
        target_file = xml_path
    elif os.path.exists(zip_path):
        target_file = zip_path

    if target_file:
        print(f"Loading initial file: {target_file}...")
        tally_db.parse_file(target_file, filename=os.path.basename(target_file))
        print("Tally Data loaded successfully!")
        
        auto_load_r2b_files()

    server_address = ('', port)
    httpd = HTTPServer(server_address, RequestHandler)
    print(f"==================================================")
    print(f"🚀 Rexinux's Tally Analyser server running on:")
    print(f"👉 http://localhost:{port}")
    print(f"==================================================")
    
    try:
        import webbrowser
        webbrowser.open(f"http://localhost:{port}")
    except Exception:
        pass
        
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down server...")
        httpd.server_close()

if __name__ == '__main__':
    port = 8000
    if len(sys.argv) > 1:
        try: port = int(sys.argv[1])
        except ValueError: pass
    start_server(port)
