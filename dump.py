import sqlite3
try:
    c = sqlite3.connect('c:/MARTIN/OPERACIONES/onpe_consultas_copy.db')
    c.row_factory = sqlite3.Row
    rs = c.execute("SELECT error_msg FROM consultas WHERE estado='error' OR error_msg != ''").fetchall()
    out = "\n".join([r['error_msg'] for r in rs])
    with open('c:/MARTIN/OPERACIONES/error_dump.txt', 'w', encoding='utf-8') as f:
        f.write(out if out else 'Not found')
except Exception as e:
    with open('c:/MARTIN/OPERACIONES/error_dump.txt', 'w', encoding='utf-8') as f:
        f.write(str(e))
