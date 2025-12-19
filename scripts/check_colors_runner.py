import sys
sys.path.append(r'c:\Users\Propietario\Downloads\makinas')
from maquinas_app import get_db
conn = get_db()
rows = conn.execute('SELECT color, COUNT(*) as c FROM machines GROUP BY color').fetchall()
for r in rows:
    print(r['color'], r['c'])
print('\nSample:')
rows2 = conn.execute('SELECT id,name,priority,color FROM machines LIMIT 20').fetchall()
for r in rows2:
    print(r['id'], r['name'], r['priority'], r['color'])
conn.close()