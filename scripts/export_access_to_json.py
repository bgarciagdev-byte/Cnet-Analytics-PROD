import json
import os
import datetime as dt
import pyodbc

# CAMBIA ESTA RUTA
ACCESS_PATH = r"C:\Users\bgarcia\CENTRONET S.A.S\Datos, Analítica e IA - General\Proyectos D&A\2. EJECUCIÓN PROYECTOS\3. BASE DE DATOS ACCESS"

# salida al repo -> data/dashboard.json
BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
OUTPUT_JSON = os.path.join(BASE_DIR, "data", "dashboard.json")

EXPORTS = [
    {"name": "kpis", "sql": "SELECT * FROM KPIs"},
    {"name": "proyectos", "sql": "SELECT * FROM Proyectos"},
]

def connect_access(path: str):
    conn_str = (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        rf"DBQ={path};"
    )
    return pyodbc.connect(conn_str)

def rows_to_dicts(cursor):
    cols = [c[0] for c in cursor.description]
    out = []
    for row in cursor.fetchall():
        d = {}
        for i, v in enumerate(row):
            if hasattr(v, "isoformat"):
                d[cols[i]] = v.isoformat()
            else:
                d[cols[i]] = v
        out.append(d)
    return out

def main():
    os.makedirs(os.path.dirname(OUTPUT_JSON), exist_ok=True)

    conn = connect_access(ACCESS_PATH)
    cur = conn.cursor()

    payload = {"generated_at": dt.datetime.now().isoformat(), "datasets": {}}

    for item in EXPORTS:
        cur.execute(item["sql"])
        payload["datasets"][item["name"]] = rows_to_dicts(cur)

    conn.close()

    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    print("OK ->", OUTPUT_JSON)

if __name__ == "__main__":
    main()