import sys
import ast
import sqlite3
from pathlib import Path

try:
    # Pokud je přítomen writer (Excel režim)
    from writer import open_isolated_excel, write_rows_to_workbook, cleanup_excel
    writer_loaded = True
except ImportError:
    writer_loaded = False

import sqlite3

def zapis_do_sql(db_path, sheet_num, cell, value):
    """Jednoduchý zápis do SQLite tabulky zapis, s ošetřením chyb a vytvořením DB/tabulky pokud neexistuje."""
    conn = None
    try:
        # Pokus o připojení (vytvoří DB, pokud není)
        conn = sqlite3.connect(db_path)
        c = conn.cursor()
        # Vytvoření tabulky, pokud neexistuje
        c.execute('''
            CREATE TABLE IF NOT EXISTS zapis (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sheet INTEGER,
                cell TEXT,
                value TEXT
            )
        ''')
        # Vložení dat
        c.execute(
            'INSERT INTO zapis (sheet, cell, value) VALUES (?, ?, ?)',
            (int(sheet_num), cell, value)
        )
        conn.commit()
    except Exception as e:
        print(f"Chyba při zápisu do databáze: {e}")
    finally:
        if conn:
            conn.close()


def _show_help():
    print(
        "Použití:\n\n"
        "1) Jednotlivý zápis (Excel):\n"
        "   main.py Excel <excel_path> [/H | /S] /<sheet>:<cell>=<value>\n"
        "   main.py SQL   <sqlite_db>  /<sheet>:<cell>=<value>\n"
        "   main.py [Excel|SQL] --csv <data.csv>\n"
        "\n"
        "Příklad (Excel): main.py Excel soubor.xlsx /1:A1=100\n"
        "Příklad (SQL):   main.py SQL   db.sqlite /1:A1=100\n"
        "CSV dávka:       main.py Excel --csv data.csv\n"
    )
    sys.exit(0)

def _rows_from_csv(csv_path: Path):
    if not csv_path.exists():
        print(f"CSV nenalezeno: {csv_path}", file=sys.stderr)
        sys.exit(1)
    rows = []
    with csv_path.open("r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                rows.append(ast.literal_eval(line))
            except Exception as e:
                print(f"Chyba v CSV: {line}\n{e}", file=sys.stderr)
                sys.exit(1)
    # po načtení vyprázdni CSV
    csv_path.write_text("", encoding="utf-8")
    return rows

def _cli_parse():
    argv = sys.argv[1:]
    if not argv or argv[0] in ("--help", "-h"):
        _show_help()

    # Prefix (Excel nebo SQL)
    if argv[0] in ("Excel", "SQL"):
        mode = argv[0]
        argv = argv[1:]
    else:
        mode = "Excel"  # Defaultně Excel

    # CSV dávka
    if argv and argv[0] == "--csv":
        if len(argv) != 2:
            _show_help()
        return mode, False, _rows_from_csv(Path(argv[1]))

    if not argv:
        _show_help()
    file_path = argv[0]
    idx = 1

    # viditelnost pro Excel
    excel_visible = False
    if mode == "Excel" and idx < len(argv) and argv[idx] in ("/H", "/S"):
        excel_visible = (argv[idx] == "/S")
        idx += 1

    if idx >= len(argv):
        _show_help()
    raw = argv[idx].lstrip("/")  # Remove leading slash, e.g., /1:A6=456

    # --- ZDE je hlavní změna ---
    try:
        # Rozděluje pouze podle ":", NIKOLIV podle lomítek
        sheet_str, rest = raw.split(":", 1)
        cell_str, value = rest.split("=", 1)
        assert sheet_str.isdigit() and cell_str[0].isalpha()
    except Exception:
        print(f"Chybný formát zápisu: {raw}", file=sys.stderr)
        print("Očekávaný tvar: 1:A6=456 (list:buňka=hodnota)", file=sys.stderr)
        sys.exit(1)

    file_path = str(Path(file_path).resolve())
    return mode, excel_visible, [[file_path, sheet_str, cell_str, value]]



def main():
    print(f"\nSpouštím {__file__}...", sys.argv)
    mode, excel_visible, rows = _cli_parse()
    print(f"[INFO] Režim: {mode}, Viditelnost Excelu: {excel_visible}")
    if mode == "Excel":
        if not writer_loaded:
            print("Chybí knihovna writer.py nebo její závislosti!", file=sys.stderr)
            sys.exit(1)
        excel = None
        try:
            excel = open_isolated_excel(visible=excel_visible)
            # The writer expects a list of lists, where the inner lists have the format:
            # [excel_path, assignment_string]
            # The assignment string should be in the format: "/sheet:cell=value"
            # We need to reconstruct this format from the parsed values.
            processed_rows = [[rows[0][0], f"/{rows[0][1]}:{rows[0][2]}={rows[0][3]}"]]
            write_rows_to_workbook(excel, processed_rows)
            print(f"[OK] Zapsáno do Excelu: {rows}")
        except Exception as e:
            print(f"[CHYBA] Excel zápis: {e}", file=sys.stderr)
            sys.exit(1)
        finally:
            cleanup_excel(excel)

    elif mode == "SQL":
        try:
            for (db_path, sheet, cell, value) in rows:
                zapis_do_sql(db_path, sheet, cell, value)
                print(f"[OK] SQL: {db_path} sheet={sheet} cell={cell} value={value}")
        except Exception as e:
            print(f"[CHYBA] SQL zápis: {e}", file=sys.stderr)
            sys.exit(1)
    else:
        print(f"Neznámý režim: {mode}")
        sys.exit(1)

if __name__ == "__main__":
    print("Spouštím main.py...", sys.argv)
    main()
