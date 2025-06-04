# PokExcel

A simple Python utility for writing values to Excel files or SQLite databases from the command line.

## Features

- Write values to Excel files or SQLite databases
- Support for batch operations via CSV files
- Excel mode with visibility control (hidden/shown)
- Simple command-line interface

## Usage

### Single Value Write

#### Excel Mode
```bash
main.py Excel <excel_path> [/H | /S] /<sheet>:<cell>=<value>
```
- `/H` - Hidden Excel window
- `/S` - Shown Excel window

Example:
```bash
main.py Excel document.xlsx /1:A1=100
```

#### SQL Mode
```bash
main.py SQL <sqlite_db> /<sheet>:<cell>=<value>
```

Example:
```bash
main.py SQL database.sqlite /1:A1=100
```

### Batch Operations (CSV)

Both modes support batch operations using CSV files:
```bash
main.py [Excel|SQL] --csv data.csv
```

## Requirements

- Python 3.x
- SQLite3 (included in Python)
- Excel COM interface (for Excel operations)

## Error Handling

The utility includes error handling for:
- Database operations
- File access
- Invalid input formats
- Excel operations

## License

This project is available under open-source terms.
