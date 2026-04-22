# CORP Valuation app v4.6 by G. Tsakalos

Desktop εφαρμογή σε **Python / Tkinter / SQLite** για αποτίμηση επιχειρήσεων, εισαγωγή οικονομικών δεδομένων, υπολογισμό χρηματοοικονομικών δεικτών, σύγκριση εταιρειών και εξαγωγή αναφορών.

---

## 🔄 Τι νέο υπάρχει από v4.1

Σε σχέση με την έκδοση v4.1, η v4.6 περιλαμβάνει:

### UI / UX Improvements
- ✔ Μορφοποίηση αριθμών με διαχωριστικά χιλιάδων (.)
- ✔ Smart input behavior (focus/blur formatting)
- ✔ Βελτιωμένο scrolling
- ✔ Καλύτερο layout

### Data Handling
- ✔ Βελτιωμένος parser αριθμών
- ✔ Ασφαλής μετατροπή σε float
- ✔ Consistent formatting σε UI και exports

### Reports & Export
- ✔ XLSX formatting
- ✔ PDF reports με charts
- ✔ Sheet 'Charts' στο Excel

### Νέα Features
- ✔ Compare tab
- ✔ Notes ανά δείκτη
- ✔ Multi-company PDF export
- ✔ Excel import
- ✔ URL buttons
- ✔ Config file support

---

## 📊 Βασικά χαρακτηριστικά

- Διαχείριση εταιρειών
- Ισολογισμός
- Αποτελέσματα Χρήσης
- Δείκτες
- Compare
- Export PDF / XLSX
- SQLite βάση

---

## 📂 Αρχεία έργου

- app.py
- app.cfg
- madka_values.sqlite
- requirements.txt
- IndexMap.jpg

---

## ▶️ Εκτέλεση

```bash
python -m pip install -r requirements.txt
python app.py
```

---

## 📦 Build (.exe)

```bash
pyinstaller --onefile --windowed app.py
```
