# DEKRA Gutachten-Erstellungs-App

Eine professionelle Web-Anwendung zur Erstellung von DEKRA Unfallgutachten mit Live-Vorschau und Export-Funktionen.

## Features

- ✅ Live-Vorschau im professionellen DEKRA Word-Dokument Style
- ✅ PDF-Export mit vollständiger Formatierung und Anhängen
- ✅ Word-Export (.docx) mit exaktem Layout
- ✅ Anhänge-Verwaltung (PDF, Word, Bilder)
- ✅ Vorlagen-System zum Speichern und Laden
- ✅ Auto-Save Funktion
- ✅ Responsive Design

## Deployment auf Vercel

### Option 1: Deploy mit Vercel CLI

1. Installiere Vercel CLI:
```bash
npm i -g vercel
```

2. Im Projektordner ausführen:
```bash
vercel
```

3. Folge den Anweisungen im Terminal

### Option 2: Deploy über GitHub

1. Pushe das Projekt zu GitHub:
```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/DEIN-USERNAME/dekra-gutachten-app.git
git push -u origin main
```

2. Gehe zu [vercel.com](https://vercel.com)
3. Klicke auf "New Project"
4. Importiere dein GitHub Repository
5. Klicke auf "Deploy"

### Option 3: Direkter Upload

1. Gehe zu [vercel.com](https://vercel.com)
2. Klicke auf "New Project"
3. Wähle "Upload Folder"
4. Ziehe den Projektordner in das Upload-Feld
5. Klicke auf "Deploy"

## Lokale Entwicklung

```bash
# Installiere Dependencies (optional)
npm install

# Starte lokalen Server
npm start
# oder einfach
npx serve
```

## Struktur

```
├── index.html          # Haupt-HTML mit Formular und Vorschau
├── styles.css          # Styling für die App
├── print.css           # Print-Styles
├── app.js              # Haupt-JavaScript Logik
├── vercel.json         # Vercel Konfiguration
├── package.json        # NPM Konfiguration
└── README.md           # Diese Datei
```

## Technologien

- Vanilla JavaScript (kein Framework benötigt)
- jsPDF für PDF-Generierung
- PDFLib für PDF-Zusammenführung
- docx.js für Word-Export
- LocalStorage für Auto-Save

## Browser-Kompatibilität

- Chrome/Edge (empfohlen)
- Firefox
- Safari
- Opera

## Lizenz

Privates Projekt für DEKRA Gutachten-Erstellung.