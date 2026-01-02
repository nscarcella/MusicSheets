# MusicSheets

Google Apps Script for writting music in Google Sheets See [the MusicSheet Template](https://docs.google.com/spreadsheets/d/1h_ihN9vbaUChdbwEjw5hraGJOGdIu1QCsK9p_64hdpg/edit?usp=sharing)

## Features

- 🎵 **Automatic Lyrics Sync**: Lyrics sheet automatically syncs to chords sheet
- 🎹 **Chord Transposition**: Transpose up/down with automatic key signature update
- 🔄 **Auto-transpose**: Automatically transpose all chords when changing the key
- 📊 **Smart Formatting**: Grid layout with proper spacing and fonts
- 🛡️ **Protected Lyric Rows**: Even rows in chords sheet auto-revert to lyrics

## Development Setup

### Prerequisites

- Node.js and npm installed
- Google account with access to Google Sheets

### Installation

1. **Install dependencies:**

   ```bash
   npm install
   ```

2. **Authenticate with Google:**

   ```bash
   npx clasp login
   ```

3. **Link to existing spreadsheet OR create new:**

   **Option A - Link existing spreadsheet:**

   - Open your Google Sheet
   - Go to Extensions > Apps Script
   - Copy the Script ID from the URL (between `/d/` and `/edit`)
   - Update `.clasp.json` with your `scriptId`:
     ```json
     {
       "scriptId": "YOUR_SCRIPT_ID_HERE",
       "rootDir": "./dist"
     }
     ```

   **Option B - Create new project:**

   ```bash
   npx clasp create --type sheets --title "MusicSheets"
   ```

### Development Workflow

1. **Edit TypeScript files** in `src/`:

   - `src/Chords.ts` - Chord manipulation logic
   - `src/MusicSheet.ts` - Main spreadsheet logic

2. **Build TypeScript:**

   ```bash
   npm run build
   ```

3. **Push to Google Apps Script:**

   ```bash
   npm run push
   ```

4. **Watch mode** (auto-compile on save):
   ```bash
   npm run watch
   ```

## Project Structure

```
MusicSheets/
├── src/
│   ├── Chords.ts          # Chord parsing and transposition
│   ├── MusicSheet.ts      # Main spreadsheet logic
│   └── appsscript.json    # Apps Script manifest
├── dist/                  # Compiled JavaScript (git-ignored)
├── package.json           # Dependencies and scripts
├── tsconfig.json          # TypeScript configuration
└── .clasp.json           # Clasp configuration (git-ignored)
```

## Usage

### Named Ranges (required in your Google Sheet)

- `Tonalidad` - Key signature cell
- `Auto_Trasponer` - Auto-transpose checkbox
- `Título` - Document title cell
- `Ideas_Sueltas` - Right tray area in Lyrics sheet
- `Encabezado` - Print header area
- `Pie_de_Página` - Print footer area

### Available Functions

- `transposeUp()` - Transpose all chords up by 1 semitone
- `transposeDown()` - Transpose all chords down by 1 semitone
- `setupTriggers()` - Install onChange trigger for auto-sync
- `resetFormatting()` - Apply grid formatting to all sheets

## Type Safety

The project uses strict TypeScript with:

- Full type checking for Google Apps Script APIs
- No `any` types
- Null safety checks
- Unused variable/parameter detection

## License

Private project
