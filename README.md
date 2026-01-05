# MusicSheets

A TypeScript-powered Google Apps Script for creating and managing music chord sheets in Google Sheets with intelligent chord transposition and automatic lyrics synchronization.

📖 [**See the MusicSheet Template**](https://docs.google.com/spreadsheets/d/1h_ihN9vbaUChdbwEjw5hraGJOGdIu1QCsK9p_64hdpg/edit?usp=sharing)

## ✨ Features

### 🎵 Core Functionality
- **Dual-Sheet Architecture**: Separate sheets for lyrics ("Letra") and chords ("Acordes") with automatic synchronization
- **Smart Chord Parser**: Recognizes all standard chord notations including:
  - Root notes with sharps/flats (C, C#, Db, etc.)
  - Chord qualities (m, maj7, sus4, dim, aug, etc.)
  - Slash chords (e.g., C/G, D/F#)
  - Multiple accidentals (e.g., C##, Dbb)

### 🎹 Transposition Engine
- **Instant Transposition**: Transpose all chords up or down by semitones
- **Auto-Transpose Mode**: Automatically transpose all chords when the key signature changes
- **Smart Key Updates**: Key signature updates automatically when transposing
- **Invalid Chord Handling**: Chords that can't be parsed are marked with `!` prefix

### 🔄 Automatic Synchronization
- **Real-time Lyrics Sync**: Changes in the Lyrics sheet automatically propagate to the Chords sheet
- **Bidirectional Updates**: Editing lyrics rows in the Chords sheet syncs back to the Lyrics sheet
- **Structural Change Detection**: Automatically adjusts to row/column insertions and deletions
- **Working Area Awareness**: Only syncs content within the active working area (respects frozen rows/columns)

### 📐 Smart Formatting
- **Monospace Grid Layout**: Uses Space Mono font for perfect chord alignment
- **Dynamic Column Widths**: Alternates between normal (15px) and wide (17px) columns every 6 columns
- **Special Column Handling**: Extra padding for frozen columns
- **Print Sheet Support**: Special formatting for header/footer areas
- **Consistent Row Heights**: Fixed 21px row height for uniform spacing

### 🛡️ Data Protection
- **Automatic Validation**: Invalid chords are marked for easy identification
- **Auto-Transpose Safety**: Disables auto-transpose if key signature becomes invalid
- **Error Handling**: All hooks wrapped in try-catch with user-friendly toast notifications

## 🏗️ Architecture

### Sheet Structure

The spreadsheet expects three sheets:

1. **"Letra" (Lyrics Sheet)**: Clean lyrics editing interface
2. **"Acordes" (Chords Sheet)**: Interleaved chords and lyrics with 2:1 row ratio
3. **"Impresión" (Print Sheet)**: Optional print-formatted output

### Required Named Ranges

Configure these named ranges in your Google Sheet:

| Range Name | Purpose | Example |
|------------|---------|---------|
| `Tonalidad` | Key signature cell | `C`, `G`, `Bb`, etc. |
| `Auto_Trasponer` | Auto-transpose toggle checkbox | `TRUE`/`FALSE` |
| `Título` | Document title (auto-populated) | Sheet name |
| `Ideas_Sueltas` | Right-side notes area in Lyrics sheet | Multi-column range |
| `Encabezado` | Print sheet header area | Header rows |
| `Pie_de_Página` | Print sheet footer area | Footer rows |

### Code Organization

```
src/
├── Chords.ts          # Chord and pitch manipulation
│   ├── Chord          # Main chord class with parse/transpose
│   └── Pitch          # Musical pitch with enharmonic equivalents
├── Range.ts           # Google Sheets Range extensions
│   ├── translate()    # Move range by offset
│   ├── scale()        # Multiply range dimensions
│   ├── resize()       # Add/subtract rows/columns
│   ├── intersect()    # Find overlap between ranges
│   └── projectInto()  # Copy range to different sheet
├── MusicSheet.ts      # Main spreadsheet logic
│   ├── Hooks          # onOpen, onEdit, onChange
│   ├── Sync Logic     # Lyrics ↔ Chords synchronization
│   └── Actions        # Transpose, formatting, triggers
└── Utils.ts           # Helper functions
```

## 🚀 Getting Started

### Prerequisites

- Node.js 16+ and npm
- Google account with Google Sheets access
- [Google Apps Script CLI (clasp)](https://github.com/google/clasp)

### Installation

1. **Clone and install dependencies:**
   ```bash
   git clone <repository-url>
   cd MusicSheets
   npm install
   ```

2. **Authenticate with Google:**
   ```bash
   npx clasp login
   ```

3. **Connect to your spreadsheet:**

   **Option A - Use existing spreadsheet:**
   - Open your Google Sheet
   - Go to Extensions → Apps Script
   - Copy the Script ID from the URL (between `/d/` and `/edit`)
   - Create `.clasp.json`:
     ```json
     {
       "scriptId": "YOUR_SCRIPT_ID_HERE",
       "rootDir": "./dist"
     }
     ```

   **Option B - Create new spreadsheet:**
   ```bash
   npx clasp create --type sheets --title "MusicSheets"
   ```

4. **Configure your sheet:**
   - Create three sheets: "Letra", "Acordes", "Impresión"
   - Set up named ranges (see Required Named Ranges above)
   - Freeze first row and column in each sheet

### Development Workflow

```bash
# Run tests
npm test

# Build TypeScript
npm run build

# Deploy to Google Apps Script (includes tests, lint, build, push)
npm run push

# Deploy and create version
npm run deploy

# Watch mode for development
npm run watch

# Lint code
npm run lint
npm run lint:fix
```

## 📝 Usage

### Available Functions

Call these from Tools → Macros or create custom menu buttons:

```javascript
// Transpose all chords up by 1 semitone
transposeUp()

// Transpose all chords down by 1 semitone
transposeDown()

// Install onChange trigger for automatic sync on structural changes
setupTriggers()

// Reset all formatting (fonts, column widths, row heights)
resetFormatting()
```

### Setting Up Triggers

The script uses three types of triggers:

1. **onOpen** (automatic): Updates document title when spreadsheet opens
2. **onEdit** (automatic): Syncs lyrics and handles key changes on cell edits
3. **onChange** (manual setup): Handles structural changes (row/column insertions)

To enable structural change detection:
1. Run `setupTriggers()` from Apps Script editor
2. Or call it once from Tools → Script editor → Run

### Chord Format Examples

The parser recognizes these chord formats:

```
Valid:
C, D, E, F, G, A, B
C#, Db, F##, Gbb
Cm, CM7, Cmaj7, C7, Csus4, Cadd9
Cdim, Caug, C9, C13
C/G, Am/C, D/F#

Invalid (will be marked with !):
H, I, J (not musical notes)
C#b (conflicting accidentals)
C// (invalid format)
```

## 🔧 Technical Details

### Range Extensions

The project extends Google Apps Script's Range class with convenient methods:

```typescript
// Move range by offset
range.translate(3, 2)  // Move 3 columns right, 2 rows down

// Move range to absolute position
range.translateTo(5, 10)  // Move to column E, row 10

// Scale range dimensions
range.scale(2, 1.5)  // Double width, multiply height by 1.5

// Add/subtract dimensions
range.resize(-2, 3)  // Subtract 2 columns, add 3 rows

// Set exact dimensions
range.resizeTo(10, 5)  // Resize to exactly 10 columns, 5 rows

// Check overlap
range.overlapsWith(otherRange)  // Returns boolean

// Find intersection
range.intersect(otherRange)  // Returns Range | undefined

// Project to different sheet
range.projectInto(targetSheet)  // Same position on different sheet
```

### Sync Algorithm

The lyrics synchronization uses a sophisticated mapping algorithm:

1. **Lyrics → Chords**: Each lyrics row maps to every other row (odd rows) in the chords sheet with 2:1 scaling
2. **Chords → Lyrics**: Edits to even rows in chords sheet sync back to lyrics sheet with 0.5:1 scaling
3. **Working Area Isolation**: Only syncs content within the working area (excludes frozen rows/columns and right tray)
4. **Offset Preservation**: Maintains relative positioning when source range is subset of working area

### Type Safety

This project uses strict TypeScript with:
- ✅ No `any` types allowed
- ✅ Strict null checks enabled
- ✅ Unused variables/parameters detected
- ✅ Full type coverage for Google Apps Script APIs
- ✅ Custom type declarations for Range extensions

### Testing

Tests are written using Vitest:

```bash
npm test              # Run all tests
npm run test:watch    # Watch mode
npm run test:ui       # Visual test UI
```

Test coverage includes:
- ✅ Chord parsing and transposition
- ✅ Pitch manipulation with enharmonics
- ✅ Range extension methods
- ✅ Utility functions

## 🎨 Customization

### Formatting Constants

Edit these in `src/MusicSheet.ts`:

```typescript
const FONT_FAMILY = "Space Mono"
const FONT_SIZE = 10
const ROW_HEIGHT = 21
const NORMAL_COLUMN_WIDTH = 15
const WIDE_COLUMN_WIDTH = 17
const WIDE_COLUMN_PERIODICITY = 6  // Wide column every N columns
const PADDING = 3
```

### Sheet Names

Change sheet names by modifying:

```typescript
const LYRICS_SHEET_NAME = "Letra"
const CHORDS_SHEET_NAME = "Acordes"
const PRINT_SHEET_NAME = "Impresión"
```

### Named Range References

Update named range names:

```typescript
const KEY_RANGE_NAME = "Tonalidad"
const AUTOTRANSPOSE_RANGE_NAME = "Auto_Trasponer"
const DOCUMENT_TITLE_RANGE_NAME = "Título"
// ... etc
```

## 🤝 Contributing

This is a private project, but contributions are welcome:

1. Write tests for new features
2. Follow existing code style (enforced by ESLint)
3. Use TypeScript strict mode
4. Run `npm run lint:fix` before committing

## 📄 License

Private project

## 🐛 Troubleshooting

**Lyrics not syncing?**
- Check that sheets are named correctly ("Letra", "Acordes")
- Verify frozen rows/columns are set up
- Run `setupTriggers()` to enable onChange hook

**Transpose not working?**
- Check that "Tonalidad" named range exists
- Verify chords are in valid format (C, Am, G7, etc.)
- Look for `!` prefix on invalid chords

**Formatting issues?**
- Run `resetFormatting()` to reapply all formatting
- Check that named ranges for header/footer exist if using print sheet

**Changes not deploying?**
- Run `npm run build` to compile TypeScript
- Check `.clasp.json` has correct scriptId
- Try `npx clasp login` to reauthenticate
