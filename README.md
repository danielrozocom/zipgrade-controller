# ZipGrade Controller 📝

A Google Apps Script library to synchronize ZipGrade exam results, consolidate auto-formulas, and manage subjects/sessions dynamically in Google Sheets with a premium HTML/JS interface.

---

## 🚀 Key Features

- **ZipGrade Sync & Auto-Alignment**: Parses and processes ZipGrade CSV exports. It automatically aligns scores with the master student list (`S1`) and flags inconsistencies (missing students, unregistered students).
- **Interactive Subject & Session Manager**: A drag-and-drop HTML modal that allows reordering subjects across multiple session sheets, updating question counts, and adding/removing evaluation columns.
- **Dynamic Session Sheets**: Automatically creates session sheets (`S1`, `S2`, etc.) with standard formatting. Names in secondary sessions are linked dynamically to the master list via formulas to maintain sync.
- **JSON Configuration Import/Export**: Allows saving, backing up, and importing the entire evaluation structure (subject ordering, sessions, and question counts) via JSON. The import algorithm is smart enough to match existing subjects and preserve student grades.
- **Auto-Formulas & Formatting**: Dynamically builds formulas for totals, averages, and zebra striping. Includes automated conditional formatting to highlight top performers.

---

## 🛠️ Installation & Setup (How to use as a Library)

This project is structured as a **Google Apps Script Library** to allow easy updates across multiple spreadsheets without copying 2800+ lines of code.

### Step 1: Deploy the Library
1. Create a new Google Apps Script project.
2. Place `Code.js` and `appsscript.json` into the editor.
3. Deploy the project as a library:
   - Go to **Deploy** > **New deployment**.
   - Select **Library** as the deployment type.
   - Enter a description (e.g., `ZipGrade Controller Library`) and click **Deploy**.
4. Copy the **Script ID** from the deployment details.

### Step 2: Add the Library to Target Sheets
1. Open the target Google Sheet where you want to use the controller.
2. Open its Apps Script editor (**Extensions** > **Apps Script**).
3. On the left pane, click the `+` icon next to **Libraries**.
4. Paste the **Script ID** copied in Step 1, click **Look up**, choose the latest version, set the Identifier to `ZipGradeController`, and click **Add**.

### Step 3: Add the Client Bridge Code
1. Copy the contents of the [`client_code.js`](./client_code.js) file from this repository.
2. Paste them into the `Code.gs` file of your target sheet.
3. Save the project and refresh your Google Sheet. The `📝 ZipGrade Controller` menu will appear automatically.

---

## 📥 JSON Configuration Schema

You can backup or import subject/session configurations using a simple JSON format. The importer automatically handles sheet creation and preserves existing columns' grade data:

```json
[
  {
    "name": "MATEMÁTICAS",
    "count": 10,
    "session": "1"
  },
  {
    "name": "FÍSICA",
    "count": 15,
    "session": "1"
  },
  {
    "name": "QUÍMICA",
    "count": 8,
    "session": "2"
  }
]
```

---

## 💻 Local Development (using clasp)

You can write and modify code locally on your computer and sync it directly to Google Apps Script using Google's official command line utility `clasp`.

### Prerequisites
Make sure you have Node.js/NPM installed.

### Commands
1. **Enable Apps Script API**: Turn on the API in your Google account settings: https://script.google.com/home/usersettings
2. **Login to Google**:
   ```bash
   npx @google/clasp login
   ```
3. **Link your Script**:
   ```bash
   npx @google/clasp clone "YOUR_SCRIPT_ID"
   ```
4. **Push local changes to Google**:
   ```bash
   npx @google/clasp push
   ```
5. **Pull cloud changes to local**:
   ```bash
   npx @google/clasp pull
   ```
