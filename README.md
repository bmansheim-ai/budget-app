Here is the documentation formatted cleanly in Markdown:

---

# Mansheim Budget Processor

The **Mansheim Budget Processor** is a standalone application designed to streamline and automate the monthly reconciliation and categorization of Israeli bank and credit card statements.

---

### Key Features

* **Multi-File Statement Upload**: Upload both bank statements (עובר ושב) and credit card statements simultaneously in a single action.
* **Smart Statement Month Detection**: Automatically detects the statement month and year from uploaded files and displays it in the header interface.
* **Automated Categorization**: Matches statement line items against global categorization rules across 56 standard budget categories.
* **One-Time Manual Overrides**: Easily override any individual transaction's category on the fly without altering global keyword rules.
* **Bank & Credit Card Reconciliation**:
* Automatically calculates beginning and ending bank balances, running totals, and reconciliation differences.
* Filters out credit card payment drafts from bank statements to prevent double-counting and verifies them against credit card statement totals.


* **Month-Based Progress Persistence**: Automatically saves reconciliation progress, uploaded statements, and manual overrides per statement month (`budget_month_YYYY-MM`), allowing seamless switching between past and current months.
* **Static Global Rules**: Categorization keyword rules remain universal across all months, with full support for importing and exporting rules as `.json` files.
* **Excel Integration**: Formats final category totals and generates copy-paste ready formulas for master budgeting spreadsheets.

---

### Build & Run Instructions

To run or package the application locally on a Windows environment:

#### 1. Prerequisites

Ensure **Node.js** is installed on the system.

#### 2. Project Setup & Installation

Open the command prompt or terminal inside the project directory (`C:\Users\Binyomin\budget-app`) and install the required dependencies:

```cmd
npm install

```

#### 3. Run in Development Mode

To test and run the application locally inside a browser during development:

```cmd
npm run dev

```

#### 4. Run as a Desktop Application

To launch the application inside a native desktop window:

```cmd
npm run electron:dev

```

#### 5. Build the Standalone Windows Executable (`.exe`)

To generate a standalone installer and portable `.exe` file:

```cmd
npm run electron:build

```

Once the build process completes, the output files will be available inside the **`dist-electron`** directory.
