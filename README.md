# 💰 Expense Tracker

A lightweight, Splitwise-style shared-expense tracker built in Google Sheets with Apps Script.
Track who paid, split bills fairly, and see who owes whom — with a one-click plan for the
fewest payments needed to settle everyone up.

<sub>Built by Athreyas, with Claude (Anthropic).</sub>

---

## What it does

- Log shared expenses with four split types: **Equal, Exact, Percentage, Shares**.
- See a live **Dashboard** — current-month spend, all-time totals, and a per-person breakdown.
- Track a single **net balance** per person (Splitwise-style): green = owed to them, red = they owe.
- Get **suggested settlements** — the smallest set of payments that clears every balance.
- Add people, record settlements, and archive old expenses without losing history.

## Install

This is a Google Apps Script project bound to a Google Sheet.

1. Create a new Google Sheet (or open the one you want to use).
2. Open **Extensions → Apps Script**.
3. Recreate these files in the editor and paste in the matching contents from this repo:
   - `Code.gs`
   - `SetupDialog.html`
   - `AddExpenseDialog.html`
   - `AddSettlementDialog.html`
   - `ArchiveDialog.html`

   Add each HTML file via **＋ → HTML** and name it *without* the `.html` extension
   (e.g. `AddExpenseDialog`).
4. **Save**, then reload the spreadsheet — a **💰 Expense Manager** menu appears.
5. Click **💰 Expense Manager → 🔧 Setup**, enter the participant names, and submit.

> Prefer the command line? You can push these files with
> [`clasp`](https://github.com/google/clasp); the included `appsscript.json` is the project manifest.

## How it works (the important part)

The design has one rule: **the Expenses tab is the only source of truth.** Everything else is
derived, so nothing can drift out of sync.

```
You add an expense ─▶ Expenses sheet (raw record)
                         │
                         ▼
                     Ledger sheet  (auto-rebuilt: one row per person per txn — Paid vs Share)
                         │
                         ▼
        Dashboard + Balances  (100% live formulas: net = Paid − Share)
```

Because balances are formulas reading a rebuilt ledger — never hand-written numbers — a single
**Refresh Balances** fully recomputes the whole system from your expenses. It is self-healing.

## The sheets

| Sheet | Purpose |
|---|---|
| **About** | Quick-start guide and split-type reference |
| **Dashboard** | Month summary, quick stats, spending by person |
| **Balances** | Net balance per person + suggested "who pays whom" |
| **Expenses** | Every transaction (the source of truth) |
| *Master Users* | People list (hidden) |
| *Ledger* | Auto-generated math cache — do not edit (hidden) |
| *Archives* | Older expenses moved out of the way (hidden) |

## Daily use

All actions live under the **💰 Expense Manager** menu:

- **Add Expense** — date, category, description, amount, split type, payer, and who it's split among.
  Defaults to *paid by the first person, split equally among everyone* (the payer is auto-included).
- **Add Settlement** — record a payment from one person to another; balances update automatically.
- **Refresh Balances** — fully recalculates everything from the Expenses sheet.
- **Add New User** — adds a person and formats them into every sheet.
- **Archive Old Expenses** — moves older entries to the Archives sheet to keep things fast.

## A live dashboard with Canvas (optional, display-only)

Google Sheets' **Canvas** (Gemini) can render a clean, real-time dashboard on top of this data.
Treat it purely as a **view** — the spreadsheet and Apps Script stay the system of record.

Canvas reads a **single tab**, so run this on the **Expenses** sheet (the one tab that holds every
raw transaction). Gemini derives the per-person balances and suggested settlements from the split
details; those derived figures are for at-a-glance display, while the authoritative balances live
on the locked **Dashboard** and **Balances** sheets computed by the script. Keep adding and settling
through the **💰 Expense Manager** menu so the math stays exact.

**To use it:** open the **Expenses** tab → **Insert → Create a canvas** → paste the prompt below.

```text
Build a clean, read-only dashboard titled "Group Expenses Summary" from this Expenses sheet.
Keep fonts and cards modest and compact — not oversized. Directly under the title, write a
short one-line summary of the group: the number of distinct people involved (derived from the
Paid By and Split Details columns) and the total number of expenses.

Below that, show four small summary cards:
  1. Current Month Expenses — total Amount of rows where Type is "Expense" dated in the current month.
  2. Month Transactions — the count of those current-month expenses.
  3. Settled Transactions — the count of rows where Type is "Settlement".
  4. Unsettled Amount — the money still outstanding across the group (the sum of all positive
     net balances; see how to compute net balance below).

Then show a "Spending by Person" table with columns Person, Total Spent, Transactions, Net
Balance, and Status, next to a bar chart of each person's Total Spent. Compute the values like so:
  - Total Spent = sum of Amount for expense rows where that person is the Paid By.
  - Transactions = count of those rows.
  - Net Balance = (total that person paid) minus (that person's fair share of every expense).
    A person's fair share of an expense is its Amount split equally among the names listed in
    Split Details ("Split equally among [A, B, C]"), or the explicit amount / percentage / share
    when one is given. Settlement rows ("X paid Y") move money from X to Y and reduce balances.
  - Show positive Net Balance in green and negative in red.
  - Status = "Owed $X" when positive, "Owes $X" when negative, "Even" when zero.

Below that, show a "Recent Activity" table that includes BOTH expenses and settlements (every
row regardless of Type), with columns Date, Type, Category, Person (Paid By), and Amount, sorted
by Date descending, showing the latest 5–6 rows. Label each row's Type as "Expense" or
"Settlement"; for settlement rows, use the Split Details ("X paid Y") in place of a category so
it is clear who paid whom. Style settlement rows subtly differently (for example a small badge or
muted tint) so they are easy to tell apart from expenses.

Just above the footer, add a highlighted "Suggested Settlements" callout listing who should pay
whom and how much, using the fewest possible payments: take each person's Net Balance, then have
those who owe (negative) pay those who are owed (positive) — repeatedly match the largest debtor
to the largest creditor for the amount that fully clears one of them — and list each payment as
"Name → Name  $Amount". If everyone is even, show "All settled up — no payments needed." Add a
small sub-line: "Record a payment via 💰 Expense Manager → 💳 Add Settlement."

Finish with a small, subtle footer line in muted gray: "Expense Tracker • built by Athreyas."
Use a soft, modern look with rounded cards and a clean compact UI with small fonts, and offer
both light and dark mode. This is display only — do not add any data-entry fields, buttons, or
editing controls.
```

Follow-up tweaks you can layer on with another prompt: *"Switch to dark mode,"* *"Make the cards
more compact,"* or *"Show the latest 10 transactions instead of 6."*

> Canvas is experimental and only reads one tab, so the balances it shows are Gemini's
> interpretation of the Expenses data. For the exact, tested figures, rely on the Dashboard and
> Balances sheets.


## Notes

- Sheets are protected; make changes through the menu rather than editing cells directly.
- A correct expense splitter always nets to **$0.00** across everyone — that's the built-in sanity check.
- Setup **erases all data**; export your Expenses tab first if you ever re-run it.
