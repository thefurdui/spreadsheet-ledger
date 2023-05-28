import { ColumnLetter, VisibleTransactionProperty } from './types'

// Define the spreadsheet ID
export const SPREADSHEET_ID = '17mXT41id8C608abnZv3vgZvMMGN8WF8amWNrdivDc5A'

// Define the ID of the update balance form
export const UPDATE_BALANCE_FORM_ID = '1jotd2iqim96fmh0Hh1XJ-uHSbHhOCdjEXQoTofkkG6o'

// Define the ID of the transactions sheet
export const TRANSACTIONS_SHEET_ID = 1732160294

// Define the field IDs for the form items
export const ACTION_FIELD_ID = 174839173

// Define the anchor columns for income and expense entries
export const EXPENSE_ANCHOR_COLUMN = 'B'
export const INCOME_ANCHOR_COLUMN = 'I'

// Define the categories used for special transaction types
export const REINIT_CATEGORY = '❖ Missing'
export const COMMISSION_CATEGORY = '❖ Commission'

// Define the ID of the balances sheet
export const BALANCES_SHEET_ID = 765911459
// Define the range name for the accounts
export const ACCOUNTS_RANGE_NAME = 'Accounts'
// Define the range name for the expense categories
export const EXPENSE_CATEGORIES_RANGE_NAME = 'ExpenseCategories'
// Define the range name for the income categories
export const INCOME_CATEGORIES_RANGE_NAME = 'IncomeCategories'

// Define the field IDs for the form items
export const ORIGIN_ACCOUNT_FIELD_ID = 1832671484
export const DESTINATION_ACCOUNT_FIELD_ID = 1737254564
export const ACCOUNT_FIELDS_IDS = [ORIGIN_ACCOUNT_FIELD_ID, DESTINATION_ACCOUNT_FIELD_ID]
export const EXPENSE_CATEGORY_FIELD_ID = 1943381995
export const INCOME_CATEGORY_FIELD_ID = 709409363

// Map to convert column letters to data keys for expense transactions
export const columnPropertyMapExpense = new Map<ColumnLetter, VisibleTransactionProperty>([
  ['B', 'date'],
  ['C', 'amount'],
  ['D', 'account'],
  ['E', 'beneficiary'],
  ['F', 'description'],
  ['G', 'category']
])

// Map to convert column letters to data keys for income transactions
export const columnPropertyMapIncome = new Map<ColumnLetter, VisibleTransactionProperty>([
  ['I', 'date'],
  ['J', 'amount'],
  ['K', 'account'],
  ['L', 'beneficiary'],
  ['M', 'description'],
  ['N', 'category']
])
