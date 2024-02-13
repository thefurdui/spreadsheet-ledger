/* GoogleAppsScript */
export type Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet
export type Sheet = GoogleAppsScript.Spreadsheet.Sheet
export type Range = GoogleAppsScript.Spreadsheet.Range
export type Date = GoogleAppsScript.Base.Date

export type FormsOnFormSubmit = GoogleAppsScript.Events.FormsOnFormSubmit
export type SheetsOnEdit = GoogleAppsScript.Events.SheetsOnEdit

/* Custom */
export type ColumnLetter = Uppercase<keyof { [K in Uppercase<string>]: true }>

export interface Transaction {
  date: Date
  amount: number
  account: string
  beneficiary?: string
  tag?: string
  expenseCategory?: string
  incomeCategory?: string
  description?: string
  destinationAccount?: string
  commission?: number
  destinationCurrencyAmount?: number
}
export type TransactionProperty = keyof Transaction
export type VisibleTransactionProperty = Exclude<
  TransactionProperty,
  'isIncome' | 'destinationAccount' | 'commission'
>

export type TransactionAction = 'spent' | 'received' | 'reinitialize' | 'transferred'
export type TransactionType = 'expense' | 'income'
export type TransactionRow = (Transaction[keyof Transaction] | null)[]
