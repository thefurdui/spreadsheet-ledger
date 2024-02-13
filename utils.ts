import { columnPropertyMapTransaction } from './constants'
import { Spreadsheet, Range, Sheet, ColumnLetter, Transaction } from './types'

// Function to get a sheet by its ID
export const getSheetById = (id: number, ss?: Spreadsheet) =>
  (ss ?? SpreadsheetApp.getActive()).getSheets().find((s) => s.getSheetId() === id)

// Function to convert a string to camel case
export const toCamelCase = (str: string) =>
  str.toLowerCase().replace(/[^a-zA-Z0-9]+(.)/g, (m, chr) => chr.toUpperCase())

// Function to generate an array of column letters based on the number of columns in the transactions sheet
export const getColumnLetters = (sheet: Sheet) => {
  const lastColumnIndex = sheet.getLastColumn() ?? 0
  const columnLetters = Array.from({ length: lastColumnIndex }, (_, index) =>
    String.fromCharCode(65 + index)
  )
  return columnLetters as Uppercase<string>[]
}

// Function to get values from a range, with an optional filter to remove empty values
export const getValuesFromRange = (range: Range, shouldFilter: boolean) => {
  const values = range.getValues().flat()
  if (!shouldFilter) {
    return values
  }
  return values.filter((v) => v !== '')
}

// Function to get the first empty row number in the transactions sheet
export const getFirstEmptyRowNumber = (sheet: Sheet, anchorColumn: ColumnLetter = 'I') => {
  const referenceColumnValues = sheet
    .getRange(`${anchorColumn}1:${anchorColumn}`)
    .getValues()
    .flat()
  const nonEmptyValues = referenceColumnValues.filter((v) => v !== '')
  const latestNonEmptyRowNumber =
    referenceColumnValues.findIndex((v) => v == nonEmptyValues[nonEmptyValues.length - 1]) + 1
  const desiredEmptyRowNumber = latestNonEmptyRowNumber + 1

  return desiredEmptyRowNumber
}

// Function to filter out spare transaction details and get transaction details row
export const getTransactionDetailsRow = (transaction: Transaction) => {
  const dataMap = columnPropertyMapTransaction
  const transactionDetailsLabels = Array.from(dataMap.values())

  const transactionDetailsEntries = transactionDetailsLabels.map((label) => [
    label,
    transaction[label] ?? ''
  ])
  const transactionDetailsRow = Object.fromEntries(transactionDetailsEntries)

  return transactionDetailsRow
}
