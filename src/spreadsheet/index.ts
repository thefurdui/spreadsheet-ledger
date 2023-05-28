import { Range, SheetsOnEdit, TransactionType } from '../../types'
import {
  ACCOUNTS_RANGE_NAME,
  ACCOUNT_FIELDS_IDS,
  EXPENSE_CATEGORIES_RANGE_NAME,
  EXPENSE_CATEGORY_FIELD_ID,
  INCOME_CATEGORIES_RANGE_NAME,
  INCOME_CATEGORY_FIELD_ID,
  UPDATE_BALANCE_FORM_ID
} from '../../constants'

// Function to check if two ranges are intersected
const isRangeIntersected = (range1: Range, range2: Range) => {
  if (range1.getLastRow() < range2.getRow()) return false
  if (range2.getLastRow() < range1.getRow()) return false
  if (range1.getLastColumn() < range2.getColumn()) return false
  if (range2.getLastColumn() < range1.getColumn()) return false
  return true
}

// Function to update a select or radio form field with options
const updateFormSelect = (fieldId: number, options: string[], as: 'select' | 'radio') => {
  const uniqueOptions = [...new Set(options)].filter((o) => !!o)
  const form = FormApp.openById(UPDATE_BALANCE_FORM_ID)
  const fieldItem = form.getItemById(fieldId)
  const field = (() => {
    switch (as) {
      case 'select':
        return fieldItem.asListItem()
      case 'radio':
        return fieldItem.asMultipleChoiceItem()
    }
  })()

  Logger.log('Options to be amended to form:')
  Logger.log(uniqueOptions)

  field.setChoices(uniqueOptions.map((o) => field.createChoice(o)))
}

// Function to handle the edit event on the accounts sheet
const onEditAccounts = (e: SheetsOnEdit) => {
  const ss = SpreadsheetApp.getActive()
  const accountsRange = ss.getRangeByName(ACCOUNTS_RANGE_NAME)
  const expenseCategoriesRange = ss.getRangeByName(EXPENSE_CATEGORIES_RANGE_NAME)
  const incomeCategoriesRange = ss.getRangeByName(INCOME_CATEGORIES_RANGE_NAME)

  if (!accountsRange || !expenseCategoriesRange || !incomeCategoriesRange) {
    Logger.log(`Accounts range ${ACCOUNTS_RANGE_NAME} not found`)
    return
  }

  // Check if the edit event intersects with the accounts range and update the form accounts
  if (isRangeIntersected(accountsRange, e.range)) {
    updateFormAccounts()
  }
  // Check if the edit event intersects with the expense categories range and update the form categories for expense
  if (isRangeIntersected(expenseCategoriesRange, e.range)) {
    updateFormCategories('expense')
  }
  // Check if the edit event intersects with the income categories range and update the form categories for income
  if (isRangeIntersected(incomeCategoriesRange, e.range)) {
    updateFormCategories('income')
  }
}

// Function to update the form accounts
const updateFormAccounts = () => {
  const ss = SpreadsheetApp.getActive()
  const accountsRange = ss.getRangeByName(ACCOUNTS_RANGE_NAME)
  const balances = accountsRange?.getValues().flat() as string[]
  ACCOUNT_FIELDS_IDS.forEach((accountField) => updateFormSelect(accountField, balances, 'select'))
}

// Function to update the form categories based on the type (expense or income)
const updateFormCategories = (type: TransactionType) => {
  const ss = SpreadsheetApp.getActive()
  const rangeName =
    type === 'expense' ? EXPENSE_CATEGORIES_RANGE_NAME : INCOME_CATEGORIES_RANGE_NAME
  const fieldId = type === 'expense' ? EXPENSE_CATEGORY_FIELD_ID : INCOME_CATEGORY_FIELD_ID

  const range = ss.getRangeByName(rangeName)

  if (!range) {
    Logger.log(`Range ${rangeName} not found`)
    return
  }

  const categories = range.getValues().flat()
  updateFormSelect(fieldId, categories, 'radio')
}
