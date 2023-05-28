import {
  FormsOnFormSubmit,
  Transaction,
  TransactionAction,
  TransactionType,
  Sheet,
  TransactionRow,
  ColumnLetter
} from '../../types'
import {
  SPREADSHEET_ID,
  UPDATE_BALANCE_FORM_ID,
  TRANSACTIONS_SHEET_ID,
  ACTION_FIELD_ID,
  EXPENSE_ANCHOR_COLUMN,
  INCOME_ANCHOR_COLUMN,
  REINIT_CATEGORY,
  COMMISSION_CATEGORY,
  columnPropertyMapIncome,
  columnPropertyMapExpense
} from '../../constants'
import {
  getSheetById,
  getColumnLetters,
  getFirstEmptyRowNumber,
  getTransactionDetailsRow,
  getValuesFromRange,
  toCamelCase
} from '../../utils'

// Open the spreadsheet by ID
const ss = SpreadsheetApp.openById(SPREADSHEET_ID)

// Get the transactions sheet by ID
const transactionsSheet = getSheetById(TRANSACTIONS_SHEET_ID, ss)

// Open the form by ID
const form = FormApp.openById(UPDATE_BALANCE_FORM_ID)

// Get the form items by their field IDs
const actionField = form.getItemById(ACTION_FIELD_ID)

// Function to handle form submission
const onFormSubmit = (event: FormsOnFormSubmit) => {
  const { response } = event
  const actionFieldResponse = response.getResponseForItem(actionField).getResponse() as string

  const action = actionFieldResponse.toLowerCase() as TransactionAction

  // Extract form response values and convert titles to camel case
  const responseTransaction = Object.fromEntries(
    response.getItemResponses().map((r) => {
      const title = toCamelCase(r.getItem().getTitle())
      // Unifying "Expense category" and "Income category" fields
      const unifiedTitle = title.includes('Category') ? 'category' : title
      return [unifiedTitle, r.getResponse()]
    })
  ) as unknown as Transaction

  const transaction = {
    ...responseTransaction,
    date: response.getTimestamp()
  }

  // Perform different actions based on the submitted form action
  switch (action) {
    case 'spent':
      appendTransactionRow(transaction, 'expense')
      break
    case 'received':
      appendTransactionRow(transaction, 'income')
      break
    case 'reinitialised':
      handleReinitialisation(transaction)
      break
    case 'transferred':
      handleTransfer(transaction)
      break
  }
}

// Function to handle reinitialisation action
const handleReinitialisation = (transaction: Transaction) => {
  const { account, amount } = transaction

  // Get the range for the accounts and their amounts
  const accountsRange = ss.getRangeByName('Accounts')
  const accountsAmountRange = ss.getRangeByName('AccountsAmount')

  if (!accountsRange || !accountsAmountRange) {
    Logger.log('Accounts or accounts amount range not found')
    return
  }

  // Find the index of the referred account in the accounts range
  const referredAccountIndex = getValuesFromRange(accountsRange, true).findIndex(
    (a) => a === account
  )

  // Get the initial amount for the referred account
  const initialAmount = getValuesFromRange(accountsAmountRange, true)[
    referredAccountIndex
  ] as number

  // Check if the initial amount matches the amount in the form response
  if (initialAmount == amount) {
    return
  }

  // Calculate the difference in amounts
  const diffAmount = Math.abs(initialAmount - amount)
  const isIncome = amount > initialAmount

  // Adjust transaction details with the difference amount and category
  const adjustedTransactionDetails = {
    ...transaction,
    category: REINIT_CATEGORY,
    description: "Account's balance reinitialised",
    amount: diffAmount,
    beneficiary: getBeneficiaryFromAccountName(transaction.account)
  }

  const transactionDetailsRow = getTransactionDetailsRow(adjustedTransactionDetails, isIncome)

  // Append a transaction row based on the difference amount
  appendTransactionRow(transactionDetailsRow, isIncome ? 'income' : 'expense')
}

// Function to append a transaction row to the transactions sheet
const appendTransactionRow = (transaction: Transaction, transactionType: TransactionType) => {
  if (!transactionsSheet) {
    Logger.log('Transactions sheet not found')
    return
  }

  const columnLetters = getColumnLetters(transactionsSheet)
  const columnPropertyMap =
    transactionType == 'income' ? columnPropertyMapIncome : columnPropertyMapExpense
  const anchorColumn = transactionType == 'income' ? INCOME_ANCHOR_COLUMN : EXPENSE_ANCHOR_COLUMN

  // Create a row array based on the transaction details and column mappings
  const row: TransactionRow = columnLetters
    .map((letter) => {
      const columnProperty = columnPropertyMap.get(letter)
      return columnProperty ? transaction[columnProperty] : null
    })
    .filter((cell) => cell !== null)

  // Append the row to the transactions sheet
  appendRow(transactionsSheet, anchorColumn, row)
}

// Function to append row from a specific anchor cell
const appendRow = (sheet: Sheet, anchorColumn: ColumnLetter, row: TransactionRow) => {
  const firstEmptyRowNumber = getFirstEmptyRowNumber(sheet, anchorColumn)

  const columnLetters = getColumnLetters(sheet)
  const indexOfAnchorColumn = columnLetters.indexOf(anchorColumn)
  const lastColumn = columnLetters[indexOfAnchorColumn + row.length - 1]

  if (sheet.getMaxRows() < firstEmptyRowNumber) {
    sheet.appendRow(columnLetters.map(() => ''))
  }

  sheet
    .getRange(`${anchorColumn}${firstEmptyRowNumber}:${lastColumn}${firstEmptyRowNumber}`)
    .setValues([row])
}

// Function to get beneficiary from Account name
const getBeneficiaryFromAccountName = (accountName: string) => {
  const accountFirstLetter = accountName[0]
  if (accountFirstLetter === 'A') {
    return 'Andrei'
  }

  if (accountFirstLetter === 'Y') {
    return 'Yasmin'
  }

  return ''
}

// Function to handle money transfer from one account to another
const handleTransfer = (transaction: Transaction) => {
  const {
    account: originAccount,
    destinationAccount,
    commission,
    destinationCurrencyAmount
  } = transaction

  const originAccountTransactionDetails = {
    ...transaction,
    description: `Transfer to ${destinationAccount}`
  }

  const destinationAccountTransactionDetails = {
    ...transaction,
    amount: destinationCurrencyAmount ?? transaction.amount,
    description: `Transfer from ${originAccount}`,
    isIncome: true
  }

  const commissionTransactionDetails = commission
    ? {
        ...transaction,
        amount: commission,
        description: `Commission for transfer from ${originAccount} to ${destinationAccount}`,
        category: COMMISSION_CATEGORY
      }
    : null

  ;[
    originAccountTransactionDetails,
    destinationAccountTransactionDetails,
    commissionTransactionDetails
  ]
    .filter((td) => td !== null)
    .forEach((td) => {
      const { isIncome } = td as Transaction
      const transactionDetailsRow = getTransactionDetailsRow(td as Transaction, isIncome)
      appendTransactionRow(transactionDetailsRow, isIncome ? 'income' : 'expense')
    })
}