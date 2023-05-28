// Define the spreadsheet ID
const SPREADSHEET_ID = "17mXT41id8C608abnZv3vgZvMMGN8WF8amWNrdivDc5A";

// Define the ID of the update balance form
const UPDATE_BALANCE_FORM_ID = "1jotd2iqim96fmh0Hh1XJ-uHSbHhOCdjEXQoTofkkG6o";

// Define the ID of the transactions sheet
const TRANSACTIONS_SHEET_ID = 1732160294;

// Define the field IDs for the form items
const ACTION_FIELD_ID = 174839173;
const EXPENSE_CATEGORY_FIELD_ID = 1943381995;
const INCOME_CATEGORY_FIELD_ID = 709409363;

// Define the anchor columns for income and expense entries
const EXPENSE_ANCHOR_COLUMN = "B";
const INCOME_ANCHOR_COLUMN = "I";

// Define the categories used for special transaction types
const REINIT_CATEGORY = "❖ Missing";
const COMISSION_CATEGORY = "❖ Comission";

// Map to convert column letters to data keys for expense transactions
const data2ColumnExpenseMap = new Map([
  ["B", "date"],
  ["C", "amount"],
  ["D", "account"],
  ["E", "beneficiary"],
  ["F", "description"],
  ["G", "category"],
]);

// Map to convert column letters to data keys for income transactions
const data2ColumnIncomeMap = new Map([
  ["I", "date"],
  ["J", "amount"],
  ["K", "account"],
  ["L", "beneficiary"],
  ["M", "description"],
  ["N", "category"],
]);

// Function to get a sheet by its ID
const getSheetById = (id, ss) =>
  (ss ?? SpreadsheetApp.getActive())
    .getSheets()
    .find((s) => s.getSheetId() === id);

// Function to convert a string to camel case
const toCamelCase = (str) =>
  str.toLowerCase().replace(/[^a-zA-Z0-9]+(.)/g, (m, chr) => chr.toUpperCase());

// Open the spreadsheet by ID
const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

// Get the transactions sheet by ID
const transactionsSheet = getSheetById(TRANSACTIONS_SHEET_ID, ss);

// Open the form by ID
const form = FormApp.openById(UPDATE_BALANCE_FORM_ID);

// Get the form items by their field IDs
const actionField = form.getItemById(ACTION_FIELD_ID);

// Function to generate an array of column letters based on the number of columns in the transactions sheet
const getColumnLetters = () => {
  const lastColumnIndex = transactionsSheet.getLastColumn();
  const columnLetters = [];
  for (let i = 0; i < lastColumnIndex; i++) {
    columnLetters.push(String.fromCharCode(65 + i));
  }
  return columnLetters;
};

// Function to get values from a range, with an optional filter to remove empty values
const getValuesFromRange = (range, shouldFilter) => {
  const values = range.getValues().flat();
  if (!shouldFilter) {
    return values;
  }
  return values.filter((v) => v !== "");
};

// Function to handle form submission
const onFormSubmit = ({ response }) => {
  const action = response
    .getResponseForItem(actionField)
    .getResponse()
    .toLowerCase();

  // Extract form response values and convert titles to camel case
  const transactionDetails = {
    ...Object.fromEntries(
      response.getItemResponses().map((r) => {
        const title = toCamelCase(r.getItem().getTitle());
        // Unifying "Expense category" and "Income category" fields
        const unifiedTitle = title.includes("Category") ? "category" : title;
        return [unifiedTitle, r.getResponse()];
      })
    ),
    date: response.getTimestamp(),
  };

  // Perform different actions based on the submitted form action
  switch (action) {
    case "spent":
      appendTransactionRow(transactionDetails, "expense");
      break;
    case "received":
      appendTransactionRow(transactionDetails, "income");
      break;
    case "reinitialised":
      handleReinitialization(transactionDetails);
      break;
    case "transferred":
      handleTransfer(transactionDetails);
      break;
  }
};

// Function to handle reinitialization action
const handleReinitialization = (transactionDetails) => {
  const { account, amount } = transactionDetails;

  // Get the range for the accounts and their amounts
  const accountsRange = ss.getRangeByName("Accounts");
  const accountsAmountRange = ss.getRangeByName("AccountsAmount");

  // Find the index of the referred account in the accounts range
  const referredAccountIndex = getValuesFromRange(
    accountsRange,
    true
  ).findIndex((a) => a === account);

  // Get the initial amount for the referred account
  const initialAmount = getValuesFromRange(accountsAmountRange, true)[
    referredAccountIndex
  ];

  // Check if the initial amount matches the amount in the form response
  if (initialAmount == amount) {
    return;
  }

  // Calculate the difference in amounts
  const diffAmount = Math.abs(initialAmount - amount);
  const isIncome = amount > initialAmount;

  // Adjust transaction details with the difference amount and category
  const adjustedTransactionDetails = {
    ...transactionDetails,
    category: REINIT_CATEGORY,
    description: "Account's balance reinitialised",
    amount: diffAmount,
    beneficiary: getBeneficiaryFromAccountName(transactionDetails.account),
  };

  const transactionDetailsRow = getTransactionDetailsRow(
    adjustedTransactionDetails,
    isIncome
  );

  // Append a transaction row based on the difference amount
  appendTransactionRow(transactionDetailsRow, isIncome ? "income" : "expense");
};

// Function to append a transaction row to the transactions sheet
const appendTransactionRow = (transactionDetails, transactionType) => {
  const columnLetters = getColumnLetters();
  const data2ColumnMap =
    transactionType == "income" ? data2ColumnIncomeMap : data2ColumnExpenseMap;
  const anchorColumn =
    transactionType == "income" ? INCOME_ANCHOR_COLUMN : EXPENSE_ANCHOR_COLUMN;

  // Create a row array based on the transaction details and column mappings
  const row = columnLetters
    .map((letter) => transactionDetails[data2ColumnMap.get(letter)] ?? null)
    .filter((cell) => cell !== null);

  // Append the row to the transactions sheet
  appendRow(transactionsSheet, anchorColumn, row);
};

// Function to get the first empty row number in the transactions sheet
const getFirstEmptyRowNumber = (anchorColumn = "I") => {
  const referenceColumnValues = transactionsSheet
    .getRange(`${anchorColumn}1:${anchorColumn}`)
    .getValues()
    .flat();
  const nonEmptyValues = referenceColumnValues.filter((v) => v !== "");
  const latestNonEmptyRowNumber =
    referenceColumnValues.findIndex(
      (v) => v == nonEmptyValues[nonEmptyValues.length - 1]
    ) + 1;
  const desiredEmptyRowNumber = latestNonEmptyRowNumber + 1;

  return desiredEmptyRowNumber;
};

// Function to append row from a specific anchor cell
const appendRow = (sheet, anchorColumn, row) => {
  const firstEmptyRowNumber = getFirstEmptyRowNumber(anchorColumn);

  const columnLetters = getColumnLetters();
  const indexOfAnchorColumn = columnLetters.indexOf(anchorColumn);
  const lastColumn = columnLetters[indexOfAnchorColumn + row.length - 1];

  if (sheet.getMaxRows() < firstEmptyRowNumber) {
    sheet.appendRow(columnLetters.map(() => ""));
  }

  sheet
    .getRange(
      `${anchorColumn}${firstEmptyRowNumber}:${lastColumn}${firstEmptyRowNumber}`
    )
    .setValues([row]);
};

// Funtion to get beneficiary from Account name
const getBeneficiaryFromAccountName = (accountName) => {
  const accountFirstLetter = accountName[0];
  if (accountFirstLetter === "A") {
    return "Andrei";
  }

  if (accountFirstLetter === "Y") {
    return "Yasmin";
  }

  return "";
};

// Function to filter out spare transaction details and get transaction details row
const getTransactionDetailsRow = (transactionDetails, isIncome) => {
  const dataMap = isIncome ? data2ColumnIncomeMap : data2ColumnExpenseMap;
  const transactionDetailsLabels = Array.from(dataMap.values());

  const transactionDetailsEntries = transactionDetailsLabels.map((label) => [
    label,
    transactionDetails[label] ?? "",
  ]);
  const transactionDetailsRow = Object.fromEntries(transactionDetailsEntries);

  return transactionDetailsRow;
};

// Function to handle money transfer from one account to another
const handleTransfer = (transactionDetails) => {
  Logger.log(Object.entries(transactionDetails));
  const {
    account: originAccount,
    destinationAccount,
    comission,
    destinationCurrencyAmount,
  } = transactionDetails;

  const originAccountTransactionDetails = {
    ...transactionDetails,
    description: `Transfer to ${destinationAccount}`,
  };

  const destinationAccountTransactionDetails = {
    ...transactionDetails,
    amount: destinationCurrencyAmount ?? transactionDetails.amount,
    description: `Transfer from ${originAccount}`,
    isIncome: true,
  };

  const comissionTransactionDetails = comission
    ? {
        ...transactionDetails,
        amount: comission,
        description: `Comission for transfer from ${originAccount} to ${destinationAccount}`,
        category: COMISSION_CATEGORY,
      }
    : null[
        (originAccountTransactionDetails,
        destinationAccountTransactionDetails,
        comissionTransactionDetails)
      ]
        .filter((td) => td !== null)
        .forEach((td) => {
          const { isIncome } = td;
          const transactionDetailsRow = getTransactionDetailsRow(td, isIncome);
          appendTransactionRow(
            transactionDetailsRow,
            isIncome ? "income" : "expense"
          );
        });
};
