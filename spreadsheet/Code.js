// Define the spreadsheet ID
const SPREADSHEET_ID = "17mXT41id8C608abnZv3vgZvMMGN8WF8amWNrdivDc5A";
// Define the ID of the balances sheet
const BALANCES_SHEET_ID = 765911459;
// Define the range name for the accounts
const ACCOUNTS_RANGE_NAME = "Accounts";
// Define the range name for the expense categories
const EXPENSE_CATEGORIES_RANGE_NAME = "ExpenseCategories";
// Define the range name for the income categories
const INCOME_CATEGORIES_RANGE_NAME = "IncomeCategories";

// Define the ID of the update balance form
const UPDATE_BALANCE_FORM_ID = "1jotd2iqim96fmh0Hh1XJ-uHSbHhOCdjEXQoTofkkG6o";
// Define the field IDs for the form items
const ORIGIN_ACCOUNT_FIELD_ID = 1832671484;
const DESTINATION_ACCOUNT_FIELD_ID = 1737254564;
const ACCOUNT_FIELDS_IDS = [
  ORIGIN_ACCOUNT_FIELD_ID,
  DESTINATION_ACCOUNT_FIELD_ID,
];
const EXPENSE_CATEGORY_FIELD_ID = 1943381995;
const INCOME_CATEGORY_FIELD_ID = 709409363;

// Function to get a sheet by its ID
const getSheetById = (id, ss) =>
  (ss ?? SpreadsheetApp.getActive())
    .getSheets()
    .find((s) => s.getSheetId() === id);

// Function to check if two ranges are intersected
const isRangeIntersected = (range1, range2) => {
  if (range1.getLastRow() < range2.getRow()) return false;
  if (range2.getLastRow() < range1.getRow()) return false;
  if (range1.getLastColumn() < range2.getColumn()) return false;
  if (range2.getLastColumn() < range1.getColumn()) return false;
  return true;
};

// Function to update a select or radio form field with options
const updateFormSelect = (fieldId, options, as) => {
  const uniqueOptions = [...new Set(options)].filter((o) => !!o || o === 0);
  const form = FormApp.openById(UPDATE_BALANCE_FORM_ID);
  const fieldItem = form.getItemById(fieldId);
  const field = (() => {
    switch (as) {
      case "select":
        return fieldItem.asListItem();
      case "radio":
        return fieldItem.asMultipleChoiceItem();
      default:
        return fieldItem.asMultipleChoiceItem();
    }
  })();

  Logger.log("Options to be ammended to form:");
  Logger.log(uniqueOptions);

  field.setChoices(uniqueOptions.map((o) => field.createChoice(o)));
};

// Function to handle the edit event on the accounts sheet
const onEditAccounts = (e) => {
  const ss = SpreadsheetApp.getActive();
  const accountsRange = ss.getRangeByName(ACCOUNTS_RANGE_NAME);
  const expenseCategoriesRange = ss.getRangeByName(
    EXPENSE_CATEGORIES_RANGE_NAME
  );
  const incomeCategoriesRange = ss.getRangeByName(INCOME_CATEGORIES_RANGE_NAME);

  // Check if the edit event intersects with the accounts range and update the form accounts
  if (isRangeIntersected(accountsRange, e.range)) {
    updateFormAccounts();
  }
  // Check if the edit event intersects with the expense categories range and update the form categories for expense
  if (isRangeIntersected(expenseCategoriesRange, e.range)) {
    updateFormCategories("expense");
  }
  // Check if the edit event intersects with the income categories range and update the form categories for income
  if (isRangeIntersected(incomeCategoriesRange, e.range)) {
    updateFormCategories("income");
  }
};

// Function to update the form accounts
const updateFormAccounts = () => {
  const ss = SpreadsheetApp.getActive();
  const accountsRange = ss.getRangeByName(ACCOUNTS_RANGE_NAME);
  const balances = accountsRange.getValues().flat();
  ACCOUNT_FIELDS_IDS.forEach((accountField) =>
    updateFormSelect(accountField, balances, "select")
  );
};

// Function to update the form categories based on the type (expense or income)
const updateFormCategories = (type) => {
  const ss = SpreadsheetApp.getActive();
  const rangeName =
    type === "expense"
      ? EXPENSE_CATEGORIES_RANGE_NAME
      : INCOME_CATEGORIES_RANGE_NAME;
  const fieldId =
    type === "expense" ? EXPENSE_CATEGORY_FIELD_ID : INCOME_CATEGORY_FIELD_ID;

  const range = ss.getRangeByName(rangeName);
  const categories = range.getValues().flat();
  updateFormSelect(fieldId, categories, "radio");
};
