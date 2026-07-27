function onOpen() {

  SpreadsheetApp.getUi()
    .createMenu("📦 Inventory")
    .addItem("Check Low Stock", "checkStock")
    .addToUi();

}

function checkStock() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Products");

  const data = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, 6)
    .getValues();

  let message = "";

  data.forEach(row => {

    const sku = row[0];
    const product = row[1];
    const stock = Number(row[4]);
    const minStock = Number(row[5]);

    if (!sku) return;

    if (stock < minStock) {
      message += `• ${product} (${stock} left)\n`;
    }

  });

  if (message === "") {

    SpreadsheetApp.getUi().alert(
      "✅ All stocks are sufficient."
    );

  } else {

    SpreadsheetApp.getUi().alert(
      "⚠ LOW STOCK\n\n" + message
    );

  }

}

function showTransactionForm() {

  const template =
    HtmlService.createTemplateFromFile("TransactionForm");

  template.products = getProducts();

  const html = template
    .evaluate()
    .setWidth(420)
    .setHeight(420);

  SpreadsheetApp.getUi()
    .showModalDialog(html, "Add Transaction");

}

function getProducts() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Products");

  return sheet
    .getRange(2, 2, sheet.getLastRow() - 1, 1)
    .getValues()
    .flat()
    .filter(String);

}

function saveTransaction(data) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const productSheet =
    ss.getSheetByName("Products");

  const transactionSheet =
    ss.getSheetByName("Transactions");

  const products =
    productSheet.getDataRange().getValues();

  let sku = "";
  let cost = 0;
  let sellPrice = 0;
  let stockRow = 0;
  let currentStock = 0;

  products.forEach((row, index) => {

    if (row[1] === data.product) {

      sku = row[0];
      cost = Number(row[2]);
      sellPrice = Number(row[3]);
      currentStock = Number(row[4]);
      stockRow = index + 1;

    }

  });

  if (!sku) {
    throw new Error("Product not found.");
  }

  const qty = Number(data.qty);

  if (data.type === "Out" && qty > currentStock) {
    throw new Error("Not enough stock.");
  }

  const profit =
    data.type === "Out"
      ? (sellPrice - cost) * qty
      : 0;

  transactionSheet.appendRow([
    new Date(),
    sku,
    data.product,
    data.type,
    qty,
    sellPrice,
    cost,
    profit
  ]);

  if (data.type === "In") {

    productSheet
      .getRange(stockRow, 5)
      .setValue(currentStock + qty);

  } else {

    productSheet
      .getRange(stockRow, 5)
      .setValue(currentStock - qty);

  }

  SpreadsheetApp.flush();

  return "Transaction Saved!";

}
