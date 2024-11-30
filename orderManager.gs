const responseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Risposte");
const orderSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ordini");
const paramSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Parametri");

const maxPizzasPerOrder = paramSheet.getRange(1,2).getValue();

function createNewOrder(){  
  // cleans the order template
  paramSheet.getRange(17, 1, 50, 3).clear();
  paramSheet.getRange(16, 2).setValue("");
  insertRandomEmoji();
  
  let copiedRows = 0; // counter to limit the pizzas per order
  const lastRow = responseSheet.getLastRow();
  let currOrder;
  let prevOrder = orderSheet.getRange(4,2).getValue();

  // checks if there are new requests to put in a order
  for(let i = 2; i <= lastRow; i++){
    const payed = responseSheet.getRange(i, 5).getValue(); // PAYED flag (E1,E2, E3 ...), must be true to order this pizza
    const orderNumber = responseSheet.getRange(i, 6).getValue(); // Order number (F1,F2, F3 ...), must be empty to order this pizza

    if (payed && orderNumber === "") {
      currOrder = prevOrder + 1;
      // updates the order number in the template
      paramSheet.getRange(16, 2).setValue(currOrder);
      const orderValues = responseSheet.getRange(i, 2, 1, 3).getValues();
      // copies temporary the request values in the order template
      paramSheet.getRange(copiedRows + 17, 1, 1, 3).setValues(orderValues);

      copiedRows++; 

      responseSheet.getRange(i, 6).setValue(currOrder);
      if (copiedRows >= maxPizzasPerOrder) break;
    }
  }
  
  if(copiedRows > 0){
    // copies the compiled order templates into the order sheet
    let orderTemplate = paramSheet.getRange(16, 1, paramSheet.getLastRow() - 5, 7);
    const orderFirstRow = orderSheet.getLastRow() + 2;
    orderTemplate.copyTo(orderSheet.getRange(orderFirstRow, 1), {contentsOnly: true});
    orderSheet.getRange(orderFirstRow + 10, 6).setHorizontalAlignment("center");
    updateCurrentOrderNumber(currOrder); 
  }     
}

function updateCurrentOrderNumber(newOrderNumber){  
  orderSheet.getRange(4,2).setValue(newOrderNumber);
}

function confirmCancelAllOrders(){
const ui = SpreadsheetApp.getUi(); 
  const response = ui.alert(
    'Annullamento ordini', 
    'Sei sicuro di voler annullare tutti gli ordini esistenti?', 
    ui.ButtonSet.YES_NO 
  );

  if (response == ui.Button.YES) {    
    cancelAllOrders();
  } else {    
    ui.alert('Azione interrotta.');
  }
}

function cancelAllOrders(){
  // clears the orders in the order sheet
  orderSheet.getRange(9, 1, orderSheet.getLastRow(), 7).clear({contentsOnly: true});

  // resets the current order number
  updateCurrentOrderNumber(0);

  // clears the order numbers in the response sheet
  responseSheet.getRange(2, 6, responseSheet.getLastRow(), 1).clear();
}

function cancelLastOrder(){
  // finds all the requests with the current order and clear the order number
  const lastRow = responseSheet.getLastRow();
  let currOrder = orderSheet.getRange(4,2).getValue();
  
  for(let i = 2; i <= lastRow; i++){
    let rowOrderNumber = responseSheet.getRange(i,6).getValue();
    if(rowOrderNumber === currOrder)
      responseSheet.getRange(i,6).clear();
  }

  // finds the starting row of the last order in the order sheet
  const lastOrderStartingRow = orderSheet.getRange(9, 2, orderSheet.getLastRow(), 1)
    .getValues()
    .flat()
    .lastIndexOf(currOrder);

  // clears the last order in the order sheet
  orderSheet.getRange(9 + lastOrderStartingRow, 1, orderSheet.getLastRow(), 7).clear({contentsOnly: true});

  // clears the range in the order sheet with the last order (how to find it?) 

  // updates the current order number at the top of the order sheet
  updateCurrentOrderNumber(--currOrder);
}

function fillOrder(){  
  // looks for the total pizzas of the last order
  const values = orderSheet.getRange(9, 5, orderSheet.getLastRow(), 1)
    .getValues()
    .flat();

  let lastNonEmptyIndex = -1;

  for(let i = values.length -1 ; i >= 0; i--){
    if(values[i] != "") { 
      lastNonEmptyIndex = i;
      break;
    }
  }

  if(lastNonEmptyIndex == -1){
    // there are no orders
    const ui = SpreadsheetApp.getUi(); 
    const response = ui.alert(
      'Completamento Ordine', 
      'Impossibile eseguire l\'operazione, non ci ordini inseriti.', 
      ui.ButtonSet.OK
    ); 
  } else {
    const lastOrderTotalPizzas = values[lastNonEmptyIndex];
    if(lastOrderTotalPizzas >= maxPizzasPerOrder){
      // the last orders is already completed with the max pizzas number
      const ui = SpreadsheetApp.getUi(); 
      const response = ui.alert(
        'Completamento Ordine', 
        'Impossibile eseguire l\'operazione, l\'ultimo ordine inserito è già completo.', 
        ui.ButtonSet.OK
      ); 
    } else {
      let copiedRows = lastOrderTotalPizzas; // counter to limit the pizzas per order
      const lastRow = responseSheet.getLastRow();

      // checks if there are new requests to put in a order
      for(let i = 2; i <= lastRow; i++){
        const payed = responseSheet.getRange(i, 5).getValue(); 
        const orderNumber = responseSheet.getRange(i, 6).getValue(); 

        if (payed && orderNumber === "") {            
          const orderValues = responseSheet.getRange(i, 2, 1, 3).getValues();
          // copies temporary the request values in the order template
          paramSheet.getRange(copiedRows + 17, 1, 1, 3).setValues(orderValues);
          copiedRows++;   
          const currOrder = orderSheet.getRange(4,2).getValue(); 
          responseSheet.getRange(i, 6).setValue(currOrder);  
          if (copiedRows >= maxPizzasPerOrder) break;
        }
      }
    }
    let orderTemplate = paramSheet.getRange(16, 1, paramSheet.getLastRow() - 5, 7);
    
    // finds the starting row of the last order in the order sheet
    const lastOrderStartingRow = orderSheet.getRange(9, 1, orderSheet.getLastRow(), 1)
      .getValues()
      .flat()
      .lastIndexOf("Ordine") + 9;
    orderTemplate.copyTo(orderSheet.getRange(lastOrderStartingRow, 1), {contentsOnly: true});
  }
}
  
function insertRandomEmoji(){
  const rndRow = Math.floor(Math.random() * 10) + 1;
  const randomEmoji = paramSheet.getRange(rndRow ,6).getValue();
  paramSheet.getRange(26, 6).setValue(randomEmoji);  
}