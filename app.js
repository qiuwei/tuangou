/**
 * A special function that inserts a custom menu when the spreadsheet opens.
 */
function onOpen() {
  var menu = [{name: 'Create order system', functionName: 'setUpOrder_'}];
  SpreadsheetApp.getActive().addMenu('Order', menu);
}


function setUpOrder_() {
  if (ScriptProperties.getProperty('calId')) {
    Browser.msgBox('Your order is already set up. Look in Google Drive!');
  }
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('inventory');
  ss.insertSheet('individual');
  ss.insertSheet('total');
  var range = sheet.getDataRange();
  var values = range.getValues();
  var formName = Browser.inputBox('Input the name for the form');
  //Browser.msgBox(values);
  setUpForm_(ss, values, formName);
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit()
      .create();
  ss.removeMenu('Order');
}



function getProductInventory(values) {
  if (!values) {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName('inventory');
    var range = sheet.getDataRange();
    values = range.getValues();
  }
  var productInventory = {};
  for (var i = 1; i < values.length; i++) {
    var productInfo = values[i];
    var productName = productInfo[0];
    var productOriginalPrice = productInfo[1];
    var productPrice = productInfo[2];
    var productUnit = productInfo[3];
    if(!productInventory[productName]) {
      productInventory[productName] = productInfo;
    }
  }
  return productInventory;
}


function setUpForm_(ss, values, formName) {
  var productInventory = getProductInventory(values);
  // Create the form and add a multiple-choice question for each timeslot.
  var form = FormApp.create(formName);

  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  form.addTextItem().setTitle('Name').setRequired(true);
  form.addTextItem().setTitle('Email').setRequired(true);
  for (var productName in productInventory) {
    var product = productInventory[productName];
    var item = form.addScaleItem().setBounds(0, 5).setTitle( product[0]).setHelpText(product[2] + '/' + product[3]);
  }

  return form;
}

function getOrders(e) {
  var values = SpreadsheetApp.getActive().getSheetByName('inventory')
     .getDataRange().getValues()
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var responseSheet = ss.getSheets();
  for(var i = 0; i < responseSheet.length; i++){
    Logger.log('sheet name: ', responseSheet[i].getName());
  }
  //Logger.log(responseSheet);
  var records = responseSheet[0].getDataRange().getValues();
  Logger.log('records are: ', records);

  var orders = {};
  var totalOrder = {};
  var fieldNames = records[0];
  Logger.log('field names: ', fieldNames);
  for(var i = 1; i < records.length; i++){
    var record = records[i];
    var name = record[1];
    var email = record[2];
    orders[email] = {}; //overwrite existing orders
    for(var j = 1; j < record.length; j++){
      //Logger.log('field names: ', fieldNames[j] );
      orders[email][fieldNames[j]] = record[j];
    }
  }
  //Logger.log('orders:', orders);
  return orders;
}

function calcPrice(order, productInfos) {
  var toPay = 0.0;
  for(key in order){
    var k = key;
    if(k !== 'Timestamp' && k !== 'Name' && k !== 'Email') {
      toPay += order[k] * productInfos[k][2];
    }
  }
  return toPay.toFixed(2);
}

function setupIndividual(orders, productInfos) {
  var individuals = SpreadsheetApp.getActive().getSheetByName('individual');
  individuals.clear();
  for(k in orders) {
    var order = orders[k];
    Logger.log('order: ', orders);
    Logger.log('price: ', calcPrice(order, productInfos));
    var individualResult = [];
    individualResult.push(order['Name']);
    individualResult.push(calcPrice(orders[k], productInfos));
    for(kk in order){
      if(order[kk] > 0 && kk !== 'Name' && kk !== 'Email'){
        individualResult.push(kk + ':' + order[kk]);
      }
    }
    Logger.log('row: ', individualResult);
    individuals.appendRow(individualResult);
  }
}
function setupTotal(orders, productInfos) {
  var totals = SpreadsheetApp.getActive().getSheetByName('total');
  totals.clear();
  var totalPrice = 0;
  var totalOriginPrice = 0;
  for(k in productInfos) {
    var product = productInfos[k];
    var originPrice = product[1];
    var price = product[2];
    var productCount = 0
    for( kk in orders) {
      var order = orders[kk];
      var count = +order[k]
      productCount += count;
    }
    if(productCount > 0){
      totals.appendRow([k, productCount]);
      totalPrice += productCount * price;
      totalOriginPrice += productCount * originPrice;
    }

  }
  totals.appendRow(['总价: ' + totalPrice.toFixed(2), '总进货价: '+ totalOriginPrice.toFixed(2) ]);
}

function updateForms() {
  var orders = getOrders();
  Logger.log('orders: ', orders)
  var productInfos = getProductInventory();
  Logger.log('productInfos: ', productInfos)

  setupIndividual(orders, productInfos);
  setupTotal(orders, productInfos);
}

/**
 * A trigger-driven function that sends out calendar invitations and a
 * personalized Google Docs itinerary after a user responds to the form.
 *
 * @param {Object} e The event parameter for form submission to a spreadsheet;
 *     see https://developers.google.com/apps-script/understanding_events
 */
function onFormSubmit(e) {
  //var formUrl = SpreadsheetApp.getActiveSpreadsheet().getFormUrl();
  //var formID = formUrl.match(/[-\w]{25,}/);
  //var form = FormApp.openById(formID);
  //Logger.log('formID: ', formID);
  //for (var response in form.getResponses()){
  //  Logger.log('response: ', response);
  //}


  var orders = getOrders();

  Logger.log('orders:', orders);
  updateForms();




  //individualResult.push(topay);
  //Logger.log('Indiviuals: ', individualResult);
  //individuals.appendRow(individualResult);

  sendDoc_(e);
}



/**
 * Create and share a personalized Google Doc that shows the user's itinerary.
 *
 * @param {Object} user An object that contains the user's name and email.
 * @param {String[][]} response An array of data for the user's session choices.
 */
function sendDoc_(e) {
  var productInventory = getProductInventory();
  var topay = 0;
  var user = {name: e.namedValues['Name'][0], email: e.namedValues['Email'][0]};
  individualResult = [];

  for ( var productName in productInventory) {
    var product = productInventory[productName];
    var price = product[2];
    var num = e.namedValues[productName][0];
    if ( num !== 0) {
      topay += num*price;
      individualResult.push([productName, num, price])

    }
  }
  var doc = DocumentApp.create(user.name + '的团购订单')
      .addEditor(user.email);
  var body = doc.getBody();
  var table = [['名称', '数量', '单价']];
  for (var i = 0; i < individualResult.length; i++) {
    table.push(individualResult[i]);
  }
  body.insertParagraph(0, doc.getName())
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  table = body.appendTable(table);
  table.getRow(0).editAsText().setBold(true);
  body.appendHorizontalRule();
  body.appendParagraph("总价: "+ topay);
  doc.saveAndClose();

  // Email a link to the Doc as well as a PDF copy.
  MailApp.sendEmail({
    to: user.email,
    subject: doc.getName(),
    body: '谢谢订购，您的订单见附件！',
    attachments: doc.getAs(MimeType.PDF),
  });

  DriveApp.getFileById(doc.getId()).setTrashed(true);
}


