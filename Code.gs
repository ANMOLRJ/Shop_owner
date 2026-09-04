
const SPREADSHEET_ID = "1r_7oN9iMuqtqkB3PL_jql1v1pbuqcVAGS-9vIQ9wjpY";

const USER_HEADERS = ["User ID","Name","Email","Phone","Shop ID","Shop Name","Created At","Last Login"];
const PRODUCT_HEADERS = [
  "Timestamp","Shop ID","Shop Name","User ID","Barcode","Product Name","Brand","Category",
  "MRP","Selling Price","Manufacturing Date","Expiry Date","Quantity",
  "Barcode Source","Details Source","OCR Text"
];
const SALES_HEADERS = [
  "Timestamp","Shop ID","Shop Name","User ID","Barcode","Product Name",
  "Selling Price","Quantity","Total","Sale ID"
];

function doGet() {
  return jsonResponse({ok:true, service:"Multi-Shop Barcode Scanner API", version:"7.0"});
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents)
      return jsonResponse({ok:false,error:"No request body received."});
    const d = JSON.parse(e.postData.contents);
    const action = String(d.action || "").trim();

    if (action === "login") return handleLogin_(d);
    if (action === "saveProduct") return handleSaveProduct_(d);
    if (action === "dashboard") return handleDashboard_(d);
    if (action === "deleteProduct") return handleDeleteProduct_(d);
    if (action === "recordSale") return handleRecordSale_(d);

    return jsonResponse({ok:false,error:"Unknown action: "+action});
  } catch(err) {
    return jsonResponse({ok:false,error:String(err && err.message ? err.message : err)});
  }
}

function handleLogin_(d) {
  const email = String(d.email || "").trim().toLowerCase();
  const phone = normalizePhone_(d.phone);
  const name = String(d.name || "").trim();
  const requestedShopName = String(d.shopName || "").trim();
  if (!email && !phone) return jsonResponse({ok:false,error:"Enter email or mobile number."});

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const users = getOrCreateSheet_(ss, "Users", USER_HEADERS);
  const data = users.getDataRange().getValues();

  for (let i=1;i<data.length;i++) {
    const rowEmail = String(data[i][2]||"").trim().toLowerCase();
    const rowPhone = normalizePhone_(data[i][3]);
    if ((email && rowEmail===email) || (phone && rowPhone===phone)) {
      const user = {
        userId:String(data[i][0]||""),
        name:name || String(data[i][1]||"Shop Owner"),
        email:rowEmail,
        phone:rowPhone,
        shopId:String(data[i][4]||""),
        shopName:String(data[i][5]||"")
      };
      if (!user.shopId) user.shopId = makeShopId_();
      if (!user.shopName) user.shopName = requestedShopName || ("Shop "+user.shopId);

      users.getRange(i+1,1,1,USER_HEADERS.length).setValues([[
        user.userId,user.name,user.email,user.phone,user.shopId,user.shopName,
        data[i][6] || new Date(),new Date()
      ]]);
      return jsonResponse({ok:true,action:"login",user:user});
    }
  }

  const userId="USR"+Utilities.getUuid().replace(/-/g,"").substring(0,10).toUpperCase();
  const shopId=makeShopId_();
  const user={
    userId:userId,name:name||"Shop Owner",email:email,phone:phone,
    shopId:shopId,shopName:requestedShopName||("Shop "+shopId)
  };
  users.appendRow([user.userId,user.name,user.email,user.phone,user.shopId,user.shopName,new Date(),new Date()]);
  return jsonResponse({ok:true,action:"created",user:user});
}

function handleSaveProduct_(d) {
  const shopId=String(d.shopId||"").trim(), userId=String(d.userId||"").trim();
  const shopName=String(d.shopName||"").trim(), barcode=String(d.barcode||"").trim();
  if(!shopId||!userId||!barcode) return jsonResponse({ok:false,error:"Shop, user and barcode are required."});

  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet=getOrCreateSheet_(ss,"Products",PRODUCT_HEADERS);
  const row=[
    new Date(),shopId,shopName,userId,barcode,String(d.name||"").trim(),
    String(d.brand||"").trim(),String(d.category||"").trim(),String(d.mrp||"").trim(),
    String(d.sellingPrice||"").trim(),String(d.manufacturingDate||"").trim(),
    String(d.expiryDate||"").trim(),Number(d.quantity||1),
    String(d.barcodeSource||"camera").trim(),String(d.detailsSource||"ocr").trim(),
    String(d.ocrText||"").trim()
  ];

  const last=sheet.getLastRow();
  if(last>=2){
    const vals=sheet.getRange(2,1,last-1,PRODUCT_HEADERS.length).getValues();
    for(let i=0;i<vals.length;i++){
      if(String(vals[i][1]||"").trim()===shopId && String(vals[i][4]||"").trim()===barcode){
        sheet.getRange(i+2,1,1,row.length).setValues([row]);
        return jsonResponse({ok:true,action:"updated",row:i+2,shopId:shopId});
      }
    }
  }
  sheet.appendRow(row);
  return jsonResponse({ok:true,action:"created",row:sheet.getLastRow(),shopId:shopId});
}

function handleDashboard_(d){
  const shopId=String(d.shopId||"").trim();
  if(!shopId) return jsonResponse({ok:false,error:"Shop ID is required."});
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  const ps=getOrCreateSheet_(ss,"Products",PRODUCT_HEADERS);
  const ssheet=getOrCreateSheet_(ss,"Sales",SALES_HEADERS);
  const products=[];
  const pvals=ps.getDataRange().getValues();
  for(let i=1;i<pvals.length;i++){
    if(String(pvals[i][1]||"").trim()===shopId){
      products.push({
        row:i+1,barcode:String(pvals[i][4]||""),
        name:String(pvals[i][5]||""),brand:String(pvals[i][6]||""),
        category:String(pvals[i][7]||""),mrp:String(pvals[i][8]||""),
        sellingPrice:String(pvals[i][9]||""),mfd:String(pvals[i][10]||""),
        exp:String(pvals[i][11]||""),quantity:Number(pvals[i][12]||0)
      });
    }
  }

  let salesCount=0, salesUnits=0, salesAmount=0;
  const svals=ssheet.getDataRange().getValues();
  for(let i=1;i<svals.length;i++){
    if(String(svals[i][1]||"").trim()===shopId){
      salesCount++;
      salesUnits+=Number(svals[i][7]||0);
      salesAmount+=Number(svals[i][8]||0);
    }
  }
  return jsonResponse({
    ok:true,shopId:shopId,
    stats:{products:products.length,salesCount:salesCount,salesUnits:salesUnits,salesAmount:salesAmount},
    products:products
  });
}

function handleDeleteProduct_(d){
  const shopId=String(d.shopId||"").trim(), barcode=String(d.barcode||"").trim();
  if(!shopId||!barcode) return jsonResponse({ok:false,error:"Shop ID and barcode are required."});
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet=getOrCreateSheet_(ss,"Products",PRODUCT_HEADERS);
  const vals=sheet.getDataRange().getValues();
  for(let i=1;i<vals.length;i++){
    if(String(vals[i][1]||"").trim()===shopId && String(vals[i][4]||"").trim()===barcode){
      sheet.deleteRow(i+1);
      return jsonResponse({ok:true,action:"deleted",barcode:barcode});
    }
  }
  return jsonResponse({ok:false,error:"Product not found in this shop."});
}

function handleRecordSale_(d){
  const shopId=String(d.shopId||"").trim(), userId=String(d.userId||"").trim();
  const shopName=String(d.shopName||"").trim(), barcode=String(d.barcode||"").trim();
  const qty=Math.max(1,Number(d.quantity||1));
  if(!shopId||!userId||!barcode) return jsonResponse({ok:false,error:"Shop, user and barcode are required."});

  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  const ps=getOrCreateSheet_(ss,"Products",PRODUCT_HEADERS);
  const vals=ps.getDataRange().getValues();
  for(let i=1;i<vals.length;i++){
    if(String(vals[i][1]||"").trim()===shopId && String(vals[i][4]||"").trim()===barcode){
      const stock=Number(vals[i][12]||0);
      if(stock<qty) return jsonResponse({ok:false,error:"Not enough stock. Available: "+stock});
      const price=Number(String(vals[i][9]||"").replace(/,/g,""))||0;
      ps.getRange(i+1,13).setValue(stock-qty);

      const sales=getOrCreateSheet_(ss,"Sales",SALES_HEADERS);
      const saleId="SALE"+Utilities.getUuid().replace(/-/g,"").substring(0,10).toUpperCase();
      sales.appendRow([
        new Date(),shopId,shopName,userId,barcode,String(vals[i][5]||""),
        price,qty,price*qty,saleId
      ]);
      return jsonResponse({ok:true,action:"sale",saleId:saleId,total:price*qty,remainingStock:stock-qty});
    }
  }
  return jsonResponse({ok:false,error:"Product not found in this shop."});
}

function getOrCreateSheet_(ss,name,headers){
  let sheet=ss.getSheetByName(name);
  if(!sheet) sheet=ss.insertSheet(name);
  if(sheet.getLastRow()===0 || sheet.getLastColumn()<headers.length)
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
  else
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  return sheet;
}

function makeShopId_(){
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet=getOrCreateSheet_(ss,"Users",USER_HEADERS);
  const ids=sheet.getRange(2,5,Math.max(1,sheet.getLastRow()-1),1).getValues();
  let max=0;
  ids.forEach(r=>{const m=String(r[0]||"").match(/^SHOP(\d+)$/);if(m)max=Math.max(max,Number(m[1]));});
  return "SHOP"+String(max+1).padStart(3,"0");
}

function normalizePhone_(phone){
  let p=String(phone||"").replace(/[^\d+]/g,"");
  if(p.startsWith("+91"))p=p.substring(3);
  if(p.startsWith("91")&&p.length===12)p=p.substring(2);
  return p;
}

function jsonResponse(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
