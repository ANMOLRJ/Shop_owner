const SPREADSHEET_ID = "1uwBIIhB3E-AhWJ_lCtzB_n8ygQzBfOJzNQWq4K4jCoI";

const USER_HEADERS = ["User ID","Name","Email","Phone","Shop ID","Shop Name","Created At","Last Login","Firebase UID","Exit Count"];
const PRODUCT_HEADERS = ["Timestamp","Shop ID","Shop Name","User ID","Barcode","Product Name","Brand","Category","MRP","Selling Price","Manufacturing Date","Expiry Date","Quantity","RFID UID","RFID Status","Barcode Source","Details Source","OCR Text"];
const ASSIGN_HEADERS = ["Timestamp","Assignment ID","Shop ID","Shop Name","User ID","Barcode","Product Name","Status","RFID UID","Assigned At"];
const RFID_HEADERS = ["Timestamp","RFID UID","Shop ID","Shop Name","Barcode","Product Name","Status","Assigned At","Exit Event ID","Last Updated"];
const SALES_HEADERS = ["Timestamp","Shop ID","Shop Name","User ID","Barcode","Product Name","Selling Price","Quantity","Total","Sale ID"];
const EXIT_HEADERS = ["Timestamp","Exit Event ID","Shop ID","Shop Name","Exit No","RFID UID","Barcode","Product Name","Selling Price","Status","Sale ID"];

function doGet(e){
  try{
    const p=(e&&e.parameter)||{}; const action=String(p.action||"health");
    if(action==="rfidPending") return handleRfidPending_(p);
    if(action==="exitEvents") return handleExitEvents_(p);
    if(action==="dashboard") return handleDashboard_(p);
    if(action==="loginFirebase") return handleFirebaseLogin_(p);
    if(action==="health") return jsonResponse({ok:true,service:"Shop Barcode OCR RFID API",version:"25.0-CHECKOUT-DASHBOARD-FIXED"});
    return jsonResponse({ok:false,error:"Unknown GET action: "+action});
  }catch(err){return jsonResponse({ok:false,error:String(err&&err.message||err)});}
}
function doPost(e){
  try{
    let d={};
    const p=(e&&e.parameter)||{};
    const raw=(e&&e.postData&&e.postData.contents)?String(e.postData.contents).trim():"";
    const contentType=String((e&&e.postData&&e.postData.type)||"").toLowerCase();

    // Accept JSON bodies regardless of the browser's Content-Type, plus normal form POSTs.
    // This prevents the dashboard from becoming "Unknown action:" when a browser sends
    // JSON with text/plain (a common CORS-simple request from GitHub Pages).
    if(raw){
      const looksJson = raw.charAt(0)==="{" || raw.charAt(0)==="[";
      if(looksJson || contentType.indexOf("application/json")!==-1){
        try{ d=JSON.parse(raw); }catch(jsonErr){ d={}; }
      }
    }

    // Apps Script exposes form/query parameters here. They also safely fill
    // in any fields missing from a JSON request.
    Object.keys(p).forEach(k=>{if(d[k]===undefined)d[k]=p[k];});

    // Form posts stringify arrays/objects. Convert them back before routing.
    ["eventIds"].forEach(function(k){
      if(typeof d[k]==="string"){
        const s=d[k].trim();
        if(s && (s.charAt(0)==="[" || s.charAt(0)==="{")){
          try{ d[k]=JSON.parse(s); }catch(parseErr){}
        }
      }
    });

    const action=String(d.action||"").trim();
    if(action==="loginFirebase") return handleFirebaseLogin_(d);
    if(action==="saveProduct") return handleSaveProduct_(d);
    if(action==="dashboard") return handleDashboard_(d);
    if(action==="deleteProduct") return handleDeleteProduct_(d);
    if(action==="recordSale") return handleRecordSale_(d);
    if(action==="assignRFID") return handleAssignRFID_(d);
    if(action==="exitDetect") return handleExitDetect_(d);
    if(action==="removeExitEvent") return handleRemoveExitEvent_(d);
    if(action==="payExitEvents") return handlePayExitEvents_(d);
    if(action==="updateExitCount") return handleUpdateExitCount_(d);
    return jsonResponse({ok:false,error:"Unknown action: "+action});
  }catch(err){return jsonResponse({ok:false,error:String(err&&err.message||err)});}
}

function handleFirebaseLogin_(d){
  const phone=normalizePhone_(d.phone), email=String(d.email||"").trim().toLowerCase(), name=String(d.name||"").trim(), shopName=String(d.shopName||"").trim(), uid=String(d.firebaseUid||"").trim();
  if(!uid||!phone) return jsonResponse({ok:false,error:"Verified Firebase UID and phone are required."});
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID), sh=getOrCreateSheet_(ss,"Users",USER_HEADERS), vals=sh.getDataRange().getValues();
  for(let i=1;i<vals.length;i++){
    if(String(vals[i][8]||"")===uid || normalizePhone_(vals[i][3])===phone){
      const user={userId:String(vals[i][0]||""),name:name||String(vals[i][1]||"Shop Owner"),email:email||String(vals[i][2]||""),phone:phone,shopId:String(vals[i][4]||""),shopName:String(vals[i][5]||"")||shopName||"My Shop",exitCount:Math.max(1,Number(vals[i][9]||1))};
      sh.getRange(i+1,1,1,USER_HEADERS.length).setValues([[user.userId,user.name,user.email,user.phone,user.shopId,user.shopName,vals[i][6]||new Date(),new Date(),uid,user.exitCount]]);
      return jsonResponse({ok:true,action:"login",user:user});
    }
  }
  const user={userId:"USR"+Utilities.getUuid().replace(/-/g,"").substring(0,10).toUpperCase(),name:name||"Shop Owner",email:email,phone:phone,shopId:makeShopId_(),shopName:shopName||"New Shop",exitCount:1};
  sh.appendRow([user.userId,user.name,user.email,user.phone,user.shopId,user.shopName,new Date(),new Date(),uid,user.exitCount]);
  return jsonResponse({ok:true,action:"created",user:user});
}

function handleSaveProduct_(d){
  const shopId=String(d.shopId||"").trim(), shopName=String(d.shopName||"").trim(), userId=String(d.userId||"").trim(), barcode=String(d.barcode||"").trim();
  const requestedQty=Math.max(1,Math.floor(Number(d.quantity||1)));
  if(!shopId||!userId||!barcode)return jsonResponse({ok:false,error:"Shop, user and barcode are required."});
  const lock=LockService.getScriptLock();
  try{
    lock.waitLock(10000);
    const ss=SpreadsheetApp.openById(SPREADSHEET_ID), sh=getOrCreateSheet_(ss,"Products",PRODUCT_HEADERS);
    const row=[new Date(),shopId,shopName,userId,barcode,String(d.name||"").trim(),String(d.brand||"").trim(),String(d.category||"").trim(),String(d.mrp||"").trim(),String(d.sellingPrice||"").trim(),String(d.manufacturingDate||"").trim(),String(d.expiryDate||"").trim(),requestedQty,"","PENDING",String(d.barcodeSource||"camera"),String(d.detailsSource||"ocr"),String(d.ocrText||"").trim()];
    const vals=sh.getDataRange().getValues();
    for(let i=1;i<vals.length;i++){
      if(String(vals[i][1]||"")===shopId&&String(vals[i][4]||"")===barcode){
        const oldQty=Math.max(0,Number(vals[i][12]||0));
        const inv=getOrCreateSheet_(ss,"RFID_Inventory",RFID_HEADERS);
        migrateLegacyRfid_(ss,vals[i]);
        const activeCount=countActiveRfids_(inv,shopId,barcode);
        if(requestedQty<activeCount)return jsonResponse({ok:false,error:"Quantity cannot be lower than the number of active RFID tags already assigned ("+activeCount+"). Remove/sell those units first."});
        row[13]=String(vals[i][13]||"");
        row[14]=activeCount>=requestedQty?"ASSIGNED":"PENDING";
        sh.getRange(i+1,1,1,row.length).setValues([row]);
        ensurePendingAssignments_(ss,shopId,shopName,userId,barcode,String(row[5]||""),requestedQty);
        return jsonResponse({ok:true,action:"updated",row:i+1,rfidRequired:activeCount<requestedQty,quantity:requestedQty,assignedRfids:activeCount,remainingRfids:Math.max(0,requestedQty-activeCount),previousQuantity:oldQty});
      }
    }
    sh.appendRow(row);
    ensurePendingAssignments_(ss,shopId,shopName,userId,barcode,String(row[5]||""),requestedQty);
    return jsonResponse({ok:true,action:"created",row:sh.getLastRow(),rfidRequired:true,quantity:requestedQty,assignedRfids:0,remainingRfids:requestedQty});
  }catch(err){return jsonResponse({ok:false,error:String(err&&err.message||err)});}
  finally{try{lock.releaseLock();}catch(e){}}
}
function migrateLegacyRfid_(ss,productRow){
  const rfid=normalizeRfid_(productRow[13]); if(!rfid)return;
  const shopId=String(productRow[1]||"").trim(),shopName=String(productRow[2]||"").trim(),barcode=String(productRow[4]||"").trim(),name=String(productRow[5]||"").trim();
  const inv=getOrCreateSheet_(ss,"RFID_Inventory",RFID_HEADERS),v=inv.getDataRange().getValues();
  for(let i=1;i<v.length;i++)if(String(v[i][1]||"")===rfid)return;
  inv.appendRow([new Date(),rfid,shopId,shopName,barcode,name,"ACTIVE",new Date(),"",new Date()]);
}
function countActiveRfids_(inv,shopId,barcode){
  const v=inv.getDataRange().getValues(); let n=0;
  for(let i=1;i<v.length;i++)if(String(v[i][2]||"")===shopId&&String(v[i][4]||"")===barcode&&String(v[i][6]||"")==="ACTIVE")n++;
  return n;
}
function countAssignedRfids_(ss,shopId,barcode){
  const inv=getOrCreateSheet_(ss,"RFID_Inventory",RFID_HEADERS); return countActiveRfids_(inv,shopId,barcode);
}
function ensurePendingAssignments_(ss,shopId,shopName,userId,barcode,productName,quantity){
  const a=getOrCreateSheet_(ss,"RFID_Assignments",ASSIGN_HEADERS);
  const active=countAssignedRfids_(ss,shopId,barcode);
  const targetPending=Math.max(0,Math.floor(quantity)-active);

  let v=a.getDataRange().getValues();
  let pending=0;
  for(let i=1;i<v.length;i++){
    if(String(v[i][2]||"")===shopId&&String(v[i][5]||"")===barcode&&String(v[i][7]||"")==="PENDING") pending++;
  }

  if(pending>targetPending){
    let remove=pending-targetPending;
    for(let i=v.length-1;i>=1&&remove>0;i--){
      if(String(v[i][2]||"")===shopId&&String(v[i][5]||"")===barcode&&String(v[i][7]||"")==="PENDING"){
        a.deleteRow(i+1);
        remove--;
      }
    }
  }

  // Re-read after any deletes, then create every missing assignment.
  v=a.getDataRange().getValues();
  pending=0;
  for(let i=1;i<v.length;i++){
    if(String(v[i][2]||"")===shopId&&String(v[i][5]||"")===barcode&&String(v[i][7]||"")==="PENDING") pending++;
  }

  while(pending<targetPending){
    const id="ASN"+Utilities.getUuid().replace(/-/g,"").substring(0,10).toUpperCase();
    a.appendRow([new Date(),id,shopId,shopName,userId,barcode,productName,"PENDING","",""]);
    pending++;
  }
}

function handleRfidPending_(p){
  const shopName=String(p.shopName||"").trim(), shopId=String(p.shopId||"").trim();
  if(!shopName&&!shopId)return jsonResponse({ok:false,error:"shopName or shopId required"});

  // IMPORTANT: RFID polling is a frequent READ operation. Do not hold the
  // global script lock for the entire poll. Older versions did this and a
  // slow Google Sheets request could block the next poll for 10+ seconds.
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  const ps=getOrCreateSheet_(ss,"Products",PRODUCT_HEADERS);
  let pv=ps.getDataRange().getValues();
  const a=getOrCreateSheet_(ss,"RFID_Assignments",ASSIGN_HEADERS);

  // First look for an existing pending assignment. This path is READ-ONLY
  // and therefore remains fast even if another write request is running.
  let vals=a.getDataRange().getValues();
  let target=-1;
  for(let i=1;i<vals.length;i++){
    if(String(vals[i][7]||"")==="PENDING" &&
       ((shopId&&String(vals[i][2]||"")===shopId)||(shopName&&String(vals[i][3]||"")===shopName))){
      target=i;break;
    }
  }

  // If no queue item exists, repair ONLY the first incomplete product and
  // use a short lock. Never wait 10 seconds just to poll.
  if(target<0){
    for(let j=1;j<pv.length;j++){
      const sameShop=(shopId&&String(pv[j][1]||"")===shopId)||(shopName&&String(pv[j][2]||"")===shopName);
      if(!sameShop)continue;
      const sid=String(pv[j][1]||shopId), sname=String(pv[j][2]||shopName);
      const barcode=String(pv[j][4]||"");
      const qty=Math.max(0,Math.floor(Number(pv[j][12]||0)));
      if(!barcode||qty<=0)continue;
      const active=countAssignedRfids_(ss,sid,barcode);
      const pending=countPendingAssignments_(a,sid,barcode);
      const needed=Math.max(0,qty-active);
      if(needed>pending){
        const lock=LockService.getScriptLock();
        if(lock.tryLock(1500)){
          try{
            // Re-check after obtaining the lock because another request may
            // have repaired the queue while we were reading.
            const active2=countAssignedRfids_(ss,sid,barcode);
            const pending2=countPendingAssignments_(a,sid,barcode);
            const needed2=Math.max(0,qty-active2);
            if(needed2>pending2){
              ensurePendingAssignments_(ss,sid,sname,String(pv[j][3]||""),barcode,String(pv[j][5]||""),qty);
              ps.getRange(j+1,15).setValue(active2>=qty?"ASSIGNED":"PENDING");
            }
          }finally{try{lock.releaseLock();}catch(e){}}
        }
        break;
      }
    }
    vals=a.getDataRange().getValues();
    for(let i=1;i<vals.length;i++){
      if(String(vals[i][7]||"")==="PENDING" &&
         ((shopId&&String(vals[i][2]||"")===shopId)||(shopName&&String(vals[i][3]||"")===shopName))){
        target=i;break;
      }
    }
  }

  if(target<0){
    let lastQty=0,lastAssigned=0;
    pv=ps.getDataRange().getValues();
    for(let j=1;j<pv.length;j++){
      const sameShop=(shopId&&String(pv[j][1]||"")===shopId)||(shopName&&String(pv[j][2]||"")===shopName);
      if(!sameShop)continue;
      const q=Math.max(0,Number(pv[j][12]||0)), b=String(pv[j][4]||"");
      const ac=countAssignedRfids_(ss,String(pv[j][1]||shopId),b);
      if(ac>0||q>0){lastQty=q;lastAssigned=ac;}
    }
    return jsonResponse({ok:true,pending:false,quantity:lastQty,assignedCount:lastAssigned,remainingCount:Math.max(0,lastQty-lastAssigned)});
  }

  const row=vals[target];
  const productBarcode=String(row[5]||"");
  let quantity=0;
  const freshProducts=ps.getDataRange().getValues();
  for(let j=1;j<freshProducts.length;j++){
    if(String(freshProducts[j][1]||"")===String(row[2]||shopId)&&String(freshProducts[j][4]||"")===productBarcode){
      quantity=Number(freshProducts[j][12]||0);break;
    }
  }
  const active=countAssignedRfids_(ss,String(row[2]||shopId),productBarcode);
  const pendingForProduct=countPendingAssignments_(a,String(row[2]||shopId),productBarcode);

  return jsonResponse({
    ok:true,pending:true,assignmentId:String(row[1]),
    shopId:String(row[2]),shopName:String(row[3]),barcode:productBarcode,
    productName:String(row[6]),quantity:quantity,
    assignedCount:active,remainingCount:Math.max(0,quantity-active),
    pendingCount:pendingForProduct
  });
}

function countPendingAssignments_(a,shopId,barcode){const v=a.getDataRange().getValues();let n=0;for(let i=1;i<v.length;i++)if(String(v[i][2]||"")===shopId&&String(v[i][5]||"")===barcode&&String(v[i][7]||"")==="PENDING")n++;return n;}
function handleAssignRFID_(d){
  const shopName=String(d.shopName||"").trim(), shopId=String(d.shopId||"").trim(), aid=String(d.assignmentId||"").trim(), rfid=normalizeRfid_(d.rfidUid);
  if(!rfid||(!shopName&&!shopId))return jsonResponse({ok:false,error:"shopName/shopId and rfidUid required"});
  const lock=LockService.getScriptLock();
  try{
    lock.waitLock(10000);
    const ss=SpreadsheetApp.openById(SPREADSHEET_ID), a=getOrCreateSheet_(ss,"RFID_Assignments",ASSIGN_HEADERS), vals=a.getDataRange().getValues(); let target=-1;
    for(let i=1;i<vals.length;i++)if(String(vals[i][7]||"")==="PENDING"&&((aid&&String(vals[i][1]||"")===aid)||(!aid&&((shopName&&String(vals[i][3]||"")===shopName)||(shopId&&String(vals[i][2]||"")===shopId))))){target=i;break;}
    if(target<0)return jsonResponse({ok:false,error:"No pending product assignment found for this shop."});
    const targetShopId=String(vals[target][2]||shopId), barcode=String(vals[target][5]||""), productName=String(vals[target][6]||"");
    const ps=getOrCreateSheet_(ss,"Products",PRODUCT_HEADERS),pvals=ps.getDataRange().getValues(); let productRow=-1,quantity=0;
    for(let i=1;i<pvals.length;i++)if(String(pvals[i][1]||"")===targetShopId&&String(pvals[i][4]||"")===barcode){productRow=i+1;quantity=Math.max(0,Number(pvals[i][12]||0));migrateLegacyRfid_(ss,pvals[i]);break;}
    if(productRow<0)return jsonResponse({ok:false,error:"Product not found for this RFID assignment."});
    const inv=getOrCreateSheet_(ss,"RFID_Inventory",RFID_HEADERS),iv=inv.getDataRange().getValues();
    for(let i=1;i<iv.length;i++)if(String(iv[i][1]||"")===rfid)return jsonResponse({ok:false,error:"This RFID tag is already assigned/used. Scan a different RFID tag.",rfidUid:rfid});
    const active=countActiveRfids_(inv,targetShopId,barcode);
    if(active>=quantity)return jsonResponse({ok:false,error:"All RFID tags required for this product are already assigned ("+active+"/"+quantity+").",assignedCount:active,quantity:quantity});
    a.getRange(target+1,8,1,3).setValues([["ASSIGNED",rfid,new Date()]]);
    inv.appendRow([new Date(),rfid,targetShopId,String(vals[target][3]||shopName),barcode,productName,"ACTIVE",new Date(),"",new Date()]);
    if(productRow>0){
      const firstRfid=String(pvals[productRow-1][13]||"");
      const assignedStatus=(active+1)>=quantity?"ASSIGNED":"PENDING";
      if(!firstRfid)ps.getRange(productRow,14,1,2).setValues([[rfid,assignedStatus]]); else ps.getRange(productRow,15).setValue(assignedStatus);
    }
    const assignedNow=active+1;
    ensurePendingAssignments_(ss,targetShopId,String(vals[target][3]||shopName),String(vals[target][4]||""),barcode,productName,quantity);
    const remaining=Math.max(0,quantity-assignedNow);
    return jsonResponse({ok:true,action:"assigned",assignmentId:String(vals[target][1]),barcode:barcode,rfidUid:rfid,assignedCount:assignedNow,quantity:quantity,remainingCount:remaining,complete:remaining===0});
  }catch(err){return jsonResponse({ok:false,error:String(err&&err.message||err)});}
  finally{try{lock.releaseLock();}catch(e){}}
}

function handleDashboard_(d){
  const shopId=String(d.shopId||"").trim(); if(!shopId)return jsonResponse({ok:false,error:"Shop ID required"});
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID), ps=getOrCreateSheet_(ss,"Products",PRODUCT_HEADERS), products=[]; const pv=ps.getDataRange().getValues();
  for(let i=1;i<pv.length;i++) if(String(pv[i][1]||"")===shopId) products.push({row:i+1,barcode:String(pv[i][4]||""),name:String(pv[i][5]||""),brand:String(pv[i][6]||""),category:String(pv[i][7]||""),mrp:String(pv[i][8]||""),sellingPrice:String(pv[i][9]||""),mfd:String(pv[i][10]||""),exp:String(pv[i][11]||""),quantity:Number(pv[i][12]||0),rfidUid:String(pv[i][13]||""),rfidStatus:String(pv[i][14]||"PENDING")});
  let salesCount=0,units=0,amount=0; const sh=getOrCreateSheet_(ss,"Sales",SALES_HEADERS), sv=sh.getDataRange().getValues(); for(let i=1;i<sv.length;i++)if(String(sv[i][1]||"")===shopId){salesCount++;units+=Number(sv[i][7]||0);amount+=Number(sv[i][8]||0);}
  const users=getOrCreateSheet_(ss,"Users",USER_HEADERS), uv=users.getDataRange().getValues(); let exitCount=1,shopName=""; for(let i=1;i<uv.length;i++)if(String(uv[i][4]||"")===shopId){exitCount=Math.max(1,Number(uv[i][9]||1));shopName=String(uv[i][5]||"");break;}
  return jsonResponse({ok:true,shopId,shopName,exitCount,stats:{products:products.length,salesCount:salesCount,salesUnits:units,salesAmount:amount},products:products});
}
function handleUpdateExitCount_(d){
  const shopId=String(d.shopId||"").trim(), count=Math.min(20,Math.max(1,Math.floor(Number(d.exitCount||1)))); if(!shopId)return jsonResponse({ok:false,error:"Shop ID required"});
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID), sh=getOrCreateSheet_(ss,"Users",USER_HEADERS), vals=sh.getDataRange().getValues(); for(let i=1;i<vals.length;i++)if(String(vals[i][4]||"")===shopId){sh.getRange(i+1,10).setValue(count);return jsonResponse({ok:true,exitCount:count});}
  return jsonResponse({ok:false,error:"Shop not found"});
}

function handleExitDetect_(d){
  const shopName=String(d.shopName||"").trim(),shopId=String(d.shopId||"").trim(),exitNo=Math.max(1,Math.floor(Number(d.exitNo||1))),rfid=normalizeRfid_(d.rfidUid);
  if(!rfid||(!shopName&&!shopId))return jsonResponse({ok:false,error:"shopName/shopId, exitNo and rfidUid required"});
  const lock=LockService.getScriptLock();
  try{
    lock.waitLock(10000);
    const ss=SpreadsheetApp.openById(SPREADSHEET_ID), inv=getOrCreateSheet_(ss,"RFID_Inventory",RFID_HEADERS),iv=inv.getDataRange().getValues();
    let rfidRow=-1,rfidBarcode="",rfidShopId="",rfidShopName="",rfidStatus="";
    for(let i=1;i<iv.length;i++)if(String(iv[i][1]||"")===rfid&&((shopId&&String(iv[i][2]||"")===shopId)||(shopName&&String(iv[i][3]||"")===shopName))){rfidRow=i+1;rfidBarcode=String(iv[i][4]||"");rfidShopId=String(iv[i][2]||shopId);rfidShopName=String(iv[i][3]||shopName);rfidStatus=String(iv[i][6]||"");break;}
    if(rfidRow<0)return jsonResponse({ok:false,error:"RFID not assigned to a product in this shop",rfidUid:rfid});
    if(rfidStatus!=="ACTIVE")return jsonResponse({ok:false,error:"This RFID tag has already been used for this product. Scan the RFID of another available unit.",rfidUid:rfid,status:rfidStatus});
    const ps=getOrCreateSheet_(ss,"Products",PRODUCT_HEADERS),pv=ps.getDataRange().getValues(); let productRow=-1,prod=null,stock=0;
    for(let i=1;i<pv.length;i++)if(String(pv[i][1]||"")===rfidShopId&&String(pv[i][4]||"")===rfidBarcode){productRow=i+1;stock=Math.max(0,Number(pv[i][12]||0));prod={barcode:String(pv[i][4]||""),name:String(pv[i][5]||""),price:Number(String(pv[i][9]||"").replace(/,/g,""))||0};break;}
    if(!prod)return jsonResponse({ok:false,error:"Product is no longer in active inventory for this RFID",rfidUid:rfid});
    if(stock<=0)return jsonResponse({ok:false,error:"Product is out of stock",rfidUid:rfid,barcode:prod.barcode,quantity:0});
    const newQuantity=stock-1;
    const ex=getOrCreateSheet_(ss,"Exit_Events",EXIT_HEADERS),evId="EXT"+Utilities.getUuid().replace(/-/g,"").substring(0,10).toUpperCase();
    ex.appendRow([new Date(),evId,rfidShopId,rfidShopName,exitNo,rfid,prod.barcode,prod.name,prod.price,"WAITING_PAYMENT",""]);
    inv.getRange(rfidRow,7,1,4).setValues([["EXITED",iv[rfidRow-1][7]||new Date(),evId,new Date()]]);
    if(newQuantity<=0)ps.deleteRow(productRow); else ps.getRange(productRow,13).setValue(newQuantity);
    return jsonResponse({ok:true,eventId:evId,product:prod,exitNo:exitNo,rfidUid:rfid,previousQuantity:stock,remainingQuantity:newQuantity,productDeleted:newQuantity<=0});
  }catch(err){return jsonResponse({ok:false,error:String(err&&err.message||err)});}
  finally{try{lock.releaseLock();}catch(e){}}
}
function handleExitEvents_(p){const shopId=String(p.shopId||"").trim();if(!shopId)return jsonResponse({ok:false,error:"shopId required"});const ss=SpreadsheetApp.openById(SPREADSHEET_ID),sh=getOrCreateSheet_(ss,"Exit_Events",EXIT_HEADERS),v=sh.getDataRange().getValues(),events=[];for(let i=1;i<v.length;i++)if(String(v[i][2]||"")===shopId&&String(v[i][9]||"")==="WAITING_PAYMENT")events.push({row:i+1,eventId:String(v[i][1]||""),exitNo:Number(v[i][4]||1),rfidUid:String(v[i][5]||""),barcode:String(v[i][6]||""),productName:String(v[i][7]||""),price:Number(v[i][8]||0),time:v[i][0] instanceof Date?v[i][0].toISOString():String(v[i][0]||"")});return jsonResponse({ok:true,events:events,count:events.length});}
function handleRemoveExitEvent_(d){
  const shopId=String(d.shopId||"").trim(),id=String(d.eventId||"").trim(); if(!shopId||!id)return jsonResponse({ok:false,error:"shopId and eventId required"});
  const lock=LockService.getScriptLock();
  try{
    lock.waitLock(10000);
    const ss=SpreadsheetApp.openById(SPREADSHEET_ID),sh=getOrCreateSheet_(ss,"Exit_Events",EXIT_HEADERS),v=sh.getDataRange().getValues();
    for(let i=1;i<v.length;i++)if(String(v[i][2]||"")===shopId&&String(v[i][1]||"")===id&&String(v[i][9]||"")==="WAITING_PAYMENT"){
      const rfid=normalizeRfid_(v[i][5]),barcode=String(v[i][6]||"");
      const ps=getOrCreateSheet_(ss,"Products",PRODUCT_HEADERS),pv=ps.getDataRange().getValues();let restored=false;
      for(let j=1;j<pv.length;j++)if(String(pv[j][1]||"")===shopId&&String(pv[j][4]||"")===barcode){ps.getRange(j+1,13).setValue(Math.max(0,Number(pv[j][12]||0)+1));restored=true;break;}
      if(!restored)return jsonResponse({ok:false,error:"Cannot restore stock because the product was deleted at zero stock. Re-add the product before removing this exit event."});
      const inv=getOrCreateSheet_(ss,"RFID_Inventory",RFID_HEADERS),iv=inv.getDataRange().getValues();
      for(let j=1;j<iv.length;j++)if(String(iv[j][1]||"")===rfid&&String(iv[j][2]||"")===shopId&&String(iv[j][6]||"")==="EXITED"){inv.getRange(j+1,7,1,4).setValues([["ACTIVE",iv[j][7]||new Date(),"",new Date()]]);break;}
      sh.getRange(i+1,10).setValue("REMOVED");return jsonResponse({ok:true,stockRestored:true});
    }
    return jsonResponse({ok:false,error:"Exit event not found or already processed"});
  }catch(err){return jsonResponse({ok:false,error:String(err&&err.message||err)});}finally{try{lock.releaseLock();}catch(e){}}
}
function handlePayExitEvents_(d){
  const shopId=String(d.shopId||"").trim();
  const userId=String(d.userId||"").trim();
  const shopName=String(d.shopName||"").trim();
  let ids=[];
  if(Array.isArray(d.eventIds)) ids=d.eventIds.map(String);
  else if(d.eventIds!==undefined && d.eventIds!==null && String(d.eventIds).trim()){
    const rawIds=String(d.eventIds).trim();
    try{
      const parsed=JSON.parse(rawIds);
      ids=Array.isArray(parsed)?parsed.map(String):[String(parsed)];
    }catch(e){
      ids=rawIds.split(",").map(function(x){return x.trim();}).filter(String);
    }
  }
if(!shopId||!userId||!ids.length)return jsonResponse({ok:false,error:"shopId, userId and eventIds required"});const ss=SpreadsheetApp.openById(SPREADSHEET_ID),ex=getOrCreateSheet_(ss,"Exit_Events",EXIT_HEADERS),ev=ex.getDataRange().getValues(),sales=getOrCreateSheet_(ss,"Sales",SALES_HEADERS),created=[];let grand=0;for(let i=1;i<ev.length;i++)if(String(ev[i][2]||"")===shopId&&ids.indexOf(String(ev[i][1]||""))>=0&&String(ev[i][9]||"")==="WAITING_PAYMENT"){const price=Number(ev[i][8]||0),sid="SALE"+Utilities.getUuid().replace(/-/g,"").substring(0,10).toUpperCase();sales.appendRow([new Date(),shopId,shopName,userId,String(ev[i][6]||""),String(ev[i][7]||""),price,1,price,sid]);ex.getRange(i+1,10,1,2).setValues([["PAID",sid]]);grand+=price;created.push(sid);}return jsonResponse({ok:true,total:grand,saleIds:created,message:"Payment recorded. Bill generated on screen."});}
function handleDeleteProduct_(d){
  const shopId=String(d.shopId||"").trim(),barcode=String(d.barcode||"").trim();
  if(!shopId||!barcode)return jsonResponse({ok:false,error:"Shop ID and barcode required"});
  const lock=LockService.getScriptLock();
  try{
    lock.waitLock(10000);
    const ss=SpreadsheetApp.openById(SPREADSHEET_ID),sh=getOrCreateSheet_(ss,"Products",PRODUCT_HEADERS),v=sh.getDataRange().getValues();
    let found=false;
    for(let i=1;i<v.length;i++)if(String(v[i][1]||"")===shopId&&String(v[i][4]||"")===barcode){sh.deleteRow(i+1);found=true;break;}
    if(!found)return jsonResponse({ok:false,error:"Product not found"});
    const inv=getOrCreateSheet_(ss,"RFID_Inventory",RFID_HEADERS),iv=inv.getDataRange().getValues();
    for(let i=iv.length-1;i>=1;i--)if(String(iv[i][2]||"")===shopId&&String(iv[i][4]||"")===barcode&&String(iv[i][6]||"")==="ACTIVE")inv.getRange(i+1,7,1,4).setValues([["UNASSIGNED",iv[i][7]||new Date(),"",new Date()]]);
    const a=getOrCreateSheet_(ss,"RFID_Assignments",ASSIGN_HEADERS),av=a.getDataRange().getValues();
    for(let i=av.length-1;i>=1;i--)if(String(av[i][2]||"")===shopId&&String(av[i][5]||"")===barcode&&String(av[i][7]||"")==="PENDING")a.deleteRow(i+1);
    return jsonResponse({ok:true});
  }catch(err){return jsonResponse({ok:false,error:String(err&&err.message||err)});}finally{try{lock.releaseLock();}catch(e){}}
}
function handleRecordSale_(d){
  const shopId=String(d.shopId||"").trim(),userId=String(d.userId||"").trim(),shopName=String(d.shopName||"").trim(),barcode=String(d.barcode||"").trim(),qty=Math.max(1,Math.floor(Number(d.quantity||1)));
  if(!shopId||!userId||!barcode)return jsonResponse({ok:false,error:"shopId, userId and barcode required"});
  const lock=LockService.getScriptLock();
  try{
    lock.waitLock(10000);
    const ss=SpreadsheetApp.openById(SPREADSHEET_ID),ps=getOrCreateSheet_(ss,"Products",PRODUCT_HEADERS),v=ps.getDataRange().getValues();
    for(let i=1;i<v.length;i++)if(String(v[i][1]||"")===shopId&&String(v[i][4]||"")===barcode){
      const stock=Math.max(0,Number(v[i][12]||0)); if(stock<qty)return jsonResponse({ok:false,error:"Not enough stock. Available: "+stock});
      const inv=getOrCreateSheet_(ss,"RFID_Inventory",RFID_HEADERS);migrateLegacyRfid_(ss,v[i]);
      const active=countActiveRfids_(inv,shopId,barcode); if(active>stock)return jsonResponse({ok:false,error:"RFID inventory is inconsistent with stock. Please refresh/repair RFID assignments."});
      if(active<Math.min(qty,stock))return jsonResponse({ok:false,error:"Only "+active+" active RFID tag(s) are available for this product. Assign one RFID per physical unit before selling."});
      const price=Number(String(v[i][9]||"").replace(/,/g,""))||0,newStock=stock-qty,sh=getOrCreateSheet_(ss,"Sales",SALES_HEADERS),sid="SALE"+Utilities.getUuid().replace(/-/g,"").substring(0,10).toUpperCase();
      sh.appendRow([new Date(),shopId,shopName,userId,barcode,String(v[i][5]||""),price,qty,price*qty,sid]);
      const iv=inv.getDataRange().getValues();let marked=0;for(let j=1;j<iv.length&&marked<qty;j++)if(String(iv[j][2]||"")===shopId&&String(iv[j][4]||"")===barcode&&String(iv[j][6]||"")==="ACTIVE"){inv.getRange(j+1,7,1,4).setValues([["SOLD",iv[j][7]||new Date(),sid,new Date()]]);marked++;}
      if(newStock<=0)ps.deleteRow(i+1);else ps.getRange(i+1,13).setValue(newStock);
      return jsonResponse({ok:true,total:price*qty,saleId:sid,remainingQuantity:newStock,productDeleted:newStock<=0,rfidsConsumed:marked});
    }
    return jsonResponse({ok:false,error:"Product not found"});
  }catch(err){return jsonResponse({ok:false,error:String(err&&err.message||err)});}finally{try{lock.releaseLock();}catch(e){}}
}

function getOrCreateSheet_(ss,name,headers){let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);sh.getRange(1,1,1,headers.length).setValues([headers]);sh.setFrozenRows(1);return sh;}
function makeShopId_(){const ss=SpreadsheetApp.openById(SPREADSHEET_ID),sh=getOrCreateSheet_(ss,"Users",USER_HEADERS),v=sh.getDataRange().getValues();let max=0;for(let i=1;i<v.length;i++){const m=String(v[i][4]||"").match(/^SHOP(\d+)$/);if(m)max=Math.max(max,Number(m[1]));}return "SHOP"+String(max+1).padStart(3,"0");}
function findShopIdByName_(ss,name){const sh=getOrCreateSheet_(ss,"Users",USER_HEADERS),v=sh.getDataRange().getValues();for(let i=1;i<v.length;i++)if(String(v[i][5]||"")===name)return String(v[i][4]||"");return "";}
function normalizePhone_(p){let x=String(p||"").replace(/[^\d+]/g,"");if(x.startsWith("+91"))x=x.substring(3);if(x.startsWith("91")&&x.length===12)x=x.substring(2);return x;}
function normalizeRfid_(x){return String(x||"").replace(/[^0-9A-Fa-f]/g,"").toUpperCase();}
function jsonResponse(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
