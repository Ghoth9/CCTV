 // ==========================================
// การตั้งค่าชื่อ Sheet ที่ต้องการใช้งาน
// ==========================================
const SHEET_DATA = 'CCTV_DB';  
const SHEET_ADMIN = 'USERS_DB'; 
const SHEET_MAINTAIN = 'MAINTAIN_DB';
const SHEET_INVENTORY = 'INVENTORY_DB'; 
const SHEET_BORROW = 'BORROW_DB'; 
const SHEET_CLAIM = 'CLAIM_DB'; 
const SHEET_INSTALL = 'INSTALL_DB';
const SHEET_LOG = 'LOG_DB';

function testAuth() {
  DriveApp.getFiles(); 
  getActiveSS();
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('CCTV Management System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

// ==========================================
// 1. API ระบบกล้อง (CCTV)
// ==========================================
function getCCTVData() {
  try {
    const sheet = getActiveSS().getSheetByName(SHEET_DATA);
    if (!sheet) return { success: false, error: "ไม่พบ Sheet: " + SHEET_DATA };

    const data = sheet.getDataRange().getValues();
    let result = [];
    for(let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        result.push({
          id: data[i][0].toString(),
          name: data[i][1].toString(),
          // ใส่ Fallback || 0 ป้องกันค่า NaN ทำให้ Map แครช
          lat: parseFloat(data[i][2]) || 0,
          lng: parseFloat(data[i][3]) || 0,
          status: data[i][4].toString(),
          lastUpdate: data[i][5] ? data[i][5].toString() : "",
          mediaUrl: data[i][6] ? data[i][6].toString() : "", 
          project: data[i][7] ? data[i][7].toString() : "-"  
        });
      }
    }
    return { success: true, data: result };
  } catch (err) { 
    return { success: false, error: err.toString() };
  }
}

function updateCamera(camData) {
  try {
    const sheet = getActiveSS().getSheetByName(SHEET_DATA);
    const data = sheet.getDataRange().getValues();
    for(let i = 1; i < data.length; i++) {
      if(data[i][0].toString() === camData.id.toString()) {
        if(camData.status) sheet.getRange(i + 1, 5).setValue(camData.status);
        sheet.getRange(i + 1, 6).setValue(getThaiDate());
        return { success: true };
      }
    }
    return { success: false, error: "Not found" };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function addCamera(camData) {
  try {
    const sheet = getActiveSS().getSheetByName(SHEET_DATA);
    
    // ++ 1. เพิ่มการตรวจสอบ ID ซ้ำ ++
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      // ตรวจสอบว่าคอลัมน์แรก (ID) ตรงกับที่ส่งมาหรือไม่
      if (data[i][0] && data[i][0].toString().trim() === camData.id.toString().trim()) {
        return { success: false, error: "รหัสกล้องนี้ (ID) มีอยู่ในระบบแล้ว กรุณาใช้รหัสอื่น" };
      }
    }

    // ++ 2. ดักจับค่าว่างที่ฝั่ง Backend (โค้ดเดิม) ++
    if (!camData.project || camData.project.trim() === "") {
        return { success: false, error: "กรุณาระบุชื่อโครงการ / โซน" };
    }

    let imgUrl = camData.mediaUrl || "";
    if (camData.imgData && camData.imgData !== "") {
      const folder = getOrCreateFolder("CCTV_Images");
      const blob = Utilities.newBlob(Utilities.base64Decode(camData.imgData), camData.mimeType, camData.fileName);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      imgUrl = "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w800";
    }

    // บันทึกข้อมูลลงชีตเมื่อผ่านการตรวจสอบทั้งหมด
    sheet.appendRow([
      camData.id, camData.name, camData.lat, camData.lng, camData.status, 
      getThaiDate(), imgUrl, camData.project 
    ]);
    
    return { success: true, imgUrl: imgUrl };
  } catch (err) { 
    return { success: false, error: err.toString() };
  }
}


function getUsersData() {
  try {
    const sheet = getActiveSS().getSheetByName('USERS_DB');
    if (!sheet) return { success: true, data: [] };
    
    const data = sheet.getDataRange().getValues();
    let result = [];
    
    for(let i = 1; i < data.length; i++) {
      if(data[i][0] && data[i][0].toString().trim() !== "") {
        result.push({ 
          username: data[i][0].toString().trim(), 
          password: data[i][1] ? data[i][1].toString() : "", 
          role: data[i][2] ? data[i][2].toString() : "Member", 
          fullName: data[i][3] ? data[i][3].toString() : "",
          assignedZones: data[i][4] ? data[i][4].toString() : "",
          assignedMembers: data[i][5] ? data[i][5].toString() : "",
          phone: data[i][6] ? data[i][6].toString() : "" // ดึงเบอร์โทร
        });
      }
    }
    return { success: true, data: result };
  } catch(err) { return { success: false, error: err.toString() }; }
}



// ==========================================
// 3. API แจ้งซ่อม (Maintenance)
// ==========================================
function addMaintenanceTicket(ticketData) {
  try {
    let sheet = getActiveSS().getSheetByName(SHEET_MAINTAIN);
    if(!sheet) {
      sheet = getActiveSS().insertSheet(SHEET_MAINTAIN);
      sheet.appendRow(["Ticket ID", "Cam ID", "Issue", "Status", "Reported By", "Date", "Images"]);
    } else {
      if (sheet.getRange(1, 7).getValue() === "") {
        sheet.getRange(1, 7).setValue("Images");
      }
    }

    let imgUrls = [];
    if (ticketData.images && ticketData.images.length > 0) {
      const folder = getOrCreateFolder("Ticket_Images"); 
      for (let i = 0; i < ticketData.images.length; i++) {
        const img = ticketData.images[i];
        if (img.imgData) {
          const blob = Utilities.newBlob(Utilities.base64Decode(img.imgData), img.mimeType, img.fileName);
          const file = folder.createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          imgUrls.push("https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w800");
        }
      }
    }

    // สร้างตัวแปรเวลาไทย
    const thaiDate = Utilities.formatDate(new Date(), "Asia/Bangkok", "dd/MM/yyyy HH:mm:ss");
    const ticketId = "T" + new Date().getTime().toString().slice(-6); 

    sheet.appendRow([ 
      ticketId, ticketData.camId, ticketData.issue, "รอดำเนินการ", 
      ticketData.reportedBy, thaiDate, imgUrls.join(',')
    ]);
    return { success: true, ticketId: ticketId };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function getMaintenanceTickets() {
  try {
    const sheet = getActiveSS().getSheetByName(SHEET_MAINTAIN);
    if(!sheet) return { success: true, data: [] };
    const data = sheet.getDataRange().getValues();
    let result = [];
    for(let i = 1; i < data.length; i++) {
      if(data[i][0]) {
        result.push({ 
          ticketId: data[i][0].toString(), 
          camId: data[i][1].toString(), 
          issue: data[i][2].toString(), 
          status: data[i][3].toString(), 
          reportedBy: data[i][4].toString(), 
          date: data[i][5].toString(),
          images: data[i][6] ? data[i][6].toString() : "" 
        });
      }
    }
    return { success: true, data: result.reverse() }; 
  } catch(err) { return { success: false, error: err.toString() }; }
}

function updateMaintenanceTicketStatus(ticketId, newStatus) {
  try {
    const sheet = getActiveSS().getSheetByName(SHEET_MAINTAIN);
    const data = sheet.getDataRange().getValues();
    for(let i = 1; i < data.length; i++) {
      if(data[i][0].toString() === ticketId.toString()) {
        sheet.getRange(i + 1, 4).setValue(newStatus); 
        return { success: true, message: "อัปเดตสถานะสำเร็จ" };
      }
    }
    return { success: false, error: "ไม่พบรหัสการแจ้งซ่อมนี้" };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function saveUserData(userData) {
  try {
    let sheet = getActiveSS().getSheetByName('USERS_DB');
    if (!sheet) {
      sheet = getActiveSS().insertSheet('USERS_DB');
      sheet.appendRow(["Username", "Password", "Role", "Full Name", "Zones", "Members", "Phone"]);
    }
    
    const data = sheet.getDataRange().getValues();
    let isUpdated = false;
    
    // เติม ' (Single Quote) นำหน้า เพื่อป้องกัน Google Sheets ตัดเลข 0
    let phoneText = userData.phone ? "'" + userData.phone.toString() : "";
    
    for(let i = 1; i < data.length; i++) {
      if(data[i][0] && data[i][0].toString().trim() === userData.username.toString().trim()) {
        sheet.getRange(i + 1, 2).setValue(userData.password || "");
        sheet.getRange(i + 1, 3).setValue(userData.role || "Member");
        sheet.getRange(i + 1, 4).setValue(userData.fullName || "");
        sheet.getRange(i + 1, 5).setValue(userData.assignedZones || "");
        sheet.getRange(i + 1, 6).setValue(userData.assignedMembers || "");
        sheet.getRange(i + 1, 7).setValue(phoneText); // อัปเดตเบอร์ (มี ' ดักไว้)
        isUpdated = true;
        break;
      }
    }
    
    if (!isUpdated) {
        sheet.appendRow([
            userData.username || "", 
            userData.password || "", 
            userData.role || "Member", 
            userData.fullName || "", 
            userData.assignedZones || "", 
            userData.assignedMembers || "",
            phoneText // เพิ่มข้อมูลบรรทัดใหม่ (มี ' ดักไว้)
        ]);
    }
    return { success: true, message: "บันทึกข้อมูลสำเร็จ" };
  } catch(err) { return { success: false, error: err.toString() }; }
}

function verifyLogin(username, password) {
  try {
    const sheet = getActiveSS().getSheetByName('USERS_DB');
    if (!sheet) return { success: false, message: "ไม่พบฐานข้อมูลบัญชีผู้ใช้งาน (USERS_DB)" };
    
    const data = sheet.getDataRange().getValues();
    let userFound = false;
    
    let adminName = "Admin ส่วนกลาง";
    let adminPhone = "ไม่มีข้อมูลติดต่อ";
    for(let k = 1; k < data.length; k++) {
        if(data[k][2] === "Admin") {
            adminName = data[k][3].toString();
            adminPhone = data[k][6] ? data[k][6].toString() : "ไม่มีข้อมูลติดต่อ";
            break;
        }
    }
    
    for(let i = 1; i < data.length; i++) {
      if(data[i][0].toString() === username.toString()) {
        userFound = true;
        if(data[i][1].toString() === password.toString()) {
          
          let role = data[i][2] || "Staff";
          let fullName = data[i][3] || "Unknown";
          let phone = data[i][6] ? data[i][6].toString() : ""; 
          
          let caretakerName = "";
          let caretakerPhone = "";
          let assignedMembers = data[i][5] ? data[i][5].toString() : "";
          let assignedMembersDetails = "";

          if (role === "Member") {
              for(let j = 1; j < data.length; j++) {
                  if (data[j][2] === "Staff" && data[j][5] && data[j][5].toString().includes(fullName)) {
                      caretakerName = data[j][3].toString();
                      caretakerPhone = data[j][6] ? data[j][6].toString() : "";
                      break;
                  }
              }
          } 
          else if (role === "Staff" && assignedMembers !== "") {
              let mNames = assignedMembers.split(',').map(m => m.trim());
              let detailsArr = [];
              for (let mName of mNames) {
                  let mPhone = "-";
                  for (let r = 1; r < data.length; r++) {
                      if (data[r][3] && data[r][3].toString().trim() === mName) {
                          mPhone = data[r][6] ? data[r][6].toString() : "-";
                          break;
                      }
                  }
                  detailsArr.push(`${mName} <span class="text-slate-400 font-mono">(${mPhone})</span>`);
              }
              assignedMembersDetails = detailsArr.join('<br>• ');
          }

          return { 
            success: true, role: role, fullName: fullName, phone: phone,
            assignedZones: data[i][4] ? data[i][4].toString() : "",   
            assignedMembers: assignedMembers,
            assignedMembersDetails: assignedMembersDetails,
            caretakerName: caretakerName, caretakerPhone: caretakerPhone,
            adminName: adminName, adminPhone: adminPhone
          };
        }
      }
    }
    
    if (!userFound) return { success: false, message: "ไม่พบบัญชีผู้ใช้นี้ในระบบ" };
  } catch (err) { return { success: false, message: "เกิดข้อผิดพลาด: " + err.toString() }; }
}

// **ใหม่** ฟังก์ชันแก้ไขข้อมูลการแจ้งซ่อม (รองรับการเพิ่มรูปภาพใหม่)
function editMaintenanceTicket(data) {
  try {
    const sheet = getActiveSS().getSheetByName(SHEET_MAINTAIN);
    const sheetData = sheet.getDataRange().getValues();
    
    for(let i = 1; i < sheetData.length; i++) {
      if(sheetData[i][0].toString() === data.ticketId.toString()) {
        sheet.getRange(i + 1, 3).setValue(data.issue);
        
        let finalImages = data.oldImages || ""; 
        if (data.newImages && data.newImages.length > 0) {
          const folder = getOrCreateFolder("Ticket_Images");
          let newImgUrls = [];
          for (let j = 0; j < data.newImages.length; j++) {
            const img = data.newImages[j];
            if (img.imgData) {
              const blob = Utilities.newBlob(Utilities.base64Decode(img.imgData), img.mimeType, img.fileName);
              const file = folder.createFile(blob);
              file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
              newImgUrls.push("https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w800");
            }
          }
          if (finalImages !== "") finalImages += ",";
          finalImages += newImgUrls.join(',');
        }
        sheet.getRange(i + 1, 7).setValue(finalImages);
        return { success: true, message: "แก้ไขข้อมูลการแจ้งซ่อมสำเร็จ" };
      }
    }
    return { success: false, error: "ไม่พบรหัสการแจ้งซ่อมนี้" };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// ==========================================
// 4. API ระบบสต๊อกสินค้า (Inventory)
// ==========================================
function getInventoryData() {
  try {
    let sheet = getActiveSS().getSheetByName(SHEET_INVENTORY);
    if (!sheet) {
      sheet = getActiveSS().insertSheet(SHEET_INVENTORY);
      sheet.appendRow(["Item ID", "Item Name", "Category", "Total Qty", "Available Qty", "Min Stock", "Unit", "Last Update"]);
      return { success: true, data: [] }; 
    }

    const data = sheet.getDataRange().getValues();
    let result = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) { 
        result.push({
          itemId: data[i][0].toString(),
          itemName: data[i][1].toString(),
          category: data[i][2].toString(),
          totalQty: parseInt(data[i][3]) || 0,
          availableQty: parseInt(data[i][4]) || 0,
          minStock: parseInt(data[i][5]) || 0,
          unit: data[i][6].toString(),
          lastUpdate: data[i][7] ? data[i][7].toString() : ""
        });
      }
    }
    return { success: true, data: result };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function addInventoryItem(itemData) {
  try {
    const sheet = getActiveSS().getSheetByName(SHEET_INVENTORY);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === itemData.itemId.toString()) {
        return { success: false, error: "รหัสสินค้านี้ (Item ID) มีอยู่ในระบบแล้ว กรุณาใช้รหัสอื่น" };
      }
    }
    sheet.appendRow([
      itemData.itemId, itemData.itemName, itemData.category, itemData.totalQty,
      itemData.availableQty, itemData.minStock, itemData.unit, getThaiDate()
    ]);
    return { success: true, message: "เพิ่มสินค้าใหม่ลงสต๊อกสำเร็จ" };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function updateInventoryItem(itemData) {
  try {
    const sheet = getActiveSS().getSheetByName(SHEET_INVENTORY);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === itemData.itemId.toString()) {
        sheet.getRange(i + 1, 2).setValue(itemData.itemName);
        sheet.getRange(i + 1, 3).setValue(itemData.category);
        sheet.getRange(i + 1, 4).setValue(itemData.totalQty);
        sheet.getRange(i + 1, 5).setValue(itemData.availableQty);
        sheet.getRange(i + 1, 6).setValue(itemData.minStock);
        sheet.getRange(i + 1, 7).setValue(itemData.unit);
        sheet.getRange(i + 1, 8).setValue(getThaiDate());
        return { success: true, message: "อัปเดตข้อมูลสินค้าสำเร็จ" };
      }
    }
    return { success: false, error: "ไม่พบสินค้ารหัสนี้ในระบบ" };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// ==========================================
// 5. API ระบบยืม/คืนสินค้า (Borrow & Return)
// ==========================================
function getBorrowData() {
  try {
    let sheet = getActiveSS().getSheetByName(SHEET_BORROW);
    if (!sheet) {
      sheet = getActiveSS().insertSheet(SHEET_BORROW);
      sheet.appendRow(["Borrow ID", "Item ID", "Item Name", "Borrower", "Qty", "Borrow Date", "Expected Return", "Status", "Actual Return", "Note"]);
      return { success: true, data: [] };
    }

    const data = sheet.getDataRange().getValues();
    let result = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        result.push({
          borrowId: data[i][0].toString(),
          itemId: data[i][1].toString(),
          itemName: data[i][2].toString(),
          borrower: data[i][3].toString(),
          qty: parseInt(data[i][4]) || 0,
          borrowDate: data[i][5].toString(),
          expectedReturn: data[i][6].toString(),
          status: data[i][7].toString(), 
          actualReturn: data[i][8] ? data[i][8].toString() : "",
          note: data[i][9] ? data[i][9].toString() : ""
        });
      }
    }
    return { success: true, data: result.reverse() };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function borrowItem(data) {
  // 1. เรียกใช้งาน LockService แบบ ScriptLock
  const lock = LockService.getScriptLock();
  
  try {
    // รอคิวการประมวลผลสูงสุด 3 วินาที (ป้องกัน Race Condition)
    lock.waitLock(3000);

    // 2. ตรวจสอบและแปลงค่าจำนวนที่ยืม ต้องเป็นตัวเลขและมากกว่า 0 เท่านั้น
    const borrowQty = parseInt(data.qty);
    if (isNaN(borrowQty) || borrowQty <= 0) {
      return { success: false, error: "จำนวนที่ยืมต้องมากกว่า 0 เท่านั้น" };
    }

    const ss = getActiveSS();
    const invSheet = ss.getSheetByName(SHEET_INVENTORY);
    const borrowSheet = ss.getSheetByName(SHEET_BORROW);
    
    const invData = invSheet.getDataRange().getValues();
    let itemRowIndex = -1;
    let currentAvailQty = 0;
    
    for (let i = 1; i < invData.length; i++) {
      if (invData[i][0].toString() === data.itemId.toString()) {
        itemRowIndex = i + 1;
        currentAvailQty = parseInt(invData[i][4]) || 0; 
        break;
      }
    }
    
    if (itemRowIndex === -1) return { success: false, error: "ไม่พบสินค้ารหัสนี้ในระบบ" };
    if (currentAvailQty < borrowQty) return { success: false, error: `สต๊อกไม่พอ! (มีของพร้อมใช้แค่ ${currentAvailQty})` };
    
    // อัปเดตจำนวนที่เหลือในสต๊อก (ตัดสต๊อก)
    invSheet.getRange(itemRowIndex, 5).setValue(currentAvailQty - borrowQty);
    invSheet.getRange(itemRowIndex, 8).setValue(getThaiDate());
    
    // บันทึกประวัติการยืม
    const borrowId = "BR" + new Date().getTime().toString().slice(-6);
    borrowSheet.appendRow([
      borrowId, data.itemId, data.itemName, data.borrower, borrowQty, 
      getThaiDate(), data.expectedReturn, "กำลังยืม", "", data.note
    ]);
    
    return { success: true, message: "ทำรายการยืมสำเร็จ" };
    
  } catch (err) { 
    return { success: false, error: err.toString() };
  } finally {
    // 3. ปลดล็อคเสมอเมื่อประมวลผลเสร็จ หรือเกิดข้อผิดพลาด
    lock.releaseLock();
  }
}

function returnItem(borrowId) {
  // 1. เรียกใช้งาน LockService แบบ ScriptLock เพื่อกันการแย่งอัปเดตสต๊อก
  const lock = LockService.getScriptLock();
  
  try {
    // รอคิวประมวลผลสูงสุด 3 วินาที
    lock.waitLock(3000);

    const ss = getActiveSS();
    const invSheet = ss.getSheetByName(SHEET_INVENTORY);
    const borrowSheet = ss.getSheetByName(SHEET_BORROW);
    
    const borrowData = borrowSheet.getDataRange().getValues();
    let bRowIndex = -1;
    let itemId = "";
    let qtyToReturn = 0;
    
    for (let i = 1; i < borrowData.length; i++) {
      if (borrowData[i][0].toString() === borrowId.toString() && borrowData[i][7] === "กำลังยืม") {
        bRowIndex = i + 1;
        itemId = borrowData[i][1].toString();
        qtyToReturn = parseInt(borrowData[i][4]) || 0;
        break;
      }
    }
    
    if (bRowIndex === -1) return { success: false, error: "ไม่พบรายการยืม หรือสินค้านี้ถูกคืนไปแล้ว" };
    
    const invData = invSheet.getDataRange().getValues();
    for (let i = 1; i < invData.length; i++) {
      if (invData[i][0].toString() === itemId) {
        let currentAvail = parseInt(invData[i][4]) || 0;
        invSheet.getRange(i + 1, 5).setValue(currentAvail + qtyToReturn);
        invSheet.getRange(i + 1, 8).setValue(getThaiDate());
        break;
      }
    }
    
    borrowSheet.getRange(bRowIndex, 8).setValue("คืนแล้ว");
    borrowSheet.getRange(bRowIndex, 9).setValue(getThaiDate()); 
    
    return { success: true, message: "คืนสินค้าเข้าสต๊อกสำเร็จ" };
  } catch (err) { 
    return { success: false, error: err.toString() }; 
  } finally {
    // ปลดล็อกคิวเสมอ
    lock.releaseLock();
  }
}

// ==========================================
// 6. API ระบบเคลมสินค้า (Claim)
// ==========================================
function getClaimData() {
  try {
    let sheet = getActiveSS().getSheetByName(SHEET_CLAIM);
    if (!sheet) {
      sheet = getActiveSS().insertSheet(SHEET_CLAIM);
      sheet.appendRow(["Claim ID", "Item ID", "Item Name", "Serial Number", "Issue", "Reported By", "Claim Date", "Status", "Update Date", "Note"]);
      return { success: true, data: [] };
    }

    const data = sheet.getDataRange().getValues();
    let result = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        result.push({
          claimId: data[i][0].toString(),
          itemId: data[i][1].toString(),
          itemName: data[i][2].toString(),
          serialNumber: data[i][3].toString(),
          issue: data[i][4].toString(),
          reportedBy: data[i][5].toString(),
          claimDate: data[i][6].toString(),
          status: data[i][7].toString(),
          updateDate: data[i][8] ? data[i][8].toString() : "",
          note: data[i][9] ? data[i][9].toString() : ""
        });
      }
    }
    return { success: true, data: result.reverse() };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function addClaimTicket(data) {
  try {
    const ss = getActiveSS();
    const claimSheet = ss.getSheetByName(SHEET_CLAIM);
    
    const claimId = "CM" + new Date().getTime().toString().slice(-6);
    
    claimSheet.appendRow([
      claimId, data.itemId, data.itemName, data.serialNumber, data.issue, 
      data.reportedBy, getThaiDate(), "รอส่งเคลม", "", data.note
    ]);
    
    return { success: true, message: "สร้างใบส่งเคลมสำเร็จ" };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function updateClaimStatus(claimId, newStatus, returnToInv) {
  // 1. เรียกใช้งาน LockService แบบ ScriptLock เพื่อกันการแย่งอัปเดตสต๊อก
  const lock = LockService.getScriptLock();
  
  try {
    // รอคิวประมวลผลสูงสุด 3 วินาที
    lock.waitLock(3000);

    const ss = getActiveSS();
    const claimSheet = ss.getSheetByName(SHEET_CLAIM);
    
    const claimData = claimSheet.getDataRange().getValues();
    let cRowIndex = -1;
    let itemId = "";
    
    for (let i = 1; i < claimData.length; i++) {
      if (claimData[i][0].toString() === claimId.toString()) {
        cRowIndex = i + 1;
        itemId = claimData[i][1].toString();
        break;
      }
    }
    
    if (cRowIndex === -1) return { success: false, error: "ไม่พบรหัสเคลมนี้" };
    
    claimSheet.getRange(cRowIndex, 8).setValue(newStatus);
    claimSheet.getRange(cRowIndex, 9).setValue(getThaiDate());
    
    if (newStatus === "เคลมสำเร็จ" && returnToInv) {
        const invSheet = ss.getSheetByName(SHEET_INVENTORY);
        if(invSheet) {
           const invData = invSheet.getDataRange().getValues();
           for (let j = 1; j < invData.length; j++) {
              if (invData[j][0].toString() === itemId) {
                let currentAvail = parseInt(invData[j][4]) || 0;
                invSheet.getRange(j + 1, 5).setValue(currentAvail + 1); 
                invSheet.getRange(j + 1, 8).setValue(getThaiDate());
                break;
              }
           }
        }
    }

    return { success: true, message: "อัปเดตสถานะเคลมสำเร็จ" };
  } catch (err) { 
    return { success: false, error: err.toString() }; 
  } finally {
    // ปลดล็อกคิวเสมอ
    lock.releaseLock();
  }
}

// **ใหม่** ฟังก์ชันแก้ไขข้อมูลแจ้งเคลม
function editClaimTicket(data) {
  try {
    const sheet = getActiveSS().getSheetByName(SHEET_CLAIM);
    const sheetData = sheet.getDataRange().getValues();
    for(let i = 1; i < sheetData.length; i++) {
      if(sheetData[i][0].toString() === data.claimId.toString()) {
        if(data.serialNumber !== undefined) sheet.getRange(i + 1, 4).setValue(data.serialNumber);
        if(data.issue) sheet.getRange(i + 1, 5).setValue(data.issue);
        if(data.note !== undefined) sheet.getRange(i + 1, 10).setValue(data.note);
        sheet.getRange(i + 1, 9).setValue(getThaiDate());
        return { success: true, message: "แก้ไขข้อมูลการส่งเคลมสำเร็จ" };
      }
    }
    return { success: false, error: "ไม่พบรหัสการส่งเคลมนี้" };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// ==========================================
// 8. API งานติดตั้ง (Installation) ** เพิ่มใหม่ **
// ==========================================
function getInstallData() {
  try {
    let sheet = getActiveSS().getSheetByName(SHEET_INSTALL);
    if (!sheet) {
      sheet = getActiveSS().insertSheet(SHEET_INSTALL);
      sheet.appendRow(["Install ID", "Customer Name", "Contact Info", "Install Details", "Images", "Reporter", "Report Date", "Status", "Last Update"]);
      return { success: true, data: [] };
    }
    const data = sheet.getDataRange().getValues();
    let result = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        result.push({
          installId: data[i][0].toString(),
          customerName: data[i][1].toString(),
          contactInfo: data[i][2].toString(),
          installDetails: data[i][3].toString(),
          images: data[i][4] ? data[i][4].toString() : "",
          reportedBy: data[i][5].toString(),
          reportDate: data[i][6].toString(),
          status: data[i][7].toString(),
          lastUpdate: data[i][8] ? data[i][8].toString() : ""
        });
      }
    }
    return { success: true, data: result.reverse() };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function addInstallWork(data) {
  try {
    const ss = getActiveSS();
    let sheet = ss.getSheetByName(SHEET_INSTALL);
    if(!sheet) {
      sheet = ss.insertSheet(SHEET_INSTALL);
      sheet.appendRow(["Install ID", "Customer Name", "Contact Info", "Install Details", "Images", "Reporter", "Report Date", "Status", "Last Update"]);
    }

    let imgUrls = [];
    if (data.images && data.images.length > 0) {
      const folder = getOrCreateFolder("Install_Images"); 
      for (let i = 0; i < data.images.length; i++) {
        const img = data.images[i];
        if (img.imgData) {
          const blob = Utilities.newBlob(Utilities.base64Decode(img.imgData), img.mimeType, img.fileName);
          const file = folder.createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          imgUrls.push("https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w800");
        }
      }
    }

    const installId = "IN" + new Date().getTime().toString().slice(-6); 
    sheet.appendRow([ 
      installId, data.customerName, data.contactInfo, data.installDetails, imgUrls.join(','), 
      data.reportedBy, getThaiDate(), "รอดำเนินการ", ""
    ]);
    return { success: true, installId: installId, message: "สร้างใบงานติดตั้งสำเร็จ" };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function editInstallWork(data) {
  try {
    const sheet = getActiveSS().getSheetByName(SHEET_INSTALL);
    const sheetData = sheet.getDataRange().getValues();
    for(let i = 1; i < sheetData.length; i++) {
      if(sheetData[i][0].toString() === data.installId.toString()) {
        sheet.getRange(i + 1, 2).setValue(data.customerName);
        sheet.getRange(i + 1, 3).setValue(data.contactInfo);
        sheet.getRange(i + 1, 4).setValue(data.installDetails);
        
        let finalImages = data.oldImages || ""; 
        if (data.newImages && data.newImages.length > 0) {
          const folder = getOrCreateFolder("Install_Images");
          let newImgUrls = [];
          for (let j = 0; j < data.newImages.length; j++) {
            const img = data.newImages[j];
            if (img.imgData) {
              const blob = Utilities.newBlob(Utilities.base64Decode(img.imgData), img.mimeType, img.fileName);
              const file = folder.createFile(blob);
              file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
              newImgUrls.push("https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w800");
            }
          }
          if (finalImages !== "") finalImages += ",";
          finalImages += newImgUrls.join(',');
        }
        sheet.getRange(i + 1, 5).setValue(finalImages);
        sheet.getRange(i + 1, 9).setValue(getThaiDate());
        return { success: true, message: "แก้ไขใบงานติดตั้งสำเร็จ" };
      }
    }
    return { success: false, error: "ไม่พบรหัสงานติดตั้งนี้" };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function updateInstallStatus(installId, newStatus) {
  try {
    const sheet = getActiveSS().getSheetByName(SHEET_INSTALL);
    const data = sheet.getDataRange().getValues();
    for(let i = 1; i < data.length; i++) {
      if(data[i][0].toString() === installId.toString()) {
        sheet.getRange(i + 1, 8).setValue(newStatus); 
        sheet.getRange(i + 1, 9).setValue(getThaiDate());
        return { success: true, message: "อัปเดตสถานะงานติดตั้งสำเร็จ" };
      }
    }
    return { success: false, error: "ไม่พบรหัสงานติดตั้งนี้" };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// ==========================================
// 7. API ระบบรายงาน (Dashboard / Report / PDF)
// ==========================================
function getDashboardStats() {
  try {
    const ss = getActiveSS();
    let stats = {
      cctvTotal: 0, cctvOnline: 0, cctvOffline: 0,
      maintainPending: 0,
      inventoryLow: 0,
      borrowActive: 0,
      claimPending: 0,
      installPending: 0 // เพิ่มสถิติงานติดตั้ง
    };

    const cctvSheet = ss.getSheetByName(SHEET_DATA);
    if (cctvSheet) {
      const cctvData = cctvSheet.getDataRange().getValues();
      for(let i = 1; i < cctvData.length; i++) {
        if(cctvData[i][0]) {
          stats.cctvTotal++;
          if(cctvData[i][4] === 'Online') stats.cctvOnline++;
          else if(cctvData[i][4] === 'Offline') stats.cctvOffline++;
        }
      }
    }

    const maintainSheet = ss.getSheetByName(SHEET_MAINTAIN);
    if (maintainSheet) {
      const mData = maintainSheet.getDataRange().getValues();
      for(let i = 1; i < mData.length; i++) {
        if(mData[i][0] && (mData[i][3] === 'รอดำเนินการ' || mData[i][3] === 'กำลังดำเนินการ')) {
          stats.maintainPending++;
        }
      }
    }

    const invSheet = ss.getSheetByName(SHEET_INVENTORY);
    if (invSheet) {
      const invData = invSheet.getDataRange().getValues();
      for(let i = 1; i < invData.length; i++) {
        if(invData[i][0]) {
          let avail = parseInt(invData[i][4]) || 0;
          let min = parseInt(invData[i][5]) || 0;
          if (avail <= min) stats.inventoryLow++;
        }
      }
    }

    const borrowSheet = ss.getSheetByName(SHEET_BORROW);
    if (borrowSheet) {
      const bData = borrowSheet.getDataRange().getValues();
      for(let i = 1; i < bData.length; i++) {
        if(bData[i][0] && bData[i][7] === 'กำลังยืม') {
          stats.borrowActive++;
        }
      }
    }

    const claimSheet = ss.getSheetByName(SHEET_CLAIM);
    if (claimSheet) {
      const cData = claimSheet.getDataRange().getValues();
      for(let i = 1; i < cData.length; i++) {
        if(cData[i][0] && (cData[i][7] === 'รอส่งเคลม' || cData[i][7] === 'กำลังเคลม')) {
          stats.claimPending++;
        }
      }
    }

    const installSheet = ss.getSheetByName(SHEET_INSTALL);
    if (installSheet) {
      const inData = installSheet.getDataRange().getValues();
      for(let i = 1; i < inData.length; i++) {
        if(inData[i][0] && (inData[i][7] === 'รอดำเนินการ' || inData[i][7] === 'กำลังดำเนินการ')) {
          stats.installPending++;
        }
      }
    }

    return { success: true, data: stats };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function generateReportPDF() {
  try {
    const ss = getActiveSS();
    const stats = getDashboardStats().data;
    const cctvData = getCCTVData().data.filter(c => c.status === 'Offline');
    
    let htmlContent = `
      <style>
        body { font-family: 'Tahoma', sans-serif; padding: 20px; color: #333; }
        .header { text-align: center; border-bottom: 2px solid #3b82f6; padding-bottom: 10px; }
        .summary-grid { display: flex; flex-wrap: wrap; margin-top: 20px; }
        .card { width: 30%; border: 1px solid #ddd; padding: 10px; margin: 5px; border-radius: 8px; text-align: center; display: inline-block; }
        .card h3 { margin: 0; font-size: 12px; color: #666; }
        .card p { font-size: 20px; font-weight: bold; margin: 5px 0; color: #1e293b; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 12px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f8fafc; }
        .footer { margin-top: 30px; text-align: right; font-size: 10px; color: #999; }
      </style>
      
      <div class="header">
        <h1>รายงานสรุปผลระบบ CCTV Management</h1>
        <p>วันที่ออกรายงาน: ${getThaiDate()}</p>
      </div>

      <div class="summary-grid">
        <div class="card"><h3>กล้องทั้งหมด</h3><p>${stats.cctvTotal}</p></div>
        <div class="card"><h3>กล้องขัดข้อง</h3><p style="color: red;">${stats.cctvOffline}</p></div>
        <div class="card"><h3>งานซ่อมค้าง</h3><p>${stats.maintainPending}</p></div>
        <div class="card"><h3>งานติดตั้งค้าง</h3><p>${stats.installPending}</p></div>
        <div class="card"><h3>สินค้าใกล้หมด</h3><p>${stats.inventoryLow}</p></div>
        <div class="card"><h3>รอส่งเคลม</h3><p>${stats.claimPending}</p></div>
      </div>

      <h3>รายละเอียดกล้องที่ขัดข้อง (Offline)</h3>
      <table>
        <thead>
          <tr><th>ID</th><th>ชื่อกล้อง/สถานที่</th><th>โครงการ</th><th>อัปเดตล่าสุด</th></tr>
        </thead>
        <tbody>
          ${cctvData.length > 0 ? cctvData.map(c => `<tr><td>${c.id}</td><td>${c.name}</td><td>${c.project}</td><td>${c.lastUpdate}</td></tr>`).join('') : '<tr><td colspan="4" style="text-align:center;">ไม่พบกล้องขัดข้อง</td></tr>'}
        </tbody>
      </table>

      <div class="footer">
        <p>ออกรายงานโดยระบบอัตโนมัติ (CCTV Management System)</p>
      </div>
    `;

    const blob = HtmlService.createHtmlOutput(htmlContent).getAs('application/pdf').setName('CCTV_Executive_Summary.pdf');
    const folder = getOrCreateFolder("CCTV_Reports");
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return { success: true, url: file.getUrl() };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function generateMaintainPDF() {
  try {
    const data = getMaintenanceTickets().data;
    if(data.length === 0) return { success: false, error: "ไม่มีข้อมูลให้สร้างรายงาน" };

    let htmlContent = `
      <style>
        body { font-family: 'Tahoma', sans-serif; font-size: 12px; padding: 20px; }
        h2 { text-align: center; color: #333; border-bottom: 2px solid #f59e0b; padding-bottom: 10px; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #fffbeb; color: #b45309; }
        .footer { text-align: right; margin-top: 20px; font-size: 10px; color: #888; }
      </style>
      <h2>รายงานประวัติการแจ้งซ่อม (Maintenance Report)</h2>
      <table>
        <thead>
          <tr>
            <th>Ticket ID</th><th>รหัสกล้อง</th><th>รายละเอียดปัญหา</th><th>สถานะ</th><th>ผู้แจ้ง</th><th>วันที่แจ้ง</th>
          </tr>
        </thead>
        <tbody>
          ${data.map(t => `<tr><td>${t.ticketId}</td><td>${t.camId}</td><td>${t.issue}</td><td><b>${t.status}</b></td><td>${t.reportedBy}</td><td>${t.date}</td></tr>`).join('')}
        </tbody>
      </table>
      <div class="footer">พิมพ์เมื่อ: ${getThaiDate()}</div>
    `;

    const blob = HtmlService.createHtmlOutput(htmlContent).getAs('application/pdf').setName('Maintenance_Report_' + new Date().getTime() + '.pdf');
    const folder = getOrCreateFolder("CCTV_Reports");
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return { success: true, url: file.getUrl() };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function generateClaimPDF() {
  try {
    const data = getClaimData().data;
    if(data.length === 0) return { success: false, error: "ไม่มีข้อมูลให้สร้างรายงาน" };

    let htmlContent = `
      <style>
        body { font-family: 'Tahoma', sans-serif; font-size: 12px; padding: 20px; }
        h2 { text-align: center; color: #333; border-bottom: 2px solid #f97316; padding-bottom: 10px; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #fff7ed; color: #c2410c; }
        .footer { text-align: right; margin-top: 20px; font-size: 10px; color: #888; }
      </style>
      <h2>รายงานการส่งเคลมสินค้า (Claim Report)</h2>
      <table>
        <thead>
          <tr>
            <th>รหัสเคลม</th><th>สินค้า</th><th>S/N</th><th>อาการเสีย</th><th>สถานะ</th><th>ผู้แจ้ง</th><th>วันที่แจ้ง</th>
          </tr>
        </thead>
        <tbody>
          ${data.map(c => `<tr><td>${c.claimId}</td><td>${c.itemName}</td><td>${c.serialNumber}</td><td>${c.issue}</td><td><b>${c.status}</b></td><td>${c.reportedBy}</td><td>${c.claimDate}</td></tr>`).join('')}
        </tbody>
      </table>
      <div class="footer">พิมพ์เมื่อ: ${getThaiDate()}</div>
    `;

    const blob = HtmlService.createHtmlOutput(htmlContent).getAs('application/pdf').setName('Claim_Report_' + new Date().getTime() + '.pdf');
    const folder = getOrCreateFolder("CCTV_Reports");
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return { success: true, url: file.getUrl() };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// **ใหม่** ฟังก์ชันพิมพ์รายงาน Inventory
function generateInventoryPDF() {
  try {
    const data = getInventoryData().data;
    if(data.length === 0) return { success: false, error: "ไม่มีข้อมูลให้สร้างรายงาน" };

    let htmlContent = `
      <style>
        body { font-family: 'Tahoma', sans-serif; font-size: 12px; padding: 20px; }
        h2 { text-align: center; color: #333; border-bottom: 2px solid #6366f1; padding-bottom: 10px; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #eef2ff; color: #4338ca; }
        .footer { text-align: right; margin-top: 20px; font-size: 10px; color: #888; }
      </style>
      <h2>รายงานสต๊อกสินค้าคงคลัง (Inventory Report)</h2>
      <table>
        <thead>
          <tr>
            <th>รหัสสินค้า</th><th>ชื่อสินค้า</th><th>หมวดหมู่</th><th>ทั้งหมด</th><th>พร้อมใช้</th><th>หน่วยนับ</th>
          </tr>
        </thead>
        <tbody>
          ${data.map(item => `<tr><td>${item.itemId}</td><td>${item.itemName}</td><td>${item.category}</td><td style="text-align:center;">${item.totalQty}</td><td style="text-align:center; font-weight:bold; color:${item.availableQty <= item.minStock ? 'red' : 'green'};">${item.availableQty}</td><td style="text-align:center;">${item.unit}</td></tr>`).join('')}
        </tbody>
      </table>
      <div class="footer">พิมพ์เมื่อ: ${getThaiDate()}</div>
    `;

    const blob = HtmlService.createHtmlOutput(htmlContent).getAs('application/pdf').setName('Inventory_Report_' + new Date().getTime() + '.pdf');
    const folder = getOrCreateFolder("CCTV_Reports");
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success: true, url: file.getUrl() };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// **ใหม่** ฟังก์ชันพิมพ์รายงานการยืม-คืน
function generateBorrowPDF() {
  try {
    const data = getBorrowData().data;
    if(data.length === 0) return { success: false, error: "ไม่มีข้อมูลให้สร้างรายงาน" };

    let htmlContent = `
      <style>
        body { font-family: 'Tahoma', sans-serif; font-size: 12px; padding: 20px; }
        h2 { text-align: center; color: #333; border-bottom: 2px solid #8b5cf6; padding-bottom: 10px; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f5f3ff; color: #6d28d9; }
        .footer { text-align: right; margin-top: 20px; font-size: 10px; color: #888; }
      </style>
      <h2>รายงานประวัติการยืม-คืนสินค้า (Borrowing Report)</h2>
      <table>
        <thead>
          <tr>
            <th>รหัสการยืม</th><th>สินค้า</th><th>ผู้ยืม</th><th>จำนวน</th><th>วันที่ยืม</th><th>กำหนดคืน</th><th>สถานะ</th>
          </tr>
        </thead>
        <tbody>
          ${data.map(b => `<tr><td>${b.borrowId}</td><td>${b.itemName}</td><td>${b.borrower}</td><td style="text-align:center;">${b.qty}</td><td>${b.borrowDate}</td><td>${b.expectedReturn}</td><td style="font-weight:bold;">${b.status}</td></tr>`).join('')}
        </tbody>
      </table>
      <div class="footer">พิมพ์เมื่อ: ${getThaiDate()}</div>
    `;

    const blob = HtmlService.createHtmlOutput(htmlContent).getAs('application/pdf').setName('Borrow_Report_' + new Date().getTime() + '.pdf');
    const folder = getOrCreateFolder("CCTV_Reports");
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success: true, url: file.getUrl() };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// **ใหม่** ฟังก์ชันพิมพ์รายงานงานติดตั้ง (แบบรวม)
function generateInstallPDF() {
  try {
    const data = getInstallData().data;
    if(data.length === 0) return { success: false, error: "ไม่มีข้อมูลให้สร้างรายงาน" };

    let htmlContent = `
      <style>
        body { font-family: 'Tahoma', sans-serif; font-size: 12px; padding: 20px; }
        h2 { text-align: center; color: #333; border-bottom: 2px solid #14b8a6; padding-bottom: 10px; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f0fdfa; color: #0f766e; }
        .footer { text-align: right; margin-top: 20px; font-size: 10px; color: #888; }
      </style>
      <h2>รายงานสรุปงานติดตั้ง (Installation Report)</h2>
      <table>
        <thead>
          <tr>
            <th>Install ID</th><th>ลูกค้า/สถานที่</th><th>รายละเอียดงาน</th><th>ผู้แจ้ง</th><th>วันที่แจ้ง</th><th>สถานะ</th>
          </tr>
        </thead>
        <tbody>
          ${data.map(t => `<tr><td>${t.installId}</td><td>${t.customerName}</td><td>${t.installDetails}</td><td>${t.reportedBy}</td><td>${t.reportDate}</td><td><b>${t.status}</b></td></tr>`).join('')}
        </tbody>
      </table>
      <div class="footer">พิมพ์เมื่อ: ${getThaiDate()}</div>
    `;

    const blob = HtmlService.createHtmlOutput(htmlContent).getAs('application/pdf').setName('Install_Report_' + new Date().getTime() + '.pdf');
    const folder = getOrCreateFolder("CCTV_Reports");
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success: true, url: file.getUrl() };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// ==========================================
// ฟังก์ชันช่วยแปลงลิงก์ภาพใน Drive เป็น Base64 เพื่อฝังใน PDF
// ==========================================
/**
 * ปรับปรุงการดึงรูปภาพ Base64 ให้เร็วขึ้นและกินทรัพยากรน้อยลง
 */
// ==========================================
// ฟังก์ชันช่วยแปลงลิงก์ภาพใน Drive เป็น Base64 (ปรับปรุง: w200 + CacheService)
// ==========================================
// ==========================================
// ฟังก์ชันช่วยแปลงลิงก์ภาพใน Drive เป็น Base64 (เวอร์ชันแก้รูปไม่ขึ้น + ระบบ Cache)
// ==========================================
function getBase64ImgFromUrl(url) {
  if (!url) return "";
  
  try {
    let fileId = "";
    // ดึง File ID จาก URL รูปภาพ
    let idMatch = url.match(/id=([^&]+)/);
    if (idMatch && idMatch[1]) {
      fileId = idMatch[1];
    } else {
      let dMatch = url.match(/\/d\/([^\/]+)/);
      if (dMatch) fileId = dMatch[1];
    }

    if (fileId) {
      // 1. ตรวจสอบ Cache ก่อน (ใช้ชื่อ IMG_V2_ เพื่อข้ามรูปที่พังจากโค้ดเวอร์ชันก่อนหน้า)
      const cache = CacheService.getScriptCache();
      const cacheKey = "IMG_V2_" + fileId; 
      const cachedBase64 = cache.get(cacheKey);
      
      if (cachedBase64) {
        return cachedBase64;
      }

      // 2. ใช้ DriveApp ดึงไฟล์ (วิธีนี้ชัวร์ 100% ว่าได้รูปแน่นอน ไม่ติดหน้า Login ของ Google)
      let file = DriveApp.getFileById(fileId);
      let blob = file.getBlob();
      let base64 = Utilities.base64Encode(blob.getBytes());
      let mimeType = blob.getContentType();
      
      // ตรวจสอบว่าเป็นรูปภาพจริงๆ ก่อน
      if (mimeType.indexOf('image') !== -1) {
        let finalBase64Str = "data:" + mimeType + ";base64," + base64;
        
        // 3. เก็บลง Cache (รับขนาดสูงสุด 100KB)
        try {
          if (finalBase64Str.length <= 100000) {
            cache.put(cacheKey, finalBase64Str, 21600); // เก็บไว้ 6 ชั่วโมง
          }
        } catch (cacheErr) {
          console.warn("รูปใหญ่เกินไป ไม่สามารถทำ Cache ได้: " + fileId);
        }

        return finalBase64Str;
      }
    }
  } catch (e) {
    console.error("Error in getBase64ImgFromUrl: " + e.toString());
  }
  
  return url; // ถ้าแปลงไม่ได้ ให้ส่ง URL กลับไปเผื่อระบบอ่านได้
}
function generateSingleTicketPDF(ticketId) {
  try {
    const ss = getActiveSS();
    const tickets = getMaintenanceTickets().data;
    const ticket = tickets.find(t => t.ticketId === ticketId);
    
    if (!ticket) return { success: false, error: "ไม่พบข้อมูลใบงานนี้" };

    const cams = getCCTVData().data;
    const camInfo = cams.find(c => c.id === ticket.camId) || { name: "ไม่ระบุ", project: "-" };

    let imageHtml = "";
    if (ticket.images && ticket.images.trim() !== "") {
      const urls = ticket.images.split(',');
      imageHtml = '<div class="image-section"><h3>รูปภาพประกอบ:</h3><div class="image-gallery">';
      urls.forEach((url, index) => {
        if(url.trim() !== "") {
          let base64Src = getBase64ImgFromUrl(url.trim());
          imageHtml += `<img src="${base64Src}" alt="รูปที่ ${index+1}" />`;
        }
      });
      imageHtml += '</div></div>';
    }

    let htmlContent = `
      <style>
        body { font-family: 'Tahoma', sans-serif; padding: 20px 30px; color: #333; line-height: 1.5; font-size: 13px; }
        .header-box { text-align: center; border: 2px solid #333; padding: 8px; margin-bottom: 15px; background-color: #f8fafc; }
        .ticket-title { font-size: 20px; font-weight: bold; margin: 0; }
        .info-table { width: 100%; border-collapse: collapse; margin-bottom: 15px; }
        .info-table td { padding: 6px 8px; border-bottom: 1px solid #eee; vertical-align: top; }
        .label { font-weight: bold; width: 25%; color: #555; }
        .value { width: 75%; }
        
        .image-section { margin-top: 15px; border-top: 2px solid #eee; padding-top: 10px; page-break-inside: avoid; }
        .image-section h3 { margin: 0 0 10px 0; font-size: 14px; }
        .image-gallery { display: block; text-align: center; }
        .image-gallery img { 
          max-width: 32%; max-height: 200px; width: auto; height: auto; object-fit: contain;
          border: 1px solid #ddd; border-radius: 6px; margin: 2px; display: inline-block; vertical-align: middle;
        }

        .footer { margin-top: 30px; text-align: right; font-size: 12px; }
        .status-badge { display: inline-block; padding: 3px 10px; border-radius: 4px; background: #eee; font-weight: bold; }
      </style>
      
      <div class="header-box">
        <p class="ticket-title">ใบแจ้งซ่อมบำรุงอุปกรณ์ (Work Order)</p>
        <p style="margin: 5px 0 0 0;">ID: ${ticket.ticketId}</p>
      </div>

      <table class="info-table">
        <tr><td class="label">วันที่แจ้ง:</td><td class="value">${ticket.date}</td></tr>
        <tr><td class="label">สถานะปัจจุบัน:</td><td class="value"><span class="status-badge">${ticket.status}</span></td></tr>
        <tr><td class="label">รหัสกล้อง:</td><td class="value"><b>${ticket.camId}</b></td></tr>
        <tr><td class="label">สถานที่ติดตั้ง:</td><td class="value">${camInfo.name}</td></tr>
        <tr><td class="label">โครงการ:</td><td class="value">${camInfo.project}</td></tr>
        <tr><td class="label">รายละเอียดปัญหา:</td><td class="value">${ticket.issue}</td></tr>
        <tr><td class="label">ผู้รายงาน:</td><td class="value">${ticket.reportedBy}</td></tr>
      </table>

      ${imageHtml}

      <div class="footer">
        <p>-------------------------------------------</p>
        <p>เจ้าหน้าที่ผู้รับผิดชอบ (ลงชื่อ)</p>
        <br>
        <p>วันที่ดำเนินการสำเร็จ: ____/____/____</p>
      </div>
    `;

    const blob = HtmlService.createHtmlOutput(htmlContent).getAs('application/pdf').setName(`WorkOrder_${ticket.ticketId}.pdf`);
    const folder = getOrCreateFolder("CCTV_Reports");
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return { success: true, url: file.getUrl() };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// **ใหม่** ฟังก์ชันพิมพ์ใบงานติดตั้ง (แบบเดี่ยว)
function generateSingleInstallPDF(installId) {
  try {
    const dataList = getInstallData().data;
    const work = dataList.find(t => t.installId === installId);
    
    if (!work) return { success: false, error: "ไม่พบข้อมูลใบงานติดตั้งนี้" };

    let imageHtml = "";
    if (work.images && work.images.trim() !== "") {
      const urls = work.images.split(',');
      imageHtml = '<div class="image-section"><h3>รูปภาพสถานที่/หน้างาน:</h3><div class="image-gallery">';
      urls.forEach((url, index) => {
        if(url.trim() !== "") {
          let base64Src = getBase64ImgFromUrl(url.trim());
          imageHtml += `<img src="${base64Src}" alt="รูปที่ ${index+1}" />`;
        }
      });
      imageHtml += '</div></div>';
    }

    let htmlContent = `
      <style>
        body { font-family: 'Tahoma', sans-serif; padding: 20px 30px; color: #333; line-height: 1.5; font-size: 13px; }
        .header-box { text-align: center; border: 2px solid #0f766e; padding: 8px; margin-bottom: 15px; background-color: #f0fdfa; }
        .ticket-title { font-size: 20px; font-weight: bold; margin: 0; color: #0f766e; }
        .info-table { width: 100%; border-collapse: collapse; margin-bottom: 15px; }
        .info-table td { padding: 6px 8px; border-bottom: 1px solid #eee; vertical-align: top; }
        .label { font-weight: bold; width: 25%; color: #555; }
        .value { width: 75%; }
        
        .image-section { margin-top: 15px; border-top: 2px solid #eee; padding-top: 10px; page-break-inside: avoid; }
        .image-section h3 { margin: 0 0 10px 0; font-size: 14px; }
        .image-gallery { display: block; text-align: center; }
        .image-gallery img { max-width: 32%; max-height: 200px; width: auto; height: auto; object-fit: contain; border: 1px solid #ddd; border-radius: 6px; margin: 2px; display: inline-block; vertical-align: middle; }
        .footer { margin-top: 40px; text-align: right; font-size: 12px; }
      </style>
      
      <div class="header-box">
        <p class="ticket-title">ใบงานติดตั้งอุปกรณ์ (Installation Work Order)</p>
        <p style="margin: 5px 0 0 0;">ID: ${work.installId}</p>
      </div>

      <table class="info-table">
        <tr><td class="label">วันที่รับเรื่อง:</td><td class="value">${work.reportDate}</td></tr>
        <tr><td class="label">สถานะงาน:</td><td class="value"><b>${work.status}</b></td></tr>
        <tr><td class="label">ลูกค้า / สถานที่:</td><td class="value">${work.customerName}</td></tr>
        <tr><td class="label">ข้อมูลติดต่อ:</td><td class="value">${work.contactInfo}</td></tr>
        <tr><td class="label">รายละเอียดการติดตั้ง:</td><td class="value">${work.installDetails}</td></tr>
        <tr><td class="label">ผู้รับเรื่อง/ผู้แจ้ง:</td><td class="value">${work.reportedBy}</td></tr>
      </table>

      ${imageHtml}

      <table style="width: 100%; margin-top: 40px; text-align: center;">
        <tr>
          <td>
             <p>....................................................</p>
             <p>ลูกค้า / ผู้ตรวจรับงาน</p>
             <p>วันที่: ______/______/______</p>
          </td>
          <td>
             <p>....................................................</p>
             <p>ช่างผู้ดำเนินงาน</p>
             <p>วันที่: ______/______/______</p>
          </td>
        </tr>
      </table>
    `;

    const blob = HtmlService.createHtmlOutput(htmlContent).getAs('application/pdf').setName(`InstallOrder_${work.installId}.pdf`);
    const folder = getOrCreateFolder("CCTV_Reports");
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success: true, url: file.getUrl() };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function generateSingleClaimPDF(claimId) {
  try {
    const ss = getActiveSS();
    const claims = getClaimData().data;
    const claim = claims.find(c => c.claimId === claimId);
    
    if (!claim) return { success: false, error: "ไม่พบข้อมูลใบส่งเคลมนี้" };

    let htmlContent = `
      <style>
        body { font-family: 'Tahoma', sans-serif; padding: 20px 30px; color: #333; line-height: 1.5; font-size: 13px; }
        .header-box { text-align: center; border: 2px solid #f97316; padding: 8px; margin-bottom: 15px; background-color: #fff7ed; }
        .ticket-title { font-size: 20px; font-weight: bold; margin: 0; color: #ea580c; }
        .info-table { width: 100%; border-collapse: collapse; margin-bottom: 15px; }
        .info-table td { padding: 6px 8px; border-bottom: 1px solid #eee; vertical-align: top; }
        .label { font-weight: bold; width: 25%; color: #555; }
        .value { width: 75%; }
        .footer { margin-top: 40px; text-align: right; font-size: 12px; }
        .status-badge { display: inline-block; padding: 3px 10px; border-radius: 4px; background: #ffedd5; color: #ea580c; font-weight: bold; border: 1px solid #fed7aa; }
      </style>
      
      <div class="header-box">
        <p class="ticket-title">ใบส่งเคลมสินค้า (Product Claim Form)</p>
        <p style="margin: 5px 0 0 0;">ID: ${claim.claimId}</p>
      </div>

      <table class="info-table">
        <tr><td class="label">วันที่ส่งเคลม:</td><td class="value">${claim.claimDate}</td></tr>
        <tr><td class="label">สถานะการเคลม:</td><td class="value"><span class="status-badge">${claim.status}</span></td></tr>
        <tr><td class="label">ชื่อสินค้าชำรุด:</td><td class="value"><b>${claim.itemName}</b></td></tr>
        <tr><td class="label">รหัสสินค้า (Item ID):</td><td class="value">${claim.itemId}</td></tr>
        <tr><td class="label">Serial Number (S/N):</td><td class="value">${claim.serialNumber || '-'}</td></tr>
        <tr><td class="label">อาการเสีย / รายละเอียด:</td><td class="value">${claim.issue}</td></tr>
        <tr><td class="label">ผู้รายงาน / ผู้ส่งเคลม:</td><td class="value">${claim.reportedBy}</td></tr>
        <tr><td class="label">หมายเหตุเพิ่มเติม:</td><td class="value">${claim.note || '-'}</td></tr>
        <tr><td class="label">อัปเดตล่าสุดเมื่อ:</td><td class="value">${claim.updateDate || '-'}</td></tr>
      </table>

      <table style="width: 100%; margin-top: 50px; text-align: center;">
        <tr>
          <td>
             <p>....................................................</p>
             <p>ผู้แจ้งเคลม / ผู้ส่งมอบ</p>
             <p>วันที่: ______/______/______</p>
          </td>
          <td>
             <p>....................................................</p>
             <p>ผู้รับสินค้าเคลม / ช่างเทคนิค</p>
             <p>วันที่: ______/______/______</p>
          </td>
        </tr>
      </table>
    `;

    const blob = HtmlService.createHtmlOutput(htmlContent).getAs('application/pdf').setName(`ClaimTicket_${claim.claimId}.pdf`);
    const folder = getOrCreateFolder("CCTV_Reports");
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success: true, url: file.getUrl() };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function generateSingleBorrowPDF(borrowId) {
  try {
    const ss = getActiveSS();
    const borrows = getBorrowData().data;
    const borrow = borrows.find(b => b.borrowId === borrowId);
    
    if (!borrow) return { success: false, error: "ไม่พบข้อมูลการยืม/คืนนี้" };

    let htmlContent = `
      <style>
        body { font-family: 'Tahoma', sans-serif; padding: 20px 30px; color: #333; line-height: 1.5; font-size: 13px; }
        .header-box { text-align: center; border: 2px solid #8b5cf6; padding: 8px; margin-bottom: 15px; background-color: #f5f3ff; }
        .ticket-title { font-size: 20px; font-weight: bold; margin: 0; color: #7c3aed; }
        .info-table { width: 100%; border-collapse: collapse; margin-bottom: 15px; }
        .info-table td { padding: 6px 8px; border-bottom: 1px solid #eee; vertical-align: top; }
        .label { font-weight: bold; width: 25%; color: #555; }
        .value { width: 75%; }
        .footer { margin-top: 40px; text-align: right; font-size: 12px; }
        .status-badge { display: inline-block; padding: 3px 10px; border-radius: 4px; background: #ede9fe; color: #7c3aed; font-weight: bold; border: 1px solid #ddd6fe; }
      </style>
      
      <div class="header-box">
        <p class="ticket-title">ใบยืม-คืนอุปกรณ์ (Borrow-Return Form)</p>
        <p style="margin: 5px 0 0 0;">ID: ${borrow.borrowId}</p>
      </div>

      <table class="info-table">
        <tr><td class="label">วันที่ยืมอุปกรณ์:</td><td class="value">${borrow.borrowDate}</td></tr>
        <tr><td class="label">กำหนดส่งคืน:</td><td class="value"><b>${borrow.expectedReturn}</b></td></tr>
        <tr><td class="label">สถานะปัจจุบัน:</td><td class="value"><span class="status-badge">${borrow.status}</span></td></tr>
        <tr><td class="label">ชื่ออุปกรณ์:</td><td class="value"><b>${borrow.itemName}</b></td></tr>
        <tr><td class="label">รหัสอุปกรณ์ (Item ID):</td><td class="value">${borrow.itemId}</td></tr>
        <tr><td class="label">จำนวนที่ยืม:</td><td class="value"><b>${borrow.qty} หน่วย</b></td></tr>
        <tr><td class="label">ผู้ขอเข้ารับการยืม:</td><td class="value">${borrow.borrower}</td></tr>
        <tr><td class="label">วันที่คืนสินค้าจริง:</td><td class="value">${borrow.actualReturn || '-'}</td></tr>
        <tr><td class="label">หมายเหตุการยืม:</td><td class="value">${borrow.note || '-'}</td></tr>
      </table>

      <table style="width: 100%; margin-top: 50px; text-align: center;">
        <tr>
          <td>
             <p>....................................................</p>
             <p>ผู้ยืมอุปกรณ์ (ลงชื่อ)</p>
             <p>วันที่: ______/______/______</p>
          </td>
          <td>
             <p>....................................................</p>
             <p>ผู้ส่งมอบ / เจ้าหน้าที่คลัง</p>
             <p>วันที่: ______/______/______</p>
          </td>
        </tr>
      </table>
    `;

    const blob = HtmlService.createHtmlOutput(htmlContent).getAs('application/pdf').setName(`BorrowSlip_${borrow.borrowId}.pdf`);
    const folder = getOrCreateFolder("CCTV_Reports");
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success: true, url: file.getUrl() };
  } catch (err) { return { success: false, error: err.toString() }; }
}


// ==========================================
// Utils Functions & Delete Operations
// ==========================================
function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

// ==========================================
// ฟังก์ชันบันทึกประวัติการทำงาน (Audit Log)
// ==========================================
function writeLog(username, action, details) {
  try {
    const sheet = getActiveSS().getSheetByName(SHEET_LOG);
    if (sheet) {
      sheet.appendRow([getThaiDate(), username, action, details]);
    }
  } catch (err) {
    console.error("Log Error: ", err);
  }
}

// ==========================================
// ฟังก์ชันลบข้อมูล (แก้ไขให้รองรับการส่งชื่อผู้ลบมาเก็บ Log)
// ==========================================
function deleteCCTV(id, user) { return deleteRowByFieldValue(SHEET_DATA, 0, id, "ลบข้อมูลกล้อง CCTV", user); }
function deleteMaintenanceTicket(ticketId, user) { return deleteRowByFieldValue(SHEET_MAINTAIN, 0, ticketId, "ลบใบแจ้งซ่อม", user); }
function deleteBorrowRecord(borrowId, user) { return deleteRowByFieldValue(SHEET_BORROW, 0, borrowId, "ลบประวัติการยืม", user); }
function deleteClaimRecord(claimId, user) { return deleteRowByFieldValue(SHEET_CLAIM, 0, claimId, "ลบประวัติการเคลม", user); }
function deleteInventoryItem(itemId, user) { return deleteRowByFieldValue(SHEET_INVENTORY, 0, itemId, "ลบสินค้าคงคลัง", user); }
function deleteUserData(username, user) { return deleteRowByFieldValue(SHEET_ADMIN, 0, username, "ลบบัญชีผู้ใช้", user); }
function deleteInstallRecord(installId, user) { return deleteRowByFieldValue(SHEET_INSTALL, 0, installId, "ลบใบงานติดตั้ง", user); }

function deleteRowByFieldValue(sheetName, colIndex, value, actionName, user) {
  try {
    const sheet = getActiveSS().getSheetByName(sheetName);
    if (!sheet) return { success: false, error: "ไม่พบฐานข้อมูล " + sheetName };

    const data = sheet.getDataRange().getValues();
    const searchValue = value.toString().trim();

    for (let i = 1; i < data.length; i++) {
      if (data[i][colIndex] !== "" && data[i][colIndex].toString().trim() === searchValue) {
        sheet.deleteRow(i + 1);
        
        // ++ สั่งบันทึก Log หลังจากลบข้อมูลสำเร็จ ++
        const executor = user ? user : "System/Unknown";
        writeLog(executor, actionName, `ลบข้อมูลรหัสอ้างอิง: ${searchValue}`);
        
        return { success: true, message: "ลบข้อมูลสำเร็จ" };
      }
    }
    return { success: false, error: "ไม่พบข้อมูลที่ต้องการลบ" };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ==========================================
// Utils: ตัวช่วยจัดการเวลาให้เป็นเวลาไทย (Asia/Bangkok)
// ==========================================
function getThaiDate() {
  return Utilities.formatDate(new Date(), "Asia/Bangkok", "dd/MM/yyyy HH:mm:ss");
}

// ==========================================
// ฟังก์ชันนำเข้าข้อมูล CSV จาก Google Drive Folder อัตโนมัติ
// ==========================================
function importCsvFromFolder() {
  const folderId = "1uS5hRGQpSJ9Z1XzVPNfu8neu7HRYj1bx";
  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    const ss = getActiveSS();
    
    // รายชื่อตารางในระบบ
    const dbSheets = [
      SHEET_DATA, SHEET_ADMIN, SHEET_MAINTAIN, SHEET_INVENTORY, 
      SHEET_BORROW, SHEET_CLAIM, SHEET_INSTALL, SHEET_LOG
    ];
    
    let importCount = 0;
    
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      
      // ค้นหาว่าชื่อไฟล์ตรงกับตารางใดในระบบ
      let targetSheetName = "";
      for (let name of dbSheets) {
        if (fileName.indexOf(name) !== -1) {
          targetSheetName = name;
          break;
        }
      }
      
      if (targetSheetName !== "") {
        let sheet = ss.getSheetByName(targetSheetName);
        if (!sheet) {
          sheet = ss.insertSheet(targetSheetName);
        } else {
          sheet.clear();
        }
        
        // อ่านข้อมูล CSV โดยรองรับภาษาไทย (พยายามถอดรหัสเป็น UTF-8 หรือ TIS-620)
        let blob = file.getBlob();
        let csvContent = "";
        try {
          csvContent = blob.getDataAsString("UTF-8");
        } catch(e) {
          csvContent = blob.getDataAsString("TIS-620");
        }
        
        const csvData = Utilities.parseCsv(csvContent);
        
        if (csvData.length > 0) {
          sheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
          importCount++;
        }
      }
    }
    
    return { success: true, message: "นำเข้าข้อมูลสำเร็จทั้งหมด " + importCount + " ตาราง" };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ==========================================
// ฟังก์ชันช่วยดึงสิทธิ์การเข้าถึง Spreadsheet (รองรับทั้งแบบผูกติดและแบบแยกเดี่ยว)
// ==========================================
function getActiveSS() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
  } catch (e) {}
  return SpreadsheetApp.openById("1X-vMNivhxv8CjA1koTJqDNXt70-k-Rb_bVFhbNKhBVI");
}

// ==========================================
// ฟังก์ชัน doPost (API Gateway) รองรับการเรียกใช้ฟังก์ชันหลังบ้านจากโดเมนภายนอก (เช่น Vercel)
// ==========================================
function doPost(e) {
  try {
    const request = JSON.parse(e.postData.contents);
    const functionName = request.functionName;
    const args = request.arguments || [];
    
    // เรียกใช้งานฟังก์ชันที่ระบุแบบไดนามิก
    const result = this[functionName].apply(this, args);
    
    return ContentService.createTextOutput(JSON.stringify({ success: true, result: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
