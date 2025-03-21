function doGet() {
  const html = HtmlService.createTemplateFromFile('Index');
  return html.evaluate()
    .setTitle('บันทึกการแก้ไฟฟ้าขัดข้อง กฟส.วัดสิงห์')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setFaviconUrl('https://cdn.jsdelivr.net/gh/Aonint12/images/PEA114.png')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getAllRecords(limit, filterMonth, filterYear, filterProtection) {
  try {
    const ss = SpreadsheetApp.openById('14LYcoB2AIn2K-GPZp0JqISjPDo2dHjoIw73eBEZgTO8');
    const sheet = ss.getActiveSheet();
    const data = sheet.getDataRange().getValues();
    
    Logger.log('Data length: ' + data.length);
    if (data.length <= 1) {
      return { success: true, records: [], pendingCount: 0, message: 'ยังไม่มีข้อมูลบันทึก' };
    }
    
    let records = data.slice(1).map((row, index) => {
      return {
        id: index + 1,
        timestamp: row[0] ? new Date(row[0]).toISOString() : '',
        outageStartTime: row[1] ? new Date(row[1]).toISOString() : '',
        powerRestoreTime: row[2] ? new Date(row[2]).toISOString() : '',
        officeReturnTime: row[3] ? new Date(row[3]).toISOString() : '',
        location: row[4] || '',
        protectionEquipment: row[5] || '',
        transformerPEA: row[6] || '',
        cause: row[7] || '',
        phase: row[8] || '',
        equipmentDetails: row[9] || '',
        operator: row[10] || '',
        weatherCondition: row[11] || '',
        notes: row[12] || '',
        orderNumber: row[13] || '',
        transformerSize: row[14] || ''
      };
    });
    
    if (filterMonth && filterYear) {
      records = records.filter(record => {
        if (!record.outageStartTime) return false;
        const date = new Date(record.outageStartTime);
        const thaiYear = date.getFullYear() + 543;
        return date.getMonth() + 1 === parseInt(filterMonth) && thaiYear === parseInt(filterYear);
      });
    }
    
    if (filterProtection) {
      records = records.filter(record => {
        if (filterProtection.type === 'หม้อแปลง' && filterProtection.value) {
          return record.protectionEquipment === 'หม้อแปลง' && record.transformerPEA === filterProtection.value;
        } else if (filterProtection.type === 'ดรอฟเอาร์ฟิวส์' && filterProtection.value) {
          return record.protectionEquipment === 'ดรอฟเอาร์ฟิวส์' && record.transformerPEA === filterProtection.value;
        } else if (filterProtection.type === 'เบรกเกอร์' && filterProtection.value) {
          return record.protectionEquipment === 'เบรกเกอร์' && record.transformerPEA === filterProtection.value;
        }
        return record.protectionEquipment === filterProtection.type;
      });
    }
    
    records = records.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    records = records.slice(0, limit || records.length);
    records = records.map((record, index) => ({ ...record, id: index + 1 }));
    
    const pendingCount = records.filter(record => !record.orderNumber).length;
    return { success: true, records: records, pendingCount: pendingCount };
  } catch (error) {
    Logger.log('Error in getAllRecords: ' + error.toString());
    return { success: false, message: 'เกิดข้อผิดพลาดในการดึงข้อมูล: ' + error.toString() };
  }
}

function updateOrderNumber(recordId, orderNumber) {
  try {
    const ss = SpreadsheetApp.openById('14LYcoB2AIn2K-GPZp0JqISjPDo2dHjoIw73eBEZgTO8');
    const sheet = ss.getActiveSheet();
    const data = sheet.getDataRange().getValues();
    const totalRows = data.length - 1;
    const rowIndex = totalRows - recordId;
    if (rowIndex >= 0 && rowIndex < totalRows) {
      sheet.getRange(rowIndex + 2, 14).setValue(orderNumber);
      return { success: true, message: 'อัปเดตเลขใบสั่งสำเร็จ' };
    } else {
      return { success: false, message: 'ไม่พบรายการที่ต้องการอัปเดต' };
    }
  } catch (error) {
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}

function getOptions() {
  try {
    const ss = SpreadsheetApp.openById('14LYcoB2AIn2K-GPZp0JqISjPDo2dHjoIw73eBEZgTO8');
    const equipmentSheet = ss.getSheetByName('อุปกรณ์');
    const optionsSheet = ss.getSheetByName('ตัวเลือก');
    const transformerSheet = ss.getSheetByName('Transformers');
    
    if (!equipmentSheet) throw new Error('ไม่พบชีท "อุปกรณ์"');
    if (!optionsSheet) throw new Error('ไม่พบชีท "ตัวเลือก"');
    if (!transformerSheet) throw new Error('ไม่พบชีท "Transformers"');
    
    const equipmentData = equipmentSheet.getDataRange().getValues();
    const equipmentTypes = [];
    const equipmentUnits = {};
    for (let i = 1; i < equipmentData.length; i++) {
      if (equipmentData[i][0]) {
        equipmentTypes.push(equipmentData[i][0]);
        equipmentUnits[equipmentData[i][0]] = equipmentData[i][1] || '';
      }
    }
    
    const optionsData = optionsSheet.getDataRange().getValues();
    const protectionEquipment = ['หม้อแปลง', 'ดรอฟเอาร์ฟิวส์', 'เบรกเกอร์'];
    const operators = [];
    const causes = [];
    for (let i = 1; i < optionsData.length; i++) {
      if (optionsData[i][0] && !causes.includes(optionsData[i][0])) causes.push(optionsData[i][0]);
      if (optionsData[i][2] && !operators.includes(optionsData[i][2])) operators.push(optionsData[i][2]);
    }
    
    const transformerData = transformerSheet.getDataRange().getValues();
    const transformerPEAs = [];
    const dropFuses = [];
    const breakers = [];
    const transformerLocations = {};
    const transformerSizes = {};
    for (let i = 1; i < transformerData.length; i++) {
      if (transformerData[i][0]) {
        transformerPEAs.push(transformerData[i][0]);
        transformerLocations[transformerData[i][0]] = transformerData[i][1] || '';
        transformerSizes[transformerData[i][0]] = transformerData[i][2] || '';
      }
      if (transformerData[i][3]) dropFuses.push(transformerData[i][3]);
      if (transformerData[i][4]) breakers.push(transformerData[i][4]);
    }
    
    return {
      success: true,
      equipmentTypes: equipmentTypes,
      equipmentUnits: equipmentUnits,
      protectionEquipment: protectionEquipment,
      operators: operators,
      causes: causes,
      transformerPEAs: transformerPEAs,
      dropFuses: dropFuses,
      breakers: breakers,
      transformerLocations: transformerLocations,
      transformerSizes: transformerSizes
    };
  } catch (error) {
    Logger.log('Error in getOptions: ' + error.toString());
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}

function submitPowerOutageData(data) {
  try {
    const ss = SpreadsheetApp.openById('14LYcoB2AIn2K-GPZp0JqISjPDo2dHjoIw73eBEZgTO8');
    const sheet = ss.getActiveSheet();
    const transformerSheet = ss.getSheetByName('Transformers');
    const timestamp = new Date();
    
    if (!data.outageStartTime && data.transformerPEA) {
      const transformerData = transformerSheet.getDataRange().getValues();
      const existingPEAs = transformerData.slice(1).map(row => row[0]);
      if (!existingPEAs.includes(data.transformerPEA)) {
        transformerSheet.appendRow([data.transformerPEA, data.location, data.transformerSize, '', '']);
      }
      return { success: true, message: 'บันทึกหม้อแปลงใหม่สำเร็จ' };
    }
    
    let transformerPEA = data.transformerPEA || '';
    if (['หม้อแปลง', 'ดรอฟเอาร์ฟิวส์', 'เบรกเกอร์'].includes(data.protectionEquipment) && !transformerPEA) {
      return { success: false, message: `กรุณาระบุรหัสสำหรับ ${data.protectionEquipment}` };
    }
    
    const rowData = [
      timestamp,
      data.outageStartTime,
      data.powerRestoreTime,
      data.officeReturnTime,
      data.location,
      data.protectionEquipment,
      transformerPEA,
      data.cause,
      data.phase,
      data.equipmentDetails,
      data.operator,
      data.weatherCondition,
      data.notes,
      data.orderNumber || '',
      data.transformerSize || ''
    ];

    sheet.appendRow(rowData);
    return { success: true, message: 'บันทึกข้อมูลสำเร็จ' };
  } catch (error) {
    Logger.log('Error in submitPowerOutageData: ' + error.toString());
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}

function addNewCause(newCause) {
  try {
    const ss = SpreadsheetApp.openById('14LYcoB2AIn2K-GPZp0JqISjPDo2dHjoIw73eBEZgTO8');
    const optionsSheet = ss.getSheetByName('ตัวเลือก');
    const data = optionsSheet.getDataRange().getValues();
    const existingCauses = data.slice(1).map(row => row[0]);

    if (!existingCauses.includes(newCause)) {
      optionsSheet.appendRow([newCause, '', '']);
    }
    return { success: true, message: 'บันทึกสาเหตุใหม่สำเร็จ' };
  } catch (error) {
    Logger.log('Error in addNewCause: ' + error.toString());
    return { success: false, message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}

function checkPassword(password) {
  try {
    const ss = SpreadsheetApp.openById('14LYcoB2AIn2K-GPZp0JqISjPDo2dHjoIw73eBEZgTO8');
    const passwordSheet = ss.getSheetByName('รหัสผ่าน');
    
    if (!passwordSheet) {
      return { success: false, message: 'ไม่พบชีท "รหัสผ่าน"' };
    }
    
    const data = passwordSheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: false, message: 'ไม่พบข้อมูลรหัสผ่าน' };
    }
    
    const passwords = data.slice(1).reduce((acc, row) => {
      acc[row[0]] = { isValid: true, branch: row[1] || '', username: row[2] || '' };
      return acc;
    }, {});
    
    const result = passwords[password];
    if (result) {
      return { success: true, isValid: true, branch: result.branch, username: result.username };
    } else {
      return { success: true, isValid: false, message: 'รหัสผ่านไม่ถูกต้อง' };
    }
  } catch (error) {
    Logger.log('Error in checkPassword: ' + error.toString());
    return { success: false, message: 'เกิดข้อผิดพลาดในการตรวจสอบรหัสผ่าน: ' + error.toString() };
  }
}