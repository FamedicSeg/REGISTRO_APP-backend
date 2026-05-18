const { getAccessToken } = require('../auth/microsoftAuth');
const axios = require('axios');
const ExcelJS = require('exceljs');

// Leer archivo Excel desde OneDrive y devolver como Workbook de ExcelJS
async function getWorkbookFromOneDrive(filePath) {
  try {
    const token = await getAccessToken();
    
    // URL para descargar el contenido del archivo
    const url = `https://graph.microsoft.com/v1.0/me/drive/root:${filePath}:/content`;
    
    const response = await axios.get(url, {
      headers: {
        'Authorization': `Bearer ${token}`
      },
      responseType: 'arraybuffer'
    });

    // Cargar el buffer en ExcelJS
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(response.data);
    
    return workbook;
  } catch (error) {
    console.error('Error leyendo archivo de OneDrive:', error.response?.data || error.message);
    throw new Error(`No se pudo leer el archivo: ${filePath}`);
  }
}

// Guardar Workbook de ExcelJS en OneDrive
async function saveWorkbookToOneDrive(filePath, workbook) {
  try {
    const token = await getAccessToken();
    
    // Convertir workbook a buffer
    const buffer = await workbook.xlsx.writeBuffer();
    
    // URL para subir/actualizar el archivo
    const url = `https://graph.microsoft.com/v1.0/me/drive/root:${filePath}:/content`;
    
    const response = await axios.put(url, buffer, {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      }
    });
    
    return response.data;
  } catch (error) {
    console.error('Error guardando archivo en OneDrive:', error.response?.data || error.message);
    throw new Error(`No se pudo guardar el archivo: ${filePath}`);
  }
}

// Leer datos de una hoja específica (devuelve array de objetos)
async function readSheetData(filePath, sheetName, hasHeaders = true) {
  const workbook = await getWorkbookFromOneDrive(filePath);
  const worksheet = workbook.getWorksheet(sheetName);
  
  if (!worksheet) {
    throw new Error(`Hoja "${sheetName}" no encontrada en el archivo`);
  }
  
  const data = [];
  const rows = worksheet.getWorksheet().getSheetValues();
  
  // Obtener headers si existen
  let headers = [];
  if (hasHeaders) {
    const headerRow = worksheet.getRow(1);
    headers = headerRow.values.slice(1);
  }
  
  // Leer filas
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (hasHeaders && rowNumber === 1) return; // Saltar headers
    
    const rowData = {};
    const values = row.values.slice(1);
    
    if (hasHeaders) {
      headers.forEach((header, index) => {
        rowData[header] = values[index];
      });
    } else {
      rowData[`col_${rowNumber}`] = values;
    }
    
    data.push(rowData);
  });
  
  return data;
}

// Agregar una nueva fila a una hoja específica
async function appendRowToSheet(filePath, sheetName, rowData) {
  const workbook = await getWorkbookFromOneDrive(filePath);
  const worksheet = workbook.getWorksheet(sheetName);
  
  if (!worksheet) {
    throw new Error(`Hoja "${sheetName}" no encontrada en el archivo`);
  }
  
  // Encontrar la última fila con datos
  const lastRow = worksheet.lastRow;
  const newRowNumber = lastRow ? lastRow.number + 1 : 1;
  
  // Agregar nueva fila
  const newRow = worksheet.addRow(Object.values(rowData));
  
  // Guardar cambios
  await saveWorkbookToOneDrive(filePath, workbook);
  
  return {
    success: true,
    rowNumber: newRow.number,
    data: rowData
  };
}

// Buscar filas que cumplan una condición
async function searchInSheet(filePath, sheetName, searchField, searchValue) {
  const data = await readSheetData(filePath, sheetName, true);
  
  const results = data.filter(row => {
    const fieldValue = row[searchField];
    // Búsqueda case-insensitive
    if (typeof fieldValue === 'string' && typeof searchValue === 'string') {
      return fieldValue.toLowerCase().includes(searchValue.toLowerCase());
    }
    return fieldValue == searchValue;
  });
  
  return results;
}

module.exports = {
  getWorkbookFromOneDrive,
  saveWorkbookToOneDrive,
  readSheetData,
  appendRowToSheet,
  searchInSheet
};