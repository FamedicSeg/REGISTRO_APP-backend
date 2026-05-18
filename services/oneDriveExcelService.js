// backend/services/oneDriveExcelService.js
const { getGraphAccessToken } = require('./erpAuth');
const axios = require('axios');
const ExcelJS = require('exceljs');

const ROOT_FOLDER = process.env.ONEDRIVE_ROOT_FOLDER;
const USER_EMAIL = process.env.ONEDRIVE_USER_EMAIL;

const norm = (s) =>
  String(s ?? "")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");

async function getWorkbookFromOneDrive(fileName) {
  try {
    const token = await getGraphAccessToken();
    const encodedFileName = encodeURIComponent(fileName);
    
    if (!USER_EMAIL) {
      throw new Error("Falta ONEDRIVE_USER_EMAIL en .env");
    }
    
    // Usar la API de /users/{email}/drive (funciona con permisos de aplicación)
    const url = `https://graph.microsoft.com/v1.0/users/${USER_EMAIL}/drive/root:/${ROOT_FOLDER}${encodedFileName}:/content`;
    
    console.log(`📁 Leyendo: ${fileName}`);
    
    const response = await axios.get(url, {
      headers: { 
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      },
      responseType: 'arraybuffer'
    });
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(response.data);
    console.log(`✅ Leído: ${fileName}`);
    return workbook;
  } catch (error) {
    console.error(`❌ Error leyendo ${fileName} de OneDrive:`);
    if (error.response) {
      console.error(`Status: ${error.response.status}`);
    } else {
      console.error(`Error: ${error.message}`);
    }
    throw new Error(`No se pudo leer ${fileName} desde OneDrive`);
  }
}

async function saveWorkbookToOneDrive(fileName, workbook) {
  try {
    const token = await getGraphAccessToken();
    const encodedFileName = encodeURIComponent(fileName);
    const buffer = await workbook.xlsx.writeBuffer();
    
    const url = `https://graph.microsoft.com/v1.0/users/${USER_EMAIL}/drive/root:/${ROOT_FOLDER}${encodedFileName}:/content`;
    
    await axios.put(url, buffer, {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      }
    });
    console.log(`✅ Guardado: ${fileName}`);
    return true;
  } catch (error) {
    console.error(`❌ Error guardando ${fileName} en OneDrive:`, error.message);
    throw error;
  }
}

// ============================================
// RESTO DE FUNCIONES (igual que antes)
// ============================================

async function getOPListaFromOneDrive() {
  const workbook = await getWorkbookFromOneDrive("Planificación Semanal.xlsx");
  
  const HEADER_ROWS = {
    "PLAN CONFECCION": 7,
    "PLAN AUTOMATICAS": 6
  };
  
  const hojasBuscadas = ["PLAN CONFECCION", "PLAN AUTOMATICAS"];
  const resultado = { hojas: {} };

  for (const nombreHoja of hojasBuscadas) {
    const HEADER_ROW = HEADER_ROWS[nombreHoja] || 7;
    const ws = workbook.worksheets.find((w) => norm(w.name) === norm(nombreHoja));
    
    if (!ws) {
      resultado.hojas[nombreHoja] = { error: "Hoja no encontrada", datos: [] };
      continue;
    }

    const headerRow = ws.getRow(HEADER_ROW);
    const columnasOP = [];
    
    for (let col = 1; col <= 50; col++) {
      const cellValue = String(headerRow.getCell(col).value || "").trim();
      if (norm(cellValue) === "OP" || (norm(cellValue) && norm(cellValue).includes("OP"))) {
        columnasOP.push({ numero: col, nombre: cellValue });
      }
    }

    const datosHoja = [];
    const lastRow = ws.rowCount;
    
    for (let rowNum = HEADER_ROW + 1; rowNum <= lastRow; rowNum++) {
      const row = ws.getRow(rowNum);
      const filaData = { fila: rowNum, valores: {} };
      let tieneDatos = false;
      
      columnasOP.forEach(columna => {
        let valor = row.getCell(columna.numero).value;
        if (valor && typeof valor === 'object') {
          valor = valor.result || valor.text || "";
        }
        const valorStr = valor ? String(valor).trim() : "";
        if (valorStr !== "") tieneDatos = true;
        filaData.valores[columna.numero] = { valor: valorStr };
      });
      
      if (tieneDatos) datosHoja.push(filaData);
    }
    
    resultado.hojas[nombreHoja] = { datos: datosHoja };
  }
  
  return resultado;
}

async function getModuloPersonalFromOneDrive(modulo) {
  const workbook = await getWorkbookFromOneDrive("Modulos.xlsx");
  
  const wanted = [];
  if (/^\d+$/.test(modulo)) {
    wanted.push(`MODULO ${modulo}`);
    wanted.push(`MÓDULO ${modulo}`);
  } else {
    wanted.push(modulo);
  }
  const wantedNorms = wanted.map(norm);
  
  const ws = workbook.worksheets.find((w) => wantedNorms.includes(norm(w.name)));
  if (!ws) {
    throw new Error(`No existe la hoja para '${modulo}'`);
  }
  
  const HEADER_ROW = 4;
  const headerRow = ws.getRow(HEADER_ROW);
  
  let COL_SUP = -1, COL_LID = -1, COL_INT = -1, COL_CARGO = -1;
  for (let c = 1; c <= 20; c++) {
    const h = norm(headerRow.getCell(c).text);
    if (COL_SUP === -1 && h.includes("SUPERVISOR")) COL_SUP = c;
    if (COL_LID === -1 && h.includes("LIDER")) COL_LID = c;
    if (COL_INT === -1 && h.startsWith("INTEGRANTES")) COL_INT = c;
    if (COL_CARGO === -1 && h.includes("CARGO")) COL_CARGO = c;
  }
  
  const supervisoresSet = new Set();
  const lideresSet = new Set();
  const integrantes = [];
  
  for (let r = 5; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const sup = String(row.getCell(COL_SUP).text || "").trim();
    const lid = String(row.getCell(COL_LID).text || "").trim();
    const nombre = String(row.getCell(COL_INT).text || "").trim();
    const cargo = String(row.getCell(COL_CARGO).text || "").trim();
    
    if (sup) supervisoresSet.add(sup);
    if (lid) lideresSet.add(lid);
    if (nombre) integrantes.push({ nombre, cargo });
  }
  
  return {
    hoja: ws.name,
    supervisores: [...supervisoresSet],
    lideres: [...lideresSet],
    modulo: modulo,
    integrantes
  };
}

async function getLoteInfoFromOneDrive(codigo) {
  const workbook = await getWorkbookFromOneDrive("Lote.xlsx");
  const ws = workbook.worksheets[0];
  
  const headerRow = ws.getRow(1).values.slice(1).map(v => String(v || "").trim());
  const idxCodigo = headerRow.findIndex(h => norm(h) === "CODIGO" || norm(h) === "CÓDIGO");
  const idxLoteInfo = headerRow.findIndex(h => norm(h) === "LOTE");
  
  if (idxCodigo === -1 || idxLoteInfo === -1) {
    throw new Error("No encuentro columnas CÓDIGO y/o LOTE");
  }
  
  let loteInfo = null;
  
  for (let rowNum = 2; rowNum <= ws.rowCount; rowNum++) {
    const row = ws.getRow(rowNum);
    const codCelda = row.getCell(idxCodigo + 1);
    let cod = String(codCelda.value?.result || codCelda.value?.text || codCelda.value || "");
    cod = cod.trim();
    
    if (cod === String(codigo).trim()) {
      const loteCelda = row.getCell(idxLoteInfo + 1);
      loteInfo = String(loteCelda.value?.result || loteCelda.value?.text || loteCelda.value || "").trim();
      break;
    }
  }
  
  if (!loteInfo) throw new Error("Código no encontrado");
  return { codigo: String(codigo).trim(), loteInfo };
}

async function getInsumoLoteFromOneDrive(codigo) {
  const workbook = await getWorkbookFromOneDrive("Insumos.xlsx");
  const ws = workbook.worksheets[0];
  
  const headerRow = ws.getRow(1);
  const headers = [];
  headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    headers[colNumber] = norm(String(cell.text || cell.value || "").trim());
  });
  
  let idxCodigo = -1;
  for (let i = 1; i < headers.length; i++) {
    if (headers[i]?.includes("CODIGO DE INSUMO") || headers[i]?.includes("CÓDIGO DE INSUMO")) {
      idxCodigo = i;
      break;
    }
  }
  
  let idxLote = -1;
  for (let i = 1; i < headers.length; i++) {
    if (headers[i]?.includes("LOTE")) {
      idxLote = i;
      break;
    }
  }
  
  if (idxCodigo === -1) throw new Error("No se encontró columna CÓDIGO DE INSUMO");
  
  let lote = "";
  let encontrado = false;
  
  for (let rowNum = 2; rowNum <= ws.rowCount; rowNum++) {
    const row = ws.getRow(rowNum);
    const celdaCodigo = row.getCell(idxCodigo);
    let valorInsumo = String(celdaCodigo.value?.result || celdaCodigo.value?.text || celdaCodigo.value || "").trim();
    
    if (valorInsumo === codigo) {
      encontrado = true;
      if (idxLote !== -1) {
        const celdaLote = row.getCell(idxLote);
        lote = String(celdaLote.value?.result || celdaLote.value?.text || celdaLote.value || "").trim();
      }
      break;
    }
  }
  
  if (!encontrado) throw new Error("Insumo no encontrado");
  return { codigo, lote };
}

async function getActividadCantidadPorHoraFromOneDrive(actividad) {
  const workbook = await getWorkbookFromOneDrive("Cantidades por Actividad.xlsx");
  const ws = workbook.worksheets[0];
  
  const headerRow = ws.getRow(1);
  const headers = [];
  headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    headers[colNumber] = String(cell.text || cell.value || "").trim().toUpperCase();
  });
  
  let idxActividad = -1;
  let idxCantidadHora = -1;
  
  for (let i = 1; i < headers.length; i++) {
    if (headers[i]?.includes("DESCRIPCIÓN") || headers[i]?.includes("DESCRIPCION")) idxActividad = i;
    if (headers[i]?.includes("CANTIDAD POR HORA")) idxCantidadHora = i;
  }
  
  if (idxActividad === -1 || idxCantidadHora === -1) {
    throw new Error("No se encontraron columnas requeridas");
  }
  
  let cantidadPorHora = "";
  
  for (let rowNum = 2; rowNum <= ws.rowCount; rowNum++) {
    const row = ws.getRow(rowNum);
    const celdaActividad = row.getCell(idxActividad);
    let valorActividad = String(celdaActividad.value?.result || celdaActividad.value?.text || celdaActividad.value || "").trim().toUpperCase();
    
    if (valorActividad === actividad.toUpperCase()) {
      const celdaCantidad = row.getCell(idxCantidadHora);
      cantidadPorHora = String(celdaCantidad.value?.result || celdaCantidad.value?.text || celdaCantidad.value || "").trim();
      break;
    }
  }
  
  if (!cantidadPorHora) throw new Error("Actividad no encontrada");
  return { actividad, cantidad_por_hora: cantidadPorHora };
}

async function getProcesosProductoFromOneDrive(codigo) {
  const workbook = await getWorkbookFromOneDrive("Insumos.xlsx");
  const ws = workbook.worksheets[0];
  
  const headerRow = ws.getRow(1);
  const headers = [];
  headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    headers[colNumber] = norm(String(cell.text || cell.value || "").trim());
  });
  
  const idxProducto = headers.findIndex(h => h === "CÓDIGO" || h === "CODIGO");
  const idxProceso = headers.findIndex(h => h === "PROCESO");
  
  const procesos = [];
  
  for (let rowNum = 2; rowNum <= ws.rowCount; rowNum++) {
    const row = ws.getRow(rowNum);
    const producto = String(row.getCell(idxProducto).text || row.getCell(idxProducto).value || "").trim();
    
    if (producto === codigo) {
      const proceso = String(row.getCell(idxProceso).text || row.getCell(idxProceso).value || "").trim();
      if (proceso) procesos.push(proceso);
    }
  }
  
  return { detalles: procesos.join("\n") };
}

async function getCantidadesProductoFromOneDrive(codigo) {
  const workbook = await getWorkbookFromOneDrive("Productos.xlsx");
  const ws = workbook.worksheets.find(w => norm(w.name) === "PRODUCTOS");
  
  if (!ws) throw new Error("No existe hoja PRODUCTOS");
  
  const headers = ws.getRow(1).values.slice(1).map(v => norm(v));
  const idxCodigo = headers.findIndex(h => h === "CÓDIGO" || h === "CODIGO");
  const idxMeta = headers.findIndex(h => h === "CANTIDAD POR HORA");
  
  if (idxCodigo === -1 || idxMeta === -1) throw new Error("No existe columna CANTIDAD POR HORA o CÓDIGO");
  
  let metaEncontrada = null;
  
  for (let rowNum = 2; rowNum <= ws.rowCount; rowNum++) {
    const row = ws.getRow(rowNum);
    const codigoFila = String(row.getCell(idxCodigo + 1).value || "").trim();
    
    if (norm(codigoFila) === norm(codigo)) {
      const celdaValor = row.getCell(idxMeta + 1);
      let valorExtraido = "";
      
      if (celdaValor.value && typeof celdaValor.value === 'object') {
        valorExtraido = String(celdaValor.value.result || celdaValor.value.text || "");
      } else {
        valorExtraido = String(celdaValor.value || "");
      }
      metaEncontrada = valorExtraido.trim();
      break;
    }
  }
  
  if (metaEncontrada === null) throw new Error("Código de producto no encontrado");
  return { meta: metaEncontrada };
}

async function getModuloMaquinasFromOneDrive(modulo) {
  const workbook = await getWorkbookFromOneDrive("Modulos.xlsx");
  
  const wanted = [];
  if (/^\d+$/.test(modulo)) {
    wanted.push(`MODULO ${modulo}`);
    wanted.push(`MÓDULO ${modulo}`);
  } else {
    wanted.push(modulo);
  }
  const wantedNorms = wanted.map(norm);
  
  const ws = workbook.worksheets.find((w) => wantedNorms.includes(norm(w.name)));
  if (!ws) throw new Error(`No existe la hoja para '${modulo}'`);
  
  const HEADER_ROW = 4;
  const headerRow = ws.getRow(HEADER_ROW);
  
  let COL_MAQUINARIA = -1;
  let COL_CANTIDAD = -1;
  
  for (let c = 1; c <= 50; c++) {
    const cellValue = String(headerRow.getCell(c).text || headerRow.getCell(c).value || "").trim().toUpperCase();
    if (cellValue.includes("MAQUINARIA") && !cellValue.includes("CANTIDAD")) COL_MAQUINARIA = c;
    if (cellValue.includes("CANTIDAD") || cellValue.includes("CANTIDAD DE MAQUINARIA")) COL_CANTIDAD = c;
  }
  
  const maquinasMap = new Map();
  
  for (let r = 5; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const maquina = String(row.getCell(COL_MAQUINARIA).text || row.getCell(COL_MAQUINARIA).value || "").trim();
    const cantidadStr = String(row.getCell(COL_CANTIDAD).text || row.getCell(COL_CANTIDAD).value || "").trim();
    
    if (maquina && cantidadStr) {
      const cantidad = parseInt(cantidadStr);
      if (!isNaN(cantidad) && cantidad > 0) {
        maquinasMap.set(maquina, (maquinasMap.get(maquina) || 0) + cantidad);
      }
    }
  }
  
  const resultados = [];
  maquinasMap.forEach((cantidad, nombre) => {
    resultados.push({ nombre, cantidad });
  });
  
  return { modulo, maquinas: resultados };
}

async function saveRegistroToExcel(registro) {
  const fileName = "Registro de Confección o Automáticas RG-GPR-10.xlsx";
  let workbook;
  
  try {
    workbook = await getWorkbookFromOneDrive(fileName);
  } catch (error) {
    workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet("Registros");
    const HEADERS = [
      "FECHA", "OP", "TURNO", "AREA", "MÓDULO", "RESPONSABLE", "SUPERVISOR",
      "PERSONAL ASIGNADO", "PERSONAL OTRO", "PERSONAL PRESENTE", "CÓDIGO PRODUCTO",
      "DESCRIPCIÓN", "CANTIDAD PLANIFICADA", "LOTE", "REPOSICIÓN NO CONFORME",
      "CANTIDAD ELABORADO", "CANTIDAD PROCESO", "CANTIDAD MERMA",
      "FECHA FINAL DE PRODUCTO EN PROCESO", "HORA INICIO", "HORA FIN", "DESTINO",
      "N. CLIENTE", "ESTÉRIL", "LEYENDA", "LEYENDA SI", "LEYENDA OTRA",
      "DETALLES ACTIVIDADES", "INSUMOS", "ETIQUETAS", "INTEGRANTES",
      "ACTIVIDADES_POR_INTEGRANTE", "MAQUINAS", "OBSERVACIONES GENERALES"
    ];
    ws.addRow(HEADERS);
  }
  
  const ws = workbook.worksheets[0];
  
  const insumos = typeof registro.insumos === 'string' ? JSON.parse(registro.insumos || '[]') : registro.insumos || [];
  const etiquetas = typeof registro.etiquetas === 'string' ? JSON.parse(registro.etiquetas || '[]') : registro.etiquetas || [];
  const integrantes = typeof registro.integrantes === 'string' ? JSON.parse(registro.integrantes || '[]') : registro.integrantes || [];
  const actividadesPorIntegrante = typeof registro.actividades_por_integrante === 'string' 
    ? JSON.parse(registro.actividades_por_integrante || '{}') 
    : registro.actividades_por_integrante || {};
  const maquinarias = typeof registro.maquinarias === 'string' ? JSON.parse(registro.maquinarias || '[]') : registro.maquinarias || [];
  
  ws.addRow([
    registro.fecha || "No Aplica",
    registro.op || "No Aplica",
    registro.turno || "No Aplica",
    registro.area || "No Aplica",
    registro.modulo || "No Aplica",
    registro.responsable || "No Aplica",
    registro.supervisor || "No Aplica",
    registro.personal_asignado || "No Aplica",
    registro.personal_otro || "No Aplica",
    registro.personal_presente || "No Aplica",
    registro.codigo_producto || "No Aplica",
    registro.descripcion || "No Aplica",
    registro.cantidad_planificada || 0,
    registro.loteUnido || "No Aplica",
    registro.reposicion_no_conforme || "No Aplica",
    registro.cantidad_elaborado || 0,
    registro.cantidad_proceso || 0,
    registro.cantidad_merma || 0,
    registro.fecha_final_producto || "No Aplica",
    registro.hora_inicio || "No Aplica",
    registro.hora_fin || "No Aplica",
    registro.destino || "No Aplica",
    registro.cliente || "No Aplica",
    registro.esteril || "No Aplica",
    registro.leyenda || "No Aplica",
    registro.leyenda_si || "No Aplica",
    registro.leyenda_otra || "No Aplica",
    registro.detalles_actividades || "No Aplica",
    JSON.stringify(insumos),
    JSON.stringify(etiquetas),
    JSON.stringify(integrantes),
    JSON.stringify(actividadesPorIntegrante),
    JSON.stringify(maquinarias),
    registro.observaciones || "No Aplica"
  ]);
  
  await saveWorkbookToOneDrive(fileName, workbook);
  return true;
}

module.exports = {
  getOPListaFromOneDrive,
  getModuloPersonalFromOneDrive,
  getLoteInfoFromOneDrive,
  getInsumoLoteFromOneDrive,
  getActividadCantidadPorHoraFromOneDrive,
  getProcesosProductoFromOneDrive,
  getCantidadesProductoFromOneDrive,
  getModuloMaquinasFromOneDrive,
  saveRegistroToExcel
};