const express = require('express');
const router = express.Router();
const excelService = require('../services/excelService');

const ROOT_FOLDER = process.env.ONEDRIVE_ROOT_FOLDER;
const REGISTROS_PATH = process.env.ONEDRIVE_REGISTROS_PATH;

// ============================================
// LECTURA: Obtener todos los datos de un archivo
// ============================================
router.post('/read', async (req, res) => {
  try {
    const { fileName, sheetName, hasHeaders = true } = req.body;
    
    if (!fileName) {
      return res.status(400).json({ error: 'fileName es requerido' });
    }
    
    const filePath = `${ROOT_FOLDER}${fileName}`;
    const data = await excelService.readSheetData(filePath, sheetName, hasHeaders);
    
    res.json({
      success: true,
      fileName,
      sheetName,
      data
    });
  } catch (error) {
    console.error('Error en /read:', error);
    res.status(500).json({ error: error.message });
  }
});

// ============================================
// BÚSQUEDA: Buscar en un archivo específico
// ============================================
router.post('/search', async (req, res) => {
  try {
    const { fileName, sheetName, searchField, searchValue } = req.body;
    
    if (!fileName || !sheetName || !searchField || !searchValue) {
      return res.status(400).json({ 
        error: 'Se requiere fileName, sheetName, searchField y searchValue' 
      });
    }
    
    const filePath = `${ROOT_FOLDER}${fileName}`;
    const results = await excelService.searchInSheet(filePath, sheetName, searchField, searchValue);
    
    res.json({
      success: true,
      fileName,
      sheetName,
      searchField,
      searchValue,
      results
    });
  } catch (error) {
    console.error('Error en /search:', error);
    res.status(500).json({ error: error.message });
  }
});

// ============================================
// ESCRITURA: Agregar registro aprobado al archivo principal
// ============================================
router.post('/append-registro', async (req, res) => {
  try {
    const { sheetName, registroData } = req.body;
    
    if (!sheetName || !registroData) {
      return res.status(400).json({ 
        error: 'Se requiere sheetName y registroData' 
      });
    }
    
    // Agregar timestamp automático si no viene
    if (!registroData.fechaAprobacion) {
      registroData.fechaAprobacion = new Date().toISOString();
    }
    
    const result = await excelService.appendRowToSheet(
      REGISTROS_PATH,
      sheetName,
      registroData
    );
    
    res.json({
      success: true,
      message: 'Registro guardado exitosamente en OneDrive',
      ...result
    });
  } catch (error) {
    console.error('Error en /append-registro:', error);
    res.status(500).json({ error: error.message });
  }
});

// ============================================
// BÚSQUEDA ESPECÍFICA PARA LOTES
// ============================================
router.post('/buscar-lote', async (req, res) => {
  try {
    const { loteId } = req.body;
    
    if (!loteId) {
      return res.status(400).json({ error: 'loteId es requerido' });
    }
    
    const filePath = `${ROOT_FOLDER}Lote.xlsx`;
    const results = await excelService.searchInSheet(filePath, 'Hoja1', 'ID', loteId);
    
    res.json({
      success: true,
      loteId,
      encontrado: results.length > 0,
      data: results.length > 0 ? results[0] : null
    });
  } catch (error) {
    console.error('Error en /buscar-lote:', error);
    res.status(500).json({ error: error.message });
  }
});

// ============================================
// BÚSQUEDA ESPECÍFICA PARA INSUMOS
// ============================================
router.post('/buscar-insumo', async (req, res) => {
  try {
    const { insumoId } = req.body;
    
    if (!insumoId) {
      return res.status(400).json({ error: 'insumoId es requerido' });
    }
    
    const filePath = `${ROOT_FOLDER}Insumos.xlsx`;
    const results = await excelService.searchInSheet(filePath, 'Hoja1', 'ID', insumoId);
    
    res.json({
      success: true,
      insumoId,
      encontrado: results.length > 0,
      data: results.length > 0 ? results[0] : null
    });
  } catch (error) {
    console.error('Error en /buscar-insumo:', error);
    res.status(500).json({ error: error.message });
  }
});

// ============================================
// BÚSQUEDA ESPECÍFICA PARA PRODUCTOS
// ============================================
router.post('/buscar-producto', async (req, res) => {
  try {
    const { productoId } = req.body;
    
    if (!productoId) {
      return res.status(400).json({ error: 'productoId es requerido' });
    }
    
    const filePath = `${ROOT_FOLDER}Productos.xlsx`;
    const results = await excelService.searchInSheet(filePath, 'Hoja1', 'ID', productoId);
    
    res.json({
      success: true,
      productoId,
      encontrado: results.length > 0,
      data: results.length > 0 ? results[0] : null
    });
  } catch (error) {
    console.error('Error en /buscar-producto:', error);
    res.status(500).json({ error: error.message });
  }
});

module.exports = router;