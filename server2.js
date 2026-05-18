const express = require("express");
const cors = require("cors");
const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const db = require('./database/db');
const bcrypt = require("bcrypt");
const app = express();
const {getAccessToken} = require("./services/erpAuth");
const axios = require ("axios");
require("dotenv").config();

app.use(
  cors({
    origin: [
      "http://localhost:5173",
      "http://127.0.0.1:5173",
      "http://192.168.1.46:5173",      // ✅ TU SERVIDOR PRINCIPAL
      "http://192.168.4.147:5173",
      "http://192.168.3.106:5173",
      /https:\/\/.*\.vercel\.app$/,
      /^http:\/\/192\.168\.\d{1,3}\.\d{1,3}:5173$/,  // CUALQUIER IP en la red local
      /https:\/\/.*\.loca\.lt$/,
      /https:\/\/.*\.trycloudflare\.com$/,
    ],
    methods: ["GET", "POST", "PUT", "PATCH", "DELETE", "OPTIONS"],
    allowedHeaders: ["Content-Type", "Authorization"],
    credentials: true,  // ✅ Permite cookies si usas autenticación
  })
);

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
/* =========================
   PATHS (LOCAL)
========================= */
const BASE_NETWORK_PATH = process.env.NETWORK_PATH;
const normalizePath = (p) => p.replace(/\\/g, '/');

const FILE_PATH = normalizePath(path.join(BASE_NETWORK_PATH, "Registro de Confección o Automáticas RG-GPR-10.xlsx"));
const PLAN_PATH = normalizePath(path.join(BASE_NETWORK_PATH, "Planificación Semanal.xlsx"));
const PRODUCTOS_PATH = normalizePath(path.join(BASE_NETWORK_PATH, "Productos.xlsx"));
const MODULOS_PATH = normalizePath(path.join(BASE_NETWORK_PATH, "Modulos.xlsx"));
const LOTE_PATH = normalizePath(path.join(BASE_NETWORK_PATH, "Lote.xlsx"));
const INSUMOS_PATH = normalizePath(path.join(BASE_NETWORK_PATH, "Insumos.xlsx"));
const ACTIVIDAD_PATH = normalizePath(path.join(BASE_NETWORK_PATH, "Cantidades por Actividad.xlsx"));
const SHEET_NAME = "Insumos";

console.log("Existe archivo PLAN:", fs.existsSync(PLAN_PATH));

const norm = (s) =>
  String(s ?? "")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");

const cellText = (cell) => String(cell?.text ?? cell?.value ?? "").trim();

const toNum = (v, def = 0) => {
  if (v === "" || v === null || v === undefined) return def;
  const n = Number(v);
  return Number.isFinite(n) ? n : def;
};

/* =========================
   1) CREAR EXCEL SI NO EXISTE
========================= */
const HEADERS = [
  "FECHA","OP", "TURNO", "AREA", "MÓDULO", "RESPONSABLE", "SUPERVISOR", "PERSONAL ASIGNADO", "PERSONAL OTRO", "PERSONAL PRESENTE",
  "CÓDIGO PRODUCTO", "DESCRIPCIÓN", "CANTIDAD PLANIFICADA", "LOTE",
  "REPOSICIÓN NO CONFORME", "CANTIDAD ELABORADO", "CANTIDAD PROCESO", "CANTIDAD MERMA","FECHA FINAL DE PRODUCTO EN PROCESO",
  "HORA INICIO", "HORA FIN", "DESTINO", "N. CLIENTE",
  "ESTÉRIL", "LEYENDA", "LEYENDA SI", "LEYENDA OTRA",
  "DETALLES ACTIVIDADES",
  "CANTIDAD PLANIFICADA DETALLES",
  "CANTIDAD ELABORADA DETALLES",
  "PLANIFICADA TOTAL POR DETALLE",  
  "ELABORADA TOTAL POR DETALLE",     
  "INSUMOS",
  "ETIQUETAS",
  "INTEGRANTES",
  "ACTIVIDADES_POR_INTEGRANTE", 
  "MAQUINAS",
  "OBSERVACIONES GENERALES",
  "MOTIVO DEL RECHAZO",             
  "RECHAZADO POR",                
  "FECHA DE RECHAZO",
  "VERIFICADO POR",
  "FECHA DE VERIFICACIÓN",
  "APROBADO POR",
  "FECHA DE APROBACIÓN",
                  
];

async function asegurarExcel() {
  if (!fs.existsSync(FILE_PATH)) {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet(SHEET_NAME);
    ws.addRow(HEADERS);
    await wb.xlsx.writeFile(FILE_PATH);
  }
}

// ============================================
// FUNCIÓN PARA GUARDAR EN EXCEL (APROBADOS)
// ============================================
async function guardarEnExcel(registro) {
  try {
    await asegurarExcel();
    
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(FILE_PATH);
    const ws = wb.getWorksheet(SHEET_NAME);

    const insumos = typeof registro.insumos === 'string' 
      ? JSON.parse(registro.insumos || '[]') 
      : registro.insumos || [];
    
    const etiquetas = typeof registro.etiquetas === 'string'
      ? JSON.parse(registro.etiquetas || '[]')
      : registro.etiquetas || [];
    
    const integrantes = typeof registro.integrantes === 'string'
      ? JSON.parse(registro.integrantes || '[]')
      : registro.integrantes || [];
    
    const actividadesPorIntegrante = typeof registro.actividades_por_integrante === 'string'
      ? JSON.parse(registro.actividades_por_integrante || '{}')
      : registro.actividades_por_integrante || {};
    
    const planificadaPorDetalle = typeof registro.planificada_total_por_detalle === 'string'
      ? JSON.parse(registro.planificada_total_por_detalle || '{}')
      : registro.planificada_total_por_detalle || {};
    
    const elaboradaPorDetalle = typeof registro.elaborada_total_por_detalle === 'string'
      ? JSON.parse(registro.elaborada_total_por_detalle || '{}')
      : registro.elaborada_total_por_detalle || {};
    
    const maquinarias = typeof registro.maquinarias === 'string' 
      ? JSON.parse(registro.maquinarias || '[]')
      : registro.maquinarias || [];

    // Agregar fila al Excel
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
      registro.lote || "No Aplica",
      registro.reposicion_no_conforme || "No Aplica",
      registro.cantidad_elaborada || 0,
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
      registro.cantidad_planificada || 0,
      registro.cantidad_elaborada_detalles || 0,
      JSON.stringify(planificadaPorDetalle),  
      JSON.stringify(elaboradaPorDetalle),     
      JSON.stringify(insumos),
      JSON.stringify(etiquetas),
      JSON.stringify(integrantes),
      JSON.stringify(actividadesPorIntegrante),
      JSON.stringify(maquinarias),
      registro.observaciones_generales || "No Aplica",
      registro.motivo_rechazo || "No Aplica",        
      registro.rechazado_por || "No Aplica",         
      registro.fecha_rechazo || "No Aplica" ,
      registro.verificado_por || "No Aplica",
      registro.fecha_verificacion || "No Aplica",
      registro.aprobado_por || "No Aplica",
      registro.fecha_aprobacion || "No Aplica",
               
    ]);
    
    await wb.xlsx.writeFile(FILE_PATH);
    console.log(`Registro ${registro.id} guardado en Excel (APROBADO)`);
    return true;
    
  } catch (error) {
    console.error(`Error guardando en Excel:`, error);
    throw error;
  }
}

// ============================================
// ENDPOINT PARA RECHAZAR REGISTRO (ANALISTA)
// ============================================
app.put("/api/registros/:id/rechazar", async (req, res) => {
  const { id } = req.params;
  const { motivo, usuario, rol } = req.body;

  console.log("Intentando rechazar registro:", id);
  console.log("Usuario:", usuario, "Rol:", rol);
  console.log("Motivo:", motivo);

  // ANALISTA o SUPERVISOR pueden rechazar
  if (!["ANALISTA DE PRODUCCIÓN", "SUPERVISOR"].includes(rol)) {
    return res.status(403).json({ error: "Solo el analista o supervisor pueden rechazar registros" });
  }

  if (!motivo || motivo.trim() === "") {
    return res.status(400).json({ error: "Debe proporcionar un motivo de rechazo" });
  }

  try {
    // Verificar que el registro existe
    const registro = await new Promise((resolve, reject) => {
      db.get("SELECT * FROM registros WHERE id = ?", [id], (err, row) => {
        if (err) reject(err);
        else resolve(row);
      });
    });

    if (!registro) {
      return res.status(404).json({ error: "Registro no encontrado" });
    }

    // Actualizar estado a RECHAZADO
    const sql = `
      UPDATE registros
      SET 
        estado = 'rechazado',
        motivo_rechazo = ?,
        rechazado_por = ?,
        fecha_rechazo = datetime('now', '-5 hours')
      WHERE id = ?
    `;

    await new Promise((resolve, reject) => {
      db.run(sql, [motivo, usuario, id], function(err) {
        if (err) reject(err);
        else resolve();
      });
    });

    res.json({ 
      mensaje: "Registro rechazado correctamente",
      id: id
    });

  } catch (err) {
    console.error("❌ Error al rechazar registro:", err);
    res.status(500).json({ error: "Error al rechazar registro" });
  }
});

// ============================================
// ENDPOINT PARA APROBAR (AHORA GUARDA EN EXCEL)
// ============================================
app.put("/api/registros/:id/aprobar", async (req, res) => {
  const { id } = req.params;
  const { usuario, rol } = req.body;

  console.log("Intentando aprobar registro:", id);
  console.log("Usuario:", usuario, "Rol:", rol);

  if (rol !== "ANALISTA DE PRODUCCIÓN") {
    console.log("Rol incorrecto:", rol);
    return res.status(403).json({ error: "Solo el analista puede aprobar" });
  }

  try {
    // 1. Obtener registro actual
    const registro = await new Promise((resolve, reject) => {
      db.get("SELECT * FROM registros WHERE id = ?", [id], (err, row) => {
        if (err) reject(err);
        else {
          console.log("Registro ANTES:", row);
          resolve(row);
        }
      });
    });

    if (!registro) {
      return res.status(404).json({ error: "Registro no encontrado" });
    }

    // 2. Actualizar estado
    const result = await new Promise((resolve, reject) => {
      const sql = `
        UPDATE registros
        SET 
          aprobado_por = ?,
          fecha_aprobacion = datetime('now', '-5 hours'),
          estado = 'aprobado'
        WHERE id = ?
      `;
      db.run(sql, [usuario, id], function(err) {
        if (err) reject(err);
        else {
          console.log(`Filas actualizadas: ${this.changes}`);
          resolve(this.changes);
        }
      });
    });

    if (result === 0) {
      console.log("No se actualizó ninguna fila");
    }

    // 3. Obtener registro actualizado
    const registroActualizado = await new Promise((resolve, reject) => {
      db.get("SELECT * FROM registros WHERE id = ?", [id], (err, row) => {
        if (err) reject(err);
        else {
          console.log("Registro DESPUÉS:", row);
          resolve(row);
        }
      });
    });

    // 4. Guardar en Excel
    await guardarEnExcel(registroActualizado);

    res.json({ 
      mensaje: "Registro aprobado y guardado en Excel correctamente",
      guardadoEnExcel: true,
      registro: registroActualizado
    });

  } catch (err) {
    console.error("❌ Error en aprobación:", err);
    res.status(500).json({ error: "Error al aprobar registro" });
  }
});

// ============================================
// ENDPOINT ORIGINAL POST (MODIFICADO)
// ============================================
app.post("/api/registros", async (req, res) => {
  try {
    const data = req.body;

    // =========================
    // CONVERTIR ARRAYS A JSON
    // =========================
    const insumosJSON = JSON.stringify(data.insumos || []);
    const etiquetasJSON = JSON.stringify(data.etiquetas || []);
    const integrantesJSON = JSON.stringify(data.integrantes || []);
    const maquinasJSON = JSON.stringify(data.maquinarias || []);
    const actividadesPorIntegranteJSON = data.actividades_por_integrante || "{}";
    
    const planificadaPorDetalle = data.planificada_por_detalle || {};
    const elaboradaPorDetalle = data.elaborada_por_detalle || {};
    
    const planificadaPorDetalleJSON = JSON.stringify(planificadaPorDetalle);
    const elaboradaPorDetalleJSON = JSON.stringify(elaboradaPorDetalle);

    // =========================
    // VALIDACIÓN BÁSICA
    // =========================
    if (!data.fecha || !data.op || !data.modulo) {
      return res.status(400).json({
        error: "Faltan campos obligatorios (fecha, op, modulo).",
      });
    }

    console.log("BODY COMPLETO:", data);
    console.log("Planificada por detalle:", planificadaPorDetalle);
    console.log("Elaborada por detalle:", elaboradaPorDetalle);

    // =========================
    // INSERT SQLITE (MODIFICADO)
    // =========================
    const sqlRegistro = `
    INSERT INTO registros (
      fecha, op, turno, area, modulo, responsable, supervisor,
      personal_asignado, personal_otro, personal_presente,
      codigo_producto, descripcion, hora_planificada, cantidad_planificada, lotePrincipal, loteSecundario, loteUnido,
      reposicion_no_conforme, cantidad_elaborado, cantidad_proceso,
      cantidad_merma, fecha_final_producto, hora_inicio, hora_fin, destino, cliente,
      esteril, leyenda, leyenda_si, leyenda_otra,
      detalles_actividades, cantidad_planificada_detalles,
      cantidad_elaborada_detalles,
      planificada_total_por_detalle,    
      elaborada_total_por_detalle,      
      etiquetas, insumos, integrantes, maquinarias,
      actividades_por_integrante,
      actividades_con_horas,
      observaciones,
      estado
    )
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
  `;

    const valores = [
      data.fecha,
      data.op,
      data.turno,
      data.area,
      data.modulo,
      data.responsable,
      data.supervisor,
      data.personal_asignado,
      data.personal_otro,
      data.personal_presente,
      data.codigo_producto,
      data.descripcion,
      data.hora_planificada,
      data.cantidad_planificada,
      data.lotePrincipal,
      data.loteSecundario,
      data.loteUnido,
      data.reposicion_no_conforme,
      data.cantidad_elaborado,
      data.cantidad_proceso,
      data.cantidad_merma,
      data.fecha_final_producto,
      data.hora_inicio,
      data.hora_fin,
      data.destino,
      data.n_cliente,
      data.esteril,
      data.leyenda,
      data.leyenda_si,
      data.leyenda_otra,
      typeof data.detalles_actividades === 'string' ? data.detalles_actividades : JSON.stringify(data.detalles_actividades || []),
      data.cantidad_planificada_detalles,
      data.cantidad_elaborada_detalles,
      planificadaPorDetalleJSON,   
      elaboradaPorDetalleJSON,      
      etiquetasJSON,
      insumosJSON,
      integrantesJSON,
      maquinasJSON,
      actividadesPorIntegranteJSON,
      JSON.stringify(data.actividades_con_horas || []),
      data.observaciones,
      "pendiente_SUPERVISOR"
    ];

    db.run(sqlRegistro, valores, async function (err) {
      if (err) {
        console.error("Error al guardar registro:", err);
        return res.status(500).json({ error: "Error al guardar registro" });
      }

      const registroId = this.lastID;

      // =========================
      // GUARDAR INTEGRANTES
      // =========================
      if (Array.isArray(data.integrantes)) {
        const sqlIntegrante = `
          INSERT INTO registro_integrantes (
            registro_id,
            nombre,
            cargo
          )
          VALUES (?, ?, ?)
        `;

        for (const integrante of data.integrantes) {
          db.run(sqlIntegrante, [
            registroId,
            integrante.nombre,
            integrante.cargo
          ]);
        }
      }

      res.json({
        mensaje: "Registro guardado correctamente en BD",
        id: registroId
      });
    });

  } catch (err) {
    console.error("ERROR POST /api/registros:", err);
    return res.status(500).json({ error: "Error guardando registro" });
  }
});

// ============================================
// ENDPOINT PARA ACTUALIZAR ESTADO (GENÉRICO)
// ============================================
app.put("/api/registros/:id/estado", async (req, res) => {
  const { id } = req.params;
  const { estado, usuario, rol } = req.body;

  try {
    // Validar según el rol y el estado
    let sql = "UPDATE registros SET estado = ?";
    let params = [estado];

    if (estado === 'aprobado' && rol === 'ANALISTA DE PRODUCCIÓN') {
      sql += ", aprobado_por = ?, fecha_aprobacion = datetime('now', '-5 hours')";
      params.push(usuario);
    } else if (estado === 'pendiente_ANALISTA_PRODUCCION' && rol === 'SUPERVISOR') {
      sql += ", verificado_por = ?, fecha_verificacion = datetime('now', '-5 hours')";
      params.push(usuario);
    }

    sql += " WHERE id = ?";
    params.push(id);

    await new Promise((resolve, reject) => {
      db.run(sql, params, function(err) {
        if (err) reject(err);
        else resolve();
      });
    });

    // Si el nuevo estado es "aprobado", guardar en Excel
    if (estado === 'aprobado') {
      const registro = await new Promise((resolve, reject) => {
        db.get("SELECT * FROM registros WHERE id = ?", [id], (err, row) => {
          if (err) reject(err);
          else resolve(row);
        });
      });
      
      await guardarEnExcel(registro);
    }

    res.json({ 
      mensaje: estado === 'aprobado' 
        ? "Registro aprobado y guardado en Excel" 
        : "Estado actualizado correctamente" 
    });

  } catch (err) {
    console.error("Error actualizando estado:", err);
    res.status(500).json({ error: "Error al actualizar estado" });
  }
});

// =========================================
// ENDPOINT DE PRUEBA 
// =======================================

app.get("/api/productos/test", async (req, res) => {
  try {
    const token = await getAccessToken();

    const companiesUrl = `https://api.businesscentral.dynamics.com/v2.0/${process.env.TENANT_ID}/${process.env.BC_ENVIRONMENT}/api/v2.0/companies`;

    const companiesRes = await axios.get(companiesUrl, {
      headers: {
        Authorization: `Bearer ${token}`,
      },
    });

    if (!companiesRes.data?.value?.length) {
      return res.status(500).json({ error: "No se encontraron compañías" });
    }

    const companyId = companiesRes.data.value[0].id;

    const itemsUrl = `https://api.businesscentral.dynamics.com/v2.0/${process.env.TENANT_ID}/${process.env.BC_ENVIRONMENT}/api/v2.0/companies(${companyId})/items?$top=20`;

    const response = await axios.get(itemsUrl, {
      headers: {
        Authorization: `Bearer ${token}`,
      },
    });

    const items = (response.data.value || []).map(item => ({
      codigo: item.number,
      descripcion: item.displayName,
    }));

    res.json(items);
  } catch (err) {
    console.error(err.response?.data || err.message);
    res.status(500).json({
      error: "Error al consultar productos",
      detalle: err.response?.data || err.message,
    });
  }
});

app.get("/api/companies/test", async (req, res) => {
  try {
    const token = await getAccessToken();

    const companiesUrl = `https://api.businesscentral.dynamics.com/v2.0/${process.env.TENANT_ID}/${process.env.BC_ENVIRONMENT}/api/v2.0/companies`;

    const companiesRes = await axios.get(companiesUrl, {
      headers: {
        Authorization: `Bearer ${token}`,
      },
    });

    return res.json(
      (companiesRes.data.value || []).map(c => ({
        id: c.id,
        name: c.name,
      }))
    );
  } catch (err) {
    console.error("Error consultando compañías:", err.response?.data || err.message);
    return res.status(500).json({
      error: "Error al consultar compañías",
      detalle: err.response?.data || err.message,
    });
  }
});

app.get("/api/productos/detalle", async (req, res) => {
  try {
    const { codigo } = req.query;

    console.log("Código recibido:", codigo);

    if (!codigo) {
      return res.status(400).json({ error: "Falta el código" });
    }

    const codigoLimpio = String(codigo).trim();
    const nombreCompaniaBuscada = "FAMEDIC";

    console.log("Código limpio:", codigoLimpio);

    const token = await getAccessToken();
    console.log("Token obtenido correctamente");

    const baseUrl = `https://api.businesscentral.dynamics.com/v2.0/${process.env.TENANT_ID}/${process.env.BC_ENVIRONMENT}/api/v2.0`;


    const companiesUrl = `${baseUrl}/companies`;
    console.log("Consultando compañías:", companiesUrl);

    const companiesRes = await axios.get(companiesUrl, {
      headers: {
        Authorization: `Bearer ${token}`,
      },
    });

    console.log("Compañías obtenidas:", companiesRes.data?.value?.length || 0);

    const company = (companiesRes.data.value || []).find(
      (c) => String(c.name).trim().toUpperCase() === nombreCompaniaBuscada.toUpperCase()
    );

    if (!company) {
      console.log("No se encontró la compañía:", nombreCompaniaBuscada);
      return res.status(404).json({
        error: `No se encontró la compañía '${nombreCompaniaBuscada}'`,
      });
    }

    const companyId = company.id;

    console.log("COMPANY ID:", companyId);
    console.log("COMPANIA:", company.name);

    const itemsUrl = `${baseUrl}/companies(${companyId})/items`;
    const filtro = `number eq '${codigoLimpio.replace(/'/g, "''")}'`;

    console.log("Consultando items:", itemsUrl);
    console.log("Filtro:", filtro);

    const response = await axios.get(itemsUrl, {
      headers: {
        Authorization: `Bearer ${token}`,
      },
      params: {
        $filter: filtro,
      },
    });

    console.log("RESPUESTA ITEMS:", JSON.stringify(response.data, null, 2));

    const item = response.data?.value?.[0];

    if (!item) {
      console.log("Producto no encontrado:", codigoLimpio);
      return res.status(404).json({
        error: "Producto no encontrado en ERP",
        codigo_buscado: codigoLimpio,
        compania: company.name,
      });
    }

    console.log("Producto encontrado:", item.number, item.displayName);

    return res.json({
      codigo: item.number || "",
      descripcion: item.displayName || "",
      compania: company.name,
    });
  } catch (err) {
    console.error("=== ERROR EN /api/productos/detalle ===");
    console.error("Mensaje:", err.message);
    console.error("Status:", err.response?.status);
    console.error("Data:", JSON.stringify(err.response?.data, null, 2));

    return res.status(500).json({
      error: "Error al consultar el producto en ERP",
      detalle: err.response?.data || err.message,
    });
  }
});

// ============================================
// RUTAS PARA CRUD DE USUARIOS CON PRIMER LOGIN
// ============================================

// Obtener todos los usuarios
app.get("/api/usuarios", async (req, res) => {
  try {
    const usuarios = await new Promise((resolve, reject) => {
      db.all(`
        SELECT id, username, nombre, cedula_identidad, rol, area, activo, primer_login,
               created_at, updated_at 
        FROM usuarios 
        ORDER BY id DESC
      `, [], (err, rows) => {
        if (err) reject(err);
        else resolve(rows);
      });
    });
    res.json(usuarios);
  } catch (error) {
    console.error("Error al obtener usuarios:", error);
    res.status(500).json({ error: error.message });
  }
});

// Obtener un usuario por ID
app.get("/api/usuarios/:id", async (req, res) => {
  const { id } = req.params;
  try {
    const usuario = await new Promise((resolve, reject) => {
      db.get(`
        SELECT id, username, nombre, cedula_identidad, rol, area, activo, primer_login 
        FROM usuarios WHERE id = ?
      `, [id], (err, row) => {
        if (err) reject(err);
        else resolve(row);
      });
    });
    
    if (!usuario) {
      return res.status(404).json({ error: 'Usuario no encontrado' });
    }
    
    res.json(usuario);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// ============================================
// ENDPOINT PARA LOGIN CON PRIMER LOGIN
// ============================================
app.post("/api/login", async(req, res)=>{
  const {username, password} = req.body;
  
  db.get(
    "SELECT * FROM usuarios WHERE username = ?",
    [username],
    async (err, user) => {
      if (err) {
        return res.status(500).json({ error: "Error del servidor" });
      }

      if (!user) {
        return res.status(401).json({ error: "Usuario no encontrado" });
      }

      const match = await bcrypt.compare(password, user.password);

      if (!match) {
        return res.status(401).json({ error: "Contraseña incorrecta" });
      }

      // Verificar si es primer login (1 = primera vez)
      const esPrimerLogin = user.primer_login === 1;

      res.json({
        mensaje: "Login exitoso",
        user: {
          id: user.id,
          username: user.username,
          nombre: user.nombre,
          rol: user.rol,
          primerLogin: esPrimerLogin
        }
      });
    }
  );
});

// ============================================
// ENDPOINT PARA CAMBIAR CONTRASEÑA EN PRIMER LOGIN
// ============================================
app.put("/api/usuarios/:id/cambiar-primer-login", async (req, res) => {
  const { id } = req.params;
  const { nuevaPassword } = req.body;

  if (!nuevaPassword || nuevaPassword.length < 4) {
    return res.status(400).json({ error: "La contraseña debe tener al menos 4 caracteres" });
  }

  try {
    // Hashear la nueva contraseña
    const hashedPassword = await bcrypt.hash(nuevaPassword, 10);

    // Actualizar contraseña y marcar que ya no es primer login
    await new Promise((resolve, reject) => {
      db.run(
        'UPDATE usuarios SET password = ?, primer_login = 0 WHERE id = ?',
        [hashedPassword, id],
        function(err) {
          if (err) reject(err);
          else resolve();
        }
      );
    });

    res.json({ 
      mensaje: 'Contraseña cambiada correctamente',
      primerLogin: false
    });

  } catch (error) {
    console.error('Error al cambiar contraseña:', error);
    res.status(500).json({ error: error.message });
  }
});

// ============================================
// ENDPOINT PARA QUE JEFE DE PRODUCCIÓN VEA CONTRASEÑAS
// ============================================
app.get("/api/usuarios/:id/con-password", async (req, res) => {
  const { id } = req.params;
  const { rolSolicitante } = req.query;

  // Solo JEFE DE PRODUCCIÓN puede ver contraseñas
  if (rolSolicitante !== "ADMINISTRADOR") {
    return res.status(403).json({ error: "No tienes permisos para ver contraseñas" });
  }

  try {
    const usuario = await new Promise((resolve, reject) => {
      db.get(`
        SELECT id, username, nombre, cedula_identidad, rol, area, activo, primer_login, password 
        FROM usuarios WHERE id = ?
      `, [id], (err, row) => {
        if (err) reject(err);
        else resolve(row);
      });
    });
    
    if (!usuario) {
      return res.status(404).json({ error: 'Usuario no encontrado' });
    }
    
    res.json(usuario);
  } catch (error) {
    console.error("Error al obtener usuario con contraseña:", error);
    res.status(500).json({ error: error.message });
  }
});

// Crear nuevo usuario (con primer_login = 1)
app.post("/api/usuarios", async (req, res) => {
  const { username, password, nombre, cedula_identidad, rol, area, activo } = req.body;
  
  // Validaciones básicas
  if (!username || !password || !nombre || !cedula_identidad || !rol) {
    return res.status(400).json({ 
      error: 'Faltan campos requeridos: username, password, nombre, cedula_identidad, rol' 
    });
  }

  try {
    // Verificar si el username o cedula de identidad ya existen
    const existe = await new Promise((resolve, reject) => {
      db.get(
        'SELECT id FROM usuarios WHERE username = ? OR cedula_identidad = ?',
        [username, cedula_identidad],
        (err, row) => {
          if (err) reject(err);
          else resolve(row);
        }
      );
    });
    
    if (existe) {
      return res.status(400).json({ 
        error: 'El username o cedula_identidad ya está registrado' 
      });
    }

    // Hashear la contraseña
    const hashedPassword = await bcrypt.hash(password, 10);
    
    const result = await new Promise((resolve, reject) => {
      db.run(
        `INSERT INTO usuarios (username, password, nombre, cedula_identidad, rol, area, activo, primer_login) 
         VALUES (?, ?, ?, ?, ?, ?, ?, 1)`, // primer_login = 1 por defecto
        [username, hashedPassword, nombre, cedula_identidad, rol, area || null, activo !== undefined ? activo : 1],
        function(err) {
          if (err) reject(err);
          else resolve({ lastID: this.lastID });
        }
      );
    });

    const nuevoUsuario = await new Promise((resolve, reject) => {
      db.get(
        'SELECT id, username, nombre, cedula_identidad, rol, area, activo, primer_login FROM usuarios WHERE id = ?',
        [result.lastID],
        (err, row) => {
          if (err) reject(err);
          else resolve(row);
        }
      );
    });

    res.status(201).json(nuevoUsuario);
  } catch (error) {
    console.error('Error al crear usuario:', error);
    res.status(500).json({ error: error.message });
  }
});

// Endpoint para que JEFE DE PRODUCCIÓN cambie cualquier contraseña
app.put("/api/usuarios/:id/cambiar-password", async (req, res) => {
  const { id } = req.params;
  const { nuevaPassword, rolSolicitante } = req.body;

  // Solo JEFE DE PRODUCCIÓN puede cambiar contraseñas
  if (rolSolicitante !== "ADMINISTRADOR") {
    return res.status(403).json({ error: "No tienes permisos para cambiar contraseñas" });
  }

  if (!nuevaPassword || nuevaPassword.length < 4) {
    return res.status(400).json({ error: "La contraseña debe tener al menos 4 caracteres" });
  }

  try {
    // Verificar que el usuario existe
    const usuario = await new Promise((resolve, reject) => {
      db.get('SELECT id FROM usuarios WHERE id = ?', [id], (err, row) => {
        if (err) reject(err);
        else resolve(row);
      });
    });

    if (!usuario) {
      return res.status(404).json({ error: 'Usuario no encontrado' });
    }

    // Hashear la nueva contraseña
    const hashedPassword = await bcrypt.hash(nuevaPassword, 10);

    // Actualizar la contraseña y mantener primer_login (no se modifica)
    await new Promise((resolve, reject) => {
      db.run(
        'UPDATE usuarios SET password = ? WHERE id = ?',
        [hashedPassword, id],
        function(err) {
          if (err) reject(err);
          else resolve();
        }
      );
    });

    res.json({ 
      mensaje: 'Contraseña actualizada correctamente',
      id: parseInt(id)
    });

  } catch (error) {
    console.error('Error al cambiar contraseña:', error);
    res.status(500).json({ error: error.message });
  }
});

// Activar/Desactivar usuario
app.patch("/api/usuarios/:id/toggle-activo", async (req, res) => {
  const { id } = req.params;

  try {
    const usuario = await new Promise((resolve, reject) => {
      db.get('SELECT activo FROM usuarios WHERE id = ?', [id], (err, row) => {
        if (err) reject(err);
        else resolve(row);
      });
    });
    
    if (!usuario) {
      return res.status(404).json({ error: 'Usuario no encontrado' });
    }

    const nuevoEstado = usuario.activo ? 0 : 1;
    
    await new Promise((resolve, reject) => {
      db.run(
        'UPDATE usuarios SET activo = ? WHERE id = ?',
        [nuevoEstado, id],
        function(err) {
          if (err) reject(err);
          else resolve();
        }
      );
    });

    res.json({ 
      id: parseInt(id), 
      activo: nuevoEstado === 1,
      message: `Usuario ${nuevoEstado ? 'activado' : 'desactivado'} correctamente` 
    });
  } catch (error) {
    console.error('Error al cambiar estado:', error);
    res.status(500).json({ error: error.message });
  }
});

// Eliminar usuario
app.delete("/api/usuarios/:id", async (req, res) => {
  const { id } = req.params;

  try {
    const usuario = await new Promise((resolve, reject) => {
      db.get('SELECT id FROM usuarios WHERE id = ?', [id], (err, row) => {
        if (err) reject(err);
        else resolve(row);
      });
    });
    
    if (!usuario) {
      return res.status(404).json({ error: 'Usuario no encontrado' });
    }

    await new Promise((resolve, reject) => {
      db.run('DELETE FROM usuarios WHERE id = ?', [id], function(err) {
        if (err) reject(err);
        else resolve();
      });
    });

    res.json({ 
      message: 'Usuario eliminado correctamente',
      id: parseInt(id)
    });
  } catch (error) {
    console.error('Error al eliminar usuario:', error);
    res.status(500).json({ error: error.message });
  }
});

// ============================================
// RESTO DE TUS ENDPOINTS (SIN CAMBIOS)
// ============================================

app.get("/api/registros/excel", async (req, res) => {
  await asegurarExcel();
  res.download(FILE_PATH);
});

app.get("/api/planificacion", async (req, res) => {
  try {
    if (!fs.existsSync(PLAN_PATH)) {
      return res
        .status(404)
        .json({ error: "No existe planificacion.xlsx en backend" });
    }

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(PLAN_PATH);

    const ws = wb.worksheets[0];
    if (!ws) return res.status(404).json({ error: "Planificación sin hojas" });

    const rows = [];
    ws.eachRow({ includeEmpty: false }, (row) => {
      rows.push(row.values.slice(1));
    });

    const headers = rows[0] || [];
    const data = rows.slice(1).map((r) => {
      const obj = {};
      headers.forEach((h, i) => (obj[String(h)] = r[i] ?? ""));
      return obj;
    });

    res.json({ headers, data });
  } catch (err) {
    console.error("ERROR GET /api/planificacion:", err);
    res.status(500).json({ error: "Error leyendo planificación" });
  }
});



app.get("/api/productos/detalle", async (req, res) => {
  try {
    const { codigo } = req.query;
    if (!codigo) return res.status(400).json({ error: "Falta el código" });

    console.log("Buscando producto:", codigo);

    // 1. Token
    const token = await getAccessToken();

    // 2. Obtener companies (AQUÍ ESTABA TU ERROR)
    const companiesUrl = `https://api.businesscentral.dynamics.com/v2.0/${process.env.TENANT_ID}/${process.env.BC_ENVIRONMENT}/api/v2.0/companies`;

    const companiesRes = await axios.get(companiesUrl, {
      headers: {
        Authorization: `Bearer ${token}`,
      },
    });

    const companyId = companiesRes.data.value[0]?.id;

    if (!companyId) {
      return res.status(500).json({ error: "No se encontró companyId" });
    }

    console.log("🏢 Company ID:", companyId);

    // 3. Obtener items
    const itemsUrl = `https://api.businesscentral.dynamics.com/v2.0/${process.env.TENANT_ID}/${process.env.BC_ENVIRONMENT}/api/v2.0/companies(${companyId})/items?$filter=contains(number,'${codigo}')`;

    const response = await axios.get(itemsUrl, {
      headers: {
        Authorization: `Bearer ${token}`,
      },
    });

    const item = response.data.value[0];

    if (!item) {
      return res.status(404).json({ error: "Producto no encontrado en ERP" });
    }

    console.log("Producto encontrado:", item);

    res.json({
      codigo: item.number,
      detalle: item.displayName,
    });

  } catch (err) {
    console.error("❌ ERROR COMPLETO:");
    console.error(err.response?.data || err.message);

    res.status(500).json({
      error: "Error al consultar el producto en ERP",
      detalle: err.response?.data || err.message,
    });
  }
});

app.get("/api/op/lista", async (req, res) => {
  try {
    if (!fs.existsSync(PLAN_PATH)) {
      return res.status(404).json({ error: "No existe planificacion.xlsx" });
    }

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(PLAN_PATH);

    // DEBUGUEO: Mostrar todas las hojas disponibles
    console.log("====== HOJAS DISPONIBLES EN EXCEL ======");
    wb.worksheets.forEach((ws, idx) => {
      console.log(`[${idx}] "${ws.name}"`);
    });
    console.log("==========================================");

    const HEADER_ROWS = {
      "PLAN CONFECCION": 7,
      "PLAN AUTOMATICAS": 6
    };
    
    const hojasBuscadas = ["PLAN CONFECCION", "PLAN AUTOMATICAS"];
    
    const resultado = {
      hojas: {}
    };

    for (const nombreHoja of hojasBuscadas) {
      const HEADER_ROW = HEADER_ROWS[nombreHoja] || 7; // Usar la fila específica o default 7
      const ws = wb.worksheets.find((w) => norm(w.name) === norm(nombreHoja));
      
      console.log(`\n🔍 Buscando hoja: "${nombreHoja}" (normalizado: "${norm(nombreHoja)}")`);
      console.log(`   Buscando encabezados en fila: ${HEADER_ROW}`);
      
      if (!ws) {
        console.log(`❌ Hoja NO encontrada: ${nombreHoja}`);
        resultado.hojas[nombreHoja] = { error: "Hoja no encontrada", datos: [] };
        continue;
      }
      
      console.log(`✅ Hoja ENCONTRADA: "${ws.name}"`);
      console.log(`   Total de filas: ${ws.rowCount}`);

      // Obtener la fila de encabezados
      const headerRow = ws.getRow(HEADER_ROW);
      const totalCells = headerRow.cellCount || 100;
      
      console.log(`   Leyendo encabezados de fila ${HEADER_ROW}...`);

      // Encontrar TODAS las columnas que contengan "OP"
      const columnasOP = [];
      const todasLasColumnasLog = [];

      for (let col = 1; col <= totalCells; col++) {
        const cellValue = String(headerRow.getCell(col).value || "").trim();
        const cellValueNorm = norm(cellValue);
        
        if (cellValue) {
          todasLasColumnasLog.push(`Col ${col}: "${cellValue}" (norm: "${cellValueNorm}")`);
        }
        
        if (cellValueNorm === "OP" || (cellValueNorm && cellValueNorm.includes("OP"))) {
          const colLetter = getColumnLetter(col);
          columnasOP.push({
            numero: col,
            letra: colLetter,
            nombre: cellValue
          });
          console.log(`   ✅ ENCONTRADA COLUMNA OP: ${cellValue} (col ${colLetter})`);
        }
      }
      
      console.log(`   Todas las columnas encontradas:`);
      todasLasColumnasLog.forEach(log => console.log(`     ${log}`));

      if (columnasOP.length === 0) {
        console.log(`❌ No hay columnas OP en hoja ${nombreHoja}`);
        resultado.hojas[nombreHoja] = { error: "No se encontraron columnas OP", datos: [], columnasEncontradas: todasLasColumnasLog };
        continue;
      }

      console.log(`Hoja ${nombreHoja}: Columnas OP encontradas:`, columnasOP.map(c => `${c.nombre} (col ${c.letra})`).join(", "));

      // Recoger TODOS los datos de todas las filas (sin saltar ninguna)
      const datosHoja = [];
      
      // Obtener la última fila con datos de la hoja
      const lastRow = ws.rowCount;
      
      // Recorrer TODAS las filas desde HEADER_ROW+1 hasta el final
      for (let rowNum = HEADER_ROW + 1; rowNum <= lastRow; rowNum++) {
        const row = ws.getRow(rowNum);
        const filaData = {
          fila: rowNum,
          valores: {}
        };
        
        let tieneDatos = false;
        
        // Para cada columna OP, extraer su valor
        columnasOP.forEach(columna => {
          const cell = row.getCell(columna.numero);
          let valor = cell.value;
          
          // Manejar diferentes tipos de valores
          if (valor && typeof valor === 'object') {
            if (valor.result) valor = valor.result;
            else if (valor.text) valor = valor.text;
            else valor = JSON.stringify(valor);
          }
          
          const valorStr = valor ? String(valor).trim() : "";
          
          if (valorStr !== "") {
            tieneDatos = true;
          }
          
          filaData.valores[columna.letra] = {
            columna: columna.letra,
            nombre_encabezado: columna.nombre,
            valor: valorStr
          };
        });
        
        // Solo agregar si al menos una columna OP tiene datos
        if (tieneDatos) {
          datosHoja.push(filaData);
        }
      }

      // También puedes agrupar por rangos si lo necesitas
      const rangos = agruparPorRangos(datosHoja.map(d => d.fila));
      
      resultado.hojas[nombreHoja] = {
        columnas_op: columnasOP.map(c => `${c.nombre} (col ${c.letra})`),
        rangos_encontrados: rangos,
        total_filas_con_datos: datosHoja.length,
        datos: datosHoja
      };
      
      console.log(`${nombreHoja}: ${datosHoja.length} filas con datos en columnas OP`);
      console.log(`Rangos encontrados: ${rangos.join(", ")}`);
    }

    const totalDatos = Object.values(resultado.hojas).reduce((sum, hoja) => sum + (hoja.datos?.length || 0), 0);
    
    if (totalDatos === 0) {
      return res.status(400).json({
        error: "No se encontraron datos en las columnas OP",
        detalle: resultado
      });
    }

    console.log(`Total de filas con datos: ${totalDatos}`);
    res.json(resultado);
    
  } catch (err) {
    console.error("ERROR GET /api/op/lista:", err);
    res.status(500).json({ error: "Error leyendo OPs", detalle: err.message });
  }
});

// Función para convertir número de columna a letra (1=A, 2=B, 26=Z, 27=AA, etc.)
function getColumnLetter(colNum) {
  let letter = '';
  while (colNum > 0) {
    const temp = (colNum - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    colNum = (colNum - temp - 1) / 26;
  }
  return letter;
}

// Función para agrupar filas en rangos (ej: "8-10", "15-20", "25-30")
function agruparPorRangos(filas) {
  if (filas.length === 0) return [];
  
  filas.sort((a, b) => a - b);
  const rangos = [];
  let inicio = filas[0];
  let fin = filas[0];
  
  for (let i = 1; i < filas.length; i++) {
    if (filas[i] === fin + 1) {
      fin = filas[i];
    } else {
      rangos.push(inicio === fin ? `${inicio}` : `${inicio}-${fin}`);
      inicio = filas[i];
      fin = filas[i];
    }
  }
  rangos.push(inicio === fin ? `${inicio}` : `${inicio}-${fin}`);
  
  return rangos;
}

app.get("/api/modulos/personal", async (req, res) => {
  try {
    const { modulo } = req.query;
    if (!modulo) return res.status(400).json({ error: "Falta modulo" });

    if (!fs.existsSync(MODULOS_PATH)) {
      return res.status(404).json({ error: "No existe modulos.xlsx en backend" });
    }

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(MODULOS_PATH);

    const norm = (s) =>
      String(s ?? "")
        .trim()
        .toUpperCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/\s+/g, " ");

    const moduloParam = String(modulo).trim();

    const wanted = [];
    if (/^\d+$/.test(moduloParam)) {
      wanted.push(`MODULO ${moduloParam}`);
      wanted.push(`MÓDULO ${moduloParam}`);
    } else {
      wanted.push(moduloParam);
    }
    const wantedNorms = wanted.map(norm);

    const ws = wb.worksheets.find((w) => wantedNorms.includes(norm(w.name)));

    if (!ws) {
      return res.status(404).json({
        error: `No existe la hoja para '${moduloParam}'`,
        hojasDisponibles: wb.worksheets.map((w) => w.name),
      });
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

    if (COL_SUP === -1 || COL_LID === -1 || COL_INT === -1 || COL_CARGO === -1) {
      return res.status(400).json({
        error: "No se detectaron columnas SUPERVISOR, LIDER, INTEGRANTE o CARGO",
        hoja: ws.name,
        filaEncabezado: HEADER_ROW,
      });
    }

    const START_ROW = 5;
    const supervisoresSet = new Set();
    const lideresSet = new Set();
    const integrantes = [];

    for (let r = START_ROW; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);

      const sup = String(row.getCell(COL_SUP).text ?? "").trim();
      const lid = String(row.getCell(COL_LID).text ?? "").trim();
      const nombre = String(row.getCell(COL_INT).text ?? "").trim();
      const cargo = String(row.getCell(COL_CARGO).text ?? "").trim();

      if (sup) supervisoresSet.add(sup);
      if (lid) lideresSet.add(lid);
      if (nombre) {
        integrantes.push({
          nombre: nombre,
          cargo: cargo || "",
        });
      }
    }

    console.log("Integrantes cargados exitosamente");

    return res.json({
      hoja: ws.name,
      supervisores: [...supervisoresSet],
      lideres: [...lideresSet],
      modulo: moduloParam,
      integrantes,
    });

  } catch (err) {
    console.error("ERROR /api/modulos/personal:", err);
    return res.status(500).json({
      error: "Error leyendo Excel de módulos",
      message: err.message,
    });
  }
});

app.get("/api/lote/info", async (req, res) => {
  try {
    const { codigo } = req.query;
    if (!codigo) return res.status(400).json({ error: "Falta codigo" });
    
    if (!fs.existsSync(LOTE_PATH)) {
      return res.status(404).json({ error: "No existe Lote.xlsx en backend" });
    }

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(LOTE_PATH, {
      calcChain: true 
    });
    
    const ws = wb.worksheets[0];
    if (!ws) return res.status(404).json({ error: "Lote.xlsx sin hojas" });

    const headerRow = ws.getRow(1).values
      .slice(1)
      .map((v) => String(v ?? "").trim()); 

    const idxCodigo = headerRow.findIndex((h) => norm(h) === "CODIGO" || norm(h) === "CÓDIGO");
    const idxLoteInfo = headerRow.findIndex((h) => norm(h) === "LOTE");

    if (idxCodigo === -1 || idxLoteInfo === -1) {
      return res.status(400).json({
        error: "No encuentro columnas CÓDIGO y/o LOTE en Lote.xlsx",
        headers_detectados: headerRow,
      });
    }

    let loteInfo = null;
    
    // Función para obtener valor CORRECTO de celda (con fórmulas)
    const getCellValue = (cell) => {
      if (!cell) return "";
      
      // Si es objeto de fórmula (tu caso)
      if (cell.value && typeof cell.value === 'object') {
        // Caso 1: Tiene propiedad 'result' (fórmula calculada)
        if (cell.value.result !== undefined) {
          return String(cell.value.result).trim();
        }
        // Caso 2: Tiene propiedad 'text'
        if (cell.value.text !== undefined) {
          return String(cell.value.text).trim();
        }
        // Caso 3: Es objeto compartido de fórmula
        if (cell.value.sharedFormula) {
          // Buscar la celda original de la fórmula
          const formulaRef = cell.value.sharedFormula; // "C3"
          const formulaCell = ws.getCell(formulaRef);
          if (formulaCell && formulaCell.value && formulaCell.value.result) {
            return String(formulaCell.value.result).trim();
          }
        }
      }
      
      // Si no, usar el texto de la celda (ya calculado por ExcelJS)
      return String(cell.text || "").trim();
    };

    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber === 1) return;
      
      const codCell = row.getCell(idxCodigo + 1);
      const cod = getCellValue(codCell);
      
      if (cod === String(codigo).trim()) {
        const loteCell = row.getCell(idxLoteInfo + 1);
        loteInfo = getCellValue(loteCell);
        console.log("Celda encontrada:", {
          valorCrudo: loteCell.value,
          texto: loteCell.text,
          formula: loteCell.formula,
          tipo: typeof loteCell.value
        });
      }
    });

    if (!loteInfo) return res.status(404).json({ error: "Código no encontrado" });
    
    res.json({ 
      codigo: String(codigo).trim(), 
      loteInfo: loteInfo 
    });
    
  } catch (err) {
    console.error("ERROR GET /api/lote/info:", err);
    res.status(500).json({ error: "Error leyendo Lote.xlsx" });
  }
});
app.get("/api/test/bom-header", async (req, res) => {
  try {
    const token = await getAccessToken();

    const url = `https://api.businesscentral.dynamics.com/v2.0/${process.env.TENANT_ID}/${process.env.BC_ENVIRONMENT}/ODataV4/Company('FAMEDIC')/ProductionBOMList`;

    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
      },
    });

    res.json(response.data);
  } catch (err) {
    console.error("ERROR BOM HEADER:", err.response?.data || err.message);
    res.status(500).json({
      error: "Error consultando ProductionBOMList",
      detalle: err.response?.data || err.message,
    });
  }
});
app.get("/api/test/bom-header", async (req, res) => {
  try {
    const token = await getAccessToken();

    const url = `https://api.businesscentral.dynamics.com/v2.0/${process.env.TENANT_ID}/${process.env.BC_ENVIRONMENT}/ODataV4/Company('FAMEDIC')/ProductionBOMList`;

    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
      },
    });

    res.json(response.data);
  } catch (err) {
    console.error("ERROR BOM HEADER:", err.response?.data || err.message);
    res.status(500).json({
      error: "Error consultando ProductionBOMList",
      detalle: err.response?.data || err.message,
    });
  }
});

app.get("/api/insumos/producto", async (req, res) => {
  try {
    const { codigo } = req.query;

    if (!codigo) {
      return res.status(400).json({
        success: false,
        error: "Falta codigo",
      });
    }

    const codigoLimpio = String(codigo).trim();
    const nombreCompania = "FAMEDIC";

    console.log("Buscando insumos ERP para producto:", codigoLimpio);

    // 1) Obtener token
    const token = await getAccessToken();

    // 2) Base OData
    const odataBase = `https://api.businesscentral.dynamics.com/v2.0/${process.env.TENANT_ID}/${process.env.BC_ENVIRONMENT}/ODataV4/Company('${encodeURIComponent(nombreCompania)}')`;

    // 3) Buscar cabecera BOM
    const bomHeaderUrl = `${odataBase}/ProductionBOMList?$filter=No eq '${codigoLimpio.replace(/'/g, "''")}'&$format=json`;

    const bomHeaderRes = await axios.get(bomHeaderUrl, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
      },
    });

    const bomHeader = bomHeaderRes.data?.value?.[0];

    if (!bomHeader) {
      return res.status(404).json({
        success: false,
        error: "No se encontró la cabecera BOM para ese producto",
        codigo_buscado: codigoLimpio,
      });
    }

    const bomLinesUrl = `${odataBase}/ProductionBOMLines?$filter=Production_BOM_No eq '${codigoLimpio.replace(/'/g, "''")}'&$format=json`;

    const bomLinesRes = await axios.get(bomLinesUrl, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
      },
    });

    const lineas = bomLinesRes.data?.value || [];

    if (!lineas.length) {
      return res.status(404).json({
        success: false,
        error: "La BOM existe pero no tiene líneas",
        codigo_buscado: codigoLimpio,
      });
    }

    // 5) Transformar para frontend
    const insumos = lineas.map((linea) => ({
      tipo: linea.Type || "",
      codigo: linea.No || "",
      descripcion: linea.Description || "",
      unidad_medida: linea.Unit_of_Measure_Code || "",
      cantidad: Number(linea.Quantity_per || 0),
      rechazo: Number(linea.Scrap_Percent || 0),
      linea_no: Number(linea.Line_No || 0),
      version_codigo: linea.Version_Code || "",
      producto_bom: linea.Production_BOM_No || "",
      posicion: linea.Position || "",
    }));

    return res.json({
      success: true,
      producto: {
        codigo: bomHeader.No || codigoLimpio,
        descripcion: bomHeader.Description || "",
        unidad_medida: bomHeader.Unit_of_Measure_Code || "",
        estado: bomHeader.Status || "",
        version_activa: bomHeader.Version_Code || bomHeader.ActiveVersionCode || "",
      },
      insumos,
    });
  } catch (err) {
    console.error("❌ ERROR obteniendo insumos desde ERP:");
    console.error("STATUS:", err.response?.status);
    console.error("DATA:", JSON.stringify(err.response?.data, null, 2) || err.message);

    return res.status(500).json({
      success: false,
      error: "Error obteniendo insumos desde ERP",
      detalle: err.response?.data || err.message,
    });
  }
});

app.get("/api/insumos/detalle", async (req, res) => {
  try {
    const { codigo } = req.query;

    if (!codigo || !codigo.trim()) {
      return res.status(400).json({
        success: false,
        error: "Falta código de insumo",
      });
    }

    const codigoLimpio = codigo.trim().toUpperCase();

    console.log("Buscando detalle de insumo en ERP:", codigoLimpio);

    const token = await getAccessToken();

    if (!token) {
      return res.status(500).json({
        success: false,
        error: "No se pudo obtener token del ERP",
      });
    }

    const tenantId = process.env.TENANT_ID;
    const environment = process.env.BC_ENVIRONMENT;

    if (!tenantId || !environment) {
      return res.status(500).json({
        success: false,
        error: "Faltan variables TENANT_ID o BC_ENVIRONMENT",
      });
    }

    const baseUrl = `https://api.businesscentral.dynamics.com/v2.0/${tenantId}/${environment}/api/v2.0`;

    // 1. Obtener compañías
    const companiesRes = await axios.get(`${baseUrl}/companies`, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
      },
    });

    const companies = companiesRes.data?.value || [];

    if (!companies.length) {
      return res.status(404).json({
        success: false,
        error: "No se encontraron compañías en Business Central",
      });
    }

    // Puedes dejar la primera o luego fijar una por nombre si quieres
    const companyId = companies[0].id;

    // 2. Buscar el item por código
    const itemsRes = await axios.get(
      `${baseUrl}/companies(${companyId})/items`,
      {
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: "application/json",
        },
        params: {
          $filter: `number eq '${codigoLimpio}'`,
        },
      }
    );

    const items = itemsRes.data?.value || [];

    if (!items.length) {
      return res.status(404).json({
        success: false,
        error: "Insumo no encontrado en ERP",
        codigoBuscado: codigoLimpio,
      });
    }

    const item = items[0];

    return res.json({
      success: true,
      insumo: {
        codigo: item.number || codigoLimpio,
        descripcion: item.displayName || item.description || "Sin descripción",
      },
      datos: {
        id: item.id || "",
        number: item.number || "",
        displayName: item.displayName || "",
        description: item.description || "",
        baseUnitOfMeasure: item.baseUnitOfMeasure || "",
        unitPrice: item.unitPrice || 0,
        blocked: item.blocked || false,
        inventory: item.inventory || 0,
        type: item.type || "",
        lastModifiedDateTime: item.lastModifiedDateTime || "",
      },
    });
  } catch (err) {
    console.error("❌ ERROR en /api/insumos/detalle ERP:");
    console.error("Mensaje:", err.message);
    console.error("Detalle ERP:", err.response?.data || "Sin detalle");

    return res.status(500).json({
      success: false,
      error: "Error interno al consultar ERP",
      detalle: err.response?.data || err.message,
    });
  }
});

app.get("/api/insumos/lote", async (req, res) => {
  try {
    const { codigo } = req.query;

    if (!codigo) {
      return res.status(400).json({ error: "Falta código de insumo" });
    }

    console.log("Buscando insumo:", codigo);

    if (!fs.existsSync(INSUMOS_PATH)) {
      return res.status(404).json({
        error: "No existe INSUMOS.xlsx en backend",
        path: INSUMOS_PATH,
        sugerencia: "Verifica la ruta del archivo"
      });
    }

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(INSUMOS_PATH);

    const ws = wb.worksheets[0];

    if (!ws) {
      return res.status(404).json({
        error: "INSUMOS.xlsx sin hojas"
      });
    }

    console.log(`Excel cargado. Filas: ${ws.rowCount}, Columnas: ${ws.columnCount}`);

    const headerRow = ws.getRow(1);
    const headers = [];

    headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const headerText = norm(String(cell.text || cell.value || "").trim());
      headers[colNumber] = headerText;
    });

    console.log("Encabezados encontrados:", headers.filter(h => h));

    // buscar columna codigo
    let idxCodigo = -1;

    for (let i = 1; i < headers.length; i++) {
      const header = headers[i] || "";

      console.log(`Columna ${i}: "${header}"`);

      if (header.includes("CODIGO DE INSUMO") || header.includes("CÓDIGO DE INSUMO")) {
        idxCodigo = i;
        console.log(`Columna código encontrada en índice ${i}`);
        break;
      }
    }

    if (idxCodigo === -1) {
      return res.status(400).json({
        error: "No se encontró columna 'CÓDIGO DE INSUMO'",
        headers_encontrados: headers.filter(h => h),
        sugerencia: "Revisa los nombres de las columnas en tu Excel"
      });
    }

    // buscar columna lote
    let idxLote = -1;

    for (let i = 1; i < headers.length; i++) {
      const header = headers[i] || "";

      if (header.includes("LOTE")) {
        idxLote = i;
        console.log(`Columna lote encontrada en índice ${i}`);
        break;
      }
    }

    let encontrado = false;
    let lote = "";
    let datosCompletos = {};

    console.log(`Buscando código "${codigo}" en ${ws.rowCount - 1} filas...`);

    for (let rowNumber = 2; rowNumber <= ws.rowCount; rowNumber++) {

      if (encontrado) break;

      const row = ws.getRow(rowNumber);
      if (!row) continue;

      const celdaCodigo = row.getCell(idxCodigo);

      let valorInsumo = "";

      if (celdaCodigo && celdaCodigo.value !== null && celdaCodigo.value !== undefined) {

        const valor = celdaCodigo.value;

        if (typeof valor === "object" && valor.result !== undefined) {
          valorInsumo = String(valor.result);
        } else if (typeof valor === "object" && valor.text !== undefined) {
          valorInsumo = String(valor.text);
        } else {
          valorInsumo = String(valor);
        }

      } else if (celdaCodigo && celdaCodigo.text !== undefined) {
        valorInsumo = String(celdaCodigo.text);
      }

      valorInsumo = valorInsumo.trim();

      if (valorInsumo === codigo) {

        encontrado = true;

        console.log(`Encontrado en fila ${rowNumber}! Código: "${valorInsumo}"`);

        // obtener lote
        if (idxLote !== -1) {

          const celdaLote = row.getCell(idxLote);
          let valorLote = "";

          if (celdaLote && celdaLote.value !== null && celdaLote.value !== undefined) {

            const valor = celdaLote.value;

            if (typeof valor === "object" && valor.result !== undefined) {
              valorLote = String(valor.result);
            } else if (typeof valor === "object" && valor.text !== undefined) {
              valorLote = String(valor.text);
            } else {
              valorLote = String(valor);
            }

          } else if (celdaLote && celdaLote.text !== undefined) {
            valorLote = String(celdaLote.text);
          }

          lote = valorLote.trim();
        }

        // obtener todos los datos de la fila
        headers.forEach((header, idx) => {

          if (header && header !== "" && idx !== 0) {

            const celda = row.getCell(idx);
            let valorCelda = "";

            if (celda && celda.value !== null && celda.value !== undefined) {

              const valor = celda.value;

              if (typeof valor === "object" && valor.result !== undefined) {
                valorCelda = String(valor.result);
              } else if (typeof valor === "object" && valor.text !== undefined) {
                valorCelda = String(valor.text);
              } else {
                valorCelda = String(valor);
              }

            } else if (celda && celda.text !== undefined) {
              valorCelda = String(celda.text);
            }

            datosCompletos[header] = valorCelda.trim();
          }

        });

        console.log("Datos encontrados:", datosCompletos);
      }

    }

    if (!encontrado) {

      console.log(`Código: "${codigo}" no encontrado en el archivo`);

      return res.status(404).json({
        error: "Insumo no encontrado",
        codigoBuscado: codigo,
        total_filas_revisadas: ws.rowCount - 1,
        sugerencia: `Verifica que el código "${codigo}" exista en INSUMOS.xlsx`
      });

    }

    return res.json({
      codigo,
      lote,
      datos: datosCompletos
    });

  } catch (error) {

    console.error("Error leyendo Excel:", error);

    return res.status(500).json({
      error: "Error procesando archivo Excel",
      detalle: error.message
    });

  }
});

app.get("/api/actividad/cantidadPorHora", async (req, res) => {
  try {
    const { actividad } = req.query;
    
    if (!actividad) {
      return res.status(400).json({ error: "Falta actividad" });
    }
    
    console.log("Buscando actividad:", actividad);
    
    // Verificar que existe el archivo
    if (!fs.existsSync(ACTIVIDAD_PATH)) {
      return res.status(404).json({
        error: "No existe archivo de cantidades por actividad",
        path: ACTIVIDAD_PATH
      });
    }
    
    // Cargar Excel
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(ACTIVIDAD_PATH);
    
    const ws = wb.worksheets[0];
    if (!ws) {
      return res.status(404).json({ error: "El Excel no tiene hojas" });
    }
    
    console.log(`Excel cargado. Filas: ${ws.rowCount}, Columnas: ${ws.columnCount}`);
    
    // Leer encabezados
    const headerRow = ws.getRow(1);
    const headers = [];
    
    headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const headerText = String(cell.text || cell.value || "").trim().toUpperCase();
      headers[colNumber] = headerText;
    });
    
    console.log("Encabezados:", headers.filter(h => h));
    
    // Buscar columna de DESCRIPCIÓN (actividad)
    let idxActividad = -1;
    for (let i = 1; i < headers.length; i++) {
      if (headers[i]?.includes("DESCRIPCIÓN") || headers[i]?.includes("DESCRIPCION")) {
        idxActividad = i;
        console.log(`Columna DESCRIPCIÓN encontrada en índice ${i}`);
        break;
      }
    }
    
    // Buscar columna de CANTIDAD POR HORA
    let idxCantidadHora = -1;
    for (let i = 1; i < headers.length; i++) {
      if (headers[i]?.includes("CANTIDAD POR HORA")) {
        idxCantidadHora = i;
        console.log(`Columna CANTIDAD POR HORA encontrada en índice ${i}`);
        break;
      }
    }
    
    // Validar que encontró ambas columnas
    if (idxActividad === -1) {
      return res.status(400).json({
        error: "No se encontró columna DESCRIPCIÓN",
        headers: headers.filter(h => h)
      });
    }
    
    if (idxCantidadHora === -1) {
      return res.status(400).json({
        error: "No se encontró columna CANTIDAD POR HORA",
        headers: headers.filter(h => h)
      });
    }
    
    // Buscar la actividad en las filas
    let encontrado = false;
    let cantidadPorHora = "";
    let datosCompletos = {};
    
    console.log(`Buscando actividad "${actividad}" en ${ws.rowCount - 1} filas...`);
    
    for (let rowNumber = 2; rowNumber <= ws.rowCount; rowNumber++) {
      const row = ws.getRow(rowNumber);
      if (!row) continue;
      
      // Obtener valor de la celda de actividad
      const celdaActividad = row.getCell(idxActividad);
      let valorActividad = "";
      
      if (celdaActividad) {
        if (celdaActividad.value !== null && celdaActividad.value !== undefined) {
          valorActividad = String(celdaActividad.value?.result || celdaActividad.value?.text || celdaActividad.value || "");
        } else if (celdaActividad.text) {
          valorActividad = String(celdaActividad.text);
        }
      }
      
      valorActividad = valorActividad.trim().toUpperCase();
      
      // Comparar con la actividad buscada
      if (valorActividad === actividad.toUpperCase()) {
        encontrado = true;
        console.log(`Actividad encontrada en fila ${rowNumber}`);
        
        // Obtener cantidad por hora
        const celdaCantidad = row.getCell(idxCantidadHora);
        
        if (celdaCantidad) {
          if (celdaCantidad.value !== null && celdaCantidad.value !== undefined) {
            cantidadPorHora = String(celdaCantidad.value?.result || celdaCantidad.value?.text || celdaCantidad.value || "");
          } else if (celdaCantidad.text) {
            cantidadPorHora = String(celdaCantidad.text);
          }
        }
        
        cantidadPorHora = cantidadPorHora.trim();
        
        // Obtener todos los datos de la fila (opcional)
        headers.forEach((header, idx) => {
          if (header && idx > 0) {
            const celda = row.getCell(idx);
            let valor = "";
            
            if (celda) {
              if (celda.value !== null && celda.value !== undefined) {
                valor = String(celda.value?.result || celda.value?.text || celda.value || "");
              } else if (celda.text) {
                valor = String(celda.text);
              }
            }
            
            datosCompletos[header] = valor.trim();
          }
        });
        
        break; // Salir del bucle al encontrar
      }
    }
    
    if (!encontrado) {
      console.log(`Actividad "${actividad}" no encontrada`);
      return res.status(404).json({
        error: "Actividad no encontrada",
        actividad: actividad,
        sugerencia: "Verifica que el nombre de la actividad coincida exactamente"
      });
    }
    
    // Respuesta exitosa
    return res.json({
      actividad: actividad,
      cantidad_por_hora: cantidadPorHora,
      datos: datosCompletos
    });
    
  } catch (error) {
    console.error("❌ Error leyendo Excel:", error);
    return res.status(500).json({
      error: "Error procesando archivo Excel",
      detalle: error.message
    });
  }
});

app.get("/api/procesos/producto", async (req, res) => {
  try {
    const { codigo } = req.query;
    if (!codigo) return res.status(400).json({ error: "Falta codigo" });

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(INSUMOS_PATH);

    const ws = wb.worksheets[0];
    const procesos = [];

    // Leer encabezados
    const headerRow = ws.getRow(1);
    const headers = [];
    headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      headers[colNumber] = norm(String(cell.text || "").trim());
    });

    const idxProducto = headers.findIndex(h => h === "CÓDIGO" || h === "CODIGO");
    const idxProceso = headers.findIndex(h => h === "PROCESO");

    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber === 1) return;

      const productoCell = row.getCell(idxProducto);
      const producto = String(productoCell.text || productoCell.value || "").trim();

      if (producto === codigo) {
        const procesoCell = row.getCell(idxProceso);
        const proceso = String(procesoCell.text || procesoCell.value || "").trim();
        if (proceso) procesos.push(proceso);
      }
    });

    // Unir en texto con saltos de línea
    const texto = procesos.join("\n");

    res.json({ detalles: texto });

  } catch (err) {
    console.error("Error obteniendo procesos por producto:", err);
    res.status(500).json({ error: "Error interno" });
  }
});

app.get("/api/cantidades/producto", async (req, res) => {
  try {
    const { codigo } = req.query;
    if (!codigo) return res.status(400).json({ error: "Falta codigo" });

    console.log("Buscando producto:", codigo);

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(PRODUCTOS_PATH);

    const ws = wb.worksheets.find((w) => norm(w.name) === "PRODUCTOS");
    if (!ws) {
      return res.status(404).json({
        error: "No existe hoja PRODUCTOS",
        hojas: wb.worksheets.map((w) => w.name),
      });
    }

    const HEADER_ROW = 1;
    const headers = ws.getRow(HEADER_ROW).values.slice(1).map((v) => norm(v));

    const idxCodigo = headers.findIndex((h) => h === "CÓDIGO" || h === "CODIGO");
    const idxMeta = headers.findIndex((h) => h === "CANTIDAD POR HORA");

    if (idxCodigo === -1 || idxMeta === -1) {
      return res.status(400).json({
        error: "No existe columna CANTIDAD POR HORA o CÓDIGO en fila 1",
        headers_detectados: headers,
      });
    }

    let metaEncontrada = null;
    
    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber <= HEADER_ROW || metaEncontrada != null) return;
      
      const codigoFila = String(row.getCell(idxCodigo + 1).value ?? "").trim();
      
      if (norm(codigoFila) === norm(codigo)) {
        console.log(`Producto encontrado en fila ${rowNumber}`);
        
        // 🔥 OBTENER EL VALOR CORRECTAMENTE
        const celdaValor = row.getCell(idxMeta + 1);
        let valorExtraido = "";
        
        // Extraer el valor según su tipo
        if (celdaValor.value === null || celdaValor.value === undefined) {
          valorExtraido = "";
        } 
        else if (typeof celdaValor.value === 'object') {
          // Es un objeto (posiblemente RichText o fórmula)
          if (celdaValor.value.richText) {
            // Texto enriquecido
            valorExtraido = celdaValor.value.richText.map(r => r.text).join('');
          } else if (celdaValor.value.text) {
            // Tiene propiedad text
            valorExtraido = celdaValor.value.text;
          } else if (celdaValor.value.result) {
            // Es una fórmula con resultado
            valorExtraido = String(celdaValor.value.result);
          } else {
            // Intentar convertir a string
            valorExtraido = celdaValor.value.toString();
            if (valorExtraido === '[object Object]') {
              valorExtraido = "";
            }
          }
        } 
        else {
          // Valor simple (string, number)
          valorExtraido = String(celdaValor.value);
        }
        
        // Limpiar y guardar
        metaEncontrada = valorExtraido.trim();
        console.log(`Valor encontrado: "${metaEncontrada}"`);
      }
    });

    if (metaEncontrada === null) {
      return res.status(404).json({ error: "Código de producto no encontrado" });
    }

    // Asegurar que sea un número válido
    const numValor = parseFloat(metaEncontrada);
    if (isNaN(numValor)) {
      console.warn(`⚠️ El valor "${metaEncontrada}" no es un número válido`);
    }

    res.json({ meta: metaEncontrada });

  } catch (err) {
    console.error("ERROR GET /api/cantidades/producto:", err);
    res.status(500).json({ error: "Error leyendo Cantidades" });
  }
});


app.get("/api/registros/mi-perfil", (req, res) => {
  try {
    console.log("=== RUTA MI-PERFIL LLAMADA ===");
    console.log("Query params:", req.query);

    const { nombre, rol } = req.query;

    // Validar que nombre y rol existan
    if (!nombre || !rol) {
      return res.status(400).json({ error: "Faltan parámetros nombre o rol" });
    }

    let sql = "";
    let params = [];

    const rolNormalizado = (rol || "").toUpperCase().trim();

    if (rolNormalizado === "SUPERVISOR") {
      sql = `SELECT * FROM registros WHERE supervisor LIKE ?`;
      params = [`%${nombre}%`];
      console.log("Buscando como SUPERVISOR con LIKE");
    } 
    else if (rolNormalizado === "LÍDER") {
      sql = `SELECT * FROM registros WHERE responsable LIKE ?`;
      params = [`%${nombre}%`];
      console.log("Buscando como LÍDER con LIKE");
    } 
    else if (rolNormalizado === "ANALISTA DE PRODUCCIÓN") {
      sql = `
        SELECT * FROM registros
        WHERE estado = 'pendiente_ANALISTA_PRODUCCION' 
        ORDER BY fecha_verificacion DESC
      `;
      console.log("Buscando como ANALISTA (pendientes de aprobación)");
    } 
    else if (rolNormalizado === "JEFE DE PRODUCCIÓN") {
      sql = `
        SELECT * FROM registros
        ORDER BY fecha DESC
      `;
      console.log("Buscando como JEFE (todos los registros)");
    }
    else if (rolNormalizado === "ADMINISTRADOR") {
      sql =`
        SELECT id, username, nombre, cedula_identidad, rol, area, activo, primer_login,
               created_at, updated_at 
        FROM usuarios 
        ORDER BY id DESC
      `
    }
    else {
      console.log("Rol no válido:", rol);
      return res.json([]);
    }

    console.log("SQL:", sql);
    console.log("Parámetros:", params);

    // 🔥 FUNCIÓN SEGURA PARA PARSEAR JSON
    const safeParse = (value) => {
      if (!value) return [];
      if (typeof value === "object") return value;
      if (typeof value === "string") {
        try {
          return JSON.parse(value);
        } catch (e) {
          console.error("❌ Error parseando campo:", value.substring(0, 100));
          return []; // Array vacío en caso de error
        }
      }
      return [];
    };

    db.all(sql, params, (err, rows) => {
      if (err) {
        console.error("❌ Error SQL:", err);
        return res.status(500).json({ error: err.message });
      }

      try {
        const registros = (rows || []).map(r => {
          try {
            return {
              ...r,
              insumos: safeParse(r.insumos),
              etiquetas: safeParse(r.etiquetas),
              integrantes: safeParse(r.integrantes),
              actividades_por_integrante: safeParse(r.actividades_por_integrante),
              planificada_total_por_detalle: safeParse(r.planificada_total_por_detalle),   
              elaborada_total_por_detalle: safeParse(r.elaborada_total_por_detalle),        
              maquinarias: safeParse(r.maquinarias)
            };
          } catch (mapError) {
            console.error(" Error procesando registro:", r.id, mapError);
            return r; // Devuelve el registro sin parsear
          }
        });

        console.log(`Registros encontrados: ${registros.length}`);
        
        // Verificar que se puede convertir a JSON
        const testJson = JSON.stringify(registros);
        console.log(`📦 JSON generado: ${testJson.length} bytes`);
        
        res.json(registros);
        
      } catch (parseError) {
        console.error("❌ Error fatal procesando registros:", parseError);
        res.status(500).json({ 
          error: "Error procesando registros",
          details: parseError.message 
        });
      }
    });
    
  } catch (outerError) {
    console.error("❌ Error fatal en el endpoint:", outerError);
    res.status(500).json({ error: "Error interno del servidor" });
  }
});

app.put("/api/registros/:id/verificar", (req, res) => {
  const body = req.body || {};
  const { id } = req.params;
  const { nombre, rol } = req.body;

  // Seguridad
  if (rol !== "SUPERVISOR") {
    return res.status(403).json({ error: "Solo el supervisor puede verificar" });
  }

  const sql = `
    UPDATE registros
    SET 
      verificado_por = ?,
      fecha_verificacion = datetime('now', '-5 hours'),
      estado = 'pendiente_ANALISTA_PRODUCCION'
    WHERE id = ?
  `;

  db.run(sql, [nombre, id], function (err) {
    if (err) {
      console.error(err);
      return res.status(500).json({ error: "Error al verificar registro" });
    }

    res.json({ mensaje: "Registro verificado correctamente" });
  });
});

app.delete("/api/registros/:id",(req,res) =>{
  const {id} = req.params;
  db.serialize(()=>{
    db.run("DELETE FROM registro_insumos WHERE registro_id =?",[id]);
    db.run("DELETE FROM registro_integrantes WHERE registro_id =?",[id]);
    db.run("DELETE FROM registros WHERE id =?",[id], function(err){
      if(err){
        console.error(err);
        return res.status(500).json({error:"Error al eliminar el registro"});
      }
      res.json({message:"Registro eliminado correctamente"});
    });
  });
});

app.get("/api/registros/:id", (req, res) => {
  const { id } = req.params;

  db.get("SELECT * FROM registros WHERE id = ?", [id], (err, registro) => {
    if (err) {
      console.error(err);
      return res.status(500).json({ error: "Error al obtener registro" });
    }

    if (!registro) {
      return res.status(404).json({ error: "Registro no encontrado" });
    }

    let insumos = [];
    let integrantes = [];
    let etiquetas = [];
    let maquinarias = [];
    let actividadesPorIntegrante = {};
    let planificadaPorDetalle = {};    
    let elaboradaPorDetalle = {};       

    try {
      insumos = JSON.parse(registro.insumos || "[]");
    } catch {}

    try {
      integrantes = JSON.parse(registro.integrantes || "[]");
    } catch {}

    try {
      etiquetas = JSON.parse(registro.etiquetas || "[]");
    } catch {}

    try {
      maquinarias = JSON.parse(registro.maquinarias || "[]");
    } catch {}

    try {
      actividadesPorIntegrante = JSON.parse(registro.actividades_por_integrante || "{}");
    } catch {}

    try {
      planificadaPorDetalle = JSON.parse(registro.planificada_total_por_detalle || "{}");
    } catch {}

    try {
      elaboradaPorDetalle = JSON.parse(registro.elaborada_total_por_detalle || "{}");
    } catch {}

    res.json({
      ...registro,
      insumos,
      integrantes,
      etiquetas,
      maquinarias,
      actividadesPorIntegrante,
      planificadaPorDetalle,    
      elaboradaPorDetalle        
    });
  });
});

app.get("/api/registros/admin/pendientes", (req, res) => {
  const sql = `
    SELECT * FROM registros
    WHERE estado = 'pendiente_admin'
    ORDER BY id DESC
  `;

  db.all(sql, [], (err, rows) => {
    if (err) {
      console.error(err);
      return res.status(500).json({ error: "Error al obtener registros" });
    }
    res.json(rows);
  });
});

app.get("/api/registros/admin/aprobados", (req, res) => {
  const sql = `
    SELECT * FROM registros
    WHERE estado = 'aprobado'
    ORDER BY id DESC
  `;

  db.all(sql, [], (err, rows) => {
    if (err) {
      console.error(err);
      return res.status(500).json({ error: "Error al obtener registros" });
    }

    res.json(rows);
  });
});

app.put("/api/registros/:id", (req, res) => {
  const { id } = req.params;
  const datos = req.body;
  
  // Obtener rol y nombre del cuerpo de la petición
  const { rol, nombre } = datos;

  if (!["SUPERVISOR","LÍDER"].includes(rol)) {
    return res.status(403).json({ error: "No tienes permisos" });
  }

  const campos = [];
  const valores = [];
  
  // Campos que deben guardarse como JSON
  const camposJSON = [
    'insumos',
    'maquinarias',
    'etiquetas',
    'integrantes',
    'reposicion_no_conforme',
    'actividades_por_integrante',
    'planificada_total_por_detalle',
    'elaborada_total_por_detalle'
  ];

  // ✅ CORREGIDO: Mapeo completo de campos permitidos
  const camposPermitidos = {
    fecha: 'fecha',
    op: 'op',
    turno: 'turno',
    area: 'area',
    modulo: 'modulo',
    codigo_producto: 'codigo_producto',
    referencia: 'codigo_producto',
    descripcion: 'descripcion',
    lotePrincipal: 'lotePrincipal',
    loteSecundario: 'loteSecundario',
    loteUnido: 'loteUnido',
    responsable: 'responsable',
    supervisor: 'supervisor',
    personal_asignado: 'personal_asignado',
    personal_otro: 'personal_otro',
    personal_presente: 'personal_presente',
    cantidad_elaborada: 'cantidad_elaborado',
    cantidad_elaborado: 'cantidad_elaborado',
    cantidad_proceso: 'cantidad_proceso',
    cantidad_merma: 'cantidad_merma',
    fecha_final_producto: 'fecha_final_producto',
    hora_inicio: 'hora_inicio',
    hora_fin: 'hora_fin',
    hora_planificada: 'hora_planificada',
    destino: 'destino',
    cliente: 'cliente',
    n_cliente: 'n_cliente',
    esteril: 'esteril',
    leyenda_si: 'leyenda_si',
    leyenda_otra: 'leyenda_otra',
    leyenda: 'leyenda',
    detalles_actividades: 'detalles_actividades',
    cantidad_planificada: 'cantidad_planificada',
    cantidad_elaborada_detalles: 'cantidad_elaborada_detalles',
    observaciones: 'observaciones'
  };

  // ✅ Procesar campos normales (del camposPermitidos)
  Object.keys(camposPermitidos).forEach(campoFront => {
    if (datos[campoFront] !== undefined && datos[campoFront] !== null) {
      const campoDb = camposPermitidos[campoFront];
      // Evitar duplicados
      if (!campos.some(c => c.includes(`${campoDb} =`))) {
        campos.push(`${campoDb} = ?`);
        valores.push(datos[campoFront]);

        // 🔍 LOG ESPECIAL PARA LOTE
        if (campoDb === 'lote') {
          console.log("🎫 LOTE A GUARDAR:", datos[campoFront]);
        }
      }
    }
  });

  // ✅ Procesar campos JSON por separado
  camposJSON.forEach(campoJSON => {
    if (datos[campoJSON] !== undefined && datos[campoJSON] !== null) {
      campos.push(`${campoJSON} = ?`);
      // Si ya es string, usarlo como está; si no, convertir a JSON
      const valor = typeof datos[campoJSON] === 'string'
        ? datos[campoJSON]
        : JSON.stringify(datos[campoJSON]);

      // 🔍 LOG ESPECIAL PARA INSUMOS
      if (campoJSON === 'insumos') {
        console.log("📦 INSUMOS A GUARDAR:", valor);
      }

      // 🔍 LOG ESPECIAL PARA MAQUINARIAS
      if (campoJSON === 'maquinarias') {
        console.log("🔧 MAQUINARIAS A GUARDAR:", valor);
      }

      valores.push(valor);
    }
  });

  // ✅ Verificar que hay campos para actualizar
  if (campos.length === 0) {
    return res.status(400).json({ error: "No hay campos para actualizar" });
  }

  // Siempre actualizar estos campos
  campos.push('verificado_por = ?');
  valores.push(nombre);
  
  campos.push('fecha_verificacion = datetime("now", "-5 hours")');
  
  // Manejar el estado
  if (datos.estado !== undefined) {
    campos.push('estado = ?');
    valores.push(datos.estado);
  } else {
    // Si no se envía estado, mantener el actual
    campos.push('estado = estado');
  }

  // Agregar el ID al final para el WHERE
  valores.push(id);

  const sql = `UPDATE registros SET ${campos.join(', ')} WHERE id = ?`;
  
  // ✅ LOG PARA DEPURACIÓN
  console.log("SQL:", sql);
  console.log("Valores:", valores);

  db.run(sql, valores, function(err) {
    if (err) {
      console.error("Error en UPDATE:", err);
      return res.status(500).json({ error: err.message });
    }

    console.log(`Registro ${id} actualizado. Filas afectadas: ${this.changes}`);

    // Recuperar el registro actualizado
    db.get('SELECT * FROM registros WHERE id = ?', [id], (err, registro) => {
      if (err) {
        console.error("Error al recuperar registro:", err);
        return res.json({ 
          mensaje: "Actualizado correctamente pero error al recuperar datos" 
        });
      }

      if (!registro) {
        return res.status(404).json({ error: "Registro no encontrado después de actualizar" });
      }

      // Parsear los campos JSON antes de enviar al frontend
      camposJSON.forEach(campo => {
        if (registro[campo]) {
          try {
            registro[campo] = JSON.parse(registro[campo]);
          } catch (e) {
            console.error(`Error parseando ${campo}:`, e);
            registro[campo] = campo === 'actividades_por_integrante' ? {} : [];
          }
        } else {
          registro[campo] = campo === 'actividades_por_integrante' ? {} : [];
        }
      });

      res.json({ 
        mensaje: "Actualizado correctamente",
        registro: registro 
      });
    });
  });
});
// ============================================
// ENDPOINT PARA OBTENER MÁQUINAS POR MÓDULO (VERSIÓN CORREGIDA)
// ============================================
app.get("/api/modulos/maquinas", async (req, res) => {
  try {
    const { modulo } = req.query;
    if (!modulo) return res.status(400).json({ error: "Falta modulo" });

    if (!fs.existsSync(MODULOS_PATH)) {
      return res.status(404).json({ error: "No existe modulos.xlsx en backend" });
    }

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(MODULOS_PATH);

    const norm = (s) =>
      String(s ?? "")
        .trim()
        .toUpperCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/\s+/g, " ");

    const moduloParam = String(modulo).trim();
    
    // Buscar la hoja del módulo
    const wanted = [];
    if (/^\d+$/.test(moduloParam)) {
      wanted.push(`MODULO ${moduloParam}`);
      wanted.push(`MÓDULO ${moduloParam}`);
    } else {
      wanted.push(moduloParam);
    }
    const wantedNorms = wanted.map(norm);

    const ws = wb.worksheets.find((w) => wantedNorms.includes(norm(w.name)));

    if (!ws) {
      return res.status(404).json({
        error: `No existe la hoja para '${moduloParam}'`,
        hojasDisponibles: wb.worksheets.map((w) => w.name),
      });
    }

    // Buscar las columnas de máquinas en un rango AMPLIO (hasta columna 50)
    const HEADER_ROW = 4; // Fila donde están los encabezados
    
    const headerRow = ws.getRow(HEADER_ROW);
    let COL_MAQUINARIA = -1;
    let COL_CANTIDAD = -1;
    
    console.log(" Buscando columnas en el Excel...");
    // Buscar en un rango amplio de columnas (hasta la 50)
    for (let c = 1; c <= 50; c++) {
      const cell = headerRow.getCell(c);
      const cellValue = String(cell.text || cell.value || "").trim().toUpperCase();
      
      if (cellValue) {
        console.log(`Columna ${c}: "${cellValue}"`);
      }
      
      if (cellValue.includes("MAQUINARIA") && !cellValue.includes("CANTIDAD")) {
        COL_MAQUINARIA = c;
        console.log(`Columna MAQUINARIA encontrada en columna ${c}`);
      }
      if (cellValue.includes("CANTIDAD") || cellValue.includes("CANTIDAD DE MAQUINARIA")) {
        COL_CANTIDAD = c;
        console.log(`Columna CANTIDAD encontrada en columna ${c}`);
      }
    }

    if (COL_MAQUINARIA === -1 || COL_CANTIDAD === -1) {
      console.log("❌ No se encontraron las columnas de maquinaria");
      // Devolver datos de prueba para que el frontend pueda desarrollar
      return res.json({
        modulo: moduloParam,
        maquinas: [
          { nombre: "RECTA", cantidad: 3 },
          { nombre: "OVERLOK", cantidad: 1 },
          { nombre: "ULTRASONIDO", cantidad: 3 }
        ],
        advertencia: "Usando datos de prueba - Configure las columnas en Excel"
      });
    }

    // Leer los datos desde la fila 5 en adelante
    const START_ROW = 5;
    const maquinasMap = new Map();

    console.log("Leyendo datos de máquinas...");
    for (let r = START_ROW; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      
      const maquina = String(row.getCell(COL_MAQUINARIA).text || row.getCell(COL_MAQUINARIA).value || "").trim();
      const cantidadStr = String(row.getCell(COL_CANTIDAD).text || row.getCell(COL_CANTIDAD).value || "").trim();
      
      if (maquina && cantidadStr) {
        console.log(`Fila ${r}: Máquina="${maquina}", Cantidad="${cantidadStr}"`);
        
        const cantidad = parseInt(cantidadStr);
        if (!isNaN(cantidad) && cantidad > 0) {
          if (maquinasMap.has(maquina)) {
            maquinasMap.set(maquina, maquinasMap.get(maquina) + cantidad);
          } else {
            maquinasMap.set(maquina, cantidad);
          }
        }
      }
    }

    const resultados = [];
    maquinasMap.forEach((cantidad, nombre) => {
      resultados.push({
        nombre: nombre,
        cantidad: cantidad
      });
    });

    console.log("Máquinas encontradas:", resultados);

    res.json({
      modulo: moduloParam,
      maquinas: resultados
    });

  } catch (err) {
    console.error("ERROR /api/modulos/maquinas:", err);
    return res.status(500).json({
      error: "Error leyendo máquinas del módulo",
      message: err.message,
    });
  }
});

/* =========================
   START LOCAL SERVER
========================= */
const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => {
    console.log(` Backend corriendo en puerto ${PORT}`);
});

// Función para crear usuarios de prueba (opcional - ya se crean en db.js)
async function crearUsuariosPrueba() {
  // Esta función ya no es necesaria porque los usuarios se crean en db.js
  console.log(" Usuarios de prueba gestionados desde db.js");
}

// Ejecutar la creación de usuarios de prueba después de un pequeño delay
setTimeout(() => {
  crearUsuariosPrueba();
}, 1000);