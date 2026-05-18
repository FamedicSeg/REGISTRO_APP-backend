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

// IMPORTAR SERVICIO DE ONEDRIVE
const {
  getOPListaFromOneDrive,
  getModuloPersonalFromOneDrive,
  getLoteInfoFromOneDrive,
  getInsumoLoteFromOneDrive,
  getActividadCantidadPorHoraFromOneDrive,
  getProcesosProductoFromOneDrive,
  getCantidadesProductoFromOneDrive,
  getModuloMaquinasFromOneDrive,
  saveRegistroToExcel
} = require('./services/oneDriveExcelService');

app.use(
  cors({
    origin: [
      "http://localhost:5173",
      "http://127.0.0.1:5173",
      "http://192.168.1.46:5173",
      "http://192.168.4.147:5173",
      "http://192.168.3.106:5173",
      /https:\/\/.*\.vercel\.app$/,
      /^http:\/\/192\.168\.\d{1,3}\.\d{1,3}:5173$/,
      /https:\/\/.*\.loca\.lt$/,
      /https:\/\/.*\.trycloudflare\.com$/,
    ],
    methods: ["GET", "POST", "PUT", "PATCH", "DELETE", "OPTIONS"],
    allowedHeaders: ["Content-Type", "Authorization"],
    credentials: true,
  })
);

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

/* =========================
   PATHS (YA NO SE USAN PARA LECTURA DIRECTA, SE USA ONEDRIVE)
   Se mantienen por si acaso pero ya no se usan en los endpoints modificados
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

const norm = (s) =>
  String(s ?? "")
    .trim()
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");

const toNum = (v, def = 0) => {
  if (v === "" || v === null || v === undefined) return def;
  const n = Number(v);
  return Number.isFinite(n) ? n : def;
};

// ============================================
// FUNCIÓN PARA GUARDAR EN EXCEL (AHORA USA ONEDRIVE)
// ============================================
async function guardarEnExcel(registro) {
  try {
    await saveRegistroToExcel(registro);
    console.log(`Registro ${registro.id} guardado en OneDrive`);
    return true;
  } catch (error) {
    console.error(`Error guardando en OneDrive:`, error);
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

  if (!["ANALISTA DE PRODUCCIÓN", "SUPERVISOR"].includes(rol)) {
    return res.status(403).json({ error: "Solo el analista o supervisor pueden rechazar registros" });
  }

  if (!motivo || motivo.trim() === "") {
    return res.status(400).json({ error: "Debe proporcionar un motivo de rechazo" });
  }

  try {
    const registro = await new Promise((resolve, reject) => {
      db.get("SELECT * FROM registros WHERE id = ?", [id], (err, row) => {
        if (err) reject(err);
        else resolve(row);
      });
    });

    if (!registro) {
      return res.status(404).json({ error: "Registro no encontrado" });
    }

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
// ENDPOINT PARA APROBAR (AHORA GUARDA EN EXCEL VIA ONEDRIVE)
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

    const registroActualizado = await new Promise((resolve, reject) => {
      db.get("SELECT * FROM registros WHERE id = ?", [id], (err, row) => {
        if (err) reject(err);
        else {
          console.log("Registro DESPUÉS:", row);
          resolve(row);
        }
      });
    });

    await guardarEnExcel(registroActualizado);

    res.json({ 
      mensaje: "Registro aprobado y guardado en OneDrive correctamente",
      guardadoEnExcel: true,
      registro: registroActualizado
    });

  } catch (err) {
    console.error("❌ Error en aprobación:", err);
    res.status(500).json({ error: "Error al aprobar registro" });
  }
});

// ============================================
// ENDPOINT ORIGINAL POST (SIN CAMBIOS)
// ============================================
app.post("/api/registros", async (req, res) => {
  try {
    const data = req.body;

    const insumosJSON = JSON.stringify(data.insumos || []);
    const etiquetasJSON = JSON.stringify(data.etiquetas || []);
    const integrantesJSON = JSON.stringify(data.integrantes || []);
    const maquinasJSON = JSON.stringify(data.maquinarias || []);
    const actividadesPorIntegranteJSON = data.actividades_por_integrante || "{}";
    
    const planificadaPorDetalle = data.planificada_por_detalle || {};
    const elaboradaPorDetalle = data.elaborada_por_detalle || {};
    
    const planificadaPorDetalleJSON = JSON.stringify(planificadaPorDetalle);
    const elaboradaPorDetalleJSON = JSON.stringify(elaboradaPorDetalle);

    if (!data.fecha || !data.op || !data.modulo) {
      return res.status(400).json({
        error: "Faltan campos obligatorios (fecha, op, modulo).",
      });
    }

    console.log("BODY COMPLETO:", data);

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
        ? "Registro aprobado y guardado en OneDrive" 
        : "Estado actualizado correctamente" 
    });

  } catch (err) {
    console.error("Error actualizando estado:", err);
    res.status(500).json({ error: "Error al actualizar estado" });
  }
});

// ============================================
// ENDPOINT MODIFICADO: /api/op/lista (AHORA USA ONEDRIVE)
// ============================================
app.get("/api/op/lista", async (req, res) => {
  try {
    const resultado = await getOPListaFromOneDrive();
    res.json(resultado);
  } catch (err) {
    console.error("ERROR GET /api/op/lista desde OneDrive:", err);
    res.status(500).json({ error: "Error leyendo OPs desde OneDrive", detalle: err.message });
  }
});

// ============================================
// ENDPOINT MODIFICADO: /api/modulos/personal (AHORA USA ONEDRIVE)
// ============================================
app.get("/api/modulos/personal", async (req, res) => {
  try {
    const { modulo } = req.query;
    if (!modulo) return res.status(400).json({ error: "Falta modulo" });
    
    const resultado = await getModuloPersonalFromOneDrive(modulo);
    res.json(resultado);
  } catch (err) {
    console.error("ERROR /api/modulos/personal desde OneDrive:", err);
    res.status(500).json({ error: "Error leyendo personal desde OneDrive", message: err.message });
  }
});

// ============================================
// ENDPOINT MODIFICADO: /api/lote/info (AHORA USA ONEDRIVE)
// ============================================
app.get("/api/lote/info", async (req, res) => {
  try {
    const { codigo } = req.query;
    if (!codigo) return res.status(400).json({ error: "Falta codigo" });
    
    const resultado = await getLoteInfoFromOneDrive(codigo);
    res.json(resultado);
  } catch (err) {
    console.error("ERROR GET /api/lote/info desde OneDrive:", err);
    res.status(500).json({ error: "Error leyendo Lote.xlsx desde OneDrive", detalle: err.message });
  }
});

// ============================================
// ENDPOINT MODIFICADO: /api/insumos/lote (AHORA USA ONEDRIVE)
// ============================================
app.get("/api/insumos/lote", async (req, res) => {
  try {
    const { codigo } = req.query;
    if (!codigo) return res.status(400).json({ error: "Falta código" });
    
    const resultado = await getInsumoLoteFromOneDrive(codigo);
    res.json(resultado);
  } catch (err) {
    console.error("Error en /api/insumos/lote desde OneDrive:", err);
    res.status(500).json({ error: "Error procesando archivo Excel desde OneDrive", detalle: err.message });
  }
});

// ============================================
// ENDPOINT MODIFICADO: /api/actividad/cantidadPorHora (AHORA USA ONEDRIVE)
// ============================================
app.get("/api/actividad/cantidadPorHora", async (req, res) => {
  try {
    const { actividad } = req.query;
    if (!actividad) return res.status(400).json({ error: "Falta actividad" });
    
    const resultado = await getActividadCantidadPorHoraFromOneDrive(actividad);
    res.json(resultado);
  } catch (err) {
    console.error("Error en /api/actividad/cantidadPorHora desde OneDrive:", err);
    res.status(500).json({ error: "Error procesando archivo Excel desde OneDrive", detalle: err.message });
  }
});

// ============================================
// ENDPOINT MODIFICADO: /api/procesos/producto (AHORA USA ONEDRIVE)
// ============================================
app.get("/api/procesos/producto", async (req, res) => {
  try {
    const { codigo } = req.query;
    if (!codigo) return res.status(400).json({ error: "Falta codigo" });
    
    const resultado = await getProcesosProductoFromOneDrive(codigo);
    res.json(resultado);
  } catch (err) {
    console.error("Error en /api/procesos/producto desde OneDrive:", err);
    res.status(500).json({ error: "Error interno desde OneDrive" });
  }
});

// ============================================
// ENDPOINT MODIFICADO: /api/cantidades/producto (AHORA USA ONEDRIVE)
// ============================================
app.get("/api/cantidades/producto", async (req, res) => {
  try {
    const { codigo } = req.query;
    if (!codigo) return res.status(400).json({ error: "Falta codigo" });
    
    const resultado = await getCantidadesProductoFromOneDrive(codigo);
    res.json(resultado);
  } catch (err) {
    console.error("ERROR GET /api/cantidades/producto desde OneDrive:", err);
    res.status(500).json({ error: "Error leyendo Cantidades desde OneDrive", detalle: err.message });
  }
});

// ============================================
// ENDPOINT MODIFICADO: /api/modulos/maquinas (AHORA USA ONEDRIVE)
// ============================================
app.get("/api/modulos/maquinas", async (req, res) => {
  try {
    const { modulo } = req.query;
    if (!modulo) return res.status(400).json({ error: "Falta modulo" });
    
    const resultado = await getModuloMaquinasFromOneDrive(modulo);
    res.json(resultado);
  } catch (err) {
    console.error("ERROR /api/modulos/maquinas desde OneDrive:", err);
    return res.status(500).json({
      error: "Error leyendo máquinas del módulo desde OneDrive",
      message: err.message,
    });
  }
});

// ============================================
// LOS SIGUIENTES ENDPOINTS NO CAMBIAN (USAN ERP O BD)
// ============================================

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
// RUTAS PARA CRUD DE USUARIOS (SIN CAMBIOS)
// ============================================

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

app.put("/api/usuarios/:id/cambiar-primer-login", async (req, res) => {
  const { id } = req.params;
  const { nuevaPassword } = req.body;

  if (!nuevaPassword || nuevaPassword.length < 4) {
    return res.status(400).json({ error: "La contraseña debe tener al menos 4 caracteres" });
  }

  try {
    const hashedPassword = await bcrypt.hash(nuevaPassword, 10);

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

app.get("/api/usuarios/:id/con-password", async (req, res) => {
  const { id } = req.params;
  const { rolSolicitante } = req.query;

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

app.post("/api/usuarios", async (req, res) => {
  const { username, password, nombre, cedula_identidad, rol, area, activo } = req.body;
  
  if (!username || !password || !nombre || !cedula_identidad || !rol) {
    return res.status(400).json({ 
      error: 'Faltan campos requeridos: username, password, nombre, cedula_identidad, rol' 
    });
  }

  try {
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

    const hashedPassword = await bcrypt.hash(password, 10);
    
    const result = await new Promise((resolve, reject) => {
      db.run(
        `INSERT INTO usuarios (username, password, nombre, cedula_identidad, rol, area, activo, primer_login) 
         VALUES (?, ?, ?, ?, ?, ?, ?, 1)`,
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

app.put("/api/usuarios/:id/cambiar-password", async (req, res) => {
  const { id } = req.params;
  const { nuevaPassword, rolSolicitante } = req.body;

  if (rolSolicitante !== "ADMINISTRADOR") {
    return res.status(403).json({ error: "No tienes permisos para cambiar contraseñas" });
  }

  if (!nuevaPassword || nuevaPassword.length < 4) {
    return res.status(400).json({ error: "La contraseña debe tener al menos 4 caracteres" });
  }

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

    const hashedPassword = await bcrypt.hash(nuevaPassword, 10);

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
// ENDPOINTS ADICIONALES (SIN CAMBIOS)
// ============================================

app.get("/api/registros/excel", async (req, res) => {
  // Este endpoint ya no se usa porque los datos están en OneDrive
  res.status(501).json({ error: "Este endpoint ha sido migrado a OneDrive" });
});

app.get("/api/planificacion", async (req, res) => {
  // Este endpoint ya no se usa porque los datos están en OneDrive
  res.status(501).json({ error: "Este endpoint ha sido migrado a OneDrive" });
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

    const token = await getAccessToken();

    const odataBase = `https://api.businesscentral.dynamics.com/v2.0/${process.env.TENANT_ID}/${process.env.BC_ENVIRONMENT}/ODataV4/Company('${encodeURIComponent(nombreCompania)}')`;

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

    const companyId = companies[0].id;

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

app.get("/api/registros/mi-perfil", (req, res) => {
  try {
    console.log("=== RUTA MI-PERFIL LLAMADA ===");
    console.log("Query params:", req.query);

    const { nombre, rol } = req.query;

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
      `;
    }
    else {
      console.log("Rol no válido:", rol);
      return res.json([]);
    }

    console.log("SQL:", sql);
    console.log("Parámetros:", params);

    const safeParse = (value) => {
      if (!value) return [];
      if (typeof value === "object") return value;
      if (typeof value === "string") {
        try {
          return JSON.parse(value);
        } catch (e) {
          console.error("❌ Error parseando campo:", value.substring(0, 100));
          return [];
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
            return r;
          }
        });

        console.log(`Registros encontrados: ${registros.length}`);
        
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
  const { id } = req.params;
  const { nombre, rol } = req.body;

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
  
  const { rol, nombre } = datos;

  if (!["SUPERVISOR","LÍDER"].includes(rol)) {
    return res.status(403).json({ error: "No tienes permisos" });
  }

  const campos = [];
  const valores = [];
  
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

  Object.keys(camposPermitidos).forEach(campoFront => {
    if (datos[campoFront] !== undefined && datos[campoFront] !== null) {
      const campoDb = camposPermitidos[campoFront];
      if (!campos.some(c => c.includes(`${campoDb} =`))) {
        campos.push(`${campoDb} = ?`);
        valores.push(datos[campoFront]);
      }
    }
  });

  camposJSON.forEach(campoJSON => {
    if (datos[campoJSON] !== undefined && datos[campoJSON] !== null) {
      campos.push(`${campoJSON} = ?`);
      const valor = typeof datos[campoJSON] === 'string'
        ? datos[campoJSON]
        : JSON.stringify(datos[campoJSON]);
      valores.push(valor);
    }
  });

  if (campos.length === 0) {
    return res.status(400).json({ error: "No hay campos para actualizar" });
  }

  campos.push('verificado_por = ?');
  valores.push(nombre);
  
  campos.push('fecha_verificacion = datetime("now", "-5 hours")');
  
  if (datos.estado !== undefined) {
    campos.push('estado = ?');
    valores.push(datos.estado);
  } else {
    campos.push('estado = estado');
  }

  valores.push(id);

  const sql = `UPDATE registros SET ${campos.join(', ')} WHERE id = ?`;

  db.run(sql, valores, function(err) {
    if (err) {
      console.error("Error en UPDATE:", err);
      return res.status(500).json({ error: err.message });
    }

    console.log(`Registro ${id} actualizado. Filas afectadas: ${this.changes}`);

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

app.get("/api/test/site-id", async (req, res) => {
  try {
    const { getGraphAccessToken } = require('./services/erpAuth');
    const token = await getGraphAccessToken();
    
    // Buscar el sitio que contiene tus archivos
    const url = `https://graph.microsoft.com/v1.0/sites?search=*`;
    
    const response = await axios.get(url, {
      headers: {
        'Authorization': `Bearer ${token}`
      }
    });
    
    res.json({
      success: true,
      sites: response.data.value.map(site => ({
        name: site.name,
        id: site.id,
        webUrl: site.webUrl
      }))
    });
  } catch (error) {
    console.error("Error obteniendo sites:", error.response?.data || error.message);
    res.status(500).json({
      error: "Error obteniendo sites",
      detalle: error.response?.data || error.message
    });
  }
});

/* =========================
   START LOCAL SERVER
========================= */
const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => {
    console.log(`🚀 Backend corriendo en puerto ${PORT}`);
    console.log(`📁 Leyendo archivos desde OneDrive`);
});