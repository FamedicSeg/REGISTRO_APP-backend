const sqlite3 = require('sqlite3').verbose();
const path = require('path');
const fs = require('fs');
const bcrypt = require('bcrypt');

// Ruta del archivo de base de datos
const dbPath =
  process.env.NODE_ENV === 'production'
    ? '/data/registro.db' // Ruta para producción
    : path.resolve(__dirname, 'registro.db');

// Ensure data directory exists before SQLite tries to create the database file
const dataDir = path.dirname(dbPath);
if (!fs.existsSync(dataDir)) {
  try {
    fs.mkdirSync(dataDir, { recursive: true });
    console.log(`Created directory: ${dataDir}`);
  } catch (mkdirErr) {
    console.error(`Error creating directory ${dataDir}:`, mkdirErr.message);
  }
}

// Crear o conectar a la base de datos
const db = new sqlite3.Database(dbPath, (err) => {
  if (err) {
    console.error("Error al conectar con SQLite:", err.message);
  } else {
    console.log("Base de datos SQLite conectada");
  }
});

// Función para insertar usuarios de ejemplo
function insertarUsuariosEjemplo() {
  console.log("Insertando usuarios de ejemplo...");
  
  const usuariosEjemplo = [

    // ADMINISTRADOR
    ["administrador_tics", "1234", "ADMINISTRADOR", "Administrador TICS", "1718195827", "Técnico de TICS"],

    // JEFE DE PRODUCCIÓN
    ["enma_morales", "1234", "JEFE DE PRODUCCIÓN", "Morales Collaguazo Enma Nelly", "1716045461", "Gerencia de Producción"],
    
    // ANALISTA
    ["andrea_escobar", "1234", "ANALISTA DE PRODUCCIÓN", "Escobar Razo Andrea Vanessa", "1725489320", "Control de Calidad"],
    
    // SUPERVISORES
    ["nancy_de_la_cruz4", "7415", "SUPERVISOR", "De La Cruz Paguay Nancy Lucía", "1713319927", "Supervisión"],
    ["silvia_llumiquinga2", "1234", "SUPERVISOR", "Llumiquinga Analuisa Silvia Dolores", "1714267885", "Supervisión"],
    ["lucia_ango2", "1234", "SUPERVISOR", "Ango Chuquimarca Lucia Otilia", "1717266918", "Supervisión"],
    
    // LÍDERES
    ["john_proanio", "9631", "LÍDER", "Proaño Caiza John Alexander", "1756247712", "Producción"],
    ["nancy_barahona", "9631", "LÍDER", "Barahona Pillajo Nancy Lorena", "1720250255", "Producción"],
    ["elba_chuquimarca", "1234", "LÍDER", "Chuquimarca Chasipanta Elba Maritza", "1719999532", "Producción"],
    ["yessena_suntaxi", "1234", "LÍDER", "Suntaxi Criollo Yesseña Alexandra", "1722415005", "Producción"],
    ["vanessa_chuquimarca", "1234", "LÍDER", "Chuquimarca Chuquimarca Vanessa Michell", "1727290213", "Producción"],
    ["sandra_chiliquinga", "1234", "LÍDER", "Chiliquinga Cuichan Sandra Rocio", "1716871957", "Producción"],
    ["katherine_fernandez", "1234", "LÍDER", "Fernandez Zurita Katherine Belen", "1721990818", "Producción"],
    ["rosio_pilataxi", "1234", "LÍDER", "Pilataxi Monta Rocio Elizabeth", "1720130523", "Producción"],
    ["ana_pincay", "1234", "LÍDER", "Pincay Ruiz Ana Maria", "1716087463", "Producción"],
    ["sandra_criollo", "1234", "LÍDER", "Criollo Logacho Sandra Jimena", "1718194630", "Producción"],
    ["lucia_guiz", "1234", "LÍDER", "Guiz Guiz Lucia", "0400793741", "Producción"],
    ["sofia_gualotunia", "1234", "LÍDER", "Gualotuña Gualotuña Sofia Maria", "1719228841", "Producción"],
    ["alejandro_perugache","1234","LÍDER","Perugache Quimbiurco Manuel Alejandro","1721084257","Producción"],
    ["david_zapata","1234","LÍDER","Zapata Moina David Marcelo","1723643191","Producción"],
    ["priscila_alvarez","1234","LÍDER","Alvarez Morales Priscila Vanessa","1727586479","Producción"],
    ["patricia_cabascango","1234","LÍDER","Cabascango Quishpe Patricia Lorena","1716330145","Producción"],
    ["karla_pachucho","1234","LÍDER","Pachucho Pachacama Karla Soledad","1721960092","Producción"],
    ["grace_almachi","1234","LÍDER","Almachi Toapanta Grace Melany","1729058691","Producción"]
  ];

  const stmt = db.prepare(`
    INSERT INTO usuarios (username, password, rol, nombre, cedula_identidad, area, activo, primer_login) 
    VALUES (?, ?, ?, ?, ?, ?, 1, 1)
  `);

  let pendientes = usuariosEjemplo.length;
  
  usuariosEjemplo.forEach(user => {
    bcrypt.hash(user[1], 10).then(hash => {
      stmt.run([user[0], hash, user[2], user[3], user[4], user[5]], function(err) {
        pendientes--;
        if (err) {
          console.error(`Error al insertar usuario ${user[0]}:`, err.message);
        } else {
          console.log(`Usuario ${user[0]} creado con ID: ${this.lastID} (primer login = 1)`);
        }
        
        if (pendientes === 0) {
          stmt.finalize();
          console.log("Todos los usuarios de ejemplo insertados correctamente");
        }
      });
    });
  });
}

// Crear tablas si no existen
db.serialize(() => {

  // Tabla principal de registros (CON TODOS LOS CAMPOS)
  db.run(`
    CREATE TABLE IF NOT EXISTS registros ( 
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      fecha TEXT,
      op TEXT,
      turno TEXT,
      area TEXT,
      modulo TEXT,
      responsable TEXT,
      supervisor TEXT,
      personal_asignado TEXT,
      personal_otro TEXT,
      personal_presente TEXT,
      codigo_producto TEXT,
      descripcion TEXT,
      tiempo_planificado INTEGER,
      hora_planificada TEXT,
      cantidad_planificada INTEGER,
      lotePrimario TEXT,
      loteSecundario TEXT,
      loteUnido TEXT,
      reposicion_no_conforme TEXT,
      cantidad_elaborado INTEGER,
      cantidad_proceso INTEGER,
      cantidad_merma TEXT,
      fecha_final_producto TEXT,
      hora_inicio TEXT,
      hora_fin TEXT,
      destino TEXT,
      cliente TEXT,
      esteril TEXT,
      leyenda TEXT,
      leyenda_si TEXT,
      leyenda_otra TEXT,
      talla TEXT,
      detalles_actividades TEXT,
      cantidad_planificada_detalles INTEGER,
      cantidad_elaborada_detalles INTEGER,
      planificada_total_por_detalle TEXT,  
      elaborada_total_por_detalle TEXT,      
      etiquetas TEXT,
      observaciones TEXT,
      estado TEXT DEFAULT 'pendiente_SUPERVISOR',
      motivo_rechazo TEXT,                    
      rechazado_por TEXT,                     
      fecha_rechazo TEXT,                     
      creado_por TEXT,
      fecha_creacion TEXT,
      verificado_por TEXT,
      fecha_verificacion TEXT,
      aprobado_por TEXT,
      fecha_aprobacion TEXT,
      insumos TEXT,
      integrantes TEXT,
      maquinarias TEXT,
      actividades_por_integrante TEXT,
      actividades_con_horas TEXT
    )
  `);

  // Tabla de insumos del registro
  db.run(`
    CREATE TABLE IF NOT EXISTS registro_insumos (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      registro_id INTEGER,
      codigo_insumo TEXT,
      descripcion_insumo TEXT,
      cantidad_insumo INTEGER,
      lote_insumo TEXT,
      entrega TEXT,
      recepcion TEXT,
      FOREIGN KEY (registro_id) REFERENCES registros(id)
    )
  `);

  // Tabla de integrantes del registro
  db.run(`
    CREATE TABLE IF NOT EXISTS registro_integrantes (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      registro_id INTEGER,
      nombre TEXT,
      cargo TEXT,
      cargoOtro TEXT,
      FOREIGN KEY (registro_id) REFERENCES registros(id)
    )
  `);
  
  // ============================================
  // TABLA DE USUARIOS CON PRIMER LOGIN
  // ============================================
  
  // 1. Verificar si la tabla usuarios existe
  db.get("SELECT name FROM sqlite_master WHERE type='table' AND name='usuarios'", (err, row) => {
    if (err) {
      console.error("Error al verificar tabla usuarios:", err);
      return;
    }

    if (!row) {
      // La tabla no existe, crearla directamente
      console.log("Creando tabla usuarios por primera vez...");
      db.run(`
        CREATE TABLE usuarios (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          username TEXT UNIQUE NOT NULL,
          password TEXT NOT NULL,
          rol TEXT NOT NULL,
          nombre TEXT NOT NULL,
          cedula_identidad TEXT UNIQUE,
          area TEXT,
          activo INTEGER DEFAULT 1,
          primer_login INTEGER DEFAULT 1,
          created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
          updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )
      `, (err) => {
        if (err) {
          console.error("Error al crear tabla usuarios:", err);
        } else {
          console.log("Tabla usuarios creada correctamente");
          
          // Crear trigger
          db.run(`
            CREATE TRIGGER IF NOT EXISTS update_usuarios_timestamp 
            AFTER UPDATE ON usuarios
            BEGIN
              UPDATE usuarios SET updated_at = CURRENT_TIMESTAMP WHERE id = NEW.id;
            END;
          `);
          
          // Insertar usuarios de ejemplo
          insertarUsuariosEjemplo();
        }
      });
    } else {
      // La tabla existe, verificar si tiene la columna primer_login
      db.all("PRAGMA table_info(usuarios)", (err, columns) => {
        if (err) {
          console.error("Error al obtener columnas:", err);
          return;
        }
        
        const columnas = columns.map(c => c.name);
        
        // Verificar si falta primer_login
        if (!columnas.includes('primer_login')) {
          console.log("Agregando columna primer_login a tabla existente...");
          
          // Primero agregamos la columna sin DEFAULT
          db.run(`ALTER TABLE usuarios ADD COLUMN primer_login INTEGER`, (err) => {
            if (err) {
              console.error("Error al agregar primer_login:", err);
            } else {
              // Luego actualizamos los registros existentes con valor por defecto
              db.run(`UPDATE usuarios SET primer_login = 1 WHERE primer_login IS NULL`, (err) => {
                if (err) {
                  console.error("Error al actualizar primer_login:", err);
                } else {
                  console.log("Columna primer_login agregada y actualizada");
                }
              });
            }
          });
        }
        
        // Verificar si falta cedula_identidad
        if (!columnas.includes('cedula_identidad')) {
          db.run(`ALTER TABLE usuarios ADD COLUMN cedula_identidad TEXT UNIQUE`, (err) => {
            if (err) {
              console.error("Error al agregar cedula_identidad:", err);
            } else {
              console.log("Columna cedula_identidad agregada");
            }
          });
        }
        
        // Verificar si falta area
        if (!columnas.includes('area')) {
          db.run(`ALTER TABLE usuarios ADD COLUMN area TEXT`, (err) => {
            if (err) {
              console.error("Error al agregar area:", err);
            } else {
              console.log("Columna area agregada");
            }
          });
        }
        
        // Verificar si falta created_at
        if (!columnas.includes('created_at')) {
          db.run(`ALTER TABLE usuarios ADD COLUMN created_at DATETIME DEFAULT CURRENT_TIMESTAMP`, (err) => {
            if (err) {
              console.error("Error al agregar created_at:", err);
            } else {
              console.log("Columna created_at agregada");
            }
          });
        }
        
        // Verificar si falta updated_at
        if (!columnas.includes('updated_at')) {
          db.run(`ALTER TABLE usuarios ADD COLUMN updated_at DATETIME`, (err) => {
            if (err) {
              console.error("Error al agregar updated_at:", err);
            } else {
              console.log("Columna updated_at agregada");
              
              // Crear trigger para updated_at
              db.run(`
                CREATE TRIGGER IF NOT EXISTS update_usuarios_timestamp 
                AFTER UPDATE ON usuarios
                BEGIN
                  UPDATE usuarios SET updated_at = CURRENT_TIMESTAMP WHERE id = NEW.id;
                END;
              `, (err) => {
                if (err) {
                  console.error("Error al crear trigger:", err);
                } else {
                  console.log("Trigger de updated_at creado");
                }
              });
            }
          });
        } else {
          // Asegurar que el trigger existe
          db.run(`
            CREATE TRIGGER IF NOT EXISTS update_usuarios_timestamp 
            AFTER UPDATE ON usuarios
            BEGIN
              UPDATE usuarios SET updated_at = CURRENT_TIMESTAMP WHERE id = NEW.id;
            END;
          `, (err) => {
            if (err) {
              console.error("Error al crear trigger:", err);
            } else {
              console.log("Trigger de updated_at verificado");
            }
          });
        }
      });
    }
  });
  
  // ============================================
  // VERIFICAR Y AGREGAR COLUMNAS FALTANTES EN LA TABLA registros
  // ============================================
  
  db.all("PRAGMA table_info(registros)", (err, columns) => {
    if (err) {
      console.error("Error al verificar columnas de registros:", err);
      return;
    }
    
    const columnas = columns.map(c => c.name);
    console.log("Columnas actuales en registros:", columnas);
    
    // Agregar hora_planificada si no existe
    if (!columnas.includes('hora_planificada')) {
      console.log("Agregando columna hora_planificada...");
      db.run(`ALTER TABLE registros ADD COLUMN hora_planificada TEXT`, (err) => {
        if (err) {
          console.error("Error al agregar hora_planificada:", err);
        } else {
          console.log("Columna hora_planificada agregada");
        }
      });
    }
    
    // Agregar reposicion_no_conforme si no existe
    if (!columnas.includes('reposicion_no_conforme')) {
      console.log("Agregando columna reposicion_no_conforme...");
      db.run(`ALTER TABLE registros ADD COLUMN reposicion_no_conforme TEXT`, (err) => {
        if (err) {
          console.error("Error al agregar reposicion_no_conforme:", err);
        } else {
          console.log("Columna reposicion_no_conforme agregada");
        }
      });
    }
    
    // Agregar actividades_con_horas si no existe
    if (!columnas.includes('actividades_con_horas')) {
      console.log("Agregando columna actividades_con_horas...");
      db.run(`ALTER TABLE registros ADD COLUMN actividades_con_horas TEXT`, (err) => {
        if (err) {
          console.error("Error al agregar actividades_con_horas:", err);
        } else {
          console.log("Columna actividades_con_horas agregada");
        }
      });
    }

    // Agregar lotePrimario si no existe
    if (!columnas.includes('lotePrimario')) {
      console.log("Agregando columna lotePrimario...");
      db.run(`ALTER TABLE registros ADD COLUMN lotePrimario TEXT`, (err) => {
        if (err) {
          console.error("Error al agregar lotePrimario:", err);
        } else {
          console.log("Columna lotePrimario agregada");
        }
      });
    }

    // Agregar loteSecundario si no existe
    if (!columnas.includes('loteSecundario')) {
      console.log("Agregando columna loteSecundario...");
      db.run(`ALTER TABLE registros ADD COLUMN loteSecundario TEXT`, (err) => {
        if (err) {
          console.error("Error al agregar loteSecundario:", err);
        } else {
          console.log("Columna loteSecundario agregada");
        }
      });
    }

    // Agregar loteUnido si no existe
    if (!columnas.includes('loteUnido')) {
      console.log("Agregando columna loteUnido...");
      db.run(`ALTER TABLE registros ADD COLUMN loteUnido TEXT`, (err) => {
        if (err) {
          console.error("Error al agregar loteUnido:", err);
        } else {
          console.log("Columna loteUnido agregada");
        }
      });
    }

    // Agregar maquinarias si no existe
    if (!columnas.includes('maquinarias')) {
      console.log("Agregando columna maquinarias...");
      db.run(`ALTER TABLE registros ADD COLUMN maquinarias TEXT`, (err) => {
        if (err) {
          console.error("Error al agregar maquinarias:", err);
        } else {
          console.log("Columna maquinarias agregada");
        }
      });
    }

    // Agregar integrantes si no existe
    if (!columnas.includes('integrantes')) {
      console.log("Agregando columna integrantes...");
      db.run(`ALTER TABLE registros ADD COLUMN integrantes TEXT`, (err) => {
        if (err) {
          console.error("Error al agregar integrantes:", err);
        } else {
          console.log("Columna integrantes agregada");
        }
      });
    }

    // Agregar insumos si no existe
    if (!columnas.includes('insumos')) {
      console.log("Agregando columna insumos...");
      db.run(`ALTER TABLE registros ADD COLUMN insumos TEXT`, (err) => {
        if (err) {
          console.error("Error al agregar insumos:", err);
        } else {
          console.log("Columna insumos agregada");
        }
      });
    }

    // Agregar etiquetas si no existe
    if (!columnas.includes('etiquetas')) {
      console.log("Agregando columna etiquetas...");
      db.run(`ALTER TABLE registros ADD COLUMN etiquetas TEXT`, (err) => {
        if (err) {
          console.error("Error al agregar etiquetas:", err);
        } else {
          console.log("Columna etiquetas agregada");
        }
      });
    }

    // Agregar actividades_por_integrante si no existe
    if (!columnas.includes('actividades_por_integrante')) {
      console.log("Agregando columna actividades_por_integrante...");
      db.run(`ALTER TABLE registros ADD COLUMN actividades_por_integrante TEXT`, (err) => {
        if (err) {
          console.error("Error al agregar actividades_por_integrante:", err);
        } else {
          console.log("Columna actividades_por_integrante agregada");
        }
      });
    }

    // Agregar actividades_con_horas si no existe
    if (!columnas.includes('actividades_con_horas')) {
      console.log("Agregando columna actividades_con_horas...");
      db.run(`ALTER TABLE registros ADD COLUMN actividades_con_horas TEXT`, (err) => {
        if (err) {
          console.error("Error al agregar actividades_con_horas:", err);
        } else {
          console.log("Columna actividades_con_horas agregada");
        }
      });
    }

    // Agregar lotePrincipal si no existe
    if (!columnas.includes('lotePrincipal')) {
      console.log("Agregando columna lotePrincipal...");
      db.run(`ALTER TABLE registros ADD COLUMN lotePrincipal TEXT`, (err) => {
        if (err) {
          console.error("Error al agregar lotePrincipal:", err);
        } else {
          console.log("Columna lotePrincipal agregada");
        }
      });
    }

    // Agregar loteSecundario si no existe
    if (!columnas.includes('loteSecundario')) {
      console.log("Agregando columna loteSecundario...");
      db.run(`ALTER TABLE registros ADD COLUMN loteSecundario TEXT`, (err) => {
        if (err) {
          console.error("Error al agregar loteSecundario:", err);
        } else {
          console.log("Columna loteSecundario agregada");
        }
      });
    }

    // Agregar loteUnido si no existe
    if (!columnas.includes('loteUnido')) {
      console.log("Agregando columna loteUnido...");
      db.run(`ALTER TABLE registros ADD COLUMN loteUnido TEXT`, (err) => {
        if (err) {
          console.error("Error al agregar loteUnido:", err);
        } else {
          console.log("Columna loteUnido agregada");
        }
      });
    }
  });
});

// Función helper para consultas con promesas
db.query = (sql, params = []) => {
  return new Promise((resolve, reject) => {
    db.all(sql, params, (err, rows) => {
      if (err) reject(err);
      else resolve(rows);
    });
  });
};

// Función helper para ejecutar una sola fila
db.getAsync = (sql, params = []) => {
  return new Promise((resolve, reject) => {
    db.get(sql, params, (err, row) => {
      if (err) reject(err);
      else resolve(row);
    });
  });
};

// Función helper para ejecutar comandos (INSERT, UPDATE, DELETE)
db.runAsync = (sql, params = []) => {
  return new Promise((resolve, reject) => {
    db.run(sql, params, function(err) {
      if (err) reject(err);
      else resolve({ lastID: this.lastID, changes: this.changes });
    });
  });
};

module.exports = db;