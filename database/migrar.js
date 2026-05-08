const sqlite3 = require('sqlite3').verbose();
const path = require('path');

const dbPath = path.resolve(__dirname, 'registros.db');
const db = new sqlite3.Database(dbPath);

db.serialize(() => {
  // Verificar si la columna ya existe
  db.all("PRAGMA table_info(registros)", (err, rows) => {
    if (err) {
      console.error("Error:", err);
      return;
    }
    
    const columnas = rows.map(r => r.name);
    console.log("Columnas actuales:", columnas);
    
    if (!columnas.includes('fecha_rechazo')) {
      console.log("Agregando fecha_rechazo...");
      
      db.run("ALTER TABLE registros ADD COLUMN fecha_rechazo TEXT", (err) => {
        if (err) {
          console.error("Error al agregar columna:", err);
        } else {
          console.log("✅ Columna agregada exitosamente");
        }
        
        // Verificar nuevamente
        db.all("PRAGMA table_info(registros)", (err, rows) => {
          console.log("Columnas después de la migración:", rows.map(r => r.name));
          db.close();
        });
      });
    } else {
      console.log("La columna ya existe, no es necesario migrar");
      db.close();
    }
  });
});