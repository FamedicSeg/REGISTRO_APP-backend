const sqlite3 = require('sqlite3').verbose();
const path = require('path');
const fs = require('fs');

const dbPath = path.resolve(__dirname, 'registros.db');

// 1. Cerrar cualquier conexión existente
console.log("🔄 Reiniciando base de datos...");

// 2. Eliminar el archivo si existe
if (fs.existsSync(dbPath)) {
  try {
    fs.unlinkSync(dbPath);
    console.log("✅ Archivo registros.db eliminado");
  } catch (err) {
    console.error("❌ Error al eliminar archivo:", err);
    process.exit(1);
  }
} else {
  console.log("ℹ️ El archivo registros.db no existe");
}

// 3. Esperar un momento y salir
console.log("✅ Listo. Reinicia el servidor para recrear la BD");
console.log("👉 Ejecuta: npm start");