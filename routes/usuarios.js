const express = require('express');
const router = express.Router();
const db = require('../db');

// Obtener todos los usuarios
router.get('/', async (req, res) => {
  try {
    const usuarios = await db.query(`
      SELECT id, username, nombre, cedula_identidad, rol, area, activo, 
             created_at, updated_at 
      FROM usuarios 
      ORDER BY id DESC
    `);
    res.json(usuarios);
  } catch (error) {
    console.error('Error al obtener usuarios:', error);
    res.status(500).json({ error: error.message });
  }
});

// Obtener un usuario por ID
router.get('/:id', async (req, res) => {
  try {
    const usuario = await db.getAsync(
      'SELECT id, username, nombre, cedula_identidad, rol, area, activo FROM usuarios WHERE id = ?',
      [req.params.id]
    );
    
    if (!usuario) {
      return res.status(404).json({ error: 'Usuario no encontrado' });
    }
    
    res.json(usuario);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Crear nuevo usuario
router.post('/', async (req, res) => {
  const { username, password, nombre, cedula_identidad, rol, area, activo } = req.body;
  
  // Validaciones básicas
  if (!username || !password || !nombre || !cedula_identidad || !rol) {
    return res.status(400).json({ 
      error: 'Faltan campos requeridos: username, password, nombre, cedula_identidad, rol' 
    });
  }

  try {
    // Verificar si el username o cedula_identidad ya existen
    const existe = await db.getAsync(
      'SELECT id FROM usuarios WHERE username = ? OR cedula_identidad = ?',
      [username, cedula_identidad]
    );
    
    if (existe) {
      return res.status(400).json({ 
        error: 'El username o cedula de identidad ya está registrado' 
      });
    }

    // En un entorno real, deberías hashear la contraseña con bcrypt
    // const hashedPassword = await bcrypt.hash(password, 10);
    
    const result = await db.runAsync(
      `INSERT INTO usuarios (username, password, nombre, cedula_identidad, rol, area, activo) 
       VALUES (?, ?, ?, ?, ?, ?, ?)`,
      [username, password, nombre, cedula_identidad, rol, area || null, activo !== undefined ? activo : 1]
    );

    const nuevoUsuario = await db.getAsync(
      'SELECT id, username, nombre, cedula_identidad, rol, area, activo FROM usuarios WHERE id = ?',
      [result.lastID]
    );

    res.status(201).json(nuevoUsuario);
  } catch (error) {
    console.error('Error al crear usuario:', error);
    res.status(500).json({ error: error.message });
  }
});

// Actualizar usuario
router.put('/:id', async (req, res) => {
  const { id } = req.params;
  const { username, password, nombre, cedula_identidad, rol, area, activo } = req.body;

  try {
    // Verificar si el usuario existe
    const usuario = await db.getAsync('SELECT id FROM usuarios WHERE id = ?', [id]);
    if (!usuario) {
      return res.status(404).json({ error: 'Usuario no encontrado' });
    }

    // Verificar si el username o cedula_identidad ya existen en OTRO usuario
    const existe = await db.getAsync(
      'SELECT id FROM usuarios WHERE (username = ? OR cedula_identidad = ?) AND id != ?',
      [username, cedula_identidad, id]
    );
    
    if (existe) {
      return res.status(400).json({ 
        error: 'El username o cedula de identidad ya está registrado por otro usuario' 
      });
    }

    let query = 'UPDATE usuarios SET username = ?, nombre = ?, cedula_identidad = ?, rol = ?, area = ?, activo = ?';
    let params = [username, nombre, cedula_identidad, rol, area, activo];

    // Si se proporciona nueva contraseña, actualizarla
    if (password) {
      query += ', password = ?';
      params.push(password);
    }

    query += ' WHERE id = ?';
    params.push(id);

    await db.runAsync(query, params);

    const usuarioActualizado = await db.getAsync(
      'SELECT id, username, nombre, cedula_identidad, rol, area, activo FROM usuarios WHERE id = ?',
      [id]
    );

    res.json(usuarioActualizado);
  } catch (error) {
    console.error('Error al actualizar usuario:', error);
    res.status(500).json({ error: error.message });
  }
});

// Activar/Desactivar usuario
router.patch('/:id/toggle-activo', async (req, res) => {
  const { id } = req.params;

  try {
    const usuario = await db.getAsync('SELECT activo FROM usuarios WHERE id = ?', [id]);
    
    if (!usuario) {
      return res.status(404).json({ error: 'Usuario no encontrado' });
    }

    const nuevoEstado = usuario.activo ? 0 : 1;
    
    await db.runAsync(
      'UPDATE usuarios SET activo = ? WHERE id = ?',
      [nuevoEstado, id]
    );

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
router.delete('/:id', async (req, res) => {
  const { id } = req.params;

  try {
    const usuario = await db.getAsync('SELECT id FROM usuarios WHERE id = ?', [id]);
    
    if (!usuario) {
      return res.status(404).json({ error: 'Usuario no encontrado' });
    }

    await db.runAsync('DELETE FROM usuarios WHERE id = ?', [id]);

    res.json({ 
      message: 'Usuario eliminado correctamente',
      id: parseInt(id)
    });
  } catch (error) {
    console.error('Error al eliminar usuario:', error);
    res.status(500).json({ error: error.message });
  }
});

module.exports = router;