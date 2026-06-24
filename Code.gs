// ═══════════════════════════════════════════════════════════════════
//  GASTRO APP — Backend Google Apps Script
//  Versión 1.0
//
//  HOJAS REQUERIDAS en el Spreadsheet:
//    • config        — parámetros globales (clave/valor)
//    • productos     — catálogo de productos
//    • usuarios      — usuarios registrados
//    • pedidos       — cabecera de cada pedido
//    • pedido_items  — líneas de cada pedido
//
//  SETUP INICIAL:
//    1. Crear un Google Spreadsheet nuevo.
//    2. Pegar este código en Apps Script (Extensions > Apps Script).
//    3. Ejecutar setupSheets() UNA SOLA VEZ para crear las hojas.
//    4. Ir a Implementar > Nueva implementación > Aplicación web.
//       - Ejecutar como: Yo
//       - Quién tiene acceso: Cualquier persona
//    5. Copiar la URL y pegarla en CONFIG.gasURL del index.html y admin.html
// ═══════════════════════════════════════════════════════════════════

// ── ID del Spreadsheet ───────────────────────────────────────────
// Completar con el ID del spreadsheet (está en la URL de Drive)
const SPREADSHEET_ID = '1aktVz2oJtttVXnw2tl3vYIRf1fOx_6zBKuBp8oAk5mk';

// ── Dominio autorizado (seguridad CORS) ─────────────────────────
// Poner el dominio de GitHub Pages. Ej: 'meteoro405.github.io'
// Dejar vacío '' para permitir cualquier origen (solo para pruebas)
const DOMINIO_AUTORIZADO = '';

// ── Token de admin ───────────────────────────────────────────────
// Cambiar por una cadena larga y aleatoria. NUNCA commitear al repo.
// Generá una en: https://www.uuidgenerator.net/
const ADMIN_TOKEN = '8af2a84b-e02b-43a6-b035-6e62142d3948';

// ── Rate limiting ────────────────────────────────────────────────
const MAX_PEDIDOS_POR_HORA = 3;   // por usuario (wsp)


// ═══════════════════════════════════════════════════════════════════
//  SETUP INICIAL — ejecutar una sola vez
// ═══════════════════════════════════════════════════════════════════
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  function getOrCreate(nombre, headers) {
    let hoja = ss.getSheetByName(nombre);
    if (!hoja) {
      hoja = ss.insertSheet(nombre);
      hoja.appendRow(headers);
      hoja.setFrozenRows(1);
      hoja.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#1A1208')
        .setFontColor('#F5EDD8');
    }
    return hoja;
  }

  getOrCreate('config', ['clave', 'valor', 'descripcion']);
  getOrCreate('productos', [
    'id', 'nombre', 'descripcion', 'precio', 'categoria', 'foto', 'emoji',
    'activo', 'soloHoy', 'fechaAlta', 'pedidosHasta',
    'diasSemana', 'horaDesde', 'horaHasta', 'orden', 'ts_creacion'
  ]);
  getOrCreate('usuarios', [
    'id', 'nombre', 'wsp', 'direccion', 'barrio', 'ref',
    'ts_registro', 'ts_ultima_vez', 'total_pedidos', 'total_gastado'
  ]);
  getOrCreate('pedidos', [
    'id', 'usuario_id', 'usuario_nombre', 'usuario_wsp',
    'direccion', 'barrio', 'ref', 'nota',
    'total', 'estado', 'ts_pedido'
  ]);
  getOrCreate('pedido_items', [
    'pedido_id', 'producto_id', 'producto_nombre', 'precio_unit', 'cantidad', 'subtotal'
  ]);

  // Config inicial
  const cfg = ss.getSheetByName('config');
  const cfgData = cfg.getDataRange().getValues();
  const claves = cfgData.slice(1).map(r => r[0]);
  const defaults = [
    ['modo_mantenimiento', 'false', '¿Mostrar modo mantenimiento? (true/false)'],
    ['mensaje_mantenimiento', 'Estamos realizando mejoras. Volvemos pronto.', 'Mensaje de mantenimiento'],
    ['pedidos_abiertos', 'true', '¿Aceptar pedidos? (true/false)'],
    ['mensaje_cerrado', 'Los pedidos están cerrados por el momento.', 'Mensaje cuando pedidos están cerrados'],
  ];
  defaults.forEach(([k, v, d]) => {
    if (!claves.includes(k)) cfg.appendRow([k, v, d]);
  });

  Logger.log('✅ Hojas creadas/verificadas correctamente.');
}


// ═══════════════════════════════════════════════════════════════════
//  PUNTO DE ENTRADA — doGet / doPost
// ═══════════════════════════════════════════════════════════════════
function doGet(e) {
  // Acepta tanto GET puro como GET con datos (para evitar CORS)
  const body = e.parameter && Object.keys(e.parameter).length > 1 ? parseParams(e.parameter) : null;
  return handleRequest(e, body);
}

function doPost(e) {
  let body = null;
  try { body = JSON.parse(e.postData.contents); } catch {}
  if (!body) body = parseParams(e.parameter);
  return handleRequest(e, body);
}

function parseParams(params) {
  // Convierte parámetros de URL en objeto, parseando JSON donde corresponde
  const obj = {};
  Object.entries(params).forEach(([k, v]) => {
    try { obj[k] = JSON.parse(v); } catch { obj[k] = v; }
  });
  return obj;
}

function handleRequest(e, body) {
  const action = (e.parameter.action || '').toLowerCase();
  const token  = e.parameter.token || (body && body.token) || '';

  // ── Rutas públicas ─────────────────────────────────────
  if (action === 'getproductos')    return resp(getProductos());
  if (action === 'getconfig')       return resp(getConfig());
  if (action === 'registrarpedido') return resp(registrarPedido(body));
  if (action === 'registrarusuario')return resp(registrarUsuario(body));

  // ── Rutas de admin (requieren token) ───────────────────
  if (!validarToken(token)) {
    return resp({ ok: false, error: 'No autorizado' }, 401);
  }

  if (action === 'adminlogin')         return resp(adminLogin(token));
  if (action === 'getadminproductos')  return resp(getAdminProductos());
  if (action === 'saveproducto')       return resp(saveProducto(body));
  if (action === 'deleteproducto')     return resp(deleteProducto(body));
  if (action === 'toggleproducto')     return resp(toggleProducto(body));
  if (action === 'getpedidos')         return resp(getPedidos(e.parameter));
  if (action === 'getusuarios')        return resp(getUsuarios());
  if (action === 'getstats')           return resp(getStats());
  if (action === 'exportusuarios')     return respCSV(exportUsuarios(), 'usuarios.csv');
  if (action === 'exportproductos')    return respCSV(exportProductos(), 'productos.csv');
  if (action === 'saveconfigval')      return resp(saveConfigVal(body));

  return resp({ ok: false, error: 'Acción desconocida' }, 400);
}

// ── Helpers de respuesta ────────────────────────────────
function resp(data, code) {
  const output = ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

function respCSV(csv, filename) {
  return ContentService
    .createTextOutput(csv)
    .setMimeType(ContentService.MimeType.CSV);
}

function validarToken(t) {
  return t === ADMIN_TOKEN;
}


// ═══════════════════════════════════════════════════════════════════
//  ACCESO A HOJAS
// ═══════════════════════════════════════════════════════════════════
function ss() { return SpreadsheetApp.openById(SPREADSHEET_ID); }
function hoja(nombre) { return ss().getSheetByName(nombre); }

function leerHoja(nombre) {
  const h = hoja(nombre);
  const data = h.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function generarId() {
  return Utilities.getUuid().replace(/-/g,'').slice(0,12);
}


// ═══════════════════════════════════════════════════════════════════
//  CONFIG
// ═══════════════════════════════════════════════════════════════════
function getConfig() {
  const rows = leerHoja('config');
  const cfg = {};
  rows.forEach(r => { cfg[r.clave] = r.valor; });
  return { ok: true, config: cfg };
}

function saveConfigVal(body) {
  const h = hoja('config');
  const data = h.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === body.clave) {
      h.getRange(i + 1, 2).setValue(body.valor);
      return { ok: true };
    }
  }
  h.appendRow([body.clave, body.valor, '']);
  return { ok: true };
}


// ═══════════════════════════════════════════════════════════════════
//  PRODUCTOS
// ═══════════════════════════════════════════════════════════════════
function getProductos() {
  const cfg = getConfig().config;
  if (cfg.modo_mantenimiento === 'true') {
    return { ok: false, mantenimiento: true, mensaje: cfg.mensaje_mantenimiento };
  }
  const todos = leerHoja('productos');
  const activos = todos
    .filter(p => p.activo === true || p.activo === 'TRUE' || p.activo === 'true')
    .map(p => ({
      id:           p.id,
      nombre:       p.nombre,
      desc:         p.descripcion,
      precio:       Number(p.precio),
      categoria:    p.categoria,
      foto:         p.foto,
      emoji:        p.emoji,
      activo:       true,
      soloHoy:      p.soloHoy === true || p.soloHoy === 'TRUE',
      fechaAlta:    p.fechaAlta,
      pedidosHasta: p.pedidosHasta,
      diasSemana:   parseDias(p.diasSemana),
      horaDesde:    p.horaDesde,
      horaHasta:    p.horaHasta,
      orden:        Number(p.orden) || 99,
    }))
    .sort((a, b) => a.orden - b.orden || a.categoria.localeCompare(b.categoria));
  return { ok: true, productos: activos };
}

function getAdminProductos() {
  const todos = leerHoja('productos');
  return {
    ok: true,
    productos: todos.map(p => ({
      ...p,
      precio:    Number(p.precio),
      activo:    p.activo === true || p.activo === 'TRUE' || p.activo === 'true',
      soloHoy:   p.soloHoy === true || p.soloHoy === 'TRUE',
      diasSemana: parseDias(p.diasSemana),
      orden:     Number(p.orden) || 99,
    }))
  };
}

function parseDias(val) {
  if (!val) return [];
  if (Array.isArray(val)) return val.map(Number);
  try { return JSON.parse(val); } catch { return []; }
}

function saveProducto(body) {
  const h = hoja('productos');
  const data = h.getDataRange().getValues();
  const headers = data[0];

  const ahora = new Date().toISOString();
  const diasStr = JSON.stringify(body.diasSemana || []);

  if (body.id) {
    // Actualizar existente
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === body.id) {
        const row = [
          body.id, body.nombre, body.descripcion, body.precio,
          body.categoria, body.foto || '', body.emoji || '🍽️',
          body.activo ? 'TRUE' : 'FALSE',
          body.soloHoy ? 'TRUE' : 'FALSE',
          body.fechaAlta || '',
          body.pedidosHasta || '',
          diasStr,
          body.horaDesde || '',
          body.horaHasta || '',
          body.orden || 99,
          data[i][15]  // mantener ts_creacion
        ];
        h.getRange(i + 1, 1, 1, row.length).setValues([row]);
        return { ok: true, id: body.id };
      }
    }
  }

  // Nuevo producto
  const id = generarId();
  h.appendRow([
    id, body.nombre, body.descripcion, body.precio,
    body.categoria, body.foto || '', body.emoji || '🍽️',
    body.activo !== false ? 'TRUE' : 'FALSE',
    body.soloHoy ? 'TRUE' : 'FALSE',
    body.fechaAlta || '',
    body.pedidosHasta || '',
    diasStr,
    body.horaDesde || '',
    body.horaHasta || '',
    body.orden || 99,
    ahora
  ]);
  return { ok: true, id };
}

function deleteProducto(body) {
  const h = hoja('productos');
  const data = h.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === body.id) {
      h.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Producto no encontrado' };
}

function toggleProducto(body) {
  const h = hoja('productos');
  const data = h.getDataRange().getValues();
  const headers = data[0];
  const colActivo = headers.indexOf('activo') + 1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === body.id) {
      const actual = data[i][colActivo - 1];
      const nuevo = (actual === 'TRUE' || actual === true) ? 'FALSE' : 'TRUE';
      h.getRange(i + 1, colActivo).setValue(nuevo);
      return { ok: true, activo: nuevo === 'TRUE' };
    }
  }
  return { ok: false, error: 'No encontrado' };
}


// ═══════════════════════════════════════════════════════════════════
//  USUARIOS
// ═══════════════════════════════════════════════════════════════════
function registrarUsuario(body) {
  if (!body || !body.wsp || !body.nombre) return { ok: false, error: 'Datos incompletos' };

  const h = hoja('usuarios');
  const data = h.getDataRange().getValues();
  const ahora = new Date().toISOString();

  // Buscar por WhatsApp
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] == body.wsp) {
      // Actualizar datos
      h.getRange(i + 1, 2).setValue(body.nombre);
      h.getRange(i + 1, 4).setValue(body.direccion || '');
      h.getRange(i + 1, 5).setValue(body.barrio || '');
      h.getRange(i + 1, 6).setValue(body.ref || '');
      h.getRange(i + 1, 8).setValue(ahora);
      return { ok: true, id: data[i][0], nuevo: false };
    }
  }

  // Nuevo usuario
  const id = generarId();
  h.appendRow([id, body.nombre, body.wsp, body.direccion || '', body.barrio || '', body.ref || '', ahora, ahora, 0, 0]);
  return { ok: true, id, nuevo: true };
}

function getUsuarios() {
  return { ok: true, usuarios: leerHoja('usuarios') };
}

function buscarOCrearUsuario(body) {
  const result = registrarUsuario(body);
  return result.id;
}


// ═══════════════════════════════════════════════════════════════════
//  PEDIDOS
// ═══════════════════════════════════════════════════════════════════
function registrarPedido(body) {
  if (!body || !body.usuario || !body.carrito || !body.carrito.length) {
    return { ok: false, error: 'Datos incompletos' };
  }

  // Rate limiting
  if (excedeLimite(body.usuario.wsp)) {
    return { ok: false, error: 'Demasiados pedidos en poco tiempo. Esperá un momento.' };
  }

  const cfg = getConfig().config;
  if (cfg.pedidos_abiertos === 'false') {
    return { ok: false, error: cfg.mensaje_cerrado || 'Pedidos cerrados.' };
  }

  // Registrar/actualizar usuario
  const uid = buscarOCrearUsuario(body.usuario);

  // Crear pedido
  const pid = generarId();
  const ahora = new Date().toISOString();
  const total = body.total || body.carrito.reduce((s, i) => s + i.precio * i.cantidad, 0);

  hoja('pedidos').appendRow([
    pid, uid,
    body.usuario.nombre, body.usuario.wsp,
    body.usuario.direccion, body.usuario.barrio, body.usuario.ref || '',
    body.nota || '',
    total, 'recibido', ahora
  ]);

  // Items del pedido
  const hItems = hoja('pedido_items');
  body.carrito.forEach(item => {
    hItems.appendRow([
      pid, item.id, item.nombre,
      item.precio, item.cantidad,
      item.precio * item.cantidad
    ]);
  });

  // Actualizar stats del usuario
  actualizarStatsUsuario(uid, total);

  return { ok: true, pedido_id: pid };
}

function actualizarStatsUsuario(uid, total) {
  try {
    const h = hoja('usuarios');
    const data = h.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === uid) {
        const totalPedidos = (Number(data[i][8]) || 0) + 1;
        const totalGastado = (Number(data[i][9]) || 0) + total;
        h.getRange(i + 1, 9).setValue(totalPedidos);
        h.getRange(i + 1, 10).setValue(totalGastado);
        break;
      }
    }
  } catch {}
}

function excedeLimite(wsp) {
  try {
    const pedidos = leerHoja('pedidos');
    const hace1h = new Date(Date.now() - 3600000).toISOString();
    const recientes = pedidos.filter(p =>
      p.usuario_wsp == wsp && p.ts_pedido > hace1h
    );
    return recientes.length >= MAX_PEDIDOS_POR_HORA;
  } catch { return false; }
}

function getPedidos(params) {
  const todos = leerHoja('pedidos');
  const items = leerHoja('pedido_items');

  // Filtro por fecha
  let filtrados = todos;
  if (params.desde) filtrados = filtrados.filter(p => p.ts_pedido >= params.desde);
  if (params.hasta) filtrados = filtrados.filter(p => p.ts_pedido <= params.hasta + 'T23:59:59');

  // Adjuntar items a cada pedido
  const pedidosConItems = filtrados
    .sort((a, b) => b.ts_pedido.localeCompare(a.ts_pedido))
    .map(p => ({
      ...p,
      items: items.filter(i => i.pedido_id === p.id)
    }));

  return { ok: true, pedidos: pedidosConItems };
}


// ═══════════════════════════════════════════════════════════════════
//  ESTADÍSTICAS
// ═══════════════════════════════════════════════════════════════════
function getStats() {
  const pedidos = leerHoja('pedidos');
  const items   = leerHoja('pedido_items');
  const usuarios = leerHoja('usuarios');

  // Total general
  const totalPedidos  = pedidos.length;
  const totalFacturado = pedidos.reduce((s, p) => s + (Number(p.total) || 0), 0);

  // Pedidos últimos 7 días
  const hace7d = new Date(Date.now() - 7 * 86400000).toISOString();
  const ult7 = pedidos.filter(p => p.ts_pedido >= hace7d);

  // Pedidos últimos 30 días
  const hace30d = new Date(Date.now() - 30 * 86400000).toISOString();
  const ult30 = pedidos.filter(p => p.ts_pedido >= hace30d);

  // Ranking de productos
  const conteoProductos = {};
  items.forEach(it => {
    const k = it.producto_nombre;
    if (!conteoProductos[k]) conteoProductos[k] = { nombre: k, cantidad: 0, facturado: 0 };
    conteoProductos[k].cantidad  += Number(it.cantidad) || 0;
    conteoProductos[k].facturado += Number(it.subtotal) || 0;
  });
  const rankingProductos = Object.values(conteoProductos)
    .sort((a, b) => b.cantidad - a.cantidad)
    .slice(0, 10);

  // Pedidos por día (últimos 30 días)
  const porDia = {};
  ult30.forEach(p => {
    const dia = p.ts_pedido.slice(0, 10);
    porDia[dia] = (porDia[dia] || 0) + 1;
  });

  // Top usuarios
  const topUsuarios = usuarios
    .filter(u => Number(u.total_pedidos) > 0)
    .sort((a, b) => Number(b.total_pedidos) - Number(a.total_pedidos))
    .slice(0, 10)
    .map(u => ({
      nombre:        u.nombre,
      wsp:           u.wsp,
      total_pedidos: Number(u.total_pedidos),
      total_gastado: Number(u.total_gastado)
    }));

  return {
    ok: true,
    stats: {
      totalPedidos,
      totalFacturado,
      totalUsuarios: usuarios.length,
      pedidosUlt7:   ult7.length,
      pedidosUlt30:  ult30.length,
      facturadoUlt30: ult30.reduce((s, p) => s + (Number(p.total) || 0), 0),
      rankingProductos,
      porDia,
      topUsuarios
    }
  };
}


// ═══════════════════════════════════════════════════════════════════
//  EXPORTACIONES CSV
// ═══════════════════════════════════════════════════════════════════
function exportUsuarios() {
  const usuarios = leerHoja('usuarios');
  const headers = ['Nombre', 'WhatsApp', 'Dirección', 'Barrio', 'Referencia', 'Registro', 'Última vez', 'Pedidos', 'Total gastado'];
  const rows = usuarios.map(u => [
    u.nombre, u.wsp, u.direccion, u.barrio, u.ref,
    u.ts_registro, u.ts_ultima_vez, u.total_pedidos, u.total_gastado
  ]);
  return [headers, ...rows].map(r => r.map(v => `"${String(v||'').replace(/"/g,'""')}"`).join(',')).join('\n');
}

function exportProductos() {
  const productos = leerHoja('productos');
  const headers = ['ID', 'Nombre', 'Descripción', 'Precio', 'Categoría', 'Foto', 'Emoji', 'Activo', 'Solo hoy', 'Pedidos hasta', 'Días semana', 'Hora desde', 'Hora hasta', 'Orden'];
  const rows = productos.map(p => [
    p.id, p.nombre, p.descripcion, p.precio, p.categoria,
    p.foto, p.emoji, p.activo, p.soloHoy, p.pedidosHasta,
    p.diasSemana, p.horaDesde, p.horaHasta, p.orden
  ]);
  return [headers, ...rows].map(r => r.map(v => `"${String(v||'').replace(/"/g,'""')}"`).join(',')).join('\n');
}


// ═══════════════════════════════════════════════════════════════════
//  ADMIN LOGIN (verificar token)
// ═══════════════════════════════════════════════════════════════════
function adminLogin(token) {
  return { ok: true, mensaje: 'Autenticado correctamente' };
}
