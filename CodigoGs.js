const USERS_SHEET_ID = '1-gT6lMWqsT1P6yJta5VBcHWTR7ukA78SQpjfIQsQ4PU'; // Hoja de usuarios
const EMPLOYEES_SHEET_ID = '1FZDShnJyXYe_o5SLruKTkHFer2MoHdknGIu3G7gaG2U'; // Hoja de empleados/turnos (Maestra, turnos_semanales, etc.)
// ID de la hoja de cálculo "Brigadas Doria"
const BRIGADAS_DORIA_SHEET_ID = '1-MkS9q08XjAWmoilDdEAYWBNk_L94JuDAILwAT_QfQw';
const BRIGADISTAS_SHEET_NAME = 'Brigada';

// Define the shift definitions based on the detailed analysis
const SHIFT_DEFINITIONS = {
  "T1": { "Lunes": 6.5, "Martes": 7.5, "Miercoles": 7.5, "Jueves": 7.5, "Viernes": 7.5, "Sabado": 7.5, "Domingo": 0, "total": 44 },
  "T2": { "Lunes": 7.5, "Martes": 7.5, "Miercoles": 7.5, "Jueves": 7.5, "Viernes": 7.5, "Sabado": 6.5, "Domingo": 0, "total": 44 },
  "T3": { "Lunes": 7.5, "Martes": 7.5, "Miercoles": 7.5, "Jueves": 7.5, "Viernes": 7.5, "Sabado": 0, "Domingo": 0, "total": 37.5 },
  "T1 EXTRA": { "Lunes": 7.5, "Martes": 7.5, "Miercoles": 7.5, "Jueves": 7.5, "Viernes": 7.5, "Sabado": 7.5, "Domingo": 8, "total": 53 },
  "T2 EXTRA": { "Lunes": 7.5, "Martes": 7.5, "Miercoles": 7.5, "Jueves": 7.5, "Viernes": 7.5, "Sabado": 8, "Domingo": 8, "total": 53.5 },
  "T3 EXTRA": { "Lunes": 7.5, "Martes": 7.5, "Miercoles": 7.5, "Jueves": 7.5, "Viernes": 7.5, "Sabado": 8, "Domingo": 8, "total": 53.5 }
};



function doGet() {
return HtmlService.createTemplateFromFile('index').evaluate();
}

function include(filename) {
return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


function obtenerInfoUsuario(email) {
  try {
    const spreadsheet = SpreadsheetApp.openById(USERS_SHEET_ID);
    let sheet = spreadsheet.getSheetByName('usuario');
    if (!sheet) sheet = spreadsheet.getSheets()[0];
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return {success: false, message: 'No hay usuarios registrados'};
    }
    
    const emailLimpio = email.toLowerCase().trim();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row.length < 1) continue;
      
      const emailSheet = String(row[0]).toLowerCase().trim();
      
      if (emailSheet === emailLimpio) {
        return {
          success: true,
          email: emailSheet,
          cargo: row[2] ? String(row[2]).trim() : '', // Columna 3 (índice 2) contiene el cargo
          row: i + 1
        };
      }
    }
    
    return {success: false, message: 'Usuario no encontrado'};
  } catch (error) {
    console.error('Error al obtener información del usuario:', error);
    return {success: false, message: 'Error interno del servidor'};
  }
}


function verificarCredenciales(correo, clave) {
  try {
    const spreadsheet = SpreadsheetApp.openById(USERS_SHEET_ID);
    let sheet = spreadsheet.getSheetByName('usuario');
    if (!sheet) sheet = spreadsheet.getSheets()[0];
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return {success: false, message: 'No hay usuarios registrados'};
    
    const correoLimpio = correo.toLowerCase().trim();
    const claveLimpia = clave.trim();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row.length < 2) continue;
      
      const emailSheet = String(row[0]).toLowerCase().trim();
      const passwordSheet = String(row[1]).trim();
      const cargoSheet = row[2] ? String(row[2]).trim() : '';
      
      if (emailSheet === correoLimpio && passwordSheet === claveLimpia) {
        return {
          success: true,
          message: 'Login exitoso',
          user: {
            email: emailSheet,
            cargo: cargoSheet,
            row: i + 1
          }
        };
      }
    }
    
    return {success: false, message: 'Credenciales incorrectas'};
  } catch (error) {
    console.error('Error al verificar credenciales:', error);
    return {success: false, message: 'Error interno del servidor'};
  }
}

function getUserRole(email) {
    try {
        const spreadsheet = SpreadsheetApp.openById(USERS_SHEET_ID);
        const sheet = spreadsheet.getActiveSheet();
        const data = sheet.getDataRange().getValues();
        
        // Buscar el usuario por email (primera columna según tu tabla)
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const userEmail = row[0] ? row[0].toString().trim().toLowerCase() : '';
            const userRole = row[2] ? row[2].toString().trim() : ''; // Columna "cargo"
            
            if (userEmail === email.toLowerCase()) {
                console.log(`Rol encontrado para ${email}: ${userRole}`);
                return userRole;
            }
        }
        
        console.log(`No se encontró rol para el email: ${email}`);
        return ''; // Sin rol encontrado
        
    } catch (error) {
        console.error('Error al obtener rol del usuario:', error);
        return '';
    }
}

function validateUserWithRole(email, password) {
    try {
        const spreadsheet = SpreadsheetApp.openById(USERS_SHEET_ID);
        const sheet = spreadsheet.getActiveSheet();
        const data = sheet.getDataRange().getValues();
        
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const userEmail = row[0] ? row[0].toString().trim().toLowerCase() : '';
            const userPassword = row[1] ? row[1].toString().trim() : '';
            const userRole = row[2] ? row[2].toString().trim() : '';
            
            if (userEmail === email.toLowerCase() && userPassword === password) {
                return {
                    success: true,
                    email: userEmail,
                    role: userRole,
                    message: 'Autenticación exitosa'
                };
            }
        }
        
        return {
            success: false,
            message: 'Credenciales inválidas'
        };
        
    } catch (error) {
        return {
            success: false,
            message: 'Error en la autenticación: ' + error.message
        };
    }
}

// Agregar al archivo .gs

function obtenerEmpleadosPorTurno(turno, area) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const empleados = [];
    
    // Obtener datos de las hojas de programación
    const hojas = ['Programacion_Semanal', 'Programacion_Rango', 'Programacion_Anual'];
    
    hojas.forEach(nombreHoja => {
      const hoja = ss.getSheetByName(nombreHoja);
      if (!hoja) return;
      
      const datos = hoja.getDataRange().getValues();
      const encabezados = datos[0];
      
      // Índices relevantes
      const idxDoc = encabezados.indexOf('Documento');
      const idxNombre = encabezados.indexOf('Nombre');
      const idxArea = encabezados.indexOf('Area');
      const idxTurno = nombreHoja === 'Programacion_Semanal' ? -1 : encabezados.indexOf('Turno');
      
      // Procesar datos según el tipo de programación
      for (let i = 1; i < datos.length; i++) {
        const fila = datos[i];
        
        // Verificar si corresponde al área solicitada
        if (area !== 'todos' && fila[idxArea] !== area) continue;
        
        // Verificar turno según el tipo de programación
        if (nombreHoja === 'Programacion_Semanal') {
          // Para programación semanal, revisar cada día
          const diasTurnos = encabezados.slice(4, 11); // Lunes a Domingo
          diasTurnos.forEach((dia, idx) => {
            if (fila[idx + 4] === turno) {
              empleados.push({
                documento: fila[idxDoc],
                nombre: fila[idxNombre],
                area: fila[idxArea],
                fecha: dia,
                tipoProgramacion: 'Semanal'
              });
            }
          });
        } else {
          // Para programación por rango y anual
          if (fila[idxTurno] === turno) {
            empleados.push({
              documento: fila[idxDoc],
              nombre: fila[idxNombre],
              area: fila[idxArea],
              fecha: nombreHoja === 'Programacion_Rango' ? 
                     `${fila[encabezados.indexOf('Fecha_Inicio')]} - ${fila[encabezados.indexOf('Fecha_Fin')]}` :
                     `Año ${fila[encabezados.indexOf('Año')]}`,
              tipoProgramacion: nombreHoja.replace('Programacion_', '')
            });
          }
        }
      }
    });
    
    return empleados;
    
  } catch (error) {
    throw new Error(`Error al obtener empleados por turno: ${error.message}`);
  }
}
/**
* Mapea el nombre de un área (data-option del frontend) a palabras clave
* que podrían encontrarse en la columna 'Cargo' de la hoja 'Maestra'.
* @param {string} area El nombre del área en kebab-case (data-option).
* @returns {string[]} Un array de palabras clave para buscar en el cargo.
*/

function quitarTildes(texto) {
  return texto.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function getKeywordsForArea(area) {
const mappings = {
  'materias-primas': ['materias primas', 'mat primas', 'almacén de materias primas', 'ayudante almacen mat. primas', 'analista almacén mat. primas', 'coordinador(a) almacen materias primas'],
  'molino': ['molino', 'operario de molino', 'coordinador/a turno molino', 'auxiliar molino', 'jefe molinero', 'jefe de turno molino'],
  'cedi': ['cedi', 'logística', 'distribución', 'centro de distribución', 'ayudante cedi', 'coordinador de distribucion', 'auxiliar de distribucion', 'operario de montacarga', 'conductor de tractomula', 'coordinador/a centro de distribución-cedi', 'operario de bascula', 'operario/a vehículos de carga', 'jefe de logistica', 'analista de inventario', 'lider de operaciones', 'líder de operaciones', 'operario(a) vehiculos de carga', 'lider de operaciones'],
  'empaque': ['empaque', 'empacador', 'empacadora', 'mecanico empaque', 'ayudante de empaque', 'jefe mecanico de empaque', 'coordinador/a i+d empaque'],
  'pastificio': ['pastificio', 'pasta', 'producción de pasta', 'jefe de turno pastificio', 'ayudante lavamoldes'],
  'calidad': ['calidad', 'control de calidad', 'analista de calidad', 'jefe/a de aseguramiento de calidad', 'metrologo', 'analista de laboratorio fisico-quimico', 'analista de seguridad de los alimentos', 'analista de microbiología', 'analista aseguramiento calidad', 'director de calidad integral investigación y desar'],
  'sst': ['sst', 'seguridad y salud en el trabajo', 'seguridad industrial', 'analista seguridad y salud en el trabajo', 'auxiliar seguridad y salud en el trabajo', 'líder de salud y seguridad', 'dir.dllo humano y seg. y salud en trabajo'],
  'info-manufactura': ['info manufactura', 'información manufactura', 'analista de manufactura', 'coordinador informacion manufactura', 'analista de informacion manufactura', 'mecanico empaque II', 'soporte adtvo. proceso productivo'],
  'almacen-general': [/*'almacén general', 'almacenista', 'bodega',*/ 'auxiliar de almacen', /*'analista de inventario', 'coordinador(a)0 almacén general'*/],
  'ingenieria-montaje': ['ingeniería y montaje', 'ingeniero de montaje', 'montaje', 'jefe/a de ingeniería', 'director(a) técnico', 'auxiliar admintivo ingenieria y montajes', 'mecanico sistema automático'],
  'investigacion-desarrollo': ['investigación y desarrollo', 'i+d', 'desarrollo de productos', 'analista de investigacion y desarrollo', 'coord. investigación y desarrollo'],
  'tecnico-electricista': ['técnico electricista', 'electricista', 'tecnico electricista ii'],
  'tecnico-lider': ['técnico líder', 'líder técnico', 'tecnico lider'],
  'soporte-adtvo': ['soporte adtvo', 'administrativo', 'asistente administrativo', 'soporte adtvo. proceso productivo', 'soporte procesos adtvos-presidencia', 'soporte procesos adtvos.-cabas'],
  'maquinista': ['maquinista', 'operador de máquina'],
  'lider-operaciones': ['líder de operaciones', 'jefe de operaciones', 'coordinador de operaciones'],
  'soporte-servicios-administrativos': ['soporte servicios administrativos', 'servicios administrativos', 'jefe servicios administrativos'],
  'ayudante-produccion': [
    'ayudante de producción',
    'ayudante produccion',
    'auxiliar de producción',
    'auxiliar produccion',
    'asistente de producción',
    'asistente produccion',
    'operario de producción', // Mantener específico
    'operario produccion'
    // Eliminado 'operario' genérico para evitar conflictos con montacarga
  ],
  'empacador': ['empacador', 'empacadora', 'empacador(a)'],
  'operario-montacarga': [
    'operario montacarga',
    'operario de montacarga',
    'montacarguista',
    'conductor montacarga',
    'operador montacarga',
    'operador de montacarga'
  ],
  'jefe-ingenieria': ['jefe de ingeniería', 'jefa de ingeniería', 'gerente de ingeniería', 'jefe/a de ingenierìa'],
  'coordinador-distribucion': ['info manufactura', 'información manufactura', 'analista de manufactura', 'coordinador informacion manufactura', 'analista de informacion manufactura', 'mecanico empaque II', 'soporte adtvo. proceso productivo'],
  'analista-control-gestion': ['analista control y gestión', 'control de gestión', 'analista de gestión', 'jefe/a de control gestion', 'analista de costos', 'jefe de costos', 'jefe(a) de planeacion estrategica y proyectos'],
  'produccion': ['director/a de manufactura', 'jefe/a de producción',  'ayudante de produccion','empacador', 'empacadora', 'ayudante de empaque', 'ayudante lavamoldes', 'jefe de turno pastificio', 'jefe mecanico de empaque', 'maquinista', 'mecanico empaque', 'mecanico empaque II', 'mecanico empaque ii', 'mecanico sistema automatico', 'MECÁNICO SISTEMA AUTOMÁTICO', 'profesional en formacion', 'jefe(a) de produccion'],
  'gestion-humana': ['gestión humana', 'gerente gestion humana', 'dir.dllo humano y seg. y salud en trabajo', 'coordinador(a) dllo humano', 'analista gestion humana', 'jefe/a relaciones laborales y diseño organizaciona', 'analista de calidad de vida'],
  'mercadeo': ['mercadeo', 'jefe/a de marca', 'gerente de mercadeo', 'auxiliar de mercadeo', 'analista de mercadeo', 'director(a) de marca doria', 'coordinador/a comunicación de marca'],
  'planeacion': ['planeación', 'director/a planeación estratégica y control gestió', 'jefe/a de planeación estratégica y proyectos', 'coordinador/a de planeacion', 'jefe planeación y abastecimiento'],
  'gestion-ambiental': ['gestión ambiental', 'analista de gestión ambiental', 'coordinador(a) de gestion ambiental'],
  'tecnologia': ['tecnología', 'desarrollador de software', 'analista desarrollador de tecnologías'],
 'mantenimiento': [
  'mantenimiento',
  'analista desarrollador de tecnologías',
  'auxiliar admintivo ingenieria y montajes',
  'auxiliar admintivo ingeniería y montajes',
  'director(a) técnico',
  'jefe/a de ingeniería',
  'jefe/a de ingenieria',
  'metrologo',
  'tecnico de mantenimiento',
  'técnico de mantenimiento',
  'tecnico de mantenimiento ii',
  'técnico de mantenimiento ii',
  'tecnico electricista',
  'técnico electricista',
  'tecnico electricista ii',
  'técnico electricista ii',
  'tecnico lider',
  'técnico líder',
  'jefe(a) de ingenieria',
  'JEFE(A) DE INGENIERIA'
],
  'gestion-comercial': ['gestión comercial', 'gerente/a gestión comercial', 'jefe de gestión comercial', 'jefe/a gestión comercial- canales alternativos', 'líder comercial tribiom', 'jefe/a gestión comercial- internacional y b2b'],
  'gestion-integral': ['gestión integral', 'coordinador(a) de gestión integral', 'analista de gestión integral'],
  'presidencia': ['presidente'],
};
// Normalizar el área de entrada para que coincida con las claves
const normalizedArea = area.toLowerCase().replace(/ /g, '-');
return mappings[normalizedArea] || [];
}

function obtenerEmpleados(filtro = {}) {
try {
  Logger.log('=== INICIANDO obtenerEmpleados ===');
  Logger.log('Filtro recibido:', filtro);
  const spreadsheet = SpreadsheetApp.openById(EMPLOYEES_SHEET_ID);
  // Buscar hoja "Maestra" (insensible a mayúsculas/minúsculas)
  const sheets = spreadsheet.getSheets();
  let sheet = null;
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().toLowerCase() === 'maestra') {
      sheet = sheets[i];
      break;
    }
  }
  if (!sheet) {
    const sheetNames = sheets.map(sh => sh.getName());
    Logger.log('No se encontró hoja "Maestra". Hojas disponibles:', sheetNames);
    return {error: true, message: 'No se encontró la hoja "Maestra"', availableSheets: sheetNames};
  }
  Logger.log('Hoja "Maestra" encontrada, obteniendo datos...');
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log('Hoja encontrada pero sin datos suficientes');
    return {error: true, message: 'La hoja no contiene datos suficientes'};
  }
  // Obtener encabezados
  const headers = data[0].map(h => h.toString().trim());
  Logger.log('Encabezados encontrados:', headers);
  // Buscar índices de columnas ESPECÍFICOS para tu hoja
  const docIndex = findColumnIndexExact(headers, ['03_DOCUMENTO_IDENTIDIDAD','documento_identidad','documento']);
  const nombreIndex = findColumnIndexExact(headers, ['02_EMPLEADO','empleado','nombre']);
  const cargoIndex = findColumnIndexExact(headers, ['12_NOMBRE_CARGO','nombre_cargo','cargo']);
  const areaIndex = findColumnIndexExact(headers, ['area','departamento','sector','division','zona','seccion']); // Columna de área si existe
  Logger.log('Índices encontrados:', {documento: docIndex, nombre: nombreIndex, cargo: cargoIndex, area: areaIndex});
  // Validar columnas obligatorias
  if (docIndex === -1 || nombreIndex === -1) {
    return {error: true, message: 'No se encontraron las columnas requeridas (Documento o Nombre)', headers: headers, expectedColumns: ['03_DOCUMENTO_IDENTIDIDAD', '02_EMPLEADO']};
  }
  // Procesar empleados
  const empleados = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    // Validar que la fila tenga las columnas mínimas requeridas
    if (row.length <= Math.max(docIndex, nombreIndex)) {
      Logger.log(`Fila ${i+1} omitida - no tiene suficientes columnas`);
      continue;
    }
    const documento = row[docIndex] ? row[docIndex].toString().trim() : '';
    const nombre = row[nombreIndex] ? row[nombreIndex].toString().trim() : '';
    const cargo = cargoIndex !== -1 && row[cargoIndex] ? row[cargoIndex].toString().trim() : '';
    // Si hay una columna de área, úsala; de lo contrario, déjala vacía para que el frontend la derive
    const areaEnHoja = areaIndex !== -1 && row[areaIndex] ? row[areaIndex].toString().trim() : '';
    
    Logger.log(`\n--- Procesando fila ${i+1} ---`);
    Logger.log(`  Documento: ${documento}, Nombre: ${nombre}, Cargo: "${cargo}", Área en hoja: "${areaEnHoja}"`);
    // Solo incluir si tiene documento y nombre válidos
    if (documento && nombre) {
      const empleado = {
        documento: documento,
        nombre: nombre,
        cargo: cargo,
        area: areaEnHoja, // Usar el área de la hoja si existe
        fila: i + 1
      };
      // Aplicar filtros
      let incluir = true;
      // Filtro por búsqueda general
      if (filtro.busqueda) {
        const busqueda = filtro.busqueda.toLowerCase();
        const textoCompleto = `${empleado.documento} ${empleado.nombre} ${empleado.cargo} ${empleado.area}`.toLowerCase();
        incluir = textoCompleto.includes(busqueda);
        Logger.log(`  Filtro búsqueda general "${busqueda}" en "${textoCompleto}": ${incluir}`);
      }
      // Filtro por tipo específico
      if (filtro.tipo && filtro.valor) {
        const valor = filtro.valor.toLowerCase();
        switch (filtro.tipo.toLowerCase()) {
          case 'documento':
            incluir = incluir && empleado.documento.toLowerCase().includes(valor);
            break;
          case 'nombre':
            incluir = incluir && empleado.nombre.toLowerCase().includes(valor);
            break;
          case 'cargo':
            incluir = incluir && empleado.cargo && empleado.cargo.toLowerCase().includes(valor);
            break;
        }
        Logger.log(`  Filtro por ${filtro.tipo} "${valor}": ${incluir}`);
      }
      // Filtro por área (data-option del frontend)
      if (filtro.area && filtro.area.toLowerCase() !== 'todos' && filtro.area.trim() !== '') {
        const areaFiltro = filtro.area.toLowerCase().trim(); // e.g., "coordinador-distribucion"
        const cargoEmpleadoLower = empleado.cargo.toLowerCase().trim(); // e.g., "coordinador de distribución"
        const areaEmpleadoLower = empleado.area.toLowerCase().trim(); // e.g., "logistica" or "cedi" or ""
       
         
        
        let areaMatchBySheetColumn = false;
        // Intento 1: Coincidencia por la columna 'Area' de la hoja (si existe y tiene valor)
        if (areaIndex !== -1 && areaEmpleadoLower) {
          // Comprobar si el valor del área de la hoja coincide directamente con el filtro del frontend
          // o si lo contiene (para casos como "Logística" vs "logistica")
          areaMatchBySheetColumn = areaEmpleadoLower === areaFiltro || areaEmpleadoLower.includes(areaFiltro);
          Logger.log(`  Intento 1: Coincidencia por columna 'Area' de la hoja ("${areaEmpleadoLower}" vs "${areaFiltro}"): ${areaMatchBySheetColumn}`);
        }
        let areaMatchByCargoKeywords = false;
        // Intento 2: Coincidencia por palabras clave en la columna 'Cargo'
        const keywords = getKeywordsForArea(areaFiltro).map(quitarTildes);
        areaMatchByCargoKeywords = keywords.some(keyword => cargoEmpleadoLower.includes(keyword));
        Logger.log(`  Intento 2: Coincidencia por palabras clave en 'Cargo' ("${cargoEmpleadoLower}" con keywords [${keywords.join(', ')}]): ${areaMatchByCargoKeywords}`);
        // El empleado se incluye si CUALQUIERA de los intentos de coincidencia de área es verdadero
        incluir = incluir && (areaMatchBySheetColumn || areaMatchByCargoKeywords);
        Logger.log(`  Resultado final del filtro por área: ${incluir}`);
      }
      if (incluir) {
        empleados.push(empleado);
        Logger.log(`  Empleado INCLUIDO: ${empleado.documento} - ${empleado.nombre}`);
      } else {
        Logger.log(`  Empleado EXCLUIDO por filtro: ${empleado.documento} - ${empleado.nombre}`);
      }
    } else {
      Logger.log(`Fila ${i+1} omitida - documento o nombre vacío: "${documento}" - "${nombre}"`);
    }
  }
  Logger.log(`\nProcesadas ${data.length-1} filas, ${empleados.length} empleados encontrados`);
  return empleados.map(empleado => ({
    documento: empleado.documento || '',
    nombre: empleado.nombre || '',
    cargo: empleado.cargo || '',
    area: empleado.area || ''
  }));
} catch (error) {
  console.error('Error en obtenerEmpleados:', error);
  return [];
}
}

// Función auxiliar para búsqueda exacta de columnas (mejorada para incluir múltiples posibles nombres)
function findColumnIndexExact(headers, posiblesNombres) {
for (let i = 0; i < headers.length; i++) {
  const header = headers[i].toLowerCase().trim();
  for (const nombre of posiblesNombres) {
    if (header === nombre.toLowerCase().trim() || header.includes(nombre.toLowerCase().trim())) {
      Logger.log(`Columna encontrada: "${headers[i]}" en índice ${i} para búsqueda "${nombre}"`);
      return i;
    }
  }
}
Logger.log(`No se encontró columna para: ${posiblesNombres.join(', ')}`);
return -1;
}

// Función mejorada para obtener TODOS los empleados (para debug)
function obtenerTodosLosEmpleados() {
try {
  Logger.log('=== OBTENIENDO TODOS LOS EMPLEADOS (SIN FILTROS) ===');
  const spreadsheet = SpreadsheetApp.openById(EMPLOYEES_SHEET_ID);
  const sheets = spreadsheet.getSheets();
  let sheet = null;
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().toLowerCase() === 'maestra') {
      sheet = sheets[i];
      break;
    }
  }
  if (!sheet) {
    Logger.log('No se encontró hoja "Maestra"');
    return {error: true, message: 'No se encontró la hoja "Maestra"', availableSheets: sheets.map(sh => sh.getName())};
  }
  const data = sheet.getDataRange().getValues();
  Logger.log(`Procesando ${data.length} filas...`);
  if (data.length < 2) {
    return {error: true, message: 'La hoja no contiene datos suficientes'};
  }
  const headers = data[0].map(h => h.toString().trim());
  Logger.log('Encabezados:', headers);
  // Usar los nombres exactos de tus columnas
  const docIndex = findColumnIndexExact(headers, ['03_DOCUMENTO_IDENTIDIDAD', 'documento_identidad', 'documento']);
  const nombreIndex = findColumnIndexExact(headers, ['02_EMPLEADO', 'empleado', 'nombre']);
  const cargoIndex = findColumnIndexExact(headers, ['12_NOMBRE_CARGO', 'nombre_cargo', 'cargo']);
  const areaIndex = findColumnIndexExact(headers, ['area','departamento','sector','division','zona','seccion']);
  Logger.log('Índices:', {docIndex, nombreIndex, cargoIndex, areaIndex});
  if (docIndex === -1 || nombreIndex === -1) {
    return {error: true, message: 'No se encontraron las columnas requeridas (Documento o Nombre)', headers: headers, expectedColumns: ['03_DOCUMENTO_IDENTIDIDAD', '02_EMPLEADO']};
  }
  const empleados = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row.length <= Math.max(docIndex, nombreIndex)) {
      Logger.log(`Fila ${i+1} omitida - no tiene suficientes columnas`);
      continue;
    }
    const documento = row[docIndex] ? row[docIndex].toString().trim() : '';
    const nombre = row[nombreIndex] ? row[nombreIndex].toString().trim() : '';
    const cargo = cargoIndex !== -1 && row[cargoIndex] ? row[cargoIndex].toString().trim() : '';
    const area = areaIndex !== -1 && row[areaIndex] ? row[areaIndex].toString().trim() : 'Sin área'; // Usar el área de la hoja si existe
    if (documento && nombre) {
      empleados.push({
        documento,
        nombre,
        cargo,
        area,
        fila: i + 1
      });
      Logger.log(`Empleado agregado: ${documento} - ${nombre} - ${cargo} - ${area}`);
    } else {
      Logger.log(`Fila ${i+1} omitida - documento o nombre vacío: "${documento}" - "${nombre}"`);
    }
  }
  Logger.log(`Total empleados encontrados: ${empleados.length}`);
  return empleados;
} catch (error) {
  console.error('Error en obtenerTodosLosEmpleados:', error);
  return {error: true, message: 'Error: ' + error.message};
}
}

// Función para obtener estructura de datos (para diagnóstico)
function obtenerEstructuraHoja() {
try {
  const spreadsheet = SpreadsheetApp.openById(EMPLOYEES_SHEET_ID);
  const sheets = spreadsheet.getSheets();
  const sheet = sheets.find(sh => sh.getName().toLowerCase() === 'maestra');
  if (!sheet) {
    return {success: false, message: 'No se encontró la hoja "Maestra"', availableSheets: sheets.map(sh => sh.getName())};
  }
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => h.toString().trim());
  const sampleRow = data.length > 1 ? data[1].map(c => c.toString().trim()) : [];
  return {success: true, sheetName: sheet.getName(), headers: headers, sampleRow: sampleRow, totalRows: data.length};
} catch (error) {
  return {success: false, message: 'Error: ' + error.message};
}
}

// Helper function to parse shift string and get its components
function parseShiftString(shiftString) {
const parts = shiftString.split(' ');
if (parts.length < 4) {
  return { shiftType: shiftString, scenario: null, dominical: null };
}
const shiftType = parts[0]; // e.g., T1, T2, T3
const scenario = parts[1] + ' ' + parts[2]; // e.g., ESCENARIO 1
const dominical = parts[3] + (parts[4] ? ' ' + parts[4] : ''); // e.g., NO DOMINICAL, SI DOMINICAL
return { shiftType, scenario, dominical };
}

// Function to calculate total hours for a weekly schedule
function isExtraTurn(turn) {
  return turn.includes('EXTRA');
}

function calculateWeeklyHours(turnos) {
  let totalHours = 0;
  const days = ['lunes', 'martes', 'miercoles', 'jueves', 'viernes', 'sabado', 'domingo'];
  
  days.forEach(day => {
    const shiftValue = turnos[day];
    if (shiftValue && (shiftValue.startsWith('T1') || shiftValue.startsWith('T2') || shiftValue.startsWith('T3'))) {
      const hours = SHIFT_DEFINITIONS[shiftValue][capitalizeFirstLetter(day)];
      if (hours !== undefined) {
        totalHours += hours;
      }
    }
  });
  
  return totalHours;
}

// Function to calculate total hours for a range schedule
function calculateRangeHours(fechaInicio, fechaFin, turnoAsignado) {
let totalHours = 0;
const { shiftType, scenario, dominical } = parseShiftString(turnoAsignado);
Logger.log(`Calculating range hours for ${fechaInicio} to ${fechaFin} with shift: ${turnoAsignado}`);
if (!shiftType || !scenario || !dominical || !SHIFT_DEFINITIONS[scenario] || !SHIFT_DEFINITIONS[scenario][dominical] || !SHIFT_DEFINITIONS[scenario][dominical][shiftType]) {
  Logger.log(`Definición de turno no encontrada para rango: ${turnoAsignado}`);
  return 0;
}
const dailyHoursMap = SHIFT_DEFINITIONS[scenario][dominical][shiftType];
const startDate = new Date(fechaInicio);
const endDate = new Date(fechaFin);
let currentDate = new Date(startDate);
while (currentDate <= endDate) {
  const dayOfWeek = currentDate.getDay(); // 0 for Sunday, 1 for Monday, ..., 6 for Saturday
  let dayName;
  switch (dayOfWeek) {
    case 0: dayName = 'Domingo'; break;
    case 1: dayName = 'Lunes'; break;
    case 2: dayName = 'Martes'; break;
    case 3: dayName = 'Miercoles'; break;
    case 4: dayName = 'Jueves'; break;
    case 5: dayName = 'Viernes'; break;
    case 6: dayName = 'Sabado'; break;
  }
  const hours = dailyHoursMap[dayName];
  if (hours !== undefined) {
    totalHours += hours;
    Logger.log(`  Date: ${currentDate.toDateString()}, Day: ${dayName}, Hours: ${hours}. Current total: ${totalHours}`);
  } else {
    Logger.log(`Horas no definidas para ${shiftType} en ${scenario} ${dominical} para el día ${dayName}`);
  }
  currentDate.setDate(currentDate.getDate() + 1);
}
Logger.log('Final range total hours: ' + totalHours);
return totalHours;
}

// Function to calculate total hours for an annual schedule
function calculateAnnualHours(año, tipoTurno, turnoInicio, fijo) {
Logger.log(`Calculating annual hours for year: ${año}, type: ${tipoTurno}, start: ${turnoInicio}, fixed: ${fijo}`);
// For annual shifts, we'll use the 'total' hours defined in SHIFT_DEFINITIONS
// and multiply by 52 weeks, or use a simplified average if not fixed.
// This is a simplification, as annual shifts are more complex.
// For now, if it's a fixed T-shift, we'll use its weekly total.
// If it's not a T-shift or not fixed, we'll assume 44 hours for calculation.
if (fijo && turnoInicio.startsWith('T')) {
  const { shiftType, scenario, dominical } = parseShiftString(turnoInicio);
  if (SHIFT_DEFINITIONS[scenario] && SHIFT_DEFINITIONS[scenario][dominical] && SHIFT_DEFINITIONS[scenario][dominical][shiftType]) {
    const annualTotal = SHIFT_DEFINITIONS[scenario][dominical][shiftType].total * 52; // Total hours for the year
    Logger.log(`  Fixed T-shift annual total: ${annualTotal}`);
    return annualTotal;
  }
}
// Default for non-T shifts or non-fixed annual shifts
const defaultAnnualTotal = 44 * 52; // Assume 44 hours per week for 52 weeks
Logger.log(`  Default annual total (non-T or non-fixed): ${defaultAnnualTotal}`);
return defaultAnnualTotal;
}

// Funciones para gestión de turnos (usando EMPLOYEES_SHEET_ID)
function guardarProgramacionSemanal(datos) {
try {
  const spreadsheet = SpreadsheetApp.openById(EMPLOYEES_SHEET_ID);
  let sheet = spreadsheet.getSheetByName('turnos_semanales');
  if (!sheet) {
    sheet = spreadsheet.insertSheet('turnos_semanales');
    // Agregar columna de comentarios al header
    sheet.getRange(1, 1, 1, 14).setValues([[
      'Fecha_Creacion', 'Empleado_Doc', 'Empleado_Nombre',
      'Empleado_Cargo', 'Empleado_Area', 'Semana_Inicio',
      'Lunes', 'Martes', 'Miercoles', 'Jueves',
      'Viernes', 'Sabado', 'Domingo', 'Comentario'
    ]]);
    // Formatear encabezados
    const headerRange = sheet.getRange(1, 1, 1, 14);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4CAF50');
    headerRange.setFontColor('white');
  }

  const totalHours = calculateWeeklyHours(datos.turnos);
  const overHours = totalHours > 44;

  const nuevaFila = [
    new Date(),
    datos.empleado.documento,
    datos.empleado.nombre,
    datos.empleado.cargo || 'Sin cargo',
    datos.empleado.area || 'Sin área',
    datos.semanaInicio,
    datos.turnos.lunes || 'DESCANSO',
    datos.turnos.martes || 'DESCANSO',
    datos.turnos.miercoles || 'DESCANSO',
    datos.turnos.jueves || 'DESCANSO',
    datos.turnos.viernes || 'DESCANSO',
    datos.turnos.sabado || 'DESCANSO',
    datos.turnos.domingo || 'DESCANSO',
    datos.comentario || '' // Nueva columna para comentarios
  ];

  sheet.appendRow(nuevaFila);
  return {
    success: true,
    message: 'Programación semanal guardada correctamente',
    overHours: overHours,
    totalHours: totalHours
  };
} catch (error) {
  console.error('Error al guardar programación semanal:', error);
  return {success: false, message: 'Error al guardar: ' + error.toString()};
}
}

function guardarProgramacionRango(datos) {
try {
  const spreadsheet = SpreadsheetApp.openById(EMPLOYEES_SHEET_ID);
  let sheet = spreadsheet.getSheetByName('turnos_rango');
  if (!sheet) {
    sheet = spreadsheet.insertSheet('turnos_rango');
    // Agregar columna de comentarios al header
    sheet.getRange(1, 1, 1, 9).setValues([[
      'Fecha_Creacion', 'Empleado_Doc', 'Empleado_Nombre', 
      'Empleado_Cargo', 'Empleado_Area', 'Fecha_Inicio', 
      'Fecha_Fin', 'Turno_Asignado', 'Comentario'
    ]]);
    // Formatear encabezados
    const headerRange = sheet.getRange(1, 1, 1, 9);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#2196F3');
    headerRange.setFontColor('white');
  }

  const totalHours = calculateRangeHours(datos.fechaInicio, datos.fechaFin, datos.turnoAsignado);
  const overHours = totalHours > 44;

  const nuevaFila = [
    new Date(),
    datos.empleado.documento,
    datos.empleado.nombre,
    datos.empleado.cargo || 'Sin cargo',
    datos.empleado.area || 'Sin área',
    datos.fechaInicio,
    datos.fechaFin, 
    datos.turnoAsignado,
    datos.comentario || '' // Nueva columna para comentarios
  ];

  sheet.appendRow(nuevaFila);
  return {success: true, message: 'Programación por rango guardada correctamente', overHours: overHours, totalHours: totalHours};
} catch (error) {
  console.error('Error al guardar programación por rango:', error);
  return {success: false, message: 'Error al guardar: ' + error.toString()};
}
}

function guardarProgramacionAnual(datos) {
try {
  const spreadsheet = SpreadsheetApp.openById(EMPLOYEES_SHEET_ID);
  let sheet = spreadsheet.getSheetByName('turnos_anuales');
  if (!sheet) {
    sheet = spreadsheet.insertSheet('turnos_anuales');
    // Agregar columna de comentarios al header
    sheet.getRange(1, 1, 1, 10).setValues([[
      'Fecha_Creacion', 'Empleado_Doc', 'Empleado_Nombre', 
      'Empleado_Cargo', 'Empleado_Area', 'Año', 
      'Tipo_Turno', 'Turno_Inicio', 'Fijo', 'Comentario'
    ]]);
    // Formatear encabezados
    const headerRange = sheet.getRange(1, 1, 1, 10);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#FF9800');
    headerRange.setFontColor('white');
  }

  const totalHours = calculateAnnualHours(datos.año, datos.tipoTurno, datos.turnoInicio, datos.fijo);
  const overHours = totalHours > (44 * 52);

  const nuevaFila = [
    new Date(),
    datos.empleado.documento,
    datos.empleado.nombre,
    datos.empleado.cargo || 'Sin cargo',
    datos.empleado.area || 'Sin área',
    datos.año,
    datos.tipoTurno,
    datos.turnoInicio,
    datos.fijo ? 'SI' : 'NO',
    datos.comentario || '' // Nueva columna para comentarios
  ];

  sheet.appendRow(nuevaFila);
  return {success: true, message: 'Programación anual guardada correctamente', overHours: overHours, totalHours: totalHours};
} catch (error) {
  console.error('Error al guardar programación anual:', error);
  return {success: false, message: 'Error al guardar: ' + error.toString()};
}
}

function obtenerProgramaciones(tipo, busqueda = null, areaFiltro = null) {
  try {
    const spreadsheet = SpreadsheetApp.openById(EMPLOYEES_SHEET_ID);
    let sheet;
    let headers;
    let data;
    let programaciones = [];
    switch (tipo) {
      case 'semanal':
        sheet = spreadsheet.getSheetByName('turnos_semanales');
        if (!sheet) return [];
        data = sheet.getDataRange().getValues();
        if (data.length <= 1) return [];
        headers = data[0];

        // Solo la última programación por documento
        let ultimosPorDocSemanal = {};
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          if (!row || row.length < 13) continue;
          const empleadoDoc = row[1] ? row[1].toString() : '';
          const empleadoArea = row[4] ? row[4].toString() : 'Sin área';
          if (areaFiltro && empleadoArea.toLowerCase().trim() !== areaFiltro.toLowerCase().trim()) continue;
          if (busqueda) {
            const busquedaLower = busqueda.toLowerCase();
            if (!(empleadoDoc.toLowerCase().includes(busquedaLower) ||
                  (row[2] && row[2].toString().toLowerCase().includes(busquedaLower)))) continue;
          }
          // Guardar solo la última (la más abajo en la hoja)
          ultimosPorDocSemanal[empleadoDoc] = { row, fila: i + 1 };
        }
        Object.values(ultimosPorDocSemanal).forEach(obj => {
          const row = obj.row;
          programaciones.push({
            fila: obj.fila,
            fechaCreacion: formatDate(row[0]),
            empleadoDoc: row[1] ? row[1].toString() : '',
            empleadoNombre: row[2] ? row[2].toString() : '',
            empleadoCargo: row[3] ? row[3].toString() : 'Sin cargo',
            empleadoArea: row[4] ? row[4].toString() : 'Sin área',
            semanaInicio: row[5] ? row[5].toString() : '',
            turnos: {
              lunes: row[6] ? row[6].toString() : 'DESCANSO',
              martes: row[7] ? row[7].toString() : 'DESCANSO',
              miercoles: row[8] ? row[8].toString() : 'DESCANSO',
              jueves: row[9] ? row[9].toString() : 'DESCANSO',
              viernes: row[10] ? row[10].toString() : 'DESCANSO',
              sabado: row[11] ? row[11].toString() : 'DESCANSO',
              domingo: row[12] ? row[12].toString() : 'DESCANSO'
            }
          });
        });
        break;
      case 'rango':
        sheet = spreadsheet.getSheetByName('turnos_rango');
        if (!sheet) return [];
        data = sheet.getDataRange().getValues();
        if (data.length <= 1) return [];
        headers = data[0];

        // Solo la última programación por documento
        let ultimosPorDocRango = {};
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          if (!row || row.length < 8) continue;
          const empleadoDoc = row[1] ? row[1].toString() : '';
          const empleadoArea = row[4] ? row[4].toString() : 'Sin área';
          if (areaFiltro && empleadoArea.toLowerCase().trim() !== areaFiltro.toLowerCase().trim()) continue;
          if (busqueda) {
            const busquedaLower = busqueda.toLowerCase();
            if (!(empleadoDoc.toLowerCase().includes(busquedaLower) ||
                  (row[2] && row[2].toString().toLowerCase().includes(busquedaLower)))) continue;
          }
          ultimosPorDocRango[empleadoDoc] = { row, fila: i + 1 };
        }
        Object.values(ultimosPorDocRango).forEach(obj => {
          const row = obj.row;
          programaciones.push({
            fila: obj.fila,
            fechaCreacion: formatDate(row[0]),
            empleadoDoc: row[1] ? row[1].toString() : '',
            empleadoNombre: row[2] ? row[2].toString() : '',
            empleadoCargo: row[3] ? row[3].toString() : 'Sin cargo',
            empleadoArea: row[4] ? row[4].toString() : 'Sin área',
            fechaInicio: row[5] ? row[5].toString() : '',
            fechaFin: row[6] ? row[6].toString() : '',
            turnoAsignado: row[7] ? row[7].toString() : ''
          });
        });
        break;
      case 'anual':
        sheet = spreadsheet.getSheetByName('turnos_anuales');
        if (!sheet) return [];
        data = sheet.getDataRange().getValues();
        if (data.length <= 1) return [];
        headers = data[0];

        // Solo la última programación por documento
        let ultimosPorDocAnual = {};
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          if (!row || row.length < 9) continue;
          const empleadoDoc = row[1] ? row[1].toString() : '';
          const empleadoArea = row[4] ? row[4].toString() : 'Sin área';
          if (areaFiltro && empleadoArea.toLowerCase().trim() !== areaFiltro.toLowerCase().trim()) continue;
          if (busqueda) {
            const busquedaLower = busqueda.toLowerCase();
            // Si la búsqueda es un año, filtra por año
            if (!isNaN(busquedaLower) && !isNaN(parseFloat(busquedaLower))) {
              if (row[5] && row[5].toString() !== busqueda) continue;
            } else {
              if (!(empleadoDoc.toLowerCase().includes(busquedaLower) ||
                    (row[2] && row[2].toString().toLowerCase().includes(busquedaLower)))) continue;
            }
          }
          ultimosPorDocAnual[empleadoDoc] = { row, fila: i + 1 };
        }
        Object.values(ultimosPorDocAnual).forEach(obj => {
          const row = obj.row;
          programaciones.push({
            fila: obj.fila,
            fechaCreacion: formatDate(row[0]),
            empleadoDoc: row[1] ? row[1].toString() : '',
            empleadoNombre: row[2] ? row[2].toString() : '',
            empleadoCargo: row[3] ? row[3].toString() : 'Sin cargo',
            empleadoArea: row[4] ? row[4].toString() : 'Sin área',
            año: row[5] ? row[5].toString() : '',
            tipoTurno: row[6] ? row[6].toString() : '',
            turnoInicio: row[7] ? row[7].toString() : '',
            fijo: row[8] ? (row[8].toString().toUpperCase() === 'SI') : false
          });
        });
        break;
      default:
        console.warn('Tipo de programación no reconocido:', tipo);
        return [];
    }
    return programaciones;
  } catch (error) {
    console.error('Error en obtenerProgramaciones:', error);
    return [];
  }
}
// Función auxiliar para obtener el área del usuario (se mantiene por si se usa en otro lugar, pero no para filtrar en obtenerProgramaciones)
function obtenerAreaUsuario(email) {
try {
  const spreadsheet = SpreadsheetApp.openById(USERS_SHEET_ID);
  const sheet = spreadsheet.getSheetByName('usuario'); // Cambiado a 'usuario' si esa es tu hoja de usuarios
  if (!sheet) {
    Logger.log("Hoja 'usuario' no encontrada para obtener el área del usuario.");
    return 'Sin área';
  }
  const data = sheet.getDataRange().getValues();
  // Asumiendo que la columna de email es la primera (índice 0) y la de área es la cuarta (índice 3)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] && row[0].toString().trim().toLowerCase() === email.toLowerCase()) {
      return row[3] ? row[3].toString().trim() : 'Sin área'; // Asume que el área está en la columna 4 (índice 3)
    }
  }
  Logger.log(`Área no encontrada para el usuario: ${email}`);
  return 'Sin área';
} catch (error) {
  console.error('Error en obtenerAreaUsuario:', error);
  return 'Error al obtener área';
}
}

function formatDate(date) {
if (!date) return '';
if (typeof date === 'string') return date;
if (date instanceof Date) return date.toISOString();
return date.toString();
}

function eliminarProgramacion(tipo, fila) {
try {
  const spreadsheet = SpreadsheetApp.openById(EMPLOYEES_SHEET_ID);
  const sheetNames = {
    'semanal': 'turnos_semanales',
    'rango': 'turnos_rango',
    'anual': 'turnos_anuales'
  };
  const sheet = spreadsheet.getSheetByName(sheetNames[tipo]);
  if (!sheet) return {success: false, message: 'Hoja no encontrada'};
  sheet.deleteRow(fila);
  return {success: true, message: 'Programación eliminada correctamente'};
} catch (error) {
  console.error('Error al eliminar programación:', error);
  return {success: false, message: 'Error al eliminar: ' + error.toString()};
}
}

function obtenerProgramacionPorFila(tipo, fila) {
  try {
    const spreadsheet = SpreadsheetApp.openById(EMPLOYEES_SHEET_ID);
    let sheet;
    let programacion = null;
    Logger.log(`[obtenerProgramacionPorFila] Solicitud para tipo: ${tipo}, fila: ${fila}`);
    switch (tipo) {
      case 'semanal':
        sheet = spreadsheet.getSheetByName('turnos_semanales');
        if (!sheet) {
          Logger.log('[obtenerProgramacionPorFila] No se encontró la hoja turnos_semanales');
          return null;
        }
        const dataSemanal = sheet.getDataRange().getValues();
        if (fila < 2 || fila > dataSemanal.length) {
          Logger.log(`[obtenerProgramacionPorFila] Número de fila inválido para semanal: ${fila}. Total filas: ${dataSemanal.length}`);
          return null;
        }
        const rowSemanal = dataSemanal[fila - 1];
        const expectedSemanalLength = 13;
        if (!rowSemanal || rowSemanal.length < expectedSemanalLength) {
          Logger.log(`[obtenerProgramacionPorFila] Fila semanal incompleta o inválida en fila ${fila}. Longitud actual: ${rowSemanal ? rowSemanal.length : 'null'}. Esperado: ${expectedSemanalLength}`);
          return null;
        }
        // Buscar el índice de la columna Comentario
        const headers = dataSemanal[0];
        const idxComentario = headers.indexOf('Comentario');
        programacion = {
          fila: fila,
          fechaCreacion: rowSemanal[0],
          empleadoDoc: rowSemanal[1] ? rowSemanal[1].toString() : '',
          empleadoNombre: rowSemanal[2] ? rowSemanal[2].toString() : '',
          empleadoCargo: rowSemanal[3] ? rowSemanal[3].toString() : 'Sin cargo',
          empleadoArea: rowSemanal[4] ? rowSemanal[4].toString() : 'Sin área',
          semanaInicio: rowSemanal[5] ? rowSemanal[5].toString() : '',
          turnos: {
            lunes: rowSemanal[6] ? rowSemanal[6].toString() : 'DESCANSO',
            martes: rowSemanal[7] ? rowSemanal[7].toString() : 'DESCANSO',
            miercoles: rowSemanal[8] ? rowSemanal[8].toString() : 'DESCANSO',
            jueves: rowSemanal[9] ? rowSemanal[9].toString() : 'DESCANSO',
            viernes: rowSemanal[10] ? rowSemanal[10].toString() : 'DESCANSO',
            sabado: rowSemanal[11] ? rowSemanal[11].toString() : 'DESCANSO',
            domingo: rowSemanal[12] ? rowSemanal[12].toString() : 'DESCANSO'
          },
          comentario: idxComentario !== -1 ? rowSemanal[idxComentario] || '' : ''
        };
        break;
    case 'rango':
      sheet = spreadsheet.getSheetByName('turnos_rango');
      if (!sheet) {
        Logger.log('[obtenerProgramacionPorFila] No se encontró la hoja turnos_rango');
        return null;
      }
      const dataRango = sheet.getDataRange().getValues();
      if (fila < 2 || fila > dataRango.length) {
        Logger.log(`[obtenerProgramacionPorFila] Número de fila inválido para rango: ${fila}. Total filas: ${dataRango.length}`);
        return null;
      }
      const rowRango = dataRango[fila - 1];
      const expectedRangoLength = 8;
      if (!rowRango || rowRango.length < expectedRangoLength) {
        Logger.log(`[obtenerProgramacionPorFila] Fila rango incompleta o inválida en fila ${fila}. Longitud actual: ${rowRango ? rowRango.length : 'null'}. Esperado: ${expectedRangoLength}`);
        return null;
      }
      programacion = {
        fila: fila,
        fechaCreacion: rowRango[0],
        empleadoDoc: rowRango[1] ? rowRango[1].toString() : '',
        empleadoNombre: rowRango[2] ? rowRango[2].toString() : '',
        empleadoCargo: rowRango[3] ? rowRango[3].toString() : 'Sin cargo',
        empleadoArea: rowRango[4] ? rowRango[4].toString() : 'Sin área',
        fechaInicio: rowRango[5] ? rowRango[5].toString() : '',
        fechaFin: rowRango[6] ? rowRango[6].toString() : '',
        turnoAsignado: rowRango[7] ? rowRango[7].toString() : ''
      };
      break;
    case 'anual':
      sheet = spreadsheet.getSheetByName('turnos_anuales');
      if (!sheet) {
        Logger.log('[obtenerProgramacionPorFila] No se encontró la hoja turnos_anuales');
        return null;
      }
      const dataAnual = sheet.getDataRange().getValues();
      if (fila < 2 || fila > dataAnual.length) {
        Logger.log(`[obtenerProgramacionPorFila] Número de fila inválido para anual: ${fila}. Total filas: ${dataAnual.length}`);
        return null;
      }
      const rowAnual = dataAnual[fila - 1];
      const expectedAnualLength = 9;
      if (!rowAnual || rowAnual.length < expectedAnualLength) {
        Logger.log(`[obtenerProgramacionPorFila] Fila anual incompleta o inválida en fila ${fila}. Longitud actual: ${rowAnual ? rowAnual.length : 'null'}. Esperado: ${expectedAnualLength}`);
        return null;
      }
      programacion = {
        fila: fila,
        fechaCreacion: rowAnual[0],
        empleadoDoc: rowAnual[1] ? rowAnual[1].toString() : '',
        empleadoNombre: rowAnual[2] ? rowAnual[2].toString() : '',
        empleadoCargo: rowAnual[3] ? rowAnual[3].toString() : 'Sin cargo',
        empleadoArea: rowAnual[4] ? rowAnual[4].toString() : 'Sin área',
        año: rowAnual[5] ? rowAnual[5].toString() : '',
        tipoTurno: rowAnual[6] ? rowAnual[6].toString() : '',
        turnoInicio: rowAnual[7] ? rowAnual[7].toString() : '',
        fijo: rowAnual[8] ? (rowAnual[8].toString().toUpperCase() === 'SI') : false
      };
      break;
    default:
      Logger.log('[obtenerProgramacionPorFila] Tipo de programación no reconocido:', tipo);
      return null;
  }
  // Convertir el objeto a una cadena JSON antes de devolverlo
   const jsonProgramacion = JSON.stringify(programacion);
    Logger.log('[obtenerProgramacionPorFila] Devolviendo JSON:', jsonProgramacion);
    return jsonProgramacion;
  } catch (error) {
    console.error('[obtenerProgramacionPorFila] Error en obtenerProgramacionPorFila:', error);
    return null;
  }
}

// Función para actualizar una programación semanal
// ...existing code...
// Función para actualizar una programación semanal
// ...existing code...
function actualizarProgramacionSemanal(fila, turnos) {
  try {
    const spreadsheet = SpreadsheetApp.openById(EMPLOYEES_SHEET_ID);
    const sheet = spreadsheet.getSheetByName('turnos_semanales');
    if (!sheet) return {success: false, message: 'Hoja no encontrada'};
    
    const totalHours = calculateWeeklyHours(turnos);
    const overHours = totalHours > 44;
    Logger.log(`Actualizando semanal (fila ${fila}): Total Horas = ${totalHours}, OverHours = ${overHours}`);

    // Obtener todos los headers para encontrar la columna de comentario
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const idxComentario = headers.indexOf('comentario');
    
    // Actualizar solo los turnos (columnas 7 a 13)
    sheet.getRange(fila, 7, 1, 7).setValues([[
      turnos.lunes  || 'DESCANSO',
      turnos.martes || 'DESCANSO',
      turnos.miercoles || 'DESCANSO',
      turnos.jueves || 'DESCANSO',
      turnos.viernes || 'DESCANSO',
      turnos.sabado || 'DESCANSO',
      turnos.domingo || 'DESCANSO'
    ]]);

    // Actualizar comentario si existe la columna
    if (idxComentario !== -1 && turnos.comentario !== undefined) {
      sheet.getRange(fila, idxComentario + 1).setValue(turnos.comentario);
      Logger.log(`[actualizarProgramacionSemanal] Comentario actualizado en columna ${idxComentario + 1}: ${turnos.comentario}`);
    } else {
      Logger.log(`[actualizarProgramacionSemanal] No se pudo actualizar comentario. Índice: ${idxComentario}, Valor: ${turnos.comentario}`);
    }

    // Actualizar fecha de modificación
    sheet.getRange(fila, 1).setValue(new Date());

    return {success: true, message: 'Programación actualizada correctamente', overHours: overHours, totalHours: totalHours};
  } catch (error) {
    console.error('Error en actualizarProgramacionSemanal:', error);
    return {success: false, message: 'Error al actualizar: ' + error.message};
  }
}
// ...existing code...
// Función para actualizar una programación por rango
function actualizarProgramacionRango(fila, datos) {
Logger.log(`[actualizarProgramacionRango] Recibido: fila=${fila}, datos=`, JSON.stringify(datos));
try {
  const spreadsheet = SpreadsheetApp.openById(EMPLOYEES_SHEET_ID);
  const sheet = spreadsheet.getSheetByName('turnos_rango');
  if (!sheet) {
    Logger.log('[actualizarProgramacionRango] Hoja turnos_rango no encontrada');
    return { success: false, message: 'Hoja no encontrada' };
  }

  if (fila < 2 || fila > sheet.getLastRow()) {
    Logger.log(`[actualizarProgramacionRango] Fila inválida: ${fila}. Última fila: ${sheet.getLastRow()}`);
    return { success: false, message: 'Fila de programación no válida' };
  }

  const totalHours = calculateRangeHours(datos.fechaInicio, datos.fechaFin, datos.turnoAsignado);
  const overHours = totalHours > 44;

  // Actualizar los datos incluyendo el comentario
  const range = sheet.getRange(fila, 6, 1, 4);
  const valuesToSet = [
    [
      datos.fechaInicio,
      datos.fechaFin,
      datos.turnoAsignado,
      datos.comentario || '' // Mantener comentario existente si no hay nuevo
    ]
  ];

  Logger.log('[actualizarProgramacionRango] Valores a establecer:', JSON.stringify(valuesToSet));
  range.setValues(valuesToSet);
  
  // Actualizar fecha de modificación
  sheet.getRange(fila, 1).setValue(new Date());
  
  Logger.log(`[actualizarProgramacionRango] Programación en fila ${fila} actualizada correctamente.`);
  return { success: true, message: 'Programación actualizada correctamente', overHours: overHours, totalHours: totalHours };
} catch (error) {
  console.error('Error en actualizarProgramacionRango:', error);
  return { success: false, message: 'Error al actualizar: ' + error.message };
}
}

function actualizarProgramacionAnual(fila, datos) {
Logger.log(`[actualizarProgramacionAnual] Recibido: fila=${fila}, datos=`, JSON.stringify(datos));
try {
  const spreadsheet = SpreadsheetApp.openById(EMPLOYEES_SHEET_ID);
  const sheet = spreadsheet.getSheetByName('turnos_anuales');
  if (!sheet) {
    Logger.log('[actualizarProgramacionAnual] Hoja turnos_anuales no encontrada');
    return { success: false, message: 'Hoja no encontrada' };
  }

  if (fila < 2 || fila > sheet.getLastRow()) {
    Logger.log(`[actualizarProgramacionAnual] Fila inválida: ${fila}. Última fila: ${sheet.getLastRow()}`);
    return { success: false, message: 'Fila de programación no válida' };
  }

  const totalHours = calculateAnnualHours(datos.año, datos.tipoTurno, datos.turnoInicio, datos.fijo);
  const overHours = totalHours > (44 * 52);

  // Actualizar los datos incluyendo el comentario
  const range = sheet.getRange(fila, 6, 1, 5);
  const valuesToSet = [
    [
      datos.año,
      datos.tipoTurno,
      datos.turnoInicio,
      datos.fijo ? 'SI' : 'NO',
      datos.comentario || '' // Mantener comentario existente si no hay nuevo
    ]
  ];

  Logger.log('[actualizarProgramacionAnual] Valores a establecer:', JSON.stringify(valuesToSet));
  range.setValues(valuesToSet);
  
  // Actualizar fecha de modificación
  sheet.getRange(fila, 1).setValue(new Date());
  
  Logger.log(`[actualizarProgramacionAnual] Programación en fila ${fila} actualizada correctamente.`);
  return { success: true, message: 'Programación actualizada correctamente', overHours: overHours, totalHours: totalHours };
} catch (error) {
  console.error('Error en actualizarProgramacionAnual:', error);
  return { success: false, message: 'Error al actualizar: ' + error.message };
}
}

function actualizarHojasExistentes() {
try {
  const spreadsheet = SpreadsheetApp.openById(EMPLOYEES_SHEET_ID);
  const sheetNames = ['turnos_semanales', 'turnos_rango', 'turnos_anuales'];
  const resultados = [];
  for (const sheetName of sheetNames) {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      if (data.length > 0) {
        const headers = data[0];
        // Verificar si ya tiene las columnas de cargo y área
        const tieneCargo = headers.includes('Empleado_Cargo');
        const tieneArea = headers.includes('Empleado_Area');
        if (!tieneCargo || !tieneArea) {
          resultados.push({
            hoja: sheetName,
            mensaje: 'Hoja necesita actualización manual - tiene datos existentes',
            headers: headers
          });
        } else {
          resultados.push({
            hoja: sheetName,
            mensaje: 'Hoja ya tiene las columnas necesarias',
            headers: headers
          });
        }
      } else {
        resultados.push({
          hoja: sheetName,
          mensaje: 'Hoja vacía - se actualizará automáticamente en el próximo guardado'
        });
      }
    } else {
      resultados.push({
        hoja: sheetName,
        mensaje: 'Hoja no existe - se creará automáticamente en el próximo guardado'
      });
    }
  }
  return {success: true, resultados: resultados};
} catch (error) {
  console.error('Error al verificar hojas:', error);
  return {success: false, message: 'Error al verificar hojas: ' + error.message};
}
}

function testConnection() {
return {success: true, message: "Conexión exitosa"};
}


// ...existing code...
function getPublicShifts(filtro = {}) {
  try {
    const ss = SpreadsheetApp.openById(EMPLOYEES_SHEET_ID);
    let turnos = [];

    // --- Procesar turnos semanales ---
    const hojaSemanales = ss.getSheetByName('turnos_semanales');
    if (hojaSemanales) {
      const datos = hojaSemanales.getDataRange().getValues();
      const headers = datos[0];
      const idxNombre = headers.indexOf('Empleado_Nombre');
      const idxDoc = headers.indexOf('Empleado_Doc');
      const idxCargo = headers.indexOf('Empleado_Cargo');
      const idxArea = headers.indexOf('Empleado_Area');
      const idxComentario = headers.indexOf('comentario');
      const idxPeriodo = headers.indexOf('Semana_Inicio');
      const idxDias = [
        headers.indexOf('Lunes'),
        headers.indexOf('Martes'),
        headers.indexOf('Miercoles'),
        headers.indexOf('Jueves'),
        headers.indexOf('Viernes'),
        headers.indexOf('Sabado'),
        headers.indexOf('Domingo')
      ];

      let ultimosPorDoc = {};
      for (let i = 1; i < datos.length; i++) {
        const row = datos[i];
        const doc = row[idxDoc] ? row[idxDoc].toString() : '';
        const nombre = row[idxNombre] ? row[idxNombre].toString() : '';
        if (
          (filtro.documento && doc.toLowerCase().includes(filtro.documento.toLowerCase())) ||
          (filtro.nombre && nombre.toLowerCase().includes(filtro.nombre.toLowerCase()))
        ) {
          ultimosPorDoc[doc] = row;
        }
        // Si no hay filtro, mostrar todos
        if (!filtro.documento && !filtro.nombre) {
          ultimosPorDoc[doc] = row;
        }
      }
      Object.values(ultimosPorDoc).forEach(row => {
        turnos.push({
          nombre: row[idxNombre] || '',
          documento: row[idxDoc] || '',
          cargo: row[idxCargo] || '',
          area: row[idxArea] || '',
          lunes: row[idxDias[0]] || '',
          martes: row[idxDias[1]] || '',
          miercoles: row[idxDias[2]] || '',
          jueves: row[idxDias[3]] || '',
          viernes: row[idxDias[4]] || '',
          sabado: row[idxDias[5]] || '',
          domingo: row[idxDias[6]] || '',
          periodo: row[idxPeriodo] || '',
          comentario: row[idxComentario] || '',
          tipo: 'Semanal'
        });
      });
    }

    // --- Procesar turnos de rango ---
    const hojaRango = ss.getSheetByName('turnos_rango');
    if (hojaRango) {
      const datos = hojaRango.getDataRange().getValues();
      const headers = datos[0];
      const idxNombre = headers.indexOf('Empleado_Nombre');
      const idxDoc = headers.indexOf('Empleado_Doc');
      const idxCargo = headers.indexOf('Empleado_Cargo');
      const idxArea = headers.indexOf('Empleado_Area');
      const idxTurno = headers.indexOf('Turno_Asignado');
      const idxComentario = headers.indexOf('comentario');
      const idxFechaInicio = headers.indexOf('Fecha_Inicio');
      const idxFechaFin = headers.indexOf('Fecha_Fin');

      let ultimosPorDoc = {};
      for (let i = 1; i < datos.length; i++) {
        const row = datos[i];
        const doc = row[idxDoc] ? row[idxDoc].toString() : '';
        const nombre = row[idxNombre] ? row[idxNombre].toString() : '';
        if (
          (filtro.documento && doc.toLowerCase().includes(filtro.documento.toLowerCase())) ||
          (filtro.nombre && nombre.toLowerCase().includes(filtro.nombre.toLowerCase()))
        ) {
          ultimosPorDoc[doc] = row;
        }
        if (!filtro.documento && !filtro.nombre) {
          ultimosPorDoc[doc] = row;
        }
      }
      Object.values(ultimosPorDoc).forEach(row => {
        turnos.push({
          nombre: row[idxNombre] || '',
          documento: row[idxDoc] || '',
          cargo: row[idxCargo] || '',
          area: row[idxArea] || '',
          lunes: row[idxTurno] || '',
          martes: row[idxTurno] || '',
          miercoles: row[idxTurno] || '',
          jueves: row[idxTurno] || '',
          viernes: row[idxTurno] || '',
          sabado: row[idxTurno] || '',
          domingo: row[idxTurno] || '',
          periodo: (row[idxFechaInicio] && row[idxFechaFin]) ? `${row[idxFechaInicio]} - ${row[idxFechaFin]}` : '',
          comentario: row[idxComentario] || '',
          tipo: 'Rango'
        });
      });
    }

    // --- Procesar turnos anuales ---
    const hojaAnual = ss.getSheetByName('turnos_anuales');
    if (hojaAnual) {
      const datos = hojaAnual.getDataRange().getValues();
      const headers = datos[0];
      const idxNombre = headers.indexOf('Empleado_Nombre');
      const idxDoc = headers.indexOf('Empleado_Doc');
      const idxCargo = headers.indexOf('Empleado_Cargo');
      const idxArea = headers.indexOf('Empleado_Area');
      const idxTurno = headers.indexOf('Turno_Inicio');
      const idxComentario = headers.indexOf('comentario');
      const idxAño = headers.indexOf('Año');

      let ultimosPorDoc = {};
      for (let i = 1; i < datos.length; i++) {
        const row = datos[i];
        const doc = row[idxDoc] ? row[idxDoc].toString() : '';
        const nombre = row[idxNombre] ? row[idxNombre].toString() : '';
        if (
          (filtro.documento && doc.toLowerCase().includes(filtro.documento.toLowerCase())) ||
          (filtro.nombre && nombre.toLowerCase().includes(filtro.nombre.toLowerCase()))
        ) {
          ultimosPorDoc[doc] = row;
        }
        if (!filtro.documento && !filtro.nombre) {
          ultimosPorDoc[doc] = row;
        }
      }
      Object.values(ultimosPorDoc).forEach(row => {
        turnos.push({
          nombre: row[idxNombre] || '',
          documento: row[idxDoc] || '',
          cargo: row[idxCargo] || '',
          area: row[idxArea] || '',
          lunes: row[idxTurno] || '',
          martes: row[idxTurno] || '',
          miercoles: row[idxTurno] || '',
          jueves: row[idxTurno] || '',
          viernes: row[idxTurno] || '',
          sabado: row[idxTurno] || '',
          domingo: row[idxTurno] || '',
          periodo: row[idxAño] ? `Año: ${row[idxAño]}` : '',
          comentario: row[idxComentario] || '',
          tipo: 'Anual'
        });
      });
    }

    return turnos;
  } catch (error) {
    console.error('Error en getPublicShifts:', error);
    return [];
  }
}
// ...existing code...
// NEW FUNCTION: Get Brigadistas Data
function getBrigadistasData() {
  try {
    const spreadsheet = SpreadsheetApp.openById(BRIGADAS_DORIA_SHEET_ID);
    const sheet = spreadsheet.getSheetByName(BRIGADISTAS_SHEET_NAME);
    if (!sheet) return { error: true, message: `Hoja "${BRIGADISTAS_SHEET_NAME}" no encontrada.` };
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    const headers = data[1];
    const brigadistas = [];
    const headerMap = {
      'CEDULA': -1,
      'NOMBRE': -1,
      'PROCESO': -1,
      'ALTURAS Y CONFINADOS': -1,
      'TELÉFONO': -1,
      'M': -1
    };
    headers.forEach((header, index) => {
      const normalizedHeader = header.toString().trim().toUpperCase();
      if (headerMap.hasOwnProperty(normalizedHeader)) {
        headerMap[normalizedHeader] = index;
      }
    });

    // --- Cargar turnos semanales ---
    const ssTurnos = SpreadsheetApp.openById(EMPLOYEES_SHEET_ID);
    const sheetTurnos = ssTurnos.getSheetByName('turnos_semanales');
    let turnosData = [];
    if (sheetTurnos) {
      turnosData = sheetTurnos.getDataRange().getValues();
    }
    // Mapear última programación semanal por cédula
    const ultimosTurnosPorCedula = {};
    if (turnosData.length > 1) {
      const turnosHeaders = turnosData[0];
      const idxDoc = turnosHeaders.indexOf('Empleado_Doc');
      const dias = ['Lunes', 'Martes', 'Miercoles', 'Jueves', 'Viernes', 'Sabado', 'Domingo'];
      const idxDias = dias.map(d => turnosHeaders.indexOf(d));
      for (let i = 1; i < turnosData.length; i++) {
        const row = turnosData[i];
        const doc = row[idxDoc] ? row[idxDoc].toString() : '';
        if (doc) {
          ultimosTurnosPorCedula[doc] = idxDias.map(idx => row[idx] ? row[idx].toString() : '');
        }
      }
    }

    for (let i = 2; i < data.length; i++) {
      const row = data[i];
      let genero = 'N/A';
      if (headerMap['M'] !== -1 && row[headerMap['M']] !== undefined) {
        const valorM = row[headerMap['M']].toString().trim().toUpperCase();
        genero = valorM === 'M' ? 'M' : (valorM === 'F' ? 'F' : 'N/A');
      }
      const cedula = headerMap['CEDULA'] !== -1 && row[headerMap['CEDULA']] !== undefined ? row[headerMap['CEDULA']].toString() : '';
      // Obtener turnos de la semana para este brigadista
      const turnosSemana = ultimosTurnosPorCedula[cedula] || ["", "", "", "", "", "", ""];
      const brigadista = {
        cedula: cedula,
        nombre: headerMap['NOMBRE'] !== -1 && row[headerMap['NOMBRE']] !== undefined ? row[headerMap['NOMBRE']].toString() : '',
        proceso: headerMap['PROCESO'] !== -1 && row[headerMap['PROCESO']] !== undefined ? row[headerMap['PROCESO']].toString() : '',
        alturasConfinados: headerMap['ALTURAS Y CONFINADOS'] !== -1 && row[headerMap['ALTURAS Y CONFINADOS']] !== undefined ? row[headerMap['ALTURAS Y CONFINADOS']].toString() : '',
        telefono: headerMap['TELÉFONO'] !== -1 && row[headerMap['TELÉFONO']] !== undefined ? row[headerMap['TELÉFONO']].toString() : '',
        genero: genero,
        turnos: {
          lunes: turnosSemana[0] || "",
          martes: turnosSemana[1] || "",
          miercoles: turnosSemana[2] || "",
          jueves: turnosSemana[3] || "",
          viernes: turnosSemana[4] || "",
          sabado: turnosSemana[5] || "",
          domingo: turnosSemana[6] || ""
        }
      };
      brigadistas.push(brigadista);
    }
    return brigadistas;
  } catch (error) {
    Logger.log(`Error en getBrigadistasData: ${error.message}`);
    return { error: true, message: error.message };
  }
}

// Helper function to capitalize first letter
function capitalizeFirstLetter(string) {
return string.charAt(0).toUpperCase() + string.slice(1);
}

// Función para obtener áreas permitidas según el rol del usuario
function obtenerAreasPermitidas(email) {
  try {
    if (!email || typeof email !== 'string') {
      console.error('Email inválido:', email);
      return ['Sin acceso'];
    }

    const userInfo = obtenerInfoUsuario(email.toString());
    if (!userInfo || !userInfo.success) {
      console.error('No se pudo obtener información del usuario:', email);
      return ['Sin acceso'];
    }

    // Mapeo completo y detallado de roles a áreas
    const areasPorRol = {
      'COMPLETO': [
        'todos',
        'materias-primas',
        'molino',
        'cedi',
        'empaque',
        'pastificio',
        'calidad',
        'sst',
        'info-manufactura',
        'almacen-general',
        'ingenieria-montaje',
        'investigacion-desarrollo',
        'tecnico-electricista',
        'tecnico-lider',
        'soporte-adtvo',
        'maquinista',
        'lider-operaciones',
        'soporte-servicios-administrativos',
        'ayudante-produccion',
        'empacador',
        'operario-montacarga',
        'jefe-ingenieria',
        'coordinador-distribucion',
        'analista-control-gestion',
        'produccion',
        'gestion-humana',
        'mercadeo',
        'planeacion',
        'gestion-ambiental',
        'tecnologia',
        'mantenimiento',
        'gestion-comercial',
        'gestion-integral',
        'presidencia'
      ],
      'JEFE DE PRODUCCION': [
        'materias-primas',
        'molino',
        'empaque',
        'pastificio',
        'produccion',
        'ayudante-produccion',
        'maquinista'
      ],
      'JEFE DE TURNO MOLINO': [
        'molino',
        'maquinista',
        'ayudante-produccion'
      ],
      'JEFE DE TURNO PASTIFICIO': [
        'pastificio',
        'maquinista',
        'ayudante-produccion'
      ],
      'COORDINADOR(A) ALMACEN MATERIAS PRIMAS': [
        'materias-primas',
        'almacen-general',
        'operario-montacarga'
      ],
      'COORDINADOR/A ALMACÉN GENERAL': [
        'almacen-general',
        'operario-montacarga'
      ],
      'COORDINADOR/A CENTRO DE DISTRIBUCIÓN-CEDI': [
        'cedi',
        'operario-montacarga',
        'coordinador-distribucion'
      ],
      'JEFE MECANICO DE EMPAQUE': [
        'empaque',
        'mantenimiento',
        'tecnico-electricista',
        'tecnico-lider'
      ],
      'JEFE/A DE INGENIERÍA': [
        'ingenieria-montaje',
        'mantenimiento',
        'tecnico-electricista',
        'tecnico-lider'
      ],
      'LÍDER DE OPERACIONES': [
        'produccion',
        'materias-primas',
        'molino',
        'empaque',
        'pastificio',
        'maquinista',
        'ayudante-produccion'
      ],
      'ANALISTA SEGURIDAD Y SALUD EN EL TRABAJO': ['sst', 'SST'],
      'AUXILIAR SEGURIDAD Y SALUD EN EL TRABAJO': ['sst', 'SST'],
      'COORDINADOR INFORMACION MANUFACTURA': ['info-manufactura'],
      'DIR.DLLO HUMANO Y SEG. Y SALUD EN TRABAJ': ['sst', 'gestion-humana'],
      'JEFE/A DE ASEGURAMIENTO DE CALIDAD': ['calidad'],
      'JEFE SERVICIOS ADMINISTRATIVOS': ['soporte-servicios-administrativos', 'soporte-adtvo'],
    };

    const rol = userInfo.cargo ? userInfo.cargo.toUpperCase() : 'SIN ROL';
    
    // Si el rol existe en el mapeo, devolver sus áreas permitidas
    // Si no, devolver solo el área específica del rol si existe
    return areasPorRol[rol] || [obtenerAreaPorCargo(rol)];

  } catch (error) {
    console.error('Error en obtenerAreasPermitidas:', error);
    return ['Error'];
  }
}

function obtenerAreaPorCargo(cargo) {
  // Convertir cargo a minúsculas para comparación
  const cargoLower = cargo.toLowerCase();
  
  // Iterar sobre el mapeo de áreas y sus cargos relacionados
  for (const [area, cargos] of Object.entries(getKeywordsForArea())) {
    if (cargos.some(c => cargoLower.includes(c.toLowerCase()))) {
      return area;
    }
  }
  
  return 'Sin acceso';
}
// Agregar después de la función obtenerAreasPermitidas:

// Función para verificar si un usuario tiene acceso a un área específica
function tieneAccesoArea(area, cargoUsuario) {
  // Si el cargo es COMPLETO, tiene acceso a todas las áreas
  if (cargoUsuario === 'COMPLETO') return true;

  // Mapeado completo de cargos y sus áreas permitidas
  const permisosAreas = {
    'ANALISTA SEGURIDAD Y SALUD EN EL TRABAJO': ['SST'],
    'AUXILIAR SEGURIDAD Y SALUD EN EL TRABAJO': ['SST'],
    'COORDINADOR INFORMACION MANUFACTURA': ['Info Manufactura'],
    'COORDINADOR(A) ALMACEN MATERIAS PRIMAS': ['Materias Primas'],
    'COORDINADOR/A ALMACÉN GENERAL': ['Almacén General'],
    'COORDINADOR/A CENTRO DE DISTRIBUCIÓN-CEDI': ['CEDI'],
    'COORDINADOR/A TURNO MOLINO': ['Molino'],
    'DIR.DLLO HUMANO Y SEG. Y SALUD EN TRABAJ': ['SST', 'Gestión Humana'],
    'JEFE DE TURNO MOLINO': ['Molino'],
    'JEFE DE TURNO PASTIFICIO': ['Pastificio'],
    'JEFE MECANICO DE EMPAQUE': ['Empaque', 'Mantenimiento'],
    'JEFE MOLINERO': ['Molino'],
    'JEFE SERVICIOS ADMINISTRATIVOS': ['Soporte Servicios Administrativos'],
    'JEFE/A DE ASEGURAMIENTO DE CALIDAD': ['Calidad'],
    'JEFE/A DE INGENIERÍA': ['Ingeniería y Montaje', 'Mantenimiento'],
    'JEFE/A DE PRODUCCIÓN': ['Producción'],
    'LÍDER DE OPERACIONES': ['Líder de Operaciones', 'Producción'],
    'LÍDER DE SALUD Y SEGURIDAD': ['SST']
  };

  // Si el cargo no está en el mapeado, no tiene acceso
  if (!permisosAreas[cargoUsuario]) return false;

  // Verificar si el área está en las áreas permitidas para el cargo
  return permisosAreas[cargoUsuario].some(areaPermitida => 
    area.toLowerCase().includes(areaPermitida.toLowerCase())
  );
}
// Función auxiliar para verificar si un turno coincide con el patrón
function verificarTurnoCoincide(row, turnoPattern, tipoHoja) {
  try {
    switch (tipoHoja) {
      case 'turnos_semanales':
        // Para turnos semanales, verificar cada día
        const diasIndices = [6, 7, 8, 9, 10, 11, 12]; // índices de Lunes a Domingo
        return diasIndices.some(idx => {
          const turno = row[idx] ? row[idx].toString() : '';
          return turno.includes(turnoPattern);
        });
      
      case 'turnos_rango':
        const turnoRangoIdx = 7; // índice de la columna Turno_Asignado
        const turnoRango = row[turnoRangoIdx] ? row[turnoRangoIdx].toString() : '';
        return turnoRango.includes(turnoPattern);
      
      case 'turnos_anuales':
        const turnoAnualIdx = 7; // índice de la columna Turno_Inicio
        const turnoAnual = row[turnoAnualIdx] ? row[turnoAnualIdx].toString() : '';
        return turnoAnual.includes(turnoPattern);
      
      default:
        return false;
    }
  } catch (error) {
    Logger.log(`Error en verificarTurnoCoincide: ${error}`);
    return false;
  }
}

// Función auxiliar para obtener el período según el tipo de programación
function obtenerPeriodo(row, tipoHoja) {
  try {
    switch (tipoHoja) {
      case 'turnos_semanales':
        return `Semana: ${row[5]}`; // índice de Semana_Inicio
      
      case 'turnos_rango':
        return `${row[5]} - ${row[6]}`; // índices de Fecha_Inicio y Fecha_Fin
      
      case 'turnos_anuales':
        return `Año ${row[5]}`; // índice de Año
      
      default:
        return 'Período no especificado';
    }
  } catch (error) {
    console.error('Error en obtenerPeriodo:', error);
    return 'Error al obtener período';
  }
}

function obtenerPersonalEnTurno(turnoPattern, email) {
  try {
    if (!email || typeof email !== 'string') {
      throw new Error('Email inválido');
    }

    const userInfo = obtenerInfoUsuario(email);
    if (!userInfo.success) {
      throw new Error('Usuario no encontrado');
    }

    const ss = SpreadsheetApp.openById(EMPLOYEES_SHEET_ID);
    const hojas = ['turnos_semanales', 'turnos_rango', 'turnos_anuales'];
    let resultado = [];
    
    const cargo = userInfo.cargo;
    Logger.log(`Usuario: ${email}, Cargo: ${cargo}`);

    // Obtener áreas permitidas para el usuario
    const areasPermitidas = obtenerAreasPermitidas(email);
    const tieneAccesoTotal = cargo === 'COMPLETO' || areasPermitidas.includes('todos');
    
    Logger.log(`Áreas permitidas: ${areasPermitidas.join(', ')}`);

    hojas.forEach(nombreHoja => {
      const hoja = ss.getSheetByName(nombreHoja);
      if (!hoja) return;

      const datos = hoja.getDataRange().getValues();
      const headers = datos[0];
      const indices = obtenerIndicesColumnas(headers);

      // Procesar cada fila de datos
      for (let i = 1; i < datos.length; i++) {
        const row = datos[i];
        const areaEmpleado = (row[indices.area] || '').toString().toLowerCase();
        
        // Normalizar área del empleado para comparación
        const areaNormalizada = areaEmpleado.replace(/\s+/g, '-');
        
        Logger.log(`Verificando fila ${i}: Área empleado: ${areaNormalizada}`);

        // Verificar si el usuario tiene acceso al área
        const tieneAcceso = tieneAccesoTotal || 
                           areasPermitidas.some(area => 
                             area.toLowerCase() === areaNormalizada ||
                             areaNormalizada.includes(area.toLowerCase()));

        if (tieneAcceso) {
          if (verificarTurnoCoincide(row, turnoPattern, nombreHoja, indices)) {
            const programacion = crearObjetoProgramacion(row, indices, nombreHoja);
            resultado.push(programacion);
            Logger.log(`Agregada programación para: ${programacion.nombre}`);
          }
        }
      }
    });

    Logger.log(`Total de programaciones encontradas: ${resultado.length}`);
    return resultado;

  } catch (error) {
    Logger.log(`Error en obtenerPersonalEnTurno: ${error}`);
    throw new Error('Error al obtener personal en turno: ' + error.message);
  }
}

function obtenerIndicesColumnas(headers) {
  return {
    nombre: headers.indexOf('Empleado_Nombre'),
    documento: headers.indexOf('Empleado_Doc'),
    cargo: headers.indexOf('Empleado_Cargo'),
    area: headers.indexOf('Empleado_Area'),
    // Índices específicos para cada tipo de programación
    semanaInicio: headers.indexOf('Semana_Inicio'),
    fechaInicio: headers.indexOf('Fecha_Inicio'),
    fechaFin: headers.indexOf('Fecha_Fin'),
    turnoAsignado: headers.indexOf('Turno_Asignado'),
    año: headers.indexOf('Año'),
    tipoTurno: headers.indexOf('Tipo_Turno'),
    turnoInicio: headers.indexOf('Turno_Inicio'),
    fijo: headers.indexOf('Fijo'),
    // Días de la semana para turnos semanales
    lunes: headers.indexOf('Lunes'),
    martes: headers.indexOf('Martes'),
    miercoles: headers.indexOf('Miercoles'),
    jueves: headers.indexOf('Jueves'),
    viernes: headers.indexOf('Viernes'),
    sabado: headers.indexOf('Sabado'),
    domingo: headers.indexOf('Domingo')
  };
}

function crearObjetoProgramacion(row, indices, tipoHoja) {
  const base = {
    nombre: row[indices.nombre] || '',
    documento: row[indices.documento] || '',
    cargo: row[indices.cargo] || 'Sin cargo',
    area: row[indices.area] || 'Sin área',
    tipo: tipoHoja
  };

  switch (tipoHoja) {
    case 'turnos_semanales':
      base.horario = {
        Lunes: row[indices.lunes] || '-',
        Martes: row[indices.martes] || '-',
        Miercoles: row[indices.miercoles] || '-',
        Jueves: row[indices.jueves] || '-',
        Viernes: row[indices.viernes] || '-',
        Sabado: row[indices.sabado] || '-',
        Domingo: row[indices.domingo] || '-'
      };
      base.dominical = Object.values(base.horario).some(turno => 
        turno.includes('SI DOMINICAL'));
      break;

    case 'turnos_rango':
      base.fechaInicio = formatDate(row[indices.fechaInicio]);
      base.fechaFin = formatDate(row[indices.fechaFin]);
      base.turnoAsignado = row[indices.turnoAsignado] || '';
      base.dominical = base.turnoAsignado.includes('SI DOMINICAL');
      break;

    case 'turnos_anuales':
      base.año = row[indices.año];
      base.tipoTurno = row[indices.tipoTurno] || '';
      base.turnoInicio = row[indices.turnoInicio] || '';
      base.fijo = row[indices.fijo] ? 'Sí' : 'No';
      base.dominical = base.turnoInicio.includes('SI DOMINICAL');
      break;
  }

  return base;
}

function verificarTurnoCoincide(row, turnoPattern, tipoHoja, indices) {
  switch (tipoHoja) {
    case 'turnos_semanales':
      const dias = ['lunes', 'martes', 'miercoles', 'jueves', 'viernes', 'sabado', 'domingo'];
      return dias.some(dia => {
        const turno = row[indices[dia]];
        return turno && turno.includes(turnoPattern);
      });
    
    case 'turnos_rango':
      return row[indices.turnoAsignado] && 
             row[indices.turnoAsignado].includes(turnoPattern);
    
    case 'turnos_anuales':
      return row[indices.turnoInicio] && 
             row[indices.turnoInicio].includes(turnoPattern);
    
    default:
      return false;
  }
}

// Función auxiliar para verificar si un turno coincide con el patrón
function verificarTurnoCoincide(row, turnoPattern, tipoHoja) {
  const turnoIndex = row.indexOf(turnoPattern);
  return turnoIndex !== -1;
}

// Función auxiliar para obtener el período según el tipo de programación
function obtenerPeriodo(row, tipoHoja) {
  switch(tipoHoja) {
    case 'turnos_semanales':
      return `Semana: ${row[2]}`; // Ajusta el índice según tu estructura
    case 'turnos_rango':
      return `${row[2]} - ${row[3]}`; // Ajusta los índices
    case 'turnos_anuales':
      return `Año: ${row[2]}`; // Ajusta el índice
    default:
      return '';
  }
}

// Función auxiliar para obtener el horario
function obtenerHorario(row, headers) {
  const dias = ['Lunes', 'Martes', 'Miercoles', 'Jueves', 'Viernes', 'Sabado', 'Domingo'];
  let horario = {};
  
  dias.forEach(dia => {
    const index = headers.indexOf(dia);
    horario[dia] = index !== -1 ? row[index] : '-';
  });
  
  return horario;
}


// ...existing code...
function obtenerTodosTurnos(areaFiltro = '') {
  try {
    const ss = SpreadsheetApp.openById(EMPLOYEES_SHEET_ID);
    let turnos = [];
    
    console.log('Iniciando obtención de turnos...');

    // --- Procesar turnos semanales ---
    const hojaSemanales = ss.getSheetByName('turnos_semanales');
    if (hojaSemanales) {
      console.log('Procesando turnos_semanales');
      const datos = hojaSemanales.getDataRange().getValues();
      const headers = datos[0];
      console.log('Headers semanales:', headers);
      
      const idxNombre = headers.indexOf('Empleado_Nombre');
      const idxDoc = headers.indexOf('Empleado_Doc');
      const idxCargo = headers.indexOf('Empleado_Cargo');
      const idxArea = headers.indexOf('Empleado_Area');
      const idxComentario = headers.indexOf('comentario'); // CAMBIADO: minúscula
      console.log('Índice comentario semanal:', idxComentario);
      
      const dias = ['Lunes', 'Martes', 'Miercoles', 'Jueves', 'Viernes', 'Sabado', 'Domingo'];
      const idxDias = dias.map(d => headers.indexOf(d));

      let ultimosPorDoc = {};
      for (let i = 1; i < datos.length; i++) {
        const row = datos[i];
        const doc = row[idxDoc] || '';
        if (!doc) continue;
        if (areaFiltro && row[idxArea].toLowerCase() !== areaFiltro.toLowerCase()) continue;
        ultimosPorDoc[doc] = row;
        
        // DEBUG: Mostrar datos de una fila con comentario
        if (idxComentario !== -1 && row[idxComentario]) {
          console.log('Comentario encontrado en semanales - Doc:', doc, 'Comentario:', row[idxComentario]);
        }
      }
      
      Object.values(ultimosPorDoc).forEach(row => {
        const comentario = idxComentario !== -1 ? (row[idxComentario] || '') : '';
        console.log('Agregando semanal - Doc:', row[idxDoc], 'Comentario:', comentario);
        
        turnos.push({
          nombre: row[idxNombre] || '',
          documento: row[idxDoc] || '',
          cargo: row[idxCargo] || '',
          area: row[idxArea] || '',
          lunes: row[idxDias[0]] || '',
          martes: row[idxDias[1]] || '',
          miercoles: row[idxDias[2]] || '',
          jueves: row[idxDias[3]] || '',
          viernes: row[idxDias[4]] || '',
          sabado: row[idxDias[5]] || '',
          domingo: row[idxDias[6]] || '',
          comentario: comentario,
          periodo: 'Semanal'
        });
      });
    }

    // --- Procesar turnos de rango ---
    const hojaRango = ss.getSheetByName('turnos_rango');
    if (hojaRango) {
      console.log('Procesando turnos_rango');
      const datos = hojaRango.getDataRange().getValues();
      const headers = datos[0];
      console.log('Headers rango:', headers);
      
      const idxNombre = headers.indexOf('Empleado_Nombre');
      const idxDoc = headers.indexOf('Empleado_Doc');
      const idxCargo = headers.indexOf('Empleado_Cargo');
      const idxArea = headers.indexOf('Empleado_Area');
      const idxTurno = headers.indexOf('Turno_Asignado');
      const idxComentario = headers.indexOf('comentario'); // CAMBIADO: minúscula
      console.log('Índice comentario rango:', idxComentario);
      
      const idxFechaInicio = headers.indexOf('Fecha_Inicio');
      const idxFechaFin = headers.indexOf('Fecha_Fin');
      
      let ultimosPorDoc = {};
      for (let i = 1; i < datos.length; i++) {
        const row = datos[i];
        const doc = row[idxDoc] || '';
        if (!doc) continue;
        if (areaFiltro && row[idxArea].toLowerCase() !== areaFiltro.toLowerCase()) continue;
        ultimosPorDoc[doc] = row;
        
        // DEBUG: Mostrar datos de una fila con comentario
        if (idxComentario !== -1 && row[idxComentario]) {
          console.log('Comentario encontrado en rango - Doc:', doc, 'Comentario:', row[idxComentario]);
        }
      }
      
      Object.values(ultimosPorDoc).forEach(row => {
        const turnoAsignado = row[idxTurno] || '';
        const comentario = idxComentario !== -1 ? (row[idxComentario] || '') : '';
        console.log('Agregando rango - Doc:', row[idxDoc], 'Comentario:', comentario);
        
        turnos.push({
          nombre: row[idxNombre] || '',
          documento: row[idxDoc] || '',
          cargo: row[idxCargo] || '',
          area: row[idxArea] || '',
          lunes: turnoAsignado,
          martes: turnoAsignado,
          miercoles: turnoAsignado,
          jueves: turnoAsignado,
          viernes: turnoAsignado,
          sabado: turnoAsignado,
          domingo: turnoAsignado,
          comentario: comentario,
          periodo: (row[idxFechaInicio] && row[idxFechaFin]) ? 
                   `${formatDate(row[idxFechaInicio])} - ${formatDate(row[idxFechaFin])}` : ''
        });
      });
    }

    // --- Procesar turnos anuales ---
    const hojaAnual = ss.getSheetByName('turnos_anuales');
    if (hojaAnual) {
      console.log('Procesando turnos_anuales');
      const datos = hojaAnual.getDataRange().getValues();
      const headers = datos[0];
      console.log('Headers anuales:', headers);
      
      const idxNombre = headers.indexOf('Empleado_Nombre');
      const idxDoc = headers.indexOf('Empleado_Doc');
      const idxCargo = headers.indexOf('Empleado_Cargo');
      const idxArea = headers.indexOf('Empleado_Area');
      const idxTurno = headers.indexOf('Turno_Inicio');
      const idxComentario = headers.indexOf('comentario'); // CAMBIADO: minúscula
      console.log('Índice comentario anual:', idxComentario);
      
      const idxAño = headers.indexOf('Año');
      
      let ultimosPorDoc = {};
      for (let i = 1; i < datos.length; i++) {
        const row = datos[i];
        const doc = row[idxDoc] || '';
        if (!doc) continue;
        if (areaFiltro && row[idxArea].toLowerCase() !== areaFiltro.toLowerCase()) continue;
        ultimosPorDoc[doc] = row;
        
        // DEBUG: Mostrar datos de una fila con comentario
        if (idxComentario !== -1 && row[idxComentario]) {
          console.log('Comentario encontrado en anuales - Doc:', doc, 'Comentario:', row[idxComentario]);
        }
      }
      
      Object.values(ultimosPorDoc).forEach(row => {
        const turnoAsignado = row[idxTurno] || '';
        const comentario = idxComentario !== -1 ? (row[idxComentario] || '') : '';
        console.log('Agregando anual - Doc:', row[idxDoc], 'Comentario:', comentario);
        
        turnos.push({
          nombre: row[idxNombre] || '',
          documento: row[idxDoc] || '',
          cargo: row[idxCargo] || '',
          area: row[idxArea] || '',
          lunes: turnoAsignado,
          martes: turnoAsignado,
          miercoles: turnoAsignado,
          jueves: turnoAsignado,
          viernes: turnoAsignado,
          sabado: turnoAsignado,
          domingo: turnoAsignado,
          comentario: comentario,
          periodo: row[idxAño] ? `Año: ${row[idxAño]}` : ''
        });
      });
    }

    console.log('Total de turnos obtenidos:', turnos.length);
    turnos.forEach((turno, index) => {
      if (turno.comentario) {
        console.log(`Turno ${index + 1} CON COMENTARIO:`, {
          documento: turno.documento,
          comentario: turno.comentario,
          tieneComentario: !!turno.comentario
        });
      }
    });

    return turnos;
  } catch (error) {
    console.error('Error al obtener turnos:', error);
    throw new Error('Error al obtener turnos: ' + error.message);
  }
}

// Función auxiliar para formatear fechas
function formatDate(date) {
  if (!date) return '';
  if (typeof date === 'object' && date instanceof Date) {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  return date;
}
// ...existing code...


// Función para obtener los días festivos de Colombia
function obtenerDiasFestivos() {
  return {
    "2024": [
      "2024-01-01", "2024-01-08", "2024-03-25", "2024-03-28", "2024-03-29",
      "2024-05-01", "2024-05-13", "2024-06-03", "2024-06-10", "2024-07-01",
      "2024-07-20", "2024-08-07", "2024-08-19", "2024-10-14", "2024-11-04",
      "2024-11-11", "2024-12-08", "2024-12-25"
    ],
    "2025": [
      "2025-01-01", "2025-01-06", "2025-03-24", "2025-04-17", "2025-04-18",
      "2025-05-01", "2025-06-02", "2025-06-23", "2025-06-30", "2025-07-20",
      "2025-08-07", "2025-08-18", "2025-10-13", "2025-11-03", "2025-11-17",
      "2025-12-08", "2025-12-25"
    ]
  };
}

