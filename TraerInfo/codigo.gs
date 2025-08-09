// ============= CONFIGURACIÓN =============
// IDs de tus archivos en Google Drive
const ARCHIVO_ORIGEN_ID = '1pR41SR6Yz8NCN73VC3ciqGtpNvLkG43aDxlR-ePclGI';  // Formato Nuevos - Inquebrantables
const ARCHIVO_DESTINO_ID = '1rZyjLYeuSUA5bWfKvWIyi-c091mBtX4T2An7t3bVJZU';  // Birthdays

// Nombres de las hojas (ajusta si es necesario)
const HOJA_ORIGEN = 'Hoja1';  // Del archivo "Formato Nuevos - Inquebrantables"
const HOJA_DESTINO = 'Worksheet';  // Del archivo "Birthdays"

// Configuración para archivos en unidades compartidas
const ARCHIVO_ORIGEN_ES_COMPARTIDO = true;  // true porque está en unidad compartida
const SOLO_LECTURA_ORIGEN = true;  // true porque solo tienes permisos de lectura

// ============= FUNCIÓN PRINCIPAL =============
function procesarDatos() {
  try {
    console.log('🚀 Iniciando análisis de datos...');
    
    // Abrir archivos
    const archivoOrigen = DriveApp.getFileById(ARCHIVO_ORIGEN_ID);
    const archivoDestino = DriveApp.getFileById(ARCHIVO_DESTINO_ID);
    
    // Convertir a Google Sheets temporalmente si son archivos Excel
    const hojaOrigen = obtenerHojaDeExcel(archivoOrigen, HOJA_ORIGEN, ARCHIVO_ORIGEN_ES_COMPARTIDO);
    const hojaDestino = obtenerHojaDeExcel(archivoDestino, HOJA_DESTINO, false);
    
    // Obtener datos
    const datosOrigen = obtenerDatosOrigen(hojaOrigen);
    const datosDestino = obtenerDatosDestino(hojaDestino);
    
    console.log(`📊 Datos encontrados - Origen: ${datosOrigen.length}, Destino: ${datosDestino.length}`);
    
    // VALIDACIÓN Y ANÁLISIS DETALLADO
    const analisis = analizarDatos(datosOrigen, datosDestino);
    
    // Mostrar reporte completo
    mostrarReporteDetallado(analisis);
    
    // Copiar automáticamente si hay registros válidos
    if (analisis.paraCopiAr.length > 0) {
      agregarDatosADestino(hojaDestino, analisis.paraCopiAr);
      console.log('🎉 ¡Proceso completado exitosamente!');
    } else {
      console.log('ℹ️ No hay datos nuevos válidos para copiar.');
    }
    
    // Mostrar resumen final
    mostrarResumen(analisis);
    
  } catch (error) {
    console.error('❌ Error durante el proceso:', error);
    throw error;
  }
}

// ============= FUNCIONES AUXILIARES =============

/**
 * Obtiene una hoja de un archivo Excel o Google Sheets
 * Maneja archivos en unidades compartidas con permisos limitados
 */
function obtenerHojaDeExcel(archivo, nombreHoja, esCompartido = false) {
  try {
    // Intentar abrir como Google Sheets primero
    const spreadsheet = SpreadsheetApp.openById(archivo.getId());
    return spreadsheet.getSheetByName(nombreHoja) || spreadsheet.getSheets()[0];
  } catch (e) {
    console.log(`📝 El archivo parece ser Excel o no tenemos permisos directos: ${archivo.getName()}`);
    
    // Para archivos compartidos con solo lectura, crear una copia en nuestro Drive
    if (esCompartido || SOLO_LECTURA_ORIGEN) {
      console.log(`📂 Creando copia temporal del archivo compartido...`);
      
      try {
        // Crear copia en nuestro Drive personal
        const copiaTemp = archivo.makeCopy(`TEMP_${archivo.getName()}_${Date.now()}`);
        
        // Convertir a Google Sheets si es necesario
        let spreadsheetTemp;
        try {
          spreadsheetTemp = SpreadsheetApp.openById(copiaTemp.getId());
        } catch (conversionError) {
          // Si no se puede abrir directamente, convertir usando Drive API
          const archivoConvertido = Drive.Files.copy({
            title: `TEMP_CONVERTED_${archivo.getName()}_${Date.now()}`
          }, copiaTemp.getId(), {
            convert: true
          });
          
          // Eliminar la copia no convertida
          copiaTemp.setTrashed(true);
          
          spreadsheetTemp = SpreadsheetApp.openById(archivoConvertido.id);
          
          // Programar eliminación del archivo convertido después del procesamiento
          Utilities.sleep(1000); // Dar tiempo para que se procese
        }
        
        const hoja = spreadsheetTemp.getSheetByName(nombreHoja) || spreadsheetTemp.getSheets()[0];
        
        // Programar eliminación de archivos temporales (después de 30 segundos)
        setTimeout(() => {
          try {
            if (copiaTemp) copiaTemp.setTrashed(true);
            console.log('🗑️ Archivo temporal eliminado correctamente');
          } catch (e) {
            console.warn('⚠️ No se pudo eliminar archivo temporal automáticamente');
          }
        }, 30000);
        
        return hoja;
        
      } catch (copyError) {
        console.error('❌ Error al crear copia del archivo:', copyError);
        throw new Error(`No se puede acceder al archivo compartido: ${copyError.message}`);
      }
    }
    
    // Método original para archivos locales
    console.log(`📝 Convirtiendo archivo Excel local: ${archivo.getName()}`);
    const copiaTemp = Drive.Files.copy({
      title: `TEMP_${archivo.getName()}_${Date.now()}`
    }, archivo.getId(), {
      convert: true
    });
    
    const spreadsheetTemp = SpreadsheetApp.openById(copiaTemp.id);
    const hoja = spreadsheetTemp.getSheetByName(nombreHoja) || spreadsheetTemp.getSheets()[0];
    
    // Programar eliminación de archivo temporal
    setTimeout(() => {
      try {
        DriveApp.getFileById(copiaTemp.id).setTrashed(true);
      } catch (e) {
        console.warn('No se pudo eliminar archivo temporal:', copiaTemp.id);
      }
    }, 30000);
    
    return hoja;
  }
}

/**
 * Obtiene y procesa los datos del archivo origen
 */
function obtenerDatosOrigen(hoja) {
  const datos = hoja.getDataRange().getValues();
  const headers = datos[0];
  
  // Encontrar índices de columnas importantes
  const indiceNombre = headers.indexOf('Nombre');
  const indiceFechaNac = headers.indexOf('Fecha de Nacimientos');
  
  if (indiceNombre === -1) {
    throw new Error('No se encontró la columna "Nombre" en el archivo origen');
  }
  
  const datosProcessed = [];
  
  // Procesar cada fila (saltando headers)
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const nombreCompleto = fila[indiceNombre];
    const fechaNacimiento = fila[indiceFechaNac];
    
    // Saltar filas vacías o sin nombre
    if (!nombreCompleto || nombreCompleto.toString().trim() === '') {
      continue;
    }
    
    // Separar nombre y apellido
    const { nombre, apellido } = separarNombreApellido(nombreCompleto);
    
    // Procesar fecha de nacimiento
    const cumpleanos = procesarFechaNacimiento(fechaNacimiento);
    
    datosProcessed.push({
      nombre: nombre,
      apellido: apellido,
      cumpleanos: cumpleanos,
      nombreCompleto: nombreCompleto.toString().trim()
    });
  }
  
  return datosProcessed;
}

/**
 * Separa nombre completo en nombre y apellido de manera más inteligente
 */
function separarNombreApellido(nombreCompleto) {
  const partes = nombreCompleto.toString().trim().split(' ');
  
  if (partes.length === 1) {
    return { nombre: partes[0], apellido: '' };
  } else if (partes.length === 2) {
    return { nombre: partes[0], apellido: partes[1] };
  } else if (partes.length === 3) {
    // Para 3 partes, tomar las primeras 2 como nombre y la última como apellido
    // Ejemplo: "Eloy David wilches" → Nombre: "Eloy David", Apellido: "wilches"
    const nombre = partes.slice(0, 2).join(' ');
    const apellido = partes[2];
    return { nombre: nombre, apellido: apellido };
  } else {
    // Para más de 3 partes, tomar las primeras 2 como nombre y el resto como apellido
    const nombre = partes.slice(0, 2).join(' ');
    const apellido = partes.slice(2).join(' ');
    return { nombre: nombre, apellido: apellido };
  }
}

/**
 * Procesa fecha de nacimiento y la convierte a formato MM/DD
 */
function procesarFechaNacimiento(fechaNacimiento) {
  if (!fechaNacimiento) {
    return '';
  }
  
  let fecha;
  
  // Si es un número (formato serial de Excel)
  if (typeof fechaNacimiento === 'number') {
    // Convertir número serial de Excel a fecha
    fecha = new Date((fechaNacimiento - 25569) * 86400 * 1000);
  } 
  // Si es una fecha
  else if (fechaNacimiento instanceof Date) {
    fecha = fechaNacimiento;
  }
  // Si es string, intentar parsear
  else if (typeof fechaNacimiento === 'string') {
    fecha = new Date(fechaNacimiento);
  }
  
  // Verificar si la fecha es válida
  if (!fecha || isNaN(fecha.getTime())) {
    return '';
  }
  
  // Formatear como MM/DD
  const mes = (fecha.getMonth() + 1).toString().padStart(2, '0');
  const dia = fecha.getDate().toString().padStart(2, '0');
  
  return `${mes}/${dia}`;
}

/**
 * Obtiene los datos existentes del archivo destino
 */
function obtenerDatosDestino(hoja) {
  const datos = hoja.getDataRange().getValues();
  const datosExistentes = [];
  
  // Saltar header (fila 0)
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    if (fila[0] && fila[1]) { // Si tiene nombre y apellido
      datosExistentes.push({
        nombre: fila[0].toString().trim().toLowerCase(),
        apellido: fila[1].toString().trim().toLowerCase(),
        cumpleanos: fila[2] ? fila[2].toString() : ''
      });
    }
  }
  
  return datosExistentes;
}

/**
 * Analiza los datos y categoriza las personas según los criterios especificados
 */
function analizarDatos(datosOrigen, datosDestino) {
  console.log('\n🔍 Iniciando análisis detallado...');
  
  const yaExisten = [];           // Personas que ya están en el archivo destino
  const nuevasSinFecha = [];      // Personas nuevas pero sin fecha de nacimiento
  const paraCopiAr = [];          // Personas nuevas con fecha de nacimiento (listas para copiar)
  
  for (const persona of datosOrigen) {
    // Verificar si ya existe usando similitud
    const personaExistente = datosDestino.find(existente => 
      sonSimilares(persona.nombre, persona.apellido, existente.nombre, existente.apellido)
    );
    
    if (personaExistente) {
      // Ya existe en el archivo destino
      yaExisten.push({
        ...persona,
        coincideCon: `${personaExistente.nombre} ${personaExistente.apellido}`.trim()
      });
    } else {
      // No existe en el archivo destino
      if (!persona.cumpleanos || persona.cumpleanos === '') {
        // Sin fecha de nacimiento - solo listar para revisión
        nuevasSinFecha.push(persona);
      } else {
        // Con fecha de nacimiento - listo para copiar
        paraCopiAr.push(persona);
      }
    }
  }
  
  return {
    yaExisten,
    nuevasSinFecha,
    paraCopiAr,
    totalOrigen: datosOrigen.length,
    totalDestino: datosDestino.length
  };
}

/**
 * Muestra un reporte detallado del análisis
 */
function mostrarReporteDetallado(analisis) {
  console.log('\n' + '='.repeat(60));
  console.log('📋 REPORTE DETALLADO DE ANÁLISIS');
  console.log('='.repeat(60));
  
  console.log(`\n📊 RESUMEN GENERAL:`);
  console.log(`   • Total registros en origen: ${analisis.totalOrigen}`);
  console.log(`   • Total registros en destino: ${analisis.totalDestino}`);
  console.log(`   • Personas que ya existen: ${analisis.yaExisten.length}`);
  console.log(`   • Personas nuevas sin fecha: ${analisis.nuevasSinFecha.length}`);
  console.log(`   • Personas listas para copiar: ${analisis.paraCopiAr.length}`);
  
  // PERSONAS QUE YA EXISTEN
  if (analisis.yaExisten.length > 0) {
    console.log(`\n✅ PERSONAS QUE YA EXISTEN (${analisis.yaExisten.length}) - OMITIDAS:`);
    console.log('-'.repeat(50));
    analisis.yaExisten.forEach((persona, index) => {
      console.log(`${index + 1}. "${persona.nombreCompleto}" → coincide con "${persona.coincideCon}"`);
    });
  }
  
  // PERSONAS NUEVAS SIN FECHA DE NACIMIENTO
  if (analisis.nuevasSinFecha.length > 0) {
    console.log(`\n⚠️ PERSONAS NUEVAS SIN FECHA DE NACIMIENTO (${analisis.nuevasSinFecha.length}) - REQUIEREN REVISIÓN:`);
    console.log('-'.repeat(50));
    analisis.nuevasSinFecha.forEach((persona, index) => {
      console.log(`${index + 1}. "${persona.nombreCompleto}" - Sin fecha de nacimiento`);
    });
  }
  
  // PERSONAS LISTAS PARA COPIAR
  if (analisis.paraCopiAr.length > 0) {
    console.log(`\n🎂 PERSONAS LISTAS PARA COPIAR (${analisis.paraCopiAr.length}):`);
    console.log('-'.repeat(50));
    analisis.paraCopiAr.forEach((persona, index) => {
      console.log(`${index + 1}. ${persona.nombre} ${persona.apellido} - ${persona.cumpleanos}`);
    });
  }
  
  console.log('\n' + '='.repeat(60));
  
  // Mostrar alertas si hay casos que requieren atención
  if (analisis.nuevasSinFecha.length > 0) {
    console.log(`\n🔔 ATENCIÓN: Hay ${analisis.nuevasSinFecha.length} personas sin fecha de nacimiento que requieren revisión manual.`);
  }
  
  if (analisis.paraCopiAr.length === 0) {
    console.log(`\n💡 INFO: No hay personas nuevas con fecha de nacimiento para copiar.`);
  }
}

/**
 * Compara si dos personas son similares (nombres parecidos)
 */
function sonSimilares(nombre1, apellido1, nombre2, apellido2) {
  // Normalizar strings (minúsculas, sin acentos, sin espacios extra)
  const normalizar = (str) => {
    return str.toString()
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '') // Remover acentos
      .replace(/[^a-z0-9]/g, '') // Solo letras y números
      .trim();
  };
  
  const n1 = normalizar(nombre1);
  const a1 = normalizar(apellido1);
  const n2 = normalizar(nombre2);
  const a2 = normalizar(apellido2);
  
  // Verificar coincidencias exactas
  if (n1 === n2 && a1 === a2) {
    return true;
  }
  
  // Verificar si el nombre está contenido en el otro (para nombres compuestos)
  const nombresSimilares = (n1.includes(n2) || n2.includes(n1)) && 
                           Math.abs(n1.length - n2.length) <= 3;
  const apellidosSimilares = (a1.includes(a2) || a2.includes(a1)) && 
                            Math.abs(a1.length - a2.length) <= 5;
  
  if (nombresSimilares && apellidosSimilares) {
    return true;
  }
  
  // Verificar similaridad cruzada (por si los nombres están en diferentes posiciones)
  const todoTexto1 = `${n1} ${a1}`.replace(/\s+/g, '');
  const todoTexto2 = `${n2} ${a2}`.replace(/\s+/g, '');
  
  // Si uno contiene al otro con alta similitud
  if (todoTexto1.includes(todoTexto2) || todoTexto2.includes(todoTexto1)) {
    return true;
  }
  
  // Verificar similaridad usando distancia de Levenshtein
  if (todoTexto1.length >= 5 && todoTexto2.length >= 5) {
    const distancia = calcularDistanciaLevenshtein(todoTexto1, todoTexto2);
    const maxLength = Math.max(todoTexto1.length, todoTexto2.length);
    const similitud = 1 - (distancia / maxLength);
    
    // Considerar similares si la similitud es mayor al 75% (más permisivo)
    return similitud > 0.75;
  }
  
  return false;
}

/**
 * Calcula la distancia de Levenshtein entre dos strings
 */
function calcularDistanciaLevenshtein(str1, str2) {
  const matrix = [];
  
  for (let i = 0; i <= str2.length; i++) {
    matrix[i] = [i];
  }
  
  for (let j = 0; j <= str1.length; j++) {
    matrix[0][j] = j;
  }
  
  for (let i = 1; i <= str2.length; i++) {
    for (let j = 1; j <= str1.length; j++) {
      if (str2.charAt(i - 1) === str1.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(
          matrix[i - 1][j - 1] + 1,
          matrix[i][j - 1] + 1,
          matrix[i - 1][j] + 1
        );
      }
    }
  }
  
  return matrix[str2.length][str1.length];
}

/**
 * Agrega los datos nuevos al archivo destino
 */
function agregarDatosADestino(hoja, datosNuevos) {
  // Encontrar la primera fila vacía
  const ultimaFila = hoja.getLastRow();
  let filaInicio = ultimaFila + 1;
  
  // Preparar datos para insertar
  const filasParaInsertar = datosNuevos.map(persona => [
    persona.nombre,
    persona.apellido,
    persona.cumpleanos
  ]);
  
  // Insertar datos
  if (filasParaInsertar.length > 0) {
    const rango = hoja.getRange(filaInicio, 1, filasParaInsertar.length, 3);
    rango.setValues(filasParaInsertar);
    
    console.log(`✅ Se agregaron ${filasParaInsertar.length} nuevas personas al archivo Birthdays`);
    
    // Mostrar detalles de las personas agregadas
    console.log('\n🎂 Personas agregadas:');
    datosNuevos.forEach((persona, index) => {
      console.log(`${index + 1}. ${persona.nombre} ${persona.apellido} - ${persona.cumpleanos}`);
    });
  }
}

/**
 * Muestra un resumen del proceso
 */
function mostrarResumen(analisis) {
  if (!analisis || typeof analisis !== 'object') {
    console.log('\n📋 PROCESO COMPLETADO');
    return;
  }
  
  console.log('\n📋 RESUMEN FINAL DEL PROCESO:');
  console.log('===============================');
  console.log(`✅ Personas copiadas: ${analisis.paraCopiAr ? analisis.paraCopiAr.length : 0}`);
  console.log(`⚠️ Personas sin fecha (requieren revisión): ${analisis.nuevasSinFecha ? analisis.nuevasSinFecha.length : 0}`);
  console.log(`🔍 Personas que ya existían (omitidas): ${analisis.yaExisten ? analisis.yaExisten.length : 0}`);
  console.log('\n✅ Proceso completado exitosamente!');
}

// ============= FUNCIÓN DE SOLO ANÁLISIS (SIN COPIA) =============
/**
 * Función para hacer solo el análisis sin copiar nada
 * Útil para revisar los datos antes de hacer la copia real
 */
function soloAnalizar() {
  try {
    console.log('🔍 Iniciando solo análisis (sin copiar datos)...');
    
    // Abrir archivos
    const archivoOrigen = DriveApp.getFileById(ARCHIVO_ORIGEN_ID);
    const archivoDestino = DriveApp.getFileById(ARCHIVO_DESTINO_ID);
    
    // Convertir a Google Sheets temporalmente si son archivos Excel
    const hojaOrigen = obtenerHojaDeExcel(archivoOrigen, HOJA_ORIGEN, ARCHIVO_ORIGEN_ES_COMPARTIDO);
    const hojaDestino = obtenerHojaDeExcel(archivoDestino, HOJA_DESTINO, false);
    
    // Obtener datos
    const datosOrigen = obtenerDatosOrigen(hojaOrigen);
    const datosDestino = obtenerDatosDestino(hojaDestino);
    
    console.log(`📊 Datos encontrados - Origen: ${datosOrigen.length}, Destino: ${datosDestino.length}`);
    
    // SOLO ANÁLISIS - NO COPIA
    const analisis = analizarDatos(datosOrigen, datosDestino);
    
    // Mostrar reporte completo
    mostrarReporteDetallado(analisis);
    
    console.log('\n💡 TIP: Si los resultados se ven bien, ejecuta procesarDatos() para hacer la copia real.');
    
  } catch (error) {
    console.error('❌ Error durante el análisis:', error);
    throw error;
  }
}

// ============= FUNCIÓN DE PRUEBA =============
/**
 * Función para probar la configuración antes de ejecutar
 */
function probarConfiguracion() {
  try {
    console.log('🔍 Probando configuración...');
    
    // Verificar acceso a archivos
    const archivoOrigen = DriveApp.getFileById(ARCHIVO_ORIGEN_ID);
    const archivoDestino = DriveApp.getFileById(ARCHIVO_DESTINO_ID);
    
    console.log(`✅ Archivo origen encontrado: ${archivoOrigen.getName()}`);
    console.log(`📂 Ubicación origen: ${ARCHIVO_ORIGEN_ES_COMPARTIDO ? 'Unidad compartida' : 'Drive personal'}`);
    console.log(`🔒 Permisos origen: ${SOLO_LECTURA_ORIGEN ? 'Solo lectura' : 'Lectura/escritura'}`);
    console.log(`✅ Archivo destino encontrado: ${archivoDestino.getName()}`);
    
    // Verificar si podemos acceder al contenido
    try {
      const hojaOrigen = obtenerHojaDeExcel(archivoOrigen, HOJA_ORIGEN, ARCHIVO_ORIGEN_ES_COMPARTIDO);
      console.log(`✅ Acceso a hoja origen: ${hojaOrigen.getName()}`);
      
      const hojaDestino = obtenerHojaDeExcel(archivoDestino, HOJA_DESTINO, false);
      console.log(`✅ Acceso a hoja destino: ${hojaDestino.getName()}`);
      
      console.log('🎉 Configuración correcta. Puedes ejecutar procesarDatos()');
      
    } catch (accessError) {
      console.error('❌ Error de acceso a las hojas:', accessError);
      console.log('\n📝 Posibles soluciones:');
      console.log('1. Verificar que los nombres de las hojas sean correctos');
      console.log('2. Asegurar permisos de acceso a archivos compartidos');
      console.log('3. Verificar que el script tenga permisos para Drive API');
    }
    
  } catch (error) {
    console.error('❌ Error en la configuración:', error);
    console.log('\n📝 Asegúrate de:');
    console.log('1. Reemplazar los IDs de archivo en la sección CONFIGURACIÓN');
    console.log('2. Configurar correctamente las variables de archivo compartido');
    console.log('3. Dar permisos de acceso a Google Drive y Drive API');
    console.log('4. Verificar que los archivos existan y tengas acceso a ellos');
    console.log('\n🔧 Para obtener IDs de archivos:');
    console.log('   - Abre el archivo en Drive');
    console.log('   - Copia el ID de la URL: drive.google.com/file/d/[ID_AQUI]/view');
  }
}