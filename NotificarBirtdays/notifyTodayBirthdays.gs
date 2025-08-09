/**
 * ========================================
 * PROYECTO: NOTIFICADOR DE CUMPLEAÑOS HOY
 * ========================================
 * 
 * Este proyecto envía notificaciones de Telegram 
 * únicamente el día del cumpleaños.
 * 
 * Configuración necesaria:
 * 1. Hoja de Google Sheets llamada "Cumples" 
 * 2. Formato de fechas: MM/DD (Mes/Día)
 * 3. Bot de Telegram configurado
 * 
 * Funciones principales:
 * - notifyTodayBirthdays(): Función principal
 * - debugCumpleañosHoy(): Para verificar datos
 * - testCumpleañosHoy(): Para pruebas
 */

// ==========================================
// CONFIGURACIÓN DEL PROYECTO
// ==========================================
const CONFIG_TODAY = {
  NOMBRE_HOJA: "Cumples",
  TELEGRAM_TOKEN: '7972204638:AAHIpEUZbE-vey3xqpynQX-OMnxxg8mRwDc',
  CHAT_ID : '-1002844387414',
  //CHAT_ID: '-4963845348',// PRUEBAS
  // Personalizar mensajes aquí
  MENSAJE_CUMPLEANOS: (nombre, apellidos) => 
    `🎂 ¡Hoy está de cumpleaños *${nombre} ${apellidos}*! 🥳\n¡A llenarlo/a de buenos deseos! 🎉`,
  MENSAJE_ERROR_CRITICO: "🚨 *ERROR CRÍTICO EN NOTIFICADOR DE CUMPLEAÑOS*"
};

// ==========================================
// FUNCIÓN PRINCIPAL
// ==========================================
function notifyTodayBirthdays() {
  console.log("🚀 Iniciando notificador de cumpleaños para HOY...");
  
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_TODAY.NOMBRE_HOJA);
    if (!sheet) {
      throw new Error(`No se encontró la hoja "${CONFIG_TODAY.NOMBRE_HOJA}"`);
    }
    
    const datos = sheet.getDataRange().getValues();
    const hoy = new Date();
    const diaHoy = hoy.getDate();
    const mesHoy = hoy.getMonth() + 1;

    console.log(`📅 Fecha actual: ${diaHoy}/${mesHoy}/${hoy.getFullYear()}`);
    console.log(`🔍 Buscando cumpleaños para: ${diaHoy}/${mesHoy}`);

    let cumpleañosEncontrados = 0;
    let personasNotificadas = [];
    let erroresEncontrados = [];

    // Procesar cada fila (saltando el header)
    for (let i = 1; i < datos.length; i++) {
      const [nombre, apellidos, cumpleDato] = datos[i];
      
      // Saltar filas vacías
      if (!nombre && !apellidos) continue;

      try {
        const { dia, mes } = parsearFechaCumple(cumpleDato, i + 1);
        
        console.log(`👤 Procesando: ${nombre} ${apellidos} - Cumple: ${dia}/${mes}`);

        // ¡Verificar si cumple HOY!
        if (dia === diaHoy && mes === mesHoy) {
          console.log(`🎉 ¡CUMPLEAÑOS DETECTADO! ${nombre} ${apellidos}`);
          
          const mensaje = CONFIG_TODAY.MENSAJE_CUMPLEANOS(nombre, apellidos);
          
          if (enviarMensajeTelegram(mensaje)) {
            cumpleañosEncontrados++;
            personasNotificadas.push(`${nombre} ${apellidos}`);
            console.log(`✅ Notificación enviada para: ${nombre} ${apellidos}`);
          } else {
            console.log(`❌ Falló envío para: ${nombre} ${apellidos}`);
          }
        }

      } catch (error) {
        const errorMsg = `Fila ${i + 1} (${nombre || 'N/A'} ${apellidos || 'N/A'}): ${error.message}`;
        console.log(`⚠️ ${errorMsg}`);
        erroresEncontrados.push(errorMsg);
      }
    }

    // ==========================================
    // REPORTE FINAL
    // ==========================================
    console.log(`\n📊 RESUMEN DE EJECUCIÓN:`);
    console.log(`✅ Cumpleaños HOY: ${cumpleañosEncontrados}`);
    console.log(`⚠️ Errores encontrados: ${erroresEncontrados.length}`);
    
    if (personasNotificadas.length > 0) {
      console.log(`🎂 Personas notificadas: ${personasNotificadas.join(', ')}`);
    }

    // Enviar reporte de errores solo si los hay
    if (erroresEncontrados.length > 0) {
      enviarReporteErrores(cumpleañosEncontrados, personasNotificadas, erroresEncontrados);
    }

    // Log final
    if (cumpleañosEncontrados === 0 && erroresEncontrados.length === 0) {
      console.log('📝 No hay cumpleaños hoy. Ejecución completada exitosamente.');
    }

    return {
      cumpleanos: cumpleañosEncontrados,
      notificados: personasNotificadas,
      errores: erroresEncontrados.length
    };

  } catch (error) {
    console.error('🚨 ERROR CRÍTICO EN LA EJECUCIÓN:', error);
    
    const mensajeError = `${CONFIG_TODAY.MENSAJE_ERROR_CRITICO}\n\n` +
                        `📅 Fecha: ${new Date().toLocaleString('es-CO')}\n` +
                        `❌ Error: ${error.message}\n` +
                        `🔍 Ubicación: Función principal`;
    
    enviarMensajeTelegram(mensajeError);
    throw error;
  }
}

// ==========================================
// FUNCIONES DE PROCESAMIENTO
// ==========================================

function parsearFechaCumple(cumpleDato, numeroFila) {
  if (!cumpleDato || cumpleDato === '' || cumpleDato === null || cumpleDato === undefined) {
    throw new Error('Fecha de cumpleaños vacía o nula');
  }

  let dia, mes;

  try {
    // CASO 1: String en formato MM/DD
    if (typeof cumpleDato === 'string' && cumpleDato.includes('/')) {
      const partes = cumpleDato.trim().split('/');
      if (partes.length !== 2) {
        throw new Error(`Formato inválido: "${cumpleDato}" - debe ser MM/DD`);
      }
      
      mes = parseInt(partes[0].trim());  // Primer número es MES
      dia = parseInt(partes[1].trim());  // Segundo número es DÍA
      
      if (isNaN(mes) || isNaN(dia)) {
        throw new Error(`Valores no numéricos en: "${cumpleDato}"`);
      }
      
      console.log(`📅 Fecha parseada (string): ${dia}/${mes} desde "${cumpleDato}"`);
    }
    // CASO 2: Objeto Date (Google Sheets puede malinterpretar MM/DD como DD/MM)
    else if (cumpleDato instanceof Date) {
      const diaFromDate = cumpleDato.getDate();
      const mesFromDate = cumpleDato.getMonth() + 1;
      
      // Si ambos valores están entre 1-12, es posible que esté mal interpretado
      // Asumir que el formato original era MM/DD y Google lo interpretó como DD/MM
      if (diaFromDate <= 12 && mesFromDate <= 12) {
        dia = mesFromDate;  // Lo que Google pensó que era "mes" es realmente el día
        mes = diaFromDate;  // Lo que Google pensó que era "día" es realmente el mes
        console.log(`🔄 Fecha corregida (Date): ${dia}/${mes} (Google interpretó como ${diaFromDate}/${mesFromDate})`);
      } else {
        dia = diaFromDate;
        mes = mesFromDate;
        console.log(`📅 Fecha parseada (Date): ${dia}/${mes}`);
      }
    }
    // CASO 3: Objeto que puede convertirse a Date
    else if (typeof cumpleDato === 'object' && cumpleDato.toString) {
      const fechaConvertida = new Date(cumpleDato.toString());
      if (isNaN(fechaConvertida.getTime())) {
        throw new Error(`No se pudo convertir objeto a fecha: "${cumpleDato}"`);
      }
      
      const diaFromDate = fechaConvertida.getDate();
      const mesFromDate = fechaConvertida.getMonth() + 1;
      
      if (diaFromDate <= 12 && mesFromDate <= 12) {
        dia = mesFromDate;
        mes = diaFromDate;
        console.log(`🔄 Fecha corregida (Object): ${dia}/${mes}`);
      } else {
        dia = diaFromDate;
        mes = mesFromDate;
        console.log(`📅 Fecha parseada (Object): ${dia}/${mes}`);
      }
    }
    // CASO 4: Número serial de Excel/Sheets
    else if (typeof cumpleDato === 'number') {
      // Conversión de número serial de Excel a fecha
      const fechaConvertida = new Date((cumpleDato - 25569) * 86400 * 1000);
      if (isNaN(fechaConvertida.getTime())) {
        throw new Error(`Número serial inválido: ${cumpleDato}`);
      }
      
      const diaFromDate = fechaConvertida.getDate();
      const mesFromDate = fechaConvertida.getMonth() + 1;
      
      if (diaFromDate <= 12 && mesFromDate <= 12) {
        dia = mesFromDate;
        mes = diaFromDate;
        console.log(`🔄 Fecha corregida (Serial): ${dia}/${mes}`);
      } else {
        dia = diaFromDate;
        mes = mesFromDate;
        console.log(`📅 Fecha parseada (Serial): ${dia}/${mes}`);
      }
    } else {
      throw new Error(`Tipo de dato no compatible: ${typeof cumpleDato} - "${cumpleDato}"`);
    }

    // Validaciones finales
    if (mes < 1 || mes > 12) {
      throw new Error(`Mes fuera de rango: ${mes} (debe estar entre 1-12)`);
    }
    
    if (dia < 1 || dia > 31) {
      throw new Error(`Día fuera de rango: ${dia} (debe estar entre 1-31)`);
    }

    // Validación adicional para días según el mes
    const diasPorMes = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    if (dia > diasPorMes[mes - 1]) {
      throw new Error(`Día ${dia} inválido para el mes ${mes}`);
    }

    return { dia, mes };

  } catch (error) {
    throw new Error(`Error al procesar fecha en fila ${numeroFila}: ${error.message}`);
  }
}

// ==========================================
// FUNCIONES DE COMUNICACIÓN
// ==========================================

function enviarMensajeTelegram(texto) {
  const url = `https://api.telegram.org/bot${CONFIG_TODAY.TELEGRAM_TOKEN}/sendMessage`;
  
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        chat_id: CONFIG_TODAY.CHAT_ID,
        text: texto,
        parse_mode: 'Markdown',
        disable_web_page_preview: true
      }),
      muteHttpExceptions: true
    });
    
    const result = JSON.parse(response.getContentText());
    
    if (result.ok) {
      console.log(`📤 Mensaje enviado exitosamente a Telegram`);
      return true;
    } else {
      console.log(`📤❌ Error de Telegram: ${result.description || 'Error desconocido'}`);
      return false;
    }
  } catch (error) {
    console.log(`📤❌ Error de conexión a Telegram: ${error.message}`);
    return false;
  }
}

function enviarReporteErrores(cumpleañosEncontrados, personasNotificadas, erroresEncontrados) {
  const ahora = new Date();
  let reporte = `⚠️ *REPORTE DE ERRORES - CUMPLEAÑOS HOY*\n\n`;
  reporte += `📅 Ejecutado: ${ahora.toLocaleString('es-CO')}\n`;
  reporte += `🎂 Cumpleaños detectados: ${cumpleañosEncontrados}\n`;
  
  if (personasNotificadas.length > 0) {
    reporte += `✅ Notificados: ${personasNotificadas.join(', ')}\n`;
  }
  
  reporte += `⚠️ Errores encontrados: ${erroresEncontrados.length}\n\n`;
  reporte += `*Detalles de errores:*\n`;
  
  // Mostrar máximo 5 errores para no sobrecargar el mensaje
  const erroresAMostrar = erroresEncontrados.slice(0, 5);
  erroresAMostrar.forEach(error => {
    reporte += `• ${error}\n`;
  });
  
  if (erroresEncontrados.length > 5) {
    reporte += `• ... y ${erroresEncontrados.length - 5} errores adicionales\n`;
  }

  reporte += `\n🔧 Revisar la hoja "${CONFIG_TODAY.NOMBRE_HOJA}" para corregir errores.`;

  console.log('📊 Enviando reporte de errores...');
  enviarMensajeTelegram(reporte);
}

// ==========================================
// FUNCIONES DE PRUEBA Y DEBUG
// ==========================================

function testCumpleañosHoy() {
  console.log("🧪 === MODO DE PRUEBA - CUMPLEAÑOS HOY ===");
  
  const resultado = notifyTodayBirthdays();
  
  console.log("\n🧪 === RESULTADO DE LA PRUEBA ===");
  console.log(`Cumpleaños encontrados: ${resultado.cumpleanos}`);
  console.log(`Personas notificadas: ${resultado.notificados.join(', ') || 'Ninguna'}`);
  console.log(`Errores: ${resultado.errores}`);
}

function debugCumpleañosHoy() {
  console.log("🔍 === DEBUG DE CUMPLEAÑOS HOY ===");
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_TODAY.NOMBRE_HOJA);
  if (!sheet) {
    console.log(`❌ No se encontró la hoja "${CONFIG_TODAY.NOMBRE_HOJA}"`);
    return;
  }
  
  const datos = sheet.getDataRange().getValues();
  const hoy = new Date();
  const diaHoy = hoy.getDate();
  const mesHoy = hoy.getMonth() + 1;
  
  console.log(`📅 Fecha actual: ${diaHoy}/${mesHoy}/${hoy.getFullYear()}`);
  console.log(`🎯 Buscando cumpleaños para: ${diaHoy}/${mesHoy}\n`);
  
  let coincidenciasEncontradas = 0;
  
  for (let i = 1; i < Math.min(datos.length, 15); i++) {
    const [nombre, apellidos, cumpleDato] = datos[i];
    
    console.log(`--- Fila ${i + 1} ---`);
    console.log(`👤 Persona: ${nombre || 'N/A'} ${apellidos || 'N/A'}`);
    console.log(`📊 Dato original: "${cumpleDato}" (${typeof cumpleDato})`);
    
    try {
      const { dia, mes } = parsearFechaCumple(cumpleDato, i + 1);
      console.log(`📅 Interpretado como: ${dia}/${mes}`);
      
      if (dia === diaHoy && mes === mesHoy) {
        console.log(`🎂 ¡¡¡ CUMPLE HOY !!! 🎂`);
        coincidenciasEncontradas++;
      } else {
        console.log(`📝 No cumple hoy`);
      }
    } catch (error) {
      console.log(`❌ Error: ${error.message}`);
    }
    console.log("");
  }
  
  console.log(`🎯 TOTAL DE CUMPLEAÑOS HOY: ${coincidenciasEncontradas}`);
  
  if (datos.length > 15) {
    console.log(`\n📝 Nota: Solo se mostraron las primeras 15 filas. Total de filas: ${datos.length - 1}`);
  }
}

function simularCumpleañosHoy(nombrePrueba = "Juan", apellidosPrueba = "Pérez") {
  console.log(`🎭 Simulando cumpleaños para: ${nombrePrueba} ${apellidosPrueba}`);
  
  const mensaje = CONFIG_TODAY.MENSAJE_CUMPLEANOS(nombrePrueba, apellidosPrueba);
  console.log(`📝 Mensaje a enviar: ${mensaje}`);
  
  if (enviarMensajeTelegram(mensaje)) {
    console.log(`✅ Simulación exitosa`);
  } else {
    console.log(`❌ Falló la simulación`);
  }
}

function verProximosCumpleanos(diasAdelante = 7) {
  console.log(`📅 === PRÓXIMOS CUMPLEAÑOS (${diasAdelante} DÍAS) ===`);
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_TODAY.NOMBRE_HOJA);
  if (!sheet) {
    console.log(`❌ No se encontró la hoja "${CONFIG_TODAY.NOMBRE_HOJA}"`);
    return;
  }
  
  const datos = sheet.getDataRange().getValues();
  const hoy = new Date();
  
  for (let d = 0; d <= diasAdelante; d++) {
    const fechaCheck = new Date(hoy);
    fechaCheck.setDate(hoy.getDate() + d);
    const diaCheck = fechaCheck.getDate();
    const mesCheck = fechaCheck.getMonth() + 1;
    
    let cumpleañosDia = [];
    
    for (let i = 1; i < datos.length; i++) {
      const [nombre, apellidos, cumpleDato] = datos[i];
      
      try {
        const { dia, mes } = parsearFechaCumple(cumpleDato, i + 1);
        
        if (dia === diaCheck && mes === mesCheck) {
          cumpleañosDia.push(`${nombre} ${apellidos}`);
        }
      } catch (error) {
        // Ignorar errores en esta función de vista previa
      }
    }
    
    if (cumpleañosDia.length > 0) {
      const etiqueta = d === 0 ? '🎂 HOY' : `📅 En ${d} día${d > 1 ? 's' : ''}`;
      console.log(`${etiqueta} (${diaCheck}/${mesCheck}): ${cumpleañosDia.join(', ')}`);
    }
  }
}

// ==========================================
// FUNCIÓN DE CONFIGURACIÓN INICIAL
// ==========================================

function configurarProyecto() {
  console.log("⚙️ === CONFIGURACIÓN DEL PROYECTO ===");
  console.log(`📊 Hoja de datos: "${CONFIG_TODAY.NOMBRE_HOJA}"`);
  console.log(`🤖 Chat ID: ${CONFIG_TODAY.CHAT_ID}`);
  console.log(`🔑 Token configurado: ${CONFIG_TODAY.TELEGRAM_TOKEN ? 'Sí' : 'No'}`);
  
  // Verificar que la hoja existe
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_TODAY.NOMBRE_HOJA);
    if (sheet) {
      const filas = sheet.getLastRow();
      console.log(`✅ Hoja encontrada con ${filas - 1} personas registradas`);
    } else {
      console.log(`❌ Hoja "${CONFIG_TODAY.NOMBRE_HOJA}" no encontrada`);
    }
  } catch (error) {
    console.log(`❌ Error al verificar hoja: ${error.message}`);
  }
  
  // Probar conexión con Telegram
  console.log("\n🧪 Probando conexión con Telegram...");
  testTelegram();
}

function testTelegram() {
  const mensajePrueba = `🧪 *PRUEBA DE CONEXIÓN*\n📅 ${new Date().toLocaleString('es-CO')}\n✅ Sistema de cumpleaños funcionando correctamente`;
  
  if (enviarMensajeTelegram(mensajePrueba)) {
    console.log("✅ Conexión con Telegram exitosa");
  } else {
    console.log("❌ Falló la conexión con Telegram");
  }
}