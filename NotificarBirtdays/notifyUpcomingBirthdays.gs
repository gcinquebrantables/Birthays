// CONFIGURACIÓN
const CONFIG = {
  DIAS_ANTICIPACION: 3, // Cambiar este número según necesites
  NOMBRE_HOJA: "Cumples",
  TELEGRAM_TOKEN: '7972204638:AAHIpEUZbE-vey3xqpynQX-OMnxxg8mRwDc',
  CHAT_ID : '-1002844387414',
  //CHAT_ID: '-4963845348',// PRUEBAS
};

function notifyUpcomingBirthdays() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.NOMBRE_HOJA);
    const datos = sheet.getDataRange().getValues();

    const hoy = new Date();
    const fechaAnticipacion = new Date(hoy);
    fechaAnticipacion.setDate(hoy.getDate() + CONFIG.DIAS_ANTICIPACION);
    
    const diaTarget = fechaAnticipacion.getDate();
    const mesTarget = fechaAnticipacion.getMonth() + 1;

    console.log(`📅 Buscando cumpleaños para: ${diaTarget}/${mesTarget} (en ${CONFIG.DIAS_ANTICIPACION} días)`);

    let cumpleañosEncontrados = 0;
    let erroresEncontrados = [];

    // Procesar cada fila (saltando header)
    for (let i = 1; i < datos.length; i++) {
      const nombre = datos[i][0];
      const apellidos = datos[i][1];
      const cumpleDato = datos[i][2]; // Columna C formato MM/DD

      try {
        const { dia, mes } = parsearFechaCumple(cumpleDato, i + 1);
        
        console.log(`🧾 Evaluando: ${nombre} ${apellidos} - ${dia}/${mes}`);

        // Verificar si cumple en la fecha target
        if (dia === diaTarget && mes === mesTarget) {
          const mensaje = `🎉 *${nombre} ${apellidos}* cumple años en ${CONFIG.DIAS_ANTICIPACION} días.\n¡Vayan preparando las felicitaciones! 🎁`;
          console.log(`📨 Enviando mensaje: ${mensaje}`);
          
          if (enviarMensajeTelegram(mensaje)) {
            cumpleañosEncontrados++;
          }
        }

      } catch (error) {
        const errorMsg = `Fila ${i + 1} (${nombre} ${apellidos}): ${error.message}`;
        console.log(`⚠️ ${errorMsg}`);
        erroresEncontrados.push(errorMsg);
      }
    }

    // Reporte final
    console.log(`✅ Proceso completado:`);
    console.log(`   - Cumpleaños en ${CONFIG.DIAS_ANTICIPACION} días: ${cumpleañosEncontrados}`);
    console.log(`   - Errores encontrados: ${erroresEncontrados.length}`);

    // Solo enviar reporte de errores si los hay
    if (erroresEncontrados.length > 0) {
      enviarReporteErrores(cumpleañosEncontrados, erroresEncontrados);
    }

  } catch (error) {
    console.error('🚨 ERROR CRÍTICO:', error);
    const mensajeError = `🚨 *ERROR CRÍTICO EN NOTIFICADOR DE CUMPLEAÑOS*\n\n` +
                        `📅 Fecha: ${new Date().toLocaleString('es-CO')}\n` +
                        `❌ Error: ${error.message}`;
    enviarMensajeTelegram(mensajeError);
    throw error;
  }
}

// Función para parsear fechas en formato MM/DD
function parsearFechaCumple(cumpleDato, numeroFila) {
  if (!cumpleDato || cumpleDato === '') {
    throw new Error('Fecha de cumpleaños vacía');
  }

  let dia, mes;

  // CASO 1: String en formato MM/DD
  if (typeof cumpleDato === 'string' && cumpleDato.includes('/')) {
    const partes = cumpleDato.split('/');
    if (partes.length !== 2) {
      throw new Error(`Formato inválido: "${cumpleDato}" - debe ser MM/DD`);
    }
    
    mes = parseInt(partes[0].trim());  // Primer número es MES
    dia = parseInt(partes[1].trim());  // Segundo número es DÍA
    
    if (isNaN(mes) || isNaN(dia)) {
      throw new Error(`Formato inválido: "${cumpleDato}" - no son números válidos`);
    }
  }
  // CASO 2: Objeto Date (Google Sheets puede malinterpretar MM/DD)
  else if (cumpleDato instanceof Date) {
    const diaFromDate = cumpleDato.getDate();
    const mesFromDate = cumpleDato.getMonth() + 1;
    
    // Si ambos valores están entre 1-12, asumir que Google malinterpretó MM/DD
    if (diaFromDate <= 12 && mesFromDate <= 12) {
      dia = mesFromDate;  // Lo que Google pensó que era "mes" es realmente el día
      mes = diaFromDate;  // Lo que Google pensó que era "día" es realmente el mes
      console.log(`🔄 Fecha corregida en fila ${numeroFila}: ${dia}/${mes} (era ${diaFromDate}/${mesFromDate})`);
    } else {
      dia = diaFromDate;
      mes = mesFromDate;
    }
  }
  // CASO 3: Object convertible a Date
  else if (typeof cumpleDato === 'object' && cumpleDato.toString) {
    const fechaConvertida = new Date(cumpleDato.toString());
    if (isNaN(fechaConvertida.getTime())) {
      throw new Error('No se pudo convertir a fecha válida');
    }
    
    const diaFromDate = fechaConvertida.getDate();
    const mesFromDate = fechaConvertida.getMonth() + 1;
    
    if (diaFromDate <= 12 && mesFromDate <= 12) {
      dia = mesFromDate;
      mes = diaFromDate;
    } else {
      dia = diaFromDate;
      mes = mesFromDate;
    }
  }
  // CASO 4: Número serial de Excel
  else if (typeof cumpleDato === 'number') {
    const fechaConvertida = new Date((cumpleDato - 25569) * 86400 * 1000);
    if (isNaN(fechaConvertida.getTime())) {
      throw new Error('No se pudo convertir número a fecha válida');
    }
    
    const diaFromDate = fechaConvertida.getDate();
    const mesFromDate = fechaConvertida.getMonth() + 1;
    
    if (diaFromDate <= 12 && mesFromDate <= 12) {
      dia = mesFromDate;
      mes = diaFromDate;
    } else {
      dia = diaFromDate;
      mes = mesFromDate;
    }
  } else {
    throw new Error(`Tipo de dato no reconocido: ${typeof cumpleDato}`);
  }

  // Validar rangos
  if (mes < 1 || mes > 12) {
    throw new Error(`Mes inválido: ${mes} (debe estar entre 1-12)`);
  }
  
  if (dia < 1 || dia > 31) {
    throw new Error(`Día inválido: ${dia} (debe estar entre 1-31)`);
  }

  return { dia, mes };
}

// Función para enviar reporte de errores
function enviarReporteErrores(cumpleañosEncontrados, erroresEncontrados) {
  const ahora = new Date();
  let mensajeReporte = `⚠️ *ERRORES EN NOTIFICADOR DE CUMPLEAÑOS*\n\n`;
  mensajeReporte += `📅 Ejecutado: ${ahora.toLocaleString('es-CO')}\n`;
  mensajeReporte += `🎉 Cumpleaños en ${CONFIG.DIAS_ANTICIPACION} días: ${cumpleañosEncontrados}\n`;
  mensajeReporte += `⚠️ Errores encontrados: ${erroresEncontrados.length}\n\n`;
  mensajeReporte += `*Detalles de errores:*\n`;
  
  // Mostrar máximo 5 errores
  erroresEncontrados.slice(0, 5).forEach(error => {
    mensajeReporte += `• ${error}\n`;
  });
  
  if (erroresEncontrados.length > 5) {
    mensajeReporte += `• ... y ${erroresEncontrados.length - 5} errores más\n`;
  }

  console.log('⚠️ Enviando reporte de errores...');
  enviarMensajeTelegram(mensajeReporte);
}

// Función para enviar mensajes a Telegram
function enviarMensajeTelegram(texto) {
  const url = `https://api.telegram.org/bot${CONFIG.TELEGRAM_TOKEN}/sendMessage`;
  
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        chat_id: CONFIG.CHAT_ID,
        text: texto,
        parse_mode: 'Markdown'
      }),
      muteHttpExceptions: true
    });
    
    const result = JSON.parse(response.getContentText());
    
    if (result.ok) {
      console.log(`✅ Mensaje enviado exitosamente a Telegram`);
      return true;
    } else {
      console.log(`❌ Error al enviar mensaje: ${result.description}`);
      return false;
    }
  } catch (error) {
    console.log(`❌ Error de conexión: ${error.message}`);
    return false;
  }
}

// Función de prueba
function testNotificacion() {
  console.log(`🧪 Probando notificación con ${CONFIG.DIAS_ANTICIPACION} días de anticipación...`);
  notifyUpcomingBirthdays();
}

// Función para cambiar los días de anticipación fácilmente
function cambiarDiasAnticipacion(nuevosDias) {
  CONFIG.DIAS_ANTICIPACION = nuevosDias;
  console.log(`✅ Días de anticipación cambiados a: ${nuevosDias}`);
}

// Función de debug para verificar datos
function debugBirthdayData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.NOMBRE_HOJA);
  const datos = sheet.getDataRange().getValues();
  
  console.log(`=== DEBUG DE DATOS DE CUMPLEAÑOS ===`);
  console.log(`Configuración actual: ${CONFIG.DIAS_ANTICIPACION} días de anticipación`);
  
  const hoy = new Date();
  const fechaTarget = new Date(hoy);
  fechaTarget.setDate(hoy.getDate() + CONFIG.DIAS_ANTICIPACION);
  console.log(`Fecha objetivo: ${fechaTarget.getDate()}/${fechaTarget.getMonth() + 1}`);
  
  for (let i = 1; i < Math.min(datos.length, 10); i++) {
    const nombre = datos[i][0];
    const apellidos = datos[i][1];
    const cumpleDato = datos[i][2];
    
    console.log(`\nFila ${i + 1}: ${nombre} ${apellidos}`);
    console.log(`  Valor original: "${cumpleDato}" (${typeof cumpleDato})`);
    
    try {
      const { dia, mes } = parsearFechaCumple(cumpleDato, i + 1);
      console.log(`  Interpretado como: ${dia}/${mes}`);
      
      if (dia === fechaTarget.getDate() && mes === fechaTarget.getMonth() + 1) {
        console.log(`  ⭐ COINCIDE CON FECHA OBJETIVO ⭐`);
      }
    } catch (error) {
      console.log(`  ❌ Error: ${error.message}`);
    }
  }
}