// Configuracion de Telegram
const TELEGRAM_TOKEN = '7972204638:AAHIpEUZbE-vey3xqpynQX-OMnxxg8mRwDc';
const CHAT_ID = '-4963845348';

function onFormSubmit(e) {
  // Crear un registro detallado para debugging
  var debugInfo = [];
  
  try {
    debugInfo.push("=== INICIO DEL SCRIPT ===");
    debugInfo.push("Timestamp: " + new Date().toISOString());
    
    // Verificar que el evento tiene los datos necesarios
    if (!e) {
      // Si no hay evento, verificar si hay datos recientes en la hoja
      debugInfo.push("No se recibio evento, intentando leer desde la hoja...");
      return procesarDesdeHoja(debugInfo);
    }
    
    debugInfo.push("Estructura del evento recibido: " + JSON.stringify(Object.keys(e)));
    
    var nombre = "";
    var apellido = "";
    var archivosIds = [];
    
    // MANEJAR DIFERENTES TIPOS DE EVENTOS
    if (e.response) {
      // Evento desde Google Forms (metodo preferido)
      debugInfo.push("Procesando evento de Google Forms");
      var formResponse = e.response;
      var itemResponses = formResponse.getItemResponses();
      
      debugInfo.push("Numero de respuestas recibidas: " + itemResponses.length);
      
      // Extraer datos del formulario
      for (var i = 0; i < itemResponses.length; i++) {
        var itemResponse = itemResponses[i];
        var title = itemResponse.getItem().getTitle();
        var response = itemResponse.getResponse();
        
        debugInfo.push("Pregunta " + (i+1) + ": '" + title + "' | Respuesta: '" + response + "'");
        
        if (title.toLowerCase().includes("nombre") && !title.toLowerCase().includes("apellido")) {
          nombre = response;
          debugInfo.push("NOMBRE encontrado: '" + nombre + "'");
        } else if (title.toLowerCase().includes("apellido")) {
          apellido = response;
          debugInfo.push("APELLIDO encontrado: '" + apellido + "'");
        } else if (itemResponse.getItem().getType() == FormApp.ItemType.FILE_UPLOAD) {
          var fileIds = response;
          if (fileIds && fileIds.length > 0) {
            if (Array.isArray(fileIds)) {
              archivosIds = archivosIds.concat(fileIds);
            } else {
              archivosIds.push(fileIds);
            }
            debugInfo.push("ARCHIVOS encontrados: " + fileIds.length + " archivos");
          }
        }
      }
    } else if (e.values || e.namedValues) {
      // Evento desde Google Sheets (metodo alternativo)
      debugInfo.push("Procesando evento de Google Sheets");
      
      var valores = e.values || [];
      var nombresColumnas = e.namedValues || {};
      
      debugInfo.push("Valores recibidos: " + JSON.stringify(valores));
      debugInfo.push("Nombres de columnas: " + JSON.stringify(Object.keys(nombresColumnas)));
      
      // Buscar en namedValues primero
      if (nombresColumnas) {
        for (var columna in nombresColumnas) {
          var valorColumna = nombresColumnas[columna];
          if (Array.isArray(valorColumna) && valorColumna.length > 0) {
            valorColumna = valorColumna[0]; // Tomar el primer valor
          }
          
          debugInfo.push("Columna: '" + columna + "' = '" + valorColumna + "'");
          
          if (columna.toLowerCase().includes("nombre") && !columna.toLowerCase().includes("apellido")) {
            nombre = valorColumna;
            debugInfo.push("NOMBRE encontrado en columna: '" + nombre + "'");
          } else if (columna.toLowerCase().includes("apellido")) {
            apellido = valorColumna;
            debugInfo.push("APELLIDO encontrado en columna: '" + apellido + "'");
          } else if (columna.toLowerCase().includes("foto") || columna.toLowerCase().includes("archivo")) {
            // Los archivos en Sheets aparecen como URLs o IDs
            if (valorColumna && valorColumna.toString().length > 10) {
              // Intentar extraer ID de archivo de la URL
              var archivoId = extraerIdDeArchivo(valorColumna.toString());
              if (archivoId) {
                archivosIds.push(archivoId);
                debugInfo.push("ARCHIVO ID extraido: " + archivoId);
              }
            }
          }
        }
      }
      
      // Si no encontramos en namedValues, intentar con values por posicion
      if ((!nombre || !apellido || archivosIds.length === 0) && valores && valores.length > 0) {
        debugInfo.push("Intentando extraer de valores por posicion...");
        // Asumir orden: timestamp, nombre, apellido, foto
        if (valores.length >= 4) {
          nombre = nombre || valores[1];
          apellido = apellido || valores[2];
          
          if (valores[3] && valores[3].toString().length > 10) {
            var archivoId = extraerIdDeArchivo(valores[3].toString());
            if (archivoId) {
              archivosIds.push(archivoId);
              debugInfo.push("ARCHIVO ID extraido de posicion: " + archivoId);
            }
          }
        }
      }
    } else {
      throw new Error("Estructura de evento no reconocida");
    }
    
    return procesarFormulario(nombre, apellido, archivosIds, debugInfo);
    
  } catch (error) {
    debugInfo.push("ERROR CRITICO: " + error.toString());
    debugInfo.push("Stack trace: " + error.stack);
    escribirLog(debugInfo);
    
    // Notificar error por Telegram
    enviarErrorTelegram(error.toString(), debugInfo);
    
    throw error;
  }
}

/**
 * Funcion para procesar desde la hoja cuando no hay evento
 */
function procesarDesdeHoja(debugInfo) {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    var lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      throw new Error("No hay datos en la hoja para procesar");
    }
    
    // Obtener la ultima fila de datos
    var datos = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    debugInfo.push("Datos de la ultima fila: " + JSON.stringify(datos));
    
    var nombre = datos[1] || "";
    var apellido = datos[2] || "";
    var foto = datos[3] || "";
    
    var archivosIds = [];
    if (foto) {
      var archivoId = extraerIdDeArchivo(foto.toString());
      if (archivoId) {
        archivosIds.push(archivoId);
      }
    }
    
    debugInfo.push("Extraido desde hoja - Nombre: " + nombre + ", Apellido: " + apellido + ", Archivos: " + archivosIds.length);
    
    return procesarFormulario(nombre, apellido, archivosIds, debugInfo);
    
  } catch (error) {
    debugInfo.push("Error procesando desde hoja: " + error.toString());
    throw error;
  }
}

/**
 * Funcion principal para procesar el formulario
 */
function procesarFormulario(nombre, apellido, archivosIds, debugInfo) {
  // Verificar que tenemos todos los datos
  debugInfo.push("--- RESUMEN DE DATOS ---");
  debugInfo.push("Nombre final: '" + nombre + "'");
  debugInfo.push("Apellido final: '" + apellido + "'");
  debugInfo.push("Total archivos: " + archivosIds.length);
  
  if (!nombre || !apellido || archivosIds.length === 0) {
    var errorMsg = "FALTAN DATOS - nombre: '" + nombre + "', apellido: '" + apellido + "', archivos: " + archivosIds.length;
    debugInfo.push(errorMsg);
    escribirLog(debugInfo);
    throw new Error(errorMsg);
  }
  
  // Limpiar y procesar nombres (NUEVO FORMATO)
  var nombreFormateado = formatearNombre(nombre);
  var apellidoFormateado = formatearNombre(apellido);
  var baseNombre = nombreFormateado + " " + apellidoFormateado + " ";
  
  debugInfo.push("--- PROCESAMIENTO ---");
  debugInfo.push("Nombre formateado: '" + nombreFormateado + "'");
  debugInfo.push("Apellido formateado: '" + apellidoFormateado + "'");
  debugInfo.push("Base del nombre: '" + baseNombre + "'");
  
  // USAR EL ID DE CARPETA CORRECTO
  var carpetaId = "11dlqXIdGV3bK-4ii16B_yr27hgJnICbnBgr4_CMe9QWdNI8d1CKxNMGwAST50DAiepyPYdng";
  var carpetaDestino = DriveApp.getFolderById(carpetaId);
  debugInfo.push("Carpeta accesible: " + carpetaDestino.getName());
  
  // Obtener numero secuencial
  var numeroSecuencial = obtenerSiguienteNumero(carpetaDestino, baseNombre);
  debugInfo.push("Numero secuencial inicial: " + numeroSecuencial);
  
  var archivosRenombrados = [];
  
  // Procesar cada archivo
  for (var j = 0; j < archivosIds.length; j++) {
    try {
      debugInfo.push("--- PROCESANDO ARCHIVO " + (j+1) + " ---");
      debugInfo.push("ID del archivo: " + archivosIds[j]);
      
      // Intentar acceder al archivo con mas informacion de debug
      var archivo;
      try {
        archivo = DriveApp.getFileById(archivosIds[j]);
      } catch (fileError) {
        debugInfo.push("Error accediendo al archivo: " + fileError.toString());
        
        // Intentar diferentes metodos para encontrar el archivo
        debugInfo.push("Intentando buscar archivos en la carpeta...");
        var archivosEnCarpeta = carpetaDestino.getFiles();
        var archivosEncontrados = [];
        
        while (archivosEnCarpeta.hasNext()) {
          var archivoEnCarpeta = archivosEnCarpeta.next();
          archivosEncontrados.push({
            nombre: archivoEnCarpeta.getName(),
            id: archivoEnCarpeta.getId(),
            fecha: archivoEnCarpeta.getDateCreated()
          });
        }
        
        debugInfo.push("Archivos en la carpeta: " + JSON.stringify(archivosEncontrados));
        
        // Si hay archivos en la carpeta, usar el mas reciente
        if (archivosEncontrados.length > 0) {
          // Ordenar por fecha (mas reciente primero)
          archivosEncontrados.sort(function(a, b) { return b.fecha - a.fecha; });
          var archivoMasReciente = archivosEncontrados[0];
          
          debugInfo.push("Usando archivo mas reciente: " + archivoMasReciente.nombre);
          archivo = DriveApp.getFileById(archivoMasReciente.id);
          archivosIds[j] = archivoMasReciente.id; // Actualizar el ID
        } else {
          throw new Error("No se encontraron archivos en la carpeta");
        }
      }
      
      var nombreOriginal = archivo.getName();
      var extension = obtenerExtension(nombreOriginal);
      
      debugInfo.push("Archivo encontrado: '" + nombreOriginal + "'");
      debugInfo.push("Extension: '" + extension + "'");
      
      // Crear el nuevo nombre (NUEVO FORMATO: Alberto Plaza 1.png)
      var nuevoNombre = baseNombre + numeroSecuencial + "." + extension;
      debugInfo.push("Nuevo nombre: '" + nuevoNombre + "'");
      
      // Verificar duplicados
      var intentos = 0;
      while (existeArchivo(carpetaDestino, nuevoNombre) && intentos < 100) {
        numeroSecuencial++;
        nuevoNombre = baseNombre + numeroSecuencial + "." + extension;
        intentos++;
      }
      
      if (intentos > 0) {
        debugInfo.push("Ajustado por duplicados a: '" + nuevoNombre + "'");
      }
      
      // Crear la copia
      var nuevaCopia = archivo.makeCopy(nuevoNombre, carpetaDestino);
      
      // ELIMINAR EL ARCHIVO ORIGINAL (activado)
      try {
        archivo.setTrashed(true);
        debugInfo.push("Archivo original eliminado");
      } catch (deleteError) {
        debugInfo.push("No se pudo eliminar el original: " + deleteError.toString());
      }
      
      archivosRenombrados.push({
        original: nombreOriginal,
        nuevo: nuevoNombre,
        url: nuevaCopia.getUrl()
      });
      
      debugInfo.push("Archivo renombrado exitosamente");
      numeroSecuencial++;
      
    } catch (archivoError) {
      debugInfo.push("Error procesando archivo " + (j+1) + ": " + archivoError.toString());
    }
  }
  
  // Actualizar hoja de calculo
  actualizarHojaCalculo(archivosRenombrados, nombre, apellido, debugInfo);
  
  // Enviar notificacion a Telegram
  enviarNotificacionTelegram(nombre, apellido, archivosRenombrados);
  
  debugInfo.push("PROCESO COMPLETADO");
  debugInfo.push("Archivos procesados: " + archivosRenombrados.length);
  debugInfo.push("=== FIN DEL SCRIPT ===");
  
  escribirLog(debugInfo);
}

/**
 * Funcion para extraer ID de archivo de URL de Google Drive
 */
function extraerIdDeArchivo(url) {
  try {
    // Patrones comunes de URLs de Google Drive
    var patrones = [
      /\/file\/d\/([a-zA-Z0-9-_]+)/,
      /id=([a-zA-Z0-9-_]+)/,
      /^([a-zA-Z0-9-_]{25,})$/  // ID directo
    ];
    
    for (var i = 0; i < patrones.length; i++) {
      var match = url.match(patrones[i]);
      if (match && match[1]) {
        return match[1];
      }
    }
    
    // Si no coincide con patrones, pero parece un ID valido
    if (url.length > 20 && url.length < 50 && /^[a-zA-Z0-9-_]+$/.test(url)) {
      return url;
    }
    
    return null;
  } catch (error) {
    console.error("Error extrayendo ID de archivo: " + error.toString());
    return null;
  }
}

/**
 * Funcion para enviar notificacion exitosa a Telegram
 */
function enviarNotificacionTelegram(nombre, apellido, archivosRenombrados) {
  try {
    var fecha = new Date().toLocaleString('es-ES', {
      timeZone: 'America/Bogota',
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit'
    });
    
    var mensaje = "📸 *NUEVA FOTO PROCESADA*\n\n";
    mensaje += "👤 *Persona:* " + nombre + " " + apellido + "\n";
    mensaje += "📅 *Fecha:* " + fecha + "\n";
    mensaje += "📁 *Archivos procesados:* " + archivosRenombrados.length + "\n\n";
    
    mensaje += "*Archivos renombrados:*\n";
    for (var i = 0; i < archivosRenombrados.length; i++) {
      mensaje += "• `" + archivosRenombrados[i].nuevo + "`\n";
    }
    
    mensaje += "\n✅ *Estado:* Procesado correctamente";
    
    enviarMensajeTelegram(mensaje);
    
  } catch (error) {
    console.error("Error enviando notificacion a Telegram: " + error.toString());
  }
}

/**
 * Funcion para enviar error a Telegram
 */
function enviarErrorTelegram(errorMsg, debugInfo) {
  try {
    var fecha = new Date().toLocaleString('es-ES', {
      timeZone: 'America/Bogota'
    });
    
    var mensaje = "🚨 *ERROR EN FORMULARIO*\n\n";
    mensaje += "📅 *Fecha:* " + fecha + "\n";
    mensaje += "❌ *Error:* " + errorMsg + "\n\n";
    mensaje += "*Ultimos logs:*\n";
    
    // Tomar solo las ultimas 5 lineas del debug para no sobrecargar
    var ultimosLogs = debugInfo.slice(-5);
    for (var i = 0; i < ultimosLogs.length; i++) {
      mensaje += "• " + ultimosLogs[i] + "\n";
    }
    
    enviarMensajeTelegram(mensaje);
    
  } catch (error) {
    console.error("Error enviando error a Telegram: " + error.toString());
  }
}

/**
 * Funcion base para enviar mensajes a Telegram
 */
function enviarMensajeTelegram(mensaje) {
  try {
    var url = "https://api.telegram.org/bot" + TELEGRAM_TOKEN + "/sendMessage";
    
    var payload = {
      'chat_id': CHAT_ID,
      'text': mensaje,
      'parse_mode': 'Markdown',
      'disable_web_page_preview': true
    };
    
    var options = {
      'method': 'POST',
      'headers': {
        'Content-Type': 'application/json'
      },
      'payload': JSON.stringify(payload)
    };
    
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      console.log("Mensaje enviado a Telegram correctamente");
    } else {
      console.error("Error enviando a Telegram. Codigo: " + responseCode);
      console.error("Respuesta: " + response.getContentText());
    }
    
  } catch (error) {
    console.error("Error en enviarMensajeTelegram: " + error.toString());
  }
}

/**
 * Funcion para formatear nombres con primera letra mayuscula
 */
function formatearNombre(texto) {
  if (!texto) return "";
  
  // Limpiar el texto y convertir a minusculas
  var textoLimpio = texto
    .toString()
    .trim()
    .toLowerCase()
    .replace(/[^a-zA-Z\s]/g, "") // Mantener espacios para nombres compuestos
    .replace(/\s+/g, " "); // Normalizar espacios multiples
  
  // Capitalizar primera letra de cada palabra
  return textoLimpio.replace(/\b\w/g, function(letra) {
    return letra.toUpperCase();
  });
}

/**
 * Funcion para obtener extension de archivo
 */
function obtenerExtension(nombreArchivo) {
  var partes = nombreArchivo.split('.');
  return partes.length > 1 ? partes[partes.length - 1].toLowerCase() : "jpg";
}

/**
 * Funcion para obtener el siguiente numero secuencial (ACTUALIZADA PARA NUEVO FORMATO)
 */
function obtenerSiguienteNumero(carpeta, baseNombre) {
  var archivos = carpeta.getFiles();
  var numerosUsados = [];
  
  while (archivos.hasNext()) {
    var archivo = archivos.next();
    var nombreArchivo = archivo.getName();
    
    // Escapar caracteres especiales en baseNombre para regex
    var baseEscapado = baseNombre.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    var patron = new RegExp("^" + baseEscapado + "(\\d+)\\.", "i");
    var coincidencia = nombreArchivo.match(patron);
    
    if (coincidencia) {
      var numero = parseInt(coincidencia[1]);
      if (!isNaN(numero)) {
        numerosUsados.push(numero);
      }
    }
  }
  
  if (numerosUsados.length === 0) {
    return 1;
  }
  
  numerosUsados.sort(function(a, b) { return a - b; });
  
  for (var i = 1; i <= numerosUsados.length + 1; i++) {
    if (numerosUsados.indexOf(i) === -1) {
      return i;
    }
  }
  
  return numerosUsados.length + 1;
}

/**
 * Funcion para verificar si existe un archivo
 */
function existeArchivo(carpeta, nombreArchivo) {
  var archivos = carpeta.getFilesByName(nombreArchivo);
  return archivos.hasNext();
}

/**
 * Funcion para escribir log detallado
 */
function escribirLog(debugInfo) {
  console.log(debugInfo.join("\n"));
}

/**
 * Funcion para actualizar la hoja de calculo
 */
function actualizarHojaCalculo(archivosRenombrados, nombre, apellido, debugInfo) {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    var lastRow = sheet.getLastRow();
    
    // Agregar columnas si no existen
    if (sheet.getLastColumn() < 8) {
      sheet.getRange(1, 5).setValue("Archivos Procesados");
      sheet.getRange(1, 6).setValue("Nombres Nuevos");
      sheet.getRange(1, 7).setValue("URLs");
      sheet.getRange(1, 8).setValue("Debug Info");
    }
    
    if (archivosRenombrados.length > 0) {
      var resumenArchivos = archivosRenombrados.map(function(a) { return a.original; }).join("\n");
      var resumenNombres = archivosRenombrados.map(function(a) { return a.nuevo; }).join("\n");
      var resumenUrls = archivosRenombrados.map(function(a) { return a.url; }).join("\n");
      
      sheet.getRange(lastRow, 5).setValue(resumenArchivos);
      sheet.getRange(lastRow, 6).setValue(resumenNombres);
      sheet.getRange(lastRow, 7).setValue(resumenUrls);
    }
    
    // Agregar info de debug (ultimas 10 lineas)
    var debugResumen = debugInfo.slice(-10).join(" | ");
    sheet.getRange(lastRow, 8).setValue(debugResumen);
    
  } catch (error) {
    console.error("Error actualizando hoja: " + error.toString());
  }
}