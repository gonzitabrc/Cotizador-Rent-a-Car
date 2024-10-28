/**
*
* Cotizador Online de Rent a Car
* 
* Cotiza el valor del servicio de vehículos de alquiler y sus adicionales, y permite enviar la solicitud de reserva por email
*
* Version: 6.1
*
* Aplicación de Google Apps Script y Google Sheets -desarrollado por Gonzalo Reynoso, DDW -
* https://ddw.com.ar - gonzita@gmail.com
*
* licencia MIT: podés darle cualquier uso sin garantías y bajo tu responsabilidad, 
* no podés eliminar los créditos del autor ni el copyright en los archivos
*
**/

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Cotizador de Vehículos')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function generarJsonTarifas() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();

  var datosTabla = JSON.parse(obtenerDatosTarifas());

  var tarifas = {
    seguroExt: datosTabla.seguroExt,
    tarifa_diaria: datosTabla.tarifa_diaria
  };

var fechas = Object.keys(datosTabla.tarifa_diaria).sort();

function obtenerUltimoDiaMes(fecha) {
  var partesFecha = fecha.split('-');
  var year = parseInt(partesFecha[0]);
  var month = parseInt(partesFecha[1]) - 1; 
  
  var siguienteMes = new Date(year, month + 1, 1);
  
  var ultimoDia = new Date(siguienteMes.getTime() - 86400000); 
  
  return Utilities.formatDate(ultimoDia, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

var fechas_tarifas = {
  inicio_fecha_tarifa: fechas[0],
  fin_fecha_tarifa: obtenerUltimoDiaMes(fechas[fechas.length - 1])
};

  var adicionalesRange = sheet.getRange('Adicionales!A2:D').getValues();
  var adicionales = adicionalesRange.map(function(row) {
    return {
      adicional: row[0],
      max_unidades: row[1],
      costo_unitario: row[2],
      tipo_tarifa: row[3]
    };
  }).filter(function(row) {

    return row.adicional && row.max_unidades !== null && row.max_unidades !== undefined;
  });

  var formasPagoLastRow = sheet.getRange('Adicionales!F:F').getLastRow();
  var formasPagoRange = sheet.getRange('Adicionales!F2:H' + formasPagoLastRow).getValues();
  var formas_pago = formasPagoRange.map(function(row) {
    return {
      forma_pago: row[0],
      ajuste_decimal: row[1],
      cuotas: row[2]
    };
  }).filter(function(row) {
    return row.forma_pago && row.ajuste_decimal !== null && row.ajuste_decimal !== undefined && row.cuotas;
  });

  var pickupDropoffLastRow = sheet.getRange('Adicionales!J:J').getLastRow();
  var pickupDropoffRange = sheet.getRange('Adicionales!J2:K' + pickupDropoffLastRow).getValues();
  var pickup_dropoff = pickupDropoffRange.map(function(row) {
    return {
      location: row[0],
      costo: row[1]
    };
  }).filter(function(row) {

    return row.location && row.costo !== null && row.costo !== undefined;
  });

  var horarioNocturnoRange = sheet.getRange('Adicionales!M2:O2').getValues();
  var horario_nocturno = {
    inicio_horario_nocturno: formatTime(horarioNocturnoRange[0][0]),
    fin_horario_nocturno: formatTime(horarioNocturnoRange[0][1]),
    costo: horarioNocturnoRange[0][2]
  };

  var data = {
    tarifas: tarifas,
    fechas_tarifas: fechas_tarifas,
    adicionales: adicionales,
    formas_pago: formas_pago,
    pickup_dropoff: pickup_dropoff,
    horario_nocturno: horario_nocturno
  };

  Logger.log(JSON.stringify(data));
  //Logger.log(data.fechas_tarifas.fin_fecha_tarifa);
  return JSON.stringify(data);
}


function formatTime(value) {
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(new Date(value), Session.getScriptTimeZone(), 'HH:mm:ss');
  } else {
    return value; 
  }
}

function obtenerDatosTarifas() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tarifas");
  var range = sheet.getRange("A1:I");
  var values = range.getValues();
  
  var categorias = values[0].slice(1);
  var seguroExt = values[1].slice(1);
  var tarifa_diaria = {};
  
  var resultado = {
    seguroExt: {},
    tarifa_diaria: {}
  };
  
  for (var i = 2; i < values.length; i++) {
    var fecha = values[i][0];
    if (fecha instanceof Date && !isNaN(fecha)) {
      var anio = fecha.getFullYear();
      var mes = (fecha.getMonth() + 1).toString().padStart(2, '0');
      var dia = fecha.getDate().toString().padStart(2, '0');
      var fechaFormateada = `${anio}-${mes}-${dia}`;
      
      resultado.tarifa_diaria[fechaFormateada] = {};
      
      for (var j = 1; j < values[i].length; j++) {
        var valor = values[i][j];
        
        // Procesamos el valor del seguro, solo si es un número válido
        if (!resultado.seguroExt[categorias[j-1]]) {
          var seguroValor = parseFloat(seguroExt[j-1].toString().replace('$', '').replace(',', ''));
          if (!isNaN(seguroValor) && seguroValor > 0) {
            resultado.seguroExt[categorias[j-1]] = seguroValor;
          }
        }
        
        // Procesamos la tarifa diaria, si es un valor numérico mayor a 0
        if (typeof valor === 'string') {
          valor = parseFloat(valor.replace('$', '').replace(',', ''));
        }
        if (!isNaN(valor) && valor > 0) {
          resultado.tarifa_diaria[fechaFormateada][categorias[j-1]] = valor;
        }
      }
      
      // Si no hay categorías válidas para esa fecha, eliminar la fecha del resultado
      if (Object.keys(resultado.tarifa_diaria[fechaFormateada]).length === 0) {
        delete resultado.tarifa_diaria[fechaFormateada];
      }
    }
  }
  
  //Logger.log(JSON.stringify(resultado));
  return JSON.stringify(resultado);
}

function registrarReserva(datos) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Reservas');
  
  var fechaActual = new Date();
  var fechaActualFormateada = Utilities.formatDate(fechaActual, "America/Buenos_Aires", "yyyy-MM-dd HH:mm:ss");
  
sheet.appendRow([
    fechaActualFormateada, 
    datos.fechaRetiro, 
    datos.fechaDevolucion, 
    datos.lugarRetiro, 
    datos.lugarDevolucion, 
    datos.categoria, 
    datos.detalleServicios,
    datos.valor_alquiler,
    datos.valorTotalServicios,
    datos.valorTotalGeneral,
    datos.formaPago,
    datos.valorAPagar,
    datos.nombre, 
    datos.email, 
    datos.telefono, 
    datos.mensaje
]);

  enviarCorreo(datos);

  return {status: 'ok'};
}


function enviarCorreo(datos) {
  var destinatarioEmail = '';  // coloca entre las comillas el correo del destinatario del email por ej: 'tucorreo@gmail.com'
  var nombreDestinatario = 'Gonzalo';  // Nombre de la persona que recibirá el correo
  
  function formatearFecha(fecha) {
    var opciones = { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit', hour12: false };
    return fecha.toLocaleDateString('es-AR', opciones).replace(',', '');
  }

  var fechaActual = new Date();
  var fechaActualFormateada = Utilities.formatDate(fechaActual, "America/Buenos_Aires", "yyyy-MM-dd HH:mm:ss");
  
  var fechaRetiroFormateada = formatearFecha(new Date(datos.fechaRetiro));
  var fechaDevolucionFormateada = formatearFecha(new Date(datos.fechaDevolucion));
  
  var serviciosAdicionales = JSON.parse(datos.detalleServicios);
  var serviciosTexto = '<ul>';
  
  for (var key in serviciosAdicionales) {
    if (Array.isArray(serviciosAdicionales[key])) {
      serviciosTexto += '<li><strong>' + key + ':</strong><ul>';
      serviciosAdicionales[key].forEach(function(servicio) {
        serviciosTexto += '<li>' + servicio + '</li>';
      });
      serviciosTexto += '</ul></li>';
    } else {
      serviciosTexto += '<li><strong>' + key + ':</strong> $' + serviciosAdicionales[key] + '</li>';
    }
  }
  
  serviciosTexto += '</ul>';

  var asunto = 'Solicitud de Reserva desde el sitio web';
  var cuerpo = 
    '<p>Hola ' + nombreDestinatario + '.</p>' +
    '<p>A continuación los detalles de la cotización/solicitud que recibiste:</p>' +
    '<p><strong>Fecha y hora de la reserva:</strong> ' + fechaActualFormateada + '</p>' +
    '<p><strong>Fecha de retiro:</strong> ' + fechaRetiroFormateada + '</p>' +
    '<p><strong>Fecha de devolución:</strong> ' + fechaDevolucionFormateada + '</p>' +
    '<p><strong>Lugar de retiro:</strong> ' + datos.lugarRetiro + '</p>' +
    '<p><strong>Lugar de devolución:</strong> ' + datos.lugarDevolucion + '</p>' +
    '<p><strong>Categoría del vehículo:</strong> ' + datos.categoria + '</p>' +
    '<p><strong>Servicios extra y adicionales:</strong></p>' +
    serviciosTexto +
    '<p><strong>Valor del alquiler:</strong> $' + datos.valor_alquiler + '</p>' +
    '<p><strong>Adicionales y Cargos Extra:</strong> $' + datos.valorTotalServicios + '</p>' +
    '<p><strong>Valor total:</strong> $' + datos.valorTotalGeneral + '</p>' +
    '<p><strong>Forma de pago:</strong> ' + datos.formaPago + '</p>' +
    '<p><strong>Valor a pagar:</strong> $' + datos.valorAPagar + '</p>' +
    '<br>' +
    '<p><strong>Nombre cliente:</strong> ' + datos.nombre + '</p>' +
    '<p><strong>Email:</strong> ' + datos.email + '</p>' +
    '<p><strong>Teléfono:</strong> ' + datos.telefono + '</p>' +
    '<p><strong>Mensaje:</strong> ' + datos.mensaje + '</p>' +
    '<br>';

    destinatarioEmail = destinatarioEmail || Session.getActiveUser().getEmail();

  MailApp.sendEmail({
    to: destinatarioEmail,
    subject: asunto,
    htmlBody: cuerpo,
    replyTo: datos.email
  });
}

