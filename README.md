# Cotizador-Rent-a-Car
Cotizador Online de Rent a Car y Formulario para solicitar reservas

DESCRIPCIÓN - COTIZADOR DE VEHÍCULOS DE ALQUILER V 6.1
Esta aplicación desarrollada en Google Apps Script y Google Sheets es un sistema sencillo de cotización de servicios para rentadoras de vehículos sin chofer (rent a car)
Permite al usuario (cliente) seleccionar las fechas y lugares de retiro y devolución del vehículo, seleccionar una categoría de su interés, y agregar servicios adicionales, recibiendo instantaneamente la cotización total del servicio
Como rentadora podés cargar tus categorías de vehículos, los servicios adicionales y sus costos, las formas de pago y sus costos de financiación o descuentos, los lugares de pickup/dropoff y sus costos, el costo de horario nocturno de pickup/dropoff
No es un sistema de reservas ni de gestión de disponibilidad, solo calcula la cotización del servicio y permite enviar la solicitud de reserva por parte del cliente, la cual se registra además en la pestaña "Reservas" de la hoja de cálculo. 
Es completamente integrable con cualquier sitio web de rentacar que no posee un sistema de reservas
Posee licencia MIT: podés usarlo sin cargo, modificarlo, etc, sin garantías y bajo tu propia responsabilidad, no podés eliminar ni editar los créditos del autor (Gonzalo Reynoso, DDW) ni el copyright


SETUP COTIZADOR - INSTRUCCIONES PARA IMPLEMENTAR LA APLICACIÓN
1) crea una copia de esta hoja de cálculo https://docs.google.com/spreadsheets/d/1o0SSkVmL1IZ6LSxLs7VwV1X7FCDmZC3L2y8EjKh_hiI/edit?usp=sharing en tu espacio de Google Drive: archivo >> crear una copia  (guarda el archivo y envía su link al escritorio o a favoritos de tu navegador para tenerlo siempre a mano. 
2) edita la hoja de cálculo copiada para personalizar los datos del servicio de alquiler en las pestañas 'TarifasEditable' y 'Adicionales' (solo edita las celdas amarillas)
3) en el menú superior de la hoja andá a Extensiones >> Google Apps Script: Implementar >> Nueva Implementación (selecciona "Aplicación Web, Ejecutar como: Yo, Quien tiene acceso: cualquier usuario) >> implementar
sigue los pasos, acepta los permisos, al finalizar copia la url de la aplicación web (tiene esta forma: https://script.google.com/macros/s/......../exec)
4) en el editor de Google Apps Script edita el archivo 'funciones-gas.gs' en las lineas 207 y 208 (nombre del destinatario y correo electrónico donde se enviará la solicitud de reserva)
5) en el editor de Google Apps Script edita el archivo 'iframe.html' en las lineas 94 a 99 (reemplaza allí la url de tu sitio web, el directorio y nombre del logotipo, y la url de la aplicación web que obtuviste en el paso 3)
6) con el código de ese archivo 'iframe.html' crea nuevo un archivo html y subilo al directorio web donde deseas publicar tu cotizador, por ejemplo: https://tusitioweb.com/cotizador.html
7) accede a la url del paso anterior y prueba el formulario cotizador, una vez que cotizas el servicio puedes enviar la solicitud de reserva y esta llegará por email al correo que configuraste en el paso 4)

