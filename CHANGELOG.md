# Historial de cambios

Todas las versiones notables de Renamer Avantra.

El formato sigue las ideas de [Keep a Changelog](https://keepachangelog.com/es-ES/1.1.0/)
y el proyecto usa [Versionado Semántico](https://semver.org/lang/es/).

## [1.3.0] - 2026-10-20

### Añadido
- **Detección universal de identificadores**: el modo por defecto
  reconoce cualquier estructura de nombre de archivo por capas
  (radicado → prefijo antes del primer espacio → nombre completo),
  en lugar de exigir el patrón de radicado.
- **Estructura reglada del Excel de correspondencia**: origen y destino
  en columnas contiguas, con detección automática de ambas capas
  (clásica numérica y regla general alfanumérica como `EXP-2026-001`).
- **Verificar Excel**: dictamen instantáneo de la estructura (filas,
  pares, conflictos y filas saltadas con motivo).
- **Crear plantilla**: genera un Excel de correspondencia de ejemplo
  con instrucciones.
- **Opciones avanzadas**: modos de detección (Universal / Radicado
  estricto / Personalizado) con regex propias y prueba en vivo;
  se persisten en `config.json`.
- CLI: `--modo`, `--regex-origen` y `--regex-destino`.
- Nuevas pruebas: nombres aleatorios, fuzz de 50 nombres y destino
  alfanumérico.

## [1.2.2] - 2026-09-16

### Mejorado
- Interfaz más compacta: los campos de entrada se ven blancos en cualquier
  tema de Windows y la tabla de resultados gana espacio vertical.
- Versión visible en la barra superior.
- Acción sugerida abreviada en la tabla de pendientes (el detalle completo
  permanece en `pendientes_correccion.csv`).
- Capturas y documentación renovadas.

### Cambiado
- Texto de ejemplo de la interfaz actualizado a datos totalmente ficticios.

## [1.2.1] - 2026-09-16

### Corregido
- La aplicación fallaba al iniciar con `unknown option "-fg_color"` porque
  la ventana raíz de arrastrar y soltar no soportaba el tema de CustomTkinter.
  La raíz ahora es siempre de CustomTkinter y el soporte de arrastrar y
  soltar se carga sobre ella.

## [1.2.0] - 2026-09-16

### Añadido
- Tabla de resultados con filtros (todo / listos / pendientes / anomalías),
  buscador, pestaña de registro y botón de re-verificación.
- Detección de mapeos conflictivos en el Excel de correspondencia.
- Validación del contenido de los PDFs (archivos dañados o incompletos).
- Cancelación de procesos largos.
- Deshacer la última ejecución (solo copias registradas en el reporte).
- Arrastrar y soltar carpetas y archivos Excel.
- Suite de pruebas automatizadas (24) e integración continua con GitHub Actions.

### Cambiado
- La ejecución parcial ya no se bloquea por correspondencias faltantes:
  los archivos válidos se renombran y los demás quedan registrados en
  `pendientes_correccion.csv` con la acción sugerida.
- Los archivos que ya existen en el destino se omiten con aviso (re-ejecución).

## [1.1.0] - 2026-09-16

### Añadido
- Interfaz adaptable: pie de ventana que crece con su contenido, registro
  con desplazamiento nativo, tamaño inicial según el monitor y nitidez en
  pantallas con escalado.
- Mensajes de error explicados para problemas de Excel (archivo bloqueado
  por Excel/OneDrive, archivo inválido, plantilla vacía) con reintentos.
- Ejecución parcial con seguimiento de pendientes.
- Vista previa de la carpeta, validación del Excel al seleccionarlo,
  memoria de la sesión anterior y ayuda del formato de salida.

## [1.0.0] - 2026-09-16

### Añadido
- Primera versión estable: inventario, mapeo por Excel, plan de renombrado
  con patrón configurable, simulación previa, copia segura y reporte de
  trazabilidad.
- Interfaz gráfica y consola con la misma lógica.
- Ejecutable portable para Windows.
