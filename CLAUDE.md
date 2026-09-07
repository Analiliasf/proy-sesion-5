# Qué es

Simulador para implementación de WMS (Warehouse Management System).

# Cómo se trabaja

- **Excel** como motor del simulador: fórmulas, tablas dinámicas y/o macros VBA para modelar el flujo del WMS (recepción, almacenamiento, picking, etc.)
- **PowerShell** (`New-Object -ComObject Excel.Application`) para automatizar la validación de los Excel de entrada y la generación de reportes antes de procesar
- Verificación del resultado: correr el reporte previo de hallazgos/inconsistencias y revisarlo antes de dar por bueno cualquier cálculo del simulador

# Convenciones

- Formato de moneda: contabilidad, pesos mexicanos, sin decimales

# Qué evitar

- Nunca modificar los datos originales del Excel
- Antes de procesar, validar el Excel (anomalías, inconsistencias, errores de captura)
- Si hay más de un Excel, encontrar la relación entre archivos y generar un reporte previo de hallazgos y recomendaciones antes de procesar
