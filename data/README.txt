Coloca aquí el Excel del tablero y nómbralo seguimiento_pad.xlsx.
La hoja de datos debe llamarse Control.
La columna G puede contener Aprobado en segunda instancia; G2 alimenta la
tarjeta APROBADO/COMPLETO. Si G2 está vacío, se cuentan los valores
afirmativos de la columna G.
La columna Virtualizacion se lee desde I si su encabezado está en I1, desde H
si está en H1 o, en archivos anteriores, desde G si está en G1. La tarjeta
VIRTUALIZACIÓN usa el total de la fila 2 (I2 si Virtualizacion está en I) y
la tabla de detalles muestra el valor de cada PAD de esa columna.
Este archivo se usa como fuente local cuando Supabase no está configurado. En producción, publica los cambios desde el botón Actualizar desde Excel del dashboard.
