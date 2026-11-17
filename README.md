# Registro de Inventario

App para registrar el conteo de patio (Impregnación, Polineras, Postes y Morbark/Cilindrado).
Frontend estático (GitHub Pages) + Google Apps Script como puente hacia el Google Sheet.

## Archivos

- index.html: la app completa (4 vistas de registro + Mantenedor con clave).
- Code.gs: backend que se pega en el Google Sheet (Extensiones > Apps Script).
- README.md: estos pasos.

## Google Sheet

Hoja "Registro Inventario - Patio".
- Registros: una fila por línea de conteo. Columnas: Timestamp, Vista, Fecha, Jefe de Área, Trabajador 01 a 10, Producto, Clase, Largo, Diámetro, Unidades Producción, Unidades Recuperación, Unidades Rechazo, Nivel de Secado, Sector, Comentario.
- Mantenedor: listas desplegables (Jefes, Productos, Clases, Largos, Diámetros por producto, Nivel de secado).

## Apps Script

1. En el Sheet: Extensiones > Apps Script. Borra el contenido y pega Code.gs. Guarda.
2. Configuración del proyecto > Propiedades de la secuencia de comandos > Agregar propiedad: CLAVE_MANTENEDOR = la clave que elijas.
3. Configuración del proyecto > Zona horaria: (GMT-03:00) Santiago.
4. Implementar > Nueva implementación > Tipo: Aplicación web. Ejecutar como: Yo. Quién tiene acceso: Cualquier persona.
5. Copia la URL que termina en /exec.

Cada cambio en Code.gs requiere Implementar > Administrar implementaciones > Editar > Nueva versión.

## Conectar la app

En index.html reemplaza PEGAR_AQUI_LA_URL_DEL_WEB_APP por la URL /exec copiada.

## GitHub Pages

Repositorio registro-inventario (privado). Settings > Pages: rama main, carpeta raíz.

## Notas de uso

- Fecha, jefe y trabajadores se mantienen después de guardar; los campos de la línea se limpian.
- Producción, recuperación y rechazo son opcionales, pero al menos una debe tener cantidad.
- Producto, clase, largo, diámetro y secado aceptan valores de la lista o valores nuevos escritos a mano.
- Sector y comentario son texto libre.
- La clave del mantenedor vive solo en las propiedades del script, no en el código ni en el navegador.
- La bitácora de quién registra queda pendiente para una etapa posterior.
