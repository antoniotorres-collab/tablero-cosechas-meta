# Cosechas Claude

## Descripción del proyecto

Análisis de cosechas semanales de asesores financieros.

## Estructura de carpetas

```
/datos      — Archivos fuente xlsx con los datos de cosechas
/reportes   — Reportes generados
/graficas   — Gráficas generadas
```

## Archivo base

El archivo más reciente es `Cosechas1Jun26.xlsx` (ubicado en `/datos`).

## Convención de archivos

- **Archivos CH** (ej. `CosechasCH_...xlsx`): valores puros, sin fórmulas
- **Archivos sin CH** (ej. `Cosechas_...xlsx`): valores formulados (con fórmulas Excel)

## Columna de identificación del asesor

El campo ID del asesor cambia de nombre según el archivo:

| Tipo de archivo | Nombre de columna |
|---|---|
| Archivos recientes (sin CH) | `Clave Asesor` |
| Archivos CH antiguos | `Clave_Asesor_Actual` |

Siempre verificar cuál columna existe en el archivo antes de filtrar o agrupar por asesor.
