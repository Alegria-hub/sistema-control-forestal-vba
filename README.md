# Sistema de Control Forestal e Inventarios en Excel + VBA

Proyecto de portafolio desarrollado en **Microsoft Excel y VBA** para registrar movimientos de madera, inventarios físicos, inventarios base derivados de inspecciones y balances históricos por empresa y especie.

> **Estado:** en desarrollo activo.

## Objetivo

Construir una solución práctica para centralizar el control de inventarios y movimientos forestales, reducir errores de captura, conservar trazabilidad histórica y preparar la información para análisis posteriores con Power Query, Power BI y SQL.

## Funcionalidades actuales

- Registro de movimientos de **Entrada** y **Salida**.
- Arquitectura **maestro-detalle** para movimientos con una o varias especies.
- Control de movimientos **vigentes y cancelados** sin eliminar el historial.
- Generación automática de **IDs persistentes** para registros.
- Formularios VBA para:
  - Movimientos.
  - Inventario base / inspección.
  - Inventario físico.
  - Generación de balance.
- Validaciones reutilizables para campos obligatorios, fechas, números positivos y medidas.
- Conversión de medidas de madera a **pies tabla** y **m³**.
- Acumulación de múltiples filas de inventario físico por empresa, especie y fecha.
- Cálculo de entradas y salidas dentro de un periodo seleccionado.
- Generación de **cierres históricos** de balance sin sobrescribir balances ya generados.

## Regla principal del balance

El balance se calcula con la siguiente fórmula:

```text
(Inventario inspeccionado + Entradas)
-
(Inventario físico + Salidas)
```

Cuando existe una inspección previa, los movimientos considerados cumplen:

```text
FECHA_MOVIMIENTO >= FECHA_INSPECCION
FECHA_MOVIMIENTO <= FECHA_BALANCE
```

Cuando no existe inspección previa, el inventario base parte de `0` y se toman los movimientos disponibles hasta la fecha del balance.

## Estructura del libro

El proyecto está organizado en hojas con responsabilidades separadas:

```text
MENU
CATALOGOS
MOVIMIENTOS
DETALLE_MOV
INVENTARIO_FISICO
INVENTARIO_BASE
BALANCE
REPORTES
PARAMETROS
```

## Tecnologías y conceptos aplicados

- Microsoft Excel
- VBA
- UserForms
- ListObject / tablas estructuradas
- Diccionarios (`Scripting.Dictionary`)
- Arquitectura maestro-detalle
- Validaciones reutilizables
- Manejo de IDs
- Reglas de negocio
- Trazabilidad de datos
- Preparación para Power Query / Power BI / SQL

## Arquitectura VBA

El proyecto separa responsabilidades en módulos y formularios. Entre los componentes principales se encuentran:

```text
modIDs
modValidaciones
modMenu
frmMovimiento
frmInventarioBase
frmInventarioFisico
frmBalance
```

El código fuente VBA se irá exportando al repositorio en archivos `.bas`, `.frm` y `.cls` para permitir control de versiones y revisión del código fuera del archivo `.xlsm`.

## Próximos pasos

- Exportar módulos y formularios VBA al repositorio.
- Crear una versión DEMO del archivo con datos ficticios.
- Documentar el modelo de datos y las reglas de negocio.
- Completar pruebas integrales del módulo de balance.
- Incorporar reportes y generación de PDF.
- Añadir consultas con Power Query.
- Construir visualizaciones en Power BI.
- Evaluar una futura migración de datos a SQL.

## Privacidad de los datos

Este repositorio está pensado como **proyecto de portafolio**. Los datos publicados serán ficticios o anonimizados. No se incluirán nombres reales de empresas, personas, folios, códigos, remisiones ni información interna de terceros.

## Autor

Proyecto desarrollado por **Jaime Alegría** como parte de su portafolio de automatización, análisis de datos y desarrollo con Excel/VBA.
