# Panel Administrativo CRUD - Excel/VBA (StreamDB)

Aplicacion de escritorio en Excel con UserForms y VBA que funciona como panel administrativo, conectado directamente a la base de datos SQL Server del proyecto StreamDB (PELICULAS_SERIES).

## Contexto

Proyecto academico complementario a StreamDB, desarrollado en un curso independiente enfocado en automatizacion con VBA y conexion a bases de datos externas desde Excel.

## Funcionalidad

Panel con operaciones CRUD completas (Crear, Leer, Actualizar, Eliminar) sobre las siguientes entidades:

- **Catalogo** - gestion de peliculas/series
- **Usuarios** - gestion de cuentas de usuario
- **Suscripciones** - gestion de planes y suscripciones activas

Ademas incluye un **modulo de reportes**, con una consulta SQL directa que muestra la cantidad de titulos disponibles por genero.

## Arquitectura

- **Frontend:** UserForms de VBA (Excel) como interfaz de captura y visualizacion
- **Backend:** SQL Server (misma base de datos `PELICULAS_SERIES` usada en StreamDB)
- **Conexion:** VBA ejecuta consultas SQL directamente contra el servidor mediante ADO

## Tecnologias

- Excel con macros habilitadas (.xlsm)
- VBA (Visual Basic for Applications)
- SQL Server (instancia `HowL\SQLEXPRESS`)
- ADO (ActiveX Data Objects) para la conexion

## Relacion con StreamDB

Este panel reutiliza la misma base de datos SQL Server que el proyecto principal StreamDB, ofreciendo una interfaz administrativa alternativa (fuera del stack Flask/API) para gestionar los datos del catalogo directamente desde Excel.

Repositorio principal: [StreamDB](https://github.com/Julien009/StreamDB)

## Autor

**Walter Jason Garcia Tito**
Estudiante de Big Data y Ciencia de Datos - Instituto Tecsup
