Proyecto: Sistema de Gestión de Máquinas e Inspecciones

Descripción del problema:
Actualmente los datos de inspecciones de máquinas están en Excel y se registran manualmente. Esto genera duplicación de información, dificultad para mantener el historial, problemas para consultar datos por máquina y falta de automatización en el cálculo de la condición. Se necesita digitalizar y automatizar este proceso.

Objetivo:
Crear una aplicación de escritorio que permita registrar máquinas, realizar inspecciones utilizando la matriz de evaluación, calcular automáticamente la condición y prioridad, y llevar un historial actualizado para cada máquina.

Funciones esperadas:

Gestión de máquinas: crear, editar, eliminar y consultar

Registro de inspecciones con matriz del Excel

Cálculo automático de puntaje y criticidad

Guardado y consulta del historial por máquina

Exportación de datos (opcional)

Usuarios y roles (opcional)

Restricciones:

Tiempo limitado para desarrollo

Se prioriza un MVP funcional

Mantenimiento simple y escalabilidad futura

Tipo de aplicación:
Aplicación de escritorio en Windows (posible expansión futura a Web).

Tecnologías disponibles:
Frontend de escritorio con PySide6
Backend con FastAPI
Base de datos SQLite inicialmente (posible migración a PostgreSQL)
Lenguaje Python 3.10 o superior

Opciones para hacer una página web con Python:

Django: framework completo para web con panel admin

Flask: microframework flexible para proyectos chicos

FastAPI: ideal para crear APIs rápidas y documentadas

Opciones para hacer el backend con Python (se revisan dos):

FastAPI: arquitectura limpia, alta performance, documentación automática, ideal para separar UI y backend

Django: incluye autenticación, ORM y administración, ideal si fuese proyecto web completo

Arquitectura limpia aplicada:
Separación por capas para facilitar mantenimiento y escalabilidad:

Capa UI: interfaz gráfica en PySide6

Capa Aplicación: casos de uso y control del flujo

Capa Dominio: entidades y reglas de negocio (cálculo de matriz)

Capa Infraestructura: FastAPI, base de datos y repositorios

Lógica de negocio principal:
La inspección evalúa criterios como diseño, patologías internas, inspección externa, tiempo sin revisar y exposición. Cada opción asigna un factor. La suma determina prioridad de mantenimiento. Esta lógica se implementa en la capa de dominio.

Plan básico de desarrollo (MVP):
1: CRUD de máquinas
2: Registro de inspecciones con cálculo automático
3: Historial por máquina con visualización de datos
4: Mejoras opcionales como exportación o roles

Motivo de la arquitectura:
Permite mantenimiento ágil, crecimiento futuro y mantener desacopladas la interfaz y la lógica, pudiendo convertirse luego en aplicación web sin rehacer el sistema.

Estado del proyecto:
En fase de definición y armado de prototipo.


