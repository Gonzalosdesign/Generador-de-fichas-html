# Generador de Fichas HTML para Docentes y Funcionarios

## Descripción

Este proyecto fue desarrollado como trabajo final del curso de Introducción a la IA generativa en la comunicación.

La aplicación permite automatizar la generación de fichas HTML para docentes y funcionarios a partir de información estructurada ingresada por el usuario en un Google doc, reduciendo significativamente el tiempo necesario para crear y actualizar perfiles institucionales en sitios web.

## Funcionamiento

La herramienta cuenta con una interfaz gráfica desarrollada en Python mediante Tkinter.

El usuario selecciona los archivos de entrada necesarios y completa la información requerida. A partir de esos datos, el sistema:

1. Descargar el archivo .exe desde la carpeta /dist
2. Compartir el doc Nueva Ficha con el docente o funcionario para que complete con sus datos
3. Desde Uneam se descarga el doc en formato DOCX
4. Se ejecuta la aplicación
5. Se selecciona el o los archivos docx a transformar en html
6. El programa procesa la información proporcionada.
7. Aplica una plantilla HTML predefinida.
8. Inserta automáticamente los datos correspondientes en cada campo.
9. Genera una ficha HTML individual para cada persona.
10. Guarda los archivos resultantes en la carpeta Salida y abre dicha carpeta para comodidad del usuario.
11. Copiar el contenido del archivo html en el espacio correspondiente en el perfil del funcionario en woerdpress

De esta forma se evita la edición manual de cada ficha, disminuyendo errores y mejorando la eficiencia del proceso.

## Tecnologías utilizadas

- Python
- Tkinter
- HTML
- Git
- GitHub

## Uso de Inteligencia Artificial

Durante el desarrollo del proyecto se utilizaron herramientas de Inteligencia Artificial generativa como apoyo para:

- Diseño y planificación de la arquitectura de la aplicación.
- Resolución de problemas de programación.
- Generación y optimización de fragmentos de código.
- Revisión y depuración de errores.
- Elaboración de documentación técnica.

La concepción, validación, adaptación e integración de las soluciones propuestas fueron realizadas por el autor del proyecto, utilizando la IA como herramienta de asistencia al desarrollo.

## Autor

Proyecto desarrollado por Gonzalo Souto como trabajo final de curso.
Uneam - Fagro - UdelaR