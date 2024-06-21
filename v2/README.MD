Generador de Certificados

Esta aplicación permite generar certificados personalizados a partir de una plantilla de documento de Word y una lista de nombres.

Requisitos:
- Java 8 o superior
- Apache POI
- Launch4j (para convertir el archivo JAR en un EXE)

Instrucciones:
1. Configuración del Proyecto:
   - Asegúrate de que todas las dependencias estén incluidas en tu archivo `pom.xml` si usas Maven.

2. Construir el Proyecto:
   - Ejecuta `mvn clean package` para construir el archivo JAR con todas las dependencias.

3. Convertir el JAR en un EXE:
   - Usa Launch4j para convertir el archivo JAR en un archivo EXE.

4. Ejecutar la Aplicación:
   - Ejecuta el archivo EXE generado para abrir la aplicación.

Uso:
1. Selecciona la plantilla de documento de Word.
2. Selecciona el archivo de texto que contiene los nombres.
3. Haz clic en "Generar Certificados" para crear los certificados.

Nota:
Asegúrate de que la fuente `Edwardian Script ITC` esté instalada en tu sistema para que los certificados se generen correctamente.
