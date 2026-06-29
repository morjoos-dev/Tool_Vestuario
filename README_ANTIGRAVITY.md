# Contexto del proyecto: App de Gestión de Vestuario

## Objetivo principal

Crear una aplicación web sencilla para ayudar en la gestión de vestuario laboral a partir de dos archivos Excel:

1. **Excel de stock disponible**: contiene el material disponible en almacén, separado por artículo, talla y cantidad.
2. **Excel de pedidos solicitados**: contiene los pedidos realizados por trabajadores o centros, también separados por artículo, talla y cantidad.

La finalidad de la aplicación es comparar ambos Excel y recomendar cómo repartir el material disponible, teniendo en cuenta que un mismo artículo puede tener diferentes tallas y que, en algunos casos, se puede compensar una talla faltante con otra talla cercana.

---

## Caso de uso explicado

Ejemplo de stock disponible:

```text
Pantalones talla 42 -> cantidad 4
Pantalones talla 43 -> cantidad 3
Pantalones talla 41 -> cantidad 1
```

Ejemplo de pedidos:

```text
Pantalones talla 41 -> cantidad 2
Pantalones talla 43 -> cantidad 4
Pantalones talla 42 -> cantidad 6
```

En este caso:

- Stock total de pantalones: 8 unidades.
- Pedido total de pantalones: 12 unidades.
- Por talla concreta faltan unidades en 41, 42 y 43.
- La aplicación debe intentar cubrir primero con la talla exacta.
- Si no hay suficiente talla exacta, debe recomendar sustituciones con tallas cercanas, por ejemplo entregar talla 43 a alguien que pidió talla 42, si la tolerancia lo permite.
- Si aun así no alcanza el stock total, debe marcar la diferencia como pendiente, falta o compra necesaria.

---

## Problema que resuelve

Actualmente, el control económico o de inventario puede contar todos los pantalones como el mismo tipo de material, aunque estén divididos por tallas. Sin embargo, en la entrega real, la talla importa.

La app debe resolver ambas visiones:

1. **Visión logística**: qué talla entregar a cada pedido.
2. **Visión económica**: cuántos pantalones se han consumido en total, aunque sean de tallas diferentes.

Por tanto, la app debe permitir compensar negativos y positivos entre tallas del mismo artículo, siempre siguiendo reglas configurables.

---

## Reglas de negocio actuales

### 1. Agrupación por artículo

Los artículos deben agruparse por nombre ignorando diferencias menores de escritura, mayúsculas o acentos.

Ejemplos equivalentes:

```text
Pantalones
pantalones
PANTALONES
```

Todos deben tratarse como el mismo artículo.

---

### 2. Asignación exacta primero

La aplicación debe cubrir primero los pedidos con stock de la misma talla.

Ejemplo:

```text
Pedido: pantalones talla 42 -> 6 unidades
Stock: pantalones talla 42 -> 4 unidades
```

Resultado inicial:

```text
Entregar talla 42 -> 4 unidades
Faltan talla 42 -> 2 unidades
```

---

### 3. Sustitución por talla cercana

Después de cubrir tallas exactas, la app debe buscar tallas cercanas disponibles dentro de una tolerancia configurable.

Ejemplo con tolerancia 1:

```text
Pedido talla 42
Puede sustituirse por talla 41 o talla 43
```

Ejemplo con tolerancia 2:

```text
Pedido talla 42
Puede sustituirse por talla 40, 41, 43 o 44
```

La tolerancia debe ser configurable en la interfaz.

---

### 4. Prioridad en sustituciones

Cuando existan varias tallas candidatas, la regla actual es:

1. Priorizar la talla más cercana.
2. Si hay empate, priorizar la talla superior.

Ejemplo:

```text
Pedido talla 42
Disponibles talla 41 y talla 43
Ambas están a distancia 1
Se recomienda primero talla 43
```

Esta regla puede cambiarse en el futuro si el usuario prefiere priorizar talla inferior o aplicar reglas distintas por tipo de prenda.

---

### 5. Faltantes

Si no hay talla exacta ni talla compatible suficiente, la app debe marcar la cantidad restante como:

```text
Comprar / pendiente / falta
```

---

### 6. Sobrantes

Después de todas las asignaciones, la app debe mostrar qué stock queda sin usar.

---

## Datos de entrada

La aplicación trabaja con dos Excel:

### Excel 1: Stock disponible

Debe contener el inventario actual.

Campos esperados:

```text
Artículo | Talla | Cantidad
```

También puede venir dividido por bloques.

Ejemplo por bloques:

```text
Pantalones

Talla | Cantidad
41    | 1
42    | 4
43    | 3

Camisas

Talla | Cantidad
M     | 5
L     | 2
```

### Excel 2: Pedidos solicitados

Debe contener las necesidades o pedidos.

Campos esperados:

```text
Artículo | Talla | Cantidad
```

También puede venir dividido por bloques.

Ejemplo:

```text
Pantalones

Talla | Cantidad
41    | 2
42    | 6
43    | 4
```

---

## Formatos que debería reconocer la app

La app debe intentar reconocer columnas con nombres similares.

### Para artículo

Posibles nombres:

```text
Artículo
Articulo
Item
Producto
Prenda
Material
Descripción
Descripcion
```

### Para talla

Posibles nombres:

```text
Talla
Size
Número
Numero
```

### Para cantidad

Posibles nombres:

```text
Cantidad
Cant
Stock
Unidades
Uds
Pedido
Solicitado
Disponible
```

---

## Resultados esperados

La app debe mostrar una tabla de recomendaciones con estas columnas:

```text
Artículo | Talla pedida | Talla a entregar | Cantidad | Resultado
```

Valores posibles de resultado:

```text
Exacta
Sustitución
Comprar / pendiente
```

Ejemplo:

```text
Pantalones | Talla 41 | Talla 41 | 1 | Exacta
Pantalones | Talla 41 | Talla 42 | 1 | Sustitución
Pantalones | Talla 42 | Talla 42 | 4 | Exacta
Pantalones | Talla 42 | Talla 43 | 2 | Sustitución
Pantalones | Talla 43 | Talla 43 | 1 | Exacta
Pantalones | Talla 43 | —        | 3 | Comprar / pendiente
```

El ejemplo anterior es orientativo; el resultado real depende del algoritmo de asignación y del orden de prioridad.

---

## Resumen económico esperado

Además de las recomendaciones, la app debe mostrar un resumen por artículo:

```text
Artículo
Stock total
Pedido total
Entregado total
Faltante total
Sobrante total
```

Ejemplo:

```text
Pantalones
Stock total: 8
Pedido total: 12
Entregado total: 8
Faltante total: 4
Sobrante total: 0
```

---

## Funcionalidades actuales del proyecto

La versión actual es una app React creada con Vite.

Funcionalidades incluidas:

- Carga de Excel de stock.
- Carga de Excel de pedidos.
- Lectura básica de Excel con la librería `xlsx`.
- Detección de datos en formato tabla o por bloques.
- Edición manual de las líneas importadas.
- Cálculo de recomendaciones.
- Tolerancia configurable de talla.
- Resumen económico por artículo.
- Visualización de entregas exactas, sustituciones y faltantes.

---

## Estructura técnica actual

Archivos principales:

```text
index.html
package.json
README.md
src/main.jsx
src/App.jsx
src/index.css
```

La lógica principal está en:

```text
src/App.jsx
```

Estilos principales:

```text
src/index.css
```

---

## Comandos del proyecto

Instalar dependencias:

```bash
npm install
```

Ejecutar en local:

```bash
npm run dev
```

Generar build:

```bash
npm run build
```

Vista previa del build:

```bash
npm run preview
```

---

## Mejoras futuras deseadas

### 1. Exportar resultados a Excel

Añadir un botón para exportar:

- recomendaciones de entrega;
- resumen económico;
- faltantes;
- sobrantes.

Formato sugerido:

```text
resultado_recomendaciones.xlsx
```

Con varias hojas:

```text
Recomendaciones
Resumen
Faltantes
Sobrantes
```

---

### 2. Configurar reglas por artículo

No todas las prendas deberían tratarse igual.

Ejemplos:

- Pantalones: permitir tolerancia de ±1 talla.
- Camisas: permitir equivalencias por S, M, L, XL.
- Zapatos: quizás no permitir sustitución.
- Cazadoras: permitir talla superior, pero no inferior.

---

### 3. Compatibilidad con tallas no numéricas

Actualmente la lógica está pensada principalmente para tallas numéricas.

Se debería mejorar para tallas como:

```text
XS
S
M
L
XL
XXL
```

Y para tallas combinadas como:

```text
38/40
40/42
42/44
```

---

### 4. Mejor detección de bloques Excel reales

Los Excel reales pueden tener títulos, cabeceras irregulares, celdas vacías, texto adicional o bloques con formatos diferentes.

Se debe mejorar el parser para adaptarlo al formato exacto de los Excel del usuario.

---

### 5. Guardar configuración

Guardar en navegador:

- tolerancia de talla;
- reglas por artículo;
- preferencias de sustitución;
- último formato detectado.

Se puede usar `localStorage`.

---

### 6. Validaciones

Añadir avisos si:

- una fila no tiene artículo;
- una talla no es válida;
- una cantidad no es numérica;
- se importan cero líneas;
- el Excel no tiene columnas reconocibles;
- hay artículos en pedidos que no existen en stock.

---

### 7. Mejor interfaz de revisión

Añadir filtros por:

- artículo;
- tipo de resultado;
- talla pedida;
- talla entregada;
- solo faltantes;
- solo sustituciones.

---

### 8. Despliegue web

El proyecto está subido a GitHub y podría desplegarse en:

- Vercel;
- Netlify;
- GitHub Pages;
- servidor interno.

---

## Instrucciones para Antigravity

Cuando trabajes en este proyecto, ten en cuenta lo siguiente:

1. Mantener la app sencilla y usable por personas no técnicas.
2. Priorizar lectura correcta de Excel y claridad en las recomendaciones.
3. No eliminar la edición manual de datos, porque sirve para corregir importaciones incorrectas.
4. Mantener la lógica de negocio separada y fácil de modificar.
5. Añadir comentarios en las funciones complejas.
6. Evitar dependencias innecesarias.
7. Si se modifica el parser de Excel, probar con ejemplos de tabla normal y ejemplos por bloques.
8. Si se modifica el algoritmo de reparto, documentar claramente la regla usada.
9. Preparar la app para exportar resultados a Excel en una fase posterior.
10. Mantener todos los textos de interfaz en español.

---

## Prioridad actual del desarrollo

La prioridad actual es convertir el prototipo en una herramienta práctica para uso real:

1. Confirmar que la app lee correctamente los dos Excel reales.
2. Ajustar el parser a la estructura real de los bloques.
3. Mejorar el algoritmo de sustitución si el usuario define reglas más concretas.
4. Añadir exportación a Excel de resultados.
5. Preparar despliegue para que otras personas puedan usar la app desde navegador.

---

## Resumen corto para agentes IA

Esta app compara un Excel de stock y un Excel de pedidos de vestuario laboral. Los artículos están separados por tallas, pero a nivel económico se deben contar como el mismo artículo. La app asigna primero stock de la misma talla, luego propone sustituciones con tallas cercanas según una tolerancia configurable y finalmente marca faltantes o sobrantes. Debe leer Excel en formato tabla o por bloques, mostrar recomendaciones claras y permitir exportar resultados en el futuro.
