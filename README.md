# App de Gestión de Vestuario

Aplicación web en React para comparar dos Excel:

- Excel de stock disponible.
- Excel de pedidos solicitados.

La app lee artículos por bloques, agrupa por artículo y talla, y recomienda:

- entrega exacta;
- sustitución por talla cercana;
- faltante para comprar o dejar pendiente;
- sobrante;
- resumen económico por artículo.

## Formato de Excel recomendado

Formato por bloques:

```text
Pantalones

Talla | Cantidad
41    | 1
42    | 4
43    | 3

Camisas

Talla | Cantidad
38    | 2
39    | 5
```

También acepta una tabla normal:

```text
Artículo   | Talla | Cantidad
Pantalones | 41    | 1
Pantalones | 42    | 4
```

## Cómo usar en local

```bash
npm install
npm run dev
```

## Cómo generar versión para publicar

```bash
npm run build
```

La carpeta generada será `dist`.

## Archivos principales

- `src/App.jsx`: lógica de lectura de Excel y cálculo de recomendaciones.
- `src/index.css`: estilos.
- `package.json`: dependencias y comandos.
