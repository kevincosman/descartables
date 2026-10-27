# Comparador de proveedores de descartables

Buscador que junta las listas de precios de varios proveedores y las compara por precio unitario, con y sin IVA.

## Qué hay en cada lugar

- `LISTA *.xlsx`, `LISTA *.pdf`: listas originales de los proveedores.
- `comparador/procesar.js`: lee las listas y arma la página con los precios adentro.
- `comparador/lectores.js`: cómo se lee cada formato (Excel/CSV por columnas, PDF de Vassoi). Lo usan el procesador y la página.
- `comparador/plantilla.html`: la página (buscador, comparación, carga de listas, precios raros).
- `comparador/comparador.html`: la página ya armada (versión con carga de listas, publicada en claude.ai).
- `comparador/publicar/index.html`: copia solo para consultar, lista para subir a Netlify.

## Actualizar los precios

```
cd comparador
npm install      # una sola vez
node procesar.js
```

Las reglas de cada proveedor (IVA, descuento, columnas) están al principio de `procesar.js`.
Bolsatodo se lee directo de su Google Sheets.
