import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { Plus, Trash2, Shirt, ClipboardList, RefreshCw, CheckCircle2, AlertTriangle, FileSpreadsheet, Upload } from "lucide-react";
import { motion } from "framer-motion";

function normalizarTexto(texto) {
  return String(texto || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function limpiarValor(valor) {
  if (valor === null || valor === undefined) return "";
  return String(valor).trim();
}

function esNumero(valor) {
  if (valor === null || valor === undefined || valor === "") return false;
  return !Number.isNaN(Number(String(valor).replace(",", ".")));
}

function obtenerNumero(valor) {
  return Number(String(valor).replace(",", "."));
}

function filaVacia(fila) {
  return !fila || fila.every((celda) => limpiarValor(celda) === "");
}

function contienePalabra(celda, palabras) {
  const texto = normalizarTexto(celda);
  return palabras.some((p) => texto.includes(p));
}

function detectarColumnasCabecera(fila) {
  const columnas = { articulo: -1, talla: -1, cantidad: -1 };
  fila.forEach((celda, idx) => {
    if (contienePalabra(celda, ["articulo", "artículo", "item", "producto", "prenda", "material", "descripcion", "descripción"])) columnas.articulo = idx;
    if (contienePalabra(celda, ["talla", "size", "numero", "número"])) columnas.talla = idx;
    if (contienePalabra(celda, ["cantidad", "cant", "stock", "unidades", "uds", "pedido", "solicitado", "disponible"])) columnas.cantidad = idx;
  });
  return columnas;
}

function extraerArticuloDesdeFila(fila) {
  const textos = fila.map(limpiarValor).filter(Boolean);
  const candidatos = textos.filter((t) => !esNumero(t) && !contienePalabra(t, ["talla", "cantidad", "stock", "pedido", "unidades", "uds"]));
  if (candidatos.length === 0) return "";
  return candidatos.join(" ").replace(/[:;]$/g, "").trim();
}

function parsearExcelPorBloques(rows) {
  const resultado = [];
  let articuloActual = "";
  let columnas = null;
  let filasLeidas = 0;

  for (let i = 0; i < rows.length; i++) {
    const fila = rows[i] || [];
    if (filaVacia(fila)) {
      columnas = null;
      continue;
    }

    const detectadas = detectarColumnasCabecera(fila);
    const esCabecera = detectadas.talla >= 0 && detectadas.cantidad >= 0;

    if (esCabecera) {
      columnas = detectadas;
      if (detectadas.articulo >= 0) {
        articuloActual = articuloActual || "Artículo";
      } else {
        const posibleArticulo = extraerArticuloDesdeFila(rows[i - 1] || []);
        if (posibleArticulo) articuloActual = posibleArticulo;
      }
      continue;
    }

    if (!columnas) {
      const posibleArticulo = extraerArticuloDesdeFila(fila);
      if (posibleArticulo) articuloActual = posibleArticulo;
      continue;
    }

    const talla = fila[columnas.talla];
    const cantidad = fila[columnas.cantidad];
    const articuloFila = columnas.articulo >= 0 ? limpiarValor(fila[columnas.articulo]) : "";
    const articulo = articuloFila || articuloActual || "Artículo";

    if (esNumero(talla) && esNumero(cantidad)) {
      resultado.push({ articulo, talla: obtenerNumero(talla), cantidad: obtenerNumero(cantidad) });
      filasLeidas++;
    }
  }

  return { resultado, filasLeidas };
}

async function leerExcel(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const hoja = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[hoja];
  const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
  return parsearExcelPorBloques(rows);
}

function agruparPorArticulo(lineas) {
  const grupos = {};
  for (const linea of lineas) {
    const articulo = normalizarTexto(linea.articulo);
    const talla = Number(linea.talla);
    const cantidad = Number(linea.cantidad);
    if (!articulo || Number.isNaN(talla) || Number.isNaN(cantidad)) continue;
    if (!grupos[articulo]) grupos[articulo] = [];
    grupos[articulo].push({ ...linea, articuloKey: articulo, talla, cantidad });
  }
  return grupos;
}

function calcularRecomendaciones(stock, pedidos, tolerancia) {
  const stockPorArticulo = agruparPorArticulo(stock);
  const pedidosPorArticulo = agruparPorArticulo(pedidos);
  const articulos = Array.from(new Set([...Object.keys(stockPorArticulo), ...Object.keys(pedidosPorArticulo)]));
  const recomendaciones = [];
  const resumen = [];

  for (const articulo of articulos) {
    const stockInicial = {};
    const pedidosInicial = {};

    for (const s of stockPorArticulo[articulo] || []) stockInicial[s.talla] = (stockInicial[s.talla] || 0) + s.cantidad;
    for (const p of pedidosPorArticulo[articulo] || []) pedidosInicial[p.talla] = (pedidosInicial[p.talla] || 0) + p.cantidad;

    const stockDisponible = { ...stockInicial };
    const tallasPedido = Object.keys(pedidosInicial).map(Number).sort((a, b) => a - b);
    const asignaciones = [];
    const pendientes = [];

    for (const tallaPedido of tallasPedido) {
      let falta = pedidosInicial[tallaPedido];
      const exacto = Math.min(falta, stockDisponible[tallaPedido] || 0);
      if (exacto > 0) {
        asignaciones.push({ articulo, tallaPedido, tallaEntregada: tallaPedido, cantidad: exacto, tipo: "Exacta" });
        stockDisponible[tallaPedido] -= exacto;
        falta -= exacto;
      }
      if (falta > 0) pendientes.push({ tallaPedido, cantidad: falta });
    }

    for (const pendiente of pendientes) {
      let falta = pendiente.cantidad;
      const candidatas = Object.keys(stockDisponible)
        .map(Number)
        .filter((tallaStock) => stockDisponible[tallaStock] > 0 && Math.abs(tallaStock - pendiente.tallaPedido) <= tolerancia)
        .sort((a, b) => {
          const da = Math.abs(a - pendiente.tallaPedido);
          const db = Math.abs(b - pendiente.tallaPedido);
          if (da !== db) return da - db;
          return b - a;
        });

      for (const tallaStock of candidatas) {
        if (falta <= 0) break;
        const usar = Math.min(falta, stockDisponible[tallaStock]);
        if (usar > 0) {
          asignaciones.push({ articulo, tallaPedido: pendiente.tallaPedido, tallaEntregada: tallaStock, cantidad: usar, tipo: "Sustitución" });
          stockDisponible[tallaStock] -= usar;
          falta -= usar;
        }
      }

      if (falta > 0) asignaciones.push({ articulo, tallaPedido: pendiente.tallaPedido, tallaEntregada: null, cantidad: falta, tipo: "Falta" });
    }

    const totalStock = Object.values(stockInicial).reduce((a, b) => a + b, 0);
    const totalPedido = Object.values(pedidosInicial).reduce((a, b) => a + b, 0);
    const totalEntregado = asignaciones.filter((a) => a.tipo !== "Falta").reduce((a, b) => a + b.cantidad, 0);
    const totalFalta = asignaciones.filter((a) => a.tipo === "Falta").reduce((a, b) => a + b.cantidad, 0);
    const sobrante = Object.values(stockDisponible).reduce((a, b) => a + b, 0);

    resumen.push({ articulo, totalStock, totalPedido, totalEntregado, totalFalta, sobrante, stockDisponible });
    recomendaciones.push(...asignaciones);
  }

  return { recomendaciones, resumen };
}

const stockEjemplo = [
  { articulo: "Pantalones", talla: 42, cantidad: 4 },
  { articulo: "Pantalones", talla: 43, cantidad: 3 },
  { articulo: "Pantalones", talla: 41, cantidad: 1 },
];

const pedidosEjemplo = [
  { articulo: "Pantalones", talla: 41, cantidad: 2 },
  { articulo: "Pantalones", talla: 43, cantidad: 4 },
  { articulo: "Pantalones", talla: 42, cantidad: 6 },
];

function Button({ children, variant = "primary", className = "", ...props }) {
  return <button className={`btn btn-${variant} ${className}`} {...props}>{children}</button>;
}

function Badge({ children, variant = "default" }) {
  return <span className={`badge badge-${variant}`}>{children}</span>;
}

function Card({ children, className = "" }) {
  return <div className={`card ${className}`}>{children}</div>;
}

function ImportadorExcel({ titulo, descripcion, onImportar, color }) {
  const [estado, setEstado] = useState("");

  const manejarArchivo = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      setEstado("Leyendo Excel...");
      const { resultado, filasLeidas } = await leerExcel(file);
      onImportar(resultado);
      setEstado(`${file.name}: ${filasLeidas} líneas importadas`);
    } catch (error) {
      console.error(error);
      setEstado("No se ha podido leer el Excel. Revisa que sea .xlsx/.xls y que tenga columnas de talla y cantidad.");
    } finally {
      event.target.value = "";
    }
  };

  return (
    <Card>
      <div className="card-content import-card">
        <div className={`icon-box ${color}`}><FileSpreadsheet size={22} /></div>
        <div className="grow">
          <h2>{titulo}</h2>
          <p>{descripcion}</p>
          <label className="upload-label">
            <Upload size={17} /> Seleccionar Excel
            <input type="file" accept=".xlsx,.xls" onChange={manejarArchivo} />
          </label>
          {estado && <p className="status">{estado}</p>}
        </div>
      </div>
    </Card>
  );
}

function EditorTabla({ titulo, icono: Icono, filas, setFilas, color }) {
  const actualizar = (idx, campo, valor) => setFilas((prev) => prev.map((fila, i) => (i === idx ? { ...fila, [campo]: valor } : fila)));
  const agregar = () => setFilas((prev) => [...prev, { articulo: "Pantalones", talla: "", cantidad: "" }]);
  const borrar = (idx) => setFilas((prev) => prev.filter((_, i) => i !== idx));

  return (
    <Card>
      <div className="card-content">
        <div className="section-header">
          <div className="section-title"><div className={`icon-box ${color}`}><Icono size={22} /></div><h2>{titulo}</h2></div>
          <Button onClick={agregar}><Plus size={16} /> Añadir</Button>
        </div>

        <div className="table-grid table-head"><div>Artículo</div><div>Talla</div><div>Cantidad</div><div></div></div>
        <div className="editor-list">
          {filas.map((fila, idx) => (
            <div key={idx} className="table-grid table-row">
              <input value={fila.articulo} onChange={(e) => actualizar(idx, "articulo", e.target.value)} />
              <input type="number" value={fila.talla} onChange={(e) => actualizar(idx, "talla", e.target.value)} />
              <input type="number" value={fila.cantidad} onChange={(e) => actualizar(idx, "cantidad", e.target.value)} />
              <Button variant="ghost" onClick={() => borrar(idx)}><Trash2 size={16} /></Button>
            </div>
          ))}
        </div>
      </div>
    </Card>
  );
}

export default function App() {
  const [stock, setStock] = useState(stockEjemplo);
  const [pedidos, setPedidos] = useState(pedidosEjemplo);
  const [tolerancia, setTolerancia] = useState(1);

  const { recomendaciones, resumen } = useMemo(() => calcularRecomendaciones(stock, pedidos, Number(tolerancia) || 0), [stock, pedidos, tolerancia]);

  const resetear = () => {
    setStock(stockEjemplo);
    setPedidos(pedidosEjemplo);
    setTolerancia(1);
  };

  return (
    <main className="app">
      <div className="container">
        <motion.header initial={{ opacity: 0, y: 8 }} animate={{ opacity: 1, y: 0 }} className="hero">
          <div>
            <Badge variant="muted">Prototipo de recomendador</Badge>
            <h1>Gestión de vestuario desde dos Excel</h1>
            <p>Sube un Excel de stock y otro de pedidos. La app lee artículos por bloques, agrupa por prenda y talla, y recomienda entregas exactas o sustituciones con tallas cercanas.</p>
          </div>
          <div className="controls">
            <label>Tolerancia de talla<input type="number" min="0" value={tolerancia} onChange={(e) => setTolerancia(e.target.value)} /></label>
            <Button variant="outline" onClick={resetear}><RefreshCw size={16} /> Ejemplo</Button>
          </div>
        </motion.header>

        <section className="grid two">
          <ImportadorExcel titulo="Importar Excel de stock" descripcion="Lee bloques de material disponible: artículo, talla y cantidad/stock." onImportar={setStock} color="green" />
          <ImportadorExcel titulo="Importar Excel de pedidos" descripcion="Lee bloques de pedidos solicitados: artículo, talla y cantidad/pedido." onImportar={setPedidos} color="blue" />
        </section>

        <Card>
          <div className="card-content info">
            <h2>Formato de Excel recomendado</h2>
            <p>Puede ser por bloques: una fila con el artículo, por ejemplo <b>Pantalones</b>, después cabecera <b>Talla</b> y <b>Cantidad</b>, y debajo las líneas. También acepta tabla normal con columnas <b>Artículo</b>, <b>Talla</b> y <b>Cantidad</b>.</p>
          </div>
        </Card>

        <section className="grid two">
          <EditorTabla titulo="Stock disponible" icono={Shirt} filas={stock} setFilas={setStock} color="green" />
          <EditorTabla titulo="Pedidos solicitados" icono={ClipboardList} filas={pedidos} setFilas={setPedidos} color="blue" />
        </section>

        <section className="grid results">
          <Card>
            <div className="card-content">
              <h2>Resumen económico</h2>
              <div className="summary-list">
                {resumen.map((r) => (
                  <div key={r.articulo} className="summary-box">
                    <h3>{r.articulo}</h3>
                    <div><span>Stock total</span><b>{r.totalStock}</b></div>
                    <div><span>Pedido total</span><b>{r.totalPedido}</b></div>
                    <div><span>Entregado</span><b className="ok">{r.totalEntregado}</b></div>
                    <div><span>Falta</span><b className="bad">{r.totalFalta}</b></div>
                    <div><span>Sobrante</span><b>{r.sobrante}</b></div>
                  </div>
                ))}
              </div>
            </div>
          </Card>

          <Card className="recommendations-card">
            <div className="card-content">
              <div className="section-header"><h2>Recomendaciones de entrega</h2><Badge variant="outline">Exactas + sustituciones</Badge></div>
              <div className="recommendations">
                <div className="rec-grid rec-head"><div>Artículo</div><div>Pedido</div><div>Entregar</div><div>Cantidad</div><div>Resultado</div></div>
                {recomendaciones.length === 0 ? <div className="empty">No hay datos suficientes para calcular.</div> : recomendaciones.map((rec, idx) => (
                  <div key={idx} className="rec-grid rec-row">
                    <div>{rec.articulo}</div>
                    <div>Talla {rec.tallaPedido}</div>
                    <div>{rec.tallaEntregada ? `Talla ${rec.tallaEntregada}` : "—"}</div>
                    <div><b>{rec.cantidad}</b></div>
                    <div>
                      {rec.tipo === "Exacta" && <Badge variant="success"><CheckCircle2 size={13} /> Exacta</Badge>}
                      {rec.tipo === "Sustitución" && <Badge variant="warning"><AlertTriangle size={13} /> Sustituir</Badge>}
                      {rec.tipo === "Falta" && <Badge variant="danger"><AlertTriangle size={13} /> Comprar / pendiente</Badge>}
                    </div>
                  </div>
                ))}
              </div>
              <p className="note">Regla actual: primero se cubre la misma talla. Si falta stock, se usan tallas cercanas dentro de la tolerancia. En empate se prioriza la talla superior.</p>
            </div>
          </Card>
        </section>
      </div>
    </main>
  );
}
