import React, { useMemo, useState, useEffect } from "react";
import * as XLSX from "xlsx";
import { 
  Plus, 
  Trash2, 
  Shirt, 
  ClipboardList, 
  RefreshCw, 
  CheckCircle2, 
  AlertTriangle, 
  FileSpreadsheet, 
  Upload, 
  Download, 
  Euro, 
  Sliders, 
  Filter, 
  Info, 
  TrendingUp, 
  Wallet,
  Settings,
  HelpCircle
} from "lucide-react";
import { motion, AnimatePresence } from "framer-motion";

// --- HELPERS DE NORMALIZACIÓN Y PARSEO ---

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

// Extrae números de cadenas como "Talla 42" o "Nº 38"
function limpiarTallaString(talla) {
  const str = String(talla || "").trim().toUpperCase();
  const numMatch = str.match(/(?:TALLA|Nº|NUMERO|SIZE|NÚMERO)\s*([0-9]+(?:[\/\-][0-9]+)?)/i);
  if (numMatch) {
    return numMatch[1];
  }
  return str;
}

// Devuelve información estructurada de la talla para calcular distancias
function obtenerValorTalla(talla) {
  const tStr = limpiarTallaString(talla);
  if (!tStr) return { tipo: 'desconocido', valor: null, original: tStr };

  // 1. Número puro (ej. "42")
  const num = Number(tStr.replace(",", "."));
  if (!Number.isNaN(num)) {
    return { tipo: 'numero', valor: num, original: tStr };
  }

  // 2. Rango numérico (ej. "38/40", "40-42")
  const partesRangoNum = tStr.split(/[\/-]/).map(p => Number(p.trim().replace(",", ".")));
  if (partesRangoNum.length === 2 && !Number.isNaN(partesRangoNum[0]) && !Number.isNaN(partesRangoNum[1])) {
    return { tipo: 'numero', valor: (partesRangoNum[0] + partesRangoNum[1]) / 2, original: tStr };
  }

  // Mapa de pesos para tallas alfabéticas
  const mapLetras = {
    "XXS": 0, "XS": 1, "S": 2, "M": 3, "L": 4, "XL": 5, "XXL": 6, "2XL": 6, "XXXL": 7, "3XL": 7, "4XL": 8, "5XL": 9
  };

  // 3. Rango de letras (ej. "S/M", "M-L")
  const partesRangoText = tStr.split(/[\/-]/).map(p => p.trim());
  if (partesRangoText.length === 2) {
    const idx0 = mapLetras[partesRangoText[0]];
    const idx1 = mapLetras[partesRangoText[1]];
    if (idx0 !== undefined && idx1 !== undefined) {
      return { tipo: 'letra', valor: (idx0 + idx1) / 2, original: tStr };
    }
  }

  // 4. Letra única (ej. "M")
  if (mapLetras[tStr] !== undefined) {
    return { tipo: 'letra', valor: mapLetras[tStr], original: tStr };
  }

  return { tipo: 'desconocido', valor: null, original: tStr };
}

// Calcula la distancia absoluta entre dos tallas
function calcularDistanciaTallas(t1Info, t2Info) {
  if (t1Info.tipo === 'desconocido' || t2Info.tipo === 'desconocido') return Infinity;
  if (t1Info.tipo !== t2Info.tipo) return Infinity; // No se puede comparar 'M' con '42'
  return Math.abs(t1Info.valor - t2Info.valor);
}

// --- PARSER DE EXCEL ---

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

    if (talla !== "" && cantidad !== "") {
      resultado.push({ 
        articulo, 
        talla: limpiarValor(talla), 
        cantidad: esNumero(cantidad) ? obtenerNumero(cantidad) : cantidad 
      });
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

// --- ALGORITMO DE REPARTO ---

export function calcularRecomendaciones(stock, entradas, pedidos, configuraciones, toleranciaGlobal = 1) {
  // Resuelve el grupo de un artículo
  const obtenerGrupoKey = (articuloRaw) => {
    const conf = configuraciones[articuloRaw] || {};
    return conf.grupo || articuloRaw;
  };

  // Prepara los elementos agregando su origen logístico
  const stockConOrigen = stock.map(s => ({ ...s, origen: "Inventario" }));
  const entradasConOrigen = entradas.map(e => ({ ...e, origen: "Entrada" }));
  const stockDisponibleTotal = [...stockConOrigen, ...entradasConOrigen];

  // Agrupa por grupo
  const agruparPorGrupo = (lineas) => {
    const grupos = {};
    for (const linea of lineas) {
      const artRaw = normalizarTexto(linea.articulo);
      if (!artRaw) continue;
      const grpKey = obtenerGrupoKey(artRaw);
      if (!grupos[grpKey]) grupos[grpKey] = [];
      grupos[grpKey].push(linea);
    }
    return grupos;
  };

  const stockPorGrupo = agruparPorGrupo(stockDisponibleTotal);
  const pedidosPorGrupo = agruparPorGrupo(pedidos);

  const grupoKeys = Array.from(new Set([
    ...Object.keys(stockPorGrupo),
    ...Object.keys(pedidosPorGrupo)
  ]));

  const recomendaciones = [];
  const resumen = [];

  for (const grupoKey of grupoKeys) {
    const config = configuraciones[grupoKey] || {};
    const precio = config.precio !== undefined ? Number(config.precio) : 20;
    const tolerancia = config.tolerancia !== undefined ? Number(config.tolerancia) : Number(toleranciaGlobal);
    const direccion = config.direccion || "ambos";

    const stockItems = stockPorGrupo[grupoKey] || [];
    const pedidoItems = pedidosPorGrupo[grupoKey] || [];

    const articuloGrupoNombre = stockItems.find(s => s.origen === "Inventario")?.articulo || stockItems[0]?.articulo || pedidoItems[0]?.articulo || grupoKey;

    // Preparar lista de stock disponible mutable
    const stockDisponible = stockItems.map(item => ({
      articuloOriginal: item.articulo,
      tallaOriginal: item.talla,
      tallaInfo: obtenerValorTalla(item.talla),
      cantidad: Number(item.cantidad) || 0,
      origen: item.origen
    })).filter(s => s.cantidad > 0);

    const pedidosPendientes = pedidoItems.map(item => ({
      articuloOriginal: item.articulo,
      tallaOriginal: item.talla,
      tallaInfo: obtenerValorTalla(item.talla),
      cantidad: Number(item.cantidad) || 0
    })).filter(p => p.cantidad > 0);

    const asignaciones = [];

    // PASO 1: Asignación por talla exacta
    for (const pedido of pedidosPendientes) {
      // Intentar primero con talla exacta de Inventario
      let exactStock = stockDisponible.find(s => 
        s.origen === "Inventario" &&
        String(s.tallaOriginal).trim().toLowerCase() === String(pedido.tallaOriginal).trim().toLowerCase()
      );
      // Si no hay, intentar con talla exacta de Entrada
      if (!exactStock || exactStock.cantidad <= 0) {
        exactStock = stockDisponible.find(s => 
          s.origen === "Entrada" &&
          String(s.tallaOriginal).trim().toLowerCase() === String(pedido.tallaOriginal).trim().toLowerCase()
        );
      }

      if (exactStock && exactStock.cantidad > 0) {
        const usar = Math.min(pedido.cantidad, exactStock.cantidad);
        if (usar > 0) {
          asignaciones.push({
            articulo: pedido.articuloOriginal,
            articuloStock: exactStock.articuloOriginal,
            tallaPedido: pedido.tallaOriginal,
            tallaEntregada: exactStock.tallaOriginal,
            cantidad: usar,
            tipo: "Exacta",
            origenStock: exactStock.origen
          });
          exactStock.cantidad -= usar;
          pedido.cantidad -= usar;
        }
      }
    }

    // PASO 2: Asignación por sustitución según tolerancia y dirección
    if (tolerancia > 0) {
      for (const pedido of pedidosPendientes) {
        if (pedido.cantidad <= 0) continue;

        const candidatas = stockDisponible
          .filter(s => s.cantidad > 0)
          .map(s => ({ stockItem: s, distancia: calcularDistanciaTallas(pedido.tallaInfo, s.tallaInfo) }))
          .filter(cand => {
            if (cand.distancia > tolerancia) return false;
            if (direccion === "superior" && cand.stockItem.tallaInfo.valor < pedido.tallaInfo.valor) return false;
            if (direccion === "inferior" && cand.stockItem.tallaInfo.valor > pedido.tallaInfo.valor) return false;
            return true;
          })
          .sort((a, b) => {
            if (a.distancia !== b.distancia) return a.distancia - b.distancia;
            
            // Prioridad de origen: preferir Inventario sobre Entrada
            if (a.stockItem.origen === "Inventario" && b.stockItem.origen === "Entrada") return -1;
            if (b.stockItem.origen === "Inventario" && a.stockItem.origen === "Entrada") return 1;

            // Empate de distancia y origen: priorizar talla superior
            const valA = a.stockItem.tallaInfo.valor || 0;
            const valB = b.stockItem.tallaInfo.valor || 0;
            const valPed = pedido.tallaInfo.valor || 0;
            
            const diffA = valA - valPed;
            const diffB = valB - valPed;
            
            if (diffA > 0 && diffB < 0) return -1;
            if (diffB > 0 && diffA < 0) return 1;
            
            return valB - valA;
          });

        for (const cand of candidatas) {
          if (pedido.cantidad <= 0) break;
          const stockItem = cand.stockItem;
          const usar = Math.min(pedido.cantidad, stockItem.cantidad);
          if (usar > 0) {
            asignaciones.push({
              articulo: pedido.articuloOriginal,
              articuloStock: stockItem.articuloOriginal,
              tallaPedido: pedido.tallaOriginal,
              tallaEntregada: stockItem.tallaOriginal,
              cantidad: usar,
              tipo: "Sustitución",
              origenStock: stockItem.origen
            });
            stockItem.cantidad -= usar;
            pedido.cantidad -= usar;
          }
        }
      }
    }

    // PASO 3: Faltas
    for (const pedido of pedidosPendientes) {
      if (pedido.cantidad > 0) {
        asignaciones.push({
          articulo: pedido.articuloOriginal,
          articuloStock: null,
          tallaPedido: pedido.tallaOriginal,
          tallaEntregada: null,
          cantidad: pedido.cantidad,
          tipo: "Falta",
          origenStock: null
        });
      }
    }

    // Totales separados
    const totalStockInicial = stockItems.filter(s => s.origen === "Inventario").reduce((sum, s) => sum + (Number(s.cantidad) || 0), 0);
    const totalEntradas = stockItems.filter(s => s.origen === "Entrada").reduce((sum, s) => sum + (Number(s.cantidad) || 0), 0);
    const totalStockDisponible = totalStockInicial + totalEntradas;
    const totalPedido = pedidoItems.reduce((sum, p) => sum + (Number(p.cantidad) || 0), 0);
    const totalEntregado = asignaciones.filter(a => a.tipo !== "Falta").reduce((sum, a) => sum + a.cantidad, 0);
    const totalFalta = asignaciones.filter(a => a.tipo === "Falta").reduce((sum, a) => sum + a.cantidad, 0);
    const sobrante = stockDisponible.reduce((sum, s) => sum + s.cantidad, 0);

    const stockDisponibleRestante = {};
    stockDisponible.forEach(s => {
      if (s.cantidad > 0) {
        const key = `${s.tallaOriginal} (${s.origen === "Inventario" ? "Almacén" : "Llegada"})`;
        stockDisponibleRestante[key] = (stockDisponibleRestante[key] || 0) + s.cantidad;
      }
    });

    resumen.push({
      articulo: articuloGrupoNombre,
      articuloKey: grupoKey,
      totalStock: totalStockDisponible,
      totalStockInicial,
      totalEntradas,
      totalPedido,
      totalEntregado,
      totalFalta,
      sobrante,
      stockDisponible: stockDisponibleRestante
    });

    recomendaciones.push(...asignaciones);
  }

  return { recomendaciones, resumen };
}

// --- DATOS POR DEFECTO ---

// Los datos de ejemplo simulan el "Stock Teórico" del ERP:
// El stock inicial tiene valores positivos (tenemos) y negativos (necesidades).
// La app los separará de manera automática al cargarlos.
const stockEjemploERP = [
  { articulo: "Pantalones negros V", talla: "42", cantidad: 4 },
  { articulo: "Pantalones negros V", talla: "41", cantidad: 2 },
  { articulo: "Pantalones negros V", talla: "40", cantidad: -3 }, // Negativo = Necesidad
  { articulo: "Pantalones negros V", talla: "44", cantidad: -2 }, // Negativo = Necesidad
  { articulo: "Camisas Blancas", talla: "M", cantidad: 15 },
  { articulo: "Camisas Blancas", talla: "S", cantidad: -5 },      // Negativo = Necesidad
];
const entradasEjemploERP = [
  { articulo: "Pantalones negros V", talla: "42", cantidad: 1 },  // Llega 1 más de la 42
  { articulo: "Camisas Blancas", talla: "S", cantidad: 5 },      // Llega +5 de la S para cubrir el -5 teórico!
];

const configEjemplo = {
  "pantalones negros v": { precio: 28.50, tolerancia: 2, direccion: "ambos" },
  "camisas blancas": { precio: 18.20, tolerancia: 1, direccion: "ambos" },
};

// --- COMPONENTES AUXILIARES ---

function Card({ children, className = "", style = {} }) {
  return <div className={`card ${className}`} style={style}>{children}</div>;
}

function EditorTabla({ titulo, icono: Icono, filas, setFilas, todasLasPrendas = [], placeholderPrenda = "" }) {
  const actualizar = (idx, campo, valor) => {
    setFilas((prev) =>
      prev.map((fila, i) => {
        if (i === idx) {
          let val = valor;
          if (campo === "cantidad") {
            val = valor === "" ? "" : Number(valor);
          }
          return { ...fila, [campo]: val };
        }
        return fila;
      })
    );
  };

  const agregar = () => {
    setFilas((prev) => [...prev, { articulo: placeholderPrenda || "", talla: "", cantidad: 1 }]);
  };

  const borrar = (idx) => {
    setFilas((prev) => prev.filter((_, i) => i !== idx));
  };

  const datalistId = `datalist-${titulo.replace(/[^a-zA-Z0-9]/g, "-").toLowerCase()}`;

  return (
    <Card>
      <div className="card-header">
        <h2>
          <Icono size={18} style={{ color: "var(--primary)" }} /> {titulo}
        </h2>
        <button className="btn btn-outline" style={{ padding: "6px 12px", fontSize: "13.5px" }} onClick={agregar}>
          <Plus size={14} /> Añadir fila
        </button>
      </div>
      <div className="card-body">
        <datalist id={datalistId}>
          {todasLasPrendas.map((prenda) => (
            <option key={prenda} value={prenda} />
          ))}
        </datalist>

        {filas.length === 0 ? (
          <div className="empty-state" style={{ padding: "20px 0", textAlign: "center", color: "var(--text-muted)" }}>
            <p style={{ fontSize: "13.5px" }}>No hay filas. Haz clic en "Añadir fila" para empezar a editar.</p>
          </div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: "12px" }}>
            <div className="editor-grid-header">
              <div>Artículo</div>
              <div>Talla</div>
              <div>Cantidad</div>
              <div></div>
            </div>
            <div className="editor-grid-rows">
              {filas.map((fila, idx) => (
                <div key={idx} className="editor-grid-row">
                  <input
                    type="text"
                    className="input-control"
                    value={fila.articulo || ""}
                    placeholder="Nombre de la prenda..."
                    list={datalistId}
                    onChange={(e) => actualizar(idx, "articulo", e.target.value)}
                  />
                  <input
                    type="text"
                    className="input-control"
                    value={fila.talla || ""}
                    placeholder="Ej. 42, M..."
                    onChange={(e) => actualizar(idx, "talla", e.target.value)}
                  />
                  <input
                    type="number"
                    className="input-control"
                    value={fila.cantidad}
                    placeholder="Cant."
                    onChange={(e) => actualizar(idx, "cantidad", e.target.value)}
                  />
                  <button
                    className="btn btn-danger"
                    style={{
                      padding: "8px",
                      backgroundColor: "transparent",
                      color: "var(--danger)",
                      border: "none",
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center"
                    }}
                    onClick={() => borrar(idx)}
                    title="Eliminar fila"
                  >
                    <Trash2 size={16} />
                  </button>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    </Card>
  );
}


// --- APP COMPONENT ---

export default function App() {
  // Inicialización de estados desde LocalStorage
  const [stock, setStock] = useState([]);
  const [entradas, setEntradas] = useState([]);
  const [pedidos, setPedidos] = useState([]);
  const [configuraciones, setConfiguraciones] = useState({});
  const [toleranciaGlobal, setToleranciaGlobal] = useState(2);
  const [activeTab, setActiveTab] = useState("dashboard");

  // Filtros de resultados
  const [filtroArticulo, setFiltroArticulo] = useState("");
  const [filtroResultado, setFiltroResultado] = useState("");
  const [busqueda, setBusqueda] = useState("");

  // Carga inicial (desde ejemplo o localStorage)
  useEffect(() => {
    const savedStock = localStorage.getItem("agy_vestuario_stock_v2");
    const savedEntradas = localStorage.getItem("agy_vestuario_entradas_v2");
    const savedPedidos = localStorage.getItem("agy_vestuario_pedidos_v2");
    const savedConfig = localStorage.getItem("agy_vestuario_config_v2");
    const savedTolerancia = localStorage.getItem("agy_vestuario_tolerancia_v2");

    if (savedStock && savedEntradas && savedPedidos) {
      setStock(JSON.parse(savedStock));
      setEntradas(JSON.parse(savedEntradas));
      setPedidos(JSON.parse(savedPedidos));
      if (savedConfig) setConfiguraciones(JSON.parse(savedConfig));
      if (savedTolerancia) setToleranciaGlobal(Number(savedTolerancia));
    } else {
      // Si está vacío, cargar por defecto estructurado según el ERP (negativos automáticos)
      procesarYEstablecerDatosERP(stockEjemploERP, entradasEjemploERP, configEjemplo);
    }
  }, []);

  // Guardar en LocalStorage cada vez que cambien los datos
  useEffect(() => {
    if (stock.length > 0 || pedidos.length > 0) {
      localStorage.setItem("agy_vestuario_stock_v2", JSON.stringify(stock));
      localStorage.setItem("agy_vestuario_pedidos_v2", JSON.stringify(pedidos));
      localStorage.setItem("agy_vestuario_entradas_v2", JSON.stringify(entradas));
      localStorage.setItem("agy_vestuario_config_v2", JSON.stringify(configuraciones));
      localStorage.setItem("agy_vestuario_tolerancia_v2", String(toleranciaGlobal));
    }
  }, [stock, entradas, pedidos, configuraciones, toleranciaGlobal]);

  // Lista de prendas detectadas de forma dinámica (todas)
  const todasLasPrendas = useMemo(() => {
    const items = new Set();
    stock.forEach(s => s.articulo && items.add(s.articulo));
    entradas.forEach(e => e.articulo && items.add(e.articulo));
    pedidos.forEach(p => p.articulo && items.add(p.articulo));
    return Array.from(items).sort();
  }, [stock, entradas, pedidos]);

  // Asegura que todas las prendas tengan una configuración inicial en el estado
  useEffect(() => {
    if (todasLasPrendas.length === 0) return;
    let cambio = false;
    const nuevasConfig = { ...configuraciones };
    todasLasPrendas.forEach(prenda => {
      const key = normalizarTexto(prenda);
      if (!nuevasConfig[key]) {
        nuevasConfig[key] = { precio: 20, tolerancia: 2, direccion: "ambos" };
        cambio = true;
      }
    });
    if (cambio) {
      setConfiguraciones(nuevasConfig);
    }
  }, [todasLasPrendas, configuraciones]);

  // --- FUNCIÓN CLAVE: SEPARADOR AUTOMÁTICO DE VALORES ERP (POSITIVOS Y NEGATIVOS) ---
  const procesarYEstablecerDatosERP = (rawStock, rawEntradas, customConfig = null) => {
    const stockPositivos = [];
    const pedidosNegativos = [];

    // Separar el stock teórico: los positivos van a inventario, los negativos a pedidos
    rawStock.forEach(item => {
      const cant = Number(item.cantidad);
      if (!isNaN(cant)) {
        if (cant < 0) {
          // El negativo se extrae como una necesidad/pedido (demanda absoluta)
          pedidosNegativos.push({ 
            articulo: item.articulo, 
            talla: item.talla, 
            cantidad: Math.abs(cant) 
          });
        } else if (cant > 0) {
          stockPositivos.push({ 
            articulo: item.articulo, 
            talla: item.talla, 
            cantidad: cant 
          });
        }
      } else {
        stockPositivos.push(item);
      }
    });

    // Las entradas se cargan limpiando cualquier posible negativo accidental
    const entradasLimpias = rawEntradas.map(e => ({
      articulo: e.articulo,
      talla: e.talla,
      cantidad: Math.max(0, Number(e.cantidad) || 0)
    })).filter(e => e.cantidad > 0);

    setStock(stockPositivos);
    setPedidos(pedidosNegativos);
    setEntradas(entradasLimpias);
    
    if (customConfig) {
      setConfiguraciones(customConfig);
    }
    setToleranciaGlobal(2);
  };

  // Manejar importación desde Excel de Stock Teórico (separa + y -)
  const importarExcelTeorico = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      const { resultado } = await leerExcel(file);
      // Procesamos el Excel teórico separando positivos y negativos
      const stockPos = [];
      const pedidosNeg = [];
      
      resultado.forEach(item => {
        const cant = Number(item.cantidad);
        if (!isNaN(cant)) {
          if (cant < 0) {
            pedidosNeg.push({ articulo: item.articulo, talla: item.talla, cantidad: Math.abs(cant) });
          } else if (cant > 0) {
            stockPos.push({ articulo: item.articulo, talla: item.talla, cantidad: cant });
          }
        } else {
          stockPos.push(item);
        }
      });

      setStock(stockPos);
      setPedidos(pedidosNeg);
      alert(`Excel de Stock Teórico cargado con éxito. Se han detectado ${stockPos.length} líneas de stock y ${pedidosNeg.length} líneas de necesidades (cantidades negativas).`);
    } catch (error) {
      console.error(error);
      alert("Error al procesar el Excel. Verifica las columnas.");
    } finally {
      event.target.value = "";
    }
  };

  // Manejar importación desde Excel de Stock Práctico / Entrada (lo que llega)
  const importarExcelEntradas = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      const { resultado } = await leerExcel(file);
      const entradasLimpias = resultado.map(item => ({
        articulo: item.articulo,
        talla: item.talla,
        cantidad: Math.max(0, Number(item.cantidad) || 0)
      })).filter(e => e.cantidad > 0);

      setEntradas(entradasLimpias);
      alert(`Excel de Entrada/Recibo cargado con éxito. ${entradasLimpias.length} líneas de llegada importadas.`);
    } catch (error) {
      console.error(error);
      alert("Error al procesar el Excel. Verifica las columnas.");
    } finally {
      event.target.value = "";
    }
  };

  // --- DETECTOR DE PRENDAS SIMILARES NO AGRUPADAS ---
  const sugerenciasAgrupacion = useMemo(() => {
    const sugerencias = [];
    for (let i = 0; i < todasLasPrendas.length; i++) {
      for (let j = i + 1; j < todasLasPrendas.length; j++) {
        const p1 = todasLasPrendas[i];
        const p2 = todasLasPrendas[j];
        const k1 = normalizarTexto(p1);
        const k2 = normalizarTexto(p2);

        const g1 = configuraciones[k1]?.grupo || k1;
        const g2 = configuraciones[k2]?.grupo || k2;
        if (g1 === g2) continue;

        const w1 = k1.split(/\s+/).filter(w => w.length > 3);
        const w2 = k2.split(/\s+/).filter(w => w.length > 3);

        const compartenPalabra = w1.some(pal => w2.includes(pal));

        if (compartenPalabra) {
          sugerencias.push({
            originalA: p1,
            keyA: k1,
            originalB: p2,
            keyB: k2
          });
        }
      }
    }
    return sugerencias;
  }, [todasLasPrendas, configuraciones]);

  const agruparPrendas = (keyA, keyB) => {
    setConfiguraciones(prev => ({
      ...prev,
      [keyB]: {
        ...(prev[keyB] || { precio: 20, tolerancia: 2, direccion: "ambos" }),
        grupo: keyA
      }
    }));
  };

  // Cálculo de recomendaciones y resúmenes económicos en tiempo real
  const { recomendaciones, resumen } = useMemo(() => {
    return calcularRecomendaciones(stock, entradas, pedidos, configuraciones, toleranciaGlobal);
  }, [stock, entradas, pedidos, configuraciones, toleranciaGlobal]);

  const sustituciones = useMemo(() => {
    return recomendaciones.filter(r => r.tipo === "Sustitución");
  }, [recomendaciones]);

  const faltantes = useMemo(() => {
    return recomendaciones.filter(r => r.tipo === "Falta");
  }, [recomendaciones]);

  // --- CÁLCULOS GENERALES DEL PANEL (KPIs) ---
  const kpis = useMemo(() => {
    let valorStockInicialTotal = 0;
    let valorEntradasTotal = 0;
    let valorPedidosTotal = 0;
    let valorEntregadoTotal = 0;
    let valorFaltanteTotal = 0;
    let valorSobranteTotal = 0;

    let unidadesPedidas = 0;
    let unidadesEntregadas = 0;

    resumen.forEach(r => {
      const conf = configuraciones[r.articuloKey] || {};
      const precio = conf.precio !== undefined ? Number(conf.precio) : 20;

      valorStockInicialTotal += r.totalStockInicial * precio;
      valorEntradasTotal += r.totalEntradas * precio;
      valorPedidosTotal += r.totalPedido * precio;
      valorEntregadoTotal += r.totalEntregado * precio;
      valorFaltanteTotal += r.totalFalta * precio;
      valorSobranteTotal += r.sobrante * precio;

      unidadesPedidas += r.totalPedido;
      unidadesEntregadas += r.totalEntregado;
    });

    const eficienciaValor = valorPedidosTotal > 0 ? (valorEntregadoTotal / valorPedidosTotal) * 100 : 0;
    const eficienciaUnidades = unidadesPedidas > 0 ? (unidadesEntregadas / unidadesPedidas) * 100 : 0;

    return {
      valorStockInicialTotal,
      valorEntradasTotal,
      valorStockDisponibleTotal: valorStockInicialTotal + valorEntradasTotal,
      valorPedidosTotal,
      valorEntregadoTotal,
      valorFaltanteTotal,
      valorSobranteTotal,
      eficienciaValor,
      eficienciaUnidades
    };
  }, [resumen, configuraciones]);

  // --- AVISOS Y VALIDACIONES ---
  const avisos = useMemo(() => {
    const lista = [];

    // 1. Artículos en pedidos que no existen en stock
    const articulosStockNormalizados = new Set([
      ...stock.map(s => {
        const k = normalizarTexto(s.articulo);
        return configuraciones[k]?.grupo || k;
      }),
      ...entradas.map(e => {
        const k = normalizarTexto(e.articulo);
        return configuraciones[k]?.grupo || k;
      })
    ]);

    const articulosPedidosUnicos = Array.from(new Set(pedidos.map(p => p.articulo)));
    
    articulosPedidosUnicos.forEach(art => {
      if (art) {
        const k = normalizarTexto(art);
        const grupo = configuraciones[k]?.grupo || k;
        if (!articulosStockNormalizados.has(grupo)) {
          lista.push({
            tipo: "warning",
            mensaje: `El artículo "${art}" solicitado en necesidades (negativos) no tiene existencias equivalentes en stock ni en llegada.`
          });
        }
      }
    });

    // 2. Artículos sin precio asignado o con precio igual a cero
    todasLasPrendas.forEach(prenda => {
      const conf = configuraciones[normalizarTexto(prenda)];
      if (conf && (Number(conf.precio) === 0 || isNaN(conf.precio))) {
        lista.push({
          tipo: "info",
          mensaje: `El artículo "${prenda}" tiene un precio unitario de 0.00 €. Considera asignarle un precio real.`
        });
      }
    });

    return lista;
  }, [stock, entradas, pedidos, configuraciones, todasLasPrendas]);

  // --- FILTRADO DE RECOMENDACIONES ---
  const recomendacionesFiltradas = useMemo(() => {
    return recomendaciones.filter(rec => {
      const coincideArticulo = filtroArticulo ? normalizarTexto(rec.articulo) === normalizarTexto(filtroArticulo) : true;
      const coincideResultado = filtroResultado ? rec.tipo === filtroResultado : true;
      
      const searchNorm = normalizarTexto(busqueda);
      const coincideBusqueda = searchNorm
        ? normalizarTexto(rec.articulo).includes(searchNorm) || 
          String(rec.tallaPedido).toLowerCase().includes(searchNorm) || 
          String(rec.tallaEntregada || "").toLowerCase().includes(searchNorm) ||
          String(rec.articuloStock || "").toLowerCase().includes(searchNorm)
        : true;

      return coincideArticulo && coincideResultado && coincideBusqueda;
    });
  }, [recomendaciones, filtroArticulo, filtroResultado, busqueda]);

  // --- EXPORTAR A EXCEL ---
  const descargarExcel = () => {
    const wb = XLSX.utils.book_new();

    // Hoja 1: Resumen Económico
    const datosResumen = resumen.map(r => {
      const conf = configuraciones[r.articuloKey] || {};
      const precio = conf.precio !== undefined ? conf.precio : 20;
      return {
        "Artículo / Prenda": r.articulo,
        "Precio Unitario (€)": precio,
        "Stock Inicial (Uds)": r.totalStockInicial,
        "Valor Stock Inicial (€)": r.totalStockInicial * precio,
        "Entradas / Lo que llega (Uds)": r.totalEntradas,
        "Valor Entradas (€)": r.totalEntradas * precio,
        "Pedido / Lo que sale (Uds)": r.totalPedido,
        "Valor Pedido (€)": r.totalPedido * precio,
        "Asignado/Entregado (Uds)": r.totalEntregado,
        "Valor Asignado (€)": r.totalEntregado * precio,
        "Faltante Neto (Uds)": r.totalFalta,
        "Presupuesto de Compra (€)": r.totalFalta * precio,
        "Sobrante Final (Uds)": r.sobrante,
        "Valor Sobrante (€)": r.sobrante * precio,
        "Balance Neto Final (Uds)": r.sobrante - r.totalFalta,
        "Balance Neto Final (€)": (r.sobrante - r.totalFalta) * precio
      };
    });

    // Agregar fila de totales
    datosResumen.push({
      "Artículo / Prenda": "TOTAL GENERAL",
      "Precio Unitario (€)": "",
      "Stock Inicial (Uds)": resumen.reduce((sum, r) => sum + r.totalStockInicial, 0),
      "Valor Stock Inicial (€)": kpis.valorStockInicialTotal,
      "Entradas / Lo que llega (Uds)": resumen.reduce((sum, r) => sum + r.totalEntradas, 0),
      "Valor Entradas (€)": kpis.valorEntradasTotal,
      "Pedido / Lo que sale (Uds)": resumen.reduce((sum, r) => sum + r.totalPedido, 0),
      "Valor Pedido (€)": kpis.valorPedidosTotal,
      "Asignado/Entregado (Uds)": resumen.reduce((sum, r) => sum + r.totalEntregado, 0),
      "Valor Asignado (€)": kpis.valorEntregadoTotal,
      "Faltante Neto (Uds)": resumen.reduce((sum, r) => sum + r.totalFalta, 0),
      "Presupuesto de Compra (€)": kpis.valorFaltanteTotal,
      "Sobrante Final (Uds)": resumen.reduce((sum, r) => sum + r.sobrante, 0),
      "Valor Sobrante (€)": kpis.valorSobranteTotal,
      "Balance Neto Final (Uds)": resumen.reduce((sum, r) => sum + (r.sobrante - r.totalFalta), 0),
      "Balance Neto Final (€)": kpis.valorSobranteTotal - kpis.valorFaltanteTotal
    });

    const wsResumen = XLSX.utils.json_to_sheet(datosResumen);
    XLSX.utils.book_append_sheet(wb, wsResumen, "Resumen General");

    // Hoja 2: Recomendaciones Detalladas
    const datosRec = recomendaciones.map(rec => {
      const precio = configuraciones[normalizarTexto(rec.articulo)]?.precio ?? 20;
      return {
        "Artículo Solicitado": rec.articulo,
        "Talla Pedida": rec.tallaPedido,
        "Talla Entregada": rec.tallaEntregada || "—",
        "Artículo Stock Usado": rec.articuloStock || "—",
        "Origen Logístico": rec.origenStock === "Inventario" ? "Almacén (Existente)" : rec.origenStock === "Entrada" ? "Llegada (Nuevo)" : "—",
        "Cantidad (Uds)": rec.cantidad,
        "Precio Unitario (€)": precio,
        "Valor Total (€)": rec.cantidad * precio,
        "Resultado": rec.tipo,
        "Explicación": rec.tipo === "Exacta" ? "Entrega exacta" : rec.tipo === "Sustitución" ? `Sustitución por talla ${rec.tallaEntregada}` : "Falta material (Comprar)"
      };
    });
    const wsRec = XLSX.utils.json_to_sheet(datosRec);
    XLSX.utils.book_append_sheet(wb, wsRec, "Recomendaciones");

    // Hoja 3: Faltantes (Presupuesto de Compra)
    const datosFaltantes = recomendaciones.filter(rec => rec.tipo === "Falta").map(rec => {
      const precio = configuraciones[normalizarTexto(rec.articulo)]?.precio ?? 20;
      return {
        "Artículo Solicitado": rec.articulo,
        "Talla Pedida": rec.tallaPedido,
        "Cantidad Faltante (Uds)": rec.cantidad,
        "Precio Unitario (€)": precio,
        "Presupuesto Requerido (€)": rec.cantidad * precio
      };
    });
    const wsFaltantes = XLSX.utils.json_to_sheet(datosFaltantes);
    XLSX.utils.book_append_sheet(wb, wsFaltantes, "Compras Faltantes");

    // Hoja 4: Sobrantes
    const datosSobrantes = [];
    resumen.forEach(r => {
      const precio = configuraciones[r.articuloKey]?.precio ?? 20;
      Object.entries(r.stockDisponible).forEach(([tallaInfo, cant]) => {
        if (cant > 0) {
          datosSobrantes.push({
            "Artículo / Prenda": r.articulo,
            "Talla y Origen": tallaInfo,
            "Cantidad Sobrante (Uds)": cant,
            "Precio Unitario (€)": precio,
            "Valor Sobrante (€)": cant * precio
          });
        }
      });
    });
    const wsSobrantes = XLSX.utils.json_to_sheet(datosSobrantes);
    XLSX.utils.book_append_sheet(wb, wsSobrantes, "Stock Final Sobrante");

    XLSX.writeFile(wb, "informe_control_economico_recepciones.xlsx");
  };

  const resetearEjemploERP = () => {
    procesarYEstablecerDatosERP(stockEjemploERP, entradasEjemploERP, configEjemplo);
  };

  const vaciarTodo = () => {
    setStock([]);
    setEntradas([]);
    setPedidos([]);
    setConfiguraciones({});
    setToleranciaGlobal(2);
    localStorage.removeItem("agy_vestuario_stock_v2");
    localStorage.removeItem("agy_vestuario_pedidos_v2");
    localStorage.removeItem("agy_vestuario_entradas_v2");
    localStorage.removeItem("agy_vestuario_config_v2");
    localStorage.removeItem("agy_vestuario_tolerancia_v2");
  };

  return (
    <div className="app">
      {/* HEADER DE LA APLICACIÓN */}
      <header className="main-header">
        <div className="header-container">
          <div className="brand-section">
            <div className="brand-icon">
              <Shirt size={26} />
            </div>
            <div className="brand-text">
              <h1>Gestión y Control de Vestuario</h1>
              <p>Reconciliación de inventario: Stock Teórico (ERP) con entradas (lo que llega) y salidas (negativos)</p>
            </div>
          </div>
          <div className="gap-10">
            <button className="btn btn-outline" style={{ color: "#fff", borderColor: "#334155" }} onClick={resetearEjemploERP}>
              <RefreshCw size={14} /> Cargar Ejemplo ERP (+ / -)
            </button>
            <button className="btn btn-danger" style={{ backgroundColor: "#311", color: "#f87171" }} onClick={vaciarTodo}>
              Vaciar Todo
            </button>
          </div>
        </div>
      </header>

      {/* NAVEGACIÓN EN PESTAÑAS */}
      <nav className="tabs-navigation">
        <div className="tabs-container">
          <button 
            className={`tab-btn ${activeTab === "dashboard" ? "active" : ""}`}
            onClick={() => setActiveTab("dashboard")}
          >
            <TrendingUp size={16} /> Panel y Carga
          </button>
          <button 
            className={`tab-btn ${activeTab === "precios" ? "active" : ""}`}
            onClick={() => setActiveTab("precios")}
          >
            <Euro size={16} /> Tarifas y Reglas
          </button>
          <button 
            className={`tab-btn ${activeTab === "editores" ? "active" : ""}`}
            onClick={() => setActiveTab("editores")}
          >
            <Sliders size={16} /> Editar Datos ({stock.length + entradas.length + pedidos.length})
          </button>
          <button 
            className={`tab-btn ${activeTab === "recomendaciones" ? "active" : ""}`}
            onClick={() => setActiveTab("recomendaciones")}
          >
            <CheckCircle2 size={16} /> Recomendaciones y Reporte ({recomendaciones.length})
          </button>
        </div>
      </nav>

      {/* CONTENIDO PRINCIPAL */}
      <main className="container">
        <AnimatePresence mode="wait">
          {activeTab === "dashboard" && (
            <motion.div 
              key="dashboard"
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -15 }}
              transition={{ duration: 0.2 }}
              style={{ display: "flex", flexDirection: "column", gap: "24px" }}
            >
              {/* INDICADORES FINANCIEROS CLAVE (KPIs) */}
              <div className="kpi-grid">
                <div className="kpi-card deficit">
                  <div className="kpi-icon-box">
                    <Wallet size={24} />
                  </div>
                  <div className="kpi-info">
                    <span className="kpi-label">Presupuesto Compra (Déficit)</span>
                    <span className="kpi-value">{kpis.valorFaltanteTotal.toLocaleString('es-ES', { style: 'currency', currency: 'EUR' })}</span>
                    <span className="kpi-sub">Valor de las prendas negativas sin cubrir</span>
                  </div>
                </div>

                <div className="kpi-card stock-val">
                  <div className="kpi-icon-box">
                    <FileSpreadsheet size={24} />
                  </div>
                  <div className="kpi-info">
                    <span className="kpi-label">Disponible (Tengo + Llega)</span>
                    <span className="kpi-value">{kpis.valorStockDisponibleTotal.toLocaleString('es-ES', { style: 'currency', currency: 'EUR' })}</span>
                    <span className="kpi-sub">Almacén ({kpis.valorStockInicialTotal.toFixed(0)}€) + Entradas ({kpis.valorEntradasTotal.toFixed(0)}€)</span>
                  </div>
                </div>

                <div className="kpi-card saved">
                  <div className="kpi-icon-box">
                    <Euro size={24} />
                  </div>
                  <div className="kpi-info">
                    <span className="kpi-label">Ahorro Logístico (Entregado)</span>
                    <span className="kpi-value">{kpis.valorEntregadoTotal.toLocaleString('es-ES', { style: 'currency', currency: 'EUR' })}</span>
                    <span className="kpi-sub">Compensado con stock disponible</span>
                  </div>
                </div>

                <div className="kpi-card coverage">
                  <div className="kpi-icon-box">
                    <TrendingUp size={24} />
                  </div>
                  <div className="kpi-info">
                    <span className="kpi-label">Cobertura de Negativos</span>
                    <span className="kpi-value">{kpis.eficienciaValor.toFixed(1)}%</span>
                    <span className="kpi-sub">Eficiencia en € ({kpis.eficienciaUnidades.toFixed(0)}% en uds)</span>
                  </div>
                </div>
              </div>

              {/* AUTOMÁTICO: SUGERENCIA DE AGRUPACIÓN (BALANCE 0) */}
              {sugerenciasAgrupacion.length > 0 && (
                <Card style={{ backgroundColor: "#eff6ff", borderColor: "#bfdbfe" }}>
                  <div className="card-body">
                    <h3 style={{ color: "#1e3a8a", fontSize: "16px", fontWeight: "700", display: "flex", alignItems: "center", gap: "8px", marginBottom: "8px" }}>
                      <Info size={18} /> Sugerencias de Agrupación (Buscar Balance 0)
                    </h3>
                    <p style={{ fontSize: "13.5px", color: "#1e40af", marginBottom: "14px" }}>
                      Hemos detectado artículos con nombres similares (ej. <b>Pantalones negros V</b> y <b>Pantalones V</b>). 
                      Agrúpalos para compensar las tallas en el cálculo y buscar un balance equilibrado (cero) utilizando tanto el stock como lo que llega:
                    </p>
                    <div style={{ display: "flex", flexDirection: "column", gap: "8px" }}>
                      {sugerenciasAgrupacion.map((sug, idx) => (
                        <div key={idx} className="flex-between" style={{ backgroundColor: "#ffffff", padding: "10px 16px", borderRadius: "var(--radius-md)", border: "1px solid #dbeafe" }}>
                          <span style={{ fontSize: "13px", color: "var(--text-main)" }}>
                            ¿Agrupar pedidos de <b>{sug.originalB}</b> bajo la prenda de stock <b>{sug.originalA}</b>?
                          </span>
                          <button 
                            className="btn btn-primary" 
                            style={{ padding: "6px 14px", fontSize: "12px" }} 
                            onClick={() => agruparPrendas(sug.keyA, sug.keyB)}
                          >
                            Agrupar ahora y balancear
                          </button>
                        </div>
                      ))}
                    </div>
                  </div>
                </Card>
              )}

              {/* COBERTURA VISUAL */}
              <Card>
                <div className="card-body">
                  <div className="flex-between">
                    <span className="bold" style={{ fontSize: "15px" }}>Eficiencia de Asignación Presupuestaria (Cobertura de Negativos)</span>
                    <span className="bold" style={{ color: "var(--success-dark)" }}>
                      {kpis.valorEntregadoTotal.toLocaleString('es-ES', { style: 'currency', currency: 'EUR' })} asignados de {kpis.valorPedidosTotal.toLocaleString('es-ES', { style: 'currency', currency: 'EUR' })} en negativo (necesidades)
                    </span>
                  </div>
                  <div className="coverage-progress-bar">
                    <div className="coverage-progress-fill" style={{ width: `${Math.min(kpis.eficienciaValor, 100)}%` }} />
                  </div>
                  <p className="helper-text" style={{ marginTop: "8px" }}>
                    Representa el porcentaje de necesidades (negativos teóricos) que han sido cubiertos y compensados gracias a existencias positivas de almacén y a nuevas entregas, aplicando las tolerancias de tallas.
                  </p>
                </div>
              </Card>

              {/* SUSTITUCIONES Y DÉFICIT DE MATERIAL DETECTADO */}
              {(stock.length > 0 || pedidos.length > 0) && (
                <div className="grid-two">
                  {/* Tarjeta de Sustituciones */}
                  <Card style={{ borderLeft: "4px solid var(--warning)" }}>
                    <div className="card-header" style={{ padding: "16px 20px" }}>
                      <h2 style={{ fontSize: "15px", fontWeight: "700", color: "var(--warning-dark)", display: "flex", alignItems: "center", gap: "8px", margin: 0 }}>
                        <RefreshCw size={16} /> Sustituciones de Talla Recomendadas ({sustituciones.length})
                      </h2>
                      <span className="badge badge-warning" style={{ fontSize: "11px", fontWeight: "700" }}>
                        Compensación
                      </span>
                    </div>
                    <div className="card-body" style={{ padding: "16px 20px", maxHeight: "300px", overflowY: "auto" }}>
                      {sustituciones.length === 0 ? (
                        <div style={{ textAlign: "center", padding: "20px 0", color: "var(--text-muted)", fontSize: "13.5px" }}>
                          <CheckCircle2 size={24} style={{ color: "var(--success)", marginBottom: "8px", margin: "0 auto" }} />
                          <p style={{ marginTop: "8px", margin: 0 }}>No se requieren sustituciones de talla. Todas las entregas son exactas.</p>
                        </div>
                      ) : (
                        <div style={{ display: "flex", flexDirection: "column", gap: "8px" }}>
                          {sustituciones.map((sug, idx) => (
                            <div key={idx} style={{ 
                              display: "flex", 
                              justifyContent: "space-between", 
                              alignItems: "center",
                              backgroundColor: "#fffbeb", 
                              padding: "10px 12px", 
                              borderRadius: "var(--radius-sm)",
                              border: "1px solid #fde68a",
                              fontSize: "13px"
                            }}>
                              <div style={{ display: "flex", flexDirection: "column", gap: "2px" }}>
                                <span style={{ fontWeight: "700", textTransform: "capitalize", color: "var(--text-main)", textAlign: "left" }}>
                                  {sug.articulo}
                                </span>
                                <span style={{ color: "var(--text-muted)", fontSize: "12px", display: "flex", alignItems: "center", gap: "4px", flexWrap: "wrap" }}>
                                  Pedida: <strong style={{ color: "var(--text-main)" }}>{sug.tallaPedido}</strong> 
                                  ➜ Entregar: <strong style={{ color: "var(--primary)" }}>{sug.tallaEntregada}</strong>
                                  <span style={{ fontSize: "11px", color: "var(--text-light)" }}>
                                    ({sug.origenStock === "Inventario" ? "Almacén" : "Llegada"})
                                  </span>
                                </span>
                              </div>
                              <span style={{ 
                                fontWeight: "800", 
                                color: "var(--warning-dark)", 
                                backgroundColor: "#fef3c7", 
                                padding: "4px 8px", 
                                borderRadius: "6px",
                                fontSize: "12.5px"
                              }}>
                                {sug.cantidad} ud{sug.cantidad > 1 ? "s" : ""}
                              </span>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  </Card>

                  {/* Tarjeta de Déficit / Compras requeridas */}
                  <Card style={{ borderLeft: "4px solid var(--danger)" }}>
                    <div className="card-header" style={{ padding: "16px 20px" }}>
                      <h2 style={{ fontSize: "15px", fontWeight: "700", color: "var(--danger-dark)", display: "flex", alignItems: "center", gap: "8px", margin: 0 }}>
                        <AlertTriangle size={16} /> Déficit Neto de Material / Compras Requeridas ({faltantes.length})
                      </h2>
                      <span className="badge badge-danger" style={{ fontSize: "11px", fontWeight: "700" }}>
                        Faltantes
                      </span>
                    </div>
                    <div className="card-body" style={{ padding: "16px 20px", maxHeight: "300px", overflowY: "auto" }}>
                      {faltantes.length === 0 ? (
                        <div style={{ textAlign: "center", padding: "20px 0", color: "var(--text-muted)", fontSize: "13.5px" }}>
                          <CheckCircle2 size={24} style={{ color: "var(--success)", marginBottom: "8px", margin: "0 auto" }} />
                          <p style={{ marginTop: "8px", margin: 0 }}>¡Todo cubierto! No se necesita comprar ninguna prenda adicional.</p>
                        </div>
                      ) : (
                        <div style={{ display: "flex", flexDirection: "column", gap: "8px" }}>
                          {faltantes.map((fal, idx) => (
                            <div key={idx} style={{ 
                              display: "flex", 
                              justifyContent: "space-between", 
                              alignItems: "center",
                              backgroundColor: "#fef2f2", 
                              padding: "10px 12px", 
                              borderRadius: "var(--radius-sm)",
                              border: "1px solid #fecaca",
                              fontSize: "13px"
                            }}>
                              <div style={{ display: "flex", flexDirection: "column", gap: "2px" }}>
                                <span style={{ fontWeight: "700", textTransform: "capitalize", color: "var(--text-main)", textAlign: "left" }}>
                                  {fal.articulo}
                                </span>
                                <span style={{ color: "var(--text-muted)", fontSize: "12px", textAlign: "left" }}>
                                  Talla pedida: <strong style={{ color: "var(--danger-dark)" }}>{fal.tallaPedido}</strong> (Sin equivalencia en stock)
                                </span>
                              </div>
                              <span style={{ 
                                fontWeight: "800", 
                                color: "var(--danger-dark)", 
                                backgroundColor: "#fee2e2", 
                                padding: "4px 8px", 
                                borderRadius: "6px",
                                fontSize: "12.5px"
                              }}>
                                {fal.cantidad} ud{fal.cantidad > 1 ? "s" : ""}
                              </span>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  </Card>
                </div>
              )}

              {/* AREA DE CARGA DE ARCHIVOS (DOS RECIPIENTES) */}
              <h2 style={{ fontFamily: "var(--font-display)", fontSize: "18px", fontWeight: "700", marginTop: "10px" }}>Carga de Archivos Excel</h2>
              <div className="grid-two">
                <Card>
                  <div className="card-header">
                    <h2><Shirt size={18} style={{ color: "var(--success)" }} /> 1. Excel de Stock Teórico (Sistema ERP)</h2>
                  </div>
                  <div className="card-body">
                    <div className="upload-dropzone" onClick={() => document.getElementById('file-input-teorico').click()}>
                      <div className="upload-icon">
                        <FileSpreadsheet size={24} />
                      </div>
                      <div>
                        <div className="upload-title">Sube el Stock Teórico</div>
                        <div className="upload-desc">Contiene tanto los sobrantes (valores positivos) como los pedidos/salidas pendientes (valores negativos). La app los dividirá automáticamente.</div>
                      </div>
                      <input 
                        id="file-input-teorico" 
                        type="file" 
                        accept=".xlsx,.xls" 
                        onChange={importarExcelTeorico} 
                      />
                    </div>
                  </div>
                </Card>

                <Card>
                  <div className="card-header">
                    <h2><Download size={18} style={{ color: "var(--primary)" }} /> 2. Excel de Stock Práctico / Entrada (Lo que llega)</h2>
                  </div>
                  <div className="card-body">
                    <div className="upload-dropzone" onClick={() => document.getElementById('file-input-entradas').click()}>
                      <div className="upload-icon">
                        <FileSpreadsheet size={24} />
                      </div>
                      <div>
                        <div className="upload-title">Sube las Entradas / Llegadas</div>
                        <div className="upload-desc">Contiene el albarán de lo que físicamente ha llegado al almacén para cubrir las necesidades (valores positivos).</div>
                      </div>
                      <input 
                        id="file-input-entradas" 
                        type="file" 
                        accept=".xlsx,.xls" 
                        onChange={importarExcelEntradas} 
                      />
                    </div>
                  </div>
                </Card>
              </div>

              {/* AJUSTES GLOBALES RÁPIDOS */}
              <Card>
                <div className="card-header">
                  <h2><Settings size={18} /> Configuración de Tolerancia de Talla Global</h2>
                </div>
                <div className="card-body" style={{ display: "flex", gap: "20px", alignItems: "center", flexWrap: "wrap" }}>
                  <div className="form-group" style={{ minWidth: "220px" }}>
                    <label className="form-label">Tolerancia de talla global</label>
                    <input 
                      type="number" 
                      min="0" 
                      className="input-control" 
                      value={toleranciaGlobal} 
                      onChange={(e) => setToleranciaGlobal(Math.max(0, Number(e.target.value)))} 
                    />
                  </div>
                  <div style={{ flex: 1, minWidth: "280px" }}>
                    <p style={{ fontSize: "13px", color: "var(--text-muted)", margin: 0 }}>
                      Determina cuántos niveles de diferencia se permiten para compensar una talla en negativo (necesidad) con stock disponible de otra talla. 
                      Por ejemplo, una tolerancia de <b>2</b> permite compensar una necesidad de talla 40 con existencias disponibles de talla 42 o 41.
                    </p>
                  </div>
                </div>
              </Card>

              {/* ALERTAS Y ADVERTENCIAS */}
              {avisos.length > 0 && (
                <Card>
                  <div className="card-header" style={{ backgroundColor: "#fffbeb" }}>
                    <h2 style={{ color: "#92400e" }}><AlertTriangle size={18} /> Alertas de Consistencia ({avisos.length})</h2>
                  </div>
                  <div className="card-body" style={{ display: "flex", flexDirection: "column", gap: "10px" }}>
                    {avisos.map((aviso, idx) => (
                      <div key={idx} className={`alert-banner ${aviso.tipo === "error" ? "error" : aviso.tipo === "warning" ? "warning" : "info"}`}>
                        <Info size={16} style={{ flexShrink: 0, marginTop: "2px" }} />
                        <div>{aviso.mensaje}</div>
                      </div>
                    ))}
                  </div>
                </Card>
              )}
            </motion.div>
          )}

          {activeTab === "precios" && (
            <motion.div 
              key="precios"
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -15 }}
              transition={{ duration: 0.2 }}
              style={{ display: "flex", flexDirection: "column", gap: "24px" }}
            >
              <Card>
                <div className="card-header">
                  <h2><Euro size={18} /> Tarifas y Asociaciones de Agrupación por Prenda</h2>
                </div>
                <div className="card-body">
                  <p style={{ fontSize: "14px", color: "var(--text-muted)", marginBottom: "20px" }}>
                    Configura el coste unitario de cada prenda, su tolerancia y agrupa nombres similares. Si seleccionas otra prenda en la columna <b>"Agrupar bajo"</b>, el recomendador considerará ambos nombres como la misma prenda y compensará sus tallas para equilibrar los stocks (buscar el balance cero).
                  </p>

                  {todasLasPrendas.length === 0 ? (
                    <div className="empty-state">
                      <Shirt size={48} className="empty-state-icon" />
                      <p>Sube archivos Excel o introduce datos manuales para configurar las tarifas por prenda.</p>
                    </div>
                  ) : (
                    <div className="table-wrapper">
                      <table className="app-table">
                        <thead>
                          <tr>
                            <th>Artículo detectado</th>
                            <th>Coste Unitario (€)</th>
                            <th>Tolerancia Particular</th>
                            <th>Dirección de Tolerancia</th>
                            <th>Agrupar bajo (Asociación)</th>
                            <th>Detalle de Regla</th>
                          </tr>
                        </thead>
                        <tbody>
                          {todasLasPrendas.map(prenda => {
                            const key = normalizarTexto(prenda);
                            const conf = configuraciones[key] || { precio: 20, tolerancia: 2, direccion: "ambos" };

                            const cambiarPrecio = (val) => {
                              setConfiguraciones(prev => ({
                                ...prev,
                                [key]: { ...conf, precio: Math.max(0, Number(val)) }
                              }));
                            };

                            const cambiarTolerancia = (val) => {
                              setConfiguraciones(prev => ({
                                ...prev,
                                [key]: { ...conf, tolerancia: Math.max(0, Number(val)) }
                              }));
                            };

                            const cambiarDireccion = (val) => {
                              setConfiguraciones(prev => ({
                                ...prev,
                                [key]: { ...conf, direccion: val }
                              }));
                            };

                            const cambiarGrupo = (val) => {
                              setConfiguraciones(prev => ({
                                ...prev,
                                [key]: { ...conf, grupo: val || undefined }
                              }));
                            };

                            return (
                              <tr key={prenda}>
                                <td className="bold" style={{ textTransform: "capitalize" }}>{prenda}</td>
                                <td>
                                  <div style={{ display: "flex", alignItems: "center", gap: "4px", maxWidth: "110px" }}>
                                    <input 
                                      type="number"
                                      step="0.01"
                                      min="0"
                                      className="input-control"
                                      value={conf.precio}
                                      onChange={(e) => cambiarPrecio(e.target.value)}
                                    />
                                    <span>€</span>
                                  </div>
                                </td>
                                <td>
                                  <div style={{ maxWidth: "80px" }}>
                                    <input 
                                      type="number"
                                      min="0"
                                      className="input-control"
                                      value={conf.tolerancia}
                                      onChange={(e) => cambiarTolerancia(e.target.value)}
                                    />
                                  </div>
                                </td>
                                <td>
                                  <select 
                                    className="input-control"
                                    value={conf.direccion}
                                    onChange={(e) => cambiarDireccion(e.target.value)}
                                    style={{ maxWidth: "150px" }}
                                  >
                                    <option value="ambos">Cualquiera (±)</option>
                                    <option value="superior">Solo Superior (+)</option>
                                    <option value="inferior">Solo Inferior (-)</option>
                                  </select>
                                </td>
                                <td>
                                  <select 
                                    className="input-control"
                                    value={conf.grupo || ""}
                                    onChange={(e) => cambiarGrupo(e.target.value)}
                                    style={{ maxWidth: "200px", fontWeight: conf.grupo ? "700" : "normal", borderColor: conf.grupo ? "var(--primary)" : "var(--border-color)" }}
                                  >
                                    <option value="">— (Grupo Propio)</option>
                                    {todasLasPrendas.filter(p => normalizarTexto(p) !== key).map(p => (
                                      <option key={p} value={normalizarTexto(p)}>{p}</option>
                                    ))}
                                  </select>
                                </td>
                                <td style={{ fontSize: "11px", color: "var(--text-muted)", maxWidth: "200px" }}>
                                  {conf.grupo ? (
                                    <span style={{ color: "var(--primary)", fontWeight: "600" }}>
                                      Agrupado bajo "{todasLasPrendas.find(p => normalizarTexto(p) === conf.grupo) || conf.grupo}"
                                    </span>
                                  ) : conf.tolerancia === 0 ? (
                                    <span style={{ color: "var(--danger-dark)", fontWeight: "600" }}>Talla exacta</span>
                                  ) : (
                                    <span>
                                      Sustitución <b>±{conf.tolerancia}</b> ({
                                        conf.direccion === "superior" ? "solo mayor" : 
                                        conf.direccion === "inferior" ? "solo menor" : "ambas"
                                      })
                                    </span>
                                  )}
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>
              </Card>

              {/* EXPLICACION DE TALLAS */}
              <Card>
                <div className="card-header">
                  <h2><HelpCircle size={18} /> Guía de Soporte para Tallas No Numéricas</h2>
                </div>
                <div className="card-body">
                  <div style={{ display: "flex", flexDirection: "column", gap: "12px", fontSize: "14px" }}>
                    <p>
                      El recomendador inteligente es compatible tanto con tallas numéricas como con tallas alfabéticas o rangos complejos:
                    </p>
                    <ul style={{ paddingLeft: "20px", display: "flex", flexDirection: "column", gap: "8px" }}>
                      <li>
                        <b>Tallas Numéricas:</b> Se calcula la distancia matemática normal. (Ej: pedido de 42, entregado 43, distancia = 1).
                      </li>
                      <li>
                        <b>Tallas Alfabéticas:</b> Mapeadas secuencialmente como: <code>XXS &lt; XS &lt; S &lt; M &lt; L &lt; XL &lt; XXL (2XL) &lt; 3XL &lt; 4XL</code>. Una tolerancia de 1 permite entregar M o XL a quien pidió L.
                      </li>
                      <li>
                        <b>Rangos combinados:</b> Formatos como <code>38/40</code> o <code>M-L</code> se calculan usando su punto medio. <code>38/40</code> equivale a 39 y <code>M-L</code> a una posición intermedia para buscar la talla más cercana adecuada.
                      </li>
                      <li>
                        <b>Compensación:</b> Si hay un empate en distancias (ej: pedido M y stock disponible S y L, ambos a distancia 1), la app prioriza de manera automática la <b>talla superior (L)</b> para evitar ropa pequeña.
                      </li>
                    </ul>
                  </div>
                </div>
              </Card>
            </motion.div>
          )}

          {activeTab === "editores" && (
            <motion.div 
              key="editores"
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -15 }}
              transition={{ duration: 0.2 }}
              style={{ display: "flex", flexDirection: "column", gap: "24px" }}
            >
              <div style={{ display: "flex", flexDirection: "column", gap: "24px" }}>
                <div className="grid-two">
                  <EditorTabla 
                    titulo="1. Inventario / Stock Inicial Positivo (Tengo)" 
                    icono={Shirt} 
                    filas={stock} 
                    setFilas={setStock}
                    todasLasPrendas={todasLasPrendas}
                    placeholderPrenda="Pantalones negros V"
                  />
                  <EditorTabla 
                    titulo="2. Nuevas Entradas Positivas (Llega)" 
                    icono={Download} 
                    filas={entradas} 
                    setFilas={setEntradas}
                    todasLasPrendas={todasLasPrendas}
                    placeholderPrenda="Pantalones negros V"
                  />
                </div>
                <div style={{ maxWidth: "50%", alignSelf: "center", width: "100%" }}>
                  <EditorTabla 
                    titulo="3. Necesidades / Pedidos (Valores en Negativo ERP)" 
                    icono={ClipboardList} 
                    filas={pedidos} 
                    setFilas={setPedidos}
                    todasLasPrendas={todasLasPrendas}
                    placeholderPrenda="Pantalones V"
                  />
                </div>
              </div>
            </motion.div>
          )}

          {activeTab === "recomendaciones" && (
            <motion.div 
              key="recomendaciones"
              initial={{ opacity: 0, y: 15 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -15 }}
              transition={{ duration: 0.2 }}
              style={{ display: "flex", flexDirection: "column", gap: "24px" }}
            >
              {/* DESGLOSE ECONÓMICO RESUMIDO POR ARTÍCULO */}
              <Card>
                <div className="card-header">
                  <h2><Euro size={18} /> Resumen de Costes y Análisis Económico de Inventario y Movimientos</h2>
                  <button className="btn btn-success" onClick={descargarExcel}>
                    <Download size={16} /> Exportar Informe Completo (Excel)
                  </button>
                </div>
                <div className="card-body">
                  {resumen.length === 0 ? (
                    <div className="empty-state">No hay suficientes datos cargados para generar un resumen.</div>
                  ) : (
                    <div className="table-wrapper">
                      <table className="app-table">
                        <thead>
                          <tr>
                            <th>Artículo</th>
                            <th className="align-right">Precio</th>
                            <th className="align-right">Stock Inic. (+)</th>
                            <th className="align-right">Entradas (Llega)</th>
                            <th className="align-right">Total Disp.</th>
                            <th className="align-right">Necesidad (-)</th>
                            <th className="align-right">Val. Pedido</th>
                            <th className="align-right">Compensado</th>
                            <th className="align-right">Ahorro (€)</th>
                            <th className="align-right">Déficit Neto</th>
                            <th className="align-right">Presupuesto Compra</th>
                            <th className="align-right">Sobrante (+)</th>
                            <th className="align-right">Val. Sobrante</th>
                            <th className="align-right" style={{ backgroundColor: "#f1f5f9" }}>Balance Neto Final</th>
                          </tr>
                        </thead>
                        <tbody>
                          {resumen.map(r => {
                            const conf = configuraciones[r.articuloKey] || {};
                            const precio = conf.precio !== undefined ? Number(conf.precio) : 20;
                            const balanceNeto = r.sobrante - r.totalFalta;

                            return (
                              <tr key={r.articuloKey}>
                                <td className="bold" style={{ textTransform: "capitalize" }}>{r.articulo}</td>
                                <td className="align-right">{precio.toLocaleString('es-ES', { style: 'currency', currency: 'EUR' })}</td>
                                <td className="align-right">{r.totalStockInicial}</td>
                                <td className="align-right" style={{ color: "var(--primary)" }}>+{r.totalEntradas}</td>
                                <td className="align-right bold">{r.totalStock}</td>
                                <td className="align-right" style={{ color: "var(--danger-dark)" }}>-{r.totalPedido}</td>
                                <td className="align-right">{(r.totalPedido * precio).toLocaleString('es-ES', { style: 'currency', currency: 'EUR' })}</td>
                                <td className="align-right" style={{ color: "var(--success-dark)" }}>{r.totalEntregado}</td>
                                <td className="align-right" style={{ color: "var(--success-dark)", fontWeight: "600" }}>
                                  {(r.totalEntregado * precio).toLocaleString('es-ES', { style: 'currency', currency: 'EUR' })}
                                </td>
                                <td className="align-right" style={{ color: r.totalFalta > 0 ? "var(--danger-dark)" : "inherit" }}>-{r.totalFalta}</td>
                                <td className="align-right" style={{ color: r.totalFalta > 0 ? "var(--danger-dark)" : "inherit", fontWeight: "600" }}>
                                  {(r.totalFalta * precio).toLocaleString('es-ES', { style: 'currency', currency: 'EUR' })}
                                </td>
                                <td className="align-right">+{r.sobrante}</td>
                                <td className="align-right">{(r.sobrante * precio).toLocaleString('es-ES', { style: 'currency', currency: 'EUR' })}</td>
                                <td className="align-right bold" style={{ 
                                  backgroundColor: "#f8fafc", 
                                  color: balanceNeto > 0 ? "var(--success-dark)" : balanceNeto < 0 ? "var(--danger-dark)" : "var(--text-main)" 
                                }}>
                                  {balanceNeto > 0 ? `+${balanceNeto}` : balanceNeto} ({(balanceNeto * precio).toLocaleString('es-ES', { style: 'currency', currency: 'EUR' })})
                                </td>
                              </tr>
                            );
                          })}
                          <tr className="table-row-totals">
                            <td>TOTALES</td>
                            <td className="align-right">—</td>
                            <td className="align-right">{resumen.reduce((s, r) => s + r.totalStockInicial, 0)}</td>
                            <td className="align-right" style={{ color: "var(--primary)" }}>+{resumen.reduce((s, r) => s + r.totalEntradas, 0)}</td>
                            <td className="align-right">{resumen.reduce((s, r) => s + r.totalStock, 0)}</td>
                            <td className="align-right" style={{ color: "var(--danger-dark)" }}>-{resumen.reduce((s, r) => s + r.totalPedido, 0)}</td>
                            <td className="align-right">{kpis.valorPedidosTotal.toLocaleString('es-ES', { style: 'currency', currency: 'EUR' })}</td>
                            <td className="align-right">{resumen.reduce((s, r) => s + r.totalEntregado, 0)}</td>
                            <td className="align-right">{kpis.valorEntregadoTotal.toLocaleString('es-ES', { style: 'currency', currency: 'EUR' })}</td>
                            <td className="align-right" style={{ color: "var(--danger-dark)" }}>-{resumen.reduce((s, r) => s + r.totalFalta, 0)}</td>
                            <td className="align-right" style={{ color: "var(--danger-dark)" }}>
                              {kpis.valorFaltanteTotal.toLocaleString('es-ES', { style: 'currency', currency: 'EUR' })}
                            </td>
                            <td className="align-right">+{resumen.reduce((s, r) => s + r.sobrante, 0)}</td>
                            <td className="align-right">{kpis.valorSobranteTotal.toLocaleString('es-ES', { style: 'currency', currency: 'EUR' })}</td>
                            <td className="align-right bold" style={{ 
                              backgroundColor: "#f1f5f9", 
                              color: (kpis.valorSobranteTotal - kpis.valorFaltanteTotal) > 0 ? "var(--success-dark)" : (kpis.valorSobranteTotal - kpis.valorFaltanteTotal) < 0 ? "var(--danger-dark)" : "var(--text-main)"
                            }}>
                              {resumen.reduce((s, r) => s + (r.sobrante - r.totalFalta), 0) > 0 ? `+` : ""}{resumen.reduce((s, r) => s + (r.sobrante - r.totalFalta), 0)} ({(kpis.valorSobranteTotal - kpis.valorFaltanteTotal).toLocaleString('es-ES', { style: 'currency', currency: 'EUR' })})
                            </td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>
              </Card>

              {/* LISTA DETALLADA DE RECOMENDACIONES DE ENTREGA LOGÍSTICA */}
              <Card>
                <div className="card-header">
                  <h2><Sliders size={18} /> Detalle de Distribución y Recomendaciones</h2>
                </div>
                <div className="card-body">
                  {/* BARRA DE FILTROS */}
                  <div className="filters-bar">
                    <div className="filters-inputs">
                      <div className="form-group" style={{ minWidth: "150px" }}>
                        <label className="form-label">Filtrar por Prenda</label>
                        <select 
                          className="input-control"
                          value={filtroArticulo}
                          onChange={(e) => setFiltroArticulo(e.target.value)}
                        >
                          <option value="">Todos los artículos</option>
                          {todasLasPrendas.map(prenda => (
                            <option key={prenda} value={prenda}>{prenda}</option>
                          ))}
                        </select>
                      </div>

                      <div className="form-group" style={{ minWidth: "150px" }}>
                        <label className="form-label">Estado Asignación</label>
                        <select 
                          className="input-control"
                          value={filtroResultado}
                          onChange={(e) => setFiltroResultado(e.target.value)}
                        >
                          <option value="">Todos los resultados</option>
                          <option value="Exacta">Talla Exacta</option>
                          <option value="Sustitución">Sustituciones</option>
                          <option value="Falta">Faltas para Compra</option>
                        </select>
                      </div>

                      <div className="form-group" style={{ minWidth: "220px" }}>
                        <label className="form-label">Buscar texto...</label>
                        <input 
                          className="input-control"
                          value={busqueda}
                          onChange={(e) => setBusqueda(e.target.value)}
                          placeholder="Buscar por prenda, talla..."
                        />
                      </div>
                    </div>
                    
                    <div style={{ fontSize: "13px", color: "var(--text-muted)" }}>
                      Mostrando <b>{recomendacionesFiltradas.length}</b> de <b>{recomendaciones.length}</b> asignaciones
                    </div>
                  </div>

                  {/* TABLA DE ASIGNACIONES */}
                  {recomendacionesFiltradas.length === 0 ? (
                    <div className="empty-state">
                      <Filter size={36} className="empty-state-icon" />
                      <p>No se encontraron recomendaciones con los filtros aplicados.</p>
                    </div>
                  ) : (
                    <div className="table-wrapper">
                      <table className="app-table">
                        <thead>
                          <tr>
                            <th>Artículo Solicitado</th>
                            <th>Talla Pedida</th>
                            <th>Talla Entregada</th>
                            <th>Origen Logístico</th>
                            <th className="align-right">Cantidad (Uds)</th>
                            <th className="align-right">Coste Unitario</th>
                            <th className="align-right">Valor Total</th>
                            <th>Resultado</th>
                          </tr>
                        </thead>
                        <tbody>
                          {recomendacionesFiltradas.map((rec, idx) => {
                            const key = normalizarTexto(rec.articulo);
                            const precio = configuraciones[key]?.precio ?? 20;

                            return (
                              <tr key={idx}>
                                <td className="bold" style={{ textTransform: "capitalize" }}>{rec.articulo}</td>
                                <td><span className="badge badge-neutral">Talla {rec.tallaPedido}</span></td>
                                <td>
                                  {rec.tallaEntregada ? (
                                    <div>
                                      <span className="badge badge-primary">Talla {rec.tallaEntregada}</span>
                                      {rec.articuloStock && normalizarTexto(rec.articulo) !== normalizarTexto(rec.articuloStock) && (
                                        <div style={{ fontSize: "11px", color: "var(--text-light)", marginTop: "2px" }}>
                                          de {rec.articuloStock}
                                        </div>
                                      )}
                                    </div>
                                  ) : (
                                    <span style={{ color: "var(--text-light)" }}>—</span>
                                  )}
                                </td>
                                <td>
                                  {rec.origenStock === "Inventario" && (
                                    <span className="badge badge-neutral" style={{ fontSize: "11px" }}>Almacén</span>
                                  )}
                                  {rec.origenStock === "Entrada" && (
                                    <span className="badge badge-success" style={{ fontSize: "11px", backgroundColor: "#ecfdf5", color: "#065f46", border: "1px solid #a7f3d0" }}>Llegada</span>
                                  )}
                                  {!rec.origenStock && (
                                    <span style={{ color: "var(--text-light)" }}>—</span>
                                  )}
                                </td>
                                <td className="align-right bold">{rec.cantidad}</td>
                                <td className="align-right">{precio.toLocaleString('es-ES', { style: 'currency', currency: 'EUR' })}</td>
                                <td className="align-right bold">{(rec.cantidad * precio).toLocaleString('es-ES', { style: 'currency', currency: 'EUR' })}</td>
                                <td>
                                  {rec.tipo === "Exacta" && (
                                    <span className="badge badge-success">Exacta</span>
                                  )}
                                  {rec.tipo === "Sustitución" && (
                                    <span className="badge badge-warning">Sustitución</span>
                                  )}
                                  {rec.tipo === "Falta" && (
                                    <span className="badge badge-danger">Compra Necesaria</span>
                                  )}
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>
              </Card>
            </motion.div>
          )}
        </AnimatePresence>
      </main>
    </div>
  );
}
