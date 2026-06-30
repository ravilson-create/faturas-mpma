import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";
import {
  LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
} from "recharts";

const SUPABASE_URL = "https://zvwwquwlesjppmzkhmtn.supabase.co";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inp2d3dxdXdsZXNqcHBtemtobXRuIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODE3OTM3MTUsImV4cCI6MjA5NzM2OTcxNX0.HALwLmkZ_xCs7wfojDJ4vaBRO7Y-KkDNCGs2cfj8yBY";
const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

const COPPER = "#BA7517";
const COPPER_DARK = "#633806";
const INK = "#2C2C2A";
const MUTED = "#5F5E5A";
const BORDER = "#D3D1C7";
const BG = "#F7F6F2";
const CARD = "#FFFFFF";

function fmtCurrency(v) {
  if (v == null || isNaN(v)) return "-";
  return v.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
}
function fmtNumber(v, decimals = 0) {
  if (v == null || isNaN(v)) return "-";
  return v.toLocaleString("pt-BR", { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
}
function monthLabel(dateStr) {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  const meses = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];
  return `${meses[d.getMonth()]}/${d.getFullYear()}`;
}

function normalizeItem(raw) {
  if (!raw) return "";
  let s = String(raw).trim().replace(/\s+/g, " ");
  const fixes = {
    "Ã§Ã£o": "ção", "Ã§": "ç", "Ã£": "ã", "Ã©": "é", "Ãª": "ê", "Ã³": "ó",
    "Ã­": "í", "Ã¡": "á", "Ãº": "ú", "Ã‰": "É", "Ã": "Á",
  };
  for (const [bad, good] of Object.entries(fixes)) s = s.split(bad).join(good);
  return s;
}

function parseValor(v) {
  if (v == null) return 0;
  if (typeof v === "number") return v;
  const s = String(v).trim();
  if (s === "" || s.toUpperCase().includes("R$ -") || s === "-") return 0;
  const cleaned = s.replace(/R\$\s*/g, "").replace(/\./g, "").replace(",", ".");
  const n = parseFloat(cleaned);
  return isNaN(n) ? 0 : n;
}

function toISODate(v) {
  if (v == null) return null;
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if (typeof v === "number") {
    const d = XLSX.SSF.parse_date_code(v);
    if (!d) return null;
    return `${d.y}-${String(d.m).padStart(2, "0")}-${String(d.d).padStart(2, "0")}`;
  }
  const d = new Date(v);
  if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
  return null;
}

const ITEM_MAP = {
  "Consumo": { consumo: "kwh", valor: "val_consumo" },
  "Consumo Ponta": { consumo: "consumo_ponta_kwh", valor: "val_consumo_ponta" },
  "Consumo Fora Ponta": { consumo: "consumo_fora_ponta_kwh", valor: "val_consumo_fora_ponta" },
  "Demanda Ativa": { consumo: "demanda_ativa_kw", valor: "val_demanda_ativa" },
  "Demanda Ponta": { consumo: "demanda_ponta_kw", valor: "val_demanda_ponta" },
  "Demanda Fora Ponta": { consumo: "demanda_fora_ponta_kw", valor: "val_demanda_fora_ponta" },
  "Demanda Ultrapassagem": { consumo: "demanda_ultrapassagem_kw", valor: "val_demanda_ultrapassagem" },
  "Demanda Ultrapassagem Ponta": { consumo: "demanda_ultrapassagem_ponta_kw", valor: "val_demanda_ultrapassagem_ponta" },
  "Demanda de Geração": { consumo: null, valor: "val_demanda_geracao" },
  "Consumo Reativo Excedente": { consumo: "reativo_excedente_kwh", valor: "val_reativo_excedente" },
  "Consumo Reativo Excedente FP": { consumo: "reativo_excedente_fp_kwh", valor: "val_reativo_excedente_fp" },
  "Consumo Reativo Excedente NP": { consumo: "reativo_excedente_np_kwh", valor: "val_reativo_excedente_np" },
  "Adicional Bandeira": { consumo: null, valor: "val_bandeira" },
  "Cip-Ilum Pub Pref Munic": { consumo: null, valor: "cip" },
  "Tributo a Reter IRPJ": { consumo: null, valor: "irpj", abs: true },
  "Tributo a Reter CSLL": { consumo: null, valor: "csll", abs: true },
  "Tributo a Reter PIS": { consumo: null, valor: "pis_retido", abs: true },
  "Tributo a Reter COFINS": { consumo: null, valor: "cofins_retido", abs: true },
  // Itens do novo formato (jan/2023–mai/2026)
  "Crédito Nível de Tensão": { consumo: null, valor: "credito_nivel_tensao" },
  "Estorno Crédito DIC/FIC": { consumo: null, valor: "estorno_credito_dic_fic" },
  "Correção Monetária CNR": { consumo: null, valor: "correcao_monetaria_cnr" },
  "Dev.Pagto.Indevido Fat.Anter.": { consumo: null, valor: "dev_pagto_indevido" },
  "Religação urgência/programada": { consumo: null, valor: "religacao_urgencia" },
  "Dev Geração": { consumo: null, valor: "dev_geracao" },
  "Dev Geração NP": { consumo: null, valor: "dev_geracao_np" },
  "Dev Geração FP": { consumo: null, valor: "dev_geracao_fp" },
  "Multa": { consumo: null, valor: "multa" },
  "Juros": { consumo: null, valor: "juros" },
  "Correção Monetária": { consumo: null, valor: "correcao_monetaria" },
  "Crédito DIC/FIC/DMIC": { consumo: null, valor: "credito_dic_fic_dmic" },
  "Crédito DICRI": { consumo: null, valor: "credito_dicri" },
  "Crédito Prazo Atendimento": { consumo: null, valor: "credito_prazo_atendimento" },
  "Crédito pagamento por conta": { consumo: null, valor: "credito_pagamento_conta" },
  "Taxa Religação": { consumo: null, valor: "taxa_religacao" },
  "Ligação Provisória": { consumo: null, valor: "ligacao_provisoria" },
  "- Parcelamento": { consumo: null, valor: "parcelamento" },
  "- Base ICMS": { consumo: null, valor: "icms_base" },
  "- Valor ICMS": { consumo: null, valor: "icms_val" },
  "- Base COFINS": { consumo: null, valor: "cofins_base" },
  "- Valor COFINS": { consumo: null, valor: "cofins_val" },
  "- Base PIS": { consumo: null, valor: "pis_base" },
  "- Valor PIS": { consumo: null, valor: "pis_val" },
  "- VALOR TOTAL": { consumo: null, valor: "valor_total" },
};

function parseWorkbook(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const rowsRaw = XLSX.utils.sheet_to_json(ws, { defval: null, raw: true });

  // Remove linhas totalmente idênticas (a exportação da Equatorial às vezes
  // repete a mesma linha de cobrança, o que inflava demanda e valores ao somar)
  const vistos = new Set();
  const rows = rowsRaw.filter((row) => {
    const assinatura = JSON.stringify([
      row["NUMERO_FATURA"], row["DETALHAMENTO_CONTA"], row["CONSUMO"], row["VALOR_FATURA"]
    ]);
    if (vistos.has(assinatura)) return false;
    vistos.add(assinatura);
    return true;
  });

  const byFatura = new Map();

  for (const row of rows) {
    const numeroFatura = row["NUMERO_FATURA"];
    if (numeroFatura == null) continue;
    const key = String(numeroFatura);

    if (!byFatura.has(key)) {
      byFatura.set(key, {
        fatura: key,
        instalacao: String(row["INSTALACAO"] ?? ""),
        fatura_agrupadora: String(row["FATURA_AGRUPADORA"] ?? ""),
        municipio: row["MUNICIPIO"] ?? "",
        complemento: row["COMPLEMENTO"] ?? "",
        endereco: row["ENDERECO"] ?? "",
        categoria: row["CATEGORIA"] ?? "",
        tipo_tarifa: row["TIPO_TARIFA"] ?? null,
        mes_referencia: toISODate(row["MES_REFERENCIA"]),
        vencimento: toISODate(row["VENCIMENTO"]),
        data_leitura_anterior: toISODate(row["DATA_LEITURA_ANTERIOR"]),
        data_leitura_atual: toISODate(row["DATA_LEITURA_ATUAL"]),
        dic_apurado: row["DIC_APURADO"] ?? null,
        fic_apurado: row["FIC_APURADO"] ?? null,
      });
    }
    const reg = byFatura.get(key);

    const itemRaw = normalizeItem(row["DETALHAMENTO_CONTA"]);
    const mapping = ITEM_MAP[itemRaw];
    const consumoCell = row["CONSUMO"];
    const valorCell = parseValor(row["VALOR_FATURA"]);

    if (mapping) {
      if (mapping.consumo && consumoCell != null) {
        const num = Number(consumoCell) || 0;
        // Campos de DEMANDA (kW) são valores de pico — usa-se o MAIOR valor positivo
        // registrado, nunca a soma (somar linhas repetidas inflava a demanda medida
        // e gerava ultrapassagem falsa). Demais campos (kWh) acumulam normalmente.
        const ehDemanda = mapping.consumo.startsWith("demanda_");
        if (ehDemanda) {
          reg[mapping.consumo] = Math.max(reg[mapping.consumo] || 0, num);
        } else {
          reg[mapping.consumo] = (reg[mapping.consumo] || 0) + num;
        }
      }
      if (mapping.valor) {
        // Tributos retidos vêm com sinal negativo no arquivo — armazenar como positivo
        const v = mapping.abs ? Math.abs(valorCell) : valorCell;
        reg[mapping.valor] = (reg[mapping.valor] || 0) + v;
      }
    } else if (!itemRaw && valorCell !== 0) {
      reg.nao_identificado = (reg.nao_identificado || 0) + valorCell;
    }

    // Demanda contratada: coluna DEMANDA (tarifas convencionais)
    if (row["DEMANDA"] != null && !isNaN(Number(row["DEMANDA"])) && Number(row["DEMANDA"]) > 0) {
      reg.demanda_contratada_kw = Number(row["DEMANDA"]);
    }
    // Demanda contratada horosazonal: DEMANDA_PONTA e DEMANDA_FORA_PONTA
    // (aparecem nas linhas de Demanda Ultrapassagem/Ativa para tarifas horo-sazonais)
    if (row["DEMANDA_PONTA"] != null && !isNaN(Number(row["DEMANDA_PONTA"])) && Number(row["DEMANDA_PONTA"]) > 0) {
      reg.demanda_contratada_ponta_kw = Number(row["DEMANDA_PONTA"]);
    }
    if (row["DEMANDA_FORA_PONTA"] != null && !isNaN(Number(row["DEMANDA_FORA_PONTA"])) && Number(row["DEMANDA_FORA_PONTA"]) > 0) {
      reg.demanda_contratada_fora_ponta_kw = Number(row["DEMANDA_FORA_PONTA"]);
    }

    // Flag direta de ultrapassagem: se a fatura tem item de Demanda Ultrapassagem com valor > 0
    if ((itemRaw === "Demanda Ultrapassagem" || itemRaw === "Demanda Ultrapassagem Ponta") && valorCell > 0) {
      reg.tem_ultrapassagem = true;
      reg.val_ultrapassagem_total = (reg.val_ultrapassagem_total || 0) + valorCell;
    }
  }

  const registros = Array.from(byFatura.values())
    // Faturas de estorno têm VALOR TOTAL negativo ou zero — descarta para não contaminar somas
    .filter((r) => (r.valor_total == null || r.valor_total > 0))
    .map((r) => {
      const consumoTotal = (r.kwh || 0) + (r.consumo_ponta_kwh || 0) + (r.consumo_fora_ponta_kwh || 0);
      const devGeracaoTotal = (r.dev_geracao || 0) + (r.dev_geracao_np || 0) + (r.dev_geracao_fp || 0);
      const temGD = devGeracaoTotal !== 0 || (r.val_demanda_geracao || 0) !== 0;
      return { ...r, consumo_total_kwh: consumoTotal, uc: r.instalacao, dev_geracao_total: devGeracaoTotal, tem_gd: temGD };
    });

  return registros;
}

export default function FaturasApp() {
  const [registros, setRegistros] = useState([]);
  const [loading, setLoading] = useState(true);
  const [processing, setProcessing] = useState(false);
  const [error, setError] = useState(null);
  const [search, setSearch] = useState("");
  const [selectedUC, setSelectedUC] = useState(null);

  // Controle de acesso por senha (válida por sessão)
  const [autenticado, setAutenticado] = useState(false);
  const [senhaInput, setSenhaInput] = useState("");
  const [erroSenha, setErroSenha] = useState(false);
  const [mostrarPainelAdmin, setMostrarPainelAdmin] = useState(false);

  // Filtro por ano e gerenciamento de períodos
  const [anoSelecionado, setAnoSelecionado] = useState("todos");
  const [periodoInicio, setPeriodoInicio] = useState("");
  const [periodoFim, setPeriodoFim] = useState("");
  const [usarFiltroPeriodo, setUsarFiltroPeriodo] = useState(false);
  const [excluindoPeriodo, setExcluindoPeriodo] = useState(false);
  const [confirmacaoExclusao, setConfirmacaoExclusao] = useState(null); // { tipo, valor, label }

  const SENHA_ADMIN = "mpma@coea";
  const fileInputRef = useRef(null);

  const carregarRegistros = useCallback(async () => {
    setLoading(true);
    try {
      const PAGE = 1000;
      let todos = [];
      let pagina = 0;
      while (true) {
        const { data, error: err } = await supabase
          .from("faturas")
          .select("*")
          .order("mes_referencia", { ascending: true })
          .range(pagina * PAGE, (pagina + 1) * PAGE - 1);
        if (err) throw new Error(err.message);
        if (!data || data.length === 0) break;
        todos = todos.concat(data);
        if (data.length < PAGE) break; // última página
        pagina++;
      }
      setRegistros(todos);
    } catch (e) {
      setError("Não foi possível carregar os dados do banco: " + e.message);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    carregarRegistros();
  }, [carregarRegistros]);

  const handleAutenticar = useCallback(() => {
    if (senhaInput === SENHA_ADMIN) {
      setAutenticado(true);
      setErroSenha(false);
      setSenhaInput("");
      setMostrarPainelAdmin(true);
    } else {
      setErroSenha(true);
    }
  }, [senhaInput]);

  const handleExcluirPeriodo = useCallback(async ({ tipo, valor }) => {
    setExcluindoPeriodo(true);
    setError(null);
    try {
      let query = supabase.from("faturas").delete();
      if (tipo === "mes") {
        // valor = "2025-04-01"
        query = query.eq("mes_referencia", valor);
      } else if (tipo === "ano") {
        // valor = "2025"
        query = query.gte("mes_referencia", `${valor}-01-01`).lte("mes_referencia", `${valor}-12-31`);
      } else if (tipo === "uc") {
        // valor = instalacao
        query = query.eq("uc", valor);
      }
      const { error: delErr } = await query;
      if (delErr) throw new Error(delErr.message);
      setConfirmacaoExclusao(null);
      setSelectedUC(null);
      await carregarRegistros();
    } catch (e) {
      setError("Erro ao excluir: " + e.message);
    } finally {
      setExcluindoPeriodo(false);
    }
  }, [carregarRegistros]);

  const handleFile = useCallback(async (file) => {
    if (!file) return;
    const isExcel = /\.(xlsx|xls)$/i.test(file.name);
    if (!isExcel) {
      setError("Selecione um arquivo Excel (.xlsx ou .xls).");
      return;
    }
    setError(null);
    setProcessing(true);
    try {
      const buffer = await file.arrayBuffer();
      const novosRegistros = parseWorkbook(buffer);

      // Envia em lotes de 500 para não estourar o limite do Supabase
      const LOTE = 500;
      for (let i = 0; i < novosRegistros.length; i += LOTE) {
        const lote = novosRegistros.slice(i, i + LOTE);
        const { error: upsertError } = await supabase
          .from("faturas")
          .upsert(lote, { onConflict: "fatura" });
        if (upsertError) throw new Error(upsertError.message);
      }

      await carregarRegistros();
    } catch (e) {
      setError("Falha ao processar o arquivo: " + e.message);
    } finally {
      setProcessing(false);
      if (fileInputRef.current) fileInputRef.current.value = "";
    }
  }, [carregarRegistros]);

  const ucsComGD = useMemo(() => {
    const set = new Set();
    registros.forEach((r) => { if (r.tem_gd) set.add(r.uc); });
    return set;
  }, [registros]);

  // Anos disponíveis no banco
  const anosDisponiveis = useMemo(() => {
    const anos = new Set();
    registros.forEach((r) => {
      if (r.mes_referencia) anos.add(r.mes_referencia.slice(0, 4));
    });
    return Array.from(anos).sort().reverse();
  }, [registros]);

  // Meses disponíveis (para exclusão por mês)
  const mesesDisponiveis = useMemo(() => {
    const map = new Map();
    registros.forEach((r) => {
      if (r.mes_referencia && !map.has(r.mes_referencia)) {
        map.set(r.mes_referencia, monthLabel(r.mes_referencia));
      }
    });
    return Array.from(map.entries()).sort((a, b) => b[0].localeCompare(a[0]));
  }, [registros]);

  // Converte MM/AAAA -> AAAA-MM-01 para comparação
  const registrosFiltrados = useMemo(() => {
    if (usarFiltroPeriodo) {
      const toISO = (mmaaaa) => {
        if (!mmaaaa || !/^\d{2}\/\d{4}$/.test(mmaaaa.trim())) return null;
        const [mm, yyyy] = mmaaaa.trim().split("/");
        return `${yyyy}-${mm.padStart(2,"0")}-01`;
      };
      const inicio = toISO(periodoInicio);
      const fim = toISO(periodoFim);
      return registros.filter((r) => {
        const ref = r.mes_referencia;
        if (!ref) return false;
        if (inicio && ref < inicio) return false;
        if (fim && ref > fim) return false;
        return true;
      });
    }
    if (anoSelecionado === "todos") return registros;
    return registros.filter((r) => (r.mes_referencia || "").startsWith(anoSelecionado));
  }, [registros, anoSelecionado, usarFiltroPeriodo, periodoInicio, periodoFim]);

  // Resumo do período/ano selecionado — sempre baseado em registrosFiltrados
  // Deduplica por UC+mês mantendo o registro de maior valor_total (mais completo).
  // Protege contra faturas duplicadas vindas da planilha da concessionária ou
  // de dados residuais de uploads anteriores com número de fatura diferente.
  const registrosDeduplicados = useMemo(() => {
    const porChave = new Map();
    for (const r of registrosFiltrados) {
      const chave = `${r.uc}|${r.mes_referencia}`;
      const existente = porChave.get(chave);
      if (!existente || (r.valor_total || 0) > (existente.valor_total || 0)) {
        porChave.set(chave, r);
      }
    }
    return Array.from(porChave.values());
  }, [registrosFiltrados]);

  const resumoAno = useMemo(() => {
    const base = registrosDeduplicados;
    const totalFaturado = base.reduce((acc, r) => acc + (r.valor_total || 0), 0);
    const totalConsumo = base.reduce((acc, r) => acc + (r.consumo_total_kwh || 0), 0);
    const totalCip = base.reduce((acc, r) => acc + (r.cip || 0), 0);
    const totalICMS = base.reduce((acc, r) => acc + (r.icms_val || 0), 0);
    const totalCOFINS = base.reduce((acc, r) => acc + (r.cofins_val || 0), 0);
    const totalPIS = base.reduce((acc, r) => acc + (r.pis_val || 0), 0);
    const totalGD = base.reduce((acc, r) => acc + Math.abs(r.dev_geracao_total || 0), 0);
    const meses = new Set(base.map((r) => r.mes_referencia)).size;
    const ucs = new Set(base.map((r) => r.uc)).size;
    return { totalFaturado, totalConsumo, totalCip, totalICMS, totalCOFINS, totalPIS, totalGD, meses, ucs };
  }, [registrosDeduplicados]);

  const ucList = useMemo(() => {
    const map = new Map();
    registrosFiltrados.forEach((r) => {
      if (!map.has(r.uc)) {
        map.set(r.uc, { uc: r.uc, municipio: r.municipio, complemento: r.complemento });
      }
    });
    return Array.from(map.values()).sort((a, b) => (a.municipio || "").localeCompare(b.municipio || ""));
  }, [registrosFiltrados]);

  const filteredUcList = useMemo(() => {
    if (!search.trim()) return ucList;
    const q = search.toLowerCase();
    return ucList.filter(
      (u) =>
        (u.municipio || "").toLowerCase().includes(q) ||
        (u.complemento || "").toLowerCase().includes(q) ||
        u.uc.includes(q)
    );
  }, [ucList, search]);

  const historicoUC = useMemo(() => {
    if (!selectedUC) return [];
    return registrosDeduplicados
      .filter((r) => r.uc === selectedUC)
      .sort((a, b) => new Date(a.mes_referencia) - new Date(b.mes_referencia));
  }, [registrosDeduplicados, selectedUC]);

  const isAltaTensao = useMemo(() => {
    return historicoUC.some((r) => (r.categoria || "").toUpperCase().includes("ALTA"));
  }, [historicoUC]);

  const temGeracaoDistribuida = useMemo(() => historicoUC.some((r) => r.tem_gd), [historicoUC]);

  const chartData = useMemo(() => {
    return historicoUC.map((r) => ({
      mes: monthLabel(r.mes_referencia),
      consumo: r.consumo_total_kwh,
      valor: r.valor_total,
    }));
  }, [historicoUC]);

  const chartDataGD = useMemo(() => {
    return historicoUC
      .filter((r) => r.tem_gd)
      .map((r) => ({ mes: monthLabel(r.mes_referencia), credito: Math.abs(r.dev_geracao_total || 0) }));
  }, [historicoUC]);

  const totalCreditoGD = useMemo(() => {
    return historicoUC.reduce((acc, r) => acc + Math.abs(r.dev_geracao_total || 0), 0);
  }, [historicoUC]);

  const mesesComGDeNaoIdentificado = useMemo(() => {
    return historicoUC.some((r) => r.tem_gd && (r.nao_identificado || 0) !== 0);
  }, [historicoUC]);

  const ultimoRegistro = historicoUC[historicoUC.length - 1];

  // ── Alertas ──────────────────────────────────────────────────────────────

  // Alerta 1: ultrapassagem de demanda
  // Fonte primária: item "Demanda Ultrapassagem" na fatura (tem_ultrapassagem=true)
  // Fonte secundária (tarifas convencionais): demanda medida > demanda contratada
  const alertasUltrapassagem = useMemo(() => {
    return historicoUC.filter((r) => {
      if (r.tem_ultrapassagem) return true;
      const contratada = r.demanda_contratada_kw || 0;
      const medida = r.demanda_ativa_kw || 0;
      return contratada > 0 && medida > contratada;
    });
  }, [historicoUC]);

  // Alerta 2: variação de consumo > 20% vs média dos últimos 6 meses
  const alertaConsumo = useMemo(() => {
    if (historicoUC.length < 2) return null;
    const ultimo = historicoUC[historicoUC.length - 1];
    const anteriores = historicoUC.slice(-7, -1); // até 6 meses antes do último
    if (anteriores.length === 0) return null;
    const media = anteriores.reduce((acc, r) => acc + (r.consumo_total_kwh || 0), 0) / anteriores.length;
    if (media === 0) return null;
    const variacao = ((ultimo.consumo_total_kwh || 0) - media) / media;
    // Alerta somente quando houver AUMENTO acima de 20%
    if (variacao > 0.2) {
      return { variacao, media, consumoAtual: ultimo.consumo_total_kwh, mes: ultimo.mes_referencia };
    }
    return null;
  }, [historicoUC]);

  // Alerta 3: reativo excedente + análise de capacitores (NBR 5410 / ANEEL)
  const alertasReativo = useMemo(() => {
    return historicoUC.filter((r) => {
      const reativo = (r.val_reativo_excedente || 0) + (r.val_reativo_excedente_fp || 0) + (r.val_reativo_excedente_np || 0);
      return reativo > 0;
    });
  }, [historicoUC]);

  const recomendacaoCapacitor = useMemo(() => {
    // Só calcula se houve mais de 4 ocorrências no ano calendário do último mês com reativo
    if (alertasReativo.length <= 4) return null;

    // Usa o mês com maior kVAr excedente como referência para o dimensionamento
    const pior = alertasReativo.reduce((max, r) => {
      const kvar = (r.reativo_excedente_kwh || 0) + (r.reativo_excedente_fp_kwh || 0) + (r.reativo_excedente_np_kwh || 0);
      const kvarMax = (max.reativo_excedente_kwh || 0) + (max.reativo_excedente_fp_kwh || 0) + (max.reativo_excedente_np_kwh || 0);
      return kvar > kvarMax ? r : max;
    }, alertasReativo[0]);

    const kW = pior.demanda_ativa_kw || pior.kwh || 0;
    const kVArExcedente = (pior.reativo_excedente_kwh || 0) + (pior.reativo_excedente_fp_kwh || 0) + (pior.reativo_excedente_np_kwh || 0);

    if (kW === 0 || kVArExcedente === 0) return null;

    // Limite ANEEL: FP mínimo = 0,92 indutivo
    // kVAr permitido = kW × tan(arccos(0,92))
    const FP_LIMITE = 0.92;
    const FP_ALVO = 0.95; // margem de segurança acima do mínimo legal
    const TENSAO_V = 13800; // média tensão 13,8 kV
    const FREQ_HZ = 60;

    const tanLimite = Math.tan(Math.acos(FP_LIMITE)); // ≈ 0.3936
    const tanAlvo = Math.tan(Math.acos(FP_ALVO));     // ≈ 0.3287
    const kVArPermitido = kW * tanLimite;
    const kVArTotal = kVArExcedente + kVArPermitido;
    const kVA = Math.sqrt(kW * kW + kVArTotal * kVArTotal);
    const fpAtual = kW / kVA;

    // Potência reativa necessária para corrigir de fpAtual para FP_ALVO
    const tanAtual = Math.tan(Math.acos(fpAtual));
    const Qc_kVAr = kW * (tanAtual - tanAlvo);

    // Capacitância por fase — banco trifásico em triângulo (ligação Δ), 13,8 kV
    // Q_total = 3 × V² × 2πfC  →  C = Q_total / (3 × 2πf × V²)
    const Qc_VAr = Qc_kVAr * 1000;
    const C_F = Qc_VAr / (3 * 2 * Math.PI * FREQ_HZ * TENSAO_V * TENSAO_V);
    const C_uF = C_F * 1e6;

    // Padronização para banco comercial (kVAr arredondado para múltiplo de 5 acima)
    const Qc_padronizado = Math.ceil(Qc_kVAr / 5) * 5;

    return {
      kW: kW.toFixed(1),
      kVArExcedente: kVArExcedente.toFixed(1),
      kVArTotal: kVArTotal.toFixed(1),
      fpAtual: fpAtual.toFixed(3),
      Qc_kVAr: Qc_kVAr.toFixed(1),
      Qc_padronizado,
      C_uF: C_uF.toFixed(4),
      tensao: TENSAO_V,
      mesPior: pior.mes_referencia,
      ocorrencias: alertasReativo.length,
    };
  }, [alertasReativo]);

  // Alerta 4: aumento da CIP vs mês anterior
  const alertasCIP = useMemo(() => {
    const alertas = [];
    for (let i = 1; i < historicoUC.length; i++) {
      const atual = historicoUC[i];
      const anterior = historicoUC[i - 1];
      const cipAtual = atual.cip || 0;
      const cipAnterior = anterior.cip || 0;
      if (cipAnterior > 0 && cipAtual > cipAnterior) {
        alertas.push({
          fatura: atual.fatura,
          mes: atual.mes_referencia,
          cipAtual,
          cipAnterior,
          aumento: cipAtual - cipAnterior,
          pct: ((cipAtual - cipAnterior) / cipAnterior) * 100,
        });
      }
    }
    return alertas;
  }, [historicoUC]);

  const temAlertas = alertasUltrapassagem.length > 0 || alertaConsumo !== null || alertasReativo.length > 0 || alertasCIP.length > 0;

  return (
    <div style={{ fontFamily: "Inter, system-ui, sans-serif", background: BG, minHeight: "100vh", color: INK }}>
      <div style={{ maxWidth: 1140, margin: "0 auto", padding: "32px 24px 64px" }}>

        {/* ── Cabeçalho ── */}
        <header style={{ marginBottom: 20, display: "flex", alignItems: "flex-start", justifyContent: "space-between", flexWrap: "wrap", gap: 12 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 16 }}>
            <img
              src="https://www.mpma.mp.br/wp-content/uploads/2022/09/mpma-hs.png"
              alt="MPMA"
              style={{ height: 52, objectFit: "contain" }}
              onError={(e) => { e.target.style.display = "none"; }}
            />
            <div>
              <h1 style={{ fontSize: 22, fontWeight: 700, margin: 0, letterSpacing: "-0.01em" }}>
                Consulta de faturas <span style={{ color: COPPER_DARK }}>· Energia MPMA</span>
              </h1>
              <p style={{ color: MUTED, fontSize: 13, marginTop: 4 }}>
                {ucList.length > 0
                  ? `${ucList.length} unidades consumidoras · ${resumoAno.meses} ${resumoAno.meses === 1 ? "mês" : "meses"}${ucsComGD.size > 0 ? ` · ${ucsComGD.size} com geração distribuída` : ""}`
                  : "Nenhum dado carregado. Use o painel de administração para importar planilhas."}
              </p>
            </div>
          </div>
          <button
            onClick={() => {
              if (autenticado) { setMostrarPainelAdmin((v) => !v); }
              else { setMostrarPainelAdmin(true); }
            }}
            style={{
              padding: "8px 16px", borderRadius: 8, border: `1px solid ${BORDER}`,
              background: autenticado ? "#FAEEDA" : CARD, cursor: "pointer",
              fontSize: 13, fontWeight: 600, color: COPPER_DARK, fontFamily: "inherit",
            }}
          >
            {autenticado ? (mostrarPainelAdmin ? "Fechar administração" : "Administração") : "🔒 Administração"}
          </button>
        </header>

        {/* ── Painel de administração (upload + exclusão, protegido por senha) ── */}
        {mostrarPainelAdmin && (
          <section style={{ background: CARD, border: `1px solid ${BORDER}`, borderRadius: 12, padding: "20px 22px", marginBottom: 24 }}>
            {!autenticado ? (
              <div>
                <p style={{ fontSize: 13.5, fontWeight: 600, marginBottom: 12, color: INK }}>
                  Digite a senha de administração para carregar ou remover dados:
                </p>
                <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
                  <input
                    type="password"
                    placeholder="Senha"
                    value={senhaInput}
                    onChange={(e) => { setSenhaInput(e.target.value); setErroSenha(false); }}
                    onKeyDown={(e) => e.key === "Enter" && handleAutenticar()}
                    style={{ border: `1px solid ${erroSenha ? "#A32D2D" : BORDER}`, borderRadius: 8, padding: "8px 12px", fontSize: 14, fontFamily: "inherit", width: 220 }}
                  />
                  <button
                    onClick={handleAutenticar}
                    style={{ padding: "8px 18px", borderRadius: 8, border: "none", background: COPPER, color: "#fff", fontWeight: 700, cursor: "pointer", fontFamily: "inherit", fontSize: 13.5 }}
                  >
                    Entrar
                  </button>
                </div>
                {erroSenha && <p style={{ color: "#A32D2D", fontSize: 12.5, marginTop: 8 }}>Senha incorreta.</p>}
              </div>
            ) : (
              <div>
                <div style={{ display: "flex", gap: 32, flexWrap: "wrap" }}>
                  {/* Upload */}
                  <div style={{ flex: 1, minWidth: 280 }}>
                    <p style={{ fontSize: 12, fontWeight: 700, color: MUTED, marginBottom: 8, textTransform: "uppercase", letterSpacing: "0.05em" }}>
                      Carregar planilha
                    </p>
                    <label style={{ fontSize: 13, color: MUTED, display: "block", marginBottom: 6 }}>
                      Arquivo .xlsx (um ou vários meses)
                    </label>
                    <input
                      ref={fileInputRef}
                      type="file"
                      accept=".xlsx,.xls"
                      disabled={processing}
                      onChange={(e) => handleFile(e.target.files[0])}
                      style={{ fontSize: 13.5 }}
                    />
                    {processing && <p style={{ marginTop: 8, fontSize: 13, color: MUTED }}>Processando planilha...</p>}
                  </div>

                  {/* Exclusão por mês */}
                  <div style={{ flex: 1, minWidth: 220 }}>
                    <p style={{ fontSize: 12, fontWeight: 700, color: MUTED, marginBottom: 8, textTransform: "uppercase", letterSpacing: "0.05em" }}>
                      Remover mês
                    </p>
                    {mesesDisponiveis.length === 0 ? (
                      <p style={{ fontSize: 13, color: MUTED }}>Nenhum mês disponível.</p>
                    ) : (
                      <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                        <select
                          id="sel-mes"
                          style={{ border: `1px solid ${BORDER}`, borderRadius: 8, padding: "7px 10px", fontSize: 13.5, fontFamily: "inherit", flex: 1 }}
                        >
                          {mesesDisponiveis.map(([val, label]) => (
                            <option key={val} value={val}>{label}</option>
                          ))}
                        </select>
                        <button
                          disabled={excluindoPeriodo}
                          onClick={() => {
                            const sel = document.getElementById("sel-mes");
                            setConfirmacaoExclusao({ tipo: "mes", valor: sel.value, label: `mês ${monthLabel(sel.value)}` });
                          }}
                          style={{ padding: "7px 14px", borderRadius: 8, border: `1px solid #F5A9A9`, background: "#FDECEA", color: "#A32D2D", fontWeight: 700, cursor: "pointer", fontFamily: "inherit", fontSize: 13 }}
                        >
                          Remover
                        </button>
                      </div>
                    )}
                  </div>

                  {/* Exclusão por ano */}
                  <div style={{ flex: 1, minWidth: 220 }}>
                    <p style={{ fontSize: 12, fontWeight: 700, color: MUTED, marginBottom: 8, textTransform: "uppercase", letterSpacing: "0.05em" }}>
                      Remover ano completo
                    </p>
                    {anosDisponiveis.length === 0 ? (
                      <p style={{ fontSize: 13, color: MUTED }}>Nenhum ano disponível.</p>
                    ) : (
                      <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                        <select
                          id="sel-ano-del"
                          style={{ border: `1px solid ${BORDER}`, borderRadius: 8, padding: "7px 10px", fontSize: 13.5, fontFamily: "inherit", flex: 1 }}
                        >
                          {anosDisponiveis.map((a) => (
                            <option key={a} value={a}>{a}</option>
                          ))}
                        </select>
                        <button
                          disabled={excluindoPeriodo}
                          onClick={() => {
                            const sel = document.getElementById("sel-ano-del");
                            setConfirmacaoExclusao({ tipo: "ano", valor: sel.value, label: `ano ${sel.value} completo` });
                          }}
                          style={{ padding: "7px 14px", borderRadius: 8, border: `1px solid #F5A9A9`, background: "#FDECEA", color: "#A32D2D", fontWeight: 700, cursor: "pointer", fontFamily: "inherit", fontSize: 13 }}
                        >
                          Remover
                        </button>
                      </div>
                    )}
                  </div>
                </div>

                {/* Exclusão por UC (visível se uma UC estiver selecionada) */}
                {selectedUC && (
                  <div style={{ marginTop: 16, paddingTop: 16, borderTop: `1px solid ${BORDER}` }}>
                    <p style={{ fontSize: 12, fontWeight: 700, color: MUTED, marginBottom: 8, textTransform: "uppercase", letterSpacing: "0.05em" }}>
                      Remover UC selecionada
                    </p>
                    <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                      <span style={{ fontSize: 13, color: INK }}>
                        <strong>{ucList.find((u) => u.uc === selectedUC)?.municipio}</strong>
                        {" — "}UC {selectedUC}
                      </span>
                      <button
                        disabled={excluindoPeriodo}
                        onClick={() => setConfirmacaoExclusao({ tipo: "uc", valor: selectedUC, label: `UC ${selectedUC} (todos os meses)` })}
                        style={{ padding: "6px 14px", borderRadius: 8, border: `1px solid #F5A9A9`, background: "#FDECEA", color: "#A32D2D", fontWeight: 700, cursor: "pointer", fontFamily: "inherit", fontSize: 13 }}
                      >
                        Remover todos os dados desta UC
                      </button>
                    </div>
                  </div>
                )}

                {error && <p style={{ marginTop: 12, fontSize: 13, color: "#A32D2D", fontWeight: 600 }}>{error}</p>}
              </div>
            )}
          </section>
        )}

        {/* ── Modal de confirmação de exclusão ── */}
        {confirmacaoExclusao && (
          <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", zIndex: 100, display: "flex", alignItems: "center", justifyContent: "center" }}>
            <div style={{ background: CARD, borderRadius: 14, padding: "28px 32px", maxWidth: 420, width: "90%", boxShadow: "0 8px 32px rgba(0,0,0,0.18)" }}>
              <h3 style={{ fontSize: 16, fontWeight: 700, marginBottom: 12, color: "#A32D2D" }}>Confirmar exclusão</h3>
              <p style={{ fontSize: 14, color: INK, marginBottom: 20, lineHeight: 1.6 }}>
                Você está prestes a remover permanentemente todos os registros do <strong>{confirmacaoExclusao.label}</strong> do banco de dados. Essa ação não pode ser desfeita.
              </p>
              <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
                <button
                  onClick={() => setConfirmacaoExclusao(null)}
                  style={{ padding: "8px 18px", borderRadius: 8, border: `1px solid ${BORDER}`, background: CARD, cursor: "pointer", fontFamily: "inherit", fontSize: 13.5, fontWeight: 600 }}
                >
                  Cancelar
                </button>
                <button
                  disabled={excluindoPeriodo}
                  onClick={() => handleExcluirPeriodo(confirmacaoExclusao)}
                  style={{ padding: "8px 18px", borderRadius: 8, border: "none", background: "#A32D2D", color: "#fff", fontWeight: 700, cursor: "pointer", fontFamily: "inherit", fontSize: 13.5 }}
                >
                  {excluindoPeriodo ? "Removendo..." : "Confirmar exclusão"}
                </button>
              </div>
            </div>
          </div>
        )}

        {/* ── Filtro por ano / período + Resumo ── */}
        {anosDisponiveis.length > 0 && (
          <div style={{ marginBottom: 20 }}>
            {/* Botões de ano */}
            <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 10, flexWrap: "wrap" }}>
              <span style={{ fontSize: 12, fontWeight: 700, color: MUTED, textTransform: "uppercase", letterSpacing: "0.05em" }}>Ano:</span>
              {["todos", ...anosDisponiveis].map((a) => (
                <button
                  key={a}
                  onClick={() => { setAnoSelecionado(a); setUsarFiltroPeriodo(false); setSelectedUC(null); }}
                  style={{
                    padding: "5px 14px", borderRadius: 20,
                    border: `1px solid ${!usarFiltroPeriodo && anoSelecionado === a ? COPPER : BORDER}`,
                    background: !usarFiltroPeriodo && anoSelecionado === a ? "#FAEEDA" : CARD,
                    color: !usarFiltroPeriodo && anoSelecionado === a ? COPPER_DARK : INK,
                    fontWeight: !usarFiltroPeriodo && anoSelecionado === a ? 700 : 400,
                    cursor: "pointer", fontSize: 13.5, fontFamily: "inherit",
                  }}
                >
                  {a === "todos" ? "Todos" : a}
                </button>
              ))}
            </div>

            {/* Filtro por período */}
            <div style={{ display: "flex", gap: 10, alignItems: "center", marginBottom: 14, flexWrap: "wrap" }}>
              <span style={{ fontSize: 12, fontWeight: 700, color: MUTED, textTransform: "uppercase", letterSpacing: "0.05em" }}>Período:</span>
              <input
                type="text"
                placeholder="De MM/AAAA"
                value={periodoInicio}
                onChange={(e) => setPeriodoInicio(e.target.value)}
                style={{ border: `1px solid ${BORDER}`, borderRadius: 8, padding: "5px 10px", fontSize: 13.5, width: 120, fontFamily: "inherit" }}
              />
              <span style={{ color: MUTED, fontSize: 13 }}>até</span>
              <input
                type="text"
                placeholder="MM/AAAA"
                value={periodoFim}
                onChange={(e) => setPeriodoFim(e.target.value)}
                style={{ border: `1px solid ${BORDER}`, borderRadius: 8, padding: "5px 10px", fontSize: 13.5, width: 120, fontFamily: "inherit" }}
              />
              <button
                onClick={() => {
                  if (periodoInicio || periodoFim) { setUsarFiltroPeriodo(true); setSelectedUC(null); }
                }}
                style={{ padding: "5px 14px", borderRadius: 8, border: `1px solid ${usarFiltroPeriodo ? COPPER : BORDER}`, background: usarFiltroPeriodo ? "#FAEEDA" : CARD, color: usarFiltroPeriodo ? COPPER_DARK : INK, fontWeight: 600, cursor: "pointer", fontSize: 13, fontFamily: "inherit" }}
              >
                Filtrar
              </button>
              {usarFiltroPeriodo && (
                <button
                  onClick={() => { setUsarFiltroPeriodo(false); setPeriodoInicio(""); setPeriodoFim(""); }}
                  style={{ padding: "5px 12px", borderRadius: 8, border: `1px solid ${BORDER}`, background: CARD, color: MUTED, cursor: "pointer", fontSize: 13, fontFamily: "inherit" }}
                >
                  Limpar
                </button>
              )}
            </div>

            {/* Resumo — 4 + 4 métricas simétricas */}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12, marginBottom: 12 }}>
              <Metric label={`Total faturado${anoSelecionado !== "todos" && !usarFiltroPeriodo ? ` ${anoSelecionado}` : ""}`} value={fmtCurrency(resumoAno.totalFaturado)} />
              <Metric label="Consumo total (kWh)" value={fmtNumber(resumoAno.totalConsumo, 0)} />
              <Metric label="Crédito geração distribuída" value={fmtCurrency(resumoAno.totalGD)} />
              <Metric label="CIP total" value={fmtCurrency(resumoAno.totalCip)} />
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12 }}>
              <Metric label="ICMS total" value={fmtCurrency(resumoAno.totalICMS)} />
              <Metric label="COFINS total" value={fmtCurrency(resumoAno.totalCOFINS)} />
              <Metric label="PIS total" value={fmtCurrency(resumoAno.totalPIS)} />
              <Metric label="Unidades no período" value={resumoAno.ucs} />
            </div>
          </div>
        )}

        {loading ? (
          <p style={{ color: MUTED, fontSize: 14 }}>Carregando dados...</p>
        ) : ucList.length === 0 ? (
          <div style={{ background: CARD, border: `1px dashed ${BORDER}`, borderRadius: 12, padding: "48px 24px", textAlign: "center", color: MUTED }}>
            Nenhuma unidade consumidora registrada ainda.<br />
            Envie a planilha Excel acima para começar.
          </div>
        ) : (
          <div style={{ display: "grid", gridTemplateColumns: "320px 1fr", gap: 20 }}>
            <div style={{ background: CARD, border: `1px solid ${BORDER}`, borderRadius: 12, padding: 14, height: "fit-content", maxHeight: 640, display: "flex", flexDirection: "column" }}>
              <input
                type="text"
                placeholder="Buscar município, UC ou unidade..."
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                style={{ border: `1px solid ${BORDER}`, borderRadius: 8, padding: "8px 10px", fontSize: 13.5, marginBottom: 10, fontFamily: "inherit" }}
              />
              <div style={{ overflowY: "auto", flex: 1 }}>
                {filteredUcList.map((u) => (
                  <button
                    key={u.uc}
                    onClick={() => setSelectedUC(u.uc)}
                    style={{
                      display: "block", width: "100%", textAlign: "left", padding: "10px 10px",
                      marginBottom: 4, borderRadius: 8, border: "none", cursor: "pointer",
                      fontFamily: "inherit", background: selectedUC === u.uc ? "#FAEEDA" : "transparent",
                    }}
                  >
                    <div style={{ fontSize: 13.5, fontWeight: 600, color: selectedUC === u.uc ? COPPER_DARK : INK, display: "flex", alignItems: "center", gap: 6 }}>
                      <span>{u.municipio}</span>
                      {ucsComGD.has(u.uc) && (
                        <span title="Possui geração distribuída" style={{ fontSize: 10, fontWeight: 700, color: "#27500A", background: "#EAF3DE", borderRadius: 4, padding: "1px 5px", letterSpacing: "0.02em" }}>
                          GD
                        </span>
                      )}
                    </div>
                    <div style={{ fontSize: 12, color: MUTED, marginTop: 1 }}>{u.complemento}</div>
                    <div style={{ fontSize: 11, color: MUTED, marginTop: 2, fontFamily: "ui-monospace, monospace" }}>UC {u.uc}</div>
                  </button>
                ))}
                {filteredUcList.length === 0 && <p style={{ color: MUTED, fontSize: 13, padding: 10 }}>Nenhuma unidade encontrada.</p>}
              </div>
            </div>

            <div>
              {!selectedUC ? (
                <div style={{ background: CARD, border: `1px dashed ${BORDER}`, borderRadius: 12, padding: "60px 24px", textAlign: "center", color: MUTED, fontSize: 14 }}>
                  Selecione uma unidade consumidora na lista ao lado para ver o histórico.
                </div>
              ) : (
                <>
                  <div style={{ background: CARD, border: `1px solid ${BORDER}`, borderRadius: 12, padding: "16px 20px", marginBottom: 16 }}>
                    <h2 style={{ fontSize: 16, fontWeight: 700, margin: 0 }}>
                      {ultimoRegistro?.municipio} <span style={{ color: MUTED, fontWeight: 400 }}>— {ultimoRegistro?.complemento}</span>
                    </h2>
                    <p style={{ fontSize: 12, color: MUTED, marginTop: 2, fontFamily: "ui-monospace, monospace" }}>UC {selectedUC}</p>

                    <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12, marginTop: 16 }}>
                      <Metric label="Último consumo" value={`${fmtNumber(ultimoRegistro?.consumo_total_kwh, 0)} kWh`} />
                      <Metric label="Último valor total" value={fmtCurrency(ultimoRegistro?.valor_total)} />
                      <Metric label="Meses registrados" value={historicoUC.length} />
                      <Metric label="Categoria" value={ultimoRegistro?.categoria === "ALTA TENSAO" ? "Alta tensão" : "Baixa tensão"} />
                    </div>

                    {isAltaTensao && (
                      <>
                        <div style={{ marginTop: 10, marginBottom: 4, fontSize: 11, fontWeight: 700, color: MUTED, letterSpacing: "0.06em", textTransform: "uppercase" }}>
                          Demanda
                        </div>
                        <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12 }}>
                          <Metric
                            label="Contratada"
                            value={`${fmtNumber(ultimoRegistro?.demanda_contratada_kw, 0)} kW`}
                          />
                          <Metric
                            label="Ativa medida"
                            value={`${fmtNumber(ultimoRegistro?.demanda_ativa_kw || ultimoRegistro?.demanda_contratada_kw, 0)} kW`}
                            alert={
                              (ultimoRegistro?.demanda_ativa_kw || 0) > (ultimoRegistro?.demanda_contratada_kw || 0)
                                ? "red"
                                : null
                            }
                          />
                          <Metric
                            label="Ponta"
                            value={`${fmtNumber(ultimoRegistro?.demanda_ponta_kw || ultimoRegistro?.demanda_ponta_col_kw, 0)} kW`}
                          />
                          <Metric
                            label="Fora ponta"
                            value={`${fmtNumber(ultimoRegistro?.demanda_fora_ponta_kw || ultimoRegistro?.demanda_fora_ponta_col_kw, 0)} kW`}
                          />
                        </div>

                        <div style={{ marginTop: 12, marginBottom: 4, fontSize: 11, fontWeight: 700, color: MUTED, letterSpacing: "0.06em", textTransform: "uppercase" }}>
                          Consumo
                        </div>
                        <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12 }}>
                          <Metric label="Total" value={`${fmtNumber(ultimoRegistro?.consumo_total_kwh, 0)} kWh`} />
                          <Metric label="Ponta" value={`${fmtNumber(ultimoRegistro?.consumo_ponta_kwh, 0)} kWh`} />
                          <Metric label="Fora ponta" value={`${fmtNumber(ultimoRegistro?.consumo_fora_ponta_kwh, 0)} kWh`} />
                          <Metric
                            label="Reativo excedente"
                            value={`${fmtNumber((ultimoRegistro?.reativo_excedente_fp_kwh || 0) + (ultimoRegistro?.reativo_excedente_np_kwh || 0), 0)} kWh`}
                            alert={
                              ((ultimoRegistro?.val_reativo_excedente || 0) + (ultimoRegistro?.val_reativo_excedente_fp || 0) + (ultimoRegistro?.val_reativo_excedente_np || 0)) > 0
                                ? "orange"
                                : null
                            }
                          />
                        </div>
                      </>
                    )}
                  </div>

                  {temAlertas && (
                    <div style={{ background: "#FFF8F0", border: "1px solid #F5C58A", borderRadius: 12, padding: "16px 20px", marginBottom: 16 }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12 }}>
                        <span style={{ fontSize: 16 }}>⚠️</span>
                        <h3 style={{ fontSize: 13, fontWeight: 700, color: "#7A3C00", margin: 0 }}>
                          Alertas desta unidade
                        </h3>
                      </div>

                      {alertasUltrapassagem.length > 0 && (
                        <div style={{ marginBottom: 12 }}>
                          <div style={{ fontSize: 12, fontWeight: 700, color: "#A32D2D", marginBottom: 6, display: "flex", alignItems: "center", gap: 6 }}>
                            <span style={{ width: 8, height: 8, borderRadius: "50%", background: "#A32D2D", display: "inline-block" }} />
                            Ultrapassagem de demanda ({alertasUltrapassagem.length} {alertasUltrapassagem.length === 1 ? "ocorrência" : "ocorrências"})
                          </div>
                          {alertasUltrapassagem.map((r) => {
                            const isHoro = r.demanda_contratada_ponta_kw > 0 || r.demanda_contratada_fora_ponta_kw > 0;
                            const contratada = r.demanda_contratada_kw || 0;
                            const medida = r.demanda_ativa_kw || 0;
                            const ultrKw = r.demanda_ultrapassagem_kw || r.demanda_ultrapassagem_ponta_kw || 0;
                            const ultrVal = r.val_ultrapassagem_total || r.val_demanda_ultrapassagem || 0;
                            return (
                              <div key={r.fatura} style={{ fontSize: 12.5, color: INK, background: "#FDECEA", borderRadius: 6, padding: "8px 10px", marginBottom: 4 }}>
                                <strong>{monthLabel(r.mes_referencia)}</strong>
                                {isHoro ? (
                                  <span> — tarifa horosazonal · contratada ponta {fmtNumber(r.demanda_contratada_ponta_kw, 0)} kW / fora ponta {fmtNumber(r.demanda_contratada_fora_ponta_kw, 0)} kW</span>
                                ) : contratada > 0 ? (
                                  <span> — medida {fmtNumber(medida, 1)} kW vs contratada {fmtNumber(contratada, 0)} kW
                                    {" "}(<span style={{ color: "#A32D2D", fontWeight: 700 }}>+{(((medida - contratada) / contratada) * 100).toFixed(1)}%</span>)
                                  </span>
                                ) : null}
                                {ultrKw > 0 && (
                                  <span> · excesso <strong>{fmtNumber(ultrKw, 2)} kW</strong></span>
                                )}
                                {ultrVal > 0 && (
                                  <span> · custo <strong style={{ color: "#A32D2D" }}>{fmtCurrency(ultrVal)}</strong></span>
                                )}
                              </div>
                            );
                          })}
                        </div>
                      )}

                      {alertaConsumo && (
                        <div style={{ marginBottom: 12 }}>
                          <div style={{ fontSize: 12, fontWeight: 700, color: "#7A3C00", marginBottom: 6, display: "flex", alignItems: "center", gap: 6 }}>
                            <span style={{ width: 8, height: 8, borderRadius: "50%", background: "#E88C00", display: "inline-block" }} />
                            Variação de consumo acima de 20%
                          </div>
                          <div style={{ fontSize: 12.5, color: INK, background: "#FFF3DC", borderRadius: 6, padding: "6px 10px" }}>
                            <strong>{monthLabel(alertaConsumo.mes)}</strong> — consumo {fmtNumber(alertaConsumo.consumoAtual, 0)} kWh vs média dos 6 meses anteriores {fmtNumber(alertaConsumo.media, 0)} kWh
                            {" "}(
                            <span style={{ color: alertaConsumo.variacao > 0 ? "#A32D2D" : "#3B6D11", fontWeight: 700 }}>
                              {alertaConsumo.variacao > 0 ? "+" : ""}{(alertaConsumo.variacao * 100).toFixed(1)}%
                            </span>
                            )
                          </div>
                        </div>
                      )}

                      {alertasReativo.length > 0 && (
                        <div>
                          <div style={{ fontSize: 12, fontWeight: 700, color: "#5A4200", marginBottom: 6, display: "flex", alignItems: "center", gap: 6 }}>
                            <span style={{ width: 8, height: 8, borderRadius: "50%", background: "#C47F00", display: "inline-block" }} />
                            Reativo excedente faturado ({alertasReativo.length} {alertasReativo.length === 1 ? "ocorrência" : "ocorrências"})
                          </div>
                          {alertasReativo.map((r) => {
                            const valReativo = (r.val_reativo_excedente || 0) + (r.val_reativo_excedente_fp || 0) + (r.val_reativo_excedente_np || 0);
                            const kvarReativo = (r.reativo_excedente_kwh || 0) + (r.reativo_excedente_fp_kwh || 0) + (r.reativo_excedente_np_kwh || 0);
                            return (
                              <div key={r.fatura} style={{ fontSize: 12.5, color: INK, background: "#FFF8DC", borderRadius: 6, padding: "6px 10px", marginBottom: 4 }}>
                                <strong>{monthLabel(r.mes_referencia)}</strong> — {fmtNumber(kvarReativo, 1)} kVAr excedente · custo {fmtCurrency(valReativo)}
                              </div>
                            );
                          })}

                          {recomendacaoCapacitor && (
                            <div style={{ marginTop: 12, background: "#F0F7FF", border: "1px solid #AED6F1", borderRadius: 8, padding: "14px 16px" }}>
                              <div style={{ fontSize: 12, fontWeight: 700, color: "#1A5276", marginBottom: 10, display: "flex", alignItems: "center", gap: 6 }}>
                                ⚡ Recomendação técnica — Instalação de banco de capacitores
                              </div>
                              <p style={{ fontSize: 12.5, color: INK, margin: "0 0 10px", lineHeight: 1.6 }}>
                                Esta unidade registrou <strong>{recomendacaoCapacitor.ocorrencias} meses</strong> com reativo excedente
                                faturado no período analisado, superando o limite de 4 ocorrências que justifica a instalação de
                                correção de fator de potência conforme a <strong>NBR 5410</strong> e resolução normativa <strong>ANEEL nº 1.000/2021</strong>
                                (FP mínimo: 0,92 indutivo).
                              </p>

                              <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 10, marginBottom: 12 }}>
                                <div style={{ background: CARD, borderRadius: 6, padding: "8px 10px" }}>
                                  <div style={{ fontSize: 10, color: MUTED, fontWeight: 700, marginBottom: 2 }}>FP ATUAL (pior mês)</div>
                                  <div style={{ fontSize: 16, fontWeight: 700, color: "#A32D2D" }}>{recomendacaoCapacitor.fpAtual}</div>
                                  <div style={{ fontSize: 10, color: MUTED }}>referência: {monthLabel(recomendacaoCapacitor.mesPior)}</div>
                                </div>
                                <div style={{ background: CARD, borderRadius: 6, padding: "8px 10px" }}>
                                  <div style={{ fontSize: 10, color: MUTED, fontWeight: 700, marginBottom: 2 }}>CARGA ATIVA</div>
                                  <div style={{ fontSize: 16, fontWeight: 700 }}>{recomendacaoCapacitor.kW} kW</div>
                                </div>
                                <div style={{ background: CARD, borderRadius: 6, padding: "8px 10px" }}>
                                  <div style={{ fontSize: 10, color: MUTED, fontWeight: 700, marginBottom: 2 }}>REATIVO TOTAL ESTIMADO</div>
                                  <div style={{ fontSize: 16, fontWeight: 700 }}>{recomendacaoCapacitor.kVArTotal} kVAr</div>
                                </div>
                              </div>

                              <div style={{ background: "#1A5276", borderRadius: 8, padding: "12px 14px", color: "#fff" }}>
                                <div style={{ fontSize: 12, fontWeight: 700, marginBottom: 8, letterSpacing: "0.04em" }}>
                                  BANCO DE CAPACITORES RECOMENDADO
                                </div>
                                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10 }}>
                                  <div>
                                    <div style={{ fontSize: 10, opacity: 0.75, marginBottom: 2 }}>Potência reativa calculada</div>
                                    <div style={{ fontSize: 15, fontWeight: 700 }}>{recomendacaoCapacitor.Qc_kVAr} kVAr</div>
                                  </div>
                                  <div>
                                    <div style={{ fontSize: 10, opacity: 0.75, marginBottom: 2 }}>Potência padronizada (comercial)</div>
                                    <div style={{ fontSize: 15, fontWeight: 700 }}>{recomendacaoCapacitor.Qc_padronizado} kVAr</div>
                                  </div>
                                  <div>
                                    <div style={{ fontSize: 10, opacity: 0.75, marginBottom: 2 }}>Tensão nominal</div>
                                    <div style={{ fontSize: 15, fontWeight: 700 }}>13,8 kV</div>
                                  </div>
                                </div>
                                <div style={{ marginTop: 10, paddingTop: 10, borderTop: "1px solid rgba(255,255,255,0.2)" }}>
                                  <div style={{ fontSize: 10, opacity: 0.75, marginBottom: 4 }}>
                                    Capacitância por fase — banco trifásico em triângulo (Δ), 60 Hz
                                  </div>
                                  <div style={{ fontSize: 14, fontWeight: 700 }}>
                                    C = {recomendacaoCapacitor.C_uF} µF / fase
                                  </div>
                                  <div style={{ fontSize: 10, opacity: 0.65, marginTop: 6 }}>
                                    Fórmula: C = Q / (3 × 2πf × V²) · Alvo de correção: FP ≥ 0,95
                                    · Base: ANEEL Res. 1.000/2021 + NBR 5410
                                  </div>
                                </div>
                              </div>
                            </div>
                          )}

                          {alertasReativo.length > 0 && alertasReativo.length <= 4 && (
                            <p style={{ fontSize: 12, color: "#5A4200", marginTop: 8, marginBottom: 0 }}>
                              Atenção: {alertasReativo.length} {alertasReativo.length === 1 ? "ocorrência" : "ocorrências"} de reativo excedente no período.
                              A recomendação de instalação de banco de capacitores é emitida a partir de 5 ocorrências.
                            </p>
                          )}
                        </div>
                      )}

                      {alertasCIP.length > 0 && (
                        <div style={{ marginTop: alertasReativo.length > 0 ? 12 : 0 }}>
                          <div style={{ fontSize: 12, fontWeight: 700, color: "#185FA5", marginBottom: 6, display: "flex", alignItems: "center", gap: 6 }}>
                            <span style={{ width: 8, height: 8, borderRadius: "50%", background: "#185FA5", display: "inline-block" }} />
                            Aumento da CIP (Contribuição de Iluminação Pública)
                          </div>
                          {alertasCIP.map((a) => (
                            <div key={a.fatura} style={{ fontSize: 12.5, color: INK, background: "#EAF2FB", borderRadius: 6, padding: "6px 10px", marginBottom: 4 }}>
                              <strong>{monthLabel(a.mes)}</strong> — CIP {fmtCurrency(a.cipAtual)} vs {fmtCurrency(a.cipAnterior)} no mês anterior
                              {" "}(<span style={{ color: "#185FA5", fontWeight: 700 }}>+{fmtCurrency(a.aumento)} · +{a.pct.toFixed(1)}%</span>)
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  )}

                  {temGeracaoDistribuida && (
                    <div style={{ background: CARD, border: `1px solid ${BORDER}`, borderRadius: 12, padding: "16px 20px", marginBottom: 16 }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12 }}>
                        <span style={{ width: 8, height: 8, borderRadius: "50%", background: "#3B6D11" }} />
                        <h3 style={{ fontSize: 13, fontWeight: 700, color: "#3B6D11", margin: 0 }}>Geração distribuída (energia solar)</h3>
                      </div>
                      <p style={{ fontSize: 12.5, color: MUTED, margin: "0 0 14px" }}>
                        Esta unidade possui sistema de geração própria conectado à rede, com crédito de
                        energia injetada abatido na fatura ("Dev Geração"). Os valores abaixo somam o
                        histórico carregado.
                      </p>
                      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 12, marginBottom: 16 }}>
                        <Metric label="Crédito total acumulado" value={fmtCurrency(totalCreditoGD)} />
                        <Metric label="Último crédito mensal" value={fmtCurrency(Math.abs(ultimoRegistro?.dev_geracao_total || 0))} />
                        <Metric label="Meses com geração" value={historicoUC.filter((r) => r.tem_gd).length} />
                      </div>
                      {chartDataGD.length > 1 && (
                        <ResponsiveContainer width="100%" height={160}>
                          <LineChart data={chartDataGD}>
                            <CartesianGrid strokeDasharray="3 3" stroke={BORDER} />
                            <XAxis dataKey="mes" tick={{ fontSize: 11, fill: MUTED }} />
                            <YAxis tick={{ fontSize: 11, fill: MUTED }} width={70} />
                            <Tooltip formatter={(v) => fmtCurrency(v)} />
                            <Line type="monotone" dataKey="credito" name="Crédito (R$)" stroke="#3B6D11" strokeWidth={2} dot={{ r: 3 }} />
                          </LineChart>
                        </ResponsiveContainer>
                      )}
                      {mesesComGDeNaoIdentificado && (
                        <p style={{ fontSize: 11.5, color: "#854F0B", marginTop: 12, marginBottom: 0, lineHeight: 1.5 }}>
                          Atenção: em pelo menos um mês desta unidade há valores na planilha original sem
                          a coluna "DETALHAMENTO_CONTA" preenchida (provavelmente componentes adicionais
                          de geração distribuída cujo rótulo se perdeu na exportação). Esses valores não
                          estão somados nas métricas acima — consulte a fatura original para conferência exata.
                        </p>
                      )}
                    </div>
                  )}

                  {chartData.length > 1 && (
                    <div style={{ background: CARD, border: `1px solid ${BORDER}`, borderRadius: 12, padding: "16px 20px", marginBottom: 16 }}>
                      <h3 style={{ fontSize: 13, fontWeight: 700, color: MUTED, margin: "0 0 12px" }}>Evolução do consumo e valor faturado</h3>
                      <ResponsiveContainer width="100%" height={220}>
                        <LineChart data={chartData}>
                          <CartesianGrid strokeDasharray="3 3" stroke={BORDER} />
                          <XAxis dataKey="mes" tick={{ fontSize: 12, fill: MUTED }} />
                          <YAxis yAxisId="left" tick={{ fontSize: 11, fill: MUTED }} width={60} />
                          <YAxis yAxisId="right" orientation="right" tick={{ fontSize: 11, fill: MUTED }} width={70} />
                          <Tooltip formatter={(v, name) => (name === "Consumo (kWh)" ? `${fmtNumber(v, 0)} kWh` : fmtCurrency(v))} />
                          <Line yAxisId="left" type="monotone" dataKey="consumo" name="Consumo (kWh)" stroke={COPPER} strokeWidth={2} dot={{ r: 3 }} />
                          <Line yAxisId="right" type="monotone" dataKey="valor" name="Valor total (R$)" stroke="#185FA5" strokeWidth={2} dot={{ r: 3 }} strokeDasharray="4 3" />
                        </LineChart>
                      </ResponsiveContainer>
                      <div style={{ display: "flex", gap: 16, fontSize: 12, color: MUTED, marginTop: 4 }}>
                        <LegendDot color={COPPER} label="Consumo (kWh)" />
                        <LegendDot color="#185FA5" label="Valor total (R$)" dashed />
                      </div>
                    </div>
                  )}

                  <div style={{ background: CARD, border: `1px solid ${BORDER}`, borderRadius: 12, padding: "16px 20px" }}>
                    <h3 style={{ fontSize: 13, fontWeight: 700, color: MUTED, margin: "0 0 12px" }}>Histórico mensal detalhado</h3>
                    <div style={{ overflowX: "auto" }}>
                      <table style={{ width: "100%", fontSize: 13, borderCollapse: "collapse" }}>
                        <thead>
                          <tr style={{ borderBottom: `1px solid ${BORDER}` }}>
                            {(isAltaTensao
                              ? ["Mês", "Consumo total", "Demanda (kW)", "ICMS", "COFINS", "PIS", "CIP", "Total fatura"]
                              : ["Mês", "Consumo (kWh)", "Valor consumo", "ICMS", "COFINS", "PIS", "CIP", "Total fatura"]
                            ).map((h) => (
                              <th key={h} style={{ textAlign: "right", padding: "6px 8px", color: MUTED, fontWeight: 600, fontSize: 12 }}>
                                {h === "Mês" ? <span style={{ float: "left" }}>{h}</span> : h}
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {historicoUC.map((r) => {
                            const temUltrapassagem = alertasUltrapassagem.some((a) => a.fatura === r.fatura);
                            const temReativo = alertasReativo.some((a) => a.fatura === r.fatura);
                            const ehMesAlertaConsumo = alertaConsumo?.mes === r.mes_referencia;
                            const temCIP = alertasCIP.some((a) => a.fatura === r.fatura);
                            const rowAlert = temUltrapassagem ? "#FDECEA" : temReativo ? "#FFF8DC" : ehMesAlertaConsumo ? "#FFF3DC" : temCIP ? "#EAF2FB" : null;
                            return (
                              <tr key={r.fatura} style={{ borderBottom: `1px solid ${BORDER}`, background: rowAlert || "transparent" }}>
                                <td style={{ padding: "8px 8px", fontWeight: 600 }}>
                                  {monthLabel(r.mes_referencia)}
                                  {temUltrapassagem && <span title="Ultrapassagem de demanda" style={{ marginLeft: 5 }}>🔴</span>}
                                  {temReativo && <span title="Reativo excedente" style={{ marginLeft: 5 }}>🟠</span>}
                                  {ehMesAlertaConsumo && <span title="Variação de consumo > 20%" style={{ marginLeft: 5 }}>🟡</span>}
                                  {temCIP && <span title="Aumento de CIP" style={{ marginLeft: 5 }}>🔵</span>}
                                </td>
                                <td style={{ padding: "8px 8px", textAlign: "right" }}>
                                  {isAltaTensao ? fmtNumber(r.consumo_total_kwh, 0) : fmtNumber(r.kwh, 0)}
                                </td>
                                <td style={{ padding: "8px 8px", textAlign: "right" }}>
                                  {isAltaTensao ? fmtNumber(r.demanda_contratada_kw, 0) : fmtCurrency(r.val_consumo)}
                                </td>
                                <td style={{ padding: "8px 8px", textAlign: "right" }}>{fmtCurrency(r.icms_val)}</td>
                                <td style={{ padding: "8px 8px", textAlign: "right" }}>{fmtCurrency(r.cofins_val)}</td>
                                <td style={{ padding: "8px 8px", textAlign: "right" }}>{fmtCurrency(r.pis_val)}</td>
                                <td style={{ padding: "8px 8px", textAlign: "right" }}>{fmtCurrency(r.cip)}</td>
                                <td style={{ padding: "8px 8px", textAlign: "right", fontWeight: 700, color: COPPER_DARK }}>{fmtCurrency(r.valor_total)}</td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </>
              )}
            </div>
          </div>
        )}
      </div>
      <style>{`
        button:hover { background: #F1EFE8 !important; }
        input[type="text"]:focus { outline: 2px solid ${COPPER}; outline-offset: 1px; }
      `}</style>
    </div>
  );
}

function Metric({ label, value, alert }) {
  const bgMap = { red: "#FDECEA", orange: "#FFF3DC", yellow: "#FFFBDC" };
  const borderMap = { red: "1px solid #F5A9A9", orange: "1px solid #F5C58A", yellow: "1px solid #F5E58A" };
  const bg = alert ? bgMap[alert] : BG;
  const border = alert ? borderMap[alert] : "none";
  return (
    <div style={{ background: bg, borderRadius: 8, padding: "10px 12px", border }}>
      <div style={{ fontSize: 11, color: MUTED, fontWeight: 600, marginBottom: 3 }}>{label}</div>
      <div style={{ fontSize: 17, fontWeight: 700 }}>{value}</div>
    </div>
  );
}

function LegendDot({ color, label, dashed }) {
  return (
    <span style={{ display: "flex", alignItems: "center", gap: 6 }}>
      <span style={{ width: 14, height: dashed ? 2 : 8, borderRadius: dashed ? 0 : "50%", background: dashed ? "none" : color, borderTop: dashed ? `2px dashed ${color}` : "none" }} />
      {label}
    </span>
  );
}
