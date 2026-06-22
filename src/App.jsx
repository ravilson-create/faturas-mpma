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
  "Demanda Ultrapassagem Ponta": { consumo: "demanda_ultrapassagem_kw", valor: "val_demanda_ultrapassagem" },
  "Demanda de Geração": { consumo: null, valor: "val_demanda_geracao" },
  "Consumo Reativo Excedente": { consumo: "reativo_excedente_kwh", valor: "val_reativo_excedente" },
  "Consumo Reativo Excedente FP": { consumo: "reativo_excedente_fp_kwh", valor: "val_reativo_excedente_fp" },
  "Consumo Reativo Excedente NP": { consumo: "reativo_excedente_np_kwh", valor: "val_reativo_excedente_np" },
  "Adicional Bandeira": { consumo: null, valor: "val_bandeira" },
  "Cip-Ilum Pub Pref Munic": { consumo: null, valor: "cip" },
  "Tributo a Reter IRPJ": { consumo: null, valor: "irpj" },
  "Tributo a Reter CSLL": { consumo: null, valor: "csll" },
  "Tributo a Reter PIS": { consumo: null, valor: "pis_retido" },
  "Tributo a Reter COFINS": { consumo: null, valor: "cofins_retido" },
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
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: true });

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
        reg[mapping.consumo] = (reg[mapping.consumo] || 0) + (Number(consumoCell) || 0);
      }
      if (mapping.valor) {
        reg[mapping.valor] = (reg[mapping.valor] || 0) + valorCell;
      }
    } else if (!itemRaw && valorCell !== 0) {
      reg.nao_identificado = (reg.nao_identificado || 0) + valorCell;
    }

    if (row["DEMANDA"] != null) reg.demanda_contratada_kw = Number(row["DEMANDA"]);
    if (row["DEMANDA_PONTA"] != null) reg.demanda_ponta_col_kw = Number(row["DEMANDA_PONTA"]);
    if (row["DEMANDA_FORA_PONTA"] != null) reg.demanda_fora_ponta_col_kw = Number(row["DEMANDA_FORA_PONTA"]);
  }

  const registros = Array.from(byFatura.values()).map((r) => {
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
  const fileInputRef = useRef(null);

  const carregarRegistros = useCallback(async () => {
    setLoading(true);
    const { data, error: err } = await supabase
      .from("faturas")
      .select("*")
      .order("mes_referencia", { ascending: true });
    if (err) {
      setError("Não foi possível carregar os dados do banco: " + err.message);
    } else {
      setRegistros(data || []);
    }
    setLoading(false);
  }, []);

  useEffect(() => {
    carregarRegistros();
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

      const { error: upsertError } = await supabase
        .from("faturas")
        .upsert(novosRegistros, { onConflict: "fatura" });

      if (upsertError) throw new Error(upsertError.message);

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

  const ucList = useMemo(() => {
    const map = new Map();
    registros.forEach((r) => {
      if (!map.has(r.uc)) {
        map.set(r.uc, { uc: r.uc, municipio: r.municipio, complemento: r.complemento });
      }
    });
    return Array.from(map.values()).sort((a, b) => (a.municipio || "").localeCompare(b.municipio || ""));
  }, [registros]);

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
    return registros
      .filter((r) => r.uc === selectedUC)
      .sort((a, b) => new Date(a.mes_referencia) - new Date(b.mes_referencia));
  }, [registros, selectedUC]);

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
  const totalMesesUnicos = useMemo(() => new Set(registros.map((r) => r.mes_referencia)).size, [registros]);

  // ── Alertas ──────────────────────────────────────────────────────────────

  // Alerta 1: ultrapassagem de demanda (medida > contratada)
  const alertasUltrapassagem = useMemo(() => {
    return historicoUC.filter((r) => {
      const contratada = r.demanda_contratada_kw || 0;
      const medida = r.demanda_ativa_kw || r.demanda_ponta_kw || 0;
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
    if (Math.abs(variacao) > 0.2) {
      return { variacao, media, consumoAtual: ultimo.consumo_total_kwh, mes: ultimo.mes_referencia };
    }
    return null;
  }, [historicoUC]);

  // Alerta 3: reativo excedente (qualquer mês com valor > 0)
  const alertasReativo = useMemo(() => {
    return historicoUC.filter((r) => {
      const reativo = (r.val_reativo_excedente || 0) + (r.val_reativo_excedente_fp || 0) + (r.val_reativo_excedente_np || 0);
      return reativo > 0;
    });
  }, [historicoUC]);

  const temAlertas = alertasUltrapassagem.length > 0 || alertaConsumo !== null || alertasReativo.length > 0;

  return (
    <div style={{ fontFamily: "Inter, system-ui, sans-serif", background: BG, minHeight: "100vh", color: INK }}>
      <div style={{ maxWidth: 1140, margin: "0 auto", padding: "32px 24px 64px" }}>
        <header style={{ marginBottom: 28 }}>
          <h1 style={{ fontSize: 24, fontWeight: 700, margin: 0, letterSpacing: "-0.01em" }}>
            Consulta de faturas <span style={{ color: COPPER_DARK }}>· Energia MPMA</span>
          </h1>
          <p style={{ color: MUTED, fontSize: 13.5, marginTop: 6 }}>
            {ucList.length > 0
              ? `${ucList.length} unidades consumidoras · ${totalMesesUnicos} ${totalMesesUnicos === 1 ? "mês" : "meses"} de histórico${ucsComGD.size > 0 ? ` · ${ucsComGD.size} com geração distribuída` : ""}`
              : "Envie a planilha Excel da Equatorial para começar a registrar o histórico."}
          </p>
        </header>

        <section style={{ background: CARD, border: `1px solid ${BORDER}`, borderRadius: 12, padding: "18px 20px", marginBottom: 24 }}>
          <label style={{ fontSize: 12, color: MUTED, fontWeight: 600, display: "block", marginBottom: 6 }}>
            Planilha de faturamento (.xlsx) — pode conter um ou vários meses
          </label>
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx,.xls"
            disabled={processing}
            onChange={(e) => handleFile(e.target.files[0])}
            style={{ fontSize: 13.5 }}
          />
          {processing && <p style={{ marginTop: 10, fontSize: 13, color: MUTED }}>Lendo planilha e salvando no banco compartilhado...</p>}
          {error && <p style={{ marginTop: 12, fontSize: 13, color: "#A32D2D", fontWeight: 600 }}>{error}</p>}
        </section>

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
                            Ultrapassagem de demanda
                          </div>
                          {alertasUltrapassagem.map((r) => {
                            const contratada = r.demanda_contratada_kw || 0;
                            const medida = r.demanda_ativa_kw || r.demanda_ponta_kw || 0;
                            const excesso = medida - contratada;
                            const pct = contratada > 0 ? ((excesso / contratada) * 100).toFixed(1) : "-";
                            return (
                              <div key={r.fatura} style={{ fontSize: 12.5, color: INK, background: "#FDECEA", borderRadius: 6, padding: "6px 10px", marginBottom: 4 }}>
                                <strong>{monthLabel(r.mes_referencia)}</strong> — medida {fmtNumber(medida, 0)} kW vs contratada {fmtNumber(contratada, 0)} kW
                                {" "}(<span style={{ color: "#A32D2D", fontWeight: 700 }}>+{pct}%</span>)
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
                            Reativo excedente faturado
                          </div>
                          {alertasReativo.map((r) => {
                            const valReativo = (r.val_reativo_excedente || 0) + (r.val_reativo_excedente_fp || 0) + (r.val_reativo_excedente_np || 0);
                            const kwhReativo = (r.reativo_excedente_kwh || 0) + (r.reativo_excedente_fp_kwh || 0) + (r.reativo_excedente_np_kwh || 0);
                            return (
                              <div key={r.fatura} style={{ fontSize: 12.5, color: INK, background: "#FFF8DC", borderRadius: 6, padding: "6px 10px", marginBottom: 4 }}>
                                <strong>{monthLabel(r.mes_referencia)}</strong> — {fmtNumber(kwhReativo, 0)} kVAr excedente · custo {fmtCurrency(valReativo)}
                              </div>
                            );
                          })}
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
                            const rowAlert = temUltrapassagem ? "#FDECEA" : temReativo ? "#FFF8DC" : ehMesAlertaConsumo ? "#FFF3DC" : null;
                            return (
                              <tr key={r.fatura} style={{ borderBottom: `1px solid ${BORDER}`, background: rowAlert || "transparent" }}>
                                <td style={{ padding: "8px 8px", fontWeight: 600 }}>
                                  {monthLabel(r.mes_referencia)}
                                  {temUltrapassagem && <span title="Ultrapassagem de demanda" style={{ marginLeft: 5 }}>🔴</span>}
                                  {temReativo && <span title="Reativo excedente" style={{ marginLeft: 5 }}>🟠</span>}
                                  {ehMesAlertaConsumo && <span title="Variação de consumo > 20%" style={{ marginLeft: 5 }}>🟡</span>}
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
