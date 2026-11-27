// Roda na página de download: recebe os dados (pequenos) pelo link, monta o .xlsx e baixa.
import { gerarXlsx, EntradaXlsx } from "../visual/src/xlsx";
import { gerarXlsxDia, EntradaDia } from "../painel-dia/src/xlsxDia";
import { gerarXlsxMes, EntradaMes, LinhaMes } from "../painel-mes/src/xlsxMes";
import { gerarXlsxPedido, EntradaPedido, LinhaPedido } from "../pedido-diario/src/xlsxPedido";
import { gerarXlsxReps, EntradaReps } from "../reps-sem-pedidos/src/xlsxReps";
import { gerarXlsxMetas, EntradaMetas } from "../painel-metas/src/xlsxMetas";
import { gerarXlsxPedidos, expandir, CargaPedidos } from "../pedidos-mes/src/xlsxPedidos";
import { gerarXlsxPainelVendas, expandir as expandirPV, CargaPV } from "../painel-vendas-aprovacao/src/xlsxPainelVendas";
import { gerarXlsxMetasRep, expandir as expandirMR, CargaRep } from "../painel-metas-rep/src/xlsxMetasRep";
import { ANTON_B64 } from "./fonte";
import { decodificarPedidosTabela, gerarXlsxTabela, mensagemPedidos } from "./pedidosTabela";
import { inflarXlsxComLimite, validarXlsxBytes, MAX_COMPRESSED_XLSX_BYTES } from "./xlsxBytes";

interface CargaDia extends Omit<EntradaDia, "imagens"> { tipo: "painelDia"; svgs: { svg: string; w: number; h: number }[]; }
// Painel Mês v1.1+ / Painel Dia v1.2+ mandam também periodo/filtros (rodapé do Excel); versões antigas não mandam
interface CargaMes { tipo: "painelMes"; titulo: string; geradoEm: string; periodo?: string; filtros?: string; linhas: LinhaMes[]; opcoes: EntradaMes["opcoes"]; card: { svg: string; w: number; h: number } | null; svgs: { svg: string; w: number; h: number }[]; }
// Pedido Diário: v2 (oficial) manda periodo/filtros e só tabela/dados; v1 mandava card/svgs (ignorados: o Excel sai sem imagens/topo)
interface CargaPedido { tipo: "pedidoDiario"; versao?: number; titulo: string; geradoEm: string; periodo?: string; filtros?: string; dias: number[]; linhas: LinhaPedido[]; opcoes: EntradaPedido["opcoes"]; card?: unknown; svgs?: { svg: string; w: number; h: number }[]; }
interface CargaReps extends EntradaReps { tipo: "repsSemPedidos"; svgs?: undefined; }
// Painel Metas v1.0: só células (aba igual à PAINEL METAS + aba Dados opcional), sem imagens
interface CargaMetas extends EntradaMetas { tipo: "painelMetas"; versao?: number; svgs?: undefined; }
// Pedidos Mês v1.0: formato compacto (tabelas de família e hierarquia rep > cliente > [família, centavos]); pelo link ou colado (k=cola)
interface CargaPM extends CargaPedidos { svgs?: undefined; }
// Painel de Vendas (Data Aprovação e Data do APP) v1.0: um tipo só ("painelVendas"), título, aba, regra e meses vêm na carga
interface CargaVendas extends CargaPV { svgs?: undefined; }
// Painel Metas | Representante v1.0: um bloco por representante (formato compacto), pelo link ou colado (k=cola, prefixo NFMR1.)
interface CargaMR extends CargaRep { svgs?: undefined; }
interface Carga extends Omit<EntradaXlsx, "imagens"> { svgs: { svg: string; w: number; h: number }[]; }

const $ = (s: string): HTMLElement => document.querySelector(s) as HTMLElement;
const erro = (m: string): void => {
    $("#st").textContent = "!"; $("#st").classList.add("erro");
    $("#tit").textContent = "Não foi possível preparar o arquivo"; $("#msg").textContent = m;
};
function b64u(s: string): Uint8Array {
    const b = s.replace(/-/g, "+").replace(/_/g, "/");
    const bin = atob(b.padEnd(Math.ceil(b.length / 4) * 4, "="));
    return Uint8Array.from(bin, c => c.charCodeAt(0));
}
async function inflar(b: Uint8Array): Promise<Uint8Array> {
    const ds = new DecompressionStream("deflate-raw");
    const w = ds.writable.getWriter(); void w.write(b as BufferSource); void w.close();
    return new Uint8Array(await new Response(ds.readable).arrayBuffer());
}
function svgParaPng(svg: string, w: number, h: number): Promise<Uint8Array> {
    return new Promise((resolve, reject) => {
        const img = new Image();
        img.onload = () => {
            const c = document.createElement("canvas"); c.width = w * 2; c.height = h * 2;
            const ctx = c.getContext("2d");
            if (!ctx) { reject(new Error("canvas indisponível")); return; }
            ctx.fillStyle = "#FFFFFF"; ctx.fillRect(0, 0, c.width, c.height);
            ctx.drawImage(img, 0, 0, c.width, c.height);
            c.toBlob(async b => { if (b) { resolve(new Uint8Array(await b.arrayBuffer())); } else { reject(new Error("falha ao gerar a imagem")); } }, "image/png");
        };
        img.onerror = () => reject(new Error("falha ao desenhar o gráfico"));
        const fonte = `@font-face{font-family:AntonPD;src:url(data:font/woff;base64,${ANTON_B64}) format("woff");}`;
        img.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svg.replace("/*FONTE*/", fonte));
    });
}
function entregar(bytes: Uint8Array, nome: string, mime: string): void {
    const a = $("#bt") as HTMLAnchorElement;
    a.href = URL.createObjectURL(new Blob([bytes as BlobPart], { type: mime }));
    a.download = nome;
    $("#st").classList.remove("erro");
    $("#st").textContent = "✓"; $("#tit").textContent = "Arquivo pronto";
    $("#msg").textContent = "Se o download não começar, use o botão abaixo.";
    $("#nome").textContent = nome;
    ["#arq", "#bt", "#nota"].forEach(s => $(s).classList.remove("oculto"));
    history.replaceState(null, "", location.pathname);
    a.click();
}

// Tabelas grandes (Pedidos Mês, Painel de Vendas, Painel Metas | Representante): o visual copia os dados ("NFPM1.", "NFPV1." ou "NFMR1." + base64url do deflate-raw)
// e abre a página com k=cola. A página pede para colar (Ctrl+V ou botão Colar dados) e monta o Excel com o texto colado.
// Nada sai do navegador.
const PREFIXOS_COLA: { [p: string]: (carga: unknown) => Uint8Array } = {
    "NFPM1.": c => gerarXlsxPedidos(expandir(c as CargaPedidos)),
    "NFPV1.": c => gerarXlsxPainelVendas(expandirPV(c as CargaPV)),
    "NFMR1.": c => gerarXlsxMetasRep(expandirMR(c as CargaRep))
};
function esperarCola(nome: string, mime: string, linhas: string, unidade: string): void {
    $("#st").textContent = "⇩"; $("#tit").textContent = "Cole os dados do painel";
    $("#msg").textContent = "A tabela" + (linhas ? " (" + Number(linhas).toLocaleString("pt-BR") + " " + (unidade || "linhas") + ")" : "") + " é grande demais para o link e foi copiada pelo painel. Clique em Colar dados ou pressione Ctrl+V nesta página.";
    const bt = $("#cola") as HTMLButtonElement;
    bt.classList.remove("oculto");
    let feito = false;
    const usar = async (texto: string): Promise<void> => {
        if (feito) { return; }
        const t = (texto || "").trim();
        const pref = t.indexOf("NFPT1.") === 0 ? "NFPT1." : t.indexOf("NFXLSX1.") === 0 ? "NFXLSX1." : Object.keys(PREFIXOS_COLA).find(k => t.indexOf(k) === 0);
        if (!pref) { erro("O texto colado não veio do painel. Volte ao Power BI, clique em Exportar > Baixar Excel de novo e cole aqui sem copiar outra coisa antes."); bt.classList.remove("oculto"); return; }
        feito = true; bt.classList.add("oculto");
        try {
            if (pref === "NFPT1.") {
                entregar(await gerarXlsxTabela(await decodificarPedidosTabela(t.slice(pref.length))), nome, mime);
            } else if (pref === "NFXLSX1.") {
                // Bound encoded input before atob allocates the compressed byte array.
                if (t.length - pref.length > Math.ceil(MAX_COMPRESSED_XLSX_BYTES / 3) * 4) {
                    throw new Error("O arquivo recebido ultrapassa 30 MB.");
                }
                entregar(await inflarXlsxComLimite(b64u(t.slice(pref.length))), nome, mime);
            } else {
                const carga = JSON.parse(new TextDecoder().decode(await inflar(b64u(t.slice(pref.length))))) as unknown;
                entregar(PREFIXOS_COLA[pref](carga), nome, mime);
            }
        } catch (e) {
            feito = false;
            if (pref === "NFPT1.") { erro(mensagemPedidos(e)); }
            else erro("Os dados colados estão incompletos. " + (e instanceof Error ? e.message : ""));
            bt.classList.remove("oculto");
        }
    };
    document.addEventListener("paste", ev => { ev.preventDefault(); void usar(ev.clipboardData ? ev.clipboardData.getData("text") : ""); });
    bt.addEventListener("click", async () => {
        try { await usar(await navigator.clipboard.readText()); }
        catch { $("#msg").textContent = "O navegador não deixou ler a área de transferência. Pressione Ctrl+V nesta página."; }
    });
    bt.focus();
}

(async () => {
    const p = new URLSearchParams(location.hash.slice(1));
    const nome = p.get("n") || "painel.xlsx";
    const mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    try {
        const j = p.get("j"), d = p.get("d");
        if (j && p.get("v") === "2") {
            try { entregar(await gerarXlsxTabela(await decodificarPedidosTabela(j, p.get("c") === "1")), nome, mime); }
            catch (e) { esperarCola(nome, mime, "", ""); erro(mensagemPedidos(e)); }
        } else if (j) {
            const bruto = p.get("c") === "1" ? await inflar(b64u(j)) : b64u(j);
            const carga = JSON.parse(new TextDecoder().decode(bruto)) as Carga | CargaDia | CargaMes | CargaPedido | CargaReps | CargaMetas | CargaPM | CargaVendas | CargaMR;
            const png = async (g: { svg: string; w: number; h: number }) => ({ png: await svgParaPng(g.svg, g.w, g.h), w: g.w, h: g.h });
            const imagens = [];
            const semImagens = (carga as CargaPedidos).tipo === "pedidosMes" || (carga as CargaPedido).tipo === "pedidoDiario" || (carga as CargaReps).tipo === "repsSemPedidos" || (carga as CargaMetas).tipo === "painelMetas" || (carga as CargaVendas).tipo === "painelVendas" || (carga as CargaMR).tipo === "painelMetasRep";
            for (const g of semImagens ? [] : carga.svgs || []) { imagens.push(await png(g)); }
            if ((carga as CargaMR).tipo === "painelMetasRep") {
                entregar(gerarXlsxMetasRep(expandirMR(carga as CargaMR)), nome, mime);
            } else if ((carga as CargaVendas).tipo === "painelVendas") {
                entregar(gerarXlsxPainelVendas(expandirPV(carga as CargaVendas)), nome, mime);
            } else if ((carga as CargaPedidos).tipo === "pedidosMes") {
                entregar(gerarXlsxPedidos(expandir(carga as CargaPedidos)), nome, mime);
            } else if ((carga as CargaMetas).tipo === "painelMetas") {
                const m = carga as CargaMetas;
                entregar(gerarXlsxMetas({ titulo: m.titulo, atualizacao: m.atualizacao || "", geradoEm: m.geradoEm, periodo: m.periodo || "", filtros: m.filtros || "", blocos: m.blocos || [], opcoes: { dados: !!(m.opcoes && m.opcoes.dados) } }), nome, mime);
            } else if ((carga as CargaReps).tipo === "repsSemPedidos") {
                const m = carga as CargaReps;
                entregar(gerarXlsxReps({ titulo: m.titulo, faixa: m.faixa, geradoEm: m.geradoEm, periodo: m.periodo, filtros: m.filtros, linhas: m.linhas, opcoes: m.opcoes }), nome, mime);
            } else if ((carga as CargaMes).tipo === "painelMes") {
                const m = carga as CargaMes;
                const card = m.card ? await png(m.card) : null;
                entregar(gerarXlsxMes({ titulo: m.titulo, geradoEm: m.geradoEm, periodo: m.periodo || "", filtros: m.filtros || "", linhas: m.linhas, opcoes: m.opcoes, card, medidores: imagens }), nome, mime);
            } else if ((carga as CargaPedido).tipo === "pedidoDiario") {
                const m = carga as CargaPedido;
                const op = m.opcoes || { tabela: true, dados: false };
                entregar(gerarXlsxPedido({ titulo: m.titulo, geradoEm: m.geradoEm, periodo: m.periodo || "", filtros: m.filtros || "", dias: m.dias || [], linhas: m.linhas || [], opcoes: { tabela: !!op.tabela, dados: !!op.dados } }), nome, mime);
            } else if ((carga as CargaDia).tipo === "painelDia") { entregar(gerarXlsxDia({ ...(carga as CargaDia), imagens }), nome, mime); }
            else { entregar(gerarXlsx({ ...(carga as Carga), imagens }), nome, mime); }
        } else if (p.get("k") === "cola") {
            esperarCola(nome, mime, p.get("q") || "", (p.get("u") || "").replace(/[^a-zà-ú ]/gi, ""));
        } else if (d) {
            const b = b64u(d);
            const tipo = p.get("t") || mime;
            const bytes = tipo === mime
                ? (p.get("c") === "1" ? await inflarXlsxComLimite(b) : validarXlsxBytes(b))
                : (p.get("c") === "1" ? await inflar(b) : b);
            entregar(bytes, nome, tipo);
        } else { erro("Nenhum arquivo recebido (página v2.6). Abra esta página pelo botão Exportar do painel."); }
    } catch (e) { erro("O arquivo chegou incompleto ou passou do limite do endereço. " + (e instanceof Error ? e.message : "")); }
})();
