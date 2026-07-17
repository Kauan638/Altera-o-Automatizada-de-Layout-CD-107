// =====================================
// VARIÁVEIS GLOBAIS
// =====================================

let dadosPosicoes = [];
let dadosABC = [];

let resultadoPosicoes = [];
let curvaABCAtual = { porSku:{}, colunaSku:null, colunaSaida:null, colunaDescricao:null, totalGeral:0, erro:null };
let sugestoesRealocacao = [];

// =====================================
// PAVILHÕES — mesma faixa de ruas (CODRUA) usada
// nos outros painéis do CD-107. Serve só para
// identificar em qual pavilhão a rua está; a REGRA
// de tamanho de posição é definida separadamente
// em REGRAS_TIPO_POSICAO (abaixo), pavilhão a
// pavilhão, e hoje só existe para o Pavilhão 2.
// =====================================

const PAVILHOES = [

    {
        nome:"Perecível",
        ruas:[
            [26,27],
            [29,31],
        ],
    },
    {
        nome:"Pavilhão 1",
        ruas:[
            [3,14],
            [21,24],
            [51,65],
        ],
    },
    {
        nome:"Pavilhão 2",
        ruas:[
            [71,106],
        ],
    },
    {
        nome:"Pavilhão 3",
        ruas:[
            [311,317],
        ],
    },

];

function obterPavilhao(rua){

    const r = Number(rua) || 0;

    const encontrado =
    PAVILHOES.find(p =>
        p.ruas.some(([ruaInicio, ruaFim]) =>
            r >= ruaInicio &&
            r <= ruaFim
        )
    );

    return encontrado
    ? encontrado.nome
    : "Sem Pavilhão";

}

// =====================================
// REGRAS DE TIPO DE POSIÇÃO POR PAVILHÃO
// =====================================
// Estrutura genérica: uma função por pavilhão que recebe
// (rua, predio) e devolve o TIPO de posição, ou null se a
// combinação rua/prédio não tiver regra conhecida. Pavilhões
// sem entrada aqui caem automaticamente em "Sem regra
// definida" — é só acrescentar uma função nova quando o
// layout de outro pavilhão for mapeado.
//
// Pavilhão 2 (única regra confirmada até agora):
//   Ruas 073-079                  -> Pallet Inteiro (SEMPRE, todo prédio)
//   Ruas 080-104, Prédio 133/134  -> Picking (Metade)
//   Ruas 080-104, Prédio 135+     -> Pallet Inteiro
//   Ruas 080-104, outro prédio    -> sem regra (fora da faixa informada)
//   Demais ruas do Pavilhão 2 (071/072/105/106) -> sem regra ainda

const REGRAS_TIPO_POSICAO = {

    "Pavilhão 2": function(rua, predio){

        const r = Number(rua);
        const p = Number(predio);

        if(r >= 73 && r <= 79){

            return "Pallet Inteiro";

        }

        if(r >= 80 && r <= 104){

            if(p === 133 || p === 134){

                return "Picking (Metade)";

            }

            if(p >= 135){

                return "Pallet Inteiro";

            }

            return null; // prédio fora da faixa 133+ informada

        }

        return null; // rua do Pavilhão 2 sem regra informada

    }

    // outros pavilhões entram aqui quando o layout for mapeado

};

function classificarTipoPosicao(rua, predio){

    const pavilhao = obterPavilhao(rua);

    const regra = REGRAS_TIPO_POSICAO[pavilhao];

    const tipo = regra ? regra(rua, predio) : null;

    if(!tipo){

        return {
            tipo:"Sem regra definida",
            tamanho:"desconhecido"
        };

    }

    const tamanho =
    tipo === "Pallet Inteiro"
    ? "grande"
    : tipo === "Picking (Metade)"
    ? "pequena"
    : "desconhecido";

    return { tipo, tamanho };

}

// =====================================
// INICIALIZAÇÃO — NOMES DE ARQUIVO
// =====================================

document
.getElementById("arquivoPosicoes")
.addEventListener("change", function(){

    const arquivo = this.files[0];

    document.getElementById("nomePosicoes").innerText =
    arquivo ? arquivo.name : "Nenhum arquivo selecionado";

});

document
.getElementById("arquivoABC")
.addEventListener("change", function(){

    const arquivo = this.files[0];

    document.getElementById("nomeABC").innerText =
    arquivo ? arquivo.name : "Nenhum arquivo selecionado";

});

// =====================================
// LOADING
// =====================================

function mostrarLoading(){

    document.getElementById("loading").style.display = "flex";

}

function ocultarLoading(){

    document.getElementById("loading").style.display = "none";

}

// =====================================
// LEITURA DE ARQUIVOS
// =====================================

function lerExcel(arquivo){

    return new Promise((resolve,reject)=>{

        const leitor = new FileReader();

        leitor.onload = e=>{

            const dados = new Uint8Array(e.target.result);

            const workbook = XLSX.read(dados, {type:"array"});

            const aba = workbook.SheetNames[0];

            const json = XLSX.utils.sheet_to_json(
                workbook.Sheets[aba],
                { defval:"" }
            );

            resolve(json);

        };

        leitor.onerror = reject;

        leitor.readAsArrayBuffer(arquivo);

    });

}

function lerTXT(arquivo){

    return new Promise((resolve,reject)=>{

        Papa.parse(arquivo, {

            header:true,
            delimiter:";",
            skipEmptyLines:true,

            complete:resultado=> resolve(resultado.data),

            error:erro=> reject(erro)

        });

    });

}

function lerCSV(arquivo){

    return new Promise((resolve,reject)=>{

        Papa.parse(arquivo, {

            header:true,
            skipEmptyLines:true,

            complete:resultado=> resolve(resultado.data),

            error:erro=> reject(erro)

        });

    });

}

// Detecta o leitor certo pra Análise ABC — normalmente vem
// de Excel (Metabase), mas aceita CSV também.

function lerArquivoABC(arquivo){

    const nome = arquivo.name.toLowerCase();

    if(nome.endsWith(".csv")){

        return lerCSV(arquivo);

    }

    return lerExcel(arquivo);

}

// =====================================
// DETECÇÃO FLEXÍVEL DE COLUNAS
// =====================================
// A Análise ABC por movimentação ainda não teve uma amostra
// real conferida — por isso a busca de colunas é por uma
// lista de nomes prováveis, normalizando acentos/maiúsculas,
// em vez de um nome fixo. Se o arquivo real usar um nome
// diferente dos candidatos abaixo, é só adicionar na lista
// (ou avisar para eu ajustar).

const CANDIDATOS_SKU_ABC = [
    "CODIGO","COD_PRODUTO","CODPRODUTO","SKU",
    "SEQPRODUTO","PRODUTO","COD"
];

const CANDIDATOS_DESCRICAO_ABC = [
    "DESCRICAO","DESCCOMPLETA","DESC_PRODUTO",
    "NOME_PRODUTO","PRODUTO_DESC","NOME"
];

const CANDIDATOS_SAIDA_ABC = [
    "SAIDA","QTD_SAIDA","QUANTIDADE_SAIDA","QUANTIDADE",
    "QTD_MOVIMENTACAO","MOVIMENTACAO","MOVIMENTACOES",
    "TOTAL_SAIDA","VOLUME","QTDE","QTD"
];

function normalizarChave(texto){

    return String(texto || "")
    .toUpperCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g,"")
    .replace(/[^A-Z0-9]/g,"");

}

function detectarColuna(linha, candidatos){

    if(!linha) return null;

    const chaves = Object.keys(linha);

    // 1) match exato (normalizado)
    for(const candidato of candidatos){

        const achado = chaves.find(k=>
            normalizarChave(k) === normalizarChave(candidato)
        );

        if(achado) return achado;

    }

    // 2) match por conteúdo (fallback)
    for(const candidato of candidatos){

        const achado = chaves.find(k=>
            normalizarChave(k).includes(normalizarChave(candidato))
        );

        if(achado) return achado;

    }

    return null;

}

// =====================================
// CURVA ABC (80% / 15% / 5%)
// =====================================
// Reconstrói a curva a partir da quantidade de saída/
// movimentação por SKU: ordena do maior giro pro menor,
// acumula o percentual do total e classifica:
//   Curva A -> até 80% acumulado
//   Curva B -> de 80% até 95% acumulado
//   Curva C -> de 95% até 100% acumulado

function calcularCurvaABC(dadosABC){

    if(!dadosABC || !dadosABC.length){

        return {
            porSku:{},
            colunaSku:null,
            colunaSaida:null,
            colunaDescricao:null,
            totalGeral:0,
            erro:"Arquivo de Análise ABC vazio ou não carregado."
        };

    }

    const colunaSku = detectarColuna(dadosABC[0], CANDIDATOS_SKU_ABC);
    const colunaSaida = detectarColuna(dadosABC[0], CANDIDATOS_SAIDA_ABC);
    const colunaDescricao = detectarColuna(dadosABC[0], CANDIDATOS_DESCRICAO_ABC);

    if(!colunaSku || !colunaSaida){

        return {
            porSku:{},
            colunaSku,
            colunaSaida,
            colunaDescricao,
            totalGeral:0,
            erro:
            "Não encontrei as colunas de SKU e/ou quantidade de saída no arquivo ABC. " +
            "Colunas disponíveis no arquivo: " + Object.keys(dadosABC[0]).join(", ")
        };

    }

    // agrupa por SKU somando a saída
    const agrupado = {};

    dadosABC.forEach(linha=>{

        const sku =
        String(linha[colunaSku] || "")
        .replace(",00","")
        .replace(".00","")
        .trim();

        if(!sku) return;

        const saida = Number(linha[colunaSaida]) || 0;

        if(!agrupado[sku]){

            agrupado[sku] = {
                sku,
                descricao: colunaDescricao ? linha[colunaDescricao] : "",
                saida:0
            };

        }

        agrupado[sku].saida += saida;

    });

    const lista =
    Object.values(agrupado)
    .sort((a,b)=> b.saida - a.saida);

    const totalGeral =
    lista.reduce((s,x)=> s + x.saida, 0);

    let acumulado = 0;

    const porSku = {};

    lista.forEach(item=>{

        acumulado += item.saida;

        const percentualAcumulado =
        totalGeral > 0
        ? (acumulado / totalGeral) * 100
        : 0;

        const percentualIndividual =
        totalGeral > 0
        ? (item.saida / totalGeral) * 100
        : 0;

        let curva = "C";

        if(percentualAcumulado <= 80){

            curva = "A";

        }
        else if(percentualAcumulado <= 95){

            curva = "B";

        }

        porSku[item.sku] = {
            sku:item.sku,
            descricao:item.descricao,
            saida:item.saida,
            percentualIndividual,
            percentualAcumulado,
            curva
        };

    });

    return {
        porSku,
        colunaSku,
        colunaSaida,
        colunaDescricao,
        totalGeral,
        erro:null
    };

}

// =====================================
// PROCESSAMENTO PRINCIPAL
// =====================================

async function processarLayout(){

    try{

        mostrarLoading();

        const arquivoPosicoes = document.getElementById("arquivoPosicoes").files[0];
        const arquivoABC = document.getElementById("arquivoABC").files[0];

        if(!arquivoPosicoes || !arquivoABC){

            alert("Selecione os dois arquivos: Posição de Endereços e Análise ABC por Movimentação.");

            ocultarLoading();

            return;

        }

        dadosPosicoes = await lerTXT(arquivoPosicoes);
        dadosABC = await lerArquivoABC(arquivoABC);

        console.log("Posições carregadas:", dadosPosicoes.length);
        console.log("Primeira posição:", dadosPosicoes[0]);

        console.log("ABC carregado:", dadosABC.length);
        console.log("Primeira linha ABC:", dadosABC[0]);

        curvaABCAtual = calcularCurvaABC(dadosABC);

        const avisoEl = document.getElementById("avisoABC");

        if(curvaABCAtual.erro){

            avisoEl.style.display = "block";

            avisoEl.innerText = "⚠️ " + curvaABCAtual.erro;

        }else{

            avisoEl.style.display = "none";

            console.log(
                "Curva ABC calculada — coluna SKU:", curvaABCAtual.colunaSku,
                "| coluna saída:", curvaABCAtual.colunaSaida,
                "| total geral:", curvaABCAtual.totalGeral
            );

        }

        montarResultadoPosicoes();

        ocultarLoading();

    }

    catch(erro){

        console.error(erro);

        ocultarLoading();

        alert("Erro ao processar arquivos.");

    }

}

// =====================================
// MONTA O RESULTADO CLASSIFICADO
// =====================================

function montarResultadoPosicoes(){

    resultadoPosicoes = [];

    dadosPosicoes.forEach(p=>{

        const rua = Number(p.CODRUA) || 0;
        const predio = Number(p.NROPREDIO) || 0;

        const classificacao = classificarTipoPosicao(rua, predio);

        const pavilhao = obterPavilhao(rua);

        const status =
        String(p.STATUS_ENDERECO || "")
        .toUpperCase()
        .trim();

        const livre = status === "DISPONIVEL";

        const sku =
        String(p.CODIGO || "")
        .replace(",00","")
        .replace(".00","")
        .trim();

        const infoABC =
        sku ? curvaABCAtual.porSku[sku] : null;

        resultadoPosicoes.push({

            pavilhao,
            rua,
            predio,
            apartamento:Number(p.NROAPARTAMENTO) || 0,
            sala:Number(p.NROSALA) || 0,

            endereco:`${p.CODRUA}.${p.NROPREDIO}.${p.NROAPARTAMENTO}.${p.NROSALA}`,

            especie: p.ESPECIE_END || "",

            tipo: classificacao.tipo,
            tamanho: classificacao.tamanho,

            livre,

            sku,
            descricao: p.DESCRICAO || (infoABC ? infoABC.descricao : ""),
            quantidade: Number(p.QTD_END || 0),

            curva: infoABC ? infoABC.curva : null,
            percentualAcumulado: infoABC ? infoABC.percentualAcumulado : null

        });

    });

    atualizarKPIsLayout();

    aplicarFiltrosPosicoes();

}

// =====================================
// KPIs
// =====================================

function atualizarKPIsLayout(){

    const comRegra =
    resultadoPosicoes.filter(x=> x.tamanho !== "desconhecido");

    document.getElementById("kpiTotalPosicoes").innerText =
    comRegra.length;

    document.getElementById("kpiVaziasGrandes").innerText =
    resultadoPosicoes.filter(x=> x.livre && x.tamanho === "grande").length;

    document.getElementById("kpiVaziasPequenas").innerText =
    resultadoPosicoes.filter(x=> x.livre && x.tamanho === "pequena").length;

    document.getElementById("kpiAltoGiroPequena").innerText =
    resultadoPosicoes.filter(x=>
        !x.livre && x.tamanho === "pequena" && x.curva === "A"
    ).length;

    document.getElementById("kpiSemRegra").innerText =
    resultadoPosicoes.filter(x=> x.tamanho === "desconhecido").length;

}

// =====================================
// RENDERIZAÇÃO DA TABELA DE POSIÇÕES
// =====================================

function badgeTamanho(tamanho){

    if(tamanho === "grande"){
        return `<span class="badge badge-grande">Pallet Inteiro</span>`;
    }

    if(tamanho === "pequena"){
        return `<span class="badge badge-pequena">Picking</span>`;
    }

    return `<span class="badge badge-indefinida">Sem regra</span>`;

}

function badgeCurva(curva){

    if(!curva){
        return "—";
    }

    return `<span class="badge badge-curva-${curva}">${curva}</span>`;

}

function renderizarTabelaPosicoes(dados){

    const tbody = document.getElementById("tbodyPosicoes");

    if(!dados.length){

        tbody.innerHTML =
        `<tr><td colspan="10" style="text-align:center;padding:30px;color:#6b7280;">
        Nenhuma posição encontrada com os filtros atuais.
        </td></tr>`;

        document.getElementById("resumoPaginacao").innerText = "";

        return;

    }

    // Tabela pode ficar grande (WMS tem ~30 mil endereços) —
    // mostra no máximo 500 linhas por vez pra não travar o
    // navegador, priorizando quem tem filtro aplicado.

    const LIMITE = 500;

    const dadosExibidos = dados.slice(0, LIMITE);

    let html = "";

    dadosExibidos.forEach(item=>{

        html += `
        <tr>
            <td>${item.pavilhao}</td>
            <td>${String(item.rua).padStart(3,"0")}</td>
            <td>${item.predio}</td>
            <td>${item.endereco}</td>
            <td>${item.especie}</td>
            <td>${badgeTamanho(item.tamanho)}</td>
            <td>${item.livre ? '<span class="badge badge-livre">Livre</span>' : '<span class="badge badge-ocupada">Ocupada</span>'}</td>
            <td>${item.sku || "—"}</td>
            <td style="text-align:left;">${item.descricao || "—"}</td>
            <td>${badgeCurva(item.curva)}</td>
        </tr>
        `;

    });

    tbody.innerHTML = html;

    document.getElementById("resumoPaginacao").innerText =
    dados.length > LIMITE
    ? `Mostrando ${LIMITE} de ${dados.length} posições — refine os filtros para ver o restante.`
    : `${dados.length} posições encontradas.`;

}

// =====================================
// FILTROS
// =====================================

function obterPosicoesFiltradas(){

    const skuFiltro =
    document.getElementById("filtroSKU").value.toLowerCase().trim();

    const ruaFiltro =
    document.getElementById("filtroRua").value.toLowerCase().trim();

    const statusFiltro =
    document.getElementById("filtroStatusPosicao").value;

    const tamanhoFiltro =
    document.getElementById("filtroTamanho").value;

    const pavilhaoFiltro =
    document.getElementById("filtroPavilhao").value;

    return resultadoPosicoes.filter(item=>{

        const skuOk =
        !skuFiltro ||
        item.sku.toLowerCase().includes(skuFiltro);

        const ruaOk =
        !ruaFiltro ||
        String(item.rua).toLowerCase().includes(ruaFiltro);

        const statusOk =
        !statusFiltro ||
        (statusFiltro === "LIVRE" && item.livre) ||
        (statusFiltro === "OCUPADA" && !item.livre);

        const tamanhoOk =
        !tamanhoFiltro ||
        item.tamanho === tamanhoFiltro;

        const pavilhaoOk =
        !pavilhaoFiltro ||
        item.pavilhao === pavilhaoFiltro;

        return skuOk && ruaOk && statusOk && tamanhoOk && pavilhaoOk;

    });

}

function aplicarFiltrosPosicoes(){

    renderizarTabelaPosicoes(
        obterPosicoesFiltradas()
    );

}

// =====================================
// MAPA DE POSIÇÕES GRANDES LIVRES (por rua)
// =====================================

function construirMapaGrandesLivres(){

    const mapa = {};

    resultadoPosicoes.forEach(p=>{

        if(!p.livre || p.tamanho !== "grande") return;

        if(!mapa[p.rua]) mapa[p.rua] = [];

        mapa[p.rua].push(p);

    });

    Object.values(mapa).forEach(lista=>{

        lista.sort((a,b)=> a.predio - b.predio);

    });

    return mapa;

}

function buscarGrandeLivre(rua, mapa){

    const candidatos = mapa[Number(rua)];

    if(!candidatos || !candidatos.length) return null;

    return candidatos.shift();

}

// =====================================
// SUGESTÃO DE REALOCAÇÃO — ALTO GIRO
// =====================================
// Candidato: SKU Curva A, ocupando hoje uma posição Picking
// (pequena) numa rua/prédio com regra definida. Destino:
// endereço Pallet Inteiro (grande) livre, priorizando a
// MESMA rua para minimizar deslocamento.

async function gerarSugestoesRealocacao(){

    if(!resultadoPosicoes.length){

        alert("Processe os arquivos primeiro.");

        return;

    }

    mostrarLoading();

    await new Promise(resolve=> setTimeout(resolve,50));

    try{

        const mapaGrandesLivres = construirMapaGrandesLivres();

        const candidatos =
        resultadoPosicoes.filter(p=>
            !p.livre &&
            p.tamanho === "pequena" &&
            p.curva === "A"
        );

        sugestoesRealocacao =
        candidatos.map(c=>{

            const destino =
            buscarGrandeLivre(c.rua, mapaGrandesLivres);

            return {

                sku:c.sku,
                descricao:c.descricao,
                curva:c.curva,
                percentualAcumulado:c.percentualAcumulado,
                rua:c.rua,
                enderecoAtual:c.endereco,
                quantidade:c.quantidade,
                enderecoDestino: destino ? destino.endereco : null

            };

        });

        // prioriza quem não tem destino ainda pendente por último,
        // mostrando primeiro quem já tem endereço pronto pra mover
        sugestoesRealocacao.sort((a,b)=>{

            if(!!a.enderecoDestino === !!b.enderecoDestino) return 0;

            return a.enderecoDestino ? -1 : 1;

        });

        abrirModalRealocacao();

    }

    catch(erro){

        console.error(erro);

        alert("Erro ao gerar sugestões de realocação.");

    }

    finally{

        ocultarLoading();

    }

}

// =====================================
// MODAL DE REALOCAÇÃO
// =====================================

function abrirModalRealocacao(){

    const modal = document.getElementById("modalRealocacao");

    if(!modal){

        alert("Modal de realocação não encontrado no HTML.");

        return;

    }

    document.getElementById("filtroRealocSKU").value = "";
    document.getElementById("filtroRealocRua").value = "";

    atualizarKPIsRealocacao();

    renderizarTabelaRealocacao(sugestoesRealocacao);

    modal.classList.add("ativo");

}

function fecharModalRealocacao(){

    document.getElementById("modalRealocacao").classList.remove("ativo");

}

function atualizarKPIsRealocacao(){

    document.getElementById("realocTotal").innerText =
    sugestoesRealocacao.length;

    document.getElementById("realocSemDestino").innerText =
    sugestoesRealocacao.filter(x=> !x.enderecoDestino).length;

    document.getElementById("realocComDestino").innerText =
    sugestoesRealocacao.filter(x=> x.enderecoDestino).length;

    const comPercentual =
    sugestoesRealocacao.filter(x=> x.percentualAcumulado !== null);

    const mediaPercentual =
    comPercentual.length
    ? comPercentual.reduce((s,x)=> s + x.percentualAcumulado, 0) / comPercentual.length
    : 0;

    document.getElementById("realocPercentual").innerText =
    `${mediaPercentual.toFixed(0)}%`;

}

function renderizarTabelaRealocacao(dados){

    const tbody = document.getElementById("tbodyRealocacao");

    if(!dados.length){

        tbody.innerHTML =
        `<tr><td colspan="6" style="text-align:center;padding:30px;color:#6b7280;">
        Nenhum item Curva A em posição pequena encontrado. Confira se os dois
        arquivos foram processados corretamente.
        </td></tr>`;

        return;

    }

    let html = "";

    dados.forEach(item=>{

        const moverParaTexto =
        item.enderecoDestino
        ? item.enderecoDestino
        : `<span style="color:#d32f2f;">Sem posição grande livre na rua ${String(item.rua).padStart(3,"0")}</span>`;

        html += `
        <tr>
            <td>${item.sku}</td>
            <td style="text-align:left;">${item.descricao || "—"}</td>
            <td>${badgeCurva(item.curva)}</td>
            <td>${item.enderecoAtual}</td>
            <td>${item.quantidade}</td>
            <td>${moverParaTexto}</td>
        </tr>
        `;

    });

    tbody.innerHTML = html;

}

function obterRealocacaoFiltrada(){

    const skuFiltro =
    document.getElementById("filtroRealocSKU").value.toLowerCase().trim();

    const ruaFiltro =
    document.getElementById("filtroRealocRua").value.toLowerCase().trim();

    return sugestoesRealocacao.filter(item=>{

        const skuOk =
        item.sku.toLowerCase().includes(skuFiltro);

        const ruaOk =
        !ruaFiltro ||
        String(item.rua).toLowerCase().includes(ruaFiltro);

        return skuOk && ruaOk;

    });

}

function aplicarFiltrosRealocacao(){

    renderizarTabelaRealocacao(
        obterRealocacaoFiltrada()
    );

}

function imprimirRealocacaoModal(){

    if(!sugestoesRealocacao.length){

        alert("Nenhuma sugestão para imprimir.");

        return;

    }

    const janela = window.open("", "_blank");

    if(!janela){

        alert("Permita pop-ups para este site.");

        return;

    }

    imprimirRealocacao(janela, obterRealocacaoFiltrada());

}

function imprimirRealocacao(janela, dadosBase){

    const dados =
    (dadosBase || sugestoesRealocacao)
    .filter(item=> item.enderecoDestino);

    if(!dados.length){

        alert("Nenhuma sugestão com posição livre para imprimir.");

        return;

    }

    let html = `
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>Sugestão de Realocação — Alto Giro</title>
<style>

@page{ size:A4 portrait; margin:8mm 8mm 14mm 8mm; }

*{ box-sizing:border-box; }

body{ font-family:Arial,Helvetica,sans-serif; margin:0; color:#222; }

h1{ margin:0; text-align:center; color:#0F4C81; font-size:22px; }

.info{ display:flex; justify-content:space-between; margin:15px 0; font-size:13px; }

table{ width:100%; border-collapse:collapse; }

th{ background:#0F4C81; color:#fff; padding:10px; border:1px solid #DDD; font-size:12px; }

td{ border:1px solid #DDD; padding:8px; font-size:11px; }

.curvaA{ background:#ffdede; }

@media print{
    th, .curvaA{ -webkit-print-color-adjust:exact; print-color-adjust:exact; }
}

</style>
</head>
<body>

<h1>🔄 SUGESTÃO DE REALOCAÇÃO — ALTO GIRO</h1>

<div class="info">
<div><b>Data:</b> ${new Date().toLocaleString("pt-BR")}</div>
<div><b>Total:</b> ${dados.length}</div>
</div>

<table>
<thead>
<tr>
<th>SKU</th>
<th>Produto</th>
<th>Curva</th>
<th>Posição Atual</th>
<th>Quantidade</th>
<th>Mover para</th>
</tr>
</thead>
<tbody>
`;

    dados.forEach(item=>{

        html += `
<tr class="curvaA">
<td><b>${item.sku}</b></td>
<td>${item.descricao || "—"}</td>
<td>${item.curva}</td>
<td>${item.enderecoAtual}</td>
<td style="text-align:center;">${item.quantidade}</td>
<td>${item.enderecoDestino}</td>
</tr>
`;

    });

    html += `
</tbody>
</table>

<script>
window.PagedConfig = {
    after: () => { window.focus(); window.print(); }
};
</script>
<script src="https://unpkg.com/pagedjs/dist/paged.polyfill.js"></script>

</body>
</html>
`;

    janela.document.open();
    janela.document.write(html);
    janela.document.close();

}

// =====================================
// LISTENERS
// =====================================

window.addEventListener("load", ()=>{

    document.getElementById("filtroSKU")
    .addEventListener("input", aplicarFiltrosPosicoes);

    document.getElementById("filtroRua")
    .addEventListener("input", aplicarFiltrosPosicoes);

    document.getElementById("filtroStatusPosicao")
    .addEventListener("change", aplicarFiltrosPosicoes);

    document.getElementById("filtroTamanho")
    .addEventListener("change", aplicarFiltrosPosicoes);

    document.getElementById("filtroPavilhao")
    .addEventListener("change", aplicarFiltrosPosicoes);

    document.getElementById("filtroRealocSKU")
    .addEventListener("input", aplicarFiltrosRealocacao);

    document.getElementById("filtroRealocRua")
    .addEventListener("input", aplicarFiltrosRealocacao);

    document.getElementById("modalRealocacao")
    .addEventListener("click", function(e){

        if(e.target === this){

            fecharModalRealocacao();

        }

    });

    document.addEventListener("keydown", function(e){

        if(e.key === "Escape"){

            fecharModalRealocacao();

        }

    });

});

console.log("SCRIPT CARREGADO — Alteração Automatizada de Layout CD 107");
