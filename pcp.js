import { atualizarRetornoPcp, emailAutenticado, solicitacoes } from "./firebase.js";
import { listaVendedores } from "./vendedores.js"; 
import { solicitacoesSelecionadasIds, setSolicitacoesSelecionadasIds } from "./script.js";

// Registrar o plugin DataLabels globalmente se estiver disponível
if (typeof ChartDataLabels !== 'undefined') {
    Chart.register(ChartDataLabels);
}

// Variáveis para armazenar as instâncias do Chart.js
let graficoStatus = null;
let graficoArea = null;
let graficoEvolucao = null;
let graficoAtendimento = null;
let graficoTopVendedores = null;
let filtroMesPcpJaPreenchido = false;
let filtroVendedoresPcpJaPreenchido = false;

/* ==========================================================================
   MODAL DE RESPOSTA PCP
   ========================================================================== */
export function abrirModal(docId) {
    setSolicitacoesSelecionadasIds([docId]);
    const item = solicitacoes.find(x => x.docId === docId);

    document.getElementById("modalTituloGeral").innerText = `Responder ID ${item.id}`;
    document.getElementById("responsavelPcp").value = item.responsavelPcp !== "-" ? item.responsavelPcp : "";
    document.getElementById("areaPcp").value = item.areaPcp !== "-" && item.areaPcp ? item.areaPcp : "Escolha";
    document.getElementById("respostaTexto").value = item.resposta || "";
    document.getElementById("dataProducao").value = item.dataProducao !== "-" ? item.dataProducao : "";
    document.getElementById("dataRetornoPcp").value = item.dataRetornoPcp !== "-" ? item.dataRetornoPcp : "";
    document.getElementById("novoStatus").value = item.status || "PENDENTE";
    
    document.getElementById("modalResposta").style.display = "flex";
}

export function abrirModalMassa() {
    const marcados = document.querySelectorAll(".chk-solicitacao-item:checked");
    let ids = Array.from(marcados).map(c => c.value);
    setSolicitacoesSelecionadasIds(ids);
    
    if (solicitacoesSelecionadasIds.length === 0) return alert("Selecione pelo menos uma linha.");

    document.getElementById("modalTituloGeral").innerText = `Responder ${solicitacoesSelecionadasIds.length} Itens Selecionados`;
    document.getElementById("responsavelPcp").value = "";
    document.getElementById("respostaTexto").value = "";
    
    document.getElementById("modalResposta").style.display = "flex";
}

export function fecharModal() {
    document.getElementById("modalResposta").style.display = "none";
}

export async function salvarResposta() {
    const marcados = document.querySelectorAll(".chk-solicitacao-item:checked");
    if (marcados.length > 0) {
        let ids = Array.from(marcados).map(c => c.value);
        setSolicitacoesSelecionadasIds(ids);
    }
    
    if (solicitacoesSelecionadasIds.length === 0) return alert("Nenhum item selecionado.");

    const respPcp = document.getElementById("responsavelPcp").value.trim();
    if (!respPcp) return alert("Insira o nome do Responsável pelo PCP.");

    const btn = document.getElementById("btnSalvarResposta");
    btn.disabled = true;
    btn.innerText = "SALVANDO...";

    const agora = new Date();
    const dados = {
        responsavelPcp: respPcp,
        areaPcp: document.getElementById("areaPcp").value,
        resposta: document.getElementById("respostaTexto").value,
        dataProducao: document.getElementById("dataProducao").value || "-",
        dataRetornoPcp: document.getElementById("dataRetornoPcp").value || "-",
        status: document.getElementById("novoStatus").value,
        dataAtendimento: agora.toLocaleDateString("pt-BR"),
        logAuditoria: `Modificado por ${emailAutenticado} em ${agora.toLocaleDateString("pt-BR")} às ${agora.toLocaleTimeString("pt-BR")}`
    };

    try {
        await atualizarRetornoPcp(solicitacoesSelecionadasIds, dados);
        fecharModal();
        alert(`Sucesso! ${solicitacoesSelecionadasIds.length} retornos atualizados.`);
        setSolicitacoesSelecionadasIds([]); 
    } catch (e) {
        alert("Erro ao salvar retorno.");
    } finally {
        btn.disabled = false;
        btn.innerText = "ENVIAR RETORNO";
    }
}

/* ==========================================================================
   INDICADORES PCP COM RÓTULOS, FILTRO MÊS E TOP VENDEDORES
   ========================================================================== */

export function toggleModoTV() {
    const container = document.getElementById("containerDashboardTV");
    if (!document.fullscreenElement && !document.webkitFullscreenElement) {
        if (container.requestFullscreen) {
            container.requestFullscreen().catch(e => console.error("Erro fullscreen:", e));
        } else if (container.webkitRequestFullscreen) {
            container.webkitRequestFullscreen();
        } else if (container.msRequestFullscreen) {
            container.msRequestFullscreen();
        }
    } else {
        if (document.exitFullscreen) {
            document.exitFullscreen();
        } else if (document.webkitExitFullscreen) {
            document.webkitExitFullscreen();
        } else if (document.msExitFullscreen) {
            document.msExitFullscreen();
        }
    }
}

document.addEventListener('fullscreenchange', handleFullscreenChange);
document.addEventListener('webkitfullscreenchange', handleFullscreenChange);

function handleFullscreenChange() {
    const container = document.getElementById("containerDashboardTV");
    const btnSair = document.getElementById("btnSairTV");
    
    if (document.fullscreenElement || document.webkitFullscreenElement) {
        container.classList.add("tv-mode-active");
        if (btnSair) btnSair.classList.remove("hidden");
    } else {
        container.classList.remove("tv-mode-active");
        if (btnSair) btnSair.classList.add("hidden");
    }
    
    setTimeout(() => {
        if(graficoStatus) graficoStatus.resize();
        if(graficoArea) graficoArea.resize();
        if(graficoEvolucao) graficoEvolucao.resize();
        if(graficoAtendimento) graficoAtendimento.resize();
        if(graficoTopVendedores) graficoTopVendedores.resize();
    }, 300);
}

function popularFiltroMesPcp() {
    const select = document.getElementById("filtroMesPcp");
    if (!select || filtroMesPcpJaPreenchido) return;

    const mesesNomes = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
    const mesesSet = new Set();

    solicitacoes.forEach(item => {
        if (item.dataSolicitacao && item.dataSolicitacao !== "-") {
            const parts = item.dataSolicitacao.split('/');
            if (parts.length === 3) mesesSet.add(`${parts[1]}/${parts[2]}`);
        }
        if (item.logAuditoria) {
            const match = item.logAuditoria.match(/em (\d{2})\/(\d{2})\/(\d{4})/);
            if (match) mesesSet.add(`${match[2]}/${match[3]}`);
        }
        if (item.dataAtendimento && item.dataAtendimento !== "-") {
            const parts = item.dataAtendimento.split('/');
            if (parts.length === 3) mesesSet.add(`${parts[1]}/${parts[2]}`);
        }
    });

    const valorAtual = select.value;
    select.innerHTML = '<option value="">Todos os Meses</option>';

    Array.from(mesesSet).sort((a,b) => {
        const [mA, yA] = a.split('/');
        const [mB, yB] = b.split('/');
        return yA !== yB ? yB - yA : mB - mA;
    }).forEach(mesAno => {
        const [m, y] = mesAno.split('/');
        const nomeMes = mesesNomes[parseInt(m) - 1];
        const option = document.createElement("option");
        option.value = mesAno; // "MM/YYYY"
        option.textContent = `${nomeMes}/${y}`;
        select.appendChild(option);
    });

    if (Array.from(mesesSet).includes(valorAtual)) {
        select.value = valorAtual;
    }
    filtroMesPcpJaPreenchido = true;
}

export function toggleMultiSelect(e) {
    e.stopPropagation();
    const dropdown = document.getElementById("dropdownVendedoresTop10");
    if(dropdown) dropdown.classList.toggle("hidden");
}

document.addEventListener("click", function(e) {
    const wrapper = document.getElementById("wrapperFiltroTop10");
    const dropdown = document.getElementById("dropdownVendedoresTop10");
    if(wrapper && dropdown && !wrapper.contains(e.target)) {
        dropdown.classList.add("hidden");
    }
});

function popularFiltroTop10Vendedores() {
    const dropdown = document.getElementById("dropdownVendedoresTop10");
    if(!dropdown || filtroVendedoresPcpJaPreenchido) return;

    dropdown.innerHTML = "";
    listaVendedores.sort().forEach(v => {
        dropdown.innerHTML += `
            <label class="multi-select-option">
                <input type="checkbox" value="${v}" class="chk-vendedor-top10" onchange="window.renderizarIndicadoresPcp()">
                ${v}
            </label>
        `;
    });
    filtroVendedoresPcpJaPreenchido = true;
}

export function renderizarIndicadoresPcp() {
    popularFiltroMesPcp();
    popularFiltroTop10Vendedores();
    
    const filtroMes = document.getElementById("filtroMesPcp").value;
    const hojeStr = new Date().toLocaleDateString("pt-BR");
    
    // Filtro 1: Mês Global
    let dadosFiltrados = solicitacoes;
    if (filtroMes) {
        dadosFiltrados = solicitacoes.filter(item => {
            let inMonth = false;
            if (item.dataSolicitacao && item.dataSolicitacao.includes(`/${filtroMes}`)) inMonth = true;
            if (item.logAuditoria && item.logAuditoria.includes(`/${filtroMes}`)) inMonth = true;
            if (item.dataAtendimento && item.dataAtendimento.includes(`/${filtroMes}`)) inMonth = true;
            return inMonth;
        });
    }

    const tituloKpiAtendimento = document.getElementById("tituloKpiAtendimento");
    if (tituloKpiAtendimento) {
        tituloKpiAtendimento.innerText = filtroMes ? "Total Atendimentos (Período)" : "Taxa Atendimento Diário (Hoje)";
    }

    // Variáveis Agregadoras Globais
    let atendimentosPcp = 0;
    const contagemStatus = { "PENDENTE": 0, "AGUARDANDO SUPRIMENTOS": 0, "AGUARDANDO COMERCIAL": 0, "PROGRAMADO": 0 };
    let dec = { pendente: 0, andamento: 0, concluido: 0 };
    let mon = { pendente: 0, andamento: 0, concluido: 0 };
    const historicoNovas = {}; 
    const historicoAtendimento = {};

    // Variável Específica para Top 10 Vendedores
    const contagemVendedores = {}; 
    const checkboxesVendedores = Array.from(document.querySelectorAll(".chk-vendedor-top10:checked")).map(cb => cb.value);

    // Iteração Única sobre Dados Filtrados
    dadosFiltrados.forEach(item => {
        // Atendimentos PCP
        if (item.logAuditoria) {
            const match = item.logAuditoria.match(/em (\d{2}\/\d{2}\/\d{4})/);
            if(match) {
                let dataLog = match[1];
                historicoAtendimento[dataLog] = (historicoAtendimento[dataLog] || 0) + 1;
                if (!filtroMes && dataLog === hojeStr) atendimentosPcp++;
                else if (filtroMes) atendimentosPcp++;
            }
        }

        // Status
        const st = item.status || "PENDENTE";
        if (contagemStatus[st] !== undefined) contagemStatus[st]++;
        else contagemStatus[st] = 1;

        // Áreas
        if (item.areaPcp === "Decoração") {
            if (st === "PENDENTE") dec.pendente++;
            else if (st === "PROGRAMADO") dec.concluido++;
            else dec.andamento++;
        } else if (item.areaPcp === "Montagem") {
            if (st === "PENDENTE") mon.pendente++;
            else if (st === "PROGRAMADO") mon.concluido++;
            else mon.andamento++;
        }

        // Evolução de Entradas
        if (item.dataSolicitacao && item.dataSolicitacao !== "-") {
            historicoNovas[item.dataSolicitacao] = (historicoNovas[item.dataSolicitacao] || 0) + 1;
        }

        // Vendedores
        if (item.vendedor && item.vendedor !== "-") {
            if (checkboxesVendedores.length === 0 || checkboxesVendedores.includes(item.vendedor)) {
                contagemVendedores[item.vendedor] = (contagemVendedores[item.vendedor] || 0) + 1;
            }
        }
    });

    document.getElementById("kpiPcpHoje").innerText = atendimentosPcp;
    document.getElementById("kpiDecPendente").innerText = `${dec.pendente} / ${dec.andamento}`;
    document.getElementById("kpiMonPendente").innerText = `${mon.pendente} / ${mon.andamento}`;

    if(graficoStatus) graficoStatus.destroy();
    if(graficoArea) graficoArea.destroy();
    if(graficoEvolucao) graficoEvolucao.destroy();
    if(graficoAtendimento) graficoAtendimento.destroy();
    if(graficoTopVendedores) graficoTopVendedores.destroy();

    // 1. Gráfico Rosca
    const ctxStatus = document.getElementById('chartStatusPcp').getContext('2d');
    graficoStatus = new Chart(ctxStatus, {
        type: 'doughnut',
        data: {
            labels: ['Pendente', 'Ag. Suprimentos', 'Ag. Comercial', 'Programado'],
            datasets: [{
                data: [contagemStatus["PENDENTE"], contagemStatus["AGUARDANDO SUPRIMENTOS"], contagemStatus["AGUARDANDO COMERCIAL"], contagemStatus["PROGRAMADO"]],
                backgroundColor: ['#f59e0b', '#0ea5e9', '#8b5cf6', '#22c55e']
            }]
        },
        options: { 
            responsive: true, maintainAspectRatio: false, layout: { padding: 20 },
            plugins: { 
                legend: { position: 'right' }, 
                title: { display: true, text: 'Status Global de Solicitações' },
                datalabels: {
                    color: '#fff',
                    font: { weight: 'bold', size: 12 },
                    textAlign: 'center',
                    formatter: (value, ctx) => {
                        if(value === 0) return null;
                        let sum = 0;
                        ctx.chart.data.datasets[0].data.forEach(data => sum += data);
                        let percentage = (value * 100 / sum).toFixed(1) + "%";
                        let label = ctx.chart.data.labels[ctx.dataIndex];
                        return `${label}
${value} (${percentage})`;
                    }
                }
            } 
        }
    });

    // 2. Gráfico Barras Empilhadas
    const ctxArea = document.getElementById('chartAreaPcp').getContext('2d');
    graficoArea = new Chart(ctxArea, {
        type: 'bar',
        data: {
            labels: ['Decoração', 'Montagem'],
            datasets: [
                { label: 'Pendente', data: [dec.pendente, mon.pendente], backgroundColor: '#f59e0b' },
                { label: 'Em Andamento', data: [dec.andamento, mon.andamento], backgroundColor: '#0ea5e9' },
                { label: 'Concluído / Programado', data: [dec.concluido, mon.concluido], backgroundColor: '#22c55e' }
            ]
        },
        options: { 
            responsive: true, maintainAspectRatio: false, 
            scales: { x: { stacked: true }, y: { stacked: true } }, 
            plugins: { 
                title: { display: true, text: 'Volume por Setor Produtivo' },
                datalabels: {
                    color: '#fff',
                    font: { weight: 'bold', size: 14 },
                    anchor: 'center',
                    align: 'center',
                    formatter: (value) => value > 0 ? value : null
                }
            } 
        }
    });

    // 3. Gráfico de Área (Evolução Diária)
    const datasEv = Object.keys(historicoNovas).sort((a,b) => {
        let pA = a.split('/'); let pB = b.split('/'); 
        return new Date(`${pA[2]}-${pA[1]}-${pA[0]}`) - new Date(`${pB[2]}-${pB[1]}-${pB[0]}`);
    }).slice(-15);
    
    const valsEv = datasEv.map(d => historicoNovas[d]);
    const ctxEv = document.getElementById('chartEvolucaoPcp').getContext('2d');
    graficoEvolucao = new Chart(ctxEv, {
        type: 'line',
        data: {
            labels: datasEv,
            datasets: [{
                label: 'Entrada de Novas Solicitações', data: valsEv, 
                borderColor: '#6366f1', backgroundColor: 'rgba(99, 102, 241, 0.2)', fill: true, tension: 0.3
            }]
        },
        options: { 
            responsive: true, maintainAspectRatio: false, layout: { padding: { top: 25 } },
            scales: { y: { suggestedMin: 0 } },
            plugins: { 
                legend: { display: false },
                title: { display: true, text: 'Evolução Diária (Últimos Dias da Seleção)' },
                datalabels: {
                    color: '#6366f1',
                    font: { weight: 'bold', size: 12 },
                    anchor: 'end',
                    align: 'top',
                    offset: 4,
                    formatter: (value) => value > 0 ? value : null
                }
            } 
        }
    });

    // 4. Gráfico de Colunas (Respostas PCP)
    const datasAt = Object.keys(historicoAtendimento).sort((a,b) => {
        let pA = a.split('/'); let pB = b.split('/'); 
        return new Date(`${pA[2]}-${pA[1]}-${pA[0]}`) - new Date(`${pB[2]}-${pB[1]}-${pB[0]}`);
    }).slice(-15);
    
    const valsAt = datasAt.map(d => historicoAtendimento[d]);
    const ctxAt = document.getElementById('chartAtendimentoPcp').getContext('2d');
    graficoAtendimento = new Chart(ctxAt, {
        type: 'bar',
        data: {
            labels: datasAt,
            datasets: [{
                label: 'Respostas do PCP', data: valsAt, backgroundColor: '#14b8a6', borderRadius: 4
            }]
        },
        options: { 
            responsive: true, maintainAspectRatio: false, layout: { padding: { top: 25 } },
            scales: { y: { suggestedMin: 0 } },
            plugins: { 
                legend: { display: false },
                title: { display: true, text: 'Produtividade de Retornos PCP (Resoluções)' },
                datalabels: {
                    color: '#14b8a6',
                    font: { weight: 'bold', size: 12 },
                    anchor: 'end',
                    align: 'top',
                    formatter: (value) => value > 0 ? value : null
                }
            } 
        }
    });

    // 5. Novo Gráfico Top 10 Vendedores
    const topVendedores = Object.entries(contagemVendedores)
        .sort((a, b) => b[1] - a[1]) // Do maior para o menor
        .slice(0, 10); // Apenas 10 primeiros
    
    const canvasContainerTop10 = document.getElementById("canvasContainerTop10");
    const emptyStateTop10 = document.getElementById("emptyStateTop10");

    if (topVendedores.length === 0) {
        canvasContainerTop10.classList.add("hidden");
        emptyStateTop10.classList.remove("hidden");
    } else {
        canvasContainerTop10.classList.remove("hidden");
        emptyStateTop10.classList.add("hidden");

        const labelsVendedores = topVendedores.map(v => v[0]);
        const dadosVendedores = topVendedores.map(v => v[1]);
        
        const ctxTopVend = document.getElementById('chartTopVendedoresPcp').getContext('2d');
        graficoTopVendedores = new Chart(ctxTopVend, {
            type: 'bar',
            data: {
                labels: labelsVendedores,
                datasets: [{
                    label: 'Solicitações', 
                    data: dadosVendedores, 
                    backgroundColor: '#8b5cf6', 
                    borderRadius: 4
                }]
            },
            options: {
                indexAxis: 'y', // Barras horizontais
                responsive: true, 
                maintainAspectRatio: false, 
                layout: { padding: { right: 35 } },
                scales: { x: { suggestedMin: 0 } },
                plugins: {
                    legend: { display: false }, 
                    title: { display: false }, 
                    datalabels: {
                        color: '#8b5cf6',
                        font: { weight: 'bold', size: 12 },
                        anchor: 'end',
                        align: 'right',
                        formatter: (value) => value > 0 ? value : null
                    }
                }
            }
        });
    }
}

/* ==========================================================================
   EXPORTAÇÕES (PCP e Geral)
   ========================================================================== */

export function exportarPendentesPcp() {
    // Exportar SOMENTE registros pendentes que requerem ação (Tudo exceto PROGRAMADO e CONCLUIDO)
    const pendentes = solicitacoes.filter(item => item.status !== 'PROGRAMADO' && item.status !== 'CONCLUIDO');
    
    if (pendentes.length === 0) {
        alert('Não há solicitações pendentes no momento para exportação.');
        return;
    }

    // A ordem obrigatória estabelecida: CLIENTE | COD ITEM (Box) | Nº PEDIDO | QTD | ID | STATUS ...
    const dadosExcel = pendentes.map(item => ({
        "CLIENTE": item.cliente || "-",
        "COD ITEM (Box)": item.codItem || "-",
        "Nº PEDIDO": item.numeroPedido || "-",
        "QTD": item.quantidade || 0,
        "ID": item.id || "-",
        "STATUS": item.status || "PENDENTE",
        "TIPO": item.tipoMaterial || "MTO",
        "MERCADO": item.mercadoSolicitante || "-",
        "VENDEDOR": item.vendedor || "-",
        "DATA SOLICITADA": item.dataSolicitacao || "-",
        "DATA CLIENTE": item.dataCliente || "-",
        "DATA PREVISTA": item.dataPrevista || "-",
        "DATA DESEJÁVEL": item.dataDesejavel || "-",
        "DATA ATENDIMENTO": item.dataAtendimento || "-",
        "DATA PRODUÇÃO": item.dataProducao || "-",
        "DATA RETORNO PCP": item.dataRetornoPcp || "-",
        "ÁREA PCP": item.areaPcp || "-",
        "RESPONSÁVEL PCP": item.responsavelPcp || "-",
        "OBSERVAÇÃO / RETORNO": item.observacao || ""
    }));

    const ws = XLSX.utils.json_to_sheet(dadosExcel);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Pendentes");

    const dataHoje = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `Solicitacoes_Pendentes_PCP_${dataHoje}.xlsx`);
}

export function exportarExcel() {
    const dados = solicitacoes.map(({ docId, ...resto }) => resto);
    const ws = XLSX.utils.json_to_sheet(dados);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "MTO");
    XLSX.writeFile(wb, "Programacao_MTO.xlsx");
}

export function exportarPDF() {
    if (window.jspdf && window.jspdf.jsPDF) {
        window.jsPDF = window.jspdf.jsPDF;
    }
    if (!window.jsPDF) {
        return alert("A biblioteca jsPDF não foi carregada. Verifique sua conexão com a internet.");
    }
    try {
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF('l', 'mm', 'a4');
        if (typeof doc.autoTable !== 'function') {
            return alert("O plugin de tabela PDF (jspdf-autotable) não foi carregado corretamente.");
        }
        doc.setFontSize(14);
        doc.text("Programação MTO - Relatório de Solicitações", 14, 15);
        doc.setFontSize(10);
        doc.setTextColor(100);
        doc.text(`Gerado em: ${new Date().toLocaleDateString('pt-BR')} ${new Date().toLocaleTimeString('pt-BR')}`, 14, 22);

        doc.autoTable({
            html: '#tabelaMtoHTML',
            startY: 28,
            theme: 'grid',
            headStyles: { fillColor: [15, 23, 42], textColor: [255, 255, 255], fontSize: 8, fontStyle: 'bold' },
            bodyStyles: { fontSize: 7, textColor: [51, 65, 85] },
            alternateRowStyles: { fillColor: [248, 250, 252] },
            margin: { top: 28, right: 10, bottom: 10, left: 10 }
        });

        doc.save(`Programacao_MTO_${new Date().toISOString().slice(0, 10)}.pdf`);
    } catch (error) {
        alert("Ocorreu um erro ao gerar o PDF. Pressione F12 e veja o Console para mais detalhes.");
    }
}

// VINCULAÇÕES DO MÓDULO PCP NO ESCOPO GLOBAL
window.abrirModal = abrirModal;
window.abrirModalMassa = abrirModalMassa;
window.fecharModal = fecharModal;
window.salvarResposta = salvarResposta;
window.toggleModoTV = toggleModoTV;
window.toggleMultiSelect = toggleMultiSelect;
window.renderizarIndicadoresPcp = renderizarIndicadoresPcp;
window.exportarPendentesPcp = exportarPendentesPcp;
window.exportarExcel = exportarExcel;
window.exportarPDF = exportarPDF;
