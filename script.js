import { 
    fazerLogin, fazerLogout, iniciarOuvinteFirestore, 
    salvarNovoPedido, excluirSolicitacaoBanco, atualizarRetornoPcp,
    emailAutenticado, solicitacoes 
} from "./firebase.js";
import { listaVendedores } from "./vendedores.js"; // Importa a lista do arquivo separado

// Variáveis Locais da Interface
let usuarioAtual = "";
let itensDoPedidoAtual = [];
let solicitacoesSelecionadasIds = [];
let filtroMesAtual = "";
let filtroVendedorAtual = ""; 
let filtroStatusAtual = "TODOS";
let limiteRegistros = 100;

/* POPULAR SELECT DE VENDEDORES (FORMULÁRIO, FILTRO E MONITORAMENTO) */
function carregarSelectVendedores() {
    const select = document.getElementById("vendedorNome");
    const selectMonitor = document.getElementById("monitorVendedorNome");
    
    if (select) select.innerHTML = '<option value="">Selecione o Vendedor...</option>';
    if (selectMonitor) selectMonitor.innerHTML = '<option value="">Selecione o Vendedor...</option>';

    listaVendedores.sort().forEach(nome => {
        if (select) {
            const option = document.createElement("option");
            option.value = nome;
            option.textContent = nome;
            select.appendChild(option);
        }
        
        if (selectMonitor) {
            const optionMonitor = document.createElement("option");
            optionMonitor.value = nome;
            optionMonitor.textContent = nome;
            selectMonitor.appendChild(optionMonitor);
        }
    });
}

function carregarSelectFiltroVendedores() {
    const select = document.getElementById("filtroVendedor");
    if (!select) return;

    select.innerHTML = '<option value="">Todos os Vendedores</option>';

    listaVendedores.sort().forEach(nome => {
        const option = document.createElement("option");
        option.value = nome;
        option.textContent = nome;
        select.appendChild(option);
    });
}

// Inicialização segura dos selects (funciona mesmo se o DOM já carregou)
function inicializarSelects() {
    carregarSelectVendedores();
    carregarSelectFiltroVendedores();
}

if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", inicializarSelects);
} else {
    inicializarSelects();
}

/* CONTROLE DE SESSÃO */
window.configurarSessaoUsuario = function(email) {
    if (email === "programacaomto@vendedor.com" || email.includes("vendedor")) {
        usuarioAtual = "Vendedor/Comercial";
    } else if (email === "atendimento@pcp.com") {
        usuarioAtual = "Atendimento-PCP"; 
        document.getElementById("btnRespostaMassa").classList.remove("hidden");
        document.querySelectorAll(".id-pcp-view").forEach(el => el.classList.remove("hidden"));
    } else {
        usuarioAtual = email.split('@')[0];
    }

    document.getElementById("loginScreen").style.display = "none";
    document.getElementById("app").style.display = "flex";
    document.getElementById("usuarioLogado").innerText = usuarioAtual;

    carregarSelectFiltroVendedores(); 

    if (email === "programacaomto@vendedor.com" || email.includes("vendedor")) {
        document.getElementById("formSolicitante").classList.remove("hidden");
        configuringDataSolicitacaoAutomatica();
        aplicarBloqueioDatasRetroativas();
        carregarSelectVendedores();
    }

    iniciarOuvinteFirestore(limiteRegistros, renderTabela);
};

/* LOGIN & LOGOUT */
async function login() {
    const email = document.getElementById("usuario").value.trim().toLowerCase();
    const senha = document.getElementById("senha").value;
    const btn = document.getElementById("btnLogin");

    if (!email || !senha) return alert("Por favor, preencha o e-mail e a senha.");

    btn.disabled = true;
    btn.innerText = "ACESSANDO...";

    try {
        await fazerLogin(email, senha);
    } catch (error) {
        alert("Credenciais inválidas ou erro de rede.");
        btn.disabled = false;
        btn.innerText = "ACESSAR SISTEMA";
    }
}

function logout() {
    fazerLogout().then(() => location.reload());
}

function abrirPagina(id, btn) {
    document.querySelectorAll(".page").forEach(p => p.classList.remove("active"));
    document.getElementById(id).classList.add("active");
    document.querySelectorAll(".menu button").forEach(b => b.classList.remove("active"));
    btn.classList.add("active");
}

/* CÁLCULOS E DATAS */
function aplicarBloqueioDatasRetroativas() {
    const hoje = new Date();
    const dataMinima = `${hoje.getFullYear()}-${String(hoje.getMonth() + 1).padStart(2, '0')}-${String(hoje.getDate()).padStart(2, '0')}`;
    document.getElementById("itemDataCliente").min = dataMinima;
    document.getElementById("itemDataDesejavel").min = dataMinima;
}

function configuringDataSolicitacaoAutomatica() {
    const hoje = new Date();
    document.getElementById("dataSolicitacao").value = `${String(hoje.getDate()).padStart(2, '0')}/${String(hoje.getMonth() + 1).padStart(2, '0')}/${hoje.getFullYear()}`;
}

function calcularPrevisaoItem() {
    const tipo = document.getElementById("itemTipoMaterial").value;
    const qtd = document.getElementById("itemQuantidade").value;
    const campo = document.getElementById("itemPrevisaoEntrega");

    if (!qtd || qtd <= 0) { 
        campo.value = ""; 
        return; 
    }
    campo.value = calcularPrevisaoEspecifica(tipo, qtd);
}

function calcularPrevisaoEspecifica(tipoMaterial, quantidade) {
    if (!quantidade || quantidade <= 0) return "-";
    
    const dataCalculo = new Date();
    const tipoNormalizado = String(tipoMaterial).trim().toUpperCase();
    
    const diasAdicionais = (tipoNormalizado === "MTO") ? 30 : 25;
    
    dataCalculo.setDate(dataCalculo.getDate() + diasAdicionais);
    
    const dia = String(dataCalculo.getDate()).padStart(2, '0');
    const mes = String(dataCalculo.getMonth() + 1).padStart(2, '0');
    const ano = dataCalculo.getFullYear();
    
    return `${dia}/${mes}/${ano}`;
}

/* IMPORTAÇÃO EXCEL DE ITENS */
function importarItensDoExcel(event) {
    const arquivo = event.target.files[0];
    const cli = document.getElementById("cliente").value.trim();
    const numPedido = document.getElementById("numeroPedido").value.trim();

    if (!cli || !numPedido) {
        alert("Preencha CLIENTE e NÚMERO DO PEDIDO antes de importar.");
        event.target.value = "";
        return;
    }
    if (!arquivo) return;

    const leitor = new FileReader();
    leitor.onload = function(e) {
        try {
            const linhas = XLSX.utils.sheet_to_json(XLSX.read(e.target.result, { type: 'binary' }).Sheets[XLSX.read(e.target.result, { type: 'binary' }).SheetNames[0]]);
            let contador = 0;

            linhas.forEach(linha => {
                const cod = String(linha["Item"] || linha["Código"] || linha["item"] || "").trim();
                const qtd = Number(linha["Quantidade"] || linha["Qtd"] || linha["qtd"] || 0);
                const tipoImp = String(linha["Tipo"] || linha["TipoMaterial"] || "MTO").trim();

                if (cod && qtd > 0) {
                    itensDoPedidoAtual.push({
                        codItem: cod,
                        tipoMaterial: tipoImp,
                        quantidade: qtd,
                        dataPrevista: calcularPrevisaoEspecifica(tipoImp, qtd),
                        dataCliente: "-",
                        dataDesejavel: "-",
                        numeroPedido: numPedido
                    });
                    contador++;
                }
            });

            renderListaItensProvisorios();
            alert(`Sucesso! ${contador} itens importados.`);
        } catch (erro) {
            alert("Erro ao ler o arquivo Excel.");
        } finally {
            event.target.value = "";
        }
    };
    leitor.readAsBinaryString(arquivo);
}

function adicionarItemNaLista() {
    const cod = document.getElementById("itemCod").value.trim();
    const tipo = document.getElementById("itemTipoMaterial").value;
    const qtd = document.getElementById("itemQuantidade").value;
    const prev = document.getElementById("itemPrevisaoEntrega").value;
    const dtCliRaw = document.getElementById("itemDataCliente").value;
    const dtDesRaw = document.getElementById("itemDataDesejavel").value;
    const numPedido = document.getElementById("numeroPedido").value.trim();

    if (!numPedido) return alert("Digite o NÚMERO DO PEDIDO.");
    if (!cod || !qtd || !prev) return alert("Preencha Código, Quantidade e gere a Previsão.");

    itensDoPedidoAtual.push({
        codItem: cod,
        tipoMaterial: tipo,
        quantidade: Number(qtd),
        dataPrevista: prev,
        dataCliente: dtCliRaw ? `${dtCliRaw.split('-')[2]}/${dtCliRaw.split('-')[1]}/${dtCliRaw.split('-')[0]}` : "-",
        dataDesejavel: dtDesRaw ? `${dtDesRaw.split('-')[2]}/${dtDesRaw.split('-')[1]}/${dtDesRaw.split('-')[0]}` : "-",
        numeroPedido: numPedido
    });

    document.getElementById("itemCod").value = "";
    document.getElementById("itemQuantidade").value = "";
    document.getElementById("itemPrevisaoEntrega").value = "";
    document.getElementById("itemDataCliente").value = "";
    document.getElementById("itemDataDesejavel").value = "";

    renderListaItensProvisorios();
}

function removerItemDaLista(index) {
    itensDoPedidoAtual.splice(index, 1);
    renderListaItensProvisorios();
}

function renderListaItensProvisorios() {
    const tbody = document.getElementById("listaItensProvisorios");
    if (itensDoPedidoAtual.length === 0) {
        tbody.innerHTML = `<tr><td colspan="3" style="text-align: center; color: #94a3b8; font-style: italic; padding: 14px;">Nenhum item adicionado.</td></tr>`;
        return;
    }

    tbody.innerHTML = "";
    itensDoPedidoAtual.forEach((item, index) => {
        tbody.innerHTML += `
            <tr class="item-row" ondblclick="removerItemDaLista(${index})" title="Duplo clique para remover">
                <td><strong>${item.codItem}</strong> <small>(${item.tipoMaterial})</small></td>
                <td><strong>${item.quantidade}</strong></td>
                <td>${item.numeroPedido}</td>
            </tr>`;
    });
}

/* ENVIAR PEDIDO */
async function enviarSolicitacaoMultiiens() {
    const vendedor = document.getElementById("vendedorNome").value.trim();
    const mercadoValor = document.getElementById("mercadoSolicitante").value;
    const cli = document.getElementById("cliente").value.trim();
    const btn = document.getElementById("btnEnviarSol");

    if (!vendedor || !cli) return alert("Preencha VENDEDOR e CLIENTE.");
    if (itensDoPedidoAtual.length === 0) return alert("Adicione pelo menos 1 item.");

    btn.disabled = true;
    btn.innerText = "ENVIANDO PEDIDO...";

    try {
        const payload = itensDoPedidoAtual.map((item, i) => ({
            id: Date.now() + i,
            vendedor: vendedor,
            tipoMaterial: item.tipoMaterial,
            mercadoSolicitante: mercadoValor,
            dataSolicitacao: document.getElementById("dataSolicitacao").value,
            dataCliente: item.dataCliente,
            dataPrevista: item.dataPrevista,
            dataDesejavel: item.dataDesejavel,
            dataAtendimento: "-", dataProducao: "-", dataRetornoPcp: "-", areaPcp: "-", responsavelPcp: "-",
            cliente: cli,
            codItem: item.codItem,
            numeroPedido: item.numeroPedido,
            quantidade: item.quantidade,
            observacao: document.getElementById("observacao").value,
            remetenteEmail: emailAutenticado,
            destinatario: document.getElementById("destinatario").value.trim().toLowerCase(),
            status: "PENDENTE", resposta: "", logAuditoria: ""
        }));

        await salvarNovoPedido(payload);
        alert(`Sucesso! ${itensDoPedidoAtual.length} itens cadastrados.`);

        document.getElementById("cliente").value = "";
        document.getElementById("numeroPedido").value = "";
        document.getElementById("observacao").value = "";
        itensDoPedidoAtual = [];
        renderListaItensProvisorios();

    } catch (error) {
        alert("Erro ao gravar dados.");
    } finally {
        btn.disabled = false;
        btn.innerText = "ENVIAR PEDIDO COMPLETO";
    }
}

/* ==========================================================================
   NOVA FUNCIONALIDADE: NAVEGAÇÃO E MONITORAMENTO VENDEDOR
   ========================================================================== */
function mudarAbaVendedor(aba) {
    const tabNova = document.getElementById("abaNovaSolicitacao");
    const tabMonitor = document.getElementById("abaMonitorarSolicitacoes");
    const btns = document.querySelectorAll(".tabs-vendedor .tab-btn");
    
    if (!tabNova || !tabMonitor) return;

    btns.forEach(b => b.classList.remove("active"));

    if (aba === 'nova') {
        tabNova.classList.remove("hidden");
        tabMonitor.classList.add("hidden");
        btns[0].classList.add("active");
    } else {
        tabNova.classList.add("hidden");
        tabMonitor.classList.remove("hidden");
        btns[1].classList.add("active");
        renderMonitoramentoVendedor(); // Atualiza a lista ao abrir
    }
}

function renderMonitoramentoVendedor() {
    const vendedor = document.getElementById("monitorVendedorNome").value;
    const tbody = document.getElementById("listaMonitoramento");
    
    if (!tbody) return;

    if (!vendedor) {
        tbody.innerHTML = `<tr><td colspan="10" style="text-align: center; color: #94a3b8; font-style: italic; padding: 14px;">Selecione um vendedor para monitorar.</td></tr>`;
        return;
    }

    // Filtra as solicitações na memória baseada no vendedor
    const filtradas = solicitacoes.filter(item => item.vendedor === vendedor);

    if (filtradas.length === 0) {
        tbody.innerHTML = `<tr><td colspan="10" style="text-align: center; color: #94a3b8; font-style: italic; padding: 14px;">Nenhuma solicitação encontrada para este vendedor.</td></tr>`;
        return;
    }

    tbody.innerHTML = "";
    filtradas.forEach(item => {
        let cssStatus = item.status ? item.status.toLowerCase().replace(/\s+/g, '-') : 'pendente';
        
        let ultimaAtualizacao = "-";
        if (item.logAuditoria) {
             const match = item.logAuditoria.match(/em (\d{2}\/\d{2}\/\d{4} às \d{2}:\d{2}:\d{2})/);
             ultimaAtualizacao = match ? match[1].replace(' às ', ' ') : item.logAuditoria.substring(0, 30) + "...";
        } else if (item.dataAtendimento && item.dataAtendimento !== "-") {
             ultimaAtualizacao = item.dataAtendimento;
        }

        tbody.innerHTML += `
            <tr>
                <td>${item.numeroPedido || "-"}</td>
                <td>${item.dataSolicitacao || "-"}</td>
                <td>${item.cliente || "-"}</td>
                <td><strong>${item.codItem || "-"}</strong> <small>(${item.tipoMaterial || "MTO"})</small></td>
                <td>${item.quantidade || 0}</td>
                <td><span class="status-badge status-${cssStatus}">${item.status || 'PENDENTE'}</span></td>
                <td>${item.responsavelPcp || "-"}</td>
                <td style="color:#1d4ed8; font-weight:700;">${item.dataPrevista || "-"}</td>
                <td>${ultimaAtualizacao}</td>
                <td>
                    <div style="max-width: 250px; white-space: normal; word-wrap: break-word;">
                        ${item.observacao ? `<b>Obs:</b> ${item.observacao}<br>` : ""}
                        ${item.resposta ? `<b style="color:#0284c7;">Retorno:</b> ${item.resposta}` : ""}
                    </div>
                </td>
            </tr>
        `;
    });
}
/* ========================================================================== */

/* TABELA E FILTROS */
function renderTabela() {
    const tabela = document.getElementById("tabelaSolicitacoes");
    const busca = document.getElementById("inputBusca").value.toLowerCase().trim();
    tabela.innerHTML = "";

    let baseDados = emailAutenticado === "atendimento@pcp.com" 
        ? solicitacoes 
        : solicitacoes.filter(item => item.remetenteEmail === emailAutenticado);

    let filtradas = baseDados.filter(item => {
        if (filtroStatusAtual !== "TODOS" && item.status !== filtroStatusAtual) return false;
        if (filtroVendedorAtual && item.vendedor !== filtroVendedorAtual) return false;
        if (filtroMesAtual && item.dataSolicitacao) {
            const p = item.dataSolicitacao.split('/');
            if (p.length === 3 && `${p[2]}-${p[1]}` !== filtroMesAtual) return false;
        }
        if (busca) {
            return (item.cliente && item.cliente.toLowerCase().includes(busca)) ||
                   (item.codItem && item.codItem.toLowerCase().includes(busca)) ||
                   (item.vendedor && item.vendedor.toLowerCase().includes(busca)) ||
                   (item.numeroPedido && String(item.numeroPedido).toLowerCase().includes(busca));
        }
        return true;
    });

    filtradas.forEach(item => {
        let isChecked = solicitacoesSelecionadasIds.includes(item.docId) ? 'checked' : '';
        let tdCheckbox = emailAutenticado === "atendimento@pcp.com"
            ? `<td class="col-checkbox"><input type="checkbox" value="${item.docId}" ${isChecked} class="chk-solicitacao-item chk-solicitacao" onclick="gerenciarSelecaoItem(this)"></td>`
            : `<td class="col-checkbox id-pcp-view hidden"></td>`;

        let botoesAcao = emailAutenticado === "atendimento@pcp.com"
            ? `<button class="action-btn" onclick="abrirModal('${item.docId}')">RESPONDER</button><button class="delete-btn" onclick="deletarSolicitacao('${item.docId}')"><i class="fa-solid fa-trash"></i></button>`
            : '-';
            
        let cssStatus = item.status ? item.status.toLowerCase().replace(/\s+/g, '-') : 'pendente';

        tabela.innerHTML += `
        <tr>
            ${tdCheckbox}
            <td>${item.id}</td>
            <td><strong>${item.tipoMaterial || "MTO"}</strong></td>
            <td>${item.mercadoSolicitante || "-"}</td>
            <td>${item.vendedor || "-"}</td>
            <td>${item.dataSolicitacao || "-"}</td>
            <td>${item.dataCliente || "-"}</td>
            <td style="color:#1d4ed8; font-weight:700;">${item.dataPrevista || "-"}</td>
            <td>${item.dataDesejavel || "-"}</td>
            <td>${item.dataAtendimento || "-"}</td>
            <td>${item.dataProducao || "-"}</td>
            <td>${item.dataRetornoPcp || "-"}</td>
            <td>${item.areaPcp || "-"}</td>
            <td><strong>${item.responsavelPcp || "-"}</strong></td>
            <td>${item.cliente || "-"}</td>
            <td>${item.codItem || "-"}</td>
            <td>${item.numeroPedido || "-"}</td>
            <td>${item.quantidade || 0}</td>
            <td>
                <div>${item.observacao || ""}</div>
                ${item.resposta ? `<div class="response"><b>RETORNO PCP:</b><br>${item.resposta}</div>` : ''}
                ${item.logAuditoria ? `<div class="log-auditoria">${item.logAuditoria}</div>` : ''}
            </td>
            <td><span class="status-badge status-${cssStatus}">${item.status || 'PENDENTE'}</span></td>
            <td>${botoesAcao}</td>
        </tr>`;
    });

    atualizarKPIs(baseDados);

    // Atualiza automaticamente o monitoramento se a aba estiver aberta
    const tabMonitor = document.getElementById("abaMonitorarSolicitacoes");
    if (tabMonitor && !tabMonitor.classList.contains("hidden")) {
        renderMonitoramentoVendedor();
    }
}

function atualizarKPIs(dados) {
    document.getElementById("kpiTotal").innerText = dados.length;
    document.getElementById("kpiPendentes").innerText = dados.filter(x => x.status === "PENDENTE").length;
    document.getElementById("kpiAguardandoSup").innerText = dados.filter(x => x.status === "AGUARDANDO SUPRIMENTOS").length;
    document.getElementById("kpiAguardandoCom").innerText = dados.filter(x => x.status === "AGUARDANDO COMERCIAL").length;
    document.getElementById("kpiProgramados").innerText = dados.filter(x => x.status === "PROGRAMADO").length;
}

function carregarMaisRegistros() {
    limiteRegistros += 50;
    iniciarOuvinteFirestore(limiteRegistros, renderTabela);
}

async function deletarSolicitacao(docId) {
    if (!confirm("Deseja apagar esta solicitação?")) return;
    try {
        await excluirSolicitacaoBanco(docId);
        alert("Solicitação deletada!");
    } catch (e) {
        alert("Erro ao remover.");
    }
}

/* MODAL E RESPOSTAS */
function abrirModal(docId) {
    solicitacoesSelecionadasIds = [docId];
    const item = solicitacoes.find(x => x.docId === docId);

    document.getElementById("modalTituloGeral").innerText = "Responder Solicitação";
    document.getElementById("responsavelPcp").value = item.responsavelPcp !== "-" ? item.responsavelPcp : "";
    document.getElementById("areaPcp").value = item.areaPcp !== "-" && item.areaPcp ? item.areaPcp : "Escolha";
    document.getElementById("respostaTexto").value = item.resposta || "";
    document.getElementById("dataProducao").value = item.dataProducao !== "-" ? item.dataProducao : "";
    document.getElementById("dataRetornoPcp").value = item.dataRetornoPcp !== "-" ? item.dataRetornoPcp : "";
    document.getElementById("novoStatus").value = item.status || "PENDENTE";
    document.getElementById("modalResposta").style.display = "flex";
}

function abrirModalMassa() {
    const marcados = document.querySelectorAll(".chk-solicitacao-item:checked");
    if (marcados.length > 0) solicitacoesSelecionadasIds = Array.from(marcados).map(c => c.value);
    if (solicitacoesSelecionadasIds.length === 0) return alert("Selecione pelo menos uma linha.");

    document.getElementById("modalTituloGeral").innerText = `Responder ${solicitacoesSelecionadasIds.length} Itens Selecionados`;
    document.getElementById("responsavelPcp").value = "";
    document.getElementById("respostaTexto").value = "";
    document.getElementById("modalResposta").style.display = "flex";
}

function fecharModal() {
    document.getElementById("modalResposta").style.display = "none";
}

async function salvarResposta() {
    const marcados = document.querySelectorAll(".chk-solicitacao-item:checked");
    if (marcados.length > 0) solicitacoesSelecionadasIds = Array.from(marcados).map(c => c.value);
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
        solicitacoesSelecionadasIds = [];
    } catch (e) {
        alert("Erro ao salvar retorno.");
    } finally {
        btn.disabled = false;
        btn.innerText = "ENVIAR RETORNO";
    }
}

/* EXPORTAÇÃO EXCEL */
function exportarExcel() {
    const dados = solicitacoes.map(({ docId, ...resto }) => resto);
    const ws = XLSX.utils.json_to_sheet(dados);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "MTO");
    XLSX.writeFile(wb, "Programacao_MTO.xlsx");
}

/* EXPORTAÇÃO PDF */
export function exportarPDF() {
    console.log("Iniciando geração do PDF...");

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
            headStyles: {
                fillColor: [15, 23, 42],
                textColor: [255, 255, 255],
                fontSize: 8,
                fontStyle: 'bold'
            },
            bodyStyles: {
                fontSize: 7,
                textColor: [51, 65, 85]
            },
            alternateRowStyles: {
                fillColor: [248, 250, 252]
            },
            margin: { top: 28, right: 10, bottom: 10, left: 10 }
        });

        doc.save(`Programacao_MTO_${new Date().toISOString().slice(0, 10)}.pdf`);
        console.log("PDF baixado com sucesso!");

    } catch (error) {
        console.error("Erro detalhado na geração do PDF:", error);
        alert("Ocorreu um erro ao gerar o PDF. Pressione F12 e veja o Console para mais detalhes.");
    }
}

/* VINCULAÇÃO GLOBAL DOS BOTÕES */
window.login = login;
window.logout = logout;
window.abrirPagina = abrirPagina;
window.calcularPrevisaoItem = calcularPrevisaoItem;
window.importarItensDoExcel = importarItensDoExcel;
window.adicionarItemNaLista = adicionarItemNaLista;
window.removerItemDaLista = removerItemDaLista;
window.enviarSolicitacaoMultiiens = enviarSolicitacaoMultiiens;
window.renderTabela = renderTabela;
window.carregarMaisRegistros = carregarMaisRegistros;
window.deletarSolicitacao = deletarSolicitacao;
window.abrirModal = abrirModal;
window.abrirModalMassa = abrirModalMassa;
window.fecharModal = fecharModal;
window.salvarResposta = salvarResposta;
window.exportarExcel = exportarExcel;
window.exportarPDF = exportarPDF;
window.filtrarPorStatus = (st) => { filtroStatusAtual = st; renderTabela(); };
window.filtrarMes = () => { 
    filtroMesAtual = document.getElementById("filtroMes").value; 
    renderTabela(); 
};
window.filtrarVendedor = () => { 
    filtroVendedorAtual = document.getElementById("filtroVendedor").value; 
    renderTabela(); 
};
window.limparFiltro = () => {
    filtroMesAtual = "";
    filtroVendedorAtual = "";
    document.getElementById("filtroMes").value = ""; 
    const selectVendedor = document.getElementById("filtroVendedor");
    if (selectVendedor) selectVendedor.value = ""; 
    renderTabela();
};
window.toggleSelecionarTodos = (m) => { 
    document.querySelectorAll(".chk-solicitacao-item").forEach(c => { 
        c.checked = m.checked; 
        gerenciarSelecaoItem(c); 
    }); 
};
window.gerenciarSelecaoItem = (c) => { 
    if (c.checked) { 
        if (!solicitacoesSelecionadasIds.includes(c.value)) solicitacoesSelecionadasIds.push(c.value); 
    } else { 
        solicitacoesSelecionadasIds = solicitacoesSelecionadasIds.filter(x => x !== c.value); 
    }
};
window.mudarAbaVendedor = mudarAbaVendedor;
window.renderMonitoramentoVendedor = renderMonitoramentoVendedor;
