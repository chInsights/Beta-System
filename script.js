import { 
    fazerLogin, fazerLogout, iniciarOuvinteFirestore, 
    salvarNovoPedido, excluirSolicitacaoBanco, 
    emailAutenticado, solicitacoes 
} from "./firebase.js";
import { listaVendedores } from "./vendedores.js"; 
import { renderizarIndicadoresPcp } from "./pcp.js"; // Importa funções específicas do PCP

let usuarioAtual = "";
let itensDoPedidoAtual = [];
export let solicitacoesSelecionadasIds = []; // Exportado para uso no pcp.js
let filtroMesAtual = "";
let filtroVendedorAtual = ""; 
let filtroStatusAtual = "TODOS";
let limiteRegistros = 100;

// Método exportado para modificar a seleção a partir do PCP
export function setSolicitacoesSelecionadasIds(ids) {
    solicitacoesSelecionadasIds.length = 0;
    if (ids && ids.length) solicitacoesSelecionadasIds.push(...ids);
}

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

function inicializarSelects() {
    carregarSelectVendedores();
}

if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", inicializarSelects);
} else {
    inicializarSelects();
}

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

    if (email === "programacaomto@vendedor.com" || email.includes("vendedor")) {
        document.getElementById("formSolicitante").classList.remove("hidden");
        configuringDataSolicitacaoAutomatica();
        aplicarBloqueioDatasRetroativas();
        carregarSelectVendedores();
    }

    iniciarOuvinteFirestore(limiteRegistros, renderTabela);
};

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
    if(btn) btn.classList.add("active");
}

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

// === FUNÇÃO CORRIGIDA ===
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
            const workbook = XLSX.read(e.target.result, { type: 'binary' });
            const linhas = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
            let contador = 0;

            linhas.forEach(linha => {
                // Normaliza os cabeçalhos (remove acentos, espaços e deixa tudo minúsculo)
                const linhaNormalizada = {};
                for (let chave in linha) {
                    const chaveLimpa = chave.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
                    linhaNormalizada[chaveLimpa] = linha[chave];
                }

                // Busca o código pelas variações mais comuns
                const codRaw = linhaNormalizada["codigo"] || linhaNormalizada["cod"] || linhaNormalizada["item"] || linhaNormalizada["produto"] || "";
                const cod = String(codRaw).trim();
                
                // Busca a quantidade pelas variações mais comuns
                const qtdRaw = linhaNormalizada["quantidade"] || linhaNormalizada["qtd"] || linhaNormalizada["quant"] || 0;
                const qtd = Number(qtdRaw);
                
                // Busca o tipo do material
                const tipoRaw = linhaNormalizada["tipo"] || linhaNormalizada["tipomaterial"] || "MTO";
                const tipoImp = String(tipoRaw).trim();

                if (cod && cod !== "undefined" && qtd > 0) {
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
            
            if (contador > 0) {
                alert(`Sucesso! ${contador} itens importados.`);
            } else {
                alert("Atenção: O arquivo foi lido, mas nenhum item válido foi encontrado. Verifique se as colunas de 'Código' e 'Quantidade' estão preenchidas corretamente na planilha.");
            }
            
        } catch (erro) {
            console.error(erro);
            alert("Erro ao ler o arquivo Excel. Verifique o formato.");
        } finally {
            event.target.value = "";
        }
    };
    leitor.readAsBinaryString(arquivo);
}
// =========================

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
        renderMonitoramentoVendedor();
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

export function renderTabela() {
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
            ? `<td class="col-checkbox"><input type="checkbox" value="${item.docId}" ${isChecked} class="chk-solicitacao-item chk-solicitacao" onclick="window.gerenciarSelecaoItem(this)"></td>`
            : `<td class="col-checkbox id-pcp-view hidden"></td>`;

        let botoesAcao = emailAutenticado === "atendimento@pcp.com"
            ? `<button class="action-btn" onclick="window.abrirModal('${item.docId}')">RESPONDER</button><button class="delete-btn" onclick="window.deletarSolicitacao('${item.docId}')"><i class="fa-solid fa-trash"></i></button>`
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

    const tabMonitor = document.getElementById("abaMonitorarSolicitacoes");
    if (tabMonitor && !tabMonitor.classList.contains("hidden")) {
        renderMonitoramentoVendedor();
    }
    
    // Atualiza indicadores caso a página PCP Indicadores esteja ativa
    const pageIndicadores = document.getElementById("indicadoresPcpPage");
    if (pageIndicadores && pageIndicadores.classList.contains("active")) {
        renderizarIndicadoresPcp();
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

window.deletarSolicitacao = async function(docId) {
    if (!confirm("Deseja apagar esta solicitação?")) return;
    try {
        await excluirSolicitacaoBanco(docId);
        alert("Solicitação deletada!");
    } catch (e) {
        alert("Erro ao remover.");
    }
}

/* VINCULAÇÃO GLOBAL DOS BOTÕES DO ESCOPO GERAL */
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
window.filtrarPorStatus = (st) => { filtroStatusAtual = st; renderTabela(); };
window.filtrarMes = () => { 
    filtroMesAtual = document.getElementById("filtroMes").value; 
    renderTabela(); 
};
window.limparFiltro = () => {
    filtroMesAtual = "";
    document.getElementById("filtroMes").value = ""; 
    renderTabela();
};
window.toggleSelecionarTodos = (m) => { 
    document.querySelectorAll(".chk-solicitacao-item").forEach(c => { 
        c.checked = m.checked; 
        window.gerenciarSelecaoItem(c); 
    }); 
};
window.gerenciarSelecaoItem = (c) => { 
    if (c.checked) { 
        if (!solicitacoesSelecionadasIds.includes(c.value)) solicitacoesSelecionadasIds.push(c.value); 
    } else { 
        const index = solicitacoesSelecionadasIds.indexOf(c.value);
        if (index > -1) solicitacoesSelecionadasIds.splice(index, 1);
    }
};
window.mudarAbaVendedor = mudarAbaVendedor;
window.renderMonitoramentoVendedor = renderMonitoramentoVendedor;