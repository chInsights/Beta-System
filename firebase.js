import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-app.js";
import { 
    getFirestore, collection, addDoc, onSnapshot, doc, updateDoc, deleteDoc, query, limit, orderBy
} from "https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js";
import { 
    getAuth, signInWithEmailAndPassword, signOut, onAuthStateChanged
} from "https://www.gstatic.com/firebasejs/10.12.0/firebase-auth.js";

// Configuração do Firebase
const firebaseConfig = {
    apiKey: "AIzaSyDz8qIZpkaj5ygJJJwyAfYH_3xKFxsZcU8",
    authDomain: "mto-webapp.firebaseapp.com",
    projectId: "mto-webapp",
    storageBucket: "mto-webapp.firebasestorage.app",
    messagingSenderId: "998118770114",
    appId: "1:998118770114:web:f29f30ed3c28ba5dc2c25b"
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app);
export const auth = getAuth(app);

// Estado Global de Autenticação e Dados
export let emailAutenticado = "";
export let solicitacoes = [];
let unsubscribeRealtime = null;

// Observador de Login
onAuthStateChanged(auth, (user) => {
    if (user && window.configurarSessaoUsuario) {
        emailAutenticado = user.email.trim().toLowerCase();
        window.configurarSessaoUsuario(emailAutenticado);
    }
});

// Realizar Login
export async function fazerLogin(email, senha) {
    return await signInWithEmailAndPassword(auth, email, senha);
}

// Realizar Logout
export function fazerLogout() {
    return signOut(auth);
}

// Escutar atualizações do Firestore em tempo real
export function iniciarOuvinteFirestore(limite, callbackRender) {
    if (unsubscribeRealtime) unsubscribeRealtime();

    const q = query(
        collection(db, "solicitacoes"),
        orderBy("id", "desc"),
        limit(limite)
    );
    
    unsubscribeRealtime = onSnapshot(q, (snapshot) => {
        solicitacoes = [];
        snapshot.forEach((docSnap) => {
            let dados = docSnap.data();
            dados.docId = docSnap.id;
            solicitacoes.push(dados);
        });
        callbackRender();
    });
}

// Salvar Novos Pedidos
export async function salvarNovoPedido(pedidosArray) {
    for (let i = 0; i < pedidosArray.length; i++) {
        await addDoc(collection(db, "solicitacoes"), pedidosArray[i]);
    }
}

// Deletar Solicitação
export async function excluirSolicitacaoBanco(docId) {
    await deleteDoc(doc(db, "solicitacoes", docId));
}

// Atualizar Retorno em Massa
export async function atualizarRetornoPcp(idsArray, dadosAtualizados) {
    for (let i = 0; i < idsArray.length; i++) {
        const docId = idsArray[i];
        await updateDoc(doc(db, "solicitacoes", docId), dadosAtualizados);
    }
}