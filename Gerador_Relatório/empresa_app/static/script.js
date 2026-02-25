// Formatação de dinheiro
function formatarMoeda(inputId) {
    const input = document.getElementById(inputId);

    input.addEventListener("input", function(e) {
        let valor = e.target.value;

        valor = valor.replace(/\D/g, "");
        valor = (Number(valor) / 100).toFixed(2) + "";
        valor = valor.replace(".", ",");
        valor = valor.replace(/\B(?=(\d{3})+(?!\d))/g, ".");

        e.target.value = valor;
    });
}

// Outra função exemplo para analisar
function mostrarAlerta() {
    alert("Sistema ativo!");
}

// Executa quando a página carrega / 2 valor separados
document.addEventListener("DOMContentLoaded", function() {
    formatarMoeda("valor_item");
    formatarMoeda("custo_produto"); // se tiver outro campo
});