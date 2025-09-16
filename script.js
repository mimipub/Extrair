// Função para gerar a planilha Excel
function gerarExcel() {
    // Pega o texto do textarea
    const texto = document.getElementById('texto').value;

    // Divide o texto usando a data como delimitador
    const blocos = texto.split(/(\d{2}\/\d{2}\/\d{4})/);
    const registros = [];

    for (let i = 1; i < blocos.length; i += 2) {
        const data = blocos[i].trim();
        const linhas = blocos[i + 1].split('\n').map(l => l.trim()).filter(l => l && l !== "-");

        // Se a quantidade de linhas for maior ou igual a 6, adiciona ao array de registros
        if (linhas.length >= 6) {
            registros.push({
                Nome: linhas[0],
                Email: linhas[1],
                Telefone: linhas[2],
                Data: data,
                Tipo: linhas[3],
                Veículo: linhas[4],
                Estado: linhas[5],
                Cidade: linhas[6] || ""
            });
        }
    }

    // Cria a planilha com os dados
    const ws = XLSX.utils.json_to_sheet(registros);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Cadastros");

    // Faz o download da planilha
    XLSX.writeFile(wb, "cadastros.xlsx");
}

// Adiciona evento ao botão
document.getElementById('gerarExcel').addEventListener('click', gerarExcel);
