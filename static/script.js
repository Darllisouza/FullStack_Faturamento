document.addEventListener('DOMContentLoaded', function () {
    // Captura o formulário de upload do PDF
    const pdfForm = document.getElementById('uploadForm');

    // Adiciona um ouvinte de evento para o envio do formulário de PDF
    pdfForm.addEventListener('submit', function (event) {
        event.preventDefault(); // Impede o envio do formulário padrão

        const formData = new FormData(pdfForm); // Cria um objeto FormData com os dados do formulário

        // Envia a requisição POST para a rota '/processar_pdf' com os dados do formulário
        fetch('/processar_pdf', {
            method: 'POST',
            body: formData // Define o corpo da requisição como FormData
        })
        .then(response => response.text()) // Converte a resposta para texto
        .then(message => {
            console.log(message);
        })
        .catch(error => {
            console.error('Erro:', error);
            alert('Ocorreu um erro ao processar o PDF.'); // Exibe uma mensagem de erro em caso de falha na requisição
        });
    });

    // Captura o formulário de geração do documento Word
    const wordForm = document.getElementById('wordForm');

    // Adiciona um ouvinte de evento para o envio do formulário de Word
    wordForm.addEventListener('submit', function (event) {
        event.preventDefault(); // Impede o envio do formulário padrão

        // Envia a requisição POST para a rota '/gerar_documento_word'
        fetch('/gerar_documento_word', {
            method: 'POST'
        })
        .then(response => response.text()) // Converte a resposta para texto
        .then(message => {
            console.log(message);
            window.open('/baixar_documento_word', '_blank');
        })
        .catch(error => {
            console.error('Erro:', error);
            alert('Ocorreu um erro ao gerar o documento Word.'); // Exibe uma mensagem de erro em caso de falha na requisição
        });
    });
});
