function verificarEstoqueCritico() {

  const planilha = SpreadsheetApp.getActiveSpreadsheet();

  const abaEstoque = planilha.getSheetByName("Estoque");
  const abaEmails = planilha.getSheetByName("E-mails");

  if (!abaEstoque || !abaEmails) {
    throw new Error("Verifique os nomes das abas: Estoque e E-mails");
  }

  const dados = abaEstoque.getDataRange().getValues();
  const listaEmails = abaEmails.getDataRange().getValues();

  const itensCriticos = [];
  const itensOK = [];

  // =========================
  // PROCESSAMENTO DOS DADOS
  // =========================
  for (let i = 1; i < dados.length; i++) {

    const [
      equipamento,
      local,
      item,
      quantidade
    ] = dados[i];

    const qtd = Number(quantidade);

    if (isNaN(qtd)) continue;

    const registro = {
      equipamento,
      local,
      item,
      quantidade: qtd
    };

    if (qtd === 0) {
      itensCriticos.push(registro);
    } else {
      itensOK.push(registro);
    }
  }

  if (itensCriticos.length === 0 && itensOK.length === 0) {
    return;
  }

  // =========================
  // EMAILS
  // =========================
  const destinatarios = [];
  const copia = [];

  for (let i = 1; i < listaEmails.length; i++) {

    if (listaEmails[i][0]) {
      destinatarios.push(listaEmails[i][0]);
    }

    if (listaEmails[i][1]) {
      copia.push(listaEmails[i][1]);
    }
  }

  // =========================
  // HTML
  // =========================
  const html = `
    <html>
      <head>
        <style>

          body{
            font-family: Arial, sans-serif;
            background-color:#f4f6f9;
            padding:20px;
            color:#333;
          }

          .container{
            max-width:1000px;
            margin:auto;
            background:#ffffff;
            border-radius:12px;
            overflow:hidden;
            box-shadow:0 2px 10px rgba(0,0,0,0.08);
          }

          .header{
            background:linear-gradient(135deg,#1f4e78,#0b2c4d);
            color:white;
            padding:25px;
          }

          .header h1{
            margin:0;
            font-size:28px;
          }

          .content{
            padding:25px;
          }

          .cards{
            display:flex;
            gap:15px;
            margin-bottom:25px;
          }

          .card{
            flex:1;
            padding:20px;
            border-radius:10px;
            color:white;
            text-align:center;
          }

          .critico{
            background:#d9534f;
          }

          .ok{
            background:#28a745;
          }

          .card h2{
            margin:0;
            font-size:32px;
          }

          .card p{
            margin-top:5px;
            font-size:14px;
          }

          .section-title{
            margin-top:30px;
            margin-bottom:10px;
            font-size:20px;
            font-weight:bold;
          }

          table{
            width:100%;
            border-collapse:collapse;
            margin-top:10px;
          }

          th{
            background:#1f4e78;
            color:white;
            padding:12px;
            text-align:left;
          }

          td{
            padding:10px;
            border-bottom:1px solid #ddd;
          }

          tr:hover{
            background:#f5f5f5;
          }

          .linha-critica{
            background:#ffe5e5;
          }

          .footer{
            padding:20px;
            font-size:13px;
            color:#666;
            text-align:center;
            border-top:1px solid #eee;
          }

        </style>
      </head>

      <body>

        <div class="container">

          <div class="header">
            <h1>📦 Relatório de Estoque</h1>
            <p>Posição atual dos itens em estoque</p>
          </div>

          <div class="content">

            <div class="cards">

              <div class="card critico">
                <h2>${itensCriticos.length}</h2>
                <p>Itens Críticos</p>
              </div>

              <div class="card ok">
                <h2>${itensOK.length}</h2>
                <p>Itens Disponíveis</p>
              </div>

            </div>

            ${gerarTabela(
              "⚠️ Itens Críticos",
              itensCriticos,
              true
            )}

            ${gerarTabela(
              "✅ Itens em Estoque",
              itensOK,
              false
            )}

          </div>

          <div class="footer">
            Este e-mail foi enviado automaticamente pelo sistema de controle de estoque.
          </div>

        </div>

      </body>
    </html>
  `;

  // =========================
  // ENVIO
  // =========================
  MailApp.sendEmail({
    to: "ubametra.admexterno@gmail.com,",
    subject: "📦 Relatório de Estoque",
    htmlBody: html
  });

}


// ======================================
// FUNÇÃO PARA GERAR TABELAS
// ======================================
function gerarTabela(titulo, dados, critico = false) {

  if (dados.length === 0) {
    return `
      <div class="section-title">${titulo}</div>
      <p>Nenhum item encontrado.</p>
    `;
  }

  let linhas = "";

  dados.forEach(item => {

    linhas += `
      <tr class="${critico ? "linha-critica" : ""}">
        <td>${item.equipamento}</td>
        <td>${item.local}</td>
        <td>${item.item}</td>
        <td><b>${item.quantidade}</b></td>
      </tr>
    `;
  });

  return `
    <div class="section-title">${titulo}</div>

    <table>

      <thead>
        <tr>
          <th>Equipamento</th>
          <th>Local</th>
          <th>Item</th>
          <th>Quantidade</th>
        </tr>
      </thead>

      <tbody>
        ${linhas}
      </tbody>

    </table>
  `;
}