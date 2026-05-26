function verificarEstoqueCritico() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const abaEstoque = planilha.getSheetByName("Estoque");
  const abaEmails = planilha.getSheetByName("E-mails");

  if (!abaEstoque || !abaEmails) {
    throw new Error("Verifique os nomes das abas: 'Estoque' e 'E-mails'");
  }

  const dados = abaEstoque.getDataRange().getValues();
  const listaEmails = abaEmails.getDataRange().getValues();

  let itensCriticos = [];
  let itensOK = [];

  // 🔍 Processa os dados
  for (let i = 1; i < dados.length; i++) {
    const equipamento = dados[i][0];
    const local = dados[i][1];
    const item = dados[i][2];
    const quantidade = Number(dados[i][3]);

    if (isNaN(quantidade)) continue;

    if (quantidade === 0) {
      itensCriticos.push([equipamento, local, item, quantidade]);
    } else if (quantidade >= 1) {
      itensOK.push([equipamento, local, item, quantidade]);
    }
  }

  if (itensCriticos.length === 0 && itensOK.length === 0) return;

  // 📧 Emails
  let destinatarios = [];
  let copia = [];

  for (let i = 1; i < listaEmails.length; i++) {
    if (listaEmails[i][0]) destinatarios.push(listaEmails[i][0]);
    if (listaEmails[i][1]) copia.push(listaEmails[i][1]);
  }

  // ✉️ INÍCIO DO HTML (ESTILO APLICADO)
  let mensagem = `
    <div style="font-family: Georgia, 'Times New Roman', serif; font-size: 16px; color: #333;">

      <h2 style="font-size:22px; border-bottom:2px solid #ccc; padding-bottom:5px;">
        📊 Relatório de Estoque
      </h2>

      <p>Prezados,</p>
      <p>Segue abaixo a posição atual do estoque:</p>
  `;

  // 🔴 TABELA CRÍTICOS
  if (itensCriticos.length > 0) {
    mensagem += `
      <h3 style="color:red; font-size:18px;">
        ⚠️ Itens Críticos (Estoque Zerado)
      </h3>
      <p>Total: <b>${itensCriticos.length}</b></p>

      <table border="1" cellpadding="8" cellspacing="0" 
      style="border-collapse: collapse; font-size:15px; width:100%;">
        <tr style="background-color:#eaeaea; font-weight:bold;">
          <th>Equipamento</th>
          <th>Local</th>
          <th>Item</th>
          <th>Quantidade</th>
        </tr>
    `;

    itensCriticos.forEach(linha => {
      mensagem += `
        <tr style="background-color:#ffcccc;">
          <td>${linha[0]}</td>
          <td>${linha[1]}</td>
          <td>${linha[2]}</td>
          <td>${linha[3]}</td>
        </tr>
      `;
    });

    mensagem += `</table><br>`;
  }

  // 🟢 TABELA OK
  if (itensOK.length > 0) {
    mensagem += `
      <h3 style="color:green; font-size:18px;">
        ✅ Itens em Estoque
      </h3>
      <p>Total: <b>${itensOK.length}</b></p>

      <table border="1" cellpadding="8" cellspacing="0" 
      style="border-collapse: collapse; font-size:15px; width:100%;">
        <tr style="background-color:#eaeaea; font-weight:bold;">
          <th>Equipamento</th>
          <th>Local</th>
          <th>Item</th>
          <th>Quantidade</th>
        </tr>
    `;

    itensOK.forEach(linha => {
      mensagem += `
        <tr>
          <td>${linha[0]}</td>
          <td>${linha[1]}</td>
          <td>${linha[2]}</td>
          <td>${linha[3]}</td>
        </tr>
      `;
    });

    mensagem += `</table>`;
  }

  // 🔚 FINALIZAÇÃO
  mensagem += `
      <p style="margin-top:20px;">
        <b>Ação necessária:</b> Repor os itens críticos.
      </p>

      <p>
        Atenciosamente,<br>
        <b>Controle de Estoque</b>
      </p>

    </div>
  `;

  // 🚀 ENVIO
  MailApp.sendEmail({
    to: destinatarios.join(","),
    cc: copia.join(","),
    subject: "📊 Relatório de Estoque (Crítico + Geral)",
    htmlBody: mensagem
  });
}