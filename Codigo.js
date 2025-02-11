function verificarValoresUnificado() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Controle de Estoque'); // Aba a ser verificada
  const LIMITE = 50; // Valor limite
  const destinatario = "destino1@email.com", "destino2@email.com"; // E-mail do destinatário
  const urlPlanilha = SpreadsheetApp.getActiveSpreadsheet().getUrl(); // Link da planilha

  // Configuração das células a serem verificadas
  const celulas = [
    { valor: "C5", titulo: "C4" },
    { valor: "E5", titulo: "E4" },
    { valor: "F5", titulo: "F4" }
  ];

  // Variável para acumular os itens com estoque abaixo do limite
  let itensAbaixoDoLimite = [];

  // Iterar sobre as células para verificar os valores
  celulas.forEach(celula => {
    const valor = sheet.getRange(celula.valor).getValue(); // Valor da célula
    const titulo = sheet.getRange(celula.titulo).getValue(); // Título correspondente

    if (valor < LIMITE) { // Se o valor estiver abaixo do limite, adiciona ao array
      itensAbaixoDoLimite.push({
        produto: titulo,
        quantidade: valor
      });
    }
  });

  // Se houver produtos abaixo do limite, monta o e-mail e envia
  if (itensAbaixoDoLimite.length > 0) {
    let listaProdutos = itensAbaixoDoLimite.map(item => `
      <li style="margin-bottom: 10px;">
        <strong>Produto:</strong> ${item.produto} <br>
        <strong>Quantidade Atual:</strong> ${item.quantidade} <br>
        <strong>Limite Mínimo Recomendado:</strong> ${LIMITE}
      </li>
    `).join("");

    // Conteúdo do e-mail em HTML
    const corpoHtml = `
    <body style="background-color: rgb(240, 240, 240); margin: 0; font-family: 'Montserrat', sans-serif; color: #333; padding: 2%;">
    <div style="background-color: white; width: 55%; height: 45%; max-width: 80vw; margin: 10px auto; padding: 20px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);">
        <img src="https://ci3.googleusercontent.com/meips/ADKq_NZzpHFKpi-GITBfP9sR_aNJp0WRhikzh5ZjXBm1lMKk1SkRaBDACIBdY1YQmT65rvt3uh_z0lASo6RhTDFpAlVmztA3LVbyVcsZF2uyGEpUZCWjhyqcxR8VMjNDUHSBkjSESGZTH6KzVJBkdIwVgDHcPF2P-ElFXgikmg=s0-d-e1-ft#https://luqoa.stripocdn.email/content/guids/CABINET_3a833994bb9ef69c4f4eecab6fda1071/images/loggi.png" alt="Logo" style="max-width: 30%; width:20%; margin-bottom: 1%; ">
        <h1 style="text-align: center; color: #e74c3c; font-size: 28px; margin-bottom: 23px;">Alerta de Estoque!</h1>
        <p style="font-size: 16px; text-align: center; ">
            <strong>Os seguintes produtos estão com quantidade baixa:</strong>
        </p>
        <ul style="font-size: 14px; line-height: 24px; padding-left: 0; list-style-type: none; margin: 20px 0;">
            ${listaProdutos}
        </ul>
        <p style="text-align: center; font-size: 14px; margin-top: 16%;">
            Acessa a <a href="${urlPlanilha}" style="color: #3498db; text-decoration: underline;">planilha</a> para mais detalhes.
        </p>
        <img src="https://ci3.googleusercontent.com/meips/ADKq_NZzpHFKpi-GITBfP9sR_aNJp0WRhikzh5ZjXBm1lMKk1SkRaBDACIBdY1YQmT65rvt3uh_z0lASo6RhTDFpAlVmztA3LVbyVcsZF2uyGEpUZCWjhyqcxR8VMjNDUHSBkjSESGZTH6KzVJBkdIwVgDHcPF2P-ElFXgikmg=s0-d-e1-ft#https://luqoa.stripocdn.email/content/guids/CABINET_3a833994bb9ef69c4f4eecab6fda1071/images/loggi.png" alt="Logo" style="max-width: 25%; width:18%; margin-bottom: 1%; margin-top:10%; ">
        <p style="margin:0px;font-size:14px;font-family:&quot;Open Sans&quot;,tahoma,sans-serif;line-height:21px;color:rgb(0,0,0);">
            Atenciosamente, <br>
            <strong>XD Cajamar II
            <br> Security
            </strong>
        </p>
        <p style="font-size: 9px; font-weight: 600; margin-top: 6%; text-align:center;">Mensagem Automâtica, não responder!!</p>
    </div>
</body>
    `;

    // Enviar o e-mail
    GmailApp.sendEmail(destinatario, "Alerta de Estoque!", '', {
      htmlBody: corpoHtml
    });
  }
}

function criarAcionadorSemanal() {
  // Remove qualquer acionador existente para evitar múltiplos acionadores
  const acionadores = ScriptApp.getProjectTriggers();
  for (let i = 0; i < acionadores.length; i++) {
    ScriptApp.deleteTrigger(acionadores[i]);
  }

  // Cria um novo acionador que executa a função `verificarValoresUnificado` toda quarta-feira às 9h
  ScriptApp.newTrigger('verificarValoresUnificado')
    .timeBased()
    .everyWeeks(1) // Executa a cada 1 semana
    .onWeekDay(ScriptApp.WeekDay.WEDNESDAY) // Executa toda quarta-feira
    .atHour(9) // Define a hora (9h)
    .create();
}
