function onEdit(e) {
  // Verifica se o evento é válido
  if (!e || !e.source || !e.range) {
    Logger.log('Evento de edição inválido.');
    return;
  }

  var sheet = e.source.getActiveSheet();
  var range = e.range;

  // Verifica se a edição ocorreu na aba 'CONTROLE' e na coluna do Menu Suspenso (coluna C)
  if (sheet.getName() === 'CONTROLE' && range.getColumn() === 3) {
    var menuSelecionado = range.getValue();
    var row = range.getRow();
    var dataInicioCell = sheet.getRange(row, 4); // Coluna D: Data e Hora de Início
    var dataFimCell = sheet.getRange(row, 5);    // Coluna E: Data e Hora de Fim
    var duracaoCell = sheet.getRange(row, 6);    // Coluna F: Duração
    var statusCell = sheet.getRange(row, 7);     // Coluna G: Status

    // Cria a aba 'TRAMITAÇÃO' se não existir
    var sheetTramitacao = e.source.getSheetByName('TRAMITAÇÃO');
    if (!sheetTramitacao) {
      sheetTramitacao = e.source.insertSheet('TRAMITAÇÃO');
    }

    // Adiciona um cabeçalho se necessário
    if (sheetTramitacao.getRange(1, 1).getValue() !== 'Menu' &&
        sheetTramitacao.getRange(1, 2).getValue() !== 'Data Início 1' && sheetTramitacao.getRange(1, 3).getValue() !== 'Data Fim 1' && sheetTramitacao.getRange(1, 4).getValue() !== 'Duração 1' &&
        sheetTramitacao.getRange(1, 5).getValue() !== 'Data Início 2' && sheetTramitacao.getRange(1, 6).getValue() !== 'Data Fim 2' && sheetTramitacao.getRange(1, 7).getValue() !== 'Duração 2' &&
        sheetTramitacao.getRange(1, 8).getValue() !== 'Data Início 3' && sheetTramitacao.getRange(1, 9).getValue() !== 'Data Fim 3' && sheetTramitacao.getRange(1, 10).getValue() !== 'Duração 3') {
      sheetTramitacao.getRange(1, 1).setValue('Menu');
      sheetTramitacao.getRange(1, 2).setValue('Data Início 1');
      sheetTramitacao.getRange(1, 3).setValue('Data Fim 1');
      sheetTramitacao.getRange(1, 4).setValue('Duração 1');
      sheetTramitacao.getRange(1, 5).setValue('Data Início 2');
      sheetTramitacao.getRange(1, 6).setValue('Data Fim 2');
      sheetTramitacao.getRange(1, 7).setValue('Duração 2');
      sheetTramitacao.getRange(1, 8).setValue('Data Início 3');
      sheetTramitacao.getRange(1, 9).setValue('Data Fim 3');
      sheetTramitacao.getRange(1, 10).setValue('Duração 3');
      // Adicionar cabeçalhos para mais etapas se necessário
    }

    // Adiciona os dados da etapa anterior na aba 'TRAMITAÇÃO' se houver uma etapa anterior
    if (e.oldValue && e.oldValue !== menuSelecionado) {
      // Se Data de Fim está vazia, defina a Data e Hora de Fim para o momento atual
      if (dataFimCell.getValue() === '') {
        var dataFim = new Date();
        dataFim = ajustarFusoHorario(dataFim, -3); // Ajusta para GMT-3 (Brasília)
        dataFimCell.setValue(dataFim);

        // Calcula a duração da etapa anterior
        var dataInicio = new Date(dataInicioCell.getValue());
        var diffMillis = dataFim - dataInicio;
        var diffDias = Math.floor(diffMillis / (1000 * 60 * 60 * 24));
        var diffHoras = Math.floor((diffMillis % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60));
        var diffMinutos = Math.floor((diffMillis % (1000 * 60 * 60)) / (1000 * 60));
        var duracaoTexto = diffDias + ' dias, ' + diffHoras + ' horas, ' + diffMinutos + ' minutos';
        duracaoCell.setValue(duracaoTexto);

        // Adiciona os dados da etapa anterior na aba 'TRAMITAÇÃO'
        var colInicio = 2 + ((parseInt(e.oldValue) - 1) * 3);
        var colFim = colInicio + 1;
        var colDuracao = colInicio + 2;

        sheetTramitacao.getRange(row, 1).setValue(e.oldValue);
        sheetTramitacao.getRange(row, colInicio).setValue(dataInicio);
        sheetTramitacao.getRange(row, colFim).setValue(dataFim);
        sheetTramitacao.getRange(row, colDuracao).setValue(duracaoTexto);
      }

      // Limpa os campos da etapa anterior e define a Data e Hora de Início para a nova etapa
      dataInicioCell.setValue('');
      dataFimCell.setValue('');
      duracaoCell.setValue('');
      statusCell.setValue('');

      // Define a Data e Hora de Início para o momento atual para a nova etapa
      var dataInicio = new Date();
      dataInicio = ajustarFusoHorario(dataInicio, -3); // Ajusta para GMT-3 (Brasília)
      dataInicioCell.setValue(dataInicio);

      // Atualiza o Status para 'Ativo'
      statusCell.setValue('Ativo');
    } else if (!e.oldValue) {
      // Se não há uma etapa antiga, define a Data e Hora de Início para o momento atual para a primeira etapa
      var dataInicio = new Date();
      dataInicio = ajustarFusoHorario(dataInicio, -3); // Ajusta para GMT-3 (Brasília)
      dataInicioCell.setValue(dataInicio);

      // Atualiza o Status para 'Ativo'
      statusCell.setValue('Ativo');
    } else {
      // Se o menu suspenso está vazio, limpa os campos de Data e Hora e Duração
      dataInicioCell.setValue('');
      dataFimCell.setValue('');
      duracaoCell.setValue('');
      statusCell.setValue('');
    }
  } else if (sheet.getName() === 'CONTROLE' && range.getColumn() === 6) {
    // Verifica se a edição ocorreu na aba 'CONTROLE' e na coluna da Duração (coluna F)
    var duracao = range.getValue();
    var row = range.getRow();
    var dataInicioCell = sheet.getRange(row, 4); // Coluna D: Data e Hora de Início
    var dataFimCell = sheet.getRange(row, 5);    // Coluna E: Data e Hora de Fim

    if (duracao && !isNaN(duracao) && dataInicioCell.getValue()) {
      // Calcula a data de fim com base na duração definida (duração em dias)
      var dataInicio = new Date(dataInicioCell.getValue());
      var dataFim = new Date(dataInicio.getTime() + (duracao * 24 * 60 * 60 * 1000));
      dataFim = ajustarFusoHorario(dataFim, -3); // Ajusta para GMT-3 (Brasília)
      dataFimCell.setValue(dataFim);
    }
  }
}

// Função para ajustar a data para um fuso horário específico
function ajustarFusoHorario(data, offsetHoras) {
  var fuso = offsetHoras * 60 * 60 * 1000;
  return new Date(data.getTime() + fuso);
}
