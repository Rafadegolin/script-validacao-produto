function plotarProdutoNaListagem() {
  const IN = 'Adicionar produtos para análise';
  const OUT = 'Listagem';

  const ss = SpreadsheetApp.getActive();
  const inS = ss.getSheetByName(IN);
  const outS = ss.getSheetByName(OUT);
  if (!inS || !outS) throw new Error('Guia não encontrada.');

  // próxima linha vazia considerando apenas um bloco de colunas
  function nextEmptyRow_(sheet, startRow, startCol, width) {
    const max = sheet.getMaxRows();
    const n = Math.max(max - startRow + 1, 1);
    const vals = sheet.getRange(startRow, startCol, n, width).getValues();
    for (let i = 0; i < vals.length; i++) {
      if (vals[i].every(v => v === '' || v === null)) return startRow + i;
    }
    return max + 1;
  }

  // lê canto superior-esquerdo do bloco; se asLink, extrai URL
  function readBlock(a1, { asLink = false } = {}) {
    const tl = inS.getRange(a1).getCell(1, 1);
    if (asLink) {
      const rt = tl.getRichTextValue();
      if (rt) {
        const d = rt.getLinkUrl(); if (d) return d;
        for (const r of rt.getRuns()) { const u = r.getLinkUrl(); if (u) return u; }
      }
    }
    return tl.getValue();
  }

  // Lê os dados
  const produto = readBlock('D3:E4');
  const novoCampo = readBlock('H13:I14'); // Para nota e P:Q
  
  // Array mantém 11 colunas (A:K) - NÃO inclui novoCampo
  const linha = [
    produto,                    // A  Produto
    readBlock('H3:I4'),         // B  Catálogo/normal
    readBlock('D5:E6'),         // C  Concorrência
    readBlock('H5:I6'),         // D  Faturamento
    readBlock('D7:E8'),         // E  Preço concorrente
    readBlock('H7:I8'),         // F  Data/dias
    readBlock('D9:E10', {asLink:true}), // G  Link
    readBlock('H9:I10'),        // H  Preço custo
    readBlock('D11:E12'),       // I  Qtd mínima
    readBlock('H11:I12'),       // J  Preço venda
    readBlock('D13:E14')        // K  Margem %
  ];

  if (linha.every(v => String(v ?? '').trim() === '')) {
    ss.toast('Nada para lançar. Preencha os campos.');
    return;
  }

  // Grava em A:K (11 colunas) - coluna L fica livre para especialista
  const WIDTH_OUT = 11;
  let target = nextEmptyRow_(outS, 2, 1, WIDTH_OUT);
  if (target > outS.getMaxRows()) outS.insertRowsAfter(outS.getMaxRows(), target - outS.getMaxRows());
  const writeRange = outS.getRange(target, 1, 1, WIDTH_OUT);
  if (writeRange.isPartOfMerge()) writeRange.breakApart();
  writeRange.setValues([linha]);

  // ADICIONA NOTA na coluna A com o conteúdo de H13:I14
  if (novoCampo && String(novoCampo).trim() !== '') {
    const cellA = outS.getRange(target, 1);
    cellA.setNote(String(novoCampo));
  }

  // lista lateral L:O e observações P:Q na aba de entrada
  atualizarListaLateral_(inS, produto, novoCampo);

  // Limpa blocos incluindo o novo campo H13:I14
  inS.getRangeList([
    'D3:E4','H3:I4',
    'D5:E6','H5:I6',
    'D7:E8','H7:I8',
    'D9:E10','H9:I10',
    'D11:E12','H11:I12',
    'D13:E14','H13:I14'
  ]).clearContent();

  ss.toast('Produto lançado e listado.');
}

// escreve o produto em L:O e observações em P:Q na aba de entrada
function atualizarListaLateral_(sheet, produto, observacao) {
  const START_ROW = 3, START_COL = 12, WIDTH = 4; // L:O
  const max = sheet.getMaxRows();
  const n = Math.max(max - START_ROW + 1, 1);
  const vals = sheet.getRange(START_ROW, START_COL, n, WIDTH).getValues();

  let row = START_ROW;
  let found = false;
  for (let i = 0; i < vals.length; i++) {
    if (vals[i].every(v => v === '' || v === null)) { row = START_ROW + i; found = true; break; }
  }
  if (!found) row = max + 1;
  if (row > max) sheet.insertRowsAfter(max, row - max);

  // Grava produto em L:O
  const rProduto = sheet.getRange(row, START_COL, 1, WIDTH); // L:O
  if (rProduto.isPartOfMerge()) rProduto.breakApart();
  rProduto.merge();
  rProduto.setValue(produto);

  // Grava observação em P:Q
  if (observacao && String(observacao).trim() !== '') {
    const rObs = sheet.getRange(row, 16, 1, 2); // P:Q (colunas 16-17)
    if (rObs.isPartOfMerge()) rObs.breakApart();
    rObs.merge();
    rObs.setValue(observacao);
  }
}

// SINCRONIZA VALIDAÇÃO - lê L:M da Listagem e escreve em R:T do resumo
function validarProdutos() {
  const IN = 'Adicionar produtos para análise';
  const OUT = 'Listagem';

  const ss  = SpreadsheetApp.getActive();
  const inS = ss.getSheetByName(IN);
  const outS= ss.getSheetByName(OUT);
  if (!inS || !outS) throw new Error('Guia não encontrada.');

  // conta itens em A (contíguos)
  const aCol = outS.getRange(2, 1, Math.max(outS.getMaxRows() - 1, 1), 1).getDisplayValues();
  let listagemCount = 0;
  for (let i = 0; i < aCol.length; i++) { if (String(aCol[i][0]).trim() === '') break; listagemCount++; }
  if (listagemCount === 0) { ss.toast('Nada para validar.'); return; }

  // conta itens na lista lateral L da aba de entrada
  const lCol = inS.getRange(3, 12, Math.max(inS.getMaxRows() - 2, 1), 1).getDisplayValues();
  let lateralCount = 0;
  for (let i = 0; i < lCol.length; i++) { if (String(lCol[i][0]).trim() === '') break; lateralCount++; }

  const total = Math.min(listagemCount, lateralCount);
  if (total === 0) { ss.toast('Sem correspondência com a lista lateral.'); return; }

  // LÊ L:M (colunas 12:13) da aba LISTAGEM
  const VAL_START = 12; // L
  const validacoes = outS.getRange(2, VAL_START, total, 2).getDisplayValues(); // L,M

  // cores
  const STATUS_COLOR = { positiva: '#18c700', negativa: '#f13c3c', neutra: '#ffad55' };
  const buildRich_ = (text) => {
    const b = SpreadsheetApp.newRichTextValue().setText(text);
    const lower = text.toLowerCase();
    for (const [word, color] of Object.entries(STATUS_COLOR)) {
      let idx = 0;
      while ((idx = lower.indexOf(word, idx)) !== -1) {
        const st = SpreadsheetApp.newTextStyle().setForegroundColor(color).build();
        b.setTextStyle(idx, idx + word.length, st);
        idx += word.length;
      }
    }
    return b.build();
  };

  // ESCREVE em R:T (colunas 18:20) da aba de entrada
  for (let i = 0; i < total; i++) {
    const comentario = (validacoes[i][0] || '').toString().trim(); // L (comentário)
    const status     = (validacoes[i][1] || '').toString().trim(); // M (status)
    
    const texto = (status && comentario) ? `${status} - ${comentario}` : (status || comentario);

    const row = 3 + i;
    const rng = inS.getRange(row, 18, 1, 3); // R:T (colunas 18:20)
    if (rng.isPartOfMerge()) rng.breakApart();
    rng.merge();

    const cell = inS.getRange(row, 18); // R
    if (texto) cell.setRichTextValue(buildRich_(texto));
    else cell.setValue('');
  }

  ss.toast('Validações sincronizadas.');
}
