function ordenarAbasComBasePorUltimo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  // Nomes das abas
  const nomes = sheets.map(s => s.getName());

  // Separa abas BASE_ das demais
  const abasBase = nomes.filter(n => n.toUpperCase().startsWith("BASE_"));
  const abasNormais = nomes.filter(n => !n.toUpperCase().startsWith("BASE_"));

  // Ordena cada grupo alfabeticamente
  abasNormais.sort((a, b) => a.localeCompare(b, 'pt-BR'));
  abasBase.sort((a, b) => a.localeCompare(b, 'pt-BR'));

  // Junta tudo: normais primeiro, BASE_ por último
  const ordemFinal = [...abasNormais, ...abasBase];

  // Reorganiza as abas
  ordemFinal.forEach((nome, index) => {
    ss.setActiveSheet(ss.getSheetByName(nome));
    ss.moveActiveSheet(index + 1);
  });
}


