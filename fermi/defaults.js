// Anagrafica di partenza: SOLO dati di esempio, per la modalità demo e le liste finché il foglio non risponde.
// Con il backend attivo, l'app legge Equipment e Liste dal foglio (privato). Qui non vanno mai dati aziendali reali.
(function () {
  // Macchine DI ESEMPIO, usate solo in modalità demo (senza API_URL).
  // L'elenco vero sta solo nel foglio Equipment (privato): l'app lo riceve dopo l'accesso con il codice reparto.
  const equipment = [
    { code: "DEMO-A-01", linea: "LINEA A", area: "Filling", macchina: "RIEMPITRICE" },
    { code: "DEMO-A-02", linea: "LINEA A", area: "Packaging", macchina: "ETICHETTATRICE" },
    { code: "DEMO-A-03", linea: "LINEA A", area: "Packaging", macchina: "ASTUCCIATRICE" },
    { code: "DEMO-A-04", linea: "LINEA A", area: "Packaging", macchina: "CARTONATRICE" },
    { code: "DEMO-B-01", linea: "LINEA B", area: "Packaging", macchina: "ASTUCCIATRICE" },
    { code: "DEMO-B-02", linea: "LINEA B", area: "Packaging", macchina: "BILANCIA" }
  ];
  window.PF_DEFAULTS = {
    equipment,
    liste: {
      'Categoria': ['MECCANICO', 'ELETTRICO', 'AUTOMATION', 'ALTRO'],
      'Tipo intervento': ['REGOLAZIONE', 'SOSTITUZIONE', 'RIPRISTINO', 'PULIZIA'],
      'Guasto Sterile': ['SENSORE', 'FORMATO', 'SPIKE CORRENTE', 'ARIA COMPRESSA', 'ALTRO'], // le voci vere stanno nel foglio Liste
      'Guasto Secondario': ['GUIDE', 'INCASTRO', 'FORMATO', 'REGOLAZIONE DA PANNELLO', 'USURA', 'SENSORE', 'PULIZIA',
        'SOSTITUZIONE', 'RIAVVIO', 'SPIKE CORRENTE', 'ARIA COMPRESSA', 'ALTRO'], // le voci vere stanno nel foglio Liste
      'Manutentori': [] // i nomi stanno solo nel foglio Liste (privato)
    }
  };
})();
