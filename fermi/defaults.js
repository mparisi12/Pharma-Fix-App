// Anagrafica di partenza: usata in modalità demo e finché il Google Sheet non risponde.
// Con il backend attivo, l'app legge Equipment e Liste direttamente dal foglio.
(function () {
  const LINEE = [['ST1MAR', 'ST1 MAR', 1, 1], ['ST1TGM', 'ST1 TGM', 1, 1], ['ST2', 'ST2', 1, 1], ['ORL', 'ORL', 1, 0],
    ['ST3', 'ST3', 1, 1], ['ST4', 'ST4', 1, 1], ['ST5', 'ST5', 1, 0], ['SPL1', 'SPL1', 0, 1], ['SPL2', 'SPL2', 0, 1], ['ST6', 'ST6', 1, 1]];
  const MACC = {
    Sterile: [['RIO', 'Riordinatore'], ['FIL', 'Filling - Tappatrice']],
    Secondario: [['PGR', 'Polmone Grande'], ['ETI', 'Etichettatrice'], ['PPI', 'Polmone Piccolo'],
      ['AST', 'Astucciatrice'], ['BIL', 'Bilancia'], ['CAR', 'Cartonatrice']]
  };
  const equipment = [];
  LINEE.forEach(([cod, nome, f, p]) => {
    if (f) MACC.Sterile.forEach(([m, n]) => equipment.push({ code: cod + '-' + m, linea: nome, area: 'Sterile', macchina: n }));
    if (p) MACC.Secondario.forEach(([m, n]) => equipment.push({ code: cod + '-' + m, linea: nome, area: 'Secondario', macchina: n }));
  });
  window.PF_DEFAULTS = {
    equipment,
    liste: {
      'Categoria': ['MECCANICO', 'ELETTRICO', 'AUTOMATION', 'ALTRO'],
      'Tipo intervento': ['SOSTITUZIONE', 'REGOLAZIONE', 'PULIZIA', 'RIAVVIO', 'ALTRO'],
      'Guasto Sterile': ['H7, Sblocco Homing', 'H7 GUANTONE, Sblocco Homing', 'H32 CTG, rimessa in dima pinze ctg',
        'H32 CPS, rimessa in dima pinze cps, Sostituzione silicone pinze ferma flaconi',
        'H10, Sostituzione e regolazione fotocellula della macchina', 'H19, Rimessa in fase della Stella',
        'G25 Regolazione della coclea flaconi', 'SEA VISION', 'SPIKE CORRENTE', 'ARIA COMPRESSA', 'FORMATO', 'ALTRO'],
      'Guasto Secondario': ['PARAM VIDEOJET', '"0" MACCHINA', 'GUIDE', 'CHIUSURE AST', 'STAMPA DATI VARIABILI',
        'SPIKE CORRENTE', 'POS. TAMPER', 'POS. BOLLINO', 'INCASTRO', 'ARIA COMPRESSA', 'MAG. ISTRUZIONI',
        'MAG. ASTUCCI', 'NASTRATORI', 'VOMERE', 'CAM. OMRON', 'CAM. ANTARES', 'FORMATO',
        'REGOLAZIONE DA PANNELLO', 'USURA', 'SENSORE', 'PULIZIA', 'SOSTITUZIONE', 'RIAVVIO', 'ALTRO'],
      'Manutentori': ['ARACHIDE Simone', 'CACIOPPO Riccardo', 'CRISPO Mattia', 'DI RUSCIO Christian',
        'DI STEFANO Alessio', 'GRECO Nicolò', 'GRITTI Alessandro', 'GUIDONE Andrea', 'LAMANUZZI Vincenzo',
        'MANGIAROTTI Marco', 'MORABITO Giulio', 'ORZI Matteo', 'PALUMBO Antonino', 'PARISI Mirko',
        'PESTRICHELLA Mattia', 'RONCONI Fabio', 'SALVINI Alvaro', 'TATA Giuseppe']
    }
  };
})();
