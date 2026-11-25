# Tirocinio-BT-Unifg

Prototipo di portale per organizzare i tirocini del **Corso di Laurea in Infermieristica** negli ospedali della provincia BAT:
- Ospedale "Mons. R. Dimiccoli" di Barletta;
- Ospedale "L. Bonomo" di Andria;
- Ospedale "Vittorio Emanuele II" di Bisceglie.

È una demo statica: una pagina HTML senza server e senza database. Quello che si modifica resta in memoria fino al ricaricamento.

## Cosa fa
- **Area coordinatore**, divisa per anno di corso (I, II, III anno):
  - panoramica;
  - assegnazione dei reparti in 3 passi (automatica o manuale);
  - gestione studenti;
  - reparti con Guide e posti;
  - presenze con QR e monte ore annuale;
  - messaggi.
- **Area studente**: reparto assegnato e motivazione, percorso, timbratura con QR, calendario dei turni, messaggi alla coordinatrice.

Le regole (assegnazione, rapporto Guida-studente 1:2, debito orario del 10%, ecc.) seguono il documento *Principi e standard del tirocinio professionale* (Conferenza Permanente delle Professioni Sanitarie, 2025). Il riassunto è in [`docs/fonti-tirocinio.md`](docs/fonti-tirocinio.md).

## Dati
- **I e III anno**: studenti e ore reali dal registro ufficiale. I nomi sono ridotti a *Cognome N.* e le matricole oscurate (es. `6342**`). Media, domicilio e tirocini svolti non sono nel registro e si completano dalla sezione Studenti.
- **II anno**: studenti di esempio inventati.
- Posti, tutor, presenze del periodo e messaggi sono di esempio.

Per aggiornare i dati reali dal file Excel (serve Excel installato):

```powershell
powershell -File tools\importa-ore.ps1 -Origine "C:\percorso\ore_tirocinio_studenti.xlsx"
```

## Provarlo in locale

```bash
python -m http.server 8765
```

Poi apri http://localhost:8765.
