# Gestionale di Tirocinio

Prototipo di portale per organizzare i tirocini del **Corso di Laurea in Infermieristica** negli ospedali della provincia BAT:
- Ospedale "Mons. R. Dimiccoli" di Barletta;
- Ospedale "L. Bonomo" di Andria;
- Ospedale "Vittorio Emanuele II" di Bisceglie.

È una demo statica: una pagina HTML senza server e senza database. Quello che si modifica resta in memoria fino al ricaricamento.

## Cosa fa
- **Schermata di accesso** con due profili, Coordinatore (Michele Massaro) e Studente, ognuno con la sua homepage. Per ora l'accesso è solo dimostrativo: la password non viene verificata e il metodo definitivo è da decidere.
- **Area coordinatore**, divisa per anno di corso (I, II, III anno):
  - panoramica;
  - assegnazione dei reparti in 3 passi (automatica o manuale), nell'ordine delle linee guida 2025: requisiti obbligatori, poi bisogni di apprendimento, poi posti e Guide, infine preferenze dello studente. La media (inserita dal coordinatore) serve solo per i pari merito;
  - calendario dei periodi e delle sessioni d'esame, con i controlli della Consensus 2025: 4–6 settimane, niente sovrapposizioni con gli esami, 5–6 esperienze nel triennio;
  - livelli di complessità dei reparti (1 base, 2 intermedio, 3 area critica): il I anno va solo al livello 1, il II fino al 2, il III fino al 3. La classificazione è una proposta da validare;
  - requisiti per studente (visita di sorveglianza sanitaria, corso sicurezza, laboratori propedeutici): chi non li ha non viene assegnato;
  - gestione studenti;
  - reparti con Guide e posti;
  - presenze con QR e monte ore annuale;
  - messaggi.
- **Area studente**: requisiti mancanti, preferenze (fino a 3 scelte, "non ho l'auto", altre esigenze), reparto assegnato e motivazione, percorso, timbratura con QR, calendario dei turni, messaggi al coordinatore.

Le regole (assegnazione, rapporto Guida-studente 1:2, debito orario del 10%, ecc.) seguono il documento *Principi e standard del tirocinio professionale* (Conferenza Permanente delle Professioni Sanitarie, 2025). Il riassunto è in [`docs/fonti-tirocinio.md`](docs/fonti-tirocinio.md).

## Dati
- **I e III anno**: studenti reali dal registro ufficiale. I nomi sono ridotti a *Cognome N.* e le matricole oscurate (es. `6342**`). Media, domicilio e tirocini svolti non sono nel registro e si completano dalla sezione Studenti.
- **II anno**: studenti di esempio inventati.
- **Monte ore e presenze azzerati** per tutti: le ore si accumulano dalle timbrature del periodo, che parte il 24/11/2026. I turni passati si inseriranno a mano più avanti.
- Posti, tutor e messaggi sono di esempio.

Per aggiornare i dati reali dal file Excel (serve Excel installato):

```powershell
powershell -File tools\importa-ore.ps1 -Origine "C:\percorso\ore_tirocinio_studenti.xlsx"
```

## Provarlo in locale

```bash
python -m http.server 8765
```

Poi apri http://localhost:8765.
