# Diario dell'assistenza

Un diario per chi assiste a casa una persona con demenza. Serve a registrare in pochi tocchi com'è andata la notte, gli episodi di agitazione o delirio di giorno e le ore di sonno di chi assiste, e a portare al medico un riepilogo mensile invece che dei ricordi.

Nasce da un foglio di calcolo usato tutti i giorni, di cui mantiene le stesse colonne.

## Cosa registra

**Le notti:** ora in cui la persona è andata a letto, numero e orario dei risvegli, minuti di assistenza, agitazione su una scala da 0 a 3, evacuazione, variazioni di farmaci e note libere. Più le ore di sonno di chi assiste, che sono un dato clinico a tutti gli effetti.

**Gli episodi di giorno:** orario di inizio e fine con durata calcolata, tipo (delirio, allucinazioni, vagabondaggio, caduta e altri), intensità, possibile causa e, soprattutto, cosa ha aiutato.

**La terapia:** un promemoria degli orari delle medicine, con l'elenco dei farmaci da dare a ciascun orario e una nota facoltativa, per esempio "a stomaco pieno". L'orario più vicino al momento in cui apri l'app è evidenziato. È solo un promemoria: non registra le somministrazioni.

**La glicemia:** le misurazioni fatte con il glucometro, con valore, ora e momento della giornata (a digiuno, prima o dopo il pasto, sera) e una nota.

**La pressione:** nella stessa scheda della glicemia. Si inseriscono due misurazioni, con massima, minima e battiti, e l'app ne calcola la media, colorata di verde, giallo o rosso secondo i valori di riferimento per la misurazione a casa: fino a 135/85 è nella norma, da 160/100 è alta, sotto 90 di massima è bassa.

**Il riepilogo del mese:** media dei risvegli, minuti medi di assistenza, media delle ore di sonno di chi assiste, notti agitate, episodi e loro durata media. Un grafico mostra le notti del mese, con l'altezza delle barre pari alle ore di sonno di chi assiste, il colore pari all'agitazione e, sotto la linea di base, un pallino nei giorni in cui c'è stata evacuazione. Segnala anche i cali del sonno di chi assiste rispetto al mese precedente e i giorni consecutivi senza evacuazione, che nella demenza è una causa frequente di agitazione.

Un secondo grafico, con la stessa scala dei giorni, mette le misurazioni della glicemia sopra i giorni in cui ci sono stati episodi, così le due cose si leggono in colonna e si vede se cadono insieme.

Un terzo grafico mostra la pressione, una barra per giorno dalla minima alla massima, con la media del mese.

Da lì si scarica un file Excel con riepilogo, notti, episodi, glicemia, pressione e terapia, da portare alla visita.

## Come si usa

L'app sta in una pagina web. Aprila sul telefono e aggiungila alla schermata Home dal menu del browser, così si comporta come un'app normale e funziona anche senza rete.

I dati restano nella memoria del browser di quel telefono: non passano da nessun server e non c'è nessun account. Questo significa due cose. La prima è che sono solo tuoi. La seconda è che, se cancelli i dati di navigazione o cambi telefono, spariscono, quindi conviene scaricare ogni tanto un backup dall'ingranaggio. Dallo stesso menu si ripristina un backup e si importa un file Excel già esistente.

Usa sempre lo stesso browser: le memorie di browser diversi sono separate, e un diario aperto altrove risulterebbe vuoto.

## Pubblicare l'app

I file vanno nella cartella principale del repository, con GitHub Pages attivo su `main` e cartella `/ (root)`. L'app è poi raggiungibile all'indirizzo delle pagine del repository.

- `index.html` — la pagina
- `app.js` — l'applicazione, già compilata (generata da `src/`, non si modifica a mano)
- `manifest.webmanifest` — nome, colori e icone per l'installazione
- `sw.js` — il service worker, che la tiene disponibile offline
- `icon-192.png`, `icon-512.png`, `icon-maskable-512.png` — le icone

Per aggiornarla basta sostituire `app.js`: l'indirizzo resta lo stesso e i dati sul telefono non si toccano.

## Modificare l'app

Il codice leggibile sta in `src/`: `app.jsx` è l'applicazione intera (un solo file, React senza framework) e `main.jsx` la aggancia alla pagina. `app.js` nella cartella principale è il risultato della compilazione, quindi va rigenerato e non corretto a mano.

```
npm install
npm run build     # riscrive app.js
npm run watch     # ricompila a ogni salvataggio, mentre si lavora
npm run serve     # http://localhost:8080 per provare in locale
```

Le versioni delle librerie sono fissate in `package.json`: con queste, `npm run build` produce esattamente l'`app.js` pubblicato. Dopo la compilazione si committano sia `src/` sia `app.js`, perché GitHub Pages serve il file compilato così com'è.

Il service worker tiene una copia dei file per farli funzionare offline, quindi dopo un aggiornamento può servire riaprire l'app una volta, o svuotare la cache del browser (solo "immagini e file memorizzati nella cache", mai "cookie e dati dei siti", che cancellerebbe il diario).

## Licenza

Il codice è rilasciato sotto la [GNU General Public License v3.0](LICENSE) o successiva: puoi usarlo, modificarlo e ridistribuirlo, a patto che anche le versioni modificate restino libere, con il sorgente disponibile e la stessa licenza.

## Note

L'app non è un dispositivo medico e non sostituisce il parere del medico. Serve a ricordare e a mostrare cosa è successo davvero, giorno per giorno.

---

Se ti piace, considera di supportarmi con un caffè ☕ [ko-fi.com/na103](https://ko-fi.com/na103)
