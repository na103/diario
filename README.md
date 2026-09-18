# Diario dell'assistenza

Un diario per chi assiste a casa una persona con demenza. Serve a registrare in pochi tocchi com'è andata la notte, gli episodi di agitazione o delirio di giorno e le ore di sonno di chi assiste, e a portare al medico un riepilogo mensile invece che dei ricordi.

Nasce da un foglio di calcolo usato tutti i giorni, di cui mantiene le stesse colonne.

## Cosa registra

**Le notti:** ora in cui la persona è andata a letto, numero e orario dei risvegli, minuti di assistenza, agitazione su una scala da 0 a 3, evacuazione, variazioni di farmaci e note libere. Più le ore di sonno di chi assiste, che sono un dato clinico a tutti gli effetti.

**Gli episodi di giorno:** orario di inizio e fine con durata calcolata, tipo (delirio, allucinazioni, vagabondaggio, caduta e altri), intensità, possibile causa e, soprattutto, cosa ha aiutato.

**La terapia:** un promemoria degli orari delle medicine, con l'elenco dei farmaci da dare a ciascun orario e una nota facoltativa, per esempio "a stomaco pieno". L'orario più vicino al momento in cui apri l'app è evidenziato. È solo un promemoria: non registra le somministrazioni.

**Il riepilogo del mese:** media dei risvegli, minuti medi di assistenza, media delle ore di sonno di chi assiste, notti agitate, episodi e loro durata media. Un grafico mostra le notti del mese, con l'altezza delle barre pari alle ore di sonno di chi assiste, il colore pari all'agitazione e, sotto la linea di base, un pallino nei giorni in cui c'è stata evacuazione. Segnala anche i cali del sonno di chi assiste rispetto al mese precedente e i giorni consecutivi senza evacuazione, che nella demenza è una causa frequente di agitazione.

Da lì si scarica un file Excel con riepilogo, notti, episodi e terapia, da portare alla visita.

## Come si usa

L'app sta in una pagina web. Aprila sul telefono e aggiungila alla schermata Home dal menu del browser, così si comporta come un'app normale e funziona anche senza rete.

I dati restano nella memoria del browser di quel telefono: non passano da nessun server e non c'è nessun account. Questo significa due cose. La prima è che sono solo tuoi. La seconda è che, se cancelli i dati di navigazione o cambi telefono, spariscono, quindi conviene scaricare ogni tanto un backup dall'ingranaggio. Dallo stesso menu si ripristina un backup e si importa un file Excel già esistente.

Usa sempre lo stesso browser: le memorie di browser diversi sono separate, e un diario aperto altrove risulterebbe vuoto.

## Pubblicare l'app

I file vanno nella cartella principale del repository, con GitHub Pages attivo su `main` e cartella `/ (root)`. L'app è poi raggiungibile all'indirizzo delle pagine del repository.

- `index.html` — la pagina
- `app.js` — l'applicazione, già compilata
- `manifest.webmanifest` — nome, colori e icone per l'installazione
- `sw.js` — il service worker, che la tiene disponibile offline
- `icon-192.png`, `icon-512.png`, `icon-maskable-512.png` — le icone

Per aggiornarla basta sostituire `app.js`: l'indirizzo resta lo stesso e i dati sul telefono non si toccano.

## Note

L'app non è un dispositivo medico e non sostituisce il parere del medico. Serve a ricordare e a mostrare cosa è successo davvero, giorno per giorno.
