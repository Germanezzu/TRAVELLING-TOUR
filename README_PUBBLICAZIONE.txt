TRAVELLING TOUR - SITO ONLINE

Questa versione usa: index.html + data.js + assets.

DATI
I dati iniziali sono stati estratti dal file Classifica Travelling Tour(1).xlsm.
Ultima tappa rilevata: 41. Giocatori: 106.

LOGHI
- image3.png = logo Travelling Tour
- image2.png = logo Pesaro Padel
- image1.png / image4.png = sfondi grafici

AGGIORNAMENTO AUTOMATICO
Il file AggiornaSitoTravellingTour.bas contiene la macro che legge la scheda Classifica e pubblica data.js su GitHub.
Dopo la configurazione iniziale, premendo AggiornaSitoTravellingTour nel file Excel il sito viene aggiornato.

PUBBLICAZIONE
La soluzione consigliata e GitHub Pages: il repository deve contenere index.html, data.js e assets/.
Attivare Settings > Pages > Deploy from branch > main > /(root).
Il link pubblico sara del tipo https://USERNAME.github.io/travelling-tour/

TOKEN
Per sicurezza la macro legge il token dalla variabile ambiente Windows TRAVELLING_GITHUB_TOKEN e non lo salva nel file Excel.
