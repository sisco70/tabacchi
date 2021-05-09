# Tabacchi

Tabacchi è una applicazione Python 3 basata su GTK3 dedicata alla gestione del magazzino di una tabaccheria e all'invio degli ordini
di prodotti di monopolio tramite il portale Logista in modo automatizzato.  
Per poter utilizzare questa funzionalità è necessaria l'autorizzazione alla rivendita di generi di monopolio e il possesso delle credenziali per l'accesso al portale rivenditori.  
  
Ho sviluppato questo software per semplificare la gestione della mia attività senza fini di lucro.

[![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](https://opensource.org/licenses/Apache-2.0)


#### Build
`poetry build`

#### Install (Linux Ubuntu 20.04LTS)
```
sudo apt install pkg-config python3-gi-cairo libcairo2-dev libgirepository1.0-dev libbluetooth-dev python3-testresources
pip install tabacchi-0.3.1.tar.gz
```

#### Run
`tabacchi`