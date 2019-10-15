# PowerEditExtExamples
Esempi di codice per creare extension PowerEdit

## Extensions PowerMES
* Progetto di tipo class library
* Estensione DLL: «.Addin.dll»
* Un solo addin per DLL
* Referenziare la libreria: 
    * Atys.PowerEDIT.BaseLibrary
* Deve essere creata una classe pubblica
    * Deve implementare l’interfaccia «IPowerEDITAddin» (e «IDisposable»)
    * Deve essere decorata con l’attributo «AddinData»
    * Aggiungere using: Atys.PowerEDIT.Extensibility, Atys.PowerEDIT

La DLL deve essere posizionata in una sotto-cartella del percorso di installazione «C:\Program Files (x86)\Atys\PowerEDIT\Addin», una DLL per ogni sotto-cartella

## Esempi
1. 010_Empty: Struttura di base per la creazione di un addin
2. 015_DocumentManagement: 
    * Creazione e salvataggio
    * Agganciare gli eventi del documento attivo
    * Muoversi tra le righe
    * Ricerca e sostituzione testo
3. 020_UseCases: 
    * Copia del programma in una locazione specifica in base al tipo della macchina
    * Copia record/programmi
    * Controllo trasmissione programmi
    * Colorazione campi in base a condizioni
    * Abilitazione menù e pulsanti
    * Gestione lista utensili