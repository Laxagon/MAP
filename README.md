# Mail Automation Project (MAP)

<img src="automatic-mail.avif" width="400" height="400">

### Problem
Salahaddin skole er en kurdisk lørdagsskole i Oslo som drives av Salahaddin Senter. Jeg oppdaget at de sendte elevenes timeplan til foreldrene på e-mail som var en tidskrevende prosess hver uke.

### Min oppgave
Jeg kom med forslaget om å automatisere hele prossessen for dem frivillig, og de aksepterte. Dermed startet oppgaven om  å utvikle en automatisk e-mail script som også skulle være brukervennlig for skolens admin å bruke.

### Hvordan jeg løste oppgaven
Jeg leste meg opp på e-mail automasjon på Python og begynte å implementere det jeg leste om. Det var 3 ting jeg merket var veldig viktig å passe på:
1. Alle e-mail og brukeropplysningerer privat sensitiv informasjon, så dette skulle bli behandlet forsiktig.
2. Skriving av tester som forsikrer at foreldrene unngår spam eller mangel av timeplan. Automasjonen skulle dobbeltsjekke at alt blir gjort riktig.
3. Sørge for at MAP-scriptet skulle lagres som en .exe fil hos admin hos Salahaddin Senter slik at han bare ved et klikk kunne sende alle mailene. Han fikk også veiledning av meg over hvordan han kunne legge til, endre eller slette databasen med alle foreldrene og e-mailene.

### Resultat
MAP ble en success! Helt siden MAP ble utviklet i januar 2024, har Salahaddin Senter spart ***4 timer i måneden***. I dag, <!-- DATE_START -->06. October 2026<!-- DATE_END -->, utgjør dette totalt omtrent ***<!-- WEEKS_START -->91<!-- WEEKS_END --> timer***.

### Refleksjon

Da MAP var ferdig utviklet, var det noen svakheter ved løsningen, selv om den fungerte som den skulle. MAP var ikke så brukervennlig som jeg hadde håpet, og databasen var tungvint å redigere. Jeg har senere gått tilbake og gjort flere endringer som har gjort både databasen og brukervennligheten bedre, men det er fortsatt ikke perfekt.

I fremtidige prosjekter har jeg derfor lært at jeg må legge mer vekt på disse delene av utviklingen. Jeg vil ta med meg erfaringene fra dette prosjektet og være mer bevisst på hva som fungerer godt og hva som kan forbedres. Dette er spesielt viktig fordi det er disse delene brukerne faktisk kommer mest i kontakt med.

