\# Script potrivire masculi - femele



\## Despre proiect



Acest script în Python automatizează procesul de potrivire între masculi și femele pe baza unor fișiere CSV, folosind reguli de validare, filtrare și prioritizare.



Scopul lui este să genereze rapid combinații valide, să elimine potrivirile nepermise și să exporte rezultatele finale atât în format CSV, cât și în fișiere Excel separate pentru fiecare controlor.



---



\## Funcționalități



Scriptul realizează următoarele operații:



\- citește fișierele sursă pentru masculi și femele;

\- normalizează valorile textuale:

&nbsp; - diacritice;

&nbsp; - denumiri de rasă;

&nbsp; - secțiuni;

\- convertește codurile de culoare în valori numerice;

\- generează combinațiile posibile între masculi și femele din același crescător;

\- aplică filtre genetice pentru eliminarea combinațiilor incompatibile;

\- aplică regula specială pentru femelele din categoria \*\*METIS\*\*;

\- prioritizează combinațiile după:

&nbsp; - tipul rasei (non-METIS înainte de METIS);

&nbsp; - culoarea femelei;

&nbsp; - secțiunea femelei;

&nbsp; - culoarea masculului;

\- face balansarea numărului de combinații atribuite fiecărui mascul;

\- exportă rezultatele finale în fișiere CSV;

\- generează fișiere Excel separate pentru fiecare controlor;

\- colorează automat celulele relevante în Excel în funcție de codul culorii.



---



\## Fișiere de intrare



Scriptul folosește două fișiere CSV:



\- `MASCULI BND - RALUCA.csv`

\- `FEMELE PENTRU POTRIVIRE - RALUCA.csv`



Acestea trebuie să existe în directorul configurat în variabila `BASE`.



---



\## Fișiere generate



La rulare, scriptul produce:



\- `POTRIVIRE.csv`  

&nbsp; conține toate combinațiile valide după filtrare;



\- `COMBINARE\_CULOARE\_FEMELE.csv`  

&nbsp; conține combinațiile finale selectate pe baza priorităților definite;



\- fișiere Excel individuale în folderul `raluca/`  

&nbsp; câte un fișier pentru fiecare controlor, cu marcarea vizuală a culorilor.



---



\## Reguli principale de potrivire



\### 1. Filtrare genetică

Sunt eliminate combinațiile în care apar incompatibilități între:

\- mascul și tatăl femelei;

\- mama masculului și mama femelei;

\- mama masculului și femela însăși;

\- tatăl masculului și tatăl femelei.



\### 2. Regula de rasă

O combinație este acceptată dacă:

\- masculul și femela au aceeași rasă, sau

\- femela este `METIS`, iar crescătorul are exact o singură rasă non-METIS în lista de femele.



\### 3. Prioritizare

Ordinea de selecție este:

1\. femele non-METIS înainte de METIS;

2\. culoarea femelei;

3\. secțiunea femelei;

4\. culoarea masculului;

5\. balansarea masculilor cu mai puține combinații.



\### 4. Limitare per mascul

Se aplică:

\- \*\*soft cap\*\* = 50 combinații;

\- \*\*hard cap\*\* = 60 combinații.



---



\## Coduri de culoare



În script, culorile sunt mapate astfel:



\- `1` = albastru

\- `2` = portocaliu / auriu

\- `3` = verde

\- `4` = fără culoare



Valorile pot veni din CSV și sub formă:

\- `A` / `a`

\- `P` / `p`

\- `V` / `v`

\- `1`, `2`, `3`



---



\## Tehnologii folosite



\- Python

\- pandas

\- openpyxl

\- os

\- re



---



\## Configurare



Căile fișierelor sunt definite direct în cod, în secțiunea:



```python

BASE = r'C:\\Users\\capri\\Desktop\\SITUATIE PENTRU PERECHI\\potrivire'

