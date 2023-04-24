Matrix-Systeme für NIR-Scheinwerfer:

Wenn ein eigenens NIR-Startsystem mit Matrix Strahlquelle verwendet werden soll, dann muss ein 4 Linsen System erstellt werden, 
dass als Linse 3 und 4 Platzhalter Linsen eingefügt hat. Diese werden dann über Set_NIR_System_V5 gelöscht und mit den Ergebnissen
der Optimierung des VIS-Systems ersetzt. Die Zielwerte werden autoamtisch angepasst auf Grundlage der Tabelle GENF_Values. Die Tabelle heißt so,
weil NIR Lichtverteilungen ursprünglich nur für Punktlichtquellen gedacht waren, die den GENF Operand benutzen. Die Berechnungsweise ist allerdings
gleich zu den CENY_Values. Entscheidend ist, dass GENF_Values mit der NIR-Lichtverteilung gefüttert ist.

Die MF muss bereits in dem Startsystem angepasst sein. Es werden keine physikalischen Operanden eingefügt oder angepasst. Nur die 
technischen Operanden (CENY) werden angepasst. 

Es dürfen neben den 4 Linsen keine zusätzlichen Oberflächen im System enthalten sein. Die Oberflächen werden anhand ihrer Position im LDE gelöscht/hinzugefügt.

Für die Berechnung des Wirkungsgrads und der Modellierung im NSC müssen die Strahlquellendaten, wie zB.: Power, Abstrahlwinkel, Quellengeometrie in Set_NIR_System_V5.m
eingegeben werden. Oben im Code ist ein Kommentar, der auf die richtigen Zeilen verweist. 

Ein eigenes Startsytem muss in Input_for_Start über die ordnerpfade mit dem Dateinamen eingefügt werden!
Das Programm entscheidend auf Grundlage der Variabel 'FileNIR' ob ein eigenes Startsystem verwendet wird oder nicht. Wenn 'FileNIR' mit einem Pfad gefült ist, wird 
das eigene System geladen und verwendet. Ansonsten wird vom Programm ein System mit einer Punktlichtquelle erstellt. 