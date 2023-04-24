RUN.m - 		Hauptprogramm
			Dort sind alle Unterprogramme in der korrekten Reihenfolge einsortiert.

Input_for_Start.m - 	Zu Beginn zu überprüfen. Dort müssen Ordnerpfade festgelegt werden, aus denen Systeme geladen und Ergebnisse gespeichert werden.

System_Load.m. - 	Stellt Verbindung zu Zemax OpticStduio her, wenn zuvor die Interactive Extension bei ZOS aktiviert wurde.
			Alle Unterprogramme starten mit System_Load.m und können so schnell getestet werden. 
			Eine erneute Aufrufung des Skriptes während eine Verbindung besteht, hat keinen negativen Einfluss, deswegen ist es überall eingebettet.

xls_to_CENY_V8.m -	'Auslesen der Lichtverteilung'. Wandelt eine Lichtverteilung in CENY Zielwerte um.

Set_VIS_System.m - 	Trägt die Zielwerte in die MF des VIS-Systems ein.

Optimization_VIS_V4.m - Globale Optimierung VIS-System mit mehreren Fields.

local_Optimization_VIS_V2.m - 		lokale Optimierung VIS-System mit mehreren Fields.

Convert_VIS_to_NSC_and_Efficiency.m - 	Konvertiert das System in den nicht-sequentiellen Modus und berechnet den Wirkungsgrad.

Get_All_Parameters_from_NSC_V2.m - 	Speichert alle notwendigen Parameter für die Rückkonvertierung in den sequentiellen Modus im Matlab Workspace ab.

Set_Full_Optimized_VIS_System.m -	Baut das System im sequentiellen Modus auf.

Differenz_Lichtverteilung.m - 		Berechnet die Differenz Lichtverteilung.

xls_to_GENF_V2.m - 			'Auslesen der Lichtverteilung'. Wandelt eine Lichtverteilung in GENF Zielwerte um.

Get_Lenses_from_LDE_for_NIR_System.m - 	Speichert die hinteren beiden Linsen des VIS-System im Matlab Workspace ab.

Set_NIR_System_V4.m -			Baut das System im sequentiellen Modus auf.

... Die restlichen Programme gleichen sich vom prinzip her mit denen des VIS-Systems.


Made by: Max Caspar Sundermeier & Lukas Hanisch
ZOS-Version: 21.1.2
Matlab-Version: R2020b

