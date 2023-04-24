n = 1;

for i=2:Zeilen
    AusgabeExcel{i, 3} = Spot1_results.SpotData.GetRMSSpotSizeFor(n,1);     % hier werden die RMS Spot Ergebnisse geladen und direkt in die AusgabeExcel Datei geschrieben
    n = n+1;                                                                % die Results werden nur temporär für die Ausgabe geladen!
end

