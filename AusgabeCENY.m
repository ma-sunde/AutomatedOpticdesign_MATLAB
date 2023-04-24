n = 1;

for i=2:Zeilen
    AusgabeExcel{i, 1} = n;
    Operand = TheMFE.GetOperandAt(CENYfirst);       % hier werden die Operanden geladen und direkt in die AusgabeExcel Datei geschrieben
    AusgabeExcel{i, 2} = Operand.Value;     % die Operanden werden nur temporär für die Ausgabe geladen!
    n = n+1;
    CENYfirst = CENYfirst +1;
end
