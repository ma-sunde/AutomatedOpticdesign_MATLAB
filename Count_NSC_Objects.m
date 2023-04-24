Anzahl_Surfaces_NSC = 0;
% Diese Schleife findet Surfaces und speichert die Positionen,
% damit in Anschluss die Lichtquelle davor und die Detektor Ebende dahinter platziert werden kann
for i=1:TheNCE.NumberOfRows
    Object = TheNCE.GetObjectAt(i);
    if startsWith(char(Object.Comment), 'surfaces') == 1
        Anzahl_Surfaces_NSC = Anzahl_Surfaces_NSC +1;
    end
end

for i=1:TheNCE.NumberOfRows
    Object = TheNCE.GetObjectAt(i);
    if startsWith(char(Object.Comment), 'surfaces') == 1
        break
    end
end

SurfacePosfirst = i; 

% Das letzte Objekt wird gespeichert um die Position vom Ende des Linsensystems zu
% finden

for i=SurfacePosfirst:TheNCE.NumberOfRows
    Object = TheNCE.GetObjectAt(i);
    if startsWith(char(Object.Comment), 'surfaces') ~= 1
        break
    end
end

SurfacePoslast = i;
