n = TheNCE.NumberOfRows;
s2 = 'OddAsphereLens';              % Vergleichsvariable für stringcompare
s3 = 'StandardLens';

% Diese Schleife findet die erste Surface und speichert Position in i,
% damit in Anschluss die Lichtquelle davor platziert werden kann
for i=1:n
    Object = TheNCE.GetObjectAt(i);
    if strcmp(Object.Type,s2) == true | strcmp(Object.Type,s3) == true
        break
    end
end

SurfacePosfirst = i;

% Diese Schleife findet alle Surfaces und speichert die letzte Position in
% k, damit im Anschluss die Detector Ebene platziert werden kann

for k=i:n
    Object = TheNCE.GetObjectAt(k);
    if strcmp(Object.Type,s2) == false & strcmp(Object.Type,s3) == false
        break
    end
end

SurfacePoslast = k-1;
