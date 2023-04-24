n = TheNCE.NumberOfRows;
s2 = 'NullObject';              % Vergleichsvariable für stringcompare

% Diese Schleife findet das Null Object und speichert Position in i

for p=1:n
    Object = TheNCE.GetObjectAt(p);
    if strcmp(Object.Type,s2) == true
        break
    end
end

PosNullObject = p;