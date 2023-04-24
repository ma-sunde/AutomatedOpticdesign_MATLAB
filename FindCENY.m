n = TheMFE.NumberOfRows;
s2 = 'CENY';


% Diese Schleife findet den 1. CENY Operanden und speichert Position in i
for i=1:n
    Operand = TheMFE.GetOperandAt(i);
    if strcmp(Operand.Type,s2) == true
        break
    end
end

CENYfirst = i;

% Diese Schleife findet den letzten CENY Operanden und speichert Position in k

for k=i:n
    Operand = TheMFE.GetOperandAt(k);
    if strcmp(Operand.Type,s2) == false
        break
    end
end

CENYlast = k-1;
AnzahlCENY = k-i;