function [count, firstIndex, lastIndex] = Find_Operand(name)
% Die Funktion durchsucht eine Liste nach einem bestimmten Namen
% und gibt die Anzahl der gefundenen Namen, die erste Position und die
% letzte Position zurück.

% In Funktionen müssen die ZOS-API Komponenten neu geladen werden, sonst
% kann auf die Befehle nicht zugegriffenw erden
TheApplication = MATLABZOSConnection;
TheSystem = TheApplication.PrimarySystem;
TheMFE = TheSystem.MFE;

% Initialisierung der Variablen
count = 0;
firstIndex = -1;
lastIndex = -1;

% Durchsuchen der Liste
for i = 1:TheMFE.NumberOfRows
    Operand = TheMFE.GetOperandAt(i);
    if strcmp(Operand.Type, name) == true
        count = count + 1;
        if firstIndex == -1
            firstIndex = i;
        end
        lastIndex = i;
    end
end

% Ausgabe der Ergebnisse
% if count == 0
%     fprintf('Der Name "%s" wurde nicht gefunden.\n', name);
% else
%     fprintf('Es wurden %d Vorkommen des Namens "%s" gefunden.\n', count, name);
%     fprintf('Erste Position: %d\n', firstIndex);
%     fprintf('Letzte Position: %d\n', lastIndex);
% end
