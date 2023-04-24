%                                                                              %
%                                                                              %
%       ! Zuerst OpticStudio öffnen und im Reiter Programming auf den          %
%                Button Interactive Extension klicken !                        %                     
%                   Danach dieses Skript starten.                              %
%                                                                              %
%                                                                              %


System_Load;        % nach Connection ausführen
SystExplorer = TheSystem.SystemData;

TheSystem.SaveAs(System.String.Concat(FolderOutput, 'NIR_vor_lokaler_Optimierung.zmx'));

Cores = 64;                               % Anzahl der CPU Kerne


MaxPixel = 20;      %Die Anzahl der maximal simulierten Pixel


k=1;    %Index für Anzahl der Durchläufe


if TheSystem.SystemData.Fields.NumberOfFields == 1

        [count, firstIndex, lastIndex] = Find_Operand('GENF');
        NumberOfFields_SET = count;     

    %Die Erhöhung der Fields soll in 5er Schritten erfolgen und bei 10 oder 15 beginnen
    if NumberOfFields_SET<10
        NumberOfFields_SET = 10;
    elseif mod(NumberOfFields_SET,15)~=0
        NumberOfFields_SET = 10;
    end    

    k=1;
    
    while NumberOfFields_SET < MaxPixel   

    %-------------- Merit-Function Editor ----------------

        [count, firstIndex, lastIndex] = Find_Operand('GENF'); 
            % Hier werden die GENF Operanden gelöscht, um sie an die neue Pixel
            % Anzahl anzupassen
        for i=firstIndex:lastIndex
            Operand = TheMFE.GetOperandAt(i);
            for n=1:count
                if strcmp(Operand.Type, 'GENF') == true
                    TheMFE.DeleteRowAt(i-1);
                end
            end
        end
        
        [count, firstIndex, lastIndex] = Find_Operand('OPVA'); 
            % Hier werden die OPVA Operanden gelöscht, um sie an die neue Pixel
            % Anzahl anzupassen
        for i=firstIndex:lastIndex
            Operand = TheMFE.GetOperandAt(i);
            for n=1:count
                if strcmp(Operand.Type, 'OPVA') == true
                    TheMFE.DeleteRowAt(i-1);
                end
            end
        end


        TheMFE = TheSystem.MFE;
        TheMFE.ShowEditor;
        n = 1;  %n wird immer dann erhöht wenn ein Operand (oder Operanden-Paar) hinzugefügt wurde, der einen Wert aus der xls_to_GENF Ausgabe bekommen hat
                %Die Variable zählt mit, bis die maximale Anzahl der Flächenschwerpunkte
                %der Diagonalen der Lichtverteilung erreicht ist

        for i=2:(size(GENF_Values,1))
            if GENF_Values{i,4} == 0
               %der Wert 0 wird ignoriert, da der ChiefRay sowieso nicht beeinflusst werden kann und an dieser Position ankommen wird. Kein Operand dafür notwendig  
            else
               %zunächst wird der GENF Operand eingefügt
               TheMFE.InsertNewOperandAt(n);
               Operand = TheMFE.GetOperandAt(n); 
               Operand.ChangeType(ZOSAPI.Editors.MFE.MeritOperandType.GENF);
               Operand.GetCellAt(3).IntegerValue = 1;      %Damit bezieht sich der Operand auf die Wellenlänge an Position 1 - hier 808nm
               Operand.GetCellAt(5).DoubleValue = GENF_Values{i,4}*10^3;
               Operand.Target = (i-1)*(1/(size(GENF_Values,1)-1)); %Target wird nicht gewichtet, nur für die Übersicht
               n = n+1;

               %danach wird der OPVA Operand eingefügt
               TheMFE.InsertNewOperandAt(n);
               Operand = TheMFE.GetOperandAt(n);
               Operand.ChangeType(ZOSAPI.Editors.MFE.MeritOperandType.OPVA);
               Operand.GetCellAt(2).IntegerValue = n-1;    %Damit bezieht sich der Operand OPVA auf den GENF Operand eine Zeile vorher
               Operand.Weight = 5;
               Operand.Target = (i-1)*(1/(size(GENF_Values,1)-1));
               n = n+1;
            end

        end


       %--- Local optimisation till completion
        LocalOpt = TheSystem.Tools.OpenLocalOptimization();
        LocalOpt.Algorithm = ZOSAPI.Tools.Optimization.OptimizationAlgorithm.DampedLeastSquares;
        LocalOpt.Cycles = ZOSAPI.Tools.Optimization.OptimizationCycles.Automatic;
        LocalOpt.NumberOfCores = Cores;
        LocalOpt.RunAndWaitForCompletion();
        LocalOpt.Close();

        % Die Erhöhung der Fields efolgt in 5er Schritten 

        if NumberOfFields_SET+5>MaxPixel
            NumberOfFields_SET = MaxPixel;
        else
            NumberOfFields_SET = NumberOfFields_SET+5;
        end
        
        [count, firstIndex, lastIndex] = Find_Operand('GENF'); 
            % Hier werden die GENF Operanden gelöscht, um sie an die neue Pixel
            % Anzahl anzupassen
        for i=firstIndex:lastIndex
            Operand = TheMFE.GetOperandAt(i);
            for n=1:count
                if strcmp(Operand.Type, 'GENF') == true
                    TheMFE.DeleteRowAt(i-1);
                end
            end
        end
        
        [count, firstIndex, lastIndex] = Find_Operand('OPVA'); 
            % Hier werden die OPVA Operanden gelöscht, um sie an die neue Pixel
            % Anzahl anzupassen
        for i=firstIndex:lastIndex
            Operand = TheMFE.GetOperandAt(i);
            for n=1:count
                if strcmp(Operand.Type, 'OPVA') == true
                    TheMFE.DeleteRowAt(i-1);
                end
            end
        end
        
        %Einfügen der Operanden als Platzhalter. Nur so erkennt xls_to_GENF
        %die aktuelle Anzahl an Operanden, die für den Durchlauf vorgesehen
        %sind
        n=1;
        
        for i=1:NumberOfFields_SET

               %zunächst wird der GENF Operand eingefügt
               TheMFE.InsertNewOperandAt(n);
               Operand = TheMFE.GetOperandAt(n); 
               Operand.ChangeType(ZOSAPI.Editors.MFE.MeritOperandType.GENF);
               n = n+1;

               %danach wird der OPVA Operand eingefügt
               TheMFE.InsertNewOperandAt(n);
               Operand = TheMFE.GetOperandAt(n);
               Operand.ChangeType(ZOSAPI.Editors.MFE.MeritOperandType.OPVA);
               n = n+1;
        end
        
        %Targets für die Operanden anhand der aktuellen Anzahl an Pixel
        %analysieren
            xls_to_GENF_V2;
        
        %---Platzhalter wieder löschen
        [count, firstIndex, lastIndex] = Find_Operand('GENF'); 
            % Hier werden die GENF Operanden gelöscht, um sie an die neue Pixel
            % Anzahl anzupassen
        for i=firstIndex:lastIndex
            Operand = TheMFE.GetOperandAt(i);
            for n=1:count
                if strcmp(Operand.Type, 'GENF') == true
                    TheMFE.DeleteRowAt(i-1);
                end
            end
        end
        
        [count, firstIndex, lastIndex] = Find_Operand('OPVA'); 
            % Hier werden die OPVA Operanden gelöscht, um sie an die neue Pixel
            % Anzahl anzupassen
        for i=firstIndex:lastIndex
            Operand = TheMFE.GetOperandAt(i);
            for n=1:count
                if strcmp(Operand.Type, 'OPVA') == true
                    TheMFE.DeleteRowAt(i-1);
                end
            end
        end


        TheMFE = TheSystem.MFE;
        TheMFE.ShowEditor;
        n = 1;  %n wird immer dann erhöht wenn ein Operand (oder Operanden-Paar) hinzugefügt wurde, der einen Wert aus der xls_to_GENF Ausgabe bekommen hat
                %Die Variable zählt mit, bis die maximale Anzahl der Flächenschwerpunkte
                %der Diagonalen der Lichtverteilung erreicht ist

        for i=2:(size(GENF_Values,1))
            if GENF_Values{i,4} == 0
               %der Wert 0 wird ignoriert, da der ChiefRay sowieso nicht beeinflusst werden kann und an dieser Position ankommen wird. Kein Operand dafür notwendig  
            else
               %zunächst wird der GENF Operand eingefügt
               TheMFE.InsertNewOperandAt(n);
               Operand = TheMFE.GetOperandAt(n); 
               Operand.ChangeType(ZOSAPI.Editors.MFE.MeritOperandType.GENF);
               Operand.GetCellAt(3).IntegerValue = 1;      %Damit bezieht sich der Operand auf die Wellenlänge an Position 1 - hier 808nm
               Operand.GetCellAt(5).DoubleValue = GENF_Values{i,4}*10^3;
               Operand.Target = (i-1)*(1/(size(GENF_Values,1)-1)); %Target wird nicht gewichtet, nur für die Übersicht
               n = n+1;

               %danach wird der OPVA Operand eingefügt
               TheMFE.InsertNewOperandAt(n);
               Operand = TheMFE.GetOperandAt(n);
               Operand.ChangeType(ZOSAPI.Editors.MFE.MeritOperandType.OPVA);
               Operand.GetCellAt(2).IntegerValue = n-1;    %Damit bezieht sich der Operand OPVA auf den GENF Operand eine Zeile vorher
               Operand.Weight = 5;
               Operand.Target = (i-1)*(1/(size(GENF_Values,1)-1));
               n = n+1;
            end

        end
    
    fprintf('Durchlauf Nr.: %d.\n',k)
    NumberOfFields_SET
    k=k+1;

    end
    
%------------------GENF---------------------

    [count, firstIndex, lastIndex] = Find_Operand('GENF');     % findet den gesuchten Operanden und gibt Anzahl(count), erste Position in MF(firstIndex) und letzte Position (lastIndex) wider.

Zeilen = count +1;                 % Die Anzahl der Zeilen ist Die Anzahl der Operanden + Zeile für Beschriftung (deswegen +1)
Analyse = cell(Zeilen, 5);              % Die Ausgabetabelle bekommt hier ihr Format: Zeilen(varriert je nach Fields) & 9 Spalten
Analyse{1, 1} = 'Operanden Nr.'; 
Analyse{1, 2} = 'Target Dist GENF';
Analyse{1, 3} = 'Target GENF';
Analyse{1, 4} = ''; 
Analyse{1, 5} = 'neue Werte GENF'; 
Analyse{1, 6} = ''; 
Analyse{1, 7} = ''; 
Analyse{1, 8} = '';
Analyse{1, 9} = ''; 

%------------------GENF----------------------------------------------------------------------------------------------------------

% Hier werden die "GENF Distance Targets" die "Targets" und "Values" abgespeichert
n = 2;  %Eintragen der Werte ab Zeile 2, wegen Überschriften
for i=firstIndex:lastIndex
    Operand = TheMFE.GetOperandAt(i);
    if strcmp(Operand.Type, 'GENF') == true
        FieldGENF = Operand.GetCellAt(5);
        Analyse{n, 1} = n-1;        %Nummerierung der Werte; Beachte n beginnt bei 2, wegen Üebrschriften der Tabelle
        Analyse{n, 2} = FieldGENF.DoubleValue;        % Distance Targets werden in Spalte 2 gespeichert
        Analyse{n, 3} = Operand.Target;               % Target der Operanden werden in Spalte 3 gespeichert
        Analyse{n, 5} = Operand.Value;                % Ausgangswert des GENF Operanden wird in Spalte 4 gespeichert
        n = n+1;
    end
end

% Hier werden die neuen Werte in die Tabelle gespeichert

TheMFE.CalculateMeritFunction; %aktualisieren der MF-Values

n = 2;
for i=firstIndex:lastIndex
    Operand = TheMFE.GetOperandAt(i);
    if strcmp(Operand.Type, 'GENF') == true
        Analyse{n, 5} = Operand.Value;                % Ausgangswert des GENF Operanden wird in Spalte 4 gespeichert
        n = n+1;
    end
end


% Hier wird die Differenz berechnet und in der Tabelle abgespeichert
n = 2;  
for i=2:size(Analyse,1) 
    Difference = abs((100/Analyse{n,3})*Analyse{n,5}-100);        % Differenz wird ohne Vorzeichen berechnet --> abs = Betrag
    Analyse{n, 6} = Difference;                                   % Die Differenz wird in Spalte 5 gespeichert                               
    SpalteDifferenz(n,1) = Difference;
    n = n+1;
end

MittelwertGENF = mean(SpalteDifferenz);

file_path = fullfile(FolderVIS, 'lokale_OptimierungNIR.xls');
xlswrite(file_path,Analyse);
winopen(file_path);

end


if TheSystem.SystemData.Fields.NumberOfFields ~= 1

    local_Optimization_VIS_V2;

end

TheSystem.SaveAs(System.String.Concat(FolderOutput, 'NIR_nach_lokaler_Optimierung.zmx'));


