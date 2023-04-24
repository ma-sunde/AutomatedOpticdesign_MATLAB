%                                                                              %
%                                                                              %
%       ! Zuerst OpticStudio öffnen und im Reiter Programming auf den          %
%                Button Interactive Extension klicken !                        %                     
%                   Danach dieses Skript starten.                              %
%                                                                              %
%                                                                              %

System_Load;

% Neues System erstellen ohne zu speichern. Alle Linsendaten werden von
% Matlab gespeichert
% TheSystem.New(false);

% if isempty(FileVIS) == 0 && isempty(FolderVIS) == 0
%     DGfile = System.String.Concat(FolderVIS, FileVIS);
%     TheSystem.LoadFile(DGfile, false);
% end

TheMFE = TheSystem.MFE;
TheMFE.ShowEditor;
SystExplorer = TheSystem.SystemData;


    %----- 1 Field (Laserlichtquelle) ---------------------------------
    
    if TheSystem.SystemData.Fields.NumberOfFields == 1
        %Es wird hier unterschieden zwischen Matrix System und einzelner
        %(Laser-)Lichtquelle. Bei einem Field werden GENF + OPVA Operanden
        %gesetzt. Bei mehreren Fields (Matrix) werden CENY Operanden eingesetzt


        
            % Hier werden die GENF Operanden gelöscht, um sie an die 
            % neue Field Anzahl anzupassen
        [count, firstIndex, lastIndex] = Find_Operand('GENF'); 

        for i=firstIndex:lastIndex
            Operand = TheMFE.GetOperandAt(i);
            for n=1:count
                if strcmp(Operand.Type, 'GENF') == true
                    TheMFE.DeleteRowAt(i-1);
                end
            end
        end
        
            % Hier werden die OPVA Operanden gelöscht, um sie an die 
            % neue Field Anzahl anzupassen
        [count, firstIndex, lastIndex] = Find_Operand('OPVA'); 

        for i=firstIndex:lastIndex
            Operand = TheMFE.GetOperandAt(i);
            for n=1:count
                if strcmp(Operand.Type, 'OPVA') == true
                    TheMFE.DeleteRowAt(i-1);
                end
            end
        end

            n = 1;  %n wird immer dann erhöht wenn ein Operand (oder Operanden-Paar) hinzugefügt wurde, der einen Wert aus der xls_to_CENY Ausgabe bekommen hat
            %Die Variable zählt mit, bis die maximale Anzahl der Flächenschwerpunkte
            %der Diagonalen der Lichtverteilung erreicht ist
        
    for i=2:(size(CENY_Values,1))
        if CENY_Values{i,4} == 0
           %der Wert 0 wird ignoriert, da der ChiefRay sowieso nicht beeinflusst werden kann und an dieser Position ankommen wird. Kein Operand dafür notwendig  
        else
           %zunächst wird der GENF Operand eingefügt
           TheMFE.InsertNewOperandAt(n);
           Operand = TheMFE.GetOperandAt(n); 
           Operand.ChangeType(ZOSAPI.Editors.MFE.MeritOperandType.GENF);
           Operand.GetCellAt(3).IntegerValue = 1;      %Damit bezieht sich der Operand auf die Wellenlänge an Position 1 - hier 808nm
           Operand.GetCellAt(5).DoubleValue = CENY_Values{i,4}*10^3;
           Operand.Target = (i-1)*(1/(size(CENY_Values,1)-1)); %Target wird nicht gewichtet, nur für die Übersicht
           n = n+1;

           %danach wird der OPVA Operand eingefügt
           TheMFE.InsertNewOperandAt(n);
           Operand = TheMFE.GetOperandAt(n);
           Operand.ChangeType(ZOSAPI.Editors.MFE.MeritOperandType.OPVA);
           Operand.GetCellAt(2).IntegerValue = n-1;    %Damit bezieht sich der Operand OPVA auf den GENF Operand eine Zeile vorher
           Operand.Weight = 5;
           Operand.Target = (i-1)*(1/(size(CENY_Values,1)-1));
           n = n+1;
        end

    end

    %------ mehrere Fields (Matrix) ---------------------------------------
        
    elseif TheSystem.SystemData.Fields.NumberOfFields > 1
        %Sind mehrere Fields vorhanden (Matrix) werden CENY Operanden verwendet

        %------------- Field Editor -------------------
        % determine maximum field in Y only
        max_field = 0;
        for i=1:TheSystem.SystemData.Fields.NumberOfFields
            if TheSystem.SystemData.Fields.GetField(i).Y > max_field 
                max_field = TheSystem.SystemData.Fields.GetField(i).Y ;
            end
        end

        %alle Fields außer das an position Y=0 und Y=max_field werden gelöscht
        for i=2:TheSystem.SystemData.Fields.NumberOfFields-1
            SystExplorer.Fields.DeleteFieldAt(2);
        end

        %die neuen Fields werden gesetzt auf Grundlage der Größe der Tabelle
        %CENY_Values, die variieren kann. Die Positionen der Fields werden
        %verteilt von Y=0 bis zur größten vorhandenen Field Position Y=max_field
        for i=1:(size(CENY_Values,1)-2)
            SystExplorer.Fields.AddField(0, (i)*(max_field/(size(CENY_Values,1)-2)), 1.0);
            if i == (size(CENY_Values,1)-2)
                SystExplorer.Fields.DeleteFieldAt(2);
            end
        end

        %-------------- Merit-Function Editor ----------------

        [count, firstIndex, lastIndex] = Find_Operand('CENY'); 

            % Hier werden die CENY Operanden gelöscht, um sie an die neue Field
            % Anzahl anzupassen
        for i=firstIndex:lastIndex
            Operand = TheMFE.GetOperandAt(i);
            for n=1:count
                if strcmp(Operand.Type, 'CENY') == true
                    TheMFE.DeleteRowAt(i-1);
                end
            end
        end

        n=1;    %Startpunkt zum Einsetzen der Operanden

        for i=2:(size(CENY_Values,1))

            TheMFE.InsertNewOperandAt(n);
            Operand = TheMFE.GetOperandAt(n); 
            Operand.ChangeType(ZOSAPI.Editors.MFE.MeritOperandType.CENY);
            Operand.GetCellAt(2).IntegerValue = TheLDE.NumberOfRows-1;      %Hier wird davon ausgegangen, dass die letzte Surface die Image Surface ist
            Operand.GetCellAt(4).IntegerValue = CENY_Values{i,1};
            Operand.Target = CENY_Values{i,4};
            Operand.Weight = 5;
            n = n+1;

        end

    
    end

