%                                                                              %
%                                                                              %
%       ! Zuerst OpticStudio öffnen und im Reiter Programming auf den          %
%                Button Interactive Extension klicken !                        %                     
%                   Danach dieses Skript starten.                              %
%                                                                              %
%                                                                              %


ystem_Load;        % nach Connection ausführen
SystExplorer = TheSystem.SystemData;


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


end


if TheSystem.SystemData.Fields.NumberOfFields ~= 1

    %Die Erhöhung der Fields soll in 5er Schritten erfolgen und bei 10 oder 15 beginnen
    if NumberOfFields_SET<10
        NumberOfFields_SET = 10;
    elseif mod(NumberOfFields_SET,15)~=0
        NumberOfFields_SET = 10;
    end


    k=1;

    while NumberOfFields_SET <= MaxPixel


    %------------------ lokale Optimierung ------------------------------------


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
        for i=1:(size(CENY_Values,1)-1)
            SystExplorer.Fields.AddField(0, (i)*(max_field/(size(CENY_Values,1)-1)), 1.0);
            if i == (size(CENY_Values,1)-1)
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
            Operand.GetCellAt(5).IntegerValue = CENY_Values{i,1};
            Operand.Target = CENY_Values{i,4};
            Operand.Weight = 5;
            n = n+1;

        end


        % Local optimisation till completion
          LocalOpt = TheSystem.Tools.OpenLocalOptimization();
          LocalOpt.Algorithm = ZOSAPI.Tools.Optimization.OptimizationAlgorithm.DampedLeastSquares;
          LocalOpt.Cycles = ZOSAPI.Tools.Optimization.OptimizationCycles.Automatic;
          LocalOpt.NumberOfCores = 8;
          LocalOpt.RunAndWaitForCompletion();
          LocalOpt.Close();

        % Die Erhöhung der Fields efolgt in 5er Schritten 

        if NumberOfFields_SET+5>MaxPixel
            NumberOfFields_SET = MaxPixel;
        else
            NumberOfFields_SET = NumberOfFields_SET+5;
        end


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
        for i=1:NumberOfFields_SET-1
            SystExplorer.Fields.AddField(0, (i)*(max_field/(NumberOfFields_SET-1)), 1.0);
            if i == NumberOfFields_SET-1
                SystExplorer.Fields.DeleteFieldAt(2);
            end
        end

        xls_to_CENY_V7;

        fprintf('lokale Optimierung Durchlauf Nr.: %d.\n',k)
        NumberOfFields_SET
        k=k+1;
    end

end


