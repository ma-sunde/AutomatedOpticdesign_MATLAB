%                                                                              %
%                                                                              %
%       ! Zuerst OpticStudio öffnen und im Reiter Programming auf den          %
%                Button Interactive Extension klicken !                        %                     
%                   Danach dieses Skript starten.                              %
%                                                                              %
%                                                                              %

System_Load;        % nach Connection ausführen
SystExplorer = TheSystem.SystemData;

TheSystem.SaveAs(System.String.Concat(FolderOutput, 'VIS_vor_lokaler_Optimierung.zmx'));

Cores = 64;                               % Anzahl der CPU Kerne


MaxPixel = 20;      %Die Anzahl der maximal simulierten Pixel


%Die Erhöhung der Fields soll in 5er Schritten erfolgen und bei 10 beginnen
NumberOfFields_SET = 10;


k=1;    %Index für Anzahl der Durchläufe

while NumberOfFields_SET < MaxPixel


%------------------ lokale Optimierung ------------------------------------


    %------------- Field Editor -------------------

    if NumberOfFields_SET ~= TheSystem.SystemData.Fields.NumberOfFields
        % determine maximum field in Y only
        max_field = TheSystem.SystemData.Fields.GetField(TheSystem.SystemData.Fields.NumberOfFields).Y;

        %alle Fields außer das an position Y=0 und Y=max_field werden gelöscht
        for i=2:TheSystem.SystemData.Fields.NumberOfFields-1
            SystExplorer.Fields.DeleteFieldAt(2);
        end
        
        %die neuen Fields werden gesetzt auf Grundlage der Größe der Variable
        %'NumberOfFields' , die variieren kann. Die Positionen der Fields werden
        %verteilt von Y=0 bis zur größten vorhandenen Field Position Y=max_field
        for i=1:NumberOfFields_SET-1
            SystExplorer.Fields.AddField(0, (i)*(max_field/(NumberOfFields_SET-1)), 1.0);
            if i == NumberOfFields_SET-1
                SystExplorer.Fields.DeleteFieldAt(2);
            end
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
    
    %Wizward Einträge löschen, um neuen Wizard einzutragen
    [count, firstIndex, lastIndex] = Find_Operand('DMFS');
        for n=firstIndex:TheMFE.NumberOfRows
            TheMFE.DeleteRowAt(firstIndex-1);
        end
    

    %Optimize for smallest RMS Spot, which is "Data" = 1
    OptWizard = TheMFE.SEQOptimizationWizard;
    OptWizard.StartAt = firstIndex;
    OptWizard.Data = 1;
    OptWizard.OverallWeight = 10000000;              %5000000
    OptWizard.IsAssumeAxialSymmetryUsed = 0;        %nur für Matrix-Systeme!
        
    %And click OK!
    OptWizard.Apply();


    fprintf('Durchlauf Nr.: %d.\n',k)
    TheSystem.SystemData.Fields.NumberOfFields
    k=k+1;

    % Local optimisation till completion
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
     
   
    %------------- Field Editor -------------------

    if NumberOfFields_SET ~= TheSystem.SystemData.Fields.NumberOfFields
        % determine maximum field in Y only
        max_field = TheSystem.SystemData.Fields.GetField(TheSystem.SystemData.Fields.NumberOfFields).Y;
        
        %alle Fields außer das an position Y=0 und Y=max_field werden gelöscht
        for i=2:TheSystem.SystemData.Fields.NumberOfFields-1
            SystExplorer.Fields.DeleteFieldAt(2);
        end
        
        %die neuen Fields werden gesetzt auf Grundlage der Größe der Variable
        %'NumberOfFields' , die variieren kann. Die Positionen der Fields werden
        %verteilt von Y=0 bis zur größten vorhandenen Field Position Y=max_field
        for i=1:NumberOfFields_SET-1
            SystExplorer.Fields.AddField(0, (i)*(max_field/(NumberOfFields_SET-1)), 1.0);
            if i == NumberOfFields_SET-1
                SystExplorer.Fields.DeleteFieldAt(2);
            end
        end
    end


    xls_to_CENY_V8;
    
end

%-------------------- letzter Durchlauf -----------------------------------

%------------------ lokale Optimierung ------------------------------------


    %------------- Field Editor -------------------

    if NumberOfFields_SET ~= TheSystem.SystemData.Fields.NumberOfFields
        % determine maximum field in Y only
        max_field = TheSystem.SystemData.Fields.GetField(TheSystem.SystemData.Fields.NumberOfFields).Y;
        
        %alle Fields außer das an position Y=0 und Y=max_field werden gelöscht
        for i=2:TheSystem.SystemData.Fields.NumberOfFields-1
            SystExplorer.Fields.DeleteFieldAt(2);
        end
        
        %die neuen Fields werden gesetzt auf Grundlage der Größe der Variable
        %'NumberOfFields' , die variieren kann. Die Positionen der Fields werden
        %verteilt von Y=0 bis zur größten vorhandenen Field Position Y=max_field
        for i=1:NumberOfFields_SET-1
            SystExplorer.Fields.AddField(0, (i)*(max_field/(NumberOfFields_SET-1)), 1.0);
            if i == NumberOfFields_SET-1
                SystExplorer.Fields.DeleteFieldAt(2);
            end
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
    
    %Wizward Einträge löschen, um neuen Wizard einzutragen
    [count, firstIndex, lastIndex] = Find_Operand('DMFS');
        for n=firstIndex:TheMFE.NumberOfRows
            TheMFE.DeleteRowAt(firstIndex-1);
        end
    

    %Optimize for smallest RMS Spot, which is "Data" = 1
    OptWizard = TheMFE.SEQOptimizationWizard;
    OptWizard.StartAt = firstIndex;
    OptWizard.Data = 1;
    OptWizard.OverallWeight = 10000000;              %5000000
    OptWizard.IsAssumeAxialSymmetryUsed = 0;        %nur für Matrix-Systeme!
        
    %And click OK!
    OptWizard.Apply();


    fprintf('Durchlauf Nr.: %d.\n',k)
    TheSystem.SystemData.Fields.NumberOfFields
    k=k+1;

    % Local optimisation till completion
       LocalOpt = TheSystem.Tools.OpenLocalOptimization();
       LocalOpt.Algorithm = ZOSAPI.Tools.Optimization.OptimizationAlgorithm.DampedLeastSquares;
       LocalOpt.Cycles = ZOSAPI.Tools.Optimization.OptimizationCycles.Automatic;
       LocalOpt.NumberOfCores = Cores;
       LocalOpt.RunAndWaitForCompletion();
       LocalOpt.Close();

 TheSystem.SaveAs(System.String.Concat(FolderOutput, 'VIS_nach_lokaler_Optimierung.zmx'));


 [count, firstIndex, lastIndex] = Find_Operand('CENY');     % findet den gesuchten Operanden und gibt Anzahl(count), erste Position in MF(firstIndex) und letzte Position (lastIndex) wider.

Zeilen = count +1;                      % Die Anzahl der Zeilen ist Die Anzahl der Operanden + Zeile für Beschriftung (deswegen +1)
Analyse = cell(Zeilen, 9);              % Die Ausgabetabelle bekommt hier ihr Format: Zeilen(varriert je nach Fields) & 9 Spalten
Analyse{1, 1} = 'Fieldnummer'; 
Analyse{1, 2} = 'Target CENY';
Analyse{1, 3} = ''; 
Analyse{1, 4} = 'neue Werte CENY'; 
Analyse{1, 5} = ''; 
Analyse{1, 7} = ''; 
Analyse{1, 8} = 'neue Werte RMS';
Analyse{1, 9} = ''; 

%------------------CENY----------------------------------------------------------------------------------------------------------

% Hier werden die CENY "Targets" und "Values" abgespeichert
n = 2;  %Eintragen der Werte ab Zeile 2, wegen Überschriften

TheMFE.CalculateMeritFunction; %aktualisieren der MF-Values

for i=firstIndex:lastIndex
    Operand = TheMFE.GetOperandAt(i);
    if strcmp(Operand.Type, 'CENY') == true
        FieldCENY = Operand.GetCellAt(4);
        Analyse{n, 1} = FieldCENY.IntegerValue;       
        Analyse{n, 2} = Operand.Target;               
        Analyse{n, 4} = Operand.Value;               
        n = n+1;
    end
end

%----------------RMS--------------------------------------------------------------------------------------------------------------

    % Hier werden die alten RMS abgespeichert
    Spot1 = TheSystem.Analyses.New_Analysis(ZOSAPI.Analysis.AnalysisIDM.StandardSpot);
    Spot1.ApplyAndWaitForCompletion();
    Spot1_results = Spot1.GetResults();
    
    % für die korrekte Zuordnung werden die Fields aus den CENY operanden ausgelesen, da die Tabelle ihren Ursprung in den CENY Werten hat
    % GetCellAt(4) ist die 4. Spalte des CENY Operanden, in der die Fields stehen
    % hier werden die RMS Spot Ergebnisse geladen und direkt in die AusgabeExcel Datei geschrieben
    % die Results werden nur temporär für die Ausgabe geladen!
    
    n = 2;  %Eintragen der Werte ab Zeile 2, wegen Überschriften
    for i=firstIndex:lastIndex
        Operand = TheMFE.GetOperandAt(i);
        if strcmp(Operand.Type, 'CENY') == true
            FieldCENY = Operand.GetCellAt(4);
            Analyse{n, 8} = Spot1_results.SpotData.GetRMSSpotSizeFor(FieldCENY.IntegerValue,1);               
            n = n+1;
        end
    end   
    Spot1.Close;

file_path = fullfile(FolderVIS, 'lokale_OptimierungVIS.xls');
xlswrite(file_path,Analyse);
winopen(file_path);
