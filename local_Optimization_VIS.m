%                                                                              %
%                                                                              %
%       ! Zuerst OpticStudio öffnen und im Reiter Programming auf den          %
%                Button Interactive Extension klicken !                        %                     
%                   Danach dieses Skript starten.                              %
%                                                                              %
%                                                                              %


System_Load;        % nach Connection ausführen

Cores = 16;                               % Anzahl der CPU Kerne


MaxPixel = 30;      %Die Anzahl der maximal simulierten Pixel


NumberOfFields = TheSystem.SystemData.Fields.NumberOfFields;

%Die Erhöhung der Fields soll in 5er Schritten erfolgen und bei 10 oder 15 beginnen
if NumberOfFields<10
    NumberOfFields = 10;
elseif mod(NumberOfFields,15)~=0
    NumberOfFields = 10;
end

k=1;

while NumberOfFields <= MaxPixel


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
    
    %die neuen Fields werden gesetzt auf Grundlage der Größe der Variable
    %'NumberOfFields' , die variieren kann. Die Positionen der Fields werden
    %verteilt von Y=0 bis zur größten vorhandenen Field Position Y=max_field
    for i=1:NumberOfFields-1
        SystExplorer.Fields.AddField(0, (i)*(max_field/(NumberOfFields-1)), 1.0);
        if i == NumberOfFields-1
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
%       LocalOpt = TheSystem.Tools.OpenLocalOptimization();
%       LocalOpt.Algorithm = ZOSAPI.Tools.Optimization.OptimizationAlgorithm.DampedLeastSquares;
%       LocalOpt.Cycles = ZOSAPI.Tools.Optimization.OptimizationCycles.Automatic;
%       LocalOpt.NumberOfCores = 8;
%       LocalOpt.RunAndWaitForCompletion();
%       LocalOpt.Close();
    
    % Die Erhöhung der Fields efolgt in 5er Schritten 

    if NumberOfFields+5>MaxPixel
        NumberOfFields = MaxPixel;
    else
        NumberOfFields = NumberOfFields+5;
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
    
    %die neuen Fields werden gesetzt auf Grundlage der Größe der Variable
    %'NumberOfFields' , die variieren kann. Die Positionen der Fields werden
    %verteilt von Y=0 bis zur größten vorhandenen Field Position Y=max_field
    for i=1:NumberOfFields-1
        SystExplorer.Fields.AddField(0, (i)*(max_field/(NumberOfFields-1)), 1.0);
        if i == NumberOfFields-1
            SystExplorer.Fields.DeleteFieldAt(2);
        end
    end
    
    xls_to_CENY_V7;
    
    fprintf('Durchlauf Nr.: %d.\n',k)
    TheSystem.SystemData.Fields.NumberOfFields
    k=k+1;
end
