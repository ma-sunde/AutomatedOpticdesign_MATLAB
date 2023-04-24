%                                                                              %
%                                                                              %
%       ! Zuerst OpticStudio öffnen und im Reiter Programming auf den          %
%                Button Interactive Extension klicken !                        %                     
%                   Danach dieses Skript starten.                              %
%                                                                              %
%                                                                              %

System_Load;        % nach Connection ausführen

TheSystem.SaveAs(System.String.Concat(FolderOutput, 'NIR_vor_globaler_Optimierung.zmx'));

MaxOptTimeInSeconds = 1200;                 % Maximale Zeit eines Durchlaufes
Faktor = 1;                               % Der Faktor dient dazu diese Zeit zu verlängern und wird innerhalb der Schleife verändert. Nur für den Start wird er auf 1 gesetzt
StartTime = 300;                           % Dauer des ersten Durchlaufes
OptTimeInSeconds = StartTime*Faktor;      % Startzeit
Cores = 64;                              % Anzahl der CPU Kerne


%für Punktlichtquellensysteme:
if TheSystem.SystemData.Fields.NumberOfFields == 1 

[count, firstIndex, lastIndex] = Find_Operand('GENF');     % findet den gesuchten Operanden und gibt Anzahl(count), erste Position in MF(firstIndex) und letzte Position (lastIndex) wider.

Zeilen = count +1;                 % Die Anzahl der Zeilen ist Die Anzahl der Operanden + Zeile für Beschriftung (deswegen +1)
Analyse = cell(Zeilen, 5);              % Die Ausgabetabelle bekommt hier ihr Format: Zeilen(varriert je nach Fields) & 9 Spalten
Analyse{1, 1} = 'Operanden Nr.'; 
Analyse{1, 2} = 'Target Dist GENF';
Analyse{1, 3} = 'Target GENF';
Analyse{1, 4} = 'alte Werte GENF'; 
Analyse{1, 5} = 'neue Werte GENF'; 
Analyse{1, 6} = 'Differenz in % GENF'; 
%Analyse{1, 7} = 'alte Werte RMS'; 
%Analyse{1, 8} = 'neue Werte RMS';
%Analyse{1, 9} = 'Differenz in % RMS'; 

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
        Analyse{n, 4} = Operand.Value;                % Ausgangswert des GENF Operanden wird in Spalte 4 gespeichert
        n = n+1;
    end
end




%----------------RMS--------------------------------------------------------------------------------------------------------------

% Hier werden die alten RMS abgespeichert
% Spot1 = TheSystem.Analyses.New_Analysis(ZOSAPI.Analysis.AnalysisIDM.StandardSpot);
% Spot1.ApplyAndWaitForCompletion();
% Spot1_results = Spot1.GetResults();
% 
% n = CENYfirst;
% for i=2:Zeilen
%     Operand = TheMFE.GetOperandAt(n);
%     FieldCENY = Operand.GetCellAt(4);                                                           % für die korrekte Zuordnung werden die Fields aus den CENY operanden ausgelesen, da die Tabelle ihren Ursprung in den CENY Werten hat
%                                                                                                 % GetCellAt(4) ist die 4. Spalte des CENY Operanden, in der die Fields stehen
%     Analyse{i, 7} = Spot1_results.SpotData.GetRMSSpotSizeFor(FieldCENY.IntegerValue,1);         % hier werden die RMS Spot Ergebnisse geladen und direkt in die AusgabeExcel Datei geschrieben
%     n = n +1;                                                                                   % die Results werden nur temporär für die Ausgabe geladen!
% end
% 
% Spot1.Close;

%----------------MF---------------------------------------------------------------------------------------------------------------

% Hier wird der MF-Value abgespeichert
MF = cell(4,3);
MF{1,1} = 'Durchlauf';
MF{1,2} = 'Value MF';
MF{1,3} = 'Dauer';
NumberOfOpt = 1;                            % Variable, die zählt wie oft eine optimierung durchgeführt wurde
MF{NumberOfOpt+1,1} = NumberOfOpt;
MF{NumberOfOpt+1,2} = TheMFE.CalculateMeritFunction;
MF{NumberOfOpt+1,3} = OptTimeInSeconds;


%---------------1. Optimierung-------------------------------------------------------------------------------------------

% Hammer Optimierung
HammerOpt = TheSystem.Tools.OpenHammerOptimization();
HammerOpt.Algorithm = ZOSAPI.Tools.Optimization.OptimizationAlgorithm.DampedLeastSquares;
HammerOpt.NumberOfCores = Cores;        

    MF_Value = TheMFE.CalculateMeritFunction;
    fprintf('Hammer Optimization for %i seconds...\n', OptTimeInSeconds);
    fprintf('Initial Merit Function %i \n', MF_Value);

HammerOpt.RunAndWaitWithTimeout(OptTimeInSeconds);
HammerOpt.Cancel();
HammerOpt.WaitForCompletion();
HammerOpt.Close();

%---------------GENF--------------------------------------------------------------------------------------------------------------

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

%---------------RMS----------------------------------------------------------------------------------------------------------------------------

% % Hier werden die neuen RMS abgespeichert
% Spot1 = TheSystem.Analyses.New_Analysis(ZOSAPI.Analysis.AnalysisIDM.StandardSpot);
% Spot1.ApplyAndWaitForCompletion();
% Spot1_results = Spot1.GetResults();
% 
% n = CENYfirst;
% for i=2:Zeilen
%     Operand = TheMFE.GetOperandAt(n);
%     FieldCENY = Operand.GetCellAt(4);                                                           % für die korrekte Zuordnung werden die Fields aus den CENY operanden ausgelesen, da die Tabelle ihren Ursprung in den CENY Werten hat
%                                                                                                 % GetCellAt(4) ist die 4. Spalte des CENY Operanden, in der die Fields stehen
%     Analyse{i, 8} = Spot1_results.SpotData.GetRMSSpotSizeFor(FieldCENY.IntegerValue,1);         % hier werden die RMS Spot Ergebnisse geladen und direkt in die AusgabeExcel Datei geschrieben
%     n = n +1;                                                                                   % die Results werden nur temporär für die Ausgabe geladen!
% end
% 
% Spot1.Close;
% 
% % Hier wird die Differenz berechnet und in der Tabelle abgespeichert
% n = 1;  % hier wird n ausnahmsweise immmer ab 1 hochgezählt, da keine Abhängigkeit zu CENY
% for i=2:Zeilen  
%     Difference = (100/Analyse{i,7})*Analyse{i,8}-100;        
%     Analyse{i, 9} = Difference;                                   % Die Differenz wird in Spalte 5 gespeichert                               
%     SpalteDifferenz(n,1) = Difference;
%     n = n+1;
% end
% 
% 
% MittelwertRMS = mean(SpalteDifferenz);
% if MittelwertRMS <0
%     disp('Warnung: RMS Spot verschlechtert sich!')
% end


%--------------weiterführende Optimierungen--------------------------------------------------------------------------------------------------------------

% Hier wird die Optimierungszeit erhöht, falls die Veränderungen zu gering waren

NumberOfMissingChange = 0;

while NumberOfMissingChange < 3                                 % Ab 3 Durchläufen ohne Veränderungen wird abgebrochen
    
    if size(MF,1) > 2
    Diff = MF{size(MF,1),2}-MF{size(MF,1)-1,2};
    prozentDiff = MF{size(MF,1),2}/100*Diff;
    end

    if prozentDiff < 10                         
        Faktor = Faktor + 1;
        OptTimeInSeconds = StartTime*Faktor;
    end
    
    if OptTimeInSeconds > MaxOptTimeInSeconds                  % limitiert die Maximale Optimierungszeit
        OptTimeInSeconds = MaxOptTimeInSeconds;
    end
    

    % Hammer Optimierung
    HammerOpt = TheSystem.Tools.OpenHammerOptimization();
    HammerOpt.Algorithm = ZOSAPI.Tools.Optimization.OptimizationAlgorithm.DampedLeastSquares;
    HammerOpt.NumberOfCores = Cores;       
    HammerOpt.RunAndWaitWithTimeout(OptTimeInSeconds);
    
    HammerOpt.Cancel();
    HammerOpt.WaitForCompletion();
    HammerOpt.Close();

    MF_Value = TheMFE.CalculateMeritFunction;
    fprintf('Hammer Optimization for %i seconds...\n', OptTimeInSeconds);
    fprintf('Current Merit Function %i \n', MF_Value);

    % Hier wird der MF-Value abgespeichert
    NumberOfOpt = NumberOfOpt + 1;                            % Variable, die zählt wie oft eine optimierung durchgeführt wurde
    MF{NumberOfOpt+1,1} = NumberOfOpt;
    MF{NumberOfOpt+1,2} = TheMFE.CalculateMeritFunction;
    MF{NumberOfOpt+1,3} = OptTimeInSeconds;
    
    if size(MF,1) > 2
    Diff = MF{size(MF,1),2}-MF{size(MF,1)-1,2};
    prozentDiff = MF{size(MF,1),2}/100*Diff;
    end

       % Hier wird das finale Abbruchkriterium abgefragt
%     if OptTimeInSeconds == MaxOptTimeInSeconds & round(MF{NumberOfOpt+1,2},2) == round(MF{NumberOfOpt,2},2)     % wenn die maximale Zeit erreicht wird, wird angefangen hu zählen,
%        NumberOfMissingChange = NumberOfMissingChange+1;                                     % wie oft sich die MF NICHT verändert
%     end 

    
    % Hier wird das finale Abbruchkriterium abgefragt
    if OptTimeInSeconds == MaxOptTimeInSeconds & prozentDiff < 1     % wenn die maximale Zeit erreicht wird, wird angefangen hu zählen,
       NumberOfMissingChange = NumberOfMissingChange+1;                                     % wie oft sich die MF NICHT verändert
    end

    if TheMFE.CalculateMeritFunction < 0.0002
        NumberOfMissingChange = 4;
    end
    
    %---------------GENF (kopiert von oben)--------------------------------------------------------------------------------------------------------------

    % Hier werden die neuen Werte in die Tabelle gespeichert
    
    TheMFE.CalculateMeritFunction; %aktualisieren der MF-Values
    
    n = 2;
    for i=firstIndex:lastIndex
        Operand = TheMFE.GetOperandAt(i);
        if strcmp(Operand.Type, 'GENF') == true
            Analyse{n, 5} = Operand.Value;                % aktueller Wert des GENF Operanden wird in Spalte 4 gespeichert
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

    %---------------RMS (von oben kopiert)----------------------------------------------------------------------------------------------------------------------------

%     % Hier werden die neuen RMS abgespeichert
%     Spot1 = TheSystem.Analyses.New_Analysis(ZOSAPI.Analysis.AnalysisIDM.StandardSpot);
%     Spot1.ApplyAndWaitForCompletion();
%     Spot1_results = Spot1.GetResults();
% 
%     n = CENYfirst;
%     for i=2:Zeilen
%         Operand = TheMFE.GetOperandAt(n);
%         FieldCENY = Operand.GetCellAt(4);                                                           % für die korrekte Zuordnung werden die Fields aus den CENY operanden ausgelesen, da die Tabelle ihren Ursprung in den CENY Werten hat
%                                                                                                     % GetCellAt(4) ist die 4. Spalte des CENY Operanden, in der die Fields stehen
%         Analyse{i, 8} = Spot1_results.SpotData.GetRMSSpotSizeFor(FieldCENY.IntegerValue,1);         % hier werden die RMS Spot Ergebnisse geladen und direkt in die AusgabeExcel Datei geschrieben
%         n = n +1;                                                                                   % die Results werden nur temporär für die Ausgabe geladen!
%     end
% 
%     Spot1.Close;
% 
%     % Hier wird die Differenz berechnet und in der Tabelle abgespeichert
%     n = 1;  % hier wird n ausnahmsweise immmer ab 1 hochgezählt, da keine Abhängigkeit zu CENY
%     for i=2:Zeilen  
%         Difference = (100/Analyse{i,7})*Analyse{i,8}-100;        
%         Analyse{i, 9} = Difference;                                   % Die Differenz wird in Spalte 5 gespeichert                               
%         SpalteDifferenz(n,1) = Difference;
%         n = n+1;
%     end
% 
% 
%     MittelwertRMS = mean(SpalteDifferenz);
%     if MittelwertRMS <0
%         disp('Warnung: RMS Spot verschlechtert sich!')
%     end
    
    
    
end

TheSystem.SaveAs(System.String.Concat(FolderOutput, 'NIR_nach_globaler_Optimierung.zmx'));

%----------------Ausgabe----------------------------------------------------------------------------------------------------------------

file_path = fullfile(FolderNIR, 'globale_OptimierungNIR.xls');
xlswrite(file_path,Analyse);
winopen(file_path);

file_path = fullfile(FolderVIS, 'MF_globale_OptimierungNIR.xls');
xlswrite(file_path,MF);
winopen(file_path);

end


%----------------------------------------------------------------------
%----------------------------------------------------------------------
%----------------------------------------------------------------------
%----------------------------------------------------------------------
%----------------------------------------------------------------------
% für Matrix-Systeme:

if TheSystem.SystemData.Fields.NumberOfFields ~= 1
    
    [count, firstIndex, lastIndex] = Find_Operand('CENY');     % findet den gesuchten Operanden und gibt Anzahl(count), erste Position in MF(firstIndex) und letzte Position (lastIndex) wider.

Zeilen = count +1;                      % Die Anzahl der Zeilen ist Die Anzahl der Operanden + Zeile für Beschriftung (deswegen +1)
Analyse = cell(Zeilen, 9);              % Die Ausgabetabelle bekommt hier ihr Format: Zeilen(varriert je nach Fields) & 9 Spalten
Analyse{1, 1} = 'Fieldnummer'; 
Analyse{1, 2} = 'Target CENY';
Analyse{1, 3} = 'alte Werte CENY'; 
Analyse{1, 4} = 'neue Werte CENY'; 
Analyse{1, 5} = 'Differenz in % CENY'; 
Analyse{1, 7} = 'alte Werte RMS'; 
Analyse{1, 8} = 'neue Werte RMS';
Analyse{1, 9} = 'Differenz in % RMS'; 

%------------------CENY----------------------------------------------------------------------------------------------------------

% Hier werden die CENY "Targets" und "Values" abgespeichert
n = 2;  %Eintragen der Werte ab Zeile 2, wegen Überschriften

TheMFE.CalculateMeritFunction; %aktualisieren der MF-Values

for i=firstIndex:lastIndex
    Operand = TheMFE.GetOperandAt(i);
    if strcmp(Operand.Type, 'CENY') == true
        FieldCENY = Operand.GetCellAt(5);
        Analyse{n, 1} = FieldCENY.IntegerValue;       
        Analyse{n, 2} = Operand.Target;               
        Analyse{n, 3} = Operand.Value;               
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
            Analyse{n, 7} = Spot1_results.SpotData.GetRMSSpotSizeFor(FieldCENY.IntegerValue,1);               
            n = n+1;
        end
    end   
    Spot1.Close;



%----------------MF---------------------------------------------------------------------------------------------------------------

% Hier wird der MF-Value abgespeichert
MF = cell(4,3);
MF{1,1} = 'Durchlauf';
MF{1,2} = 'Value MF';
MF{1,3} = 'Dauer';
NumberOfOpt = 1;                            % Variable, die zählt wie oft eine optimierung durchgeführt wurde
MF{NumberOfOpt+1,1} = NumberOfOpt;
MF{NumberOfOpt+1,2} = TheMFE.CalculateMeritFunction;
MF{NumberOfOpt+1,3} = OptTimeInSeconds;



%---------------1. Optimierung-------------------------------------------------------------------------------------------

% Hammer Optimierung
HammerOpt = TheSystem.Tools.OpenHammerOptimization();
HammerOpt.Algorithm = ZOSAPI.Tools.Optimization.OptimizationAlgorithm.DampedLeastSquares;
HammerOpt.NumberOfCores = Cores;        

    MF_Value = TheMFE.CalculateMeritFunction;
    fprintf('Hammer Optimization for %i seconds...\n', OptTimeInSeconds);
    fprintf('Initial Merit Function %i \n', MF_Value);

HammerOpt.RunAndWaitWithTimeout(OptTimeInSeconds);
HammerOpt.Cancel();
HammerOpt.WaitForCompletion();
HammerOpt.Close();

%---------------CENY--------------------------------------------------------------------------------------------------------------

% Hier werden die neuen CENY "Values" abgespeichert

TheMFE.CalculateMeritFunction; %aktualisieren der MF-Values

n = 2;  %Eintragen der Werte ab Zeile 2, wegen Überschriften
for i=firstIndex:lastIndex
    Operand = TheMFE.GetOperandAt(i);
    if strcmp(Operand.Type, 'CENY') == true            
        Analyse{n, 4} = Operand.Value;            % die neuen Values werden in Spalte 4 gespeichert   
        n = n+1;
    end
end


% Hier wird die Differenz berechnet und in der Tabelle abgespeichert
n = 1;  % hier wird n ausnahmsweise immmer ab 1 hochgezählt, da keine Abhängigkeit zu CENY
for i=3:Zeilen  
    Difference = abs((100/Analyse{i,2})*Analyse{i,4}-100);        % Differenz wird ohne Vorzeichen berechnet --> abs = Betrag
    Analyse{i, 5} = Difference;                                   % Die Differenz wird in Spalte 5 gespeichert                               
    SpalteDifferenzCENY(n,1) = Difference;
    n = n+1;
end

MittelwertCENY = mean(SpalteDifferenzCENY);

%---------------RMS----------------------------------------------------------------------------------------------------------------------------


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

% Hier wird die Differenz berechnet und in der Tabelle abgespeichert
n = 1;  % hier wird n ausnahmsweise immmer ab 1 hochgezählt, da keine Abhängigkeit zu CENY
for i=2:Zeilen  
    Difference = (100/Analyse{i,7})*Analyse{i,8}-100;        
    Analyse{i, 9} = Difference;                                   % Die Differenz wird in Spalte 5 gespeichert                               
    SpalteDifferenz(n,1) = Difference;
    n = n+1;
end


MittelwertRMS = mean(SpalteDifferenz);
if MittelwertRMS > 0
    disp('Warnung: RMS Spot verschlechtert sich!')
end


%--------------weiterführende Optimierungen--------------------------------------------------------------------------------------------------------------

% Hier wird die Optimierungszeit erhöht, falls die Veränderungen zu gering waren

NumberOfMissingChange = 0;

while NumberOfMissingChange < 3                                 % Ab 3 Durchläufen bei maximaler Zeit ohne Veränderungen wird abgebrochen
    
    if size(MF,1) > 2
    Diff = MF{size(MF,1),2}-MF{size(MF,1)-1,2};
    prozentDiff = MF{size(MF,1),2}/100*Diff;
    end

    if prozentDiff < 10                          
        Faktor = Faktor + 1;
        OptTimeInSeconds = StartTime*Faktor;
    end
    
    if OptTimeInSeconds > MaxOptTimeInSeconds                  % limitiert die Maximale Optimierungszeit
        OptTimeInSeconds = MaxOptTimeInSeconds;
    end
    

    % Hammer Optimierung
     HammerOpt = TheSystem.Tools.OpenHammerOptimization();
     HammerOpt.Algorithm = ZOSAPI.Tools.Optimization.OptimizationAlgorithm.DampedLeastSquares;
     HammerOpt.NumberOfCores = Cores;       
     HammerOpt.RunAndWaitWithTimeout(OptTimeInSeconds);
     
     HammerOpt.Cancel();
     HammerOpt.WaitForCompletion();
     HammerOpt.Close();

    MF_Value = TheMFE.CalculateMeritFunction;
    fprintf('Hammer Optimization for %i seconds...\n', OptTimeInSeconds);
    fprintf('Current Merit Function %i \n', MF_Value);

    % Hier wird der MF-Value abgespeichert
    NumberOfOpt = NumberOfOpt + 1;                            % Variable, die zählt wie oft eine optimierung durchgeführt wurde
    MF{NumberOfOpt+1,1} = NumberOfOpt;
    MF{NumberOfOpt+1,2} = TheMFE.CalculateMeritFunction;
    MF{NumberOfOpt+1,3} = OptTimeInSeconds;
    
    if size(MF,1) > 2
    Diff = MF{size(MF,1),2}-MF{size(MF,1)-1,2};
    prozentDiff = MF{size(MF,1),2}/100*Diff;
    end

       % Hier wird das finale Abbruchkriterium abgefragt
%     if OptTimeInSeconds == MaxOptTimeInSeconds & round(MF{NumberOfOpt+1,2},2) == round(MF{NumberOfOpt,2},2)     % wenn die maximale Zeit erreicht wird, wird angefangen hu zählen,
%        NumberOfMissingChange = NumberOfMissingChange+1;                                     % wie oft sich die MF NICHT verändert
%     end 

    
    % Hier wird das finale Abbruchkriterium abgefragt
    if OptTimeInSeconds == MaxOptTimeInSeconds & prozentDiff < 1     % wenn die maximale Zeit erreicht wird, wird angefangen hu zählen,
       NumberOfMissingChange = NumberOfMissingChange+1;                                     % wie oft sich die MF NICHT verändert
    end
    
    if TheMFE.CalculateMeritFunction < 0.4
        NumberOfMissingChange = 4;
    end
    
    %---------------CENY (kopiert von oben)--------------------------------------------------------------------------------------------------------------

    % Hier werden die neuen CENY "Values" abgespeichert
    n = 2;  %Eintragen der Werte ab Zeile 2, wegen Überschriften
    for i=firstIndex:lastIndex
        Operand = TheMFE.GetOperandAt(i);
        if strcmp(Operand.Type, 'CENY') == true            
            Analyse{n, 4} = Operand.Value;            % die neuen Values werden in Spalte 4 gespeichert   
            n = n+1;
        end
    end


    % Hier wird die Differenz berechnet und in der Tabelle abgespeichert
    n = 1;  % hier wird n ausnahmsweise immmer ab 1 hochgezählt, da keine Abhängigkeit zu CENY
    for i=3:Zeilen  
        Difference = abs((100/Analyse{i,2})*Analyse{i,4}-100);        % Differenz wird ohne Vorzeichen berechnet --> abs = Betrag
        Analyse{i, 5} = Difference;                                   % Die Differenz wird in Spalte 5 gespeichert                               
        SpalteDifferenzCENY(n,1) = Difference;
        n = n+1;
    end

    MittelwertCENY = mean(SpalteDifferenzCENY);
    
    %---------------RMS (von oben kopiert)----------------------------------------------------------------------------------------------------------------------------

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

    % Hier wird die Differenz berechnet und in der Tabelle abgespeichert
    n = 1;  % hier wird n ausnahmsweise immmer ab 1 hochgezählt, da keine Abhängigkeit zu CENY
    for i=2:Zeilen  
        Difference = (100/Analyse{i,7})*Analyse{i,8}-100;        
        Analyse{i, 9} = Difference;                                   % Die Differenz wird in Spalte 5 gespeichert                               
        SpalteDifferenz(n,1) = Difference;
        n = n+1;
    end


    MittelwertRMS = mean(SpalteDifferenz);
    if MittelwertRMS > 0
        disp('Warnung: RMS Spot verschlechtert sich!')
    end
    
    
    
end

TheSystem.SaveAs(System.String.Concat(FolderOutput, 'NIR_nach_globaler_Optimierung.zmx'));

%----------------Ausgabe----------------------------------------------------------------------------------------------------------------

file_path = fullfile(FolderNIR, 'globale_OptimierungNIR.xls');
xlswrite(file_path,Analyse);
winopen(file_path);

file_path = fullfile(FolderNIR, 'MF_globale_OptimierungNIR.xls');
xlswrite(file_path,MF);
winopen(file_path);
    
    
end