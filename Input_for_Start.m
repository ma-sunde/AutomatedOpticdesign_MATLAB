%Hier werden vor dem Start die Randbedingungen festgelegt

%------------------------- Allgemein --------------------------------------

%Speicherort für Simulationsergebnisse festlegen
%Gib den Pfad für die Ergebnisse in der Variable 'FolderOutput' an:
    FolderOutput = System.String.Concat('D:\MA_Hanisch\Matlab\ZemaxOutput\');    %zB.: 'D:\Dokumente\Maschbau Studium\Master\Masterarbeit\Simulation';
        


%------------------------ VIS - System ------------------------------------

%---Startsystem

%Möchtest du ein eigenes VIS-Startsystem verwenden? 
%--> Ja:
    %Gib den Pfad für dein Startsystem in der Variable 'FolderVIS' an:
    FolderVIS = 'D:\Dokumente\Maschbau Studium\Master\Masterarbeit\Simulation\';    %zB.: 'D:\Dokumente\Maschbau Studium\Master\Masterarbeit\Simulation';
    
    %Gib den Namen der .zmx-Datei in der Variable 'FileVIS' an:
    FileVIS = 'VIS_Startsystem_V8_Cone Angle.zmx';      %zB.: '\NIR_Startsystem_GENF_OPVA_1Zoll_Linsen.zmx';
    
%--> Nein: alle Variablen leer lassen mit  = '';

%---Lichtverteilung

    %Gib den Pfad für die Lichtverteilung des VIS-System in der Variable
    %'FolderVIS_Lichtverteilung' an:
    FolderVIS_Lichtverteilung = 'D:\Dokumente\Maschbau Studium\Master\Masterarbeit\Daten\';   %zB.: 'D:\Dokumente\Maschbau Studium\Master\Masterarbeit\Daten\'

    
    %Gib den Namen der Datei für die Lichtverteilung des VIS-System in der Variable
    %'FileVIS_Lichtverteilung' an:
    
    FileVIS_Lichtverteilung = 'Lichtverteilung_idealisiert.xlsx';

%---Ergebnisse

    %Sequential Mode
        %Gib den namen für die Ergebnis Datei des VIS-Systems im sequential Mode an:
        FileVIS_LDE = 'VIS_Matlab_Sequ.zmx';
        
    %Non-Sequential Mode
        %Gib den namen für die Ergebnis Datei des VIS-Systems im non-sequential Mode an:
        FileVIS_NCE = 'VIS_Matlab_Non_Sequ.zmx';
    
%----------------------- NIR - System -------------------------------------

%---Startsystem

%Möchtest du ein eigenes NIR-Startsystem verwenden?
%--> Ja:
    %Gib den Pfad für dein Startsystem in der Variable 'FolderNIR' an:
    FolderNIR = FolderVIS;      %zB.: 'D:\Dokumente\Maschbau Studium\Master\Masterarbeit\Simulation\';
    
    %Gib den Namen der .zmx-Datei in der Variable 'FileVIS' an:
    FileNIR = '';               %zB.: 'NIR_Startsystem_GENF_OPVA_1Zoll_Linsen.zmx';

%--> Nein: alle Variablen leer lassen mit  = '';

%---Lichtverteilung

    %Gib den Pfad für die Lichtverteilung des NIR-System in der Variable
    %'FolderNIR_Lichtverteilung' an:
    FolderNIR_Lichtverteilung = 'D:\Dokumente\Maschbau Studium\Master\Masterarbeit\Daten\';   %zB.: 'D:\Dokumente\Maschbau Studium\Master\Masterarbeit\Daten\'

    
    %Gib den Namen der Datei für die Lichtverteilung des NIR-System in der Variable
    %'FileNIR_Lichtverteilung' an:
    
    FileNIR_Lichtverteilung = 'Lichtverteilung_NIR.xlsx';

%---Ergebnisse

    %Sequential Mode
        %Gib den Namen für die Ergebnis Datei des NIR-Systems im sequential Mode an:
        FileNIR_LDE = 'NIR_Matlab_Sequ.zmx';
    
    %Non-Sequential Mode
        %Gib den Namen für die Ergebnis Datei des NIR-Systems im non-sequential Mode an:
        FileNIR_NCE = 'NIR_Matlab_Non_Sequ.zmx';