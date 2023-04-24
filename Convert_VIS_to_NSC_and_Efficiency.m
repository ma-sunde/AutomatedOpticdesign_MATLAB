%                                                                              %
%                                                                              %
%       ! Zuerst OpticStudio öffnen und im Reiter Programming auf den          %
%                Button Interactive Extension klicken !                        %                     
%                   Danach dieses Skript starten.                              %
%                                                                              %
%                                                                              %

System_Load;

% if isempty(FileVIS) == 0 && isempty(FolderVIS) == 0
%     DGfile = System.String.Concat(FolderVIS, FileVIS);
%     TheSystem.LoadFile(DGfile, false);
% end

% Anzahl Fields, falls später benötigt
% NumberOfFields = TheSystem.SystemData.Fields.NumberOfFields

% Merit Function CENY Operanden finden
FindCENY;                               % Es entstehen Varibalen AnzahlCENY, CENYfirst, CEMYlast, die die Position der Operanden enthalten

% Tabelle für Ausgabe erstellen und CENY Werte schreiben
Zeilen = AnzahlCENY +1;                 % Die Anzahl der Zeilen ist Die Anzahl der Operanden + Zeile für Beschriftung (deswegen +1)
AusgabeExcel = cell(Zeilen, 4);         % Die Ausgabetabelle bekommt hier ihr Format: Zeilen(varriert je nach Fields) & 4 Spalten
AusgabeExcel{1, 1} = 'Fieldnummer';     % Die 1. Zeile ist nur für Überschriften     
AusgabeExcel{1, 2} = 'CENY in mm';      % Es existieren 9 Fields + Überschrift = 10 Zeilen
AusgabeExcel{1, 3} = 'RMS in µm';       % Fieldnummer + CENY + RMS + Wirkungsgrad = 4 Spalten
AusgabeExcel{1, 4} = 'Wirkungsgrad';    
    AusgabeCENY;                        % Hier werden die Values des CENY Operanden temporär geladen und in die AusgabeExcel Tabelle geschrieben



% RMS Spot
Spot1 = TheSystem.Analyses.New_Analysis(ZOSAPI.Analysis.AnalysisIDM.StandardSpot);
Spot1.ApplyAndWaitForCompletion();
Spot1_results = Spot1.GetResults();
    AusgabeRMS;                         % Hier werden die Ergebnisse des Spot Diagramms temporär geladen und in die AusgabeExcel Tabelle geschrieben
Spot1.Close;


% Convert file to Non-sequential mode
convertNSmode = TheSystem.Tools.OpenConvertToNSCGroup();
convertNSmode.ConvertFileToNSC = true;
convertNSmode.RunAndWaitForCompletion();
convertNSmode.Close();


% Einheiten festlegen
TheSystemData = TheSystem.SystemData;
TheSystemData.Units.SourceUnits = ZOSAPI.SystemData.ZemaxSourceUnits.Lumens;          % Lichtquellen Einheit: Lumen
TheSystemData.Units.AnalysisUnits = ZOSAPI.SystemData.ZemaxAnalysisUnits.WattsPerMSq;  % Analyse Einheit: W/m²


% NSC Editor aufrufen:
TheNCE = TheSystem.NCE;

% Objekte laden & einfügen:
    Count_NSC_Objects;              % findet die Position der Surfaces, um im Anschluss die 
                                    % Lichtquelle davor und die Detektor Ebene danach einzufügen.
                                    % Erstellt SurfacePosfirst und SurfacePoslast
                                    
ObjSource = TheNCE.InsertNewObjectAt(SurfacePosfirst);      % Lichtquelle
SurfacePoslast = SurfacePoslast+1;                          % muss erhöht werden immer wenn ein neues Objekt davor eingefügt wird!
ObjDetector = TheNCE.InsertNewObjectAt(SurfacePoslast+1);   % Detector Ebene wird eingefügt hinter der letzten Linsen Oberfläche

% Definieren der Objekte im nicht-sequentiellen Modus
% Lichtquelle
ObjSource.ChangeType(ObjSource.GetObjectTypeSettings(ZOSAPI.Editors.NCE.ObjectType.SourceRectangle));
ObjSource.ObjectData.Power = 5120;                   % Power der Lichtquelle in Lumen
ObjSource.ObjectData.XHalfWidth = 2;                 % halbe Länge in X-Richtung in mm
ObjSource.ObjectData.YHalfWidth = 2;                 % halbe Länge in Y-Richtung in mm
ObjSource.ObjectData.CosineExponent = 1;             % Abstrahlwinkel 90 Grad
ObjSource.ObjectData.NumberOfLayoutRays = 10;        % displayed Rays
ObjSource.ObjectData.NumberOfAnalysisRays = 1e+06;   % Rays für Berechnung
        
% Detector Ebene
ObjDetector.ChangeType(ObjDetector.GetObjectTypeSettings(ZOSAPI.Editors.NCE.ObjectType.DetectorRectangle));   
ObjDetector.ObjectData.Row.RefObject = SurfacePoslast;   % Das Refernez Objekt wird auf die letzte Linse festgelegt
ObjDetector.ObjectData.Row.ZPosition = 25000;            % Dadurch ist die Detektor Ebene immer 25 Meter von der letzten Oberfläche entfernt
ObjDetector.ObjectData.NumberXPixels = 1920;           
ObjDetector.ObjectData.NumberYPixels = 1080;
ObjDetector.ObjectData.XHalfWidth = 20977.49;       %Entspricht bei 25m Entfernung einem halben Öffnungswinkel von 40° Horizontal
ObjDetector.ObjectData.YHalfWidth = 6698.73;        %Entspricht bei 25m Entfernung einem halben Öffnungswinkel von 15° Vertikal



% Setup and run the ray trace
NSCRayTrace = TheSystem.Tools.OpenNSCRayTrace();
NSCRayTrace.SplitNSCRays = false;
NSCRayTrace.ScatterNSCRays = true;
NSCRayTrace.UsePolarization = true;
NSCRayTrace.IgnoreErrors = true;
NSCRayTrace.SaveRays = false;
NSCRayTrace.ClearDetectors(0);
    
NSCRayTrace.RunAndWaitForCompletion();
NSCRayTrace.Close();

% Detector Viewer Analysis 
detector = TheSystem.Analyses.New_Analysis(ZOSAPI.Analysis.AnalysisIDM.DetectorViewer);
detector.ApplyAndWaitForCompletion();

detector_Results = detector.GetResults();
headerData = detector_Results.HeaderData;
peakIrradiance = string(headerData.Lines(3));
peakIrradiance = str2double(extractBefore(extractAfter(string(headerData.Lines(3)), ': '), ' Lumens/M^2'));
totalPower = str2double(extractBefore(extractAfter(string(headerData.Lines(4)), ': '), ' Lumens'));
Wirkungsgrad = totalPower/(ObjSource.ObjectData.Power*10000);        % noch unsexy. Ist totalpower noch ein string? I dont get it
detector.Close;

    AusgabeExcel{2, 4} = Wirkungsgrad;      % Hier wird der Wirkungsgrad in die AusgabeExcel Tabelle geschrieben

% Ausgabe
file_path = fullfile(FolderVIS, 'AusgabeVIS.xls');
xlswrite(file_path,AusgabeExcel);
%winopen(file_path);



