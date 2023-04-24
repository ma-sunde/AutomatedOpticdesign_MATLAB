%                                                                              %
%                                                                              %
%       ! Zuerst OpticStudio öffnen und im Reiter Programming auf den          %
%                Button Interactive Extension klicken !                        %                     
%                   Danach dieses Skript starten.                              %
%                                                                              %
%                                                                              %

%DATEN FÜR LICHTQUELLE WERDEN HIER GESETZT. Power, Größe der Matrix, etc.
%Für Punktlichtquelle von Zeile 110-114
%Für Matrix-Systeme von Zeile 219 - 224

System_Load;



%für Punktlichtquellensysteme:
if TheSystem.SystemData.Fields.NumberOfFields == 1 
System = 'Punkt';
    
    
% Tabelle für Ausgabe erstellen und CENY Werte schreiben
Zeilen = size(GENF_Values,1);
AusgabeExcel = cell(Zeilen, 4);         % Die Ausgabetabelle bekommt hier ihr Format: Zeilen(varrieren) & 4 Spalten
AusgabeExcel{1, 1} = 'Fieldnummer';     % Die 1. Zeile ist nur für Überschriften     
AusgabeExcel{1, 2} = 'GENF in µm';      
AusgabeExcel{1, 3} = 'GENF Zielwert (zw. 0 und 1)';
AusgabeExcel{1, 4} = 'GENF aktueller Wert';
AusgabeExcel{1, 5} = 'Wirkungsgrad';    

% Hier werden die Zielwerte des GENF Operanden temporär geladen und in die AusgabeExcel Tabelle geschrieben

n = 1;
for i=2:size(GENF_Values,1)
    AusgabeExcel{i, 1} = n;
    AusgabeExcel{i, 2} = GENF_Values{i,4}*(10^3);
    AusgabeExcel{i, 3} = (i-1)*(1/(size(GENF_Values,1)-1));
    n = n+1;
end

% Hier werden die neuen GENF Werte in die AusgabeExcel Tabelle geschrieben
[count, firstIndex, lastIndex] = Find_Operand('GENF');
TheSystem.UpdateStatus;     %zum Auslesen der aktuellen Values im MFE muss das System zuerst aktualisiert werden
TheMFE.CalculateMeritFunction;

n = 3;  %Überschrift und erster nicht beachteter GENF Operand mit Radius = 0 ergibt Startposition n=3
for i=firstIndex:lastIndex
    Operand = TheMFE.GetOperandAt(i);
    if strcmp(Operand.Type, 'GENF') == true
        AusgabeExcel{n, 4} = Operand.Value;                % aktueller Wert des GENF Operanden wird in Spalte 4 gespeichert
        n = n+1;
    end
end

% RMS schwierig bei nur einem Field... 
% % RMS Spot
% Spot1 = TheSystem.Analyses.New_Analysis(ZOSAPI.Analysis.AnalysisIDM.StandardSpot);
% Spot1.ApplyAndWaitForCompletion();
% Spot1_results = Spot1.GetResults();
% 
% % Hier werden die Ergebnisse des Spot Diagramms temporär geladen und in die AusgabeExcel Tabelle geschrieben
% 
% n = 1;
% for i=2:Zeilen
%     AusgabeExcel{i, 3} = Spot1_results.SpotData.GetRMSSpotSizeFor(n,1);     % hier werden die RMS Spot Ergebnisse geladen und direkt in die AusgabeExcel Datei geschrieben
%     n = n+1;                                                                % die Results werden nur temporär für die Ausgabe geladen!
% end
% 
% Spot1.Close;

%Entfernung von Linse 1 und Lichtquelle muss separat übertragen werden
Startposition = TheLDE.GetSurfaceAt(0).Thickness;

%----------------------Convert to non-sequential mode ---------------------------

%--- Convert file to Non-sequential mode
convertNSmode = TheSystem.Tools.OpenConvertToNSCGroup();
convertNSmode.ConvertFileToNSC = true;
convertNSmode.RunAndWaitForCompletion();
convertNSmode.Close();


%--- Einheiten festlegen
TheSystemData = TheSystem.SystemData;
TheSystemData.Units.SourceUnits = ZOSAPI.SystemData.ZemaxSourceUnits.Lumens;           % Lichtquellen Einheit: Lumen
TheSystemData.Units.AnalysisUnits = ZOSAPI.SystemData.ZemaxAnalysisUnits.WattsPerMSq;  % Analyse Einheit: W/m²


% NSC Editor aufrufen:
TheNCE = TheSystem.NCE;

% Objekte laden & einfügen:
    Count_NSC_Objects;              % findet die Position der Surfaces, um im Anschluss die 
                                    % Lichtquelle davor und die Detektor Ebene danach einzufügen.
                                    % Erstellt SurfacePosfirst und SurfacePoslast
NullObject = TheNCE.InsertNewObjectAt(SurfacePosfirst);
SurfacePosfirst = SurfacePosfirst+1;                        % muss erhöht werden immer wenn ein neues Objekt davor eingefügt wird!
SurfacePoslast = SurfacePoslast+1;                          % muss erhöht werden immer wenn ein neues Objekt davor eingefügt wird!
ObjSource = TheNCE.InsertNewObjectAt(SurfacePosfirst);      % Lichtquelle
SurfacePoslast = SurfacePoslast+1;                          % muss erhöht werden immer wenn ein neues Objekt davor eingefügt wird!
ObjDetector = TheNCE.InsertNewObjectAt(SurfacePoslast+1);   % Detector Ebene wird eingefügt hinter der letzten Linsen Oberfläche

% Definieren des Null Objekts

NullObject.ZPosition = 0 - Startposition;               % Entfernung der LQ zur 1. Linse

%------------- Definieren der Objekte im nicht-sequentiellen Modus --------

%---------------DATEN FÜR LICHTQUELLE HIER EINGEBEN----------------------

%--- Lichtquelle
ObjSource.ChangeType(ObjSource.GetObjectTypeSettings(ZOSAPI.Editors.NCE.ObjectType.SourcePoint));
ObjSource.ObjectData.Power = 310*10^(-3);            % Power der Lichtquelle in Lumen
ObjSource.ObjectData.ConeAngle = 7;                  % Abstrahlwinkel 14 Grad
ObjSource.ObjectData.WaveNumber = 1;                 % 
ObjSource.ObjectData.NumberOfLayoutRays = 10;        % displayed Rays
ObjSource.ObjectData.NumberOfAnalysisRays = 1e+06;   % Rays für Berechnung
        
%--- Detector Ebene
ObjDetector.ChangeType(ObjDetector.GetObjectTypeSettings(ZOSAPI.Editors.NCE.ObjectType.DetectorRectangle));   
ObjDetector.ObjectData.Row.RefObject = SurfacePoslast;   % Das Referenz Objekt wird auf die letzte Linse festgelegt
ObjDetector.ObjectData.Row.ZPosition = 25000;            % Dadurch ist die Detektor Ebene immer 25 Meter von der letzten Oberfläche entfernt
ObjDetector.ObjectData.NumberXPixels = 1920;           
ObjDetector.ObjectData.NumberYPixels = 1080;
ObjDetector.ObjectData.XHalfWidth = 20977.49;       %Entspricht bei 25m Entfernung einem halben Öffnungswinkel von 40° Horizontal
ObjDetector.ObjectData.YHalfWidth = 6698.73;        %Entspricht bei 25m Entfernung einem halben Öffnungswinkel von 15° Vertikal

%-------------------- Wirkungsgrad berechnen -----------------------------

%--- Setup and run the ray trace
NSCRayTrace = TheSystem.Tools.OpenNSCRayTrace();
NSCRayTrace.SplitNSCRays = false;
NSCRayTrace.ScatterNSCRays = true;
NSCRayTrace.UsePolarization = true;
NSCRayTrace.IgnoreErrors = true;
NSCRayTrace.SaveRays = false;
NSCRayTrace.ClearDetectors(0);
    
NSCRayTrace.RunAndWaitForCompletion();
NSCRayTrace.Close();

%--- Detector Viewer Analysis 
detector = TheSystem.Analyses.New_Analysis(ZOSAPI.Analysis.AnalysisIDM.DetectorViewer);
detector.ApplyAndWaitForCompletion();

detector_Results = detector.GetResults();
headerData = detector_Results.HeaderData;
peakIrradiance = string(headerData.Lines(3));
peakIrradiance = str2double(extractBefore(extractAfter(string(headerData.Lines(3)), ': '), ' Lumens/M^2'));
totalPower = str2double(extractBefore(extractAfter(string(headerData.Lines(4)), ': '), ' Lumens'));
Wirkungsgrad = totalPower/(ObjSource.ObjectData.Power*10000);        % noch unsexy. Ist totalpower noch ein string? I dont get it
detector.Close;

    AusgabeExcel{2, 5} = Wirkungsgrad;      % Hier wird der Wirkungsgrad in die AusgabeExcel Tabelle geschrieben

% Ausgabe
file_path = fullfile(FolderNIR, 'AusgabeNIR.xls');
xlswrite(file_path,AusgabeExcel);
%winopen(file_path);


end

%für Matrixsysteme:
if TheSystem.SystemData.Fields.NumberOfFields ~= 1 
System = 'Matrix';
ApertureValue = SystExplorer.Aperture.ApertureValue;
% determine maximum field in Y only
        max_field = TheSystem.SystemData.Fields.GetField(TheSystem.SystemData.Fields.NumberOfFields).Y;    
% Anzahl der Fields
NOF = TheSystem.SystemData.Fields.NumberOfFields;
NumberOfFields = double(NOF);   %Funktion gibt NumberOfFields als Integer zurück. Muss vorher umgewandelt werden!
        
        
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

%---------------DATEN FÜR LICHTQUELLE HIER EINGEBEN----------------------

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
file_path = fullfile(FolderNIR, 'AusgabeNIR.xls');
xlswrite(file_path,AusgabeExcel);
%winopen(file_path);
    
    
    
    
end

