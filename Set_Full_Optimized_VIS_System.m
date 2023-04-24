%                                                                              %
%                                                                              %
%       ! Zuerst OpticStudio öffnen und im Reiter Programming auf den          %
%                Button Interactive Extension klicken !                        %                     
%                   Danach dieses Skript starten.                              %
%                                                                              %
%  


System_Load;

%Switch back to sequential mode
TheSystem.MakeSequential

%Den non-sequential Editor minimieren
TheNCE.HideEditor

%--------------------- Linsendaten übertragen ---------------------------
Insert_Zaehler = 0;
for i=0:Nummer
    if i == 0                                   % Die Objekt Thickness ist nicht Teil des Cell Arrays "Oberflaechen"                                  
        Oberflaeche = TheLDE.GetSurfaceAt(i);
        Oberflaeche.Thickness = ThicknessObject;
    elseif Insert_Zaehler == 2                  % Stop Surface Standardmäig an Pos 3
        Oberflaeche = TheLDE.GetSurfaceAt(i);   % Wenn 2 Flächen eingefügt wurden, muss die Position 
        OddAsphere = Oberflaeche.GetSurfaceTypeSettings(ZOSAPI.Editors.LDE.SurfaceType.OddAsphere);
        Oberflaeche.ChangeType(OddAsphere);
               
        Insert_Zaehler = Insert_Zaehler + 1;
        i = i+1;
            
    elseif i ~= 0 && i ~= Nummer
        TheLDE.InsertNewSurfaceAt(i);
        Insert_Zaehler = Insert_Zaehler +1;
        Oberflaeche = TheLDE.GetSurfaceAt(i);
        OddAsphere = Oberflaeche.GetSurfaceTypeSettings(ZOSAPI.Editors.LDE.SurfaceType.OddAsphere);
        Oberflaeche.ChangeType(OddAsphere); 
        
    elseif i == Nummer
           
    end
    
end

for i=1:Nummer-1
    Oberflaeche = TheLDE.GetSurfaceAt(i);
     % --------- Parameter zuweisen mit readCell
     % readCell durchsucht ein Array nach Oberflächennummer und Parameter,
     % dem man der Funktion gibt. Man erhält als Antwort die Position der
     % gesuchten Werte in dem Cell Array
        [num_pos, param_pos] = readCell(Oberflaechen,i,"Thickness");
        Oberflaeche.Thickness = Oberflaechen{param_pos,num_pos};
        [num_pos, param_pos] = readCell(Oberflaechen,i,"Radius");
        Oberflaeche.Radius = Oberflaechen{param_pos,num_pos};
        [num_pos, param_pos] = readCell(Oberflaechen,i,"Clear Semi-Diameter");
        Oberflaeche.SemiDiameter = Oberflaechen{param_pos,num_pos};
        [num_pos, param_pos] = readCell(Oberflaechen,i,"Mech Semi-Diameter");
        Oberflaeche.MechanicalSemiDiameter = Oberflaechen{param_pos,num_pos};
        [num_pos, param_pos] = readCell(Oberflaechen,i,"Comment");
        Oberflaeche.Comment = Oberflaechen{param_pos,num_pos};
        
    if Oberflaeche.Comment == 'Vorderseite'                         % Das Material wird nur an der Vorderseite übergeben
        [num_pos, param_pos] = readCell(Oberflaechen,i,"Material");
        Oberflaeche.Material = Oberflaechen{param_pos,num_pos};
    end
       
        [num_pos, param_pos] = readCell(Oberflaechen,i,"Conic");
        Oberflaeche.Conic = Oberflaechen{param_pos,num_pos};
        [num_pos, param_pos] = readCell(Oberflaechen,i,"1st Order Term");
        Oberflaeche.GetCellAt(12).DoubleValue = Oberflaechen{param_pos,num_pos}; 
        [num_pos, param_pos] = readCell(Oberflaechen,i,"2nd Order Term");
        Oberflaeche.GetCellAt(13).DoubleValue = Oberflaechen{param_pos,num_pos}; 
        [num_pos, param_pos] = readCell(Oberflaechen,i,"3rd Order Term");
        Oberflaeche.GetCellAt(14).DoubleValue = Oberflaechen{param_pos,num_pos};
        [num_pos, param_pos] = readCell(Oberflaechen,i,"4th Order Term");
        Oberflaeche.GetCellAt(15).DoubleValue = Oberflaechen{param_pos,num_pos};
        [num_pos, param_pos] = readCell(Oberflaechen,i,"5th Order Term");
        Oberflaeche.GetCellAt(16).DoubleValue = Oberflaechen{param_pos,num_pos};
        [num_pos, param_pos] = readCell(Oberflaechen,i,"6th Order Term");
        Oberflaeche.GetCellAt(17).DoubleValue = Oberflaechen{param_pos,num_pos};
        [num_pos, param_pos] = readCell(Oberflaechen,i,"7th Order Term");
        Oberflaeche.GetCellAt(18).DoubleValue = Oberflaechen{param_pos,num_pos};
        [num_pos, param_pos] = readCell(Oberflaechen,i,"8th Order Term");
        Oberflaeche.GetCellAt(19).DoubleValue = Oberflaechen{param_pos,num_pos};
end

%Oberflaeche.SemiDiameterCell.CreateSolveType(ZOSAPI.Editors.SolveType.Fixed)
%Oberflaeche.MechanicalSemiDiameterCell.CreateSolveType(ZOSAPI.Editors.SolveType.Fixed)


% ---- Image Ebende

Oberflaeche = TheLDE.GetSurfaceAt(TheLDE.NumberOfRows);
Oberflaeche.SemiDiameter = SizeImage;
Oberflaeche.MechanicalSemiDiameter = SizeImage;


% --------------- Lichtquelle definieren ---------------------------------

% Apertur definieren
SystExplorer = TheSystem.SystemData;
SystExplorer.Aperture.ApertureType = ZOSAPI.SystemData.ZemaxApertureType.FloatByStopSize;

% LED Array erstellen mit beispielhaft 10 Fields
SystExplorer.Fields.SetFieldType(ZOSAPI.SystemData.FieldType.ObjectHeight);
% wie geht das???
%SystExplorer.Fields.SetFieldNormalizationType(ZOSAPI.SystemData.FieldNormalizationType.Rectangular)

SystExplorer.Fields.AddField(0, 0.314, 1.0);
SystExplorer.Fields.AddField(0, 0.629, 1.0);
SystExplorer.Fields.AddField(0, 0.943, 1.0);
SystExplorer.Fields.AddField(0, 1.257, 1.0);
SystExplorer.Fields.AddField(0, 1.571, 1.0);
SystExplorer.Fields.AddField(0, 1.886, 1.0);
SystExplorer.Fields.AddField(0, 2.200, 1.0);
SystExplorer.Fields.AddField(0, 2.514, 1.0);
SystExplorer.Fields.AddField(0, 2.828, 1.0);