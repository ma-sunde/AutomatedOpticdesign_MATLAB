%                                                                              %
%                                                                              %
%       ! Zuerst OpticStudio öffnen und im Reiter Programming auf den          %
%                Button Interactive Extension klicken !                        %                     
%                   Danach dieses Skript starten.                              %
%                                                                              %
%                                                                              %

System_Load;


%Switch back to sequential mode
TheSystem.MakeSequential

%Den non-sequential Editor minimieren
TheNCE.HideEditor;
 
%--------------------- Linsendaten übertragen ---------------------------
Insert_Zaehler = 0;
for i=0:Nummer
    if i == 0                                   % Die Objekt Thickness ist nicht Teil des Cell Arrays "Oberflaechen"                                  
        Oberflaeche = TheLDE.GetSurfaceAt(i);
        Oberflaeche.Thickness = Startposition;
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



% ---- Image Ebende

% funktioniert nicht bei Image Ebene?! Why? Fehlermeldung gibt es aber auch
% nicht. 
Oberflaeche = TheLDE.GetSurfaceAt(TheLDE.NumberOfRows);
Oberflaeche.SemiDiameter = SizeImage;
Oberflaeche.MechanicalSemiDiameter = SizeImage;


% --------------- Lichtquelle definieren ---------------------------------

if strcmp(System,'Punkt') == 1

% Apertur definieren
SystExplorer = TheSystem.SystemData;
SystExplorer.Aperture.ApertureType = ZOSAPI.SystemData.ZemaxApertureType.ObjectConeAngle;
SystExplorer.Aperture.ApertureValue = 14;

elseif strcmp(System,'Matrix') == 1

% Apertur definieren

%die neuen Fields werden gesetzt auf Grundlage der Größe der Variable
        %'NumberOfFields' , die variieren kann. Die Positionen der Fields werden
        %verteilt von Y=0 bis zur größten vorhandenen Field Position Y=max_field
        for i=1:NumberOfFields-1
            SystExplorer.Fields.AddField(0, (i)*(max_field/(NumberOfFields-1)), 1.0);
            if i == NumberOfFields-1
                SystExplorer.Fields.DeleteFieldAt(2);
            end
        end

SystExplorer = TheSystem.SystemData;
SystExplorer.Aperture.ApertureType = ZOSAPI.SystemData.ZemaxApertureType.ObjectConeAngle;
SystExplorer.Aperture.ApertureValue = ApertureValue;    
    
end


