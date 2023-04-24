%                                                                              %
%                                                                              %
%       ! Zuerst OpticStudio öffnen und im Reiter Programming auf den          %
%                Button Interactive Extension klicken !                        %                     
%                   Danach dieses Skript starten.                              %
%                                                                              %
%                                                                              %

System_Load;

% ----------- Abspeichern der Linsendaten ---------------------------------
% Achtung, geht nur mit Odd Asphere als Surface Type!

% Auslesen der Parameter aus dem NSC-Editor

OberflaechenDatenTemp = {}; % das cell Array "OberflaechenDatenTemp ist eine Art Blaupause, die
                            % immer wieder neu befüllt wird. Alle
                            % Linsendaten werden später in dem Array
                            % "Oberflaechen" abgespeichert
 

OberflaechenDatenTemp{1,1} = 'Oberfläche:';
OberflaechenDatenTemp{2,1} = 'Surface Type';
OberflaechenDatenTemp{3,1} = 'Comment';
OberflaechenDatenTemp{4,1} = 'Radius';
OberflaechenDatenTemp{5,1} = 'Thickness';
OberflaechenDatenTemp{6,1} = 'Material';
OberflaechenDatenTemp{7,1} = 'Coating';
OberflaechenDatenTemp{8,1} = 'Clear Semi-Diameter';
OberflaechenDatenTemp{9,1} = 'Chip Zone';
OberflaechenDatenTemp{10,1} = 'Mech Semi-Diameter';
OberflaechenDatenTemp{11,1} = 'Conic';
OberflaechenDatenTemp{12,1} = '1st Order Term';
OberflaechenDatenTemp{13,1} = '2nd Order Term';
OberflaechenDatenTemp{14,1} = '3rd Order Term';
OberflaechenDatenTemp{15,1} = '4th Order Term';
OberflaechenDatenTemp{16,1} = '5th Order Term';
OberflaechenDatenTemp{17,1} = '6th Order Term';
OberflaechenDatenTemp{18,1} = '7th Order Term';
OberflaechenDatenTemp{19,1} = '8th Order Term';



n=1;
Nummer = 1;
for i=TheLDE.NumberOfSurfaces-5:TheLDE.NumberOfSurfaces-2   %Die letzten 2 Linsen sind im System die letzten 5 Oberflächen ausgenommen der letzten Oberfläche (Image Ebene in 25m)
    
    % Im nicht-sequentiellen Modus werden Linsen als ein einzelnes Objekt behandelt.
    % Im sequentiellen Modus besteht eine Linse aus 2 Oberflächen, daher
    % wird zwischen Vorder- und Rückseite einer Linse unterschieden.
    
    % --- Hier wird die Vorderseite einer Linse ausgelesen ---
    Oberflaeche = TheLDE.GetSurfaceAt(i);
    OberflaechenDatenTemp{1,2} = i;
    OberflaechenDatenTemp{2,2} = char(Oberflaeche.TypeName);
    OberflaechenDatenTemp{3,2} = char(Oberflaeche.Comment);
    OberflaechenDatenTemp{4,2} = Oberflaeche.Radius;
    OberflaechenDatenTemp{5,2} = Oberflaeche.Thickness; 
    OberflaechenDatenTemp{6,2} = char(Oberflaeche.Material);
    OberflaechenDatenTemp{8,2} = Oberflaeche.SemiDiameter;
    OberflaechenDatenTemp{10,2} = Oberflaeche.MechanicalSemiDiameter;
    OberflaechenDatenTemp{11,2} = Oberflaeche.Conic;
    OberflaechenDatenTemp{12,2} = Oberflaeche.GetSurfaceCell(ZOSAPI.Editors.LDE.SurfaceColumn.Par1).DoubleValue;
    OberflaechenDatenTemp{13,2} = Oberflaeche.GetSurfaceCell(ZOSAPI.Editors.LDE.SurfaceColumn.Par2).DoubleValue;
    OberflaechenDatenTemp{14,2} = Oberflaeche.GetSurfaceCell(ZOSAPI.Editors.LDE.SurfaceColumn.Par3).DoubleValue;
    OberflaechenDatenTemp{15,2} = Oberflaeche.GetSurfaceCell(ZOSAPI.Editors.LDE.SurfaceColumn.Par4).DoubleValue;
    OberflaechenDatenTemp{16,2} = Oberflaeche.GetSurfaceCell(ZOSAPI.Editors.LDE.SurfaceColumn.Par5).DoubleValue;
    OberflaechenDatenTemp{17,2} = Oberflaeche.GetSurfaceCell(ZOSAPI.Editors.LDE.SurfaceColumn.Par6).DoubleValue;
    OberflaechenDatenTemp{18,2} = Oberflaeche.GetSurfaceCell(ZOSAPI.Editors.LDE.SurfaceColumn.Par7).DoubleValue;
    OberflaechenDatenTemp{19,2} = Oberflaeche.GetSurfaceCell(ZOSAPI.Editors.LDE.SurfaceColumn.Par8).DoubleValue;
    
    % Hier wird die Blaupause OberflaechenDatenTemp in das Array "Oberflaechen" kopiert
    for k=1:19
        Oberflaechen_for_NIR{k,n} = OberflaechenDatenTemp{k,1};
        Oberflaechen_for_NIR{k,n+1} = OberflaechenDatenTemp{k,2};
    end
    n = n+3;
    Nummer = Nummer+1;
    
end
