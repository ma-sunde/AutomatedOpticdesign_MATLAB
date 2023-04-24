% Die Umwandlung vom nicht-sequentiellen Modus funktioniert ausschließlich
% mit Odd Asphere Oberflächen als Linsentypen! Das Ergebnis dieses
% programms sind alle notwendigen Parameter, um dasselbe System wieder im
% sequentiellen Modus aufzubauen. Die Ergebnisse stehen in "Oberflaechen"

TheApplication = MATLABZOSConnection;

% danach TheSystem laden:
TheSystem = TheApplication.PrimarySystem;

% NSC Editor aufrufen:
TheNCE = TheSystem.NCE;

% Positionen der Linsen im NSC-Editor finden

    FindNSC;
    
    
% ----------- Position des "Null Objects" finden -------------------------

    FindNullObject;

NullObject = TheNCE.GetObjectAt(PosNullObject);
ThicknessObject = NullObject.ZPosition;    % entspricht im sequentiellen Modus dem Abstand von LQ zu erster Linse


% ----------- Abspeichern der Linsendaten ---------------------------------
% Achtung, geht nur mit Odd Asphere als Surface Type!

OberflaechenDatenTemp = {};     % das cell Array "OberflaechenDaten§ ist eine Art Blaupause, die
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
for i=SurfacePosfirst:SurfacePoslast
    
    % Im nicht-sequentiellen Modus werden Linsen als ein einzelnes Objekt behandelt.
    % Im sequentiellen Modus besteht eine Linse aus 2 Oberflächen, daher
    % wird zwischen Vorder- und Rückseite einer Linse unterschieden.
    
    % Hier wird die Vorderseite einer Linse ausgelesen
    Oberflaeche = TheNCE.GetObjectAt(i);
    OberflaechenDatenTemp{1,2} = Nummer;
    OberflaechenDatenTemp{2,2} = char(Oberflaeche.TypeName);
    OberflaechenDatenTemp{3,2} = 'Vorderseite';
    OberflaechenDatenTemp{4,2} = Oberflaeche.GetCellAt(15).DoubleValue;
    OberflaechenDatenTemp{5,2} = Oberflaeche.GetCellAt(12).DoubleValue; 
    OberflaechenDatenTemp{6,2} = char(Oberflaeche.Material);
    OberflaechenDatenTemp{8,2} = Oberflaeche.GetCellAt(11).DoubleValue;
    OberflaechenDatenTemp{10,2} = Oberflaeche.GetCellAt(43).DoubleValue;
    OberflaechenDatenTemp{11,2} = Oberflaeche.GetCellAt(16).DoubleValue;
    OberflaechenDatenTemp{12,2} = Oberflaeche.GetCellAt(17).DoubleValue;
    OberflaechenDatenTemp{13,2} = Oberflaeche.GetCellAt(18).DoubleValue;
    OberflaechenDatenTemp{14,2} = Oberflaeche.GetCellAt(19).DoubleValue;
    OberflaechenDatenTemp{15,2} = Oberflaeche.GetCellAt(20).DoubleValue;
    OberflaechenDatenTemp{16,2} = Oberflaeche.GetCellAt(21).DoubleValue;
    OberflaechenDatenTemp{17,2} = Oberflaeche.GetCellAt(22).DoubleValue;
    OberflaechenDatenTemp{18,2} = Oberflaeche.GetCellAt(22).DoubleValue;
    OberflaechenDatenTemp{19,2} = Oberflaeche.GetCellAt(23).DoubleValue;
    
    % Hier wird die Blaupause OberflaechenDaten in das Array "OIberflaechen" kopiert
    for k=1:19
        Oberflaechen{k,n} = OberflaechenDatenTemp{k,1};
        Oberflaechen{k,n+1} = OberflaechenDatenTemp{k,2};
    end
    n = n+3;
    Nummer = Nummer+1;
    
    % Hier wird die Rückseite einer Linse ausgelesen
    OberflaechenDatenTemp{1,2} = Nummer;
    OberflaechenDatenTemp{2,2} = char(Oberflaeche.TypeName);
    OberflaechenDatenTemp{3,2} = 'Rückseite';
    OberflaechenDatenTemp{4,2} = Oberflaeche.GetCellAt(29).DoubleValue;
    
    % Thickness der Rückseite entpsricht dem Abstand zur nächsten Linse.
    % Z Position ist im non-sequential mode relativ zur Referenzfläche, 
    % daher kurze Umrechnung notwendig:
    
    Oberflaeche = TheNCE.GetObjectAt(i+1);
    NextObjectRef = Oberflaeche.RefObject;
    Oberflaeche = TheNCE.GetObjectAt(i);
    RefObject = Oberflaeche.RefObject;

    
    
        if NextObjectRef == i
            Oberflaeche = TheNCE.GetObjectAt(i+1);
            OberflaechenDatenTemp{5,2} = Oberflaeche.ZPosition;
            Oberflaeche = TheNCE.GetObjectAt(i);
        elseif RefObject == NullObject.IndexData.RowData.ObjectNumber
            Thickness = Oberflaeche.GetCellAt(12).DoubleValue;
            ZPosition = Oberflaeche.ZPosition;
            Oberflaeche = TheNCE.GetObjectAt(i+1);
            ZPositionNext = Oberflaeche.ZPosition;
            OberflaechenDatenTemp{5,2} = ZPositionNext - ZPosition - Thickness; 
            Oberflaeche = TheNCE.GetObjectAt(i);
        end
        
    % Ende der Umrechnung
    
    OberflaechenDatenTemp{6,2} = char(Oberflaeche.Material);  
    OberflaechenDatenTemp{8,2} = Oberflaeche.GetCellAt(44).DoubleValue;
    OberflaechenDatenTemp{10,2} = Oberflaeche.GetCellAt(43).DoubleValue;
    OberflaechenDatenTemp{11,2} = Oberflaeche.GetCellAt(30).DoubleValue;
    OberflaechenDatenTemp{12,2} = Oberflaeche.GetCellAt(31).DoubleValue;
    OberflaechenDatenTemp{13,2} = Oberflaeche.GetCellAt(32).DoubleValue;
    OberflaechenDatenTemp{14,2} = Oberflaeche.GetCellAt(33).DoubleValue;
    OberflaechenDatenTemp{15,2} = Oberflaeche.GetCellAt(34).DoubleValue;
    OberflaechenDatenTemp{16,2} = Oberflaeche.GetCellAt(35).DoubleValue;
    OberflaechenDatenTemp{17,2} = Oberflaeche.GetCellAt(36).DoubleValue;
    OberflaechenDatenTemp{18,2} = Oberflaeche.GetCellAt(37).DoubleValue;
    OberflaechenDatenTemp{19,2} = Oberflaeche.GetCellAt(38).DoubleValue;
    
    for k=1:19
        Oberflaechen{k,n} = OberflaechenDatenTemp{k,1};
        Oberflaechen{k,n+1} = OberflaechenDatenTemp{k,2};
    end
    n = n+3;
    Nummer = Nummer+1;
end


% -------------- Neue Datei erstellen ------------------------------------

sampleDir = TheApplication.SamplesDir;
TheSystem = TheApplication.CreateNewSystem(ZOSAPI.SystemType.Sequential);      % Create New SC File
Converted_from_NS_to_SC = System.String.Concat(TheApplication.SamplesDir, '\Converted_from_NS_to_SC.zmx');   % Define file path and name
TheSystem.SaveAs(Converted_from_NS_to_SC);  % Save New SC File
file = '\Converted_from_NS_to_SC.zmx';
TheSystem.LoadFile(System.String.Concat(TheApplication.SamplesDir, file), false);


