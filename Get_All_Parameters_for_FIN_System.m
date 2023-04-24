%                                                                              %
%                                                                              %
%       ! Zuerst OpticStudio öffnen und im Reiter Programming auf den          %
%                Button Interactive Extension klicken !                        %                     
%                   Danach dieses Skript starten.                              %
%                                                                              %
%  


System_Load;

% NSC Editor aufrufen:
TheNCE = TheSystem.NCE;

% Positionen der Linsen im NSC-Editor finden

    FindNSC;

% Auslesen der Parameter aus dem NSC-Editor


%--- Array zum Abspeichern der Linsendaten vorbereiten

NSCDatenTemp = {}; % das cell Array "OberflaechenDatenTemp ist eine Art Blaupause, die
                            % immer wieder neu befüllt wird. Alle
                            % Linsendaten werden später in dem Array
                            % "Oberflaechen" abgespeichert
 

NSCDatenTemp{1,1} = 'ObjectType:';
NSCDatenTemp{2,1} = 'Comment';
NSCDatenTemp{3,1} = 'Ref Object';
NSCDatenTemp{4,1} = 'Inside Of';
NSCDatenTemp{5,1} = 'X Position';
NSCDatenTemp{6,1} = 'Y Position';
NSCDatenTemp{7,1} = 'Z Position';
NSCDatenTemp{8,1} = 'Tilt Abaout X';
NSCDatenTemp{9,1} = 'Tilt Abaout Y';
NSCDatenTemp{10,1} = 'Tilt Abaout Z';
NSCDatenTemp{11,1} = 'Material';
NSCDatenTemp{12,1} = 'Clear 1';
NSCDatenTemp{13,1} = 'Thickness';
NSCDatenTemp{14,1} = 'PAR 3 (unused)';
NSCDatenTemp{15,1} = 'PAR 4 (unused)';
NSCDatenTemp{16,1} = 'Radius 1';
NSCDatenTemp{17,1} = 'Conic 1';
NSCDatenTemp{18,1} = 'Coeff 1 r^1';
NSCDatenTemp{19,1} = 'Coeff 1 r^2';
NSCDatenTemp{20,1} = 'Coeff 1 r^3';
NSCDatenTemp{21,1} = 'Coeff 1 r^4';
NSCDatenTemp{22,1} = 'Coeff 1 r^5';
NSCDatenTemp{23,1} = 'Coeff 1 r^6';
NSCDatenTemp{24,1} = 'Coeff 1 r^7';
NSCDatenTemp{25,1} = 'Coeff 1 r^8';
NSCDatenTemp{26,1} = 'Coeff 1 r^9';
NSCDatenTemp{27,1} = 'Coeff 1 r^10';
NSCDatenTemp{28,1} = 'Coeff 1 r^11';
NSCDatenTemp{29,1} = 'Coeff 1 r^12';
NSCDatenTemp{30,1} = 'Radius 2';
NSCDatenTemp{31,1} = 'Conic 2';
NSCDatenTemp{32,1} = 'Coeff 2 r^1';
NSCDatenTemp{33,1} = 'Coeff 2 r^2';
NSCDatenTemp{34,1} = 'Coeff 2 r^3';
NSCDatenTemp{35,1} = 'Coeff 2 r^4';
NSCDatenTemp{36,1} = 'Coeff 2 r^5';
NSCDatenTemp{37,1} = 'Coeff 2 r^6';
NSCDatenTemp{38,1} = 'Coeff 2 r^7';
NSCDatenTemp{39,1} = 'Coeff 2 r^8';
NSCDatenTemp{40,1} = 'Coeff 2 r^9';
NSCDatenTemp{41,1} = 'Coeff 2 r^10';
NSCDatenTemp{42,1} = 'Coeff 2 r^11';
NSCDatenTemp{43,1} = 'Coeff 2 r^12';
NSCDatenTemp{44,1} = 'Edge 1';
NSCDatenTemp{45,1} = 'Clear 2';
NSCDatenTemp{46,1} = 'Edge 2';

%--------------------- Linsendaten abspeichern ----------------------------

n=1;
Nummer = 1;
for i=SurfacePosfirst:SurfacePoslast
    
        Oberflaeche = TheNCE.GetObjectAt(i);
        
    for k=1:size(NSCDatenTemp,1)
        if k==1
            NSCDatenTemp{k,2} = char(Oberflaeche.Type);
        elseif k==2
            NSCDatenTemp{k,2} = char(Oberflaeche.Comment); 
        elseif k==3
            NSCDatenTemp{k,2} = Oberflaeche.RefObject;
        elseif k==4
            NSCDatenTemp{k,2} = Oberflaeche.InsideOf;
        elseif k==5
            NSCDatenTemp{k,2} = Oberflaeche.XPosition;
        elseif k==6    
            NSCDatenTemp{k,2} = Oberflaeche.YPosition;
        elseif k==7
            NSCDatenTemp{k,2} = Oberflaeche.ZPosition;
        elseif k==8
            NSCDatenTemp{k,2} = Oberflaeche.TiltAboutX;
        elseif k==9
            NSCDatenTemp{k,2} = Oberflaeche.TiltAboutY;
        elseif k==10
            NSCDatenTemp{k,2} = Oberflaeche.TiltAboutZ;
        elseif k==11
            NSCDatenTemp{k,2} = char(Oberflaeche.Material);
        elseif  k==14 || k==15 
            NSCDatenTemp{k,2} = char(Oberflaeche.GetCellAt(k-1).Value);  %aus welchem Grund auch immer verschiebt sich hier was
        else    
            NSCDatenTemp{k,2} = Oberflaeche.GetCellAt(k-1).DoubleValue;  
        end
    end
    
    for k=1:size(NSCDatenTemp,1)
        NSCDaten_NIR{k,n} = NSCDatenTemp{k,1};
        NSCDaten_NIR{k,n+1} = NSCDatenTemp{k,2};
    end
    n = n+3;
    Nummer = Nummer+1;
    
end

%------------- Lichtquelle abspeichern ------------------------------------

NSCDatenSource = {};

NSCDatenSource{1,1} = 'ObjectType:';
NSCDatenSource{2,1} = 'Comment';
NSCDatenSource{3,1} = 'Ref Object';
NSCDatenSource{4,1} = 'Inside Of';
NSCDatenSource{5,1} = 'X Position';
NSCDatenSource{6,1} = 'Y Position';
NSCDatenSource{7,1} = 'Z Position';
NSCDatenSource{8,1} = 'Tilt Abaout X';
NSCDatenSource{9,1} = 'Tilt Abaout Y';
NSCDatenSource{10,1} = 'Tilt Abaout Z';
NSCDatenSource{11,1} = 'Material';
NSCDatenSource{12,1} = 'Layout Rays';
NSCDatenSource{13,1} = 'Analysis Rays';
NSCDatenSource{14,1} = 'Power';
NSCDatenSource{15,1} = 'Wavenumber';
NSCDatenSource{16,1} = 'Color';
NSCDatenSource{17,1} = 'Cone Angle';

Oberflaeche = TheNCE.GetObjectAt(2);
        
    for k=1:size(NSCDatenSource,1)
        if k==1
            NSCDatenSource{k,2} = char(Oberflaeche.Type);
        elseif k==2
            NSCDatenSource{k,2} = char(Oberflaeche.Comment); 
        elseif k==3
            NSCDatenSource{k,2} = Oberflaeche.RefObject;
        elseif k==4
            NSCDatenSource{k,2} = Oberflaeche.InsideOf;
        elseif k==5
            NSCDatenSource{k,2} = Oberflaeche.XPosition;
        elseif k==6    
            NSCDatenSource{k,2} = Oberflaeche.YPosition;
        elseif k==7
            NSCDatenSource{k,2} = Oberflaeche.ZPosition;
        elseif k==8
            NSCDatenSource{k,2} = Oberflaeche.TiltAboutX;
        elseif k==9
            NSCDatenSource{k,2} = Oberflaeche.TiltAboutY;
        elseif k==10
            NSCDatenSource{k,2} = Oberflaeche.TiltAboutZ;
        elseif k==11
            NSCDatenSource{k,2} = char(Oberflaeche.Material);
        else
            NSCDatenSource{k,2} = char(Oberflaeche.GetCellAt(k-1).Value);
        end
    end
    
%---------------------- Null Object abspeichern ---------------------------


% ----------- Position des "Null Objects" finden -------------------------

    FindNullObject;

NullObject = TheNCE.GetObjectAt(PosNullObject);
NullObject_ZPosition = NullObject.ZPosition;    