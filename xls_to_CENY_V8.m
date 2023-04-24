%                                                                              %
%                                                                              %
%       ! Zuerst OpticStudio öffnen und im Reiter Programming auf den          %
%                Button Interactive Extension klicken !                        %                     
%                   Danach dieses Skript starten.                              %
%                                                                              %
%                                                                              %

%System_Load;        % nach Connection ausführen


% Input Datei enthält in der 1. Zeile & Spalte die Grad Zahlen und ist in
% einem FOV von -23° bis +23° horizontal und -10° bis +5° vertikal


% Eine Anzahl von 44 Fields sollte bei der Input Datei
% Lichtverteilung_idealisiert nicht überschritten werden.
% Es können nur so viele Fields verteilt werden, wie die Diagonale lang
% ist. Bei größeren Input Tabellen überschreitet die Länge der Diagonalen
% die Auflösung der LED Matrix.

% Diese Funktion wird mehrmals aufgerufen und die enthalten Variablen mehrmals
% befüllt und wieder überschrieben. Zur Fehlervermeidung werden alle mehrmals
% genutzten Variablen vor dem Start gelöscht 

clear CENY_Values;
clear Bereiche;
clear Bereiche_V2;
clear Schwerpunkte;
clear Schwerpunkte_V2;

% Anzahl der Fields
NOF = TheSystem.SystemData.Fields.NumberOfFields;
NumberOfFields = double(NOF);   %Funktion gibt NumberOfFields als Integer zurück. Muss vorher umgewandelt werden!
 
%NumberOfFields = 10;    %zum testen

%--------------------- Lichtvereitlung einlesen --------------------------
LichtverteilungVIS = xlsread(fullfile(FolderVIS_Lichtverteilung,FileVIS_Lichtverteilung));
%size = size(Lichtverteilung);
ResolutionVertikal = abs(LichtverteilungVIS(2,1)-LichtverteilungVIS(3,1));
ResolutionHorizontal = abs(LichtverteilungVIS(1,2)-LichtverteilungVIS(1,3));

% hellsten Punkt finden
[Maximum, MaxPos] = max(LichtverteilungVIS(:));

% Position der hellsten Punkte finden
[rowsOfMaxes, colsOfMaxes] = find(LichtverteilungVIS == Maximum);

% 0° Vertikal 0° Horizontal lokalisieren // vorausgesetzt erste Zeile und
% Spalte beinhalten die Achsenbeschriftung
colofZero = find(LichtverteilungVIS(1,:) == 0);
rowofZero = find(LichtverteilungVIS(:,1) == 0);


% weit entfernteste Ecke finden
Ecke{1,1} = 'Position'; 
Ecke{1,2} = 'Pythagoras';
Ecke{1,3} = 'DistanceHorizontal';
Ecke{1,4} = 'DistanceVertikal';

 Ecke{2,1} = 'links oben';
 Ecke{2,2} = sqrt((colofZero -1)^2 + (rowofZero - 1)^2);
 Ecke{2,3} = colofZero -1;
 Ecke{2,4} = rowofZero - 1;
 
 Ecke{3,1} = 'rechts oben';
 Ecke{3,2} = sqrt((size(LichtverteilungVIS,2) - (colofZero - 1))^2 + (rowofZero - 1)^2);
 Ecke{3,3} = size(LichtverteilungVIS,2) - (colofZero - 1);
 Ecke{3,4} = rowofZero - 1;
 
 Ecke{4,1} = 'links unten';
 Ecke{4,2} = sqrt((size(LichtverteilungVIS,1) - (rowofZero - 1))^2 + (colofZero - 1)^2);
 Ecke{4,3} = colofZero - 1;
 Ecke{4,4} = size(LichtverteilungVIS,1) - (rowofZero - 1);
 
 Ecke{5,1} = 'rechts unten';
 Ecke{5,2} = sqrt((size(LichtverteilungVIS,2) - (colofZero - 1))^2 + (size(LichtverteilungVIS,1) - (rowofZero - 1))^2);
 Ecke{5,3} = size(LichtverteilungVIS,2) - (colofZero - 1);
 Ecke{5,4} = size(LichtverteilungVIS,1) - (rowofZero - 1);

 % Initialisieren der Variablen zur Speicherung des Maximums und der Position
max_val = -Inf; % Startwert ist negativ unendlich, damit jedes Element im Array größer ist
max_pos = [];

% Schleife durch die Zeilen 2 bis 5 und Suche nach dem Maximum in der zweiten Spalte
for i = 2:5
    if Ecke{i,2} > max_val
        max_val = Ecke{i,2};
        max_pos = i;
    elseif Ecke{i,2} == max_val % Speichern der Position(en) des Maximums
        max_pos(end+1) = i;
    end
end

if size(max_pos,2) > 1
    i = max_pos(end);
elseif size(max_pos,2) == 1
    i = max_pos;
end


% Anzahl der Zellen bis von 0°V 0°H bis zur weit entferntesten Ecke 
DistanceHorizontal = Ecke{i,3};
DistanceVertikal = Ecke{i,4};

% Entfernung vom Mittelpunkt zur weit entferntesten Ecke
Diagonale = Ecke{i,2};

% Schrittweite horizontal und vertikal für Diagonale
 
StepHorizontal = DistanceHorizontal/DistanceVertikal;
StepInteger = round(StepHorizontal);                    %Vom Mittelpunkt der Tabelle einen Schritt nach unten und StepInteger Schritte nach rechts um eine Diagonale zu erzeugen

%--------------- halbe Diagonale auslesen ---------------------------------


switch i
    case 2      %Abtastung nach links oben
        n=0;
        k=0;
        while k<StepInteger*DistanceVertikal
    Helligkeitsverlauf(k+1,1) = StepInteger*DistanceVertikal-k;
    Helligkeitsverlauf(k+1,2) = LichtverteilungVIS(rowofZero-n,colofZero-k);
    Helligkeitsverlauf(k+1,6) = LichtverteilungVIS(rowofZero-n,1);
    Helligkeitsverlauf(k+1,7) = LichtverteilungVIS(1,colofZero-k);
    if mod(k,StepInteger) == 0 && k~=0
       n = n+1;
       %Dieser Step sorgt dafür dass alle Werte beim stufenförmigen Ablesen mitgenommen werden. Für ein realistisches Abbild wird die Kante einer Stufe besser abgeschnitten 
       %Helligkeitsverlauf(i+2,1) = StepInteger*DistanceVertikal-i-1;                       
       %Helligkeitsverlauf(i+2,2) = Lichtverteilung(rowofZero-n,colofZero-i);
       %Helligkeitsverlauf(i+2,6) = Lichtverteilung(rowofZero-n,1);
       %Helligkeitsverlauf(i+2,7) = Lichtverteilung(1,colofZero-i);
       %i = i+1;
       %Quasi statt so:
       % 1234
       %    56789
       % werden nur
       % 1234
       %     6789
       %abgetastet.
    end
    k=k+1;
end
    case 3      %Abtastung nach rechts oben
        n=0;
        k=0;
        while k<StepInteger*DistanceVertikal
    Helligkeitsverlauf(k+1,1) = StepInteger*DistanceVertikal-k;
    Helligkeitsverlauf(k+1,2) = LichtverteilungVIS(rowofZero-n,colofZero+k);
    Helligkeitsverlauf(k+1,6) = LichtverteilungVIS(rowofZero-n,1);
    Helligkeitsverlauf(k+1,7) = LichtverteilungVIS(1,colofZero+k);
    if mod(k,StepInteger) == 0 && k~=0
       n = n+1;

    end
    k=k+1;
        end
    case 4      %Abtastung nach links unten
        n=0;
        k=0;
        while k<StepInteger*DistanceVertikal
    Helligkeitsverlauf(k+1,1) = StepInteger*DistanceVertikal-k;
    Helligkeitsverlauf(k+1,2) = LichtverteilungVIS(rowofZero+n,colofZero-k);
    Helligkeitsverlauf(k+1,6) = LichtverteilungVIS(rowofZero+n,1);
    Helligkeitsverlauf(k+1,7) = LichtverteilungVIS(1,colofZero-k);
    if mod(k,StepInteger) == 0 && k~=0
       n = n+1;

    end
    k=k+1;
        end
    case 5      %Abtastung nach rechts unten
        n=0;
        k=0;
        while k<StepInteger*DistanceVertikal
    Helligkeitsverlauf(k+1,1) = StepInteger*DistanceVertikal-k;
    Helligkeitsverlauf(k+1,2) = LichtverteilungVIS(rowofZero+n,colofZero+k);
    Helligkeitsverlauf(k+1,6) = LichtverteilungVIS(rowofZero+n,1);
    Helligkeitsverlauf(k+1,7) = LichtverteilungVIS(1,colofZero+k);
    if mod(k,StepInteger) == 0 && k~=0
       n = n+1;

    end
    k=k+1;
        end
end

%---------- Helligkeitsgraphen erstellen ---------------------------------

% Bei Bedarf Kommentarbefehl wieder löschen
 
%  plot(Helligkeitsverlauf(:,1),Helligkeitsverlauf(:,2),'r:+');
%  xlabel('Länge der Diagonalen');
%  ylabel('Helligkeitswerte in lux');
%  title('Helligkeitsverlauf der Diagonalen der Lichtverteilung','FontSize',12);

%---------- Bereiche für Flächenschwerpunkte einteilen -------------------


% Prozentualer Anteil der Helligkeitswerte mit hellstem Punkt = 100%

[MaxDiagonale, MaxPosDiagonale] = max(Helligkeitsverlauf(:,2));
Helligkeitsverlauf(MaxPosDiagonale,3) = 100;
for i=1:StepInteger*DistanceVertikal
        Helligkeitsverlauf(i,3) = 100/MaxDiagonale*Helligkeitsverlauf(i,2);
end

% Prozentualer Anteil der Helligkeitswerte zur gesamten Helligkeit

Summe = sum(Helligkeitsverlauf(:,2));
for i=1:StepInteger*DistanceVertikal
        Helligkeitsverlauf(i,4) = 100/Summe*Helligkeitsverlauf(i,2);
end

% Spalte 'Fields Ja/Nein' mit Nullen füllen für die Abfrage danach

for i=1:DistanceVertikal*StepInteger
        Helligkeitsverlauf(i,5) = 0;  
end

% Erster & letzter Wert bekommen ein Field, um das FOV abzudecken

Helligkeitsverlauf(1,5) = 1;
Helligkeitsverlauf(StepInteger*DistanceVertikal,5) = 1;

% Die restlichen Fields werden anhand eines Schwellwertes verteilt


i=1;        
n=0;
temp = 0;                   %Hier muss noch eine Ausnahme gemacht werden für Wert 1! Rest kann negativ sein und muss beachtet werden! für Gewichtung da Position 1 immer ein Field bekommt
p=1;
Rest = 0;
while i<=StepInteger*DistanceVertikal  %StepInteger*DistanceVertikal-1
    
    if i == StepInteger*DistanceVertikal                    %Der erste Wert wird als einzelner Bereich bewertet, da diese 
        Bereiche(p,1) = Helligkeitsverlauf(i,4);            %Position ein Field erhält, um das gesamte FOV abzudecken.
        Rest          = Bereiche(p,1)-(100/NumberOfFields); %Der Rest im ersten Field wird ohne Einschränkungen ausgerechnet
        Bereiche(p,2) = Rest;                               %In der Regel erreicht keine einzelne Zelle den Schwellwert 
        Bereiche(p,4) = i;                                  %Es ist davon auszugehen, dass die erste Zelle einen negativen Rest hat
        Bereiche(p,5) = i;
    end
    
    temp = temp + Helligkeitsverlauf(i,4);
    n = n+1;   %zählt mit, wie oft temp aufaddiert wird, bis Schwellwert erreicht ist 
    
    if n == 1
        temp = temp + Rest;     %Rest darf nur einmal aufaddiert werden beim ersten Durchlauf!
    end     
       
    if temp >= 100/NumberOfFields               
        Rest = temp-(100/NumberOfFields);
    end
    
    if temp >= 100/NumberOfFields
        Bereiche(p,1) = temp;
        Bereiche(p,2) = Rest;
        Bereiche(p,4) = i-n+1;
        Bereiche(p,5) = i;
        p=p+1;
        temp = 0;
        n=0;
    end
    
    if i == StepInteger*DistanceVertikal                    %Die letzten Werte, die zum Schluss aufaddiert werden, 
        Bereiche(p,1) = Helligkeitsverlauf(i,4);            %aber nicht den Schwellwert erreichen
        Rest          = Bereiche(p,1)-(100/NumberOfFields); %Der Rest im letzten Field wird ohne Einschränkungen ausgerechnet
        Bereiche(p,2) = Rest;                               %In der Regel erreicht keine einzelne Zelle den Schwellwert 
        Bereiche(p,4) = i;                                  %Es ist davon auszugehen, dass die letzte Zelle einen negativen Rest hat
        Bereiche(p,5) = i;
    end
    
        i = i+1;
end

%Aufgrund der Verteilung des ersten und letzten Fields am innersten und
%äußersten Punkt des FOV, kann es passieren, dass aufgrund von Rundungsungenauigkeiten
%mehr Fields verteilt werden als vorgesehen. Das muss überprüft werden
%Aufgrund der Chonolgie einer Schleife treten diese Rundungsfehler immer
%am Ende der Tabelle auf. Da der letzte Bereich aus einer einzelnen Zelle besteht,
%wird der vorletzte Bereich gelöscht

if size(Bereiche,1) > NumberOfFields
    Bereiche(size(Bereiche,1),1) = Bereiche(size(Bereiche,1),1)+Bereiche(size(Bereiche,1)-1,1); %Flächen werden addiert
    Bereiche(size(Bereiche,1),2) = Bereiche(size(Bereiche,1),1)-(100/NumberOfFields);           %Rest wird neu berechnet, da in der Regel der Schwellwert erreicht wird
    Bereiche(size(Bereiche,1),4) = Bereiche(size(Bereiche,1)-1,4);                              %Bereich wird angepasst
    Bereiche(size(Bereiche,1)-1,:)=[];                                                          %Vorletzte zeile wird gelöscht
end


%------------- Bereiche auslesen und Schwerpunkte berechnen -------------

[Zeilen,Spalten] = size(Bereiche);

for i=1:Zeilen     %1:Zeilen
    
    if i == 1                                                       % Erste Position bekommt immer ein Field 
        Schwerpunkte(i,1) = 1;                                      %'Fieldnummer'
        Schwerpunkte(i,2) = Helligkeitsverlauf(Bereiche(i,4),1);    %'x-Koordinate des Flächenschwerpunktes & Zelle der Diagonalen'                      
        Schwerpunkte(i,3) = Helligkeitsverlauf(Bereiche(i,4),2);    %'Flächeninhalt des Bereiches'
        Schwerpunkte(i,4) = Helligkeitsverlauf(Bereiche(i,4),6);    %'Vertikal in °'
        Schwerpunkte(i,5) = Helligkeitsverlauf(Bereiche(i,4),7);    %'Horizontal in °'    
    
    
    elseif i == Zeilen                                              % letztes Position bekommt immer ein Field
        Schwerpunkte(i,1) = Zeilen;                                 %'Fieldnummer'
        Schwerpunkte(i,2) = Helligkeitsverlauf(Bereiche(i,5),1);    %'x-Koordinate des Flächenschwerpunktes & Zelle der Diagonalen'                      
        Schwerpunkte(i,3) = Helligkeitsverlauf(Bereiche(i,5),2);    %'Flächeninhalt des Bereiches'
        Schwerpunkte(i,4) = Helligkeitsverlauf(Bereiche(i,5),6);    %'Vertikal in °'
        Schwerpunkte(i,5) = Helligkeitsverlauf(Bereiche(i,5),7);    %'Horizontal in °'  
        
        %Ausnahmeregelung für letzten und vorletzten Bereich! Schwerpunkt des vorletzten Bereiches berechenet sich
        %aus dem Rest des letzten Bereiches, dem Schwerpunkt des
        %aktuellen Bereiches und dem Rest aus dem Bereich vor dem aktuellen
        %Es ist davon auszugehen, dass eine einzelen Zelle den Schwellwert
        %nicht erreicht. Die letzte Zelle bekommt immer ein Field, um das
        %FOV in seiner Gesamtheit abzudecken. Damit die damit verbundene
        %Helligkeit trotzdem verhältnismäßig aufgeteilt wird, müssen die
        %Schwerpunkte mit dem vorherigen Field verrechnet werden. Das wird ebenfalls für das
        %erste Field gemacht. Aufgrund der chronolgischen Abfolge der
        %Schleife mit ihren Berechnungen, muss allerdings für das letzte Field
        %eine Ausnahme gemacht werden.
        
        i=i-1;  %Da hier der vorletzte Bereich bearbeitet wird, wird der Zähler i temporär zurückgesetzt, da wir uns zur Zeit in
                %einer if-Schleife befinden, wo i=Zeilen festgelegt ist.
                %Ab hier ist i die Position des vorletzten Bereiches.
        
        A0     = Schwerpunkte(i-1,3);       %Bereich vor dem Vorletzten
        A1     = Schwerpunkte(i,3);         %Vorletzter Bereich
        A2     = Schwerpunkte(i+1,3);       %letzter Bereich
        xs0    = Schwerpunkte(i-1,2);
        xs1    = Schwerpunkte(i,2);
        xs2    = Schwerpunkte(i+1,2);
        xs     = (A0*xs0+A1*xs1+A2*xs2)/(A0+A1+A2);     %Schwerpunkt des vorletzten Bereiches inklusive der Anteile vom letzten Bereich und dem Rest vom Bereich davor
        
        Schwerpunkte(i,2) = xs;                         %'x-Koordinate des Flächenschwerpunktes & Zelle der Diagonalen'                      
          
        i=i+1; %Zähler i wird wieder zurückgesetzt
        
        % lineare Interpolation der Horizontal-Koordinaten für Schwerpunkte
        y1=Helligkeitsverlauf(Bereiche(i,4),7);
        y2=Helligkeitsverlauf(Bereiche(i-1,5),7);
        x1=Helligkeitsverlauf(Bereiche(i,4),1);
        x2=Helligkeitsverlauf(Bereiche(i-1,5),1);
        x=xs;
        Koordinate_Horizontal = y1 + (y2-y1)/(x2-x1)*(x-x1);

        % lineare Interpolation der Vertikal-Koordinaten für Schwerpunkte
        y1=Helligkeitsverlauf(Bereiche(i,4),6);
        y2=Helligkeitsverlauf(Bereiche(i-1,5),6);
        x1=Helligkeitsverlauf(Bereiche(i,4),1);
        x2=Helligkeitsverlauf(Bereiche(i-1,5),1);
        x=xs;
        Koordinate_Vertikal = y1 + (y2-y1)/(x2-x1)*(x-x1);
        
        i=i-1; %Zähler i wird temporär wieder auf den vorletzten bereich zurückgestuft
        
        Schwerpunkte(i,4) = Koordinate_Vertikal;                    %'Vertikal in °'
        Schwerpunkte(i,5) = Koordinate_Horizontal;                  %'Horizontal in °'
        
        i=i+1;  %hier wird der Zähler i wieder auf seine normale Position gesetzt.
        
    else       
       
    Ai0=Bereiche(i-1,2)/100*Summe;                  %Ai0 und xsi0 sind die Rest-Fläche und der Schwerpunkt
    xsi0=Helligkeitsverlauf(Bereiche(i-1,5),1);     %aus dem Bereich davor.
    Teilflaeche = 0;
    
        for p=Bereiche(i,4):Bereiche(i,5)
            % Flächenberechnung eines Bereiches (ohne den Rest aus der Fläche davor zu berücksichtigen)
            Teilflaeche = Teilflaeche + Helligkeitsverlauf(p,2);
            
            % Gewichtete Berechnung der x-Koordinate eines Flächenschwerpunktes bestehend
            % aus mehreren Flächen und Schwerpunkten innerhalb eines Bereiches.                                                                          
            Ai(p-Bereiche(i,4)+1,1)=Helligkeitsverlauf(p,2);
            xsi(p-Bereiche(i,4)+1,1)=Helligkeitsverlauf(p,1);           
        end
        
        % Gewichtete Berechnung der x-Koordinate eines Flächenschwerpunktes bestehend
        % aus mehreren Flächen und Schwerpunkten inklusive der Fläche Ai0 mit dem Schwerpunkt xsi0
        % aus dem Bereich davor (Rest).                                    
        xs = (sum(Ai.*xsi)+Ai0*xsi0)/(sum(Ai)+Ai0);
        
        
    %-------------- Koordinaten zuweisen ----------------------------
    
        % lineare Interpolation der Horizontal-Koordinaten für Schwerpunkte
        y1=Helligkeitsverlauf(Bereiche(i,4),7);
        y2=Helligkeitsverlauf(Bereiche(i-1,5),7);
        x1=Helligkeitsverlauf(Bereiche(i,4),1);
        x2=Helligkeitsverlauf(Bereiche(i-1,5),1);
        x=xs;
        Koordinate_Horizontal = y1 + (y2-y1)/(x2-x1)*(x-x1);

        % lineare Interpolation der Vertikal-Koordinaten für Schwerpunkte
        y1=Helligkeitsverlauf(Bereiche(i,4),6);
        y2=Helligkeitsverlauf(Bereiche(i-1,5),6);
        x1=Helligkeitsverlauf(Bereiche(i,4),1);
        x2=Helligkeitsverlauf(Bereiche(i-1,5),1);
        x=xs;
        Koordinate_Vertikal = y1 + (y2-y1)/(x2-x1)*(x-x1);

        Schwerpunkte(i,1) = i;
        Schwerpunkte(i,2) = xs;                                     %'x-Koordinate des Flächenschwerpunktes & Zelle der Diagonalen'                      
        Schwerpunkte(i,3) = Teilflaeche;                            %'Flächeninhalt eines Bereiches' 
        Schwerpunkte(i,4) = Koordinate_Vertikal;                    %'Vertikal in °'
        Schwerpunkte(i,5) = Koordinate_Horizontal;                  %'Horizontal in °'
        
        
    end

end




%funktioniert noch nicht richtig
p=1;
for i=1:StepInteger*DistanceVertikal
    if round(Schwerpunkte(p,2)) == Helligkeitsverlauf(i,1) & p<=size(Schwerpunkte)
        Helligkeitsverlauf(i,5) = 1;
        p=p+1;
    else
        Helligkeitsverlauf(i,5) = 0;
    end
end


%------------ Ausgabe erstellen ------------------------------------------
% Für die Übersichtlichkeit werden noch eine Tabellen dupliziert und mit
% Überschriften versehen. Aufgrund einiger Befehle konnten diese Arten
% von Tabelle vorher nicht verwendet werden.

Helligkeitsverlauf_V2 = cell(DistanceVertikal*StepInteger,5);

Helligkeitsverlauf_V2{1,1} = 'Nummer';
Helligkeitsverlauf_V2{1,2} = 'Lux-Wert';
Helligkeitsverlauf_V2{1,3} = 'Lux-Wert in %';
Helligkeitsverlauf_V2{1,4} = 'Lux-Wert in % relativ zur gesamten Helligkeit';
Helligkeitsverlauf_V2{1,5} = 'Field ja/nein';
Helligkeitsverlauf_V2{1,6} = 'Vertikal in °';
Helligkeitsverlauf_V2{1,7} = 'Horizontal in °';

[Zeilen,Spalten] = size(Helligkeitsverlauf);

for n=1:Spalten
    for i=1:Zeilen
        Helligkeitsverlauf_V2{i+1,n} = Helligkeitsverlauf(i,n);  
    end
end

Bereiche_V2 = cell(size(Bereiche));

Bereiche_V2{1,1} = 'Lux-Wert in % relativ zur gesamten Helligkeit';
Bereiche_V2{1,2} = 'Rest aus Bereich davor';
Bereiche_V2{1,4} = 'von Zelle: ';
Bereiche_V2{1,5} = 'bis Zelle: ';

[Zeilen,Spalten] = size(Bereiche);

for n=1:Spalten
    for i=1:Zeilen
        Bereiche_V2{i+1,n} = Bereiche(i,n);  
    end
end

Schwerpunkte_V2 = cell(size(Schwerpunkte));

Schwerpunkte_V2{1,1} = 'Fieldnummer';
Schwerpunkte_V2{1,2} = 'x-Koordinate des Flächenschwerpunktes & Zelle der Diagonalen';
Schwerpunkte_V2{1,3} = 'Flächeninhalt des Bereiches';
Schwerpunkte_V2{1,4} = 'Vertikal in °';
Schwerpunkte_V2{1,5} = 'Horizontal in °';

[Zeilen,Spalten] = size(Schwerpunkte);

for n=1:Spalten
    for i=1:Zeilen
        Schwerpunkte_V2{i+1,n} = Schwerpunkte(i,n);  
    end
end

CENY_Values = cell(NumberOfFields,4);

CENY_Values{1,1} = 'Fieldnummer';
CENY_Values{1,2} = 'Vertikal in °';
CENY_Values{1,3} = 'Horizontal in °';
CENY_Values{1,4} = 'Y-Position in mm';

[Zeilen,Spalten] = size(CENY_Values);

    for i=1:Zeilen
        CENY_Values{i+1,1} = i;
        CENY_Values{i+1,2} = Schwerpunkte(i,4); 
        CENY_Values{i+1,3} = Schwerpunkte(i,5);
        CENY_Values{i+1,4} = abs(tand(Schwerpunkte(i,5))*25*10^03);     %Es wird der Betrag genommen, da es sein kann, dass die Koordinaten negativ sind
    end                                                                 
                    




