%                                                                              %
%                                                                              %
%       ! Zuerst OpticStudio öffnen und im Reiter Programming auf den          %
%                Button Interactive Extension klicken !                        %                     
%                   Danach dieses Skript starten.                              %
%                                                                              %
%                                                                              %

System_Load;

% Neues System erstellen ohne zu speichern. Alle Linsendaten werden von
% Matlab gespeichert
TheSystem.New(false);

% Den nicht-sequentiellen Editor ausblenden
%TheNCE.HideEditor;
 
if isempty(FileNIR) == 1 && isempty(FolderNIR) == 1
    %der Part wird nur ausgeführt wenn kein eigenes Startsystem vorgegeben
    %wurde. Das programm springt im Falle eines eigenen Startsystem direkt
    %zum Erstellen der Merit Function.
    
    
    % --------------- Lichtquelle definieren ---------------------------------

    % Apertur definieren
    SystExplorer = TheSystem.SystemData;
    SystExplorer.Aperture.ApertureType = ZOSAPI.SystemData.ZemaxApertureType.ObjectConeAngle;
    TheSystem.SystemData.Aperture.ApertureValue = 14;
    TheSystem.SystemData.Wavelengths.GetWavelength(1).Wavelength = 0.808;

    %---------------------- Startsystem erstellen ----------------------------

    % Die hier gewählten Radien und Abstände entsprechen einem händisch
    % erstelltem Startsystem, mit großem Abstand zwischen Linse 2 und 3 für die
    % später eingefügten Spiegel und einer Kollimation der Strahlen, sodass die
    % fixen Linsen aus dem sichtbaren System problemlos eingefügt werden
    % können.

    for i=0:4
       switch i
           case 0
                Oberflaeche = TheLDE.GetSurfaceAt(i);
                Oberflaeche.Thickness = 45;                     % Entfernung der 1. Linse so festgelegt, dass eine 1 Zoll Linse bei 14 Grad Öffnungswinkel perfekt ausgenutzt wird
           case 1
                Oberflaeche = TheLDE.GetSurfaceAt(i);
                OddAsphere = Oberflaeche.GetSurfaceTypeSettings(ZOSAPI.Editors.LDE.SurfaceType.OddAsphere);
                Oberflaeche.ChangeType(OddAsphere);
                Oberflaeche.Thickness = 10;
                Oberflaeche.Material = 'PMMA';
                Oberflaeche.MechanicalSemiDiameter = 12.7;      %entspricht einer 1 Zoll Linse
                Oberflaeche.Radius = 20;
           case 2
                TheLDE.InsertNewSurfaceAt(i);
                Oberflaeche = TheLDE.GetSurfaceAt(i);
                OddAsphere = Oberflaeche.GetSurfaceTypeSettings(ZOSAPI.Editors.LDE.SurfaceType.OddAsphere);
                Oberflaeche.ChangeType(OddAsphere);
                Oberflaeche.Thickness = 30;
                Oberflaeche.MechanicalSemiDiameter = 12.7;
                Oberflaeche.Radius = 50;
           case 3
                TheLDE.InsertNewSurfaceAt(i);
                Oberflaeche = TheLDE.GetSurfaceAt(i);
                OddAsphere = Oberflaeche.GetSurfaceTypeSettings(ZOSAPI.Editors.LDE.SurfaceType.OddAsphere);
                Oberflaeche.ChangeType(OddAsphere);
                Oberflaeche.Thickness = 5;
                Oberflaeche.Material = 'PMMA';
                Oberflaeche.MechanicalSemiDiameter = 12.7;
                Oberflaeche.Radius = 70;
           case 4
                TheLDE.InsertNewSurfaceAt(i);
                Oberflaeche = TheLDE.GetSurfaceAt(i);
                OddAsphere = Oberflaeche.GetSurfaceTypeSettings(ZOSAPI.Editors.LDE.SurfaceType.OddAsphere);
                Oberflaeche.ChangeType(OddAsphere);
                Oberflaeche.Thickness = 130;                % Abstand zwischen Linse 2 und 3. Später sollen dort ein dichrotischer Spiegel und ein Umlenkspiegel eingesetzt werden
                Oberflaeche.MechanicalSemiDiameter = 12.7;
                Oberflaeche.Radius = -50;
       end
    end

    %--------------------- Linsendaten übertragen ---------------------------

    for i=5:8

            TheLDE.InsertNewSurfaceAt(i);
            Oberflaeche = TheLDE.GetSurfaceAt(i);
            OddAsphere = Oberflaeche.GetSurfaceTypeSettings(ZOSAPI.Editors.LDE.SurfaceType.OddAsphere);
            Oberflaeche.ChangeType(OddAsphere); 

    end


    for i=5:8
        Oberflaeche = TheLDE.GetSurfaceAt(i);
         % --------- Parameter zuweisen mit readCell
         % readCell durchsucht ein Array nach Oberflächennummer und Parameter,
         % dem man der Funktion gibt. Man erhält als Antwort die Position der
         % gesuchten Werte in dem Cell Array
            [num_pos, param_pos] = readCell(Oberflaechen_for_NIR,i,"Thickness");
            Oberflaeche.Thickness = Oberflaechen_for_NIR{param_pos,num_pos};
            [num_pos, param_pos] = readCell(Oberflaechen_for_NIR,i,"Radius");
            Oberflaeche.Radius = Oberflaechen_for_NIR{param_pos,num_pos};
            [num_pos, param_pos] = readCell(Oberflaechen_for_NIR,i,"Clear Semi-Diameter");
            Oberflaeche.SemiDiameter = Oberflaechen_for_NIR{param_pos,num_pos};
            [num_pos, param_pos] = readCell(Oberflaechen_for_NIR,i,"Mech Semi-Diameter");
            Oberflaeche.MechanicalSemiDiameter = Oberflaechen_for_NIR{param_pos,num_pos};
            [num_pos, param_pos] = readCell(Oberflaechen_for_NIR,i,"Comment");
            Oberflaeche.Comment = Oberflaechen_for_NIR{param_pos,num_pos};

        if Oberflaeche.Comment == 'Vorderseite'                         % Das Material wird nur an der Vorderseite übergeben
            [num_pos, param_pos] = readCell(Oberflaechen_for_NIR,i,"Material");
            Oberflaeche.Material = Oberflaechen_for_NIR{param_pos,num_pos};
        end

            [num_pos, param_pos] = readCell(Oberflaechen_for_NIR,i,"Conic");
            Oberflaeche.Conic = Oberflaechen_for_NIR{param_pos,num_pos};
            [num_pos, param_pos] = readCell(Oberflaechen_for_NIR,i,"1st Order Term");
            Oberflaeche.GetCellAt(12).DoubleValue = Oberflaechen_for_NIR{param_pos,num_pos}; 
            [num_pos, param_pos] = readCell(Oberflaechen_for_NIR,i,"2nd Order Term");
            Oberflaeche.GetCellAt(13).DoubleValue = Oberflaechen_for_NIR{param_pos,num_pos}; 
            [num_pos, param_pos] = readCell(Oberflaechen_for_NIR,i,"3rd Order Term");
            Oberflaeche.GetCellAt(14).DoubleValue = Oberflaechen_for_NIR{param_pos,num_pos};
            [num_pos, param_pos] = readCell(Oberflaechen_for_NIR,i,"4th Order Term");
            Oberflaeche.GetCellAt(15).DoubleValue = Oberflaechen_for_NIR{param_pos,num_pos};
            [num_pos, param_pos] = readCell(Oberflaechen_for_NIR,i,"5th Order Term");
            Oberflaeche.GetCellAt(16).DoubleValue = Oberflaechen_for_NIR{param_pos,num_pos};
            [num_pos, param_pos] = readCell(Oberflaechen_for_NIR,i,"6th Order Term");
            Oberflaeche.GetCellAt(17).DoubleValue = Oberflaechen_for_NIR{param_pos,num_pos};
            [num_pos, param_pos] = readCell(Oberflaechen_for_NIR,i,"7th Order Term");
            Oberflaeche.GetCellAt(18).DoubleValue = Oberflaechen_for_NIR{param_pos,num_pos};
            [num_pos, param_pos] = readCell(Oberflaechen_for_NIR,i,"8th Order Term");
            Oberflaeche.GetCellAt(19).DoubleValue = Oberflaechen_for_NIR{param_pos,num_pos};
    end


    % ---- Image Ebende

    Oberflaeche = TheLDE.GetSurfaceAt(TheLDE.NumberOfRows);
    Oberflaeche.SemiDiameter = SizeImage;
    Oberflaeche.MechanicalSemiDiameter = SizeImage;

    % ---- 2D-Layout öffnen

    TheSystem.Analyses.New_Analysis(ZOSAPI.Analysis.AnalysisIDM.Draw2D)

    % ------------------- Optimierung vorbereiten ---------------------------



    % ---- Variablen setzen

    tools = TheSystem.Tools;

    for i=1:4

        Oberflaeche = TheLDE.GetSurfaceAt(i);
        Oberflaeche.ThicknessCell.MakeSolveVariable();
        Oberflaeche.RadiusCell.MakeSolveVariable();
        Oberflaeche.SemiDiameterCell.MakeSolveVariable();
        Oberflaeche.MechanicalSemiDiameterCell.MakeSolveVariable();
        Oberflaeche.ConicCell.MakeSolveVariable();

        if i==4 %Die Thickness von Oberfläche 4 entspricht dem Abstand von Linse 2 und 3. Der soll aufgrund der später eingesetzten Spiegel vorerst nicht verändert werden
        Oberflaeche = TheLDE.GetSurfaceAt(i);
        Oberflaeche.ThicknessCell.MakeSolveFixed();
        end

        %Das sind die Order Terms der Oberfläche "Odd Asphere"
        %1st Order wird nich auf Variable gesetzt, da
        Oberflaeche.GetCellAt(13).MakeSolveVariable();  %2nd Order
        Oberflaeche.GetCellAt(14).MakeSolveVariable();  %3rd Order
        Oberflaeche.GetCellAt(15).MakeSolveVariable();  %usw.
        Oberflaeche.GetCellAt(16).MakeSolveVariable();
        Oberflaeche.GetCellAt(17).MakeSolveVariable();
        Oberflaeche.GetCellAt(18).MakeSolveVariable();
        Oberflaeche.GetCellAt(19).MakeSolveVariable();
    end

end

% -------------------- Merit Function erstellen ---------------------------

if TheSystem.SystemData.Fields.NumberOfFields == 1
    %Es wird hier unterschieden zwischen Matrix System und einzelner
    %(Laser-)Lichtquelle. Bei einem Field werden GENF + OPVA Operanden
    %gesetzt. Bei mehreren Fields (Matrix) werden CENY Operanden eingesetzt

    TheMFE = TheSystem.MFE;
    TheMFE.ShowEditor;
    n = 1;  %n wird immer dann erhöht wenn ein Operand (oder Operanden-Paar) hinzugefügt wurde, der einen Wert aus der xls_to_CENY Ausgabe bekommen hat
            %Die Variable zählt mit, bis die maximale Anzahl der Flächenschwerpunkte
            %der Diagonalen der Lichtverteilung erreicht ist

    for i=2:(size(CENY_Values,1))
        if CENY_Values{i,4} == 0
           %der Wert 0 wird ignoriert, da der ChiefRay sowieso nicht beeinflusst werden kann und an dieser Position ankommen wird. Kein Operand dafür notwendig  
        else
           %zunächst wird der GENF Operand eingefügt
           TheMFE.InsertNewOperandAt(n);
           Operand = TheMFE.GetOperandAt(n); 
           Operand.ChangeType(ZOSAPI.Editors.MFE.MeritOperandType.GENF);
           Operand.GetCellAt(3).IntegerValue = 1;      %Damit bezieht sich der Operand auf die Wellenlänge an Position 1 - hier 808nm
           Operand.GetCellAt(5).DoubleValue = CENY_Values{i,4}*10^3;
           Operand.Target = (i-1)*(1/(size(CENY_Values,1)-1)); %Target wird nicht gewichtet, nur für die Übersicht
           n = n+1;

           %danach wird der OPVA Operand eingefügt
           TheMFE.InsertNewOperandAt(n);
           Operand = TheMFE.GetOperandAt(n);
           Operand.ChangeType(ZOSAPI.Editors.MFE.MeritOperandType.OPVA);
           Operand.GetCellAt(2).IntegerValue = n-1;    %Damit bezieht sich der Operand OPVA auf den GENF Operand eine Zeile vorher
           Operand.Weight = 5;
           Operand.Target = (i-1)*(1/(size(CENY_Values,1)-1));
           n = n+1;
        end

    end




    % Einfügen von generellen Restriktionen, damit keine unlogischen Linsenformen
    % produziert werden
    for i=1:5
        TheMFE.InsertNewOperandAt(TheMFE.NumberOfOperands+1);
        Operand = TheMFE.GetOperandAt(TheMFE.NumberOfOperands);

        switch i          
            case 1
                Operand.ChangeType(ZOSAPI.Editors.MFE.MeritOperandType.MNCA);
                Operand.GetCellAt(2).IntegerValue = 1;
                Operand.GetCellAt(3).IntegerValue = 3;
                Operand.Target = 10;
                Operand.Weight = 1;

            case 2
                Operand.ChangeType(ZOSAPI.Editors.MFE.MeritOperandType.MXCA);
                Operand.GetCellAt(2).IntegerValue = 1;
                Operand.GetCellAt(3).IntegerValue = 3;
                Operand.Target = 80;
                Operand.Weight = 1;

            case 3
                Operand.ChangeType(ZOSAPI.Editors.MFE.MeritOperandType.MNEA);
                Operand.GetCellAt(2).IntegerValue = 1;
                Operand.GetCellAt(3).IntegerValue = 3;
                Operand.Target = 10;
                Operand.Weight = 1;

            case 4
                Operand.ChangeType(ZOSAPI.Editors.MFE.MeritOperandType.MNCG);
                Operand.GetCellAt(2).IntegerValue = 1;
                Operand.GetCellAt(3).IntegerValue = 4;
                Operand.Target = 4;
                Operand.Weight = 1;

            case 5
                Operand.ChangeType(ZOSAPI.Editors.MFE.MeritOperandType.MXCG);
                Operand.GetCellAt(2).IntegerValue = 1;
                Operand.GetCellAt(3).IntegerValue = 4;
                Operand.Target = 25;
                Operand.Weight = 1;

        end
    end


    TheSystem.UpdateStatus;
    TheMFE.CalculateMeritFunction; %updatet das System und die gesetzen Operanden in der MF
    SystExplorer = TheSystem.SystemData;
    
elseif TheSystem.SystemData.Fields.NumberOfFields > 1
    %Sind mehrere Fields vorhanden (Matrix) werden CENY Operanden verwendet
    n = 1;
    
    %------------- Field Editor -------------------
    % determine maximum field in Y only
    max_field = 0;
    for i=1:TheSystem.SystemData.Fields.NumberOfFields
        if TheSystem.SystemData.Fields.GetField(i).Y > max_field 
            max_field = TheSystem.SystemData.Fields.GetField(i).Y ;
        end
    end
    
    %alle Fields außer das an position Y=0 und Y=max_field werden gelöscht
    for i=2:TheSystem.SystemData.Fields.NumberOfFields-1
        SystExplorer.Fields.DeleteFieldAt(2);
    end
    
    %die neuen Fields werden gesetzt auf Grundlage der Größe der Tabelle
    %CENY_Values, die variieren kann. Die Positionen der Fields werden
    %verteilt von Y=0 bis zur größten vorhandenen Field Position Y=max_field
    for i=1:(size(CENY_Values,1)-1)
        SystExplorer.Fields.AddField(0, (i)*(max_field/(size(CENY_Values,1)-1)), 1.0);
        if i == (size(CENY_Values,1)-1)
            SystExplorer.Fields.DeleteFieldAt(2);
        end
    end
    
    %-------------- Merit-Function Editor ----------------
    
    [count, firstIndex, lastIndex] = Find_Operand('CENY'); 
    
        % Hier werden die CENY Operanden gelöscht, um sie an die neue Field
        % Anzahl anzupassen
    for i=firstIndex:lastIndex
        Operand = TheMFE.GetOperandAt(i);
        for n=1:count
            if strcmp(Operand.Type, 'CENY') == true
                TheMFE.DeleteRowAt(i-1);
            end
        end
    end
    
    n=1;    %Startpunkt zum Einsetzen der Operanden
    
    for i=2:(size(CENY_Values,1))
        
        TheMFE.InsertNewOperandAt(n);
        Operand = TheMFE.GetOperandAt(n); 
        Operand.ChangeType(ZOSAPI.Editors.MFE.MeritOperandType.CENY);
        Operand.GetCellAt(2).IntegerValue = TheLDE.NumberOfRows-1;      %Hier wird davon ausgegangen, dass die letzte Surface die Image Surface ist
        Operand.GetCellAt(5).IntegerValue = CENY_Values{i,1};
        Operand.Target = CENY_Values{i,4};
        Operand.Weight = 5;
        n = n+1;
        
    end
    
end
