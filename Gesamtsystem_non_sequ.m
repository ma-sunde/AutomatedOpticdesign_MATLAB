%                                                                              %
%                                                                              %
%       ! Zuerst OpticStudio öffnen und im Reiter Programming auf den          %
%                Button Interactive Extension klicken !                        %                     
%                   Danach dieses Skript starten.                              %
%                                                                              %
%  


System_Load;

% Neues System erstellen ohne zu speichern
TheSystem.New(false);

%--------------------- NIR-System laden -----------------------------------

% volloptimiertes NIR-System im non-sequential mode laden. Theoretisch kann
% an dieser Stelle auch ein eigenes System eingefügt werden. 
DGfile = System.String.Concat(FolderOutput, FileNIR_NCE);
TheSystem.LoadFile(DGfile, false);

Get_All_Parameters_for_FIN_System;


%--------------------- VIS-System laden -----------------------------------

% volloptimiertes VIS-System im non-sequential mode laden. Theoretisch kann
% an dieser Stelle auch ein eigenes System eingefügt werden. 
DGfile = System.String.Concat(FolderOutput, FileVIS_NCE);
TheSystem.LoadFile(DGfile, false);

%---------------------- NIR-System hinzufügen -----------------------------

%--- Null Object
NullObject = TheNCE.InsertNewObjectAt(TheNCE.NumberOfRows+1);
NullObject.ChangeType(ObjSource.GetObjectTypeSettings(ZOSAPI.Editors.NCE.ObjectType.NullObject));
NullObject.Comment = 'Reference NIR-System';
NullObject.YPosition = 40;      % Versatz von VIS und NIR System

%--- Lichtquelle
ObjSource = TheNCE.InsertNewObjectAt(TheNCE.NumberOfRows+1)
if NSCDatenSource{1,2} == 'SourcePoint'
    ObjSource.ChangeType(ObjSource.GetObjectTypeSettings(ZOSAPI.Editors.NCE.ObjectType.SourcePoint));
elseif NSCDatenSource{1,2} == 'SourceRectangle'
    ObjSource.ChangeType(ObjSource.GetObjectTypeSettings(ZOSAPI.Editors.NCE.ObjectType.SourceRectangle));
end

%Parameter aus dem NIR-System übergeben
ObjSource.RefObject = NullObject.ObjectNumber;
ObjSource.ObjectData.Power = NSCDatenSource{14,2};
ObjSource.ObjectData.ConeAngle = NSCDatenSource{17,2};
ObjSource.ObjectData.WaveNumber = NSCDatenSource{15,2}; 
ObjSource.ObjectData.NumberOfLayoutRays = NSCDatenSource{10,2};
ObjSource.ObjectData.NumberOfAnalysisRays = NSCDatenSource{13,2};
ObjSource.Comment = NSCDatenSource{2,2};
ObjSource.InsideOf = NSCDatenSource{4,2};
ObjSource.XPosition = NSCDatenSource{5,2};
ObjSource.YPosition = NSCDatenSource{6,2};
ObjSource.ZPosition = NSCDatenSource{7,2};
ObjSource.TiltAboutX = NSCDatenSource{8,2};
ObjSource.TiltAboutY = NSCDatenSource{9,2};
ObjSource.TiltAboutZ = NSCDatenSource{10,2};

