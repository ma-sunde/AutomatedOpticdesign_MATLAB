%verbindet Matlab mit Zemax OpticStudio. Es werden die wichtigsten Objekte
%geladen die fast immer gebraucht werden: LDE & MFE

TheApplication = MATLABZOSConnection;

% danach TheSystem:

TheSystem = TheApplication.PrimarySystem;


% über die Variable "TheSystem" lassen sich der Lens Data Editor (LDE) und der Merit Function Editor (MFE) laden
% LDE und MFE müssen zunächst geladen werden um Oberflächen oder Operanden beeinflussen zu können:
    
TheLDE = TheSystem.LDE;
    
TheMFE = TheSystem.MFE;

SystExplorer = TheSystem.SystemData;