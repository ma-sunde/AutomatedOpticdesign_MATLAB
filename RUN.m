Input_for_Start;

System_Load;
if isempty(FileVIS) == 0 && isempty(FolderVIS) == 0
    DGfile = System.String.Concat(FolderVIS, FileVIS);
    TheSystem.LoadFile(DGfile, false);
end

xls_to_CENY_V8;
Set_VIS_System;
Optimization_VIS_V4;
local_Optimization_VIS_V2;
Convert_VIS_to_NSC_and_Efficiency;
    
    TheSystem.SaveAs(System.String.Concat(FolderOutput, FileVIS_NCE));

Get_All_Parameters_from_NSC_V2;
Set_Full_Optimized_VIS_System;

    TheSystem.SaveAs(System.String.Concat(FolderOutput, FileVIS_LDE));
    
Differenz_Lichtverteilung;
xls_to_GENF_V2;    

if isempty(FileNIR) == 1 
    Get_Lenses_from_LDE_for_NIR_System;
    Set_NIR_System_V5;
else
    Get_Lenses_from_LDE_for_NIR_System;
    DGfile = System.String.Concat(FolderNIR, FileNIR);
    TheSystem.LoadFile(DGfile, false);
    Set_NIR_System_V5;
end

Optimization_NIR;
local_Optimization_NIR_V2;
Convert_NIR_to_NSC_and_Efficiency;

    TheSystem.SaveAs(System.String.Concat(FolderOutput, FileNIR_NCE));

Get_All_Parameters_from_NSC_V2;
Set_Full_Optimized_NIR_System;

    TheSystem.SaveAs(System.String.Concat(FolderOutput, FileNIR_LDE));