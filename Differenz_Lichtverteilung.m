LichtverteilungVIS = xlsread(fullfile(FolderVIS_Lichtverteilung,FileVIS_Lichtverteilung));
LichtverteilungNIR = xlsread(fullfile(FolderNIR_Lichtverteilung,FileNIR_Lichtverteilung));

for i=1:size(LichtverteilungVIS,1)
    Lichtverteilung_Differenz(i,1) = LichtverteilungVIS(i,1);
end

for i=1:size(LichtverteilungVIS,2)
    Lichtverteilung_Differenz(1,i) = LichtverteilungVIS(1,i);
end

for i=2:size(LichtverteilungVIS,1)
   for n=2:size(LichtverteilungVIS,2) 
        Lichtverteilung_Differenz(i,n) = LichtverteilungNIR(i,n) - LichtverteilungVIS(i,n);
        if Lichtverteilung_Differenz(i,n) < 0
            Lichtverteilung_Differenz(i,n) = 0;
        end
   end
end

xlswrite([FolderNIR_Lichtverteilung 'Lichtverteilung_Differenz.xlsx'], Lichtverteilung_Differenz);