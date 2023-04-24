% Importieren der Excel-Tabelle
[data, text] = xlsread('Lichtverteilung_NIR.xlsx');


for i=2:size(data,1)
   for n=2:size(data,2) 
    data(i,n) = data(i,n)+5;
   end
end

   xlswrite('Neu.xlsx', data, 'Sheet1');
