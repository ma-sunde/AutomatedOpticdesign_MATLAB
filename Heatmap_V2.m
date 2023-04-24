[num,txt,raw] = xlsread('Lichtverteilung_idealisiert.xlsx');

data = num(2:end, 2:end);

yData = num(2:end,1);
xData = num(1,2:end);

for i=1:length(xData)
    if mod(i,2) == 1
        xData_new(i) = xData(i);
    else
        i=i+1;
    end
end

h = heatmap(xData, yData, data);
ax = gca;
ax.FontSize = 12;
ax.FontName = 'Arial';
ax.FontName = 'LM Roman 12';
ax.XLabel = 'Horizontale Koordinaten in °';
ax.YLabel = 'Vertikale Koordinaten in °';


ax.XDisplayLabels = xData_new;
