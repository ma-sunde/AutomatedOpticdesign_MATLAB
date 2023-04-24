[num,txt,raw] = xlsread('Lichtverteilung_idealisiert.xlsx');    %Einlesen der Excel Datei

data = num(2:end, 2:end);       %1. Zeile und Spalte werden für Daten ignoriert, da hier Achsinformationen stehen,

yData = num(2:end,1);           %Achsinformationen werden ausgelesen
xData = num(1,2:end);

% pcolor hat gegenüber heatmap einige Formatierungs-Vorteile
p = pcolor(xData, yData, data);
colormap jet;                %Farbspektrum der colorbar
cb = colorbar;
caxis([min(data(:)) max(data(:))]);     %Festelegen, dass die gesamte colorbar mit den Werten ausgenutzt wird
cb.FontSize = 12;               %Schriftgröße der colorbar
cb.FontName = 'LM Roman 12';    %Schriftart der colorbar
ax = gca;                       %Achsinformationen zum manipulieren laden
ax.FontSize = 12;               %Schriftgröße der Achsen
ax.FontName = 'LM Roman 12';    %Schriftart der Achsen

%Achsbeschriftung:
xlabel('Horizontale Koordinaten in °');     
ylabel('Vertikale Koordinaten in °');

