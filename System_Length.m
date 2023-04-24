Total_Length = 0;
LastSurf = TheLDE.NumberOfSurfaces - 2;     % Der LDE nummeriert die Oberflächen von 0 aus und die Image Oberfläche hat keine Thickness
                                            % Deswegen müssen 2 Oberflächen von der Gesamtanzahl abgezogen werden.

for i=0:LastSurf
Surf = TheLDE.GetSurfaceAt(i);
Total_Length = Total_Length + Surf.Thickness;
end
