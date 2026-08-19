%% =====================================================
%% analisis_pauV2.m
%%
%% PURPOSE
%% Global institutional analysis of examination results.
%%
%% INPUT
%% Excel workbook containing:
%%   CENTRO
%%   TIPO
%%   MATERIA
%%   TIENEEXAMEN
%%   CALIFICACION
%%   COORDENADA_X
%%   COORDENADA_Y
%%
%% OUTPUT
%% RESULTADOS_PAU\
%%   ranking_centros.xlsx
%%   histograma_global.png
%%   mapa_centros.png
%%   resultados_centros.shp
%%   informe_pauV2.docx
%%
%% =====================================================

clear
clc
close all

import mlreportgen.dom.*

%% =====================================================
%% CONFIGURACION
%% =====================================================

file = 'ejemplo_datos_1000_registros.xlsx';

outFolder = 'RESULTADOS_PAU';

if ~exist(outFolder,'dir')
    mkdir(outFolder)
end

%% =====================================================
%% LECTURA EXCEL
%% =====================================================

[~,~,raw] = xlsread(file);

headers = string(raw(2,:));
datos   = raw(3:end,:);

headers = matlab.lang.makeValidName(headers);

T = cell2table( ...
    datos,...
    'VariableNames',cellstr(headers));

fprintf('\nRecords read: %d\n',height(T));

%% =====================================================
%% COMPROBACION DE CAMPOS
%% =====================================================

camposNecesarios = { ...
    'CENTRO',...
    'TIENEEXAMEN',...
    'CALIFICACION',...
    'COORDENADA_X',...
    'COORDENADA_Y'};

for k = 1:numel(camposNecesarios)

    if ~ismember( ...
            camposNecesarios{k}, ...
            T.Properties.VariableNames)

        error( ...
            'Missing required field: %s', ...
            camposNecesarios{k})

    end

end

%% =====================================================
%% VARIABLES
%% =====================================================

centro = string(T.CENTRO);

if any(strcmp('MATERIA',T.Properties.VariableNames))

    materia = string(T.MATERIA);

else

    materia = string(T.TIPO);

end

examen = string(T.TIENEEXAMEN);

nota = str2double( ...
    string(T.CALIFICACION));

x = str2double( ...
    string(T.COORDENADA_X));

y = str2double( ...
    string(T.COORDENADA_Y));

%% =====================================================
%% VALIDACION
%% =====================================================

idx = ~isnan(nota);

centro = centro(idx);
materia = materia(idx);
nota = nota(idx);
x = x(idx);
y = y(idx);
examen = examen(idx);

idx = strcmpi( ...
    strtrim(examen), ...
    'SI');

centro = centro(idx);
materia = materia(idx);
nota = nota(idx);
x = x(idx);
y = y(idx);

idx = nota >= 0 & nota <= 10;

centro = centro(idx);
materia = materia(idx);
nota = nota(idx);
x = x(idx);
y = y(idx);

fprintf( ...
    'Valid records: %d\n', ...
    length(nota));

%% =====================================================
%% FUNCION ESTADISTICA
%% =====================================================

calc_stats = @(n) struct( ...
    'mean',mean(n), ...
    'median',median(n), ...
    'std',std(n), ...
    'n0',sum(n==0), ...
    'n_0_5',sum(n>0 & n<5), ...
    'n_5_7',sum(n>=5 & n<7), ...
    'n_7_9',sum(n>=7 & n<9), ...
    'n_9_10',sum(n>=9 & n<10), ...
    'n10',sum(n==10), ...
    'p_aprobado',sum(n>=5)/numel(n));

%% =====================================================
%% ANALISIS POR CENTRO
%% =====================================================

centrosUnicos = unique(centro);

global_stats = struct([]);

k = 1;

for i = 1:length(centrosUnicos)

    idx = centro == centrosUnicos(i);

    n = nota(idx);

    if numel(n) < 5
        continue
    end

    stats = calc_stats(n);

    global_stats(k).nombre = centrosUnicos(i);

    global_stats(k).mean = stats.mean;

    global_stats(k).median = stats.median;

    global_stats(k).std = stats.std;

    global_stats(k).p_aprobado = stats.p_aprobado;

    global_stats(k).score = ...
        0.4*stats.mean + ...
        0.2*stats.median - ...
        0.2*stats.std + ...
        0.2*stats.p_aprobado*10;

    k = k + 1;

end

G = struct2table(global_stats);

G.score = double(G.score);

G = sortrows( ...
    G,...
    'score',...
    'descend');

%% =====================================================
%% TOP Y BOTTOM
%% =====================================================

disp(' ')
disp('======== TOP 5 CENTRES ========')

disp(G(1:min(5,height(G)),:))

disp(' ')
disp('======== BOTTOM 5 CENTRES ========')

disp(G(max(1,height(G)-4):end,:))

%% =====================================================
%% RANKING
%% =====================================================

rankingFile = fullfile( ...
    outFolder,...
    'ranking_centros.xlsx');

writetable(G,rankingFile);

%% =====================================================
%% HISTOGRAMA GLOBAL
%% =====================================================

fig = figure('Visible','off');

histogram(nota,20)

xlabel('Score')
ylabel('Frequency')

title('Global score distribution')

histFile = fullfile( ...
    outFolder,...
    'histograma_global.png');

exportgraphics(fig,histFile)

close(fig)

%% =====================================================
%% MAPA DE CENTROS
%% =====================================================

Xc = zeros(height(G),1);
Yc = zeros(height(G),1);
Mc = zeros(height(G),1);

for i = 1:height(G)

    idx = centro == G.nombre(i);

    Xc(i) = mean( ...
        x(idx),...
        'omitnan');

    Yc(i) = mean( ...
        y(idx),...
        'omitnan');

    Mc(i) = mean( ...
        nota(idx),...
        'omitnan');

end

fig = figure('Visible','off');

scatter( ...
    Xc,...
    Yc,...
    80,...
    Mc,...
    'filled')

colorbar

xlabel('X')
ylabel('Y')

title('Mean score by centre')

mapFile = fullfile( ...
    outFolder,...
    'mapa_centros.png');

exportgraphics(fig,mapFile)

close(fig)

%% =====================================================
%% SHAPEFILE
%% =====================================================

if exist('shapewrite','file')

    clear S

    for i = 1:height(G)

        S(i).Geometry = 'Point';

        S(i).X = Xc(i);

        S(i).Y = Yc(i);

        S(i).Centro = char(G.nombre(i));

        S(i).Media = G.mean(i);

        S(i).Std = G.std(i);

    end

    shpFile = fullfile( ...
        outFolder,...
        'resultados_centros.shp');

    shapewrite(S,shpFile);

    disp('Shapefile generated')

else

    warning( ...
        'Mapping Toolbox not available. Shapefile not generated.')

end

%% =====================================================
%% INFORME WORD
%% =====================================================

docFile = fullfile( ...
    outFolder,...
    'informe_pauV2.docx');

doc = Document(docFile,'docx');

append(doc,...
    Heading(1,...
    'PAU INSTITUTIONAL REPORT'));

append(doc,...
    Paragraph(['Date: ' datestr(now)]));

append(doc,...
    Heading(2,...
    '1. Introduction'));

append(doc,...
    Paragraph( ...
    ['A total of ' ...
    num2str(length(nota)) ...
    ' valid examination records were analysed.']));

append(doc,...
    Heading(2,...
    '2. Global results'));

append(doc,...
    Paragraph(sprintf( ...
    'Global mean score: %.2f', ...
    mean(nota))));

append(doc,...
    Paragraph(sprintf( ...
    'Global standard deviation: %.2f', ...
    std(nota))));

append(doc,...
    Heading(2,...
    '3. Centre ranking'));

for i = 1:min(5,height(G))

    append(doc,...
        Paragraph(sprintf( ...
        '%d. %s (Mean %.2f)', ...
        i,...
        char(G.nombre(i)),...
        G.mean(i))));

end

append(doc,...
    Heading(2,...
    '4. Global distribution'));

img = Image(histFile);

img.Width = '12cm';

p = Paragraph();
p.HAlign = 'center';

append(p,img);
append(doc,p);

append(doc,...
    Paragraph( ...
    'Figure 1. Global distribution of scores.'));

append(doc,...
    Heading(2,...
    '5. Centre map'));

img = Image(mapFile);

img.Width = '12cm';

p = Paragraph();
p.HAlign = 'center';

append(p,img);
append(doc,p);

append(doc,...
    Paragraph( ...
    'Figure 2. Mean score by centre.'));

append(doc,...
    Heading(2,...
    '6. Conclusions'));

append(doc,...
    Paragraph( ...
    'Differences between educational centres were identified.'));

append(doc,...
    Paragraph( ...
    'The global score distribution reflects variability in academic performance.'));

close(doc)

%% =====================================================
%% FINAL
%% =====================================================

disp(' ')
disp('========================================')
disp('PAU ANALYSIS COMPLETED')
disp('========================================')
disp(['Generated: ' docFile])
disp(['Generated: ' rankingFile])
disp(['Generated: ' histFile])
disp(['Generated: ' mapFile])