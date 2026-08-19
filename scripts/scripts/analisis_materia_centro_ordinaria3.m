%% =====================================================
%% analisis_materia_centro_ordinaria3.m
%%
%% PURPOSE
%% Master statistical table:
%% Centre × Subject
%%
%% INPUT
%% Excel workbook containing:
%%   CENTRO
%%   TIPO
%%   CALIFICACION
%%   TIENEEXAMEN
%%   COORDENADA_X
%%   COORDENADA_Y
%%
%% OUTPUT
%% RESULTADOS_MAESTROS\
%%      resumen_materia_centro.xlsx
%%
%% This file is the main statistical product used by
%% subsequent scripts and final institutional reports.
%%
%% =====================================================

clear
clc
close all

%% =====================================================
%% CONFIGURATION
%% =====================================================

file = 'ejemplo_datos_1000_registros.xlsx';

outFolder = 'RESULTADOS_MAESTROS';

if ~exist(outFolder,'dir')
    mkdir(outFolder)
end

excelFolder = fullfile( ...
    outFolder,...
    'resumenes_excel');

if ~exist(excelFolder,'dir')
    mkdir(excelFolder)
end

%% =====================================================
%% READ EXCEL
%% =====================================================

[~,~,raw] = xlsread(file);

headers = string(raw(2,:));

datos = raw(3:end,:);

headers = matlab.lang.makeValidName(headers);

T = cell2table( ...
    datos,...
    'VariableNames', ...
    cellstr(headers));

fprintf('\nRecords read: %d\n',height(T));

%% =====================================================
%% REQUIRED FIELDS
%% =====================================================

camposNecesarios = { ...
    'CENTRO',...
    'TIPO',...
    'CALIFICACION',...
    'TIENEEXAMEN',...
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

materia = string(T.TIPO);

nota = str2double( ...
    string(T.CALIFICACION));

X = str2double( ...
    string(T.COORDENADA_X));

Y = str2double( ...
    string(T.COORDENADA_Y));

tieneExamen = string(T.TIENEEXAMEN);

%% =====================================================
%% FILTER
%% =====================================================

idx = ...
    strcmpi(strtrim(tieneExamen),'SI') & ...
    ~isnan(nota) & ...
    nota >= 0 & ...
    nota <= 10;

centro  = centro(idx);
materia = materia(idx);
nota    = nota(idx);
X       = X(idx);
Y       = Y(idx);

fprintf( ...
    'Valid examination records: %d\n', ...
    length(nota));

%% =====================================================
%% GROUPS
%% =====================================================

[G,CentroU,MateriaU] = ...
    findgroups(centro,materia);

nGrupos = max(G);

fprintf( ...
    'Centre-Subject groups: %d\n', ...
    nGrupos);

%% =====================================================
%% RESULTS TABLE
%% =====================================================

Resumen = table();

%% =====================================================
%% MAIN LOOP
%% =====================================================

for k = 1:nGrupos

    idx = (G == k);

    centroSel  = CentroU(k);

    materiaSel = MateriaU(k);

    notasGrupo = nota(idx);

    Xg = X(idx);

    Yg = Y(idx);

    Ntotal = numel(notasGrupo);

    N_NP = sum(isnan(notasGrupo));

    nv = notasGrupo(~isnan(notasGrupo));

    Nvalidos = numel(nv);

    if isempty(nv)
        continue
    end

    %% -----------------------------------------------
    %% STATISTICS
    %% -----------------------------------------------

    Media = mean(nv);

    Mediana = median(nv);

    Std = std(nv);

    if Media > 0

        CV = 100*Std/Media;

    else

        CV = NaN;

    end

    Minimo = min(nv);

    Maximo = max(nv);

    %% -----------------------------------------------
    %% CLASSES
    %% -----------------------------------------------

    N0   = sum(nv==0);

    N05  = sum(nv>0 & nv<5);

    N57  = sum(nv>=5 & nv<7);

    N79  = sum(nv>=7 & nv<9);

    N910 = sum(nv>=9 & nv<10);

    N10  = sum(nv==10);

    %% -----------------------------------------------
    %% PERCENTAGES
    %% -----------------------------------------------

    PctNP  = 100*N_NP/Ntotal;

    Pct0   = 100*N0/Ntotal;

    Pct05  = 100*N05/Ntotal;

    Pct57  = 100*N57/Ntotal;

    Pct79  = 100*N79/Ntotal;

    Pct910 = 100*N910/Ntotal;

    Pct10  = 100*N10/Ntotal;

    %% -----------------------------------------------
    %% COMPOSITE INDEX
    %% -----------------------------------------------

    IndiceCompuesto = ...
          Media ...
        - 0.10*Std ...
        + 0.20*Pct57/100 ...
        + 0.40*Pct79/100 ...
        + 0.60*Pct910/100 ...
        + 0.80*Pct10/100 ...
        - 0.40*Pct05/100;

    %% -----------------------------------------------
    %% CENTROID
    %% -----------------------------------------------

    Xcentroide = mean(Xg,'omitnan');

    Ycentroide = mean(Yg,'omitnan');

    %% -----------------------------------------------
    %% OUTPUT ROW
    %% -----------------------------------------------

    fila = table( ...
        centroSel,...
        materiaSel,...
        Ntotal,...
        Nvalidos,...
        Media,...
        Mediana,...
        Std,...
        CV,...
        Minimo,...
        Maximo,...
        N_NP,...
        N0,...
        N05,...
        N57,...
        N79,...
        N910,...
        N10,...
        PctNP,...
        Pct0,...
        Pct05,...
        Pct57,...
        Pct79,...
        Pct910,...
        Pct10,...
        IndiceCompuesto,...
        Xcentroide,...
        Ycentroide);

    Resumen = [Resumen ; fila];

end

%% =====================================================
%% COLUMN NAMES
%% =====================================================

Resumen.Properties.VariableNames = { ...
'Centro',...
'Materia',...
'NTotal',...
'NValidos',...
'Media',...
'Mediana',...
'Std',...
'CV',...
'Minimo',...
'Maximo',...
'N_NP',...
'N0',...
'N05',...
'N57',...
'N79',...
'N910',...
'N10',...
'PctNP',...
'Pct0',...
'Pct05',...
'Pct57',...
'Pct79',...
'Pct910',...
'Pct10',...
'IndiceCompuesto',...
'X_COORD',...
'Y_COORD'};

%% =====================================================
%% GLOBAL RANKING
%% =====================================================

Resumen = sortrows( ...
    Resumen,...
    'IndiceCompuesto',...
    'descend');

Resumen.RankingGlobal = ...
    (1:height(Resumen))';

%% =====================================================
%% SUBJECT RANKING
%% =====================================================

Resumen.RankingMateria = ...
    zeros(height(Resumen),1);

materias = unique(Resumen.Materia);

for i = 1:length(materias)

    idx = Resumen.Materia == materias(i);

    pos = find(idx);

    [~,ord] = sort( ...
        Resumen.IndiceCompuesto(idx), ...
        'descend');

    ranking = zeros(sum(idx),1);

    ranking(ord) = 1:length(ord);

    Resumen.RankingMateria(pos) = ranking;

end

%% =====================================================
%% EXPORT EXCEL
%% =====================================================

excelFile = fullfile( ...
    excelFolder,...
    'resumen_materia_centro.xlsx');

writetable(Resumen,excelFile);

%% =====================================================
%% END
%% =====================================================

disp(' ')
disp('========================================')
disp('MASTER TABLE GENERATED')
disp('========================================')
disp(excelFile)