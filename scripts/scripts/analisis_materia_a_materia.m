clear
clc

%% =====================================================
%% analisis_materia_a_materia.m
%%
%% PURPOSE
%% Splits a master Excel file into independent Excel
%% files, one for each subject.
%%
%% INPUT
%% Excel workbook containing at least:
%%   CENTRO
%%   TIPO
%%   MATERIA
%%   TIENEEXAMEN
%%   CALIFICACION
%%
%% OPTIONAL FIELDS
%%   COORDENADA_X
%%   COORDENADA_Y
%%
%% OUTPUT
%% MATERIAS\
%%   1\
%%   2\
%%   3\
%%   ...
%%
%% Each folder contains an Excel file with all records
%% corresponding to a single subject.
%%
%% =====================================================

%% =====================================================
%% INPUT FILE
%% =====================================================

file = 'ejemplo_datos_1000_registros.xlsx';

%% =====================================================
%% READ EXCEL
%% =====================================================

[~,~,raw] = xlsread(file);

headers = string(raw(1,:));
datos   = raw(2:end,:);

headers = matlab.lang.makeValidName(headers);

T = cell2table( ...
    datos,...
    'VariableNames',cellstr(headers));

%% =====================================================
%% REQUIRED FIELDS
%% =====================================================

camposNecesarios = { ...
    'TIPO',...
    'MATERIA',...
    'TIENEEXAMEN',...
    'CALIFICACION'};

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
%% OUTPUT FOLDER
%% =====================================================

outFolder = 'MATERIAS';

if ~exist(outFolder,'dir')
    mkdir(outFolder)
end

%% =====================================================
%% SUBJECT LIST
%% =====================================================

materias = unique(string(T.TIPO));

fprintf('\n')
fprintf('========================================\n')
fprintf('SUBJECTS FOUND: %d\n',length(materias))
fprintf('========================================\n')

%% =====================================================
%% SPLIT DATASET
%% =====================================================

for i = 1:length(materias)

    materia = materias(i);

    idx = string(T.TIPO) == materia;

    Tmat = T(idx,:);

    nombreMateria = matlab.lang.makeValidName( ...
        char(materia));

    carpetaMateria = fullfile( ...
        outFolder,...
        nombreMateria);

    if ~exist(carpetaMateria,'dir')
        mkdir(carpetaMateria)
    end

    ficheroSalida = fullfile( ...
        carpetaMateria,...
        [nombreMateria '.xlsx']);

    writetable(Tmat,ficheroSalida);

    fprintf('Generated: %s\n',ficheroSalida)

end

%% =====================================================
%% FINISHED
%% =====================================================

fprintf('\n')
fprintf('========================================\n')
fprintf('PROCESS COMPLETED\n')
fprintf('========================================\n')
fprintf('Output folder: %s\n',outFolder)
fprintf('\n')
