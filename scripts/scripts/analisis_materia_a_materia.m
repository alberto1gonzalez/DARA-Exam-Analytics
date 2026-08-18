clear
clc
%------analisis_materia_a_materia.m---
%--------------------------------------
% % PREVIAMENTE SE DEBE VERIFICAR QUE LA FASE 0 ESTÁ COMPLETA

% FASE 0. Datos de partida
% % Fichero maestro
% % listadoNotasProvisionales_070626_C4.xlsx
% % Antes de ejecutar nada, verifica que existen las columnas:
% % CENTRO
% % TIPO
% % MATERIA
% % TIENEEXAMEN
% % CALIFICACION
% % COORDENADA_X
% % COORDENADA_Y
% % IMPORTANTE: Si cambian los nombres, todos los scripts SIGUIENTES posteriores deberán ajustarse.

% UNA VEZ VERIFICADO FASE0 SE PASA A FASE 1.  
% 
% FASE 1. Separación por materias
%SCRIPT:analisis_materia_a_materia.m
% % Entrada
% % listadoNotasProvisionales_070626_C4.xlsx
% % Salida esperada
    % % MATERIAS\
    % % ├── BIOLOGIA\
    % % ├── FISICA\
    % % ├── MATEMATICAS\
    % % ├── QUIMICA\
    % % └── ...
% % Dentro de cada carpeta:
% % Excel de una única materia
% % Qué revisar
% % Comprobar que:
% % Número de carpetas = número de materias
% % y que cada materia contiene su Excel.
% % Si aquí algo falla, todo lo demás fallará.

% file = 'CARRACUCASIM_10_06_26_listadoTotal_PAU_ord_2026_6_recupera2.xls';
file = 'listadoNotasProvisionales_070626_C4.xlsx';

% T = readtable(file);

[num,txt,raw] = xlsread(file);
headers = string(raw(2,:)); % nombres de campos
datos = raw(3:end,:); % datos
T = cell2table(datos,'VariableNames',matlab.lang.makeValidName(headers));

outFolder = 'MATERIAS';

if ~exist(outFolder,'dir')
    mkdir(outFolder)
end

materias = unique(string(T.TIPO));

for i = 1:length(materias)

    materia = materias(i);

    idx = string(T.TIPO) == materia;

    Tmat = T(idx,:);

    nombreMateria = matlab.lang.makeValidName(char(materia));

    carpetaMateria = fullfile(outFolder,nombreMateria);

    if ~exist(carpetaMateria,'dir')
        mkdir(carpetaMateria)
    end

    ficheroSalida = fullfile( ...
        carpetaMateria,...
        ['listadoNotasProvisionales_070626_C4' ...
        nombreMateria '.xlsx']);

    writetable(Tmat,ficheroSalida);

    fprintf('Generado: %s\n',ficheroSalida)

end

disp('FINALIZADO')