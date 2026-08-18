%% =====================================================
%% analisis_materia_v4Aa_PARAEXTRAORDINARIA.m
%% =====================================================

clear
clc
close all

import mlreportgen.dom.*

% % FASE 5. Análisis global por materias
% % Script: analisis_materia_v4Aa_PARAEXTRAORDINARIA.m
% % Entrada
% % listadoNotasProvisionales_070626_C4.xlsx
% % Salidas
    % % resultados_extraordinaria\
    % % ├── INFORME_GLOBAL.docx
    % % ├── hist_*.png
% % Qué revisar
% % Que exista
% % INFORME_GLOBAL.docx
% % Que existan
% % Plain Text
% % hist_*.png
% % para todas las materias.
% % Éstos serán los histogramas que incorporará el informe final.


%% =====================================================
%% RUTA
%% =====================================================

cd('C:\Users\gonzalea\OneDrive - UNICAN\PAU\2026\extraordinaria')

file = 'listadoNotasProvisionales_070626_C4.xlsx';

[~,nombreCarpeta] = fileparts(pwd);

outFolder = ['resultados_' nombreCarpeta];

if ~exist(outFolder,'dir')
    mkdir(outFolder)
end

%% =====================================================
%% LECTURA ROBUSTA DEL EXCEL
%% =====================================================

[~,~,raw] = xlsread(file);

headers = string(raw(2,:));
datos   = raw(3:end,:);

headers = matlab.lang.makeValidName(headers);

T = cell2table( ...
    datos,...
    'VariableNames',cellstr(headers));

%% =====================================================
%% COMPROBAR CAMPOS
%% =====================================================

campos = string(T.Properties.VariableNames);

if ~any(strcmp(campos,'TIPO'))
    error('No existe la columna TIPO')
end

if ~any(strcmp(campos,'CALIFICACION'))
    error('No existe la columna CALIFICACION')
end

if ~any(strcmp(campos,'TIENEEXAMEN'))
    error('No existe la columna TIENEEXAMEN')
end

%% =====================================================
%% VARIABLES
%% =====================================================

codigo = string(T.TIPO);

if any(strcmp(campos,'MATERIA'))

    nombre = string(T.MATERIA);
    materia = codigo + " - " + nombre;

else

    materia = codigo;

end

nota = str2double(string(T.CALIFICACION));
examen = string(T.TIENEEXAMEN);

%% =====================================================
%% POBLACION
%% =====================================================

total_raw = size(datos,1);

validos = sum(~isnan(nota));
realizados = sum(strcmpi(strtrim(examen),'SI'));
en_blanco = total_raw - validos;

%% =====================================================
%% FILTRO
%% =====================================================

idx = ...
    ~isnan(nota) & ...
    nota>=0 & ...
    nota<=10 & ...
    strcmpi(strtrim(examen),'SI');

materia = materia(idx);
nota    = nota(idx);

fprintf('\nRegistros válidos: %d\n',length(nota))

%% =====================================================
%% MATERIAS
%% =====================================================

materiasUnicas = unique(materia);

%% =====================================================
%% INFORME GLOBAL
%% =====================================================

finalFile = fullfile(outFolder,'INFORME_GLOBAL.docx');

if isfile(finalFile)
    delete(finalFile)
end

docG = Document(finalFile,'docx');

append(docG,...
    Heading(1,...
    'INFORME GLOBAL POR MATERIAS'));

%% =====================================================
%% POBLACION GLOBAL
%% =====================================================

append(docG,Heading(2,'Población global'));

append(docG,...
    Paragraph(sprintf( ...
    'Total registros: %d', ...
    total_raw)));

append(docG,...
    Paragraph(sprintf( ...
    'Realizados: %d', ...
    realizados)));

append(docG,...
    Paragraph(sprintf( ...
    'En blanco: %d', ...
    en_blanco)));

%% =====================================================
%% BUCLE MATERIAS
%% =====================================================

for m = 1:length(materiasUnicas)

    materiaSel = materiasUnicas(m);

    idxm = materia == materiaSel;

    n = nota(idxm);

    if numel(n) < 5
        continue
    end

    total = length(n);

    %% ===============================================
    %% ESTADISTICOS
    %% ===============================================

    media   = mean(n);
    mediana = median(n);
    stdv    = std(n);

    if media > 0
        cv = stdv/media;
    else
        cv = NaN;
    end

    %% ===============================================
    %% INTERPRETACION
    %% ===============================================

    if media < 5
        nivel_medio = 'bajo';
    elseif media < 7
        nivel_medio = 'medio';
    else
        nivel_medio = 'alto';
    end

    if stdv < 1.5
        dispersion_txt = 'homogénea';
    else
        dispersion_txt = 'heterogénea';
    end

    %% ===============================================
    %% DISTRIBUCION
    %% ===============================================

    n0   = sum(n==0);
    n05  = sum(n>0 & n<5);
    n57  = sum(n>=5 & n<7);
    n79  = sum(n>=7 & n<9);
    n910 = sum(n>=9 & n<10);
    n10  = sum(n==10);

    p0   = 100*n0/total;
    p05  = 100*n05/total;
    p57  = 100*n57/total;
    p79  = 100*n79/total;
    p910 = 100*n910/total;
    p10  = 100*n10/total;

    %% ===============================================
    %% INDICE COMPUESTO
    %% ===============================================

    indice = ...
        media ...
        - 0.1*stdv ...
        + 0.3*(p57+p79+p910)/100 ...
        - 0.3*p05/100;

    %% ===============================================
    %% FIGURA
    %% ===============================================

    safeName = matlab.lang.makeValidName(char(materiaSel));

    pngFile = fullfile( ...
        outFolder,...
        ['hist_' safeName '.png']);

    fig = figure('Visible','off');

    histogram(n,20,'Normalization','pdf')
    hold on

    xgrid = linspace(min(n),max(n),100);

    ydens = ksdensity(n,xgrid);

    plot(xgrid,ydens,...
        'r','LineWidth',2)

    xline(media,...
        'k','LineWidth',2)

    xline(mediana,...
        'b--','LineWidth',2)

    title(char(materiaSel))

    legend( ...
        'Histograma',...
        'Densidad',...
        'Media',...
        'Mediana',...
        'Location','best')

    exportgraphics(fig,pngFile)

    close(fig)

    %% ===============================================
    %% DOCUMENTO WORD
    %% ===============================================

    append(docG,PageBreak);

    append(docG,...
        Heading(2,...
        char(materiaSel)));

    append(docG,...
        Heading(3,...
        'Población'));

    append(docG,...
        Paragraph(sprintf( ...
        'Total registros: %d', ...
        total_raw)));

    append(docG,...
        Paragraph(sprintf( ...
        'Realizados: %d', ...
        realizados)));

    append(docG,...
        Paragraph(sprintf( ...
        'En blanco: %d', ...
        en_blanco)));

    append(docG,...
        Heading(3,...
        'Estadísticos principales'));

    append(docG,...
        Paragraph(sprintf( ...
        'Media: %.2f (%s)', ...
        media,...
        nivel_medio)));

    append(docG,...
        Paragraph(sprintf( ...
        'Mediana: %.2f', ...
        mediana)));

    append(docG,...
        Paragraph(sprintf( ...
        'Desviación estándar: %.2f (%s)', ...
        stdv,...
        dispersion_txt)));

    append(docG,...
        Paragraph(sprintf( ...
        'Coeficiente de variación: %.2f', ...
        cv)));

    append(docG,...
        Heading(3,...
        'Distribución de calificaciones'));

    append(docG,...
        Paragraph(sprintf('0: %.2f%%',p0)));

    append(docG,...
        Paragraph(sprintf('(0–5): %.2f%%',p05)));

    append(docG,...
        Paragraph(sprintf('(5–7): %.2f%%',p57)));

    append(docG,...
        Paragraph(sprintf('(7–9): %.2f%%',p79)));

    append(docG,...
        Paragraph(sprintf('(9–10): %.2f%%',p910)));

    append(docG,...
        Paragraph(sprintf('10: %.2f%%',p10)));

    append(docG,...
        Heading(3,...
        'Índice compuesto'));

    append(docG,...
        Paragraph(sprintf( ...
        'Valor: %.2f', ...
        indice)));

    if isfile(pngFile)

        append(docG,...
            Heading(3,...
            'Figura'));

        img = Image(pngFile);
        img.Width = '12cm';

        p = Paragraph();
        p.HAlign = 'center';

        append(p,img);
        append(docG,p);

    end

    append(docG,...
        Paragraph( ...
        ['Figura ' num2str(m) ...
        '. Distribución de ' ...
        char(materiaSel)]));

end

%% =====================================================
%% GUARDAR
%% =====================================================

close(docG)

disp(' ')
disp('========================================')
disp('INFORME GLOBAL GENERADO')
disp('========================================')
disp(finalFile)