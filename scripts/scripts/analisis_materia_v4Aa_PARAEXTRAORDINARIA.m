clear
clc
close all

import mlreportgen.dom.*

%% =====================================================
%% analisis_materia_v4Aa_PARAEXTRAORDINARIA.m
%%
%% PURPOSE
%% Global statistical analysis by subject.
%%
%% INPUT
%% Excel workbook containing:
%%   CENTRO
%%   TIPO
%%   MATERIA
%%   TIENEEXAMEN
%%   CALIFICACION
%%
%% OUTPUT
%% resultados_extraordinaria\
%%   INFORME_GLOBAL.docx
%%   hist_*.png
%%
%% =====================================================

%% =====================================================
%% INPUT FILE
%% =====================================================

file = 'ejemplo_datos_1000_registros.xlsx';

%% =====================================================
%% OUTPUT FOLDER
%% =====================================================

outFolder = 'resultados_extraordinaria';

if ~exist(outFolder,'dir')
    mkdir(outFolder)
end

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
    'CALIFICACION',...
    'TIENEEXAMEN'};

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

codigo = string(T.TIPO);

nombre = string(T.MATERIA);

materia = codigo + " - " + nombre;

nota = str2double( ...
    string(T.CALIFICACION));

examen = string(T.TIENEEXAMEN);

%% =====================================================
%% POPULATION
%% =====================================================

total_raw = size(datos,1);

validos = sum(~isnan(nota));

realizados = sum( ...
    strcmpi(strtrim(examen),'SI'));

en_blanco = total_raw - validos;

%% =====================================================
%% FILTER
%% =====================================================

idx = ...
    ~isnan(nota) & ...
    nota >= 0 & ...
    nota <= 10 & ...
    strcmpi(strtrim(examen),'SI');

materia = materia(idx);
nota    = nota(idx);

fprintf('\nValid records: %d\n',length(nota))

%% =====================================================
%% SUBJECTS
%% =====================================================

materiasUnicas = unique(materia);

%% =====================================================
%% GLOBAL REPORT
%% =====================================================

finalFile = fullfile( ...
    outFolder,...
    'INFORME_GLOBAL.docx');

if isfile(finalFile)
    delete(finalFile)
end

docG = Document(finalFile,'docx');

append(docG,...
    Heading(1,...
    'GLOBAL SUBJECT REPORT'));

%% =====================================================
%% GLOBAL POPULATION
%% =====================================================

append(docG,...
    Heading(2,...
    'Global population'));

append(docG,...
    Paragraph(sprintf( ...
    'Total records: %d', ...
    total_raw)));

append(docG,...
    Paragraph(sprintf( ...
    'Examinations taken: %d', ...
    realizados)));

append(docG,...
    Paragraph(sprintf( ...
    'Blank records: %d', ...
    en_blanco)));

%% =====================================================
%% SUBJECT LOOP
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
    %% STATISTICS
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
    %% INTERPRETATION
    %% ===============================================

    if media < 5

        nivel_medio = 'Low';

    elseif media < 7

        nivel_medio = 'Medium';

    else

        nivel_medio = 'High';

    end

    if stdv < 1.5

        dispersion_txt = 'Homogeneous';

    else

        dispersion_txt = 'Heterogeneous';

    end

    %% ===============================================
    %% DISTRIBUTION
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
    %% COMPOSITE INDEX
    %% ===============================================

    indice = ...
        media ...
        - 0.1*stdv ...
        + 0.3*(p57+p79+p910)/100 ...
        - 0.3*p05/100;

    %% ===============================================
    %% HISTOGRAM
    %% ===============================================

    safeName = matlab.lang.makeValidName( ...
        char(materiaSel));

    pngFile = fullfile( ...
        outFolder,...
        ['hist_' safeName '.png']);

    fig = figure('Visible','off');

    histogram(n,20,...
        'Normalization','pdf')

    hold on

    xgrid = linspace( ...
        min(n),...
        max(n),...
        100);

    ydens = ksdensity(n,xgrid);

    plot(xgrid,ydens,...
        'r',...
        'LineWidth',2)

    xline(media,...
        'k',...
        'LineWidth',2)

    xline(mediana,...
        'b--',...
        'LineWidth',2)

    xlabel('Score')
    ylabel('Probability density')

    title(char(materiaSel))

    legend( ...
        'Histogram',...
        'Density',...
        'Mean',...
        'Median',...
        'Location','best')

    exportgraphics(fig,pngFile)

    close(fig)

    %% ===============================================
    %% REPORT
    %% ===============================================

    append(docG,PageBreak);

    append(docG,...
        Heading(2,...
        char(materiaSel)));

    append(docG,...
        Heading(3,...
        'Statistics'));

    append(docG,...
        Paragraph(sprintf( ...
        'Mean: %.2f (%s)', ...
        media,...
        nivel_medio)));

    append(docG,...
        Paragraph(sprintf( ...
        'Median: %.2f', ...
        mediana)));

    append(docG,...
        Paragraph(sprintf( ...
        'Standard deviation: %.2f (%s)', ...
        stdv,...
        dispersion_txt)));

    append(docG,...
        Paragraph(sprintf( ...
        'Coefficient of variation: %.2f', ...
        cv)));

    append(docG,...
        Heading(3,...
        'Distribution'));

    append(docG,...
        Paragraph(sprintf('0: %.2f%%',p0)));

    append(docG,...
        Paragraph(sprintf('(0-5): %.2f%%',p05)));

    append(docG,...
        Paragraph(sprintf('(5-7): %.2f%%',p57)));

    append(docG,...
        Paragraph(sprintf('(7-9): %.2f%%',p79)));

    append(docG,...
        Paragraph(sprintf('(9-10): %.2f%%',p910)));

    append(docG,...
        Paragraph(sprintf('10: %.2f%%',p10)));

    append(docG,...
        Heading(3,...
        'Composite index'));

    append(docG,...
        Paragraph(sprintf( ...
        'Value: %.2f', ...
        indice)));

    if isfile(pngFile)

        img = Image(pngFile);

        img.Width = '12cm';

        p = Paragraph();

        p.HAlign = 'center';

        append(p,img);

        append(docG,p);

    end

end

%% =====================================================
%% SAVE
%% =====================================================

close(docG)

disp(' ')
disp('========================================')
disp('GLOBAL REPORT GENERATED')
disp('========================================')
disp(finalFile)