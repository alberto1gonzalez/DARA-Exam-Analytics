clear
clc
close all

import mlreportgen.dom.*

%% =====================================================
%% analisis_materia_v4_PARAEXTRAORDINARIA.m
%%
%% PURPOSE
%% Subject-by-subject statistical analysis including
%% PNG and EMF histogram generation.
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
%%   hist_*.png
%%   hist_*.emf
%%   informe_*.docx
%%   INFORME_GLOBAL.docx
%%
%% =====================================================

%% ============================
%% INPUT FILE
%% ============================

file = 'ejemplo_datos_1000_registros.xlsx';

%% ============================
%% OUTPUT FOLDER
%% ============================

outFolder = 'resultados_extraordinaria';

if ~exist(outFolder,'dir')
    mkdir(outFolder)
end

%% ============================
%% READ EXCEL
%% ============================

[~,~,raw] = xlsread(file);

headers = string(raw(1,:));
datos   = raw(2:end,:);

headers = matlab.lang.makeValidName(headers);

T = cell2table( ...
    datos,...
    'VariableNames',cellstr(headers));

%% ============================
%% REQUIRED FIELDS
%% ============================

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

%% ============================
%% VARIABLES
%% ============================

codigo = string(T.TIPO);

nombre = string(T.MATERIA);

materia = codigo + " - " + nombre;

examen = string(T.TIENEEXAMEN);

nota = str2double(string(T.CALIFICACION));

%% ============================
%% VALIDATION
%% ============================

idx = ~isnan(nota);

materia = materia(idx);
nota = nota(idx);
examen = examen(idx);

idx = strcmpi(strtrim(examen),'SI');

materia = materia(idx);
nota = nota(idx);

idx = nota >= 0 & nota <= 10;

materia = materia(idx);
nota = nota(idx);

fprintf('\nValid records: %d\n',length(nota))

%% ============================
%% SUBJECT LIST
%% ============================

materiasUnicas = unique(materia);

disp('--- SUBJECTS AVAILABLE ---')

for i = 1:length(materiasUnicas)

    fprintf('%2d - %s\n',i,materiasUnicas(i))

end

%% ============================
%% MAIN LOOP
%% ============================

listaProcesadas = {};

while true

    modo = input( ...
        '\n1: Select subject | 2: Process ALL | 0: Exit -> ');

    if modo == 0

        break

    elseif modo == 1

        id = input('Select subject number: ');

        lista = materiasUnicas(id);

    elseif modo == 2

        lista = materiasUnicas;

    else

        disp('Invalid option')
        continue

    end

    for m = 1:length(lista)

        materiaSel = lista(m);

        fprintf('\nProcessing: %s\n',materiaSel)

        idxm = materia == materiaSel;

        n = nota(idxm);

        if numel(n) < 5
            continue
        end

        %% ===================================
        %% STATISTICS
        %% ===================================

        media = mean(n);
        mediana = median(n);
        stdv = std(n);

        if media > 0
            cv = stdv/media;
        else
            cv = NaN;
        end

        suspensos = sum(n < 5);

        total = length(n);

        psus = 100*suspensos/total;

        %% ===================================
        %% INTERPRETATION
        %% ===================================

        if abs(media-mediana) < 0.3

            simetria = 'symmetric';

        else

            simetria = 'asymmetric';

        end

        %% ===================================
        %% FIGURE
        %% ===================================

        fig = figure('Visible','off');

        histogram(n,20,'Normalization','pdf')

        hold on

        xgrid = linspace(min(n),max(n),100);

        y = ksdensity(n,xgrid);

        plot(xgrid,y,...
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
            'Median')

        safeName = matlab.lang.makeValidName( ...
            char(materiaSel));

        pngFile = fullfile( ...
            outFolder,...
            ['hist_' safeName '.png']);

        emfFile = fullfile( ...
            outFolder,...
            ['hist_' safeName '.emf']);

        exportgraphics(fig,pngFile);

        print(fig,'-dmeta',emfFile);

        close(fig)

        %% ===================================
        %% INDIVIDUAL REPORT
        %% ===================================

        docName = fullfile( ...
            outFolder,...
            ['informe_' safeName '.docx']);

        doc = Document(docName,'docx');

        append(doc,...
            Heading(1,...
            ['ANALYSIS: ' char(materiaSel)]));

        append(doc,Heading(2,'Statistics'));

        append(doc,...
            Paragraph(sprintf('Mean: %.2f',media)));

        append(doc,...
            Paragraph(sprintf('Median: %.2f',mediana)));

        append(doc,...
            Paragraph(sprintf('Std: %.2f',stdv)));

        append(doc,...
            Paragraph(sprintf( ...
            'Coefficient of variation: %.2f',cv)));

        append(doc,Heading(2,'Interpretation'));

        append(doc,...
            Paragraph(['Distribution ' simetria]));

        append(doc,...
            Paragraph(sprintf( ...
            'Failures: %.2f%% (%s)', ...
            psus,...
            nivel(psus))));

        append(doc,Heading(2,'Figure'));

        if isfile(emfFile)

            img = Image(emfFile);

            img.Width = '12cm';

            p = Paragraph();

            p.HAlign = 'center';

            append(p,img);

            append(doc,p);

        end

        append(doc,...
            Paragraph( ...
            'Histogram, density curve, mean and median.'));

        close(doc)

        listaProcesadas{end+1} = char(materiaSel);

        fprintf('Generated: %s\n',docName)

    end

end

%% ============================
%% GLOBAL REPORT
%% ============================

if ~isempty(listaProcesadas)

    finalFile = fullfile( ...
        outFolder,...
        'INFORME_GLOBAL.docx');

    finalDoc = Document(finalFile,'docx');

    append(finalDoc,...
        Heading(1,...
        'GLOBAL SUBJECT REPORT'));

    for m = 1:length(listaProcesadas)

        materiaSel = string(listaProcesadas{m});

        append(finalDoc,PageBreak);

        append(finalDoc,...
            Heading(2,...
            char(materiaSel)));

        idxm = materia == materiaSel;

        n = nota(idxm);

        media = mean(n);
        mediana = median(n);
        stdv = std(n);

        suspensos = sum(n<5);

        total = length(n);

        psus = 100*suspensos/total;

        append(finalDoc,...
            Paragraph(sprintf('Mean: %.2f',media)));

        append(finalDoc,...
            Paragraph(sprintf('Median: %.2f',mediana)));

        append(finalDoc,...
            Paragraph(sprintf( ...
            'Standard deviation: %.2f',stdv)));

        append(finalDoc,...
            Paragraph(sprintf( ...
            'Failures: %.2f%% (%s)', ...
            psus,...
            nivel(psus))));

        safeName = matlab.lang.makeValidName( ...
            char(materiaSel));

        emfFile = fullfile( ...
            outFolder,...
            ['hist_' safeName '.emf']);

        if isfile(emfFile)

            img = Image(emfFile);

            img.Width = '12cm';

            p = Paragraph();

            p.HAlign = 'center';

            append(p,img);

            append(finalDoc,p);

        end

    end

    close(finalDoc)

    disp(' ')
    disp('========================================')
    disp('GLOBAL REPORT GENERATED')
    disp('========================================')
    disp(finalFile)

end

%% ============================
%% FUNCTION
%% ============================

function out = nivel(p)

if p > 50

    out = 'high';

elseif p > 30

    out = 'moderate';

else

    out = 'low';

end

end