clear
clc
close all

import mlreportgen.dom.*

%% =====================================================
%% analisis_materia_v4Ab11.m
%%
%% PURPOSE
%% Detailed subject-by-subject and centre-by-centre analysis.
%%
%% INPUT
%% MATERIAS\
%%   BIO\
%%   FIS\
%%   MAT\
%%   ...
%%
%% OUTPUT
%% RESULTADOS\
%%   INFORME_*.docx
%%   Resumen_Centros.xlsx
%%   hist_*.png
%%
%% The script generates:
%% - Statistics by centre
%% - Grade distributions
%% - Composite performance index
%% - Rankings
%% - Automatic Word reports
%%
%% =====================================================

%% =====================================================
%% ROOT FOLDER
%% =====================================================

rootFolder = 'MATERIAS';

if ~exist(rootFolder,'dir')

    error('Folder MATERIAS does not exist.')

end

D = dir(rootFolder);

D = D([D.isdir]);

D = D(~ismember({D.name},{'.','..'}));

%% =====================================================
%% SUBJECT LOOP
%% =====================================================

for k = 1:length(D)

    carpetaMateria = fullfile( ...
        rootFolder,...
        D(k).name);

    F = dir(fullfile(carpetaMateria,'*.xlsx'));

    if isempty(F)
        continue
    end

    file = fullfile( ...
        carpetaMateria,...
        F(1).name);

    fprintf('\n====================================\n');
    fprintf('Processing %s\n',file);
    fprintf('====================================\n');

    %% =================================================
    %% OUTPUT FOLDER
    %% =================================================

    outFolder = fullfile( ...
        carpetaMateria,...
        'RESULTADOS');

    if ~exist(outFolder,'dir')
        mkdir(outFolder)
    end

    %% =================================================
    %% READ EXCEL
    %% =================================================

    T = readtable(file);

    %% =================================================
    %% REQUIRED FIELDS
    %% =================================================

    camposNecesarios = { ...
        'CENTRO',...
        'TIPO',...
        'CALIFICACION',...
        'TIENEEXAMEN'};

    for kk = 1:numel(camposNecesarios)

        if ~ismember( ...
                camposNecesarios{kk}, ...
                T.Properties.VariableNames)

            error( ...
                'Missing required field: %s', ...
                camposNecesarios{kk})

        end

    end

    %% =================================================
    %% VARIABLES
    %% =================================================

    centro = string(T.CENTRO);

    materia = string(T.TIPO);

    nota = str2double( ...
        string(T.CALIFICACION));

    examen = string(T.TIENEEXAMEN);

    %% =================================================
    %% FILTER
    %% =================================================

    idx = ...
        ~isnan(nota) & ...
        nota >= 0 & ...
        nota <= 10 & ...
        strcmpi(strtrim(examen),'SI');

    centro  = centro(idx);
    materia = materia(idx);
    nota    = nota(idx);

    materiaUnica = unique(materia);

    if isempty(materiaUnica)
        continue
    end

    nombreMateria = char(materiaUnica(1));

    %% =================================================
    %% CENTRES
    %% =================================================

    centrosUnicos = unique(centro);

    %% =================================================
    %% WORD REPORT
    %% =================================================

    docFile = fullfile( ...
        outFolder,...
        ['INFORME_' ...
        matlab.lang.makeValidName(nombreMateria) ...
        '.docx']);

    if isfile(docFile)
        delete(docFile)
    end

    doc = Document(docFile,'docx');

    append(doc,...
        Heading(1,...
        ['SUBJECT REPORT: ' nombreMateria]));

    append(doc,...
        Paragraph(sprintf( ...
        'Number of centres analysed: %d', ...
        numel(centrosUnicos))));

    %% =================================================
    %% SUMMARY TABLE
    %% =================================================

    Resumen = table();

    %% =================================================
    %% CENTRE LOOP
    %% =================================================

    for c = 1:length(centrosUnicos)

        centroSel = centrosUnicos(c);

        idxc = centro == centroSel;

        n = nota(idxc);

        NEjercicios = numel(n);

        if NEjercicios == 0
            continue
        end

        %% =============================================
        %% STATISTICS
        %% =============================================

        media   = mean(n);
        mediana = median(n);
        stdv    = std(n);

        if media > 0
            cv = stdv/media;
        else
            cv = NaN;
        end

        minimo = min(n);
        maximo = max(n);

        %% =============================================
        %% DISTRIBUTION
        %% =============================================

        n0   = sum(n==0);
        n05  = sum(n>0 & n<5);
        n57  = sum(n>=5 & n<7);
        n79  = sum(n>=7 & n<9);
        n910 = sum(n>=9 & n<10);
        n10  = sum(n==10);

        p0   = 100*n0/NEjercicios;
        p05  = 100*n05/NEjercicios;
        p57  = 100*n57/NEjercicios;
        p79  = 100*n79/NEjercicios;
        p910 = 100*n910/NEjercicios;
        p10  = 100*n10/NEjercicios;

        %% =============================================
        %% COMPOSITE INDEX
        %% =============================================

        indice = ...
            media ...
            - 0.10*stdv ...
            + 0.20*p57/100 ...
            + 0.40*p79/100 ...
            + 0.60*p910/100 ...
            + 0.80*p10/100 ...
            - 0.40*p05/100;

        %% =============================================
        %% SUMMARY TABLE
        %% =============================================

        fila = table( ...
            centroSel,...
            NEjercicios,...
            media,...
            mediana,...
            stdv,...
            cv,...
            minimo,...
            maximo,...
            p0,...
            p05,...
            p57,...
            p79,...
            p910,...
            p10,...
            indice);

        Resumen = [Resumen; fila];

        %% =============================================
        %% HISTOGRAM
        %% =============================================

        safeCentro = matlab.lang.makeValidName( ...
            char(centroSel));

        fig = figure('Visible','off');

        histogram(n,20);

        xlabel('Score');
        ylabel('Frequency');

        title(char(centroSel));

        pngFile = fullfile( ...
            outFolder,...
            ['hist_' safeCentro '.png']);

        exportgraphics(fig,pngFile);

        close(fig);

        %% =============================================
        %% WORD REPORT
        %% =============================================

        append(doc,PageBreak);

        append(doc,...
            Heading(2,...
            char(centroSel)));

        append(doc,Heading(3,'Population'));

        append(doc,...
            Paragraph(sprintf( ...
            'Number of examinations: %d', ...
            NEjercicios)));

        append(doc,Heading(3,'Statistics'));

        append(doc,...
            Paragraph(sprintf('Mean: %.2f',media)));

        append(doc,...
            Paragraph(sprintf('Median: %.2f',mediana)));

        append(doc,...
            Paragraph(sprintf( ...
            'Standard deviation: %.2f',stdv)));

        append(doc,...
            Paragraph(sprintf( ...
            'Coefficient of variation: %.2f',cv)));

        append(doc,...
            Paragraph(sprintf('Minimum: %.2f',minimo)));

        append(doc,...
            Paragraph(sprintf('Maximum: %.2f',maximo)));

        append(doc,Heading(3,'Distribution (%)'));

        append(doc,Paragraph(sprintf('0 = %.2f',p0)));
        append(doc,Paragraph(sprintf('0-5 = %.2f',p05)));
        append(doc,Paragraph(sprintf('5-7 = %.2f',p57)));
        append(doc,Paragraph(sprintf('7-9 = %.2f',p79)));
        append(doc,Paragraph(sprintf('9-10 = %.2f',p910)));
        append(doc,Paragraph(sprintf('10 = %.2f',p10)));

        append(doc,Heading(3,'Composite index'));

        append(doc,...
            Paragraph(sprintf( ...
            'Value: %.2f',indice)));

        if isfile(pngFile)

            img = Image(pngFile);

            img.Width = '12cm';

            p = Paragraph();

            p.HAlign = 'center';

            append(p,img);

            append(doc,p);

        end

    end

    %% =================================================
    %% RANKING
    %% =================================================

    Resumen.Properties.VariableNames = { ...
        'Centro',...
        'NEjercicios',...
        'Media',...
        'Mediana',...
        'Std',...
        'CV',...
        'Minimo',...
        'Maximo',...
        'Pct0',...
        'Pct05',...
        'Pct57',...
        'Pct79',...
        'Pct910',...
        'Pct10',...
        'IndiceCompuesto'};

    Resumen = sortrows( ...
        Resumen,...
        'IndiceCompuesto',...
        'descend');

    Resumen.Ranking = (1:height(Resumen))';

    %% =================================================
    %% EXCEL OUTPUT
    %% =================================================

    excelOut = fullfile( ...
        outFolder,...
        'Resumen_Centros.xlsx');

    writetable(Resumen,excelOut);

    %% =================================================
    %% SAVE WORD
    %% =================================================

    close(doc);

    fprintf('Generated: %s\n',excelOut);

end

%% =====================================================
%% END
%% =====================================================

disp('========================================');
disp('PROCESS COMPLETED');
disp('========================================');