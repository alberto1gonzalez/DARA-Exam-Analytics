clear
clc
close all
% SCRIPT CORRESPONDIENTE A LA FASE 2.
% FASE 2. Análisis detallado por materia y centro
% % % Script: analisis_materia_v4Ab11.m------
%-----------------
% % % Entrada
% % % MATERIAS\*
% % % Salidas
% % % Dentro de cada materia:
    % % % RESULTADOS\
    % % % ├── INFORME_*.docx
    % % % ├── Resumen_Centros.xlsx
    % % % └── hist_*.png
    % % % Qué revisar
% % % Para cada materia:
% % % Existe
% % % Resumen_Centros.xlsx
% % % Existe
% % % INFORME_*.docx
% % % Existe
% % % hist_*.png
% % % Importante
% % % Abrir varios:
% % % Resumen_Centros.xlsx
% % % y verificar que aparecen columnas como:
% % % Plain Text
% % % Media
% % % Mediana
% % % Std
% % % CV
% % % IndiceCompuesto
% % % Ranking
% % % Ésta será la principal fuente de datos del informe final.


import mlreportgen.dom.*

%% =====================================================
%% CARPETA RAIZ
%% =====================================================

rootFolder = ...
'C:\Users\gonzalea\OneDrive - UNICAN\PAU\2026\extraordinaria\MATERIAS';

D = dir(rootFolder);
D = D([D.isdir]);
D = D(~ismember({D.name},{'.','..'}));

%% =====================================================
%% BUCLE DE MATERIAS
%% =====================================================

for k = 1:length(D)

    carpetaMateria = fullfile(rootFolder,D(k).name);

    F = dir(fullfile(carpetaMateria,'*.xlsx'));

    if isempty(F)
        continue
    end

    file = fullfile(carpetaMateria,F(1).name);

    fprintf('\n====================================\n');
    fprintf('Procesando %s\n',file);
    fprintf('====================================\n');

    %% =================================================
    %% RESULTADOS
    %% =================================================

    outFolder = fullfile(carpetaMateria,'RESULTADOS');

    if ~exist(outFolder,'dir')
        mkdir(outFolder)
    end

    %% =================================================
    %% LECTURA
    %% =================================================

    T = readtable(file);

    centro = string(T.CENTRO);
    materia = string(T.TIPO);
    nota = double(T.CALIFICACION);
    examen = string(T.TIENEEXAMEN);

    %% =================================================
    %% FILTRO
    %% =================================================

    idx = ~isnan(nota) & ...
          nota>=0 & ...
          nota<=10 & ...
          strcmpi(examen,'SI');

    centro = centro(idx);
    materia = materia(idx);
    nota = nota(idx);

    materiaUnica = unique(materia);

    if isempty(materiaUnica)
        continue
    end

    nombreMateria = char(materiaUnica(1));

    %% =================================================
    %% CENTROS
    %% =================================================

    centrosUnicos = unique(centro);

    %% =================================================
    %% INFORME WORD
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

    append(doc,Heading(1,...
        ['INFORME MATERIA: ' nombreMateria]));

    append(doc,Paragraph(sprintf( ...
        'Número de centros: %d', ...
        numel(centrosUnicos))));

    %% =================================================
    %% TABLA RESUMEN
    %% =================================================

    Resumen = table();

    %% =================================================
    %% LOOP DE CENTROS
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
        %% ESTADISTICOS
        %% =============================================

        media = mean(n);

        mediana = median(n);

        stdv = std(n);

        cv = stdv/media;

        minimo = min(n);

        maximo = max(n);

        %% =============================================
        %% DISTRIBUCION
        %% =============================================

        n0 = sum(n==0);

        n05 = sum(n>0 & n<5);

        n57 = sum(n>=5 & n<7);

        n79 = sum(n>=7 & n<9);

        n910 = sum(n>=9 & n<10);

        n10 = sum(n==10);

        p0 = 100*n0/NEjercicios;

        p05 = 100*n05/NEjercicios;

        p57 = 100*n57/NEjercicios;

        p79 = 100*n79/NEjercicios;

        p910 = 100*n910/NEjercicios;

        p10 = 100*n10/NEjercicios;

        %% =============================================
        %% INDICE COMPUESTO
        %% =============================================

        indice = media ...
            - 0.10*stdv ...
            + 0.20*p57/100 ...
            + 0.40*p79/100 ...
            + 0.60*p910/100 ...
            + 0.80*p10/100 ...
            - 0.40*p05/100;

        %% =============================================
        %% TABLA RESUMEN
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
        %% HISTOGRAMA
        %% =============================================

        safeCentro = matlab.lang.makeValidName( ...
            char(centroSel));

        fig = figure('Visible','off');

        histogram(n);

        xlabel('Calificación');
        ylabel('Frecuencia');

        title(char(centroSel));

        pngFile = fullfile( ...
            outFolder,...
            ['hist_' safeCentro '.png']);

        exportgraphics(fig,pngFile);

        close(fig);

        %% =============================================
        %% INFORME WORD
        %% =============================================

        append(doc,PageBreak);

        append(doc,Heading(2,char(centroSel)));

        append(doc,Heading(3,'Población'));

        append(doc,Paragraph(sprintf( ...
            'Número de ejercicios tratados: %d', ...
            NEjercicios)));

        append(doc,Heading(3,'Estadísticos'));

        append(doc,Paragraph(sprintf('Media: %.2f',media)));
        append(doc,Paragraph(sprintf('Mediana: %.2f',mediana)));
        append(doc,Paragraph(sprintf('Desviación estándar: %.2f',stdv)));
        append(doc,Paragraph(sprintf('Coeficiente de variación: %.2f',cv)));
        append(doc,Paragraph(sprintf('Mínimo: %.2f',minimo)));
        append(doc,Paragraph(sprintf('Máximo: %.2f',maximo)));

        append(doc,Heading(3,'Distribución (%)'));

        append(doc,Paragraph(sprintf('0: %.2f %%',p0)));
        append(doc,Paragraph(sprintf('0-5: %.2f %%',p05)));
        append(doc,Paragraph(sprintf('5-7: %.2f %%',p57)));
        append(doc,Paragraph(sprintf('7-9: %.2f %%',p79)));
        append(doc,Paragraph(sprintf('9-10: %.2f %%',p910)));
        append(doc,Paragraph(sprintf('10: %.2f %%',p10)));

        append(doc,Heading(3,'Índice compuesto'));

        append(doc,Paragraph(sprintf( ...
            'Valor: %.2f',indice)));

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
    %% EXCEL
    %% =================================================

    excelOut = fullfile( ...
        outFolder,...
        'Resumen_Centros.xlsx');

    writetable(Resumen,excelOut);

    %% =================================================
    %% GUARDAR WORD
    %% =================================================

    close(doc);

    fprintf('Generado: %s\n',excelOut);

end

disp('========================================');
disp('PROCESO COMPLETADO');
disp('========================================');