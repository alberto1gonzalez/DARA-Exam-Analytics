%% =====================================================
%% INFORME_FINAL_NUEVOV2.m
%%
%% Final integrated institutional report
%%
%% Requiere:
%% MATLAB Report Generator
%%
%% =====================================================

clear
clc
close all

import mlreportgen.dom.*

%% =====================================================
%% CONFIGURACION
%% =====================================================

nombreInforme = ...
    'INFORME_GLOBAL_EXTRAORDINARIA_V2';

doc = Document(nombreInforme,'docx');

%% =====================================================
%% RUTAS
%% =====================================================

folderPAU = 'RESULTADOS_PAU';

folderMaestro = fullfile( ...
    'RESULTADOS_MAESTROS', ...
    'resumenes_excel');

folderMaterias = ...
    'resultados_extraordinaria';

%% =====================================================
%% PORTADA
%% =====================================================

append(doc,...
    Heading(1,...
    'INFORME GLOBAL EXTRAORDINARIA'));

append(doc,...
    Paragraph(['Fecha de generación: ' ...
    datestr(now)]));

%% =====================================================
%% INTRODUCCION
%% =====================================================

append(doc,Heading(2,'1. Introducción'));

append(doc,...
    Paragraph(['Documento generado automáticamente ' ...
    'a partir de los diferentes procesos de análisis ' ...
    'desarrollados para la explotación estadística de ' ...
    'resultados académicos procedentes de sistemas ' ...
    'DARA / LECTODARA.']));

%% =====================================================
%% METODOLOGIA
%% =====================================================

append(doc,Heading(2,'2. Metodología'));

append(doc,...
    Paragraph('Scripts utilizados:'));

append(doc,...
    Paragraph('- analisis_materia_a_materia.m'));

append(doc,...
    Paragraph('- analisis_materia_v4Ab11.m'));

append(doc,...
    Paragraph('- analisis_materia_centro_ordinaria3.m'));

append(doc,...
    Paragraph('- analisis_pauV2.m'));

append(doc,...
    Paragraph('- analisis_materia_v4Aa_PARAEXTRAORDINARIA.m'));

append(doc,...
    Paragraph('- analisis_materia_v4_PARAEXTRAORDINARIA.m'));

%% =====================================================
%% RESULTADOS GLOBALES
%% =====================================================

append(doc,Heading(2,'3. Resultados globales'));

histFile = fullfile( ...
    folderPAU,...
    'histograma_global.png');

if isfile(histFile)

    img = Image(histFile);

    img.Width = '14cm';

    append(doc,img);

    append(doc,...
        Paragraph('Figura 1. Distribución global de calificaciones.'));

end

mapFile = fullfile( ...
    folderPAU,...
    'mapa_centros.png');

if isfile(mapFile)

    img = Image(mapFile);

    img.Width = '14cm';

    append(doc,img);

    append(doc,...
        Paragraph('Figura 2. Distribución espacial de resultados por centro.'));

end

%% =====================================================
%% RANKING DE CENTROS
%% =====================================================

append(doc,Heading(2,'4. Ranking de centros'));

rankingFile = fullfile( ...
    folderPAU,...
    'ranking_centros.xlsx');

mejorCentro = '';
peorCentro = '';

if isfile(rankingFile)

    G = readtable(rankingFile);

    mejorCentro = string(G.nombre(1));

    peorCentro = string(G.nombre(end));

    tabla = Table();

    h = TableRow();

    append(h,TableEntry('Posición'));
    append(h,TableEntry('Centro'));

    append(tabla,h);

    for i=1:min(10,height(G))

        r = TableRow();

        append(r,...
            TableEntry(num2str(i)));

        append(r,...
            TableEntry(char(G.nombre(i))));

        append(tabla,r);

    end

    append(doc,tabla);

end

%% =====================================================
%% TABLA MAESTRA
%% =====================================================

append(doc,Heading(2,'5. Indicador global de materias'));

excelMaestro = fullfile( ...
    folderMaestro,...
    'resumen_materia_centro.xlsx');

mejorMateria = '';
peorMateria = '';

if isfile(excelMaestro)

    R = readtable(excelMaestro);

    materias = groupsummary( ...
        R,...
        'Materia',...
        'mean',...
        'IndiceCompuesto');

    materias = sortrows( ...
        materias,...
        'mean_IndiceCompuesto',...
        'descend');

    mejorMateria = string(materias.Materia(1));

    peorMateria = string(materias.Materia(end));

    tabla = Table();

    h = TableRow();

    append(h,TableEntry('Materia'));
    append(h,TableEntry('Indice'));

    append(tabla,h);

    for i = 1:min(20,height(materias))

        rr = TableRow();

        append(rr,...
            TableEntry( ...
            char(materias.Materia(i))));

        append(rr,...
            TableEntry( ...
            sprintf('%.2f', ...
            materias.mean_IndiceCompuesto(i))));

        append(tabla,rr);

    end

    append(doc,tabla);

end

%% =====================================================
%% RESULTADOS POR MATERIAS
%% =====================================================

append(doc,Heading(2,'6. Resultados por materias'));

pngs = dir(fullfile( ...
    folderMaterias,...
    'hist_*.png'));

for k = 1:length(pngs)

    append(doc,PageBreak);

    nombreMateria = ...
        erase(pngs(k).name,'hist_');

    nombreMateria = ...
        erase(nombreMateria,'.png');

    append(doc,...
        Heading(3,...
        nombreMateria));

    img = Image(fullfile( ...
        pngs(k).folder,...
        pngs(k).name));

    img.Width = '12cm';

    append(doc,img);

    append(doc,...
        Paragraph('Distribución de calificaciones.'));

end

%% =====================================================
%% RESULTADOS POR CENTROS
%% =====================================================

append(doc,Heading(2,'7. Resultados por centros'));

exceles = dir(fullfile( ...
    'MATERIAS',...
    '**',...
    'Resumen_Centros.xlsx'));

append(doc,...
    Paragraph(sprintf( ...
    'Materias con análisis detallado por centros: %d', ...
    length(exceles))));

%% =====================================================
%% RESULTADOS ESPACIALES
%% =====================================================

append(doc,Heading(2,'8. Resultados espaciales'));

if isfile(mapFile)

    img = Image(mapFile);

    img.Width = '14cm';

    append(doc,img);

end

%% =====================================================
%% CONCLUSIONES AUTOMATICAS
%% =====================================================

append(doc,Heading(2,'9. Conclusiones'));

if ~isempty(mejorMateria)

    append(doc,...
        Paragraph(['La materia con mejor indicador ' ...
        'compuesto fue: ' ...
        char(mejorMateria)]));

end

if ~isempty(peorMateria)

    append(doc,...
        Paragraph(['La materia con menor indicador ' ...
        'compuesto fue: ' ...
        char(peorMateria)]));

end

if ~isempty(mejorCentro)

    append(doc,...
        Paragraph(['El centro con mejor rendimiento ' ...
        'global fue: ' ...
        char(mejorCentro)]));

end

if ~isempty(peorCentro)

    append(doc,...
        Paragraph(['El centro con menor rendimiento ' ...
        'global fue: ' ...
        char(peorCentro)]));

end

append(doc,...
    Paragraph(['Los resultados muestran diferencias ' ...
    'significativas entre materias y centros educativos.']));

%% =====================================================
%% INFORMACION PENDIENTE
%% =====================================================

append(doc,...
    Heading(2,...
    '10. Información pendiente'));

append(doc,...
    Paragraph('- Comparativa histórica'));

append(doc,...
    Paragraph('- Estadísticas de reclamaciones'));

append(doc,...
    Paragraph('- Comentarios de coordinación'));

%% =====================================================
%% GUARDAR
%% =====================================================

close(doc)

disp('========================================')
disp('INFORME GLOBAL EXTRAORDINARIA V2')
disp('========================================')
disp([nombreInforme '.docx'])