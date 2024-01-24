import pandas as pd
import numpy as np
import re
import os

def excel_to_outcomes_canvas(ruta_archivo, ruta_destino, nombre_archivo):

    ## leer excel de las consiganas de las actividades en excel
    df_rubric = pd.read_excel(ruta_archivo, sheet_name='Consignas')
    df_rubric = df_rubric[df_rubric['Tiene rubrica?'] == 'Si']

    ## establecer valores por defecto para la exportacion de la rubrica df_outcomes
    calculation_method = 'latest'
    calculation_int = ''
    workflow_state = 'active'
    mastery_points = 1

    ## seleccionar las cabeceras que usaremos desde el excel
    headers = list(df_rubric.columns)
    headers_rubric = []
    
    ## validar si existe algunas columnas de criterios y id rubrica
    exist_criteria = False
    exist_id_rubric = False

    for header in headers:
        if header.startswith("Criterio"):
            exist_criteria = True
        if "id_rubrica" in header:
            exist_id_rubric = True

    if exist_criteria and exist_id_rubric:
        for header in headers:
            headers_rubric.append(header)
    else:
        message = f"No se encontró las columnas de id_rubrica o criterios. Revisa el Excel de consignas y vuelva a intentarlo"
        return (message)
  

    df_rubric = df_rubric[headers_rubric]

    ## crear el df de outcomes para la importacion
    headers_outcomes = [
        'vendor_guid', 
        'object_type', 
        'title', 
        'display_name', 
        'calculation_method', 
        'calculation_int', 
        'workflow_state', 
        'parent_guids', 
        'mastery_points', 
        'ratings_4_pts', 
        'ratings_4', 
        'ratings_3_pts', 
        'ratings_3', 
        'ratings_2_pts', 
        'ratings_2', 
        'ratings_1_pts', 
        'ratings_1'
        ]

    df_outcomes = pd.DataFrame(columns=headers_outcomes)

    ## leer la cantidad de criterios que tiene el excel
    criterios = []
    for header in headers_rubric:
        if re.compile(r'^Criterio [0-9][0-9]$').search(header):
            criterios.append(header)

    ## recorrer el excel y establecer los parametros para construir el csv outcomes para la importacion
    for indice, fila in df_rubric.iterrows():

        cod_catalogo = fila['Nro catalogo'][-4:]
        cod_sis_eva = fila['Sistema evaluación']
        cod_periodo = fila['Periodo lanzamiento'][:4]+fila['Periodo lanzamiento'][5:6]

        ## validar si id_rubrica no es null
        if pd.notnull(fila['id_rubrica']):
            ## insertar las carpetas de outcomes group
            if 'id_rubrica' in fila:
                insert_group = {
                    # 'vendor_guid': fila['id_rubrica'],
                    'vendor_guid': cod_catalogo+'-'+cod_sis_eva+'-'+cod_periodo, 
                    'object_type': 'group', 
                    'title': fila['id_rubrica'], 
                    'display_name': fila['id_rubrica']
                    }
                df_temp_1 = pd.DataFrame([insert_group])
                df_outcomes = pd.concat([df_outcomes, df_temp_1])

            ## insertar los criterios
            contador = 1
            
            for criterio in criterios:
                try:
                    insert_outcome = {
                        'vendor_guid': cod_catalogo+'-'+cod_sis_eva+'-'+criterio[-2:]+'-'+cod_periodo, 
                        'object_type': 'outcome', 
                        'title': fila[criterio], 
                        'display_name': cod_sis_eva+'-'+criterio[-2:], 
                        'calculation_method': calculation_method,
                        'calculation_int': calculation_int,
                        'workflow_state': workflow_state,
                        'parent_guids': cod_catalogo+'-'+cod_sis_eva+'-'+cod_periodo,
                        'mastery_points': mastery_points,
                        'ratings_4_pts': fila[criterio+' Esperado pts'],
                        'ratings_4': fila[criterio+' Esperado'],
                        'ratings_3_pts': fila[criterio+' En Proceso 2 pts'],
                        'ratings_3': fila[criterio+' En Proceso 2'],
                        'ratings_2_pts': fila[criterio+' En Proceso 1 pts'],
                        'ratings_2': fila[criterio+' En Proceso 1'],
                        'ratings_1_pts': fila[criterio+' Inicial pts'],
                        'ratings_1': fila[criterio+' Inicial']
                        }
                    df_temp_2 = pd.DataFrame([insert_outcome])

                    # Eliminar columnas vacías o llenas de NA antes de la concatenación
                    df_outcomes = df_outcomes.dropna(axis=1, how='all')
                    df_temp_2 = df_temp_2.dropna(axis=1, how='all')

                    df_outcomes = pd.concat([df_outcomes, df_temp_2])

                    contador += 1

                except KeyError as e:
                    message = f"No se encontró las columnas para el'{criterio}'. \n Detalles: {e}. Revisa el Excel de consignas y vuelva a intentarlo"
                    return (message)
            
    ## ultimos filtros y modificaciones
    valid_outcomes = (df_outcomes.object_type == 'outcome') & (df_outcomes.title.isnull()) & (df_outcomes.ratings_4.isnull()) & (df_outcomes.ratings_4_pts.isnull()) & (df_outcomes.ratings_3.isnull()) & (df_outcomes.ratings_3_pts.isnull()) & (df_outcomes.ratings_2.isnull()) & (df_outcomes.ratings_2_pts.isnull()) & (df_outcomes.ratings_1.isnull()) & (df_outcomes.ratings_2_pts.isnull())
    df_outcomes = df_outcomes[~valid_outcomes]

    ## agregar nombre de los niveles en cada descripcion
    levels_outcomes = {
        'ratings_4': 'ESTÁNDAR ESPERADO: ',
        'ratings_3': 'EN PROCESO 2: ',
        'ratings_2': 'EN PROCESO 1: ',
        'ratings_1': 'INICIAL: '
    }

    for level in levels_outcomes:
        df_outcomes[level] = np.where((df_outcomes.object_type == 'outcome') & (df_outcomes[level].notna()), levels_outcomes[level] + df_outcomes[level], df_outcomes[level])


    df_outcomes = df_outcomes.apply(lambda row: pd.Series(row.dropna().values), axis=1)

    df_outcomes.columns = ['vendor_guid', 'object_type', 'title', 'display_name',
              'calculation_method', 'calculation_int', 'workflow_state',
              'parent_guids', 'mastery_points', 'ratings', '', '', '', '', '', '', '']

    try:
        df_outcomes.to_csv(ruta_destino+'/'+nombre_archivo+'_outcomes_canvas.csv', encoding='utf-8', index=False)
        print("Documentos exportados exitosamente - Outcomes para LMS Canvas.")
        return (nombre_archivo+'_outcomes_canvas.csv')
    except PermissionError:
        return ("El archivo que esta intentando exportar esta en uso. Cierra cualquier programa csv o excel.")
    except Exception as e:
        return (f"Ocurrió un error: {e}")