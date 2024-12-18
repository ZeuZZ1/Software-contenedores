# -*- coding: utf-8 -*-
"""
Created on Mon May  3 15:41:11 2021

@author: bpineda
"""

dic_campos_prioridades={'prioridad_prior':'PRIORIDAD',
                        'container_prior':'CONTAINER',
                        'nave_prior':'NAVE'
                        }

dic_campos_disponibilidad={'nro_contenedor':'NroContenedor',
                           'entrega_puerto':'Horario de retiro en puerto',
                           'recep_cliente':'Recepción Cliente'
                           }

dic_campos_planificacion={'invoice_col':'invoice',
                          'conocimiento_col':'Conocimiento',
                          'proveedor_col':'Proveedor',
                          'container_col':'Container',
                          'grupo_col':'Grupo',
                          'tipo_transp_col':'Tipo_transporte',
                          'tipo_contenedor_col':'Tipo',
                          'nave_col':'Nave',
                          'puerto_col':'Puerto',
                          'n_codigos_cont_col':'Cant Codigos',
                          'n_cajas_cont_col':'Cant cajas',
                          'n_productos_cont_col':'Cant productos',
                          'pallet_col':'Pallets',
                          'temporada_col':'Temporada',
                          'depto_col':'Departamentos',
                          'productos_col':'Productos',
                          'prioridad_col':'Prioridad',
                          'indicador_col':'Indicador',
                          'eta_col':'ETA',
                          'carrier_col':'Carrier',
                          'dias_libres_col':'Dias libres',
                          'fecha_lim_col':'Fecha lim',
                          'deadline_col':'Deadline',
                          'dia_dead_col':'Dia deadline',
                          'fecha_pm':'Fecha pm',
                          'disponibilidad_col':'Disponibilidad'
                          }

def carga_plano():
    '''Función que devuelve el plano de bodega y quita a los espacios a los valores
    que los tengan
    '''
    
    try:
        plano= pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\plano\\rpt_plano_bodega_compra.csv', sep=',')
        if plano.shape[1] == 1:
            plano= pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\plano\\rpt_plano_bodega_compra.csv', sep=';')
    except:
        plano= pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\plano\\rpt_plano_bodega_compra.csv', sep='\t')
        if plano.shape[1] == 1:
            plano= pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\plano\\rpt_plano_bodega_compra.csv', sep=';')    
    #Limpieza de plano
    for i in list(plano.columns):
        if type(plano[i].iloc[0]) is str:
            plano[i]=plano[i].apply(lambda x: str(x))
            plano[i]=plano[i].apply(lambda x: x.rstrip())
    del i
    asteriscos_plano(plano)
    for i in plano['numero_invoice'].unique():
        try:
            contenedores=plano[plano[dic_campos_plano['invoice_plano']]==i][dic_campos_plano['container_plano']].nunique()
            # contenedores=plano[plano['numero_invoice']==19084]['CONTAINER'].nunique()
            # contenedores=plano[plano['numero_invoice']==19283]['CONTAINER'].nunique()
            cajas=plano[plano[dic_campos_plano['invoice_plano']]==i][dic_campos_plano['cajas_invoice_plano']].unique()
            unidades=plano[plano[dic_campos_plano['invoice_plano']]==i][dic_campos_plano['unidades_plano']].sum()
            cm3=plano[plano[dic_campos_plano['invoice_plano']]==i][dic_campos_plano['cm3_plano']].unique()
            # cajas=plano[plano['numero_invoice']==19084]['PAQUE'].unique()
            # cajas=plano[plano['numero_invoice']==19283]['PAQUE'].unique()
            cajacont=math.ceil(cajas/contenedores)
            unidadcont=math.ceil(unidades/(contenedores**2))
            cm3cont=math.ceil(cm3/contenedores)
            ininvoice=list(plano[plano['numero_invoice']==i].index)
            # ininvoice=list(plano[plano['numero_invoice']==19283].index)
            for j in ininvoice:
                plano.at[j,dic_campos_plano['cajas_invoice_plano']]=cajacont
                plano.at[j,dic_campos_plano['unidades_plano']]=unidadcont
                plano.at[j,dic_campos_plano['cm3_plano']]=cm3cont
        except:
            continue
    return plano
# 18901 3662

def corregir_plano():
    '''Función que corrige los valores del plano de bodega de compra. 
    El usuario deberá indicar el campo al cuál se le desea modificar datos.
    Luego, el usuario deberá indicar el nuevo y antiguo valor a cambiar.
    La idea es que los cambios sean hechos antes de levantar la tabla para
    la coordinación del transporte.
    '''
    campo=str(input('¿Qué campo se requiere cambiar?: '))
    tipo_dato=input('¿Es un dato de tipo númerico o texto? Indique "n" para númerico y "t" para texto: ')
    if tipo_dato=='n':
        nuevo_valor=int(input('Ingrese nuevo valor numérico: '))
        antiguo_valor=int(input('Ingrese antiguo valor numérico: '))
    elif tipo_dato=='t':
        nuevo_valor=str(input('Ingrese nuevo valor de texto: '))
        antiguo_valor=str(input('Ingrese antiguo valor de texto: '))
    else:
        print('No es un valor permitido.')
        return 0
    plano=pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\plano\\rpt_plano_bodega_compra.csv', sep=',')
    index=plano[plano[campo]==antiguo_valor].index
    for i in index:
        plano.at[i,campo]=nuevo_valor
    plano.to_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\plano\\rpt_plano_bodega_compra.csv',index=False)
    return 1

def carga_archivo_disponibilidad():
    '''Función que sirve para cargar el archivo que indica si es que un archivo
    tiene o no disponibilidad para ser programado por el cd
    '''
    
    # Búsqueda de archivo más actual
    dia_hoy=datetime.date.today()
    semana_actual=dia_hoy.isocalendar()[1]
    # carpeta_base=os.listdir(r'C:\Users\bpineda\Desktop\CD\Programacion diaria de containers\reportes demurrages')
    # fechas=[]
    # for archivo in carpeta_base:
    #     nombre_archivo=archivo
    #     nombre_archivo=nombre_archivo.split('-')
    #     nombre_archivo=nombre_archivo[1].split()
    #     nombre_archivo=nombre_archivo[1].split('.')
    #     nombre_archivo=nombre_archivo[0]
    #     fechas.append(int(nombre_archivo))
    # if max(fechas)<semana_actual:
    carpeta_base=os.listdir(r'C:\Users\bpineda\Desktop\CD\Programacion diaria de containers\disponibilidad contenedores')
    fechas=[]
    for archivo in carpeta_base:
        nombre_archivo=archivo
        nombre_archivo=nombre_archivo.split('_')
        nombre_archivo=nombre_archivo[1].split('.')
        fecha=nombre_archivo[0]
        ano=int(fecha[0:4])
        mes=int(fecha[4:6])
        dia=int(fecha[6:])
        fecha=datetime.date(year=ano,month=mes,day=dia)
        fechas.append(fecha)
        actual=max(fechas)
        if len(str(actual.month))==1:
            if len(str(actual.day))==1:
                nombre_archivo=('All Data Details_'+str(actual.year)
                                +'0'+str(actual.month)+'0'+str(actual.day)
                                +'.csv')
            elif len(str(actual.day))!=1:
                nombre_archivo=('All Data Details_'+str(actual.year)
                                +'0'+str(actual.month)+str(actual.day)
                                +'.csv')
        if len(str(actual.month))!=1:
            if len(str(actual.day))==1:
                nombre_archivo=('All Data Details_'+str(actual.year)
                                +str(actual.month)+'0'+str(actual.day)+'.csv')
            elif len(str(actual.day))!=1:
                nombre_archivo=('All Data Details_'+str(actual.year)
                                +str(actual.month)+str(actual.day)+'.csv')
        
        # else:
        #     nombre_archivo=('All Data Details_'+str(actual.year)
        #                     +str(actual.month)+'0'+str(actual.day)+'.csv')
            
    reporte_demu=pd.read_csv(
        'C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\disponibilidad contenedores\\'
        +nombre_archivo)
    # reducción del archivo
    reporte_demu=reporte_demu[[dic_campos_disponibilidad['nro_contenedor'],
                               dic_campos_disponibilidad['entrega_puerto'],
                               dic_campos_disponibilidad['recep_cliente']]]
    
    reporte_demu=reporte_demu[~reporte_demu[
        dic_campos_disponibilidad['entrega_puerto']].isnull()]
    
    reporte_demu[dic_campos_disponibilidad['entrega_puerto']]=reporte_demu[
        dic_campos_disponibilidad['entrega_puerto']].apply(fecha_eta)
    
    reporte_demu=reporte_demu[[dic_campos_disponibilidad['nro_contenedor'],
                               dic_campos_disponibilidad['entrega_puerto']]]
    reporte_demu[dic_campos_disponibilidad['nro_contenedor']]=reporte_demu[dic_campos_disponibilidad['nro_contenedor']].apply(lambda x: x.strip())
# else:
#     nombre_archivo='Reporte Demurrage Tricot - Week '+str(max(fechas))+'.xlsx'
#     reporte_demu=pd.read_excel(r'C:\Users\bpineda\Desktop\CD\Programacion diaria de containers\reportes demurrages\\'+nombre_archivo,header=8)
#     reporte_demu=reporte_demu[['NroContenedor','Horario Entrega en Puerto','Horario Recepción Cliente']]
#     reporte_demu=reporte_demu[(~reporte_demu['Horario Entrega en Puerto'].isnull())]
#     for i in list(reporte_demu.columns):
#         reporte_demu[i]=reporte_demu[i].apply(lambda x: str(x))
#         reporte_demu[i]=reporte_demu[i].apply(lambda x: x.rstrip())
#     # reporte_demu['Horario Entrega en Puerto']=reporte_demu['Horario Entrega en Puerto'].apply(fecha_eta)
#     reporte_demu=reporte_demu[['NroContenedor','Horario Entrega en Puerto']]
    
    return reporte_demu
   
def disponibilidad(i):
    '''Función que entrega el mensaje de si está disponible o no un contenedor,
    tomando como base el archivo de disponibilidad
    '''
    
    if i in list(reporte_demu[~reporte_demu[dic_campos_disponibilidad[
            'nro_contenedor']].isin(maestro_prog[dic_campos_planificacion[
                'container_col']])][dic_campos_disponibilidad['nro_contenedor']]):
        if (reporte_demu[reporte_demu[dic_campos_disponibilidad['nro_contenedor']
                                     ]==i][dic_campos_disponibilidad['entrega_puerto']].unique()[0]<=datetime.date.today()):
            return 'Disponible'
        elif (reporte_demu[reporte_demu[dic_campos_disponibilidad['nro_contenedor']
                                     ]==i][dic_campos_disponibilidad['entrega_puerto']].unique()[0]>datetime.date.today()):
            return 'Con fecha'
    else:
        return 'Sin fecha'
    
def embarcador(bl, contenedor):
    '''Función que se encarga de diferenciar el embarcador del contenedor.
    El objetivo es separar los contenedores al momento de enviar los correos
    de programación'''
    
    embarcador= plano[(plano[dic_campos_plano['bl_plano']]==bl)&(plano[dic_campos_plano['container_plano']]==contenedor)][dic_campos_plano['embarcador_plano']].unique()[0]
    if embarcador in dic_campos_plano['embarcadores_plano']:
        if embarcador == 'DHL EMB':
            return 'DHL'
        elif embarcador == 'SPARX T':
            return 'SPARX'
        elif embarcador == 'NWPT':
            return 'NOWPORTS'
        elif embarcador =='JAS':
            return 'JAS'
    else:
        return 'Embarcador no reconocido'
        
def carga_prior():
    '''Función para cargar el archivo de prioridades comerciales. Es necesario
    que cumpla con el formato establecido para que este sistema pueda leerlo
    '''
    
    direccion_prioridades=r'C:\Users\bpineda\Desktop\CD\Programacion diaria de containers\prioridades pedidos\archivo prioridades_actual'
    archivo= os.listdir(direccion_prioridades)
    priors=os.listdir(r'C:\Users\bpineda\Desktop\CD\Programacion diaria de containers\prioridades pedidos\archivo prioridades_actual')
    fechas=[]
    for archivo in priors:
        if 'Prioridad' in archivo:
            archivo=archivo.split()
            fecha=archivo[0]
            ano=int(fecha[0:4])
            mes=int(fecha[4:6])
            dia=int(fecha[6:])
            fecha=datetime.date(year=ano,month=mes,day=dia)
            fechas.append(fecha)
    actual=max(fechas)
    if len(str(actual.month))==1:
        # dd==1
        if len(str(actual.day))==1:
            nombre_prior=(str(actual.year)+'0'+str(actual.month)+'0'
            +str(actual.day)+' Prioridad comercial por contenedor.xlsx')
        # dd!=1
        elif len(str(actual.day))!=1:
            nombre_prior=(str(actual.year)+'0'+str(actual.month)
                          +str(actual.day)+' Prioridad comercial por contenedor.xlsx')
    if len(str(actual.month))!=1:
        # dd==1
        if len(str(actual.day))==1:
            nombre_prior=(str(actual.year)+str(actual.month)+'0'
                          +str(actual.day)+' Prioridad comercial por contenedor.xlsx')
        #  dd!=1
        elif len(str(actual.day))!=1:
            nombre_prior=(str(actual.year)+str(actual.month)
                          +str(actual.day)+' Prioridad comercial por contenedor.xlsx')
            
    prior= pd.read_excel(direccion_prioridades+'\\'
                          +nombre_prior,sheet_name='Prioridad contenedor', header=0)
    prior=prior[[dic_campos_prioridades['prioridad_prior'],
                 dic_campos_prioridades['container_prior']]]
    prior.drop_duplicates(keep='first', inplace=True)
    archivo_prior=nombre_prior.split(' ')
    archivo_prior=archivo_prior[0]
    return prior, archivo_prior

def prioridad(i, prior, archivo_prior):
    '''Función que entrega una la prioridad de un contenedor de acuerdo al
    archivo de prioridades definido por el área comercial más actualizado que haya
    '''
    
    if i in list(prior[dic_campos_prioridades['container_prior']]):
        rank=str(prior[prior[dic_campos_prioridades['container_prior']]==i][
            dic_campos_prioridades['prioridad_prior']].unique()[0])
        return rank+'/'+archivo_prior
    else:
        return 'No prioritario'

def cont_sin_detalle(plano, maestro):
    '''Eliminación del valor "SIN DETALLE" en el plano de bodega. Esto gracias
    a que en el archivo de maestro de manifiestos si tiene los nombres de los
    contenedores necesarios
    '''
    
    s_index=list(plano[plano[dic_campos_plano[
        'container_plano']]=='SIN DETALLE'].index)
    aux= plano.loc[s_index]
    bl= list(aux[dic_campos_plano['bl_plano']].unique())
    for i in bl:
        #container= plano[plano['CONOCIMIENTO']==i]['CONTAINER']
        container=aux[aux[dic_campos_plano['bl_plano']]==i][dic_campos_plano[
            'container_plano']]
        
        for j in container:
            print(i, j)
            if j == 'SIN DETALLE':
                print(j)
                fil_indice= plano.loc[s_index[0]]
                producto= fil_indice[dic_campos_plano['producto_plano']]
                nom_cont= maestro[(maestro[dic_campos_manifiesto['bl_col']]==i)
                                  & (maestro[dic_campos_manifiesto['item_col']]
                                     ==str(producto))][
                                         dic_campos_manifiesto['container_col']]
                print(nom_cont)
                print(str(producto))
                try:
                    plano.at[s_index[0],dic_campos_plano['container_plano']]= str(
                        list(nom_cont)[0])
                    s_index.pop(0)
                except:
                    s_index.pop(0)
                    continue 

def puerto(puerto_container):
    '''Función enfocada en índices de una serie que devuelve el nombre de un
    puerto dado su código'''
    
    puertos= {'CLSAI': 'San Antonio',
              'CLVAP': 'Valparaiso',
              'CLSVE':'San Vicente',
              'SAN ANTONIO, CL':'San Antonio',
              'VALPARAISO, CL': 'Valparaiso',
              'VALPARAÍSO, CL':'Valparaiso',
              'VALPARAÍSO,CL':'Valparaiso',
              'LIRQUÉN, CL':'Lirquen',
              'LIRQUEN, CL': 'Lirquen',
              'SAN VICENTE, CL': 'San Vicente'}
    
    return puertos[puerto_container]

def grupo_contenedor(i,invoices):
    '''Función que devuelve el grupo de contenedores que comparten un conjunto
    de invoices. estas quedan separadas por guión.
    '''
    
    cont_inv=dict()
    aux=[]
    for j in invoices:
        for k in plano[plano[dic_campos_plano['invoice_plano']]==j][
                dic_campos_plano['container_plano']].unique():
            aux.append(k)
    cont_inv[i]=set(aux)
    cont_inv[i]=list(cont_inv[i])
    cont_inv[i].sort()
    if len(cont_inv[i])==1:
        if cont_inv[i][0]==i:
            cont_inv[i]=['Contenedor','unico']
            return ' '.join(cont_inv[i])
    else:
        return '-'.join(cont_inv[i])

def pallet(i):
    '''Función que devuelve la cantidad de pallets que representa cada
    container dependiendo del tipo de container
    '''
    
    #dimensiones[m] pallet largo:1.2 ,ancho: 1, alto:1.6
    #dimensiones contenedores:
        #20 GP: 5.9, 2.34, 2.39
        #40 HC: 12.01, 2.352, 2.69
        #40 NOR: 11.48, 2.26, 2.18
        #40 GP: 12.01, 2.35, 2.39
    pallet={'20GP':18,'40HC':41,'40NOR':31,'40GP':37}
    tipo=maestro[maestro['Line Container']==i]['Type'].unique()[0]
    return pallet[tipo]

# Función para poder generar la información acerca de los contenedores 
# Tratados con otro fwd.
def kkk():
    global fwd_externo
    global nuevo_df
    fwd_externo.apply(fwd_ext,axis=1)
    nuevo_df.reset_index(inplace=True, drop=True)
    try:
        nuevo_df=nuevo_df.groupby(['invoice','Container','Grupo','Embarcador','Nave','Indicador','ETA',
                      'Dias libres','Fecha lim','Deadline','Temporada',
                      'Departamentos','Productos','Prioridad','Congelado','CD destino','VAS','Fecha pm','Fecha disp',
                      'Disponibilidad']).sum()
        nuevo_df.reset_index(inplace=True)
    except:
        print('No quedan contenedores gestionados por fwdrs diferentes al oficial')
    return nuevo_df

# 
def fwd_ext(i):
    global nuevo_df
    global ultimo_dato
    global fwd_externo
    try:
        print(fwd_externo.at[i.name,dic_campos_fwdext['container_col']])
        invoices=list(plano[plano[dic_campos_plano['container_plano']]==fwd_externo.at[i.name,dic_campos_fwdext['container_col']]]
                                  [dic_campos_plano['invoice_plano']].unique())
        lineas=list(plano[plano[dic_campos_plano['container_plano']]==fwd_externo.at[i.name,dic_campos_fwdext['container_col']]]
                    [dic_campos_plano['linea_plano']].unique())
        temporadas=list(plano[plano[dic_campos_plano['container_plano']]==fwd_externo.at[i.name,dic_campos_fwdext['container_col']]]
                        [dic_campos_plano['temporada_plano']].unique())
        departamentos=list(plano[plano[dic_campos_plano['container_plano']]==fwd_externo.at[i.name,dic_campos_fwdext['container_col']]]
                           [dic_campos_plano['departamento_plano']].unique())
        x=pd.DataFrame(data={'invoice':['-'.join(list(map(lambda x: str(x),invoices)))],
                             # 'Conocimiento':[fwd_externo.at[i.name,dic_campos_fwdext['bl_col']]],
                             'Container':[fwd_externo.at[i.name,dic_campos_fwdext['container_col']]],
                             'Grupo':[grupo_contenedor(fwd_externo.at[i.name,dic_campos_fwdext['container_col']],invoices)],
                             'Embarcador':[embarcador(fwd_externo.at[i.name,dic_campos_fwdext['bl_col']], fwd_externo.at[i.name,dic_campos_fwdext['container_col']])],
                             # 'Embarcador':[embarcador(fwd_externo.at[167,dic_campos_fwdext['bl_col']], fwd_externo.at[167,dic_campos_fwdext['container_col']])],
                             'Nave':[plano[(plano[dic_campos_plano['bl_plano']]==
                                           fwd_externo.at[i.name,dic_campos_fwdext[
                                               'bl_col']])&(plano[
                                                   dic_campos_plano['container_plano']]
                                                   ==fwd_externo.at[i.name,
                                                                    dic_campos_fwdext[
                                                                        'container_col']])]
                                                   [dic_campos_plano['nave_plano']].unique()[0]],
                             # 'Cant cajas': [fwd_externo.at[i.name,dic_campos_fwdext['cajas_col']]],
                             'Cant cajas': [int((plano[(plano[dic_campos_plano['bl_plano']]==
                                           fwd_externo.at[i.name,dic_campos_fwdext[
                                               'bl_col']])&(plano[
                                                   dic_campos_plano['container_plano']]
                                                   ==fwd_externo.at[i.name,
                                                                    dic_campos_fwdext[
                                                                        'container_col']])][[
                                                                            dic_campos_plano['bl_plano'],
                                                                            dic_campos_plano['cajas_invoice_plano']]].drop_duplicates(keep='first'))[dic_campos_plano['cajas_invoice_plano']].sum())],
                             'Cant productos':[int((plano[(plano[dic_campos_plano['bl_plano']]==
                                           fwd_externo.at[i.name,dic_campos_fwdext[
                                               'bl_col']])&(plano[
                                                   dic_campos_plano['container_plano']]
                                                   ==fwd_externo.at[i.name,
                                                                    dic_campos_fwdext[
                                                                        'container_col']])][[
                                                                            dic_campos_plano['bl_plano'],
                                                                            dic_campos_plano['unidades_plano']]].drop_duplicates(keep='first'))[dic_campos_plano['unidades_plano']].sum())],
                             'Indicador':[round(plano[plano[dic_campos_plano['container_plano']]==fwd_externo.at[i.name,dic_campos_fwdext['container_col']]][dic_campos_plano['cm3_plano']].sum()/plano[plano[dic_campos_plano['container_plano']]==fwd_externo.at[i.name,dic_campos_fwdext['container_col']]][dic_campos_plano['cajas_invoice_plano']].sum(),3)],
                             'ETA':[plano[(plano[dic_campos_plano['bl_plano']]==fwd_externo.at[i.name,dic_campos_fwdext['bl_col']])&(plano[dic_campos_plano['container_plano']]==fwd_externo.at[i.name,dic_campos_fwdext['container_col']])][dic_campos_plano['eta_plano']].unique()[0]-datetime.timedelta(days=2)],
                             'Dias libres':[fwd_externo.at[i.name,dic_campos_fwdext['dias_col']]],
                             'Fecha lim':[(plano[(plano[dic_campos_plano['bl_plano']]==fwd_externo.at[i.name,dic_campos_fwdext['bl_col']])&(plano[dic_campos_plano['container_plano']]==fwd_externo.at[i.name,dic_campos_fwdext['container_col']])][dic_campos_plano['eta_plano']].unique()[0]-datetime.timedelta(days=2))+(datetime.timedelta(days=int(fwd_externo.at[i.name,dic_campos_fwdext['dias_col']]))-datetime.timedelta(days=4))],
                             'Deadline':[(plano[(plano[dic_campos_plano['bl_plano']]==fwd_externo.at[i.name,dic_campos_fwdext['bl_col']])&(plano[dic_campos_plano['container_plano']]==fwd_externo.at[i.name,dic_campos_fwdext['container_col']])][dic_campos_plano['eta_plano']].unique()[0]-datetime.timedelta(days=2))+(datetime.timedelta(days=int(fwd_externo.at[i.name,dic_campos_fwdext['dias_col']]))-datetime.timedelta(days=1))],
                             'Temporada':['-'.join(list(map(lambda x: str(x),temporadas)))],
                             'Departamentos':['-'.join(list(map(lambda x: str(x),departamentos)))],
                             'Productos':['-'.join(list(map(lambda x: str(x),lineas)))],
                             'Prioridad':[prioridad(fwd_externo.at[i.name,dic_campos_fwdext['container_col']], prior, archivo_prior)],
                             'Congelado':[congelado(fwd_externo.at[i.name,dic_campos_fwdext['container_col']])],
                             'CD destino':[pulmon(fwd_externo.at[i.name,dic_campos_fwdext['container_col']])],
                             'VAS':[vas(fwd_externo.at[i.name,dic_campos_fwdext['container_col']])],
                             'Fecha pm':[(plano[(plano[dic_campos_plano['bl_plano']]==fwd_externo.at[i.name,dic_campos_fwdext['bl_col']])&(plano[dic_campos_plano['container_plano']]==fwd_externo.at[i.name,dic_campos_fwdext['container_col']])][dic_campos_plano['eta_plano']].unique()[0]-datetime.timedelta(days=2))+datetime.timedelta(days=ultimo_dato)],
                             'Fecha disp':[(plano[(plano[dic_campos_plano['bl_plano']]==fwd_externo.at[i.name,dic_campos_fwdext['bl_col']])&(plano[dic_campos_plano['container_plano']]==fwd_externo.at[i.name,dic_campos_fwdext['container_col']])][dic_campos_plano['eta_plano']].unique()[0]-datetime.timedelta(days=2))+datetime.timedelta(days=dias)],
                             'Disponibilidad':[disponibilidad(fwd_externo.at[i.name,dic_campos_fwdext['container_col']])]
                             }
                       )
        nuevo_df=pd.concat([nuevo_df,x])
    except:
        print('error con el contenedor')
        
        
def maestro_por_contenedor():
    '''Función que devuelve la tabla de planificación para la recepción de con
    tenedores en alguno de los CDs de Triot- Esta tabla se alimenta de los
    archivos de: prioridades comerciales, plano de bodega de compra, manifiestos
    de la empresa de transporte y el archivo de disponibiildad del contenedor
    de acuerdo a la información suministrada por la plataforma de la empresa
    de transporte.
    '''
    global nuevo_df
    #plano_aux= plano[plano['FECHA_APROX'] <= datetime.date(year=maestro['ETA'].max().year, month=maestro['ETA'].max().month, day=maestro['ETA'].max().day)+datetime.timedelta(days=2)]
    
    error=[]
    plano_aux=plano[plano[dic_campos_plano['bl_plano']].isin(maestro[dic_campos_manifiesto['bl_col']])]    
    mani_por_contenedor=pd.DataFrame()
    bl=list(plano_aux[dic_campos_plano['bl_plano']].unique())
    # Incluir bl que posiblemente no están en el maestro pero si en el plano de bodega
    bl_2=list(plano[plano[dic_campos_plano['container_plano']].isin(maestro[maestro[dic_campos_manifiesto['bl_col']].isin(pd.Series(bl))][dic_campos_manifiesto['container_col']].drop_duplicates(keep='first'))][dic_campos_plano['bl_plano']].unique())
    bl=bl+bl_2
    bl=set(bl)
    bl=list(bl)
    plano_aux=plano[plano[dic_campos_plano['bl_plano']].isin(pd.Series(bl))]    
    del bl_2
    for _ in bl:
        contenedores= list(plano_aux[plano_aux[dic_campos_plano['bl_plano']]==_][dic_campos_plano['container_plano']].unique())
        
        for i in contenedores:
            try:
                #print(_,i)
                print(i)
                print(tipo_transporte(i))
                invoices=list(plano_aux[plano_aux[dic_campos_plano['container_plano']]==i]
                              [dic_campos_plano['invoice_plano']].unique())

                lineas=list(plano_aux[plano_aux[dic_campos_plano['container_plano']]==i]
                            [dic_campos_plano['linea_plano']].unique())

                temporadas=list(plano_aux[plano_aux[dic_campos_plano['container_plano']]==i]
                                [dic_campos_plano['temporada_plano']].unique())

                departamentos=list(plano_aux[plano_aux[dic_campos_plano['container_plano']]==i]
                                   [dic_campos_plano['departamento_plano']].unique())


                x= pd.DataFrame(data={'invoice':['-'.join(list(map(lambda x: str(x),invoices)))],
                                      'Conocimiento':[_],
                                      'Proveedor':[proveedor(i)],
                                      'Container':[i],
                                      'Grupo':[grupo_contenedor(i,invoices)],
                                      'Tipo_transporte':[tipo_transporte(i)],
                                      'Embarcador':[embarcador(_,i)],
                                      'Tipo':[maestro[maestro['Line Container']==i][dic_campos_manifiesto['tipo_contenedor_col']].unique()[0]],
                                      'Nave':[plano[(plano[dic_campos_plano['bl_plano']]==_) & (plano[dic_campos_plano['container_plano']]==i)][dic_campos_plano['nave_plano']].unique()[0]],
                                      'Puerto':[puerto(maestro[(maestro[dic_campos_manifiesto['bl_col']]==_) & (maestro['Line Container']==i)][dic_campos_manifiesto['puerto_col']].unique()[0])],
                                      'Cant Codigos':[int(maestro[(maestro[dic_campos_manifiesto['bl_col']]==_) & (maestro['Line Container']==i)][dic_campos_manifiesto['item_col']].nunique())],
                                      'Cant cajas':[int((plano[(plano[dic_campos_plano['bl_plano']]==_)&(plano[dic_campos_plano['container_plano']]==i)][[dic_campos_plano['bl_plano'],dic_campos_plano['cajas_invoice_plano']]].drop_duplicates(keep='first'))[dic_campos_plano['cajas_invoice_plano']].sum())],
                                      # 'Cant cajas':[int(maestro[(maestro[dic_campos_manifiesto['bl_col']]==_) & (maestro['Line Container']==i)][dic_campos_manifiesto['cantidad_cajas_col']].sum())],
                                      'Cant productos':[int((plano[(plano[dic_campos_plano['bl_plano']]==_)&(plano[dic_campos_plano['container_plano']]==i)][[dic_campos_plano['bl_plano'],dic_campos_plano['unidades_plano']]].drop_duplicates(keep='first'))[dic_campos_plano['unidades_plano']].sum())],
                                      # 'Cant productos':[int(maestro[(maestro[dic_campos_manifiesto['bl_col']]==_) & (maestro['Line Container']==i)][dic_campos_manifiesto['cantidad_unidades_col']].sum())],
                                      'Pallets':[pallet(i)],
                                      'Temporada':['-'.join(list(map(lambda x: str(x),temporadas)))],
                                      'Departamentos':['-'.join(list(map(lambda x: str(x),departamentos)))],
                                      'Productos':['-'.join(list(map(lambda x: str(x),lineas)))],
                                      'Prioridad':[prioridad(i, prior, archivo_prior)],
                                      'Congelado':[congelado(i)],
                                      'CD destino':[pulmon(i)],
                                      'VAS':[vas(i)],
                                      'Indicador':[round(plano[plano[dic_campos_plano['container_plano']]==i][dic_campos_plano['cm3_plano']].sum()/plano[plano[dic_campos_plano['container_plano']]==i][dic_campos_plano['cajas_invoice_plano']].sum(),3)],
                                      'ETA':[maestro[(maestro[dic_campos_manifiesto['bl_col']]==_) & (maestro[dic_campos_manifiesto['container_col']]==i)][dic_campos_manifiesto['eta_col']].unique()[0]],
                                      'Carrier':[maestro[(maestro[dic_campos_manifiesto['bl_col']]==_) & (maestro[dic_campos_manifiesto['container_col']]==i)][dic_campos_manifiesto['carrier_col']].unique()[0]],
                                      'Dias libres':[maestro[(maestro[dic_campos_manifiesto['bl_col']] == _) & (maestro[dic_campos_manifiesto['container_col']]== i)]['Demu free'].unique()[0]],
                                      'Fecha lim':[maestro[(maestro[dic_campos_manifiesto['bl_col']] == _) & (maestro[dic_campos_manifiesto['container_col']]== i)]['Fecha lim'].unique()[0]],
                                      'Deadline':[maestro[(maestro[dic_campos_manifiesto['bl_col']] == _) & (maestro[dic_campos_manifiesto['container_col']]== i)]['Deadline'].unique()[0]],
                                      'Dia deadline':[maestro[(maestro[dic_campos_manifiesto['bl_col']] == _) & (maestro[dic_campos_manifiesto['container_col']]== i)]['Dia deadline'].unique()[0]],
                                      'Fecha pm':[maestro[(maestro[dic_campos_manifiesto['bl_col']] == _) & (maestro[dic_campos_manifiesto['container_col']]== i)]['Fecha pm'].unique()[0]],
                                      'Fecha disp':[maestro[(maestro[dic_campos_manifiesto['bl_col']] == _) & (maestro[dic_campos_manifiesto['container_col']]== i)]['Fecha disp'].unique()[0]],
                                      'Disponibilidad':[disponibilidad(i)]
                                      #'Hora':[np.nan],
                                      #'Fecha':[np.nan]
                                      })
                # if i == 'SZLU9468778':
                    # aux=x
                mani_por_contenedor= pd.concat([mani_por_contenedor,x])
                mani_por_contenedor.reset_index(inplace=True)
                mani_por_contenedor.drop('index',axis=1, inplace=True)
            except:
                tupla=(_,i)
                error.append(tupla)
                print(i,'a')
                continue
    #mani_por_contenedor.to_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\maestro_por_contenedores.csv',index=False)
    #mani_por_contenedor= pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\maestro_por_contenedores.csv')
    
    sd= mani_por_contenedor.groupby(['invoice',
                                     'Proveedor',
                                     'Container',
                                     'Grupo',
                                     'Tipo_transporte',
                                     'Embarcador',
                                     'Tipo',
                                     'Nave',
                                     'Puerto',
                                     'ETA',
                                     'Dias libres',
                                     'Fecha lim',
                                     'Deadline',
                                     'Dia deadline',
                                     'Fecha pm',
                                     'Fecha disp',
                                     'Temporada',
                                     'Departamentos',
                                     'Productos',
                                     'Indicador',
                                     'Prioridad',
                                     'Congelado',
                                     'CD destino',
                                     'VAS',
                                     'Pallets',
                                     'Disponibilidad'
                                     #'Hora',
                                     #'Fecha'
                                     ]).sum()
    relleno= pd.Series(np.full(shape=len(mani_por_contenedor['invoice']),fill_value='-'))
    sd.insert(loc=0, column='CD', value= relleno)
    sd.insert(loc=len(sd.columns), column='Hora', value=relleno)
    sd.insert(loc=len(sd.columns), column='Fecha', value=relleno)
    del relleno
# sd.to_csv('C:\\Users\\bpineda\\Desktop\\pruebas\\maestro_por_contenedores_gby4.csv', index=True)
# =============================================================================
#     Proceso actualización del maestro de contenedores.
#     Contempla la revisión del archivo que se va a actualizar y el archivo de
#     maestro de programaciones para evitar la redundancia de datos entre archivos
# =============================================================================
    sd.to_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\backup temporal\\temp\\maestro_por_contenedores_gby.csv')
    del sd
    sd= pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\backup temporal\\temp\\maestro_por_contenedores_gby.csv')
    # nuevo_df=pd.DataFrame()
    kkk()
    # nuevo_df.drop('Conocimiento',axis=1,inplace=True)
    sd=pd.concat([sd,nuevo_df])
    sd.reset_index(inplace=True,drop=True)
# Proceso de verificación de contenedores que aún no han sido planificados
# dentro del archivo a actualizar y que no aparecieron en el archivo nuevo.
    if 'maestro_por_contenedores_gby.csv' not in os.listdir('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos'):        
        maestro_conte= pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\backup temporal\\temp\\maestro_por_contenedores_gby.csv', sep=',')
    else:
        maestro_conte= pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\maestro_por_contenedores_gby.csv', sep=',')
        if maestro_conte.shape[1]==1:
            maestro_conte= pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\maestro_por_contenedores_gby.csv', sep=';')
    maestro_conte= maestro_conte[maestro_conte['Hora'].isna()]
    maestro_conte= maestro_conte[~maestro_conte['Container'].isin(sd['Container'])]
    if maestro_conte.size == 0:
        del maestro_conte
        maestro_programaciones()
        # sd.to_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\maestro_por_contenedores_gby.csv', index=False)
    else:
        sd= pd.concat([sd,maestro_conte])
        sd.reset_index(inplace= True)
        sd.drop('index',axis=1, inplace=True)
        maestro_programaciones()
# Proceso de eliminación de contenedores que aparecen en el archivo actualizado
# pero que ya fueron programados.
    maestro_program= pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\programaciones\\maestro_programaciones.csv')
    sd= sd[~sd['Container'].isin(maestro_program['Container'])]
    sd.to_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\maestro_por_contenedores_gby.csv', index=False)
    sd.to_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\backup temporal\\maestro_por_contenedores_gby.csv', index=False)
    f=dict(error)
    e=pd.DataFrame(data=f, index=[0])
    e.to_excel(r'C:\Users\bpineda\Desktop\CD\Programacion diaria de containers\Manifiestos\error\lista_error.xlsx')
    del nuevo_df
    del sd
    return error


def maestro_programaciones():
    '''Función que crea o alimenta al maestro de programaciones del CD de
    acuerdo a la tabla de planificaciones.
    '''
    
    # df de maestro programaciones
    if 'maestro_programaciones.csv' not in os.listdir('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\programaciones'):
        maestro_program= pd.DataFrame()
    else:
        maestro_program= pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\programaciones\\maestro_programaciones.csv', sep=',')
        if maestro_program.shape[1]==1:
            maestro_program= pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\programaciones\\maestro_programaciones.csv', sep=';')
    # df maestro por contenedor
    if 'maestro_por_contenedores_gby.csv' not in os.listdir('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos'):
        maestro=pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\backup temporal\\temp\\maestro_por_contenedores_gby.csv', sep=',')
    else:
        maestro= pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\maestro_por_contenedores_gby.csv', sep=',')
        if maestro.shape[1]==1:
            maestro= pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\maestro_por_contenedores_gby.csv', sep=';')
    #concatenando tablas
    maestro= maestro[~(maestro['Hora'].isna()) & ~(maestro['Fecha'].isna())][[k for k in maestro.columns if k not in ['Pallets','Disponibilidad','Fecha pm','Congelado']]]
    maestro= maestro[~maestro['Container'].isin(maestro_program['Container'])]
    concat_maest= pd.concat([maestro_program,maestro])
    concat_maest.reset_index(inplace= True)
    concat_maest.drop('index',axis=1, inplace=True)
    concat_maest.sort_values(by='Fecha', axis=0, inplace=True)
    concat_maest.to_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\programaciones\\maestro_programaciones.csv', index= False)
    concat_maest.to_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\backup temporal\\maestro_programaciones.csv', index= False)
# =============================================================================
#     Ejecución de las funciones
# =============================================================================
plano=carga_plano()
cont_sin_detalle(plano, maestro)
prior, archivo_prior= carga_prior()


error=maestro_por_contenedor()


asteriscos_plano(plano)

corregir_plano()    


fecha_lcl_plano(maestro,plano)
maestro_programaciones()
pi=3.99
int(round(pi,0))
cont_plano='SZLU9468778'

plano[plano['CONTAINER']==cont_plano]['CONOCIMIENTO']
maestro[maestro['Line Container']=='SZLU9468778']['House Bill']
maestro[maestro['House Bill']==(plano[plano['CONTAINER']==cont_plano]['CONOCIMIENTO'].unique()[0])]['Line Container']

sd.columns


        
