# -*- coding: utf-8 -*-
"""
Created on Mon Jul 19 15:44:20 2021
Queda pendiente como se rebajará la venta en el sistema RFID.
@author: bpineda
"""

from scipy import stats
import seaborn as sns
from sklearn.linear_model import LinearRegression
from matplotlib.ticker import MaxNLocator
#codigo para generar las programaciones diarias
    ''' Con este código se obtiene la lista/archivo de contenedores de una programación
    una vez ingresada la fecha de consulta. Se obtiene la lista tanto para
    la administradora de ASN, como la que se envía para programación
    '''

def generar_documento_asn_prog():
    fecha=str(input('Introduzca la fecha en el formato dd/mm/aaaa: '))
    maestro_contenedor= pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\maestro_por_contenedores_gby.csv',sep=';')
    '''Programacion CD'''
    programacion_cd=maestro_contenedor[(maestro_contenedor['Fecha']==fecha)][['invoice','Container','Grupo','Tipo_transporte','Tipo','Nave','ETA','Deadline','CD','Hora','Fecha','Temporada','Departamentos','Productos','Cant cajas','Cant productos','Cant predist','Embarcador']]
    programacion_cd.sort_values('Hora',axis=0,inplace=True)
    programacion_cd.reset_index(inplace=True, drop=True)
    programacion_cd_group=programacion_cd.copy()
    programacion_cd_group['Container']=1
    programacion_cd_group['Productos']= str(fecha)
    programacion_cd_group.rename(columns={'Productos':'Totales'}, inplace=True)
    programacion_cd_group=programacion_cd_group[['Totales','Container', 'Cant cajas', 'Cant productos', 'Cant predist']]
    programacion_cd_group=programacion_cd_group.groupby(['Totales']).sum()
    programacion_cd_group.reset_index(inplace=True, drop=False)
    programacion_cd=pd.concat([programacion_cd,programacion_cd_group],axis=1)
    '''Programacion Embarcadores'''
    programacion_a=maestro_contenedor[(maestro_contenedor['Fecha']==fecha)&(maestro_contenedor['Embarcador']=='DHL')][['invoice','Container','Tipo_transporte','Tipo','Nave','Puerto','ETA','Deadline','CD','Hora','Fecha']]
    programacion_b=maestro_contenedor[(maestro_contenedor['Fecha']==fecha)&((maestro_contenedor['Embarcador']=='SPARX')|(maestro_contenedor['Embarcador']=='Embarcador no reconocido')|(maestro_contenedor['Embarcador']=='NOWPORTS'))][['invoice','Container','Tipo_transporte','Tipo','Nave','Puerto','ETA','Deadline','CD','Hora','Fecha']]
    programacion_c=maestro_contenedor[(maestro_contenedor['Fecha']==fecha)&(maestro_contenedor['Embarcador']=='JAS')][['invoice','Container','Tipo_transporte','Tipo','Nave','Puerto','ETA','Deadline','CD','Hora','Fecha']]
    programacion_a.sort_values('Hora',axis=0,inplace=True)
    programacion_b.sort_values('Hora',axis=0,inplace=True)
    programacion_c.sort_values('Hora',axis=0,inplace=True)
    # programacion1.to_excel('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\archivos_asn_programacion\\programacion1.xlsx',index=False)
    #codigo para asn
    '''Info ASN'''
    asn= maestro_contenedor[maestro_contenedor['Fecha']==fecha][['invoice','Container','Grupo','CD','Cant cajas','Cant productos','Hora','Fecha','Embarcador']]
    asn.sort_values('Hora',axis=0,inplace=True)
    # asn.to_excel('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\archivos_asn_programacion\\asn1.xlsx',index=False)
    # Guardado del archivo excel con la información para la programación
    dia=fecha.split('/')[0]
    mes=fecha.split('/')[1]
    
    nombre_archivo='ASN_Prog_Sens_'+dia+mes+'.xlsx'
    '''Info sensado'''
    sens=maestro_contenedor[(maestro_contenedor['Fecha']==fecha)&(maestro_contenedor['CD']=='Pedro Montt')][['invoice','Container','Grupo','Tipo','Nave','CD','Cant Codigos','Indicador','Cant cajas','Cant productos','Hora','Fecha']]
    sens['pond']=sens['Cant cajas']/sens['Cant cajas'].sum()
    dot_product= np.dot(sens['Indicador'],sens['pond'])
    if dot_product<=(0.6*0.4*0.4):
        estimacion_induccion='Inducibles'
    else: estimacion_induccion='Poco/No inducibles'
        
    # nombre_archivo='Sensado_Prog_'+dia+mes+'.xlsx'
    capacidad_sensado=50000
    cajas_total=math.ceil(sens['Cant cajas'].sum())
    productos_total=math.ceil(sens['Cant productos'].sum())
    estimado_cajas_pre=math.ceil((sens['Cant cajas'].sum())*0.9*0.5)
    estimado_productos_pre=math.ceil((sens['Cant productos'].sum())*0.9*0.5)
    contenedores=sens['Container'].nunique()
    if estimado_productos_pre >= capacidad_sensado:
        sensado_abast=0 
    else: sensado_abast=capacidad_sensado-estimado_productos_pre
    leyenda_capacidad= 'Capacidad de sensado considerada [unidades de sku por día]: '
    leyenda_cajas_total='Cajas a recibr en total [cajas]: '
    leyenda_productos_total='Unidades de sku a recibir en total [unidades de sku]: '
    leyenda_cajas='Cajas a recibr aprox a sensado [cajas]: '
    leyenda_productos='Unidades de sku a recibr aprox a sensado [unidades de sku]: '
    leyenda_sensado='Se necesita abastecer a sensado [unidades de sku]: '
    leyenda_induccion='Se espera recibir cajas [Inducibles, Poco/No inducibles]: '
    columna=['Cantidad/Información']
    cuadro_informacion={leyenda_capacidad:capacidad_sensado,
                        leyenda_cajas_total:cajas_total,
                        leyenda_productos_total:productos_total,
                        leyenda_induccion:estimacion_induccion}
    info_sensado_df=pd.DataFrame.from_dict(cuadro_informacion,orient='index',columns=columna)
    info_sensado_df.reset_index(inplace=True)
    info_sensado_df.rename(columns={'index':'Item'}, inplace=True)
    with pd.ExcelWriter('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\archivos_asn_programacion\\ASN\\'+nombre_archivo) as writer:
        asn.to_excel(writer,sheet_name='ASN',index=False)
        programacion_a.to_excel(writer,sheet_name='Programación_DHL',index=False)
        programacion_b.to_excel(writer,sheet_name='Programación_SPARX',index=False)
        programacion_c.to_excel(writer,sheet_name='Programación_JAS',index=False)
        programacion_cd.to_excel(writer,sheet_name='Programación_CD',index=False)
        info_sensado_df.to_excel(writer,sheet_name='Sensado', index=False)
    return True
# ejecutar función de reporte ASN-PROG
generar_documento_asn_prog()
def comunicacion_sensado():
    fecha=str(input('Introduzca la fecha en el formato dd/mm/aaaa: '))
    dia=fecha.split('/')[0]
    mes=fecha.split('/')[1]
    nombre_archivo='Sensado_Prog_'+dia+mes+'.xlsx'
    maestro_contenedor= pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\maestro_por_contenedores_gby.csv',sep=';')
    asn=maestro_contenedor[maestro_contenedor['Fecha']==fecha][['Container','Hora','Fecha', 'Cant cajas', 'Cant productos']]
    del maestro_contenedor
    capacidad_sensado=55000
    cajas_total=math.ceil(asn['Cant cajas'].sum())
    productos_total=math.ceil(asn['Cant productos'].sum())
    estimado_cajas_pre=math.ceil((asn['Cant cajas'].sum())*0.9*0.5)
    estimado_productos_pre=math.ceil((asn['Cant productos'].sum())*0.9*0.5)
    contenedores=asn['Container'].nunique()
    if estimado_productos_pre >= capacidad_sensado:
        sensado_abast=0 
    else: sensado_abast=capacidad_sensado-estimado_productos_pre
    leyenda_capacidad= 'Capacidad de sensado considerada [unidades de sku por día]: '
    leyenda_cajas_total='Cajas a recibr en total [cajas]: '
    leyenda_productos_total='Unidades de sku a recibir en total [unidades de sku]: '
    leyenda_cajas='Cajas a recibr aprox a sensado [cajas]: '
    leyenda_productos='Unidades de sku a recibr aprox a sensado [unidades de sku]: '
    leyenda_sensado='Se necesita abastecer a sensado [unidades de sku]: '
    columna=['Cantidad']
    cuadro_informacion={leyenda_capacidad:capacidad_sensado,
                        leyenda_cajas_total:cajas_total,
                        leyenda_productos_total:productos_total,
                        leyenda_cajas:estimado_cajas_pre,
                        leyenda_productos:estimado_productos_pre,
                        leyenda_sensado:sensado_abast}
    info_sensado_df=pd.DataFrame.from_dict(cuadro_informacion,orient='index',columns=columna)
    info_sensado_df.reset_index(inplace=True)
    info_sensado_df.rename(columns={'index':'Item'}, inplace=True)
    info_sensado_df.to_excel('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\archivos_asn_programacion\\'+nombre_archivo,sheet_name='Sensado')
    
# Buscando cuántos días de demora hay entre la ETA y la fecha de programación a CD
def cant_cont(i):
    n=df_eta_fecha[df_eta_fecha['Fecha']==i]['Container'].nunique()
    return n
def cant_cont_eta(i):
    n=df_eta_fecha[df_eta_fecha['ETA']==i]['Container'].nunique()
    return n

maestro_prog['Fecha']=maestro_prog['Fecha'].apply(fecha_eta)
maestro_prog['ETA']=maestro_prog['ETA'].apply(fecha_eta)
# df_eta_fecha=df_eta_fecha[df_eta_fecha['Fecha']>=datetime.date(year=2021,month=7,day=15)]
# Def del dataframe
df_eta_fecha=maestro_prog[['Container','ETA','Fecha','Dias libres']]
# Def del campo lead time
df_eta_fecha['Dias entrega']=df_eta_fecha['Fecha']-df_eta_fecha['ETA']
# Ordenamiento de datos
df_eta_fecha.sort_values(by='Fecha',inplace=True)
# días de lead time
df_eta_fecha['Dias entrega']=df_eta_fecha['Dias entrega'].apply(lambda x: x.days)
#  media móvil simple de los 5 días anteiores por día
# df_eta_fecha['media movil dias_entrega']=(df_eta_fecha[['Dias entrega']].rolling(5).mean()).apply(lambda x: round(x,0))
# n de contendores por día
df_eta_fecha['n cont']=df_eta_fecha['Fecha'].apply(cant_cont)
df_eta_fecha['n cont_ETA']=df_eta_fecha['ETA'].apply(cant_cont_eta)

df_eta=df_eta_fecha[['ETA','n cont_ETA']].drop_duplicates(keep='first')
df_eta.sort_values(by='ETA', inplace=True)
df_eta1=pd.merge((df_eta_fecha['Fecha'].drop_duplicates(keep='first')),right=df_eta,how='left', left_on='Fecha',right_on='ETA')
df_eta1.drop('ETA',axis=1,inplace=True)
df_eta1.fillna(0, inplace=True)
df_eta2=pd.DataFrame(index=pd.date_range(start=df_eta1['Fecha'].iloc[0],end=df_eta1['Fecha'].iloc[-1]))
df_eta2.reset_index(inplace=True)
df_eta2.rename(columns={'index':'Fecha'}, inplace=True)
df_eta2=df_eta2['Fecha'].apply(lambda x: x.date())
df_eta2=pd.merge(left=df_eta2,right=df_eta,how='left',left_on='Fecha', right_on='ETA')
df_eta2.drop('ETA',axis=1, inplace=True)
df_eta2.fillna(0, inplace=True)
df_eta2['n cont ETA 7']=df_eta2['n cont_ETA'].rolling(7).sum()
df_eta2.fillna(0, inplace=True)
df_eta2['n cont ETA 7']=df_eta2['n cont ETA 7'].apply(lambda x: round(x))
# df_eta1['n cont ETA 7']=df_eta1.rolling(7).sum().apply(lambda x: round(x))
# df_eta1.fillna(0, inplace=True)

# preparación del df
df_eta_fecha.reset_index(inplace=True)
df_eta_fecha.drop('index',inplace=True,axis=1)
df2=df_eta_fecha[df_eta_fecha['Dias libres']!=0][['Fecha','n cont',
                                                           'Dias entrega']]

# df_aux=df2.copy()
# df2=df_eta_fecha[['Fecha','Container','Dias entrega']]
# df2['n cont']=df2['Fecha'].apply(cant_cont)
# df2=df2[['Fecha','Dias entrega','n cont']]
 
df2=df2.groupby(['Fecha','n cont']).mean()
df2.reset_index(inplace=True)
# df2.to_csv(r'C:\Users\bpineda\Desktop\a.csv')
# df2=pd.read_csv(r'C:\Users\bpineda\Desktop\a.csv')
df2['Fecha']=df2['Fecha'].apply(fecha_eta)
df2['Dias entrega']=df2['Dias entrega'].apply(lambda x:round(x,0))
df2['media movil dias_entrega']=(df2[['Dias entrega']].rolling(5).mean()).apply(lambda x: round(x,0))
df2['media movil dias_entrega_max']=(df2[['Dias entrega']].rolling(5).max())
df2['media movil dias_entrega_min']=(df2[['Dias entrega']].rolling(5).min())
# Para el gráfico cotidiano
df2['media movil dias_entrega_3m']=(df2[['Dias entrega']].rolling(90).mean()).apply(lambda x: round(x,0))
df2['n cont movil']=(df2[['n cont']].rolling(30).mean()).apply(lambda x: round(x,0))
# Para el gráfico de contenedores
df2['media movil dias_entrega_1m']=(df2[['Dias entrega']].rolling(5).mean()).apply(lambda x: round(x,0))
df2['media movil dias_entrega_1m q20 movil']=(df2[['Dias entrega']].rolling(5).quantile(quantile=0.2)).apply(lambda x: round(x,0))
df2['n cont movil']=(df2[['n cont']].rolling(5).mean()).apply(lambda x: round(x,0))
df2['n cont suma movil']=(df2[['n cont']].rolling(5).sum()).apply(lambda x: round(x,0))
df2['n cont q80 movil']=(df2[['n cont']].rolling(5).quantile(quantile=0.8)).apply(lambda x: round(x,0))
# df2.reset_index(inplace=True)
# df2['media movil dias_entrega']=(df2[['Dias entrega']].rolling(5).mean()).apply(lambda x: round(x,0))
df2=pd.merge(left=df2,right=df_eta2, how='left',on='Fecha')
df2['n cont movil ETA 7']=(df2[['n cont ETA 7']].rolling(5).mean()).apply(lambda x: round(x,0))


fig,ax=plt.subplots(figsize=(14,6))
ax.set_title('Lead time de recepción en CD')
# ax.plot(np.arange(1,((df2['media movil dias_entrega']).shape[0])+1),df2['media movil dias_entrega'],
ax.plot(df2['Fecha'],df2['media movil dias_entrega'],
        color='red',alpha=0.6,markersize=1.5,marker='s',markerfacecolor='blue',
        markeredgecolor='blue', label='Lead time')
ax.plot(df2['Fecha'],df2['media movil dias_entrega_max'],
        color='black',alpha=1, ls="--",lw=0.7,markersize=0,marker='s',markerfacecolor='red',
        markeredgecolor='blue', label='Lead time_max')
ax.plot(df2['Fecha'],df2['media movil dias_entrega_min'],
        color='black',alpha=1,ls="--",lw=0.7,markersize=0,marker='s',markerfacecolor='black',
        markeredgecolor='blue', label='Lead time_min')
ax.plot(df2['Fecha'],df2['media movil dias_entrega_3m'],
        color='orange',alpha=0.7,ls="-",lw=1.5,markersize=0,marker='s',markerfacecolor='black',
        markeredgecolor='blue', label='Lead time_90 mov')
plt.axhline(y=20, color='black',ls='-', alpha=0.2)
plt.axhline(y=15, color='black',ls='-', alpha=0.2)
plt.axhline(y=10, color='black',ls='-', alpha=0.2)
plt.axhline(y=5, color='black',ls='-', alpha=0.2)

ax1=ax.twinx()
ax1.plot(df2['Fecha'],df2['n cont ETA 7'],color='purple',ls='-',linewidth=0.6,
            label='n prog contenedores', alpha=0.4)
# ax1.plot(df2['Fecha'],df2['n cont movil ETA 7'],color='green',ls='-',
#             label='n prog contenedores', alpha=1)
# ax1.plot(range(1,((df2['Fecha']).shape[0])+1),df2['n cont'].mean()*(np.ones(range(1,((df2['Fecha']).shape[0])+1))),color='yellow',ls='-')
ax1.set_ylabel('Con ETA últimos 7 días [suma de contenedores]')
ax.set_ylabel('Lead time [días]')
ax.set_xlabel('Fechas de llegadas de contenedores desde el {} hasta el {}'.format(list(df2['Fecha'].head(1))[0],list(df2['Fecha'].tail(1))[0]))
# ax.set_xlabel('Ultimos 30 días de programaciones')
# ax.set_xlim(((df2['Fecha'].tail(30)).min())-datetime.timedelta(days=1),
            # (df2['Fecha'].tail(30)).max()+datetime.timedelta(days=1))
ax.set_ylim(int((df2['media movil dias_entrega_min']).min())-2,
            int((df2['media movil dias_entrega_max']).max())+2)
ax.legend()
# plt.grid(axis='y',color='black', alpha=0.7)
# plt.axvline(x=datetime.date(year=2021, month=6, day=25), color='black',ls='dotted', alpha=0.45)
# plt.axvline(x=datetime.date(year=2021, month=7, day=5), color='black',ls='dotted', alpha=0.45)
# plt.axvline(x=datetime.date(year=2021, month=12, day=28), color='purple',ls='dotted', alpha=0.45)
# plt.axvline(x=datetime.date(year=2022, month=1, day=6), color='purple',ls='dotted', alpha=0.45)
# plt.axvline(x=datetime.date(year=2022, month=6, day=29), color='green',ls='dotted', alpha=0.45)
# plt.axvline(x=datetime.date(year=2022, month=7, day=4), color='green',ls='dotted', alpha=0.45)
# plt.axvline(x=datetime.date(year=2022, month=12, day=27), color='red',ls='dotted', alpha=0.45)
# plt.axvline(x=datetime.date(year=2023, month=1, day=3), color='red',ls='dotted', alpha=0.45)
# plt.axvline(x=datetime.date(year=2023, month=12, day=26), color='black',ls='dotted', alpha=0.45)
# plt.axvline(x=datetime.date(year=2024, month=1, day=2), color='black',ls='dotted', alpha=0.45)

plt.savefig(r'C:\Users\bpineda\Desktop\ETA-PROG9.png',dpi=100)
# ax1.legend(loc=2)
# ax.plot(df_eta_fecha['Fecha'].tail(30),df_eta_fecha['media movil dias_entrega'].tail(30),
#         color='red',markersize=6,marker='x',markerfacecolor='blue')
# ax.set_yticks([3])
# plt.show()  
    fig,ax=plt.subplots(figsize=(14,6))
    ax.set_title('Lead time de recepción en CD')
    # ax.plot(np.arange(1,((df2['media movil dias_entrega']).shape[0])+1),df2['media movil dias_entrega'],
    # ax.plot(df2['Fecha'],df2['media movil dias_entrega'],
    #         color='red',alpha=0.6,markersize=1.5,marker='s',markerfacecolor='blue',
    #         markeredgecolor='blue', label='Lead time')
    # ax.plot(df2['Fecha'],df2['media movil dias_entrega_max'],
    #         color='black',alpha=1, ls="--",lw=0.7,markersize=0,marker='s',markerfacecolor='red',
    #         markeredgecolor='blue', label='Lead time_max')
    # ax.plot(df2['Fecha'],df2['media movil dias_entrega_min'],
    #         color='black',alpha=1,ls="--",lw=0.7,markersize=0,marker='s',markerfacecolor='black',
    #         markeredgecolor='blue', label='Lead time_min')
    ax.plot(df2['Fecha'],df2['media movil dias_entrega_1m'],
            color='goldenrod',alpha=1,ls="-",lw=1.5,markersize=0,marker='s',markerfacecolor='black',
            markeredgecolor='goldenrod', label='Lead time_5 mov')
    ax.plot(df2['Fecha'],df2['media movil dias_entrega_max'],
            color='black',alpha=0.7, ls=":",lw=0.9,markersize=0,marker='s',markerfacecolor='red',
            markeredgecolor='blue', label='Lead time_max')
    ax.plot(df2['Fecha'],df2['media movil dias_entrega_min'],
            color='black',alpha=0.7,ls=":",lw=0.9,markersize=0,marker='s',markerfacecolor='black',
            markeredgecolor='blue', label='Lead time_min')
    plt.axhline(y=20, color='goldenrod',ls='--', alpha=0.6)
    plt.axhline(y=15, color='goldenrod',ls='--', alpha=0.6)
    plt.axhline(y=10, color='goldenrod',ls='--', alpha=1)
    plt.axhline(y=5, color='goldenrod',ls='--', alpha=0.6)
    
    ax1=ax.twinx()
    ax1.plot(df2['Fecha'],df2['n cont movil'],color='purple',ls='-',linewidth=0.6,
                label='n prog contenedores', alpha=1)
    ax1.plot(df2['Fecha'],df2['n cont suma movil'],color='purple',ls='-',linewidth=0.6,
                label='n prog contenedores 7', alpha=1)
    plt.axhline(y=7, color='purple',ls='--', alpha=1)
    plt.axhline(y=4, color='purple',ls=':', alpha=1)
    # ax1.plot(df2['Fecha'],df2['n cont movil ETA 7'],color='green',ls='-',
    #             label='n prog contenedores', alpha=1)
    # ax1.plot(range(1,((df2['Fecha']).shape[0])+1),df2['n cont'].mean()*(np.ones(range(1,((df2['Fecha']).shape[0])+1))),color='yellow',ls='-')
    ax1.set_ylabel('recibidos los últimos 5 días [mean diario de contenedores]')
    ax.set_ylabel('Lead time [días]')
    ax.set_xlabel('Fechas de llegadas de contenedores desde el {} hasta el {}'.format(list(df2['Fecha'].head(1))[0],list(df2['Fecha'].tail(1))[0]))
    # ax.set_xlabel('Ultimos 30 días de programaciones')
    # ax.set_xlim(((df2['Fecha'].tail(30)).min())-datetime.timedelta(days=1),
                # (df2['Fecha'].tail(30)).max()+datetime.timedelta(days=1))
    ax.set_ylim(int((df2['media movil dias_entrega_min']).min())-2,
                int((df2['media movil dias_entrega_max']).max())+2)
    ax.legend()
    # plt.grid(axis='y',color='black', alpha=0.7)
    # plt.axvline(x=datetime.date(year=2021, month=6, day=25), color='black',ls='dotted', alpha=0.45)
    # plt.axvline(x=datetime.date(year=2021, month=7, day=5), color='black',ls='dotted', alpha=0.45)
    # plt.axvline(x=datetime.date(year=2021, month=12, day=28), color='purple',ls='dotted', alpha=0.45)
    # plt.axvline(x=datetime.date(year=2022, month=1, day=6), color='purple',ls='dotted', alpha=0.45)
    # plt.axvline(x=datetime.date(year=2022, month=6, day=29), color='green',ls='dotted', alpha=0.45)
    # plt.axvline(x=datetime.date(year=2022, month=7, day=4), color='green',ls='dotted', alpha=0.45)
    # plt.axvline(x=datetime.date(year=2022, month=12, day=27), color='red',ls='dotted', alpha=0.45)
    # plt.axvline(x=datetime.date(year=2023, month=1, day=3), color='red',ls='dotted', alpha=0.45)
    # plt.axvline(x=datetime.date(year=2023, month=12, day=26), color='black',ls='dotted', alpha=0.45)
    # plt.axvline(x=datetime.date(year=2024, month=1, day=2), color='black',ls='dotted', alpha=0.45)
    
    plt.savefig(r'C:\Users\bpineda\Desktop\ETA-PROG10.png',dpi=120)
    
    df2.to_excel(r'C:\Users\bpineda\Desktop\info_temp.xlsx')

# Informe comercial
'''Reporte de contenedores recibidos en un mes dado por el usuario.
Normalmente se usa para brindar información al área comercial.
'''
dia_inicio=1
dia_final=31
mes,ano=7,2021
dia_uno=datetime.date(year=ano,month=mes,day=dia_inicio)
dia_dos=datetime.date(year=ano,month=mes,day=dia_final)
reporte_comercial=maestro_programaciones[(maestro_programaciones['Fecha']>=dia_uno)&(maestro_programaciones['Fecha']<=dia_dos)][['invoice','Container','Fecha','Cant Codigos','Cant productos','Temporada','Departamentos','Productos']]
reporte_comercial.to_excel('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Reporte CD\\Reporte_comercial_contenedore dentro del mes.xlsx',index=False)

# Informe de la semana comercial
prog_semana_comercial=maestro_prog[(maestro_prog['Fecha']>=dia_inicio)&(maestro_prog['Fecha']<=dia_final)][['invoice','Container','Fecha','Prioridad']]
prog_semana_comercial.to_excel('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Reporte CD\\Reporte_semanal_contenedores.xlsx',index=False)

# Contenedores comercial vs programaciones:
    ''' Con este reporte se ven los contenedores que están el archivo de prioridades
    comerciales más actualizado y que ya tienen o tuvieron fecha y hora de programación a CD
    '''
prog_comercial=maestro_prog[maestro_prog['Container'].isin(prior['CONTAINER'])]
prog_comercial.sort_values('Fecha',inplace=True)
prog_comercial=prog_comercial[['invoice','Container','Prioridad','CD','Fecha']]
prog_comercial.to_excel('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Reporte CD\\Reporte_comercial_prioridades_programadas.xlsx',index=False)

# Contenedores que posiblemente se puedan recibir dado un día de referencia, de preferencia el domingo de la semana anterior a la de la consulta
    ''' Son los contenedores prioritarios aún no programados, pero que posiblemente
    se pueden recibir en la semana entrante. poniendo como referencia el día domingo
    más cercano a la semana de consulta
    '''
diaref2=datetime.date(year=2021,month=10,day=24)
diaref=datetime.date(year=2021,month=10,day=25)
planif=pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\maestro_por_contenedores_gby.csv',sep=';')
planif=planif[planif['Hora'].isna()]
planif['ETA']=planif['ETA'].apply(fecha_eta)
planif=planif[((planif['ETA']<=diaref-datetime.timedelta(days=ultimo_dato-5))&(planif['Dias libres']>0))|((planif['Dias libres']==0)&(planif['ETA']<=diaref-datetime.timedelta(days=ultimo_dato-5)))][['invoice','Container','Tipo_transporte','ETA','Prioridad']]
planif=planif[planif['Container'].isin(prior['CONTAINER'])]
planif.to_excel('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Reporte CD\\Reporte_comercial_planificación_semana.xlsx',index=False)

# Contenedores de prioridades comerciales que aún no se ha traído pero se estippula un fecha de llegada
    ''' Es una lista con todos los contenedores prioritarios de la lista,
    que aún no han sido programados al CD pero de acuerdo al promedio de días
    que el CD se está demorando en recibir contenedores desde que llegan se 
    pueden obtener las fecha estimadas de llegada a CD.
    '''
planif=pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\maestro_por_contenedores_gby.csv',sep=';')
planif=planif[planif['Hora'].isna()]
planif=planif[planif['Container'].isin(prior['CONTAINER'])][['invoice','Container','Tipo_transporte','Prioridad','ETA']]
planif['ETA']=planif['ETA'].apply(fecha_eta)
planif['Fecha estimada']=planif['Container'].apply(fecha_estimada)
planif.rename(columns={'ETA':'Fecha puerto'},inplace=True)
planif[['invoice','Container','Tipo_transporte','Prioridad','Fecha puerto','Fecha estimada']].to_excel('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Reporte CD\\Reporte_comercial_planificación_contenedores_prioritarios.xlsx',index=False)

# Lista de contenedores a enviar a Carla para compartir información con antelación
    ''' Es una lista con los contenedores que más probabilidades tienen de
    ser programados a CD dada su fecha de llegada, sin importar su prioridad.
    Esta información es para la administradora de las ASN
    '''
hoy=datetime.date.today()
hoy=datetime.date(year=2021,month=11,day=14)
planif=pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\maestro_por_contenedores_gby.csv',sep=';')
planif=planif[planif['Hora'].isna()]
planif['ETA']=planif['ETA'].apply(fecha_eta)
planif=planif.sort_values('ETA')
planif=planif[planif['ETA']<=hoy-datetime.timedelta(days=ultimo_dato-6)][['invoice','Container','Grupo','ETA','Nave','Cant Codigos','Cant cajas','Cant productos']]
planif.to_excel('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Reporte CD\\Contenedores_ASN.xlsx',index=False)

# Propuesta de programación para los encargados de supervisión
'''Reporte para presentar propuestas de programación a los supervisores de recepción. Se hace con
el fin de poder estimar la jornada de trabajo para el equipo de recepción.
'''
fecha='10/2/2022'
maestro_contenedor= pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\maestro_por_contenedores_gby.csv',sep=';')
maestro_contenedor= pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\maestro_por_contenedores_gby.csv',sep=';')
programacion=maestro_contenedor[maestro_contenedor['Fecha']==fecha][['invoice','Container','Tipo_transporte','Tipo','Departamentos','Productos','Cant cajas','Cant Codigos','Hora','Fecha']]
programacion.sort_values('Hora',axis=0,inplace=True)
dia=fecha.split('/')[0]
mes=fecha.split('/')[1]
nombre_archivo='Propuesta_Prog.xlsx'
programacion.to_excel('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\archivos_asn_programacion\\'+nombre_archivo,sheet_name='Programación',index=False)

# Prog semanal para área comercial y logística.
'''Reporte para informar de la planificación semanal de contenedores al CD Tricot y área comercial
'''
diareporte=datetime.date.today()
maestro_contenedor= pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\maestro_por_contenedores_gby.csv',sep=';')
# maestro_contenedor=maestro_contenedor[~maestro_contenedor['Hora'].isna()]
maestro_contenedor=maestro_contenedor[~maestro_contenedor['Fecha'].isna()]
# maestro_contenedor=maestro_contenedor[['invoice','Container','Grupo','Tipo','Tipo_transporte','Nave','CD','Cant Codigos','Cant cajas','Cant productos','Hora','Fecha']]
maestro_contenedor2=maestro_contenedor[['invoice','Container','Departamentos','Productos','Grupo','Cant cajas','Prioridad','Nave','ETA','Dias libres','CD','Hora','Fecha','Embarcador']]
maestro_contenedor3=maestro_contenedor[['invoice','Container','Grupo','Temporada','Departamentos','Productos','Cant cajas','Cant productos', 'Prioridad','ETA','Dias libres','CD','Hora','Fecha','Embarcador']]
maestro_contenedor2['Fecha']=maestro_contenedor2['Fecha'].apply(fecha_eta)
maestro_contenedor3['Fecha']=maestro_contenedor3['Fecha'].apply(fecha_eta)
maestro_contenedor2.sort_values(by='Fecha',inplace=True)
maestro_contenedor3.sort_values(by='Fecha',inplace=True)
# maestro_contenedor.to_excel('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\archivos_asn_programacion\\Prog_semana.xlsx',index=False)
if len(str(diareporte.month))==1 and len(str(diareporte.day))==1:
    fecha=str(diareporte.year)+'0'+str(diareporte.month)+'0'+str(diareporte.day)
elif len(str(diareporte.month))==1 and len(str(diareporte.day))!=1:
    fecha=str(diareporte.year)+'0'+str(diareporte.month)+str(diareporte.day)
elif len(str(diareporte.month))!=1 and len(str(diareporte.day))==1:
    fecha=str(diareporte.year)+str(diareporte.month)+'0'+str(diareporte.day)
elif len(str(diareporte.month))!=1 and len(str(diareporte.day))!=1:
    fecha=str(diareporte.year)+str(diareporte.month)+str(diareporte.day)

nombre_archivo='Planificación.xlsx'
parent_dir='C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\archivos_asn_programacion\\informe seguimiento'
directory=fecha
final_path=os.path.join(parent_dir, directory)
os.mkdir(final_path)
final_path=os.path.join(final_path, nombre_archivo)
with pd.ExcelWriter(final_path) as writer:
    maestro_contenedor2.to_excel(writer,index=False,sheet_name='Plan semana')
    maestro_contenedor3.to_excel(writer,index=False,sheet_name='Plan semana_CD')
    dic_embarcadores={}
    dic_emb={'DHL EMB':'DHL','SPARX T':'SPARX','NWPT':'NOWPORTS','JAS':'JAS'}
    for i in dic_campos_plano['embarcadores_plano']:
        print(i)
        nombre_embarcador="Plan_"+str(i)
        dic_embarcadores[nombre_embarcador]=maestro_contenedor[maestro_contenedor['Embarcador']==dic_emb[str(i)]][['Container','Nave','Grupo','Cant cajas','ETA','Dias libres','CD','Hora','Fecha','Embarcador']]
        dic_embarcadores[nombre_embarcador]['Fecha']=dic_embarcadores[nombre_embarcador]['Fecha'].apply(fecha_eta)
        dic_embarcadores[nombre_embarcador].sort_values(by='Fecha',inplace=True)
        dic_embarcadores[nombre_embarcador].to_excel(writer,index=False,sheet_name=nombre_embarcador)
        # contador+=1
    
    


del maestro_contenedor
del maestro_contenedor2
del maestro_contenedor3


'''Cuanto contenedores debiesen llegar hasta un fecha determinada, dada por el consultante.
Puede usarse para planificar las operaciones dadas las llegadas hasta una fecha dada.'''
def n_contenedores_fecha():
    fecha= str(input('Ingresar la fecha inicial de consulta: '))
    fecha_fin= str(input('Ingresar la fecha final de la consulta: '))
    fecha= fecha.split('/')
    fecha_fin= fecha_fin.split('/')
    fecha= datetime.date(day= int(fecha[0]),month=int(fecha[1]),year=int(fecha[2]))
    fecha_fin= datetime.date(day= int(fecha_fin[0]),month=int(fecha_fin[1]),year=int(fecha_fin[2]))
    maestro_contenedor= pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\maestro_por_contenedores_gby.csv',sep=';')
    maestro_contenedor['Fecha pm']=maestro_contenedor['Fecha pm'].apply(fecha_eta)
    maestro_contenedor['Deadline']=maestro_contenedor['Deadline'].apply(fecha_eta)
    n_contenedores=maestro_contenedor[maestro_contenedor['Fecha pm']<fecha]['Container'].nunique()
    n_contenedores_dead=maestro_contenedor[(maestro_contenedor['Tipo_transporte']=='FCL')&(maestro_contenedor['Deadline']>fecha)&(maestro_contenedor['Deadline']<fecha_fin)]['Container'].nunique()
    n_contenedores_dead_lista=maestro_contenedor[(maestro_contenedor['Tipo_transporte']=='FCL')&(maestro_contenedor['Deadline']>fecha)&(maestro_contenedor['Deadline']<fecha_fin)][['Container','Nave']]
    n_contenedores_dead_antes=maestro_contenedor[(maestro_contenedor['Tipo_transporte']=='FCL')&(maestro_contenedor['Deadline']>fecha)&(maestro_contenedor['Deadline']<fecha_fin)&(maestro_contenedor['Fecha pm']<fecha)]['Container'].nunique()    
    n_contenedores_lcl=maestro_contenedor[(maestro_contenedor['Tipo_transporte']=='LCL')&(maestro_contenedor['Fecha pm']>fecha)&(maestro_contenedor['Fecha pm']<fecha_fin)]['Container'].nunique()
    n_contenedores_lcl_lista=maestro_contenedor[(maestro_contenedor['Tipo_transporte']=='LCL')&(maestro_contenedor['Fecha pm']>fecha)&(maestro_contenedor['Fecha pm']<fecha_fin)][['Container','Nave']]
    n_contenedores_hermanos=maestro_contenedor[(maestro_contenedor['Fecha pm']<fecha)&(maestro_contenedor['Grupo']!='Contenedor unico')]['Container'].nunique()
    n_stock_hasta_inicio=maestro_contenedor[maestro_contenedor['Fecha pm']<fecha]['Cant productos'].sum()
    fecha_max_hermano=maestro_contenedor[(maestro_contenedor['Fecha pm']<fecha)&(maestro_contenedor['Grupo']!='Contenedor unico')]['Fecha pm'].max()
    fecha_min_demu=maestro_contenedor[(maestro_contenedor['Tipo_transporte']=='FCL')&(maestro_contenedor['Deadline']>fecha)&(maestro_contenedor['Deadline']<fecha_fin)]['Deadline'].min()
    fecha_max_demu=maestro_contenedor[(maestro_contenedor['Tipo_transporte']=='FCL')&(maestro_contenedor['Deadline']>fecha)&(maestro_contenedor['Deadline']<fecha_fin)]['Deadline'].max()
    fecha_max_demu_antes=maestro_contenedor[(maestro_contenedor['Tipo_transporte']=='FCL')&(maestro_contenedor['Deadline']>fecha)&(maestro_contenedor['Deadline']<fecha_fin)]['Fecha pm'].max()
    return n_contenedores, n_contenedores_dead, n_contenedores_dead_lista, n_contenedores_dead_antes, n_stock_hasta_inicio, n_contenedores_hermanos, n_contenedores_lcl, n_contenedores_lcl_lista, fecha_max_hermano, fecha_min_demu, fecha_max_demu, fecha_max_demu_antes
n_contenedores_dada_una_fecha, n_contenedores_con_demurrage, n_contenedores_con_demurrage_lista, n_contenedores_con_demurrage_antes, stock, n_hermanos, n_lcl, n_lcl_lista, fecha_max, fecha_min_demu, fecha_max_demurr, fecha_max_demu_antes=n_contenedores_fecha()


def informacion_cont_disp():
    '''Función para extraer la información de los contenedores disponibles al momento de la consulta.
    requiere tener la información de disponibilidad actualizada.
    '''
    maestro_contenedor= pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\maestro_por_contenedores_gby.csv',sep=';')    
    maestro_contenedor=maestro_contenedor[maestro_contenedor['Disponibilidad']!='Sin fecha'][['invoice','Container','Departamentos','Productos','Grupo','Cant cajas','Prioridad','Nave','ETA','Dias libres','Deadline','Fecha pm','Disponibilidad']]
    maestro_contenedor.to_excel('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Reporte CD\\Información_disponibilidad.xlsx',sheet_name='Disponibilidad',index=False)
    return 0

informacion_cont_disp()

def informacion_recepcion():
    '''Función que entrega la información de recepciones entre 2 fechas dadas
    por el usuario, incluyendolas.
    '''
    
    monday=datetime.date(year=2023, month=3, day=1)
    friday=datetime.date(year=2023, month=3, day=15)
    # datetime.datetime.today(entrega un TimeStamp)
    
    maestro_program= pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\programaciones\\maestro_programaciones.csv', sep=',')
    if maestro_program.shape[1]==1:
        maestro_program= pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\programaciones\\maestro_programaciones.csv', sep=';')
    maestro_program['Fecha']=maestro_program['Fecha'].apply(fecha_eta)
    maestro_program=maestro_program[(maestro_program['Fecha']>=monday)&(maestro_program['Fecha']<=friday)]
    maestro_program=maestro_program[['invoice','Container','Departamentos','Productos','Grupo','Cant cajas','Cant productos','Prioridad','Nave','ETA','Dias libres','Fecha','CD']]
    maestro_program.to_excel('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Reporte CD\\Información_recepción.xlsx',sheet_name='Recepcion semana',index=False)
    return 0

informacion_recepcion()

def n_o_cero(i):
    global hoy
    global stop
    global fecha_stop
    if stop=="1" and i>fecha_stop:
        return 0
    if i.strftime('%A')=='Sunday' or i.strftime('%A')=='Saturday' or i<hoy:
        return 0
    else:
        return 6
def n_o_cero_cajas_7500(i):
    
    if stop=="1" and i>fecha_stop:
        return 0
    if i.strftime('%A')=='Sunday' or i.strftime('%A')=='Saturday' or i<hoy:
        return 0
    else:
        return 7500
    
def n_o_cero_cajas_5500(i):
    
    if stop=="1" and i>fecha_stop:
        return 0
    if i.strftime('%A')=='Sunday' or i.strftime('%A')=='Saturday' or i<hoy:
        return 0
    else:
        return 5500
    
def fun_valores_eta_cajas_7500(i):
    
    if i.name==0:
        # acum_test.at[0,'Valor_remanente t<t0']=0
        if rango_fechas.at[i.name,'Cajas ETA']+rango_fechas.at[i.name,'Valor t_0 ETA_7500']-rango_fechas.at[i.name,'caja_base_7500'] < 0:
            rango_fechas.at[i.name,'Valor t_0 ETA_7500']= 0
        else:
            rango_fechas.at[i.name,'Valor t_0 ETA_7500']= rango_fechas.at[i.name,'Cajas ETA']+rango_fechas.at[i.name,'Valor t_0 ETA_7500']-rango_fechas.at[i.name,'caja_base_7500']
    else:
        if rango_fechas.at[i.name,'Cajas ETA']+rango_fechas.at[int(i.name)-1,'Valor t_0 ETA_7500']-rango_fechas.at[i.name,'caja_base_7500'] < 0:
            rango_fechas.at[i.name,'Valor t_0 ETA_7500']= 0
        else:
            rango_fechas.at[i.name,'Valor t_0 ETA_7500']= rango_fechas.at[i.name,'Cajas ETA']+rango_fechas.at[int(i.name)-1,'Valor t_0 ETA_7500']-rango_fechas.at[i.name,'caja_base_7500']
    return 0        
def fun_valores_eta_cajas_5500(i):
    
    if i.name==0:
        # acum_test.at[0,'Valor_remanente t<t0']=0
        if rango_fechas.at[i.name,'Cajas ETA']+rango_fechas.at[i.name,'Valor t_0 ETA_5500']-rango_fechas.at[i.name,'caja_base_5500'] < 0:
            rango_fechas.at[i.name,'Valor t_0 ETA_5500']= 0
        else:
            rango_fechas.at[i.name,'Valor t_0 ETA_5500']= rango_fechas.at[i.name,'Cajas ETA']+rango_fechas.at[i.name,'Valor t_0 ETA_5500']-rango_fechas.at[i.name,'caja_base_5500']
    else:
        if rango_fechas.at[i.name,'Cajas ETA']+rango_fechas.at[int(i.name)-1,'Valor t_0 ETA_5500']-rango_fechas.at[i.name,'caja_base_5500'] < 0:
            rango_fechas.at[i.name,'Valor t_0 ETA_5500']= 0
        else:
            rango_fechas.at[i.name,'Valor t_0 ETA_5500']= rango_fechas.at[i.name,'Cajas ETA']+rango_fechas.at[int(i.name)-1,'Valor t_0 ETA_5500']-rango_fechas.at[i.name,'caja_base_5500']
    return 0        
def fun_valores_disp_cajas_7500(i):
    
    if i.name==0:
        # acum_test.at[0,'Valor_remanente t<t0']=0
        if rango_fechas.at[i.name,'Cajas Disponibilidad']+rango_fechas.at[i.name,'Valor t_0 Disp_7500']-rango_fechas.at[i.name,'caja_base_7500'] < 0:
            rango_fechas.at[i.name,'Valor t_0 Disp_7500']= 0
        else:
            rango_fechas.at[i.name,'Valor t_0 Disp_7500']= rango_fechas.at[i.name,'Cajas Disponibilidad']+rango_fechas.at[i.name,'Valor t_0 Disp_7500']-rango_fechas.at[i.name,'caja_base_7500']
    else:
        if rango_fechas.at[i.name,'Cajas Disponibilidad']+rango_fechas.at[int(i.name)-1,'Valor t_0 Disp_7500']-rango_fechas.at[i.name,'caja_base_7500'] < 0:
            rango_fechas.at[i.name,'Valor t_0 Disp_7500']= 0
        else:
            rango_fechas.at[i.name,'Valor t_0 Disp_7500']= rango_fechas.at[i.name,'Cajas Disponibilidad']+rango_fechas.at[int(i.name)-1,'Valor t_0 Disp_7500']-rango_fechas.at[i.name,'caja_base_7500']
    return 0
def fun_valores_disp_cajas_5500(i):
    
    if i.name==0:
        # acum_test.at[0,'Valor_remanente t<t0']=0
        if rango_fechas.at[i.name,'Cajas Disponibilidad']+rango_fechas.at[i.name,'Valor t_0 Disp_5500']-rango_fechas.at[i.name,'caja_base_5500'] < 0:
            rango_fechas.at[i.name,'Valor t_0 Disp_5500']= 0
        else:
            rango_fechas.at[i.name,'Valor t_0 Disp_5500']= rango_fechas.at[i.name,'Cajas Disponibilidad']+rango_fechas.at[i.name,'Valor t_0 Disp_5500']-rango_fechas.at[i.name,'caja_base_5500']
    else:
        if rango_fechas.at[i.name,'Cajas Disponibilidad']+rango_fechas.at[int(i.name)-1,'Valor t_0 Disp_5500']-rango_fechas.at[i.name,'caja_base_5500'] < 0:
            rango_fechas.at[i.name,'Valor t_0 Disp_5500']= 0
        else:
            rango_fechas.at[i.name,'Valor t_0 Disp_5500']= rango_fechas.at[i.name,'Cajas Disponibilidad']+rango_fechas.at[int(i.name)-1,'Valor t_0 Disp_5500']-rango_fechas.at[i.name,'caja_base_5500']
    return 0

def fun_valores_eta(i):
    
    if i.name==0:
        # acum_test.at[0,'Valor_remanente t<t0']=0
        if rango_fechas.at[i.name,'Contenedores ETA']+rango_fechas.at[i.name,'Valor t_0 ETA']-rango_fechas.at[i.name,'cont_base'] < 0:
            rango_fechas.at[i.name,'Valor t_0 ETA']= 0
        else:
            rango_fechas.at[i.name,'Valor t_0 ETA']= rango_fechas.at[i.name,'Contenedores ETA']+rango_fechas.at[i.name,'Valor t_0 ETA']-rango_fechas.at[i.name,'cont_base']
    else:
        if rango_fechas.at[i.name,'Contenedores ETA']+rango_fechas.at[int(i.name)-1,'Valor t_0 ETA']-rango_fechas.at[i.name,'cont_base'] < 0:
            rango_fechas.at[i.name,'Valor t_0 ETA']= 0
        else:
            rango_fechas.at[i.name,'Valor t_0 ETA']= rango_fechas.at[i.name,'Contenedores ETA']+rango_fechas.at[int(i.name)-1,'Valor t_0 ETA']-rango_fechas.at[i.name,'cont_base']
    return 0        
def fun_valores_disp(i):
    
    if i.name==0:
        # acum_test.at[0,'Valor_remanente t<t0']=0
        if rango_fechas.at[i.name,'Contenedores Disponibilidad']+rango_fechas.at[i.name,'Valor t_0 Disp']-rango_fechas.at[i.name,'cont_base'] < 0:
            rango_fechas.at[i.name,'Valor t_0 Disp']= 0
        else:
            rango_fechas.at[i.name,'Valor t_0 Disp']= rango_fechas.at[i.name,'Contenedores Disponibilidad']+rango_fechas.at[i.name,'Valor t_0 Disp']-rango_fechas.at[i.name,'cont_base']
    else:
        if rango_fechas.at[i.name,'Contenedores Disponibilidad']+rango_fechas.at[int(i.name)-1,'Valor t_0 Disp']-rango_fechas.at[i.name,'cont_base'] < 0:
            rango_fechas.at[i.name,'Valor t_0 Disp']= 0
        else:
            rango_fechas.at[i.name,'Valor t_0 Disp']= rango_fechas.at[i.name,'Contenedores Disponibilidad']+rango_fechas.at[int(i.name)-1,'Valor t_0 Disp']-rango_fechas.at[i.name,'cont_base']
    return 0
def fun_valores_dead(i):
    
    if i.name==0:
        # acum_test.at[0,'Valor_remanente t<t0']=0
        if rango_fechas.at[i.name,'Contenedores Deadline']+rango_fechas.at[i.name,'Valor t_0 Dead']-rango_fechas.at[i.name,'cont_base'] < 0:
            rango_fechas.at[i.name,'Valor t_0 Dead']= 0
        else:
            rango_fechas.at[i.name,'Valor t_0 Dead']= rango_fechas.at[i.name,'Contenedores Deadline']+rango_fechas.at[i.name,'Valor t_0 Dead']-rango_fechas.at[i.name,'cont_base']
    else:
        if rango_fechas.at[i.name,'Contenedores Deadline']+rango_fechas.at[int(i.name)-1,'Valor t_0 Dead']-rango_fechas.at[i.name,'cont_base'] < 0:
            rango_fechas.at[i.name,'Valor t_0 Dead']= 0
        else:
            rango_fechas.at[i.name,'Valor t_0 Dead']= rango_fechas.at[i.name,'Contenedores Deadline']+rango_fechas.at[int(i.name)-1,'Valor t_0 Dead']-rango_fechas.at[i.name,'cont_base']
    return 0
# cont_escolar=plano[plano['TEMPORADA']=='ESCOLAR']['CONTAINER'].unique()

def reporte_llegada_disponibilidad_contenedores():
    
    fin=1
    contador=0
    while fin==1:
        estimacion=str(input("\nIngrese el tipo de visión que necesita: Contención=1 o Estado=0: "))
        fecha_referencia=str(input("\nNecesita utilizar una fecha de referencia? Sí=1  No=0: "))
        stop=str(input("\nNecesita utilizar una fecha de acumulación? Sí=1  No=0: "))
        if estimacion=="1":
            hoy=tuple(
                map(
                    int,
                    tuple(str((input("Ingrese la fecha de contención en formato: dd/mm/aaaa: "))).split("/"))))
            hoy=datetime.date(year=hoy[2], month=hoy[1], day=hoy[0])
            fin=0
        elif estimacion=="0":        
            hoy=datetime.date.today()
            fin=0
        else:
            print("\nPor favor ingrese un tipo de visión correcto...")
            contador+=1
            if contador==5:
                print("\nHa ingresdo 5 veces el tipo de manera errónea, intente el proceso nuevamente.")
                break
        if fecha_referencia=="1":
            referencia=tuple(
                map(
                    int,
                    tuple(str((input("Ingrese la fecha de referencia en formato: dd/mm/aaaa: "))).split("/"))))
            referencia=datetime.date(year=referencia[2], month=referencia[1], day=referencia[0])
        else:   referencia=hoy
        
        if stop=="1":
            fecha_stop=tuple(
                map(
                    int,
                    tuple(str((input("Ingrese la fecha de acumulación en formato: dd/mm/aaaa: "))).split("/"))))
            fecha_stop=datetime.date(year=fecha_stop[2], month=fecha_stop[1], day=fecha_stop[0])
        else: fecha_stop=hoy
        
    maestro_contenedor= pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\maestro_por_contenedores_gby.csv',sep=';')
    maestro_contenedor['ETA']=maestro_contenedor['ETA'].apply(fecha_eta)
    maestro_contenedor['Deadline']=maestro_contenedor['Deadline'].apply(fecha_eta)
    maestro_contenedor['Fecha disp']=maestro_contenedor['Fecha disp'].apply(fecha_eta)
    minf=maestro_contenedor['ETA'].min()
    rango_fechas=pd.DataFrame(index=pd.date_range(start=minf,end=hoy+datetime.timedelta(days=21)))
    rango_fechas.reset_index(inplace=True)
    rango_fechas.rename({'index':'Fechas operación'},axis=1, inplace=True)
    rango_fechas=rango_fechas['Fechas operación'].apply(lambda x: x.date())
    maestro_contenedor['Contenedores']=1
    maestro_contenedor=maestro_contenedor[['Dias libres','ETA','Deadline', 'Fecha disp','Contenedores' , 'Cant cajas', 'Cant productos']]
    gby_eta= maestro_contenedor[['ETA', 'Contenedores', 'Cant cajas', 'Cant productos']].groupby(['ETA']).sum()
    gby_dead= maestro_contenedor[maestro_contenedor['Dias libres']!=0][['Deadline', 'Contenedores', 'Cant cajas', 'Cant productos']].groupby(['Deadline']).sum()
    gby_disp= maestro_contenedor[['Fecha disp', 'Contenedores', 'Cant cajas', 'Cant productos']].groupby(['Fecha disp']).sum()
    rango_fechas=pd.merge(left=rango_fechas, right=gby_eta, how='left', left_on='Fechas operación', right_on='ETA')
    rango_fechas.rename({'Contenedores':'Contenedores ETA', 'Cant cajas':'Cajas ETA', 'Cant productos':'Unidades ETA'},axis=1, inplace=True)
    rango_fechas=pd.merge(left=rango_fechas, right=gby_dead, how='left', left_on='Fechas operación', right_on='Deadline')
    rango_fechas.rename({'Contenedores':'Contenedores Deadline', 'Cant cajas':'Cajas Deadline', 'Cant productos':'Unidades Deadline'},axis=1, inplace=True)
    rango_fechas=pd.merge(left=rango_fechas, right=gby_disp, how='left', left_on='Fechas operación', right_on='Fecha disp')
    rango_fechas.rename({'Contenedores':'Contenedores Disponibilidad', 'Cant cajas':'Cajas Disponibilidad', 'Cant productos':'Unidades Disponibilidad'},axis=1, inplace=True)
    rango_fechas.fillna(value=0, inplace=True)
    
    rango_fechas['cont_base']=rango_fechas['Fechas operación'].apply(n_o_cero)
    rango_fechas['caja_base_5500']=rango_fechas['Fechas operación'].apply(n_o_cero_cajas_5500)
    rango_fechas['caja_base_7500']=rango_fechas['Fechas operación'].apply(n_o_cero_cajas_7500)
    rango_fechas['cont_sum_ETA']=rango_fechas['Contenedores ETA'].cumsum()
    rango_fechas['cont_sum_Dead']=rango_fechas['Contenedores Deadline'].cumsum()
    rango_fechas['cont_sum_Disp']=rango_fechas['Contenedores Disponibilidad'].cumsum()
    rango_fechas['caja_sum_ETA']=rango_fechas['Cajas ETA'].cumsum()
    rango_fechas['caja_sum_Dead']=rango_fechas['Cajas Deadline'].cumsum()
    rango_fechas['caja_sum_Disp']=rango_fechas['Cajas Disponibilidad'].cumsum()
    rango_fechas['cont_sum_base']=rango_fechas['cont_base'].cumsum()
    rango_fechas['caja_sum_base_5500']=rango_fechas['caja_base_5500'].cumsum()
    rango_fechas['caja_sum_base_7500']=rango_fechas['caja_base_7500'].cumsum()
    
    rango_fechas.at[0,'Valor t_0 ETA']=0
    rango_fechas.at[0,'Valor t_0 Disp']=0
    rango_fechas.at[0,'Valor t_0 Dead']=0
    rango_fechas.at[0,'Valor t_0 ETA_5500']=0
    rango_fechas.at[0,'Valor t_0 Disp_5500']=0
    rango_fechas.at[0,'Valor t_0 ETA_7500']=0
    rango_fechas.at[0,'Valor t_0 Disp_7500']=0
    rango_fechas.apply(fun_valores_eta,1)
    rango_fechas.apply(fun_valores_disp,1)
    rango_fechas.apply(fun_valores_dead,1)
    rango_fechas.apply(fun_valores_eta_cajas_5500,1)
    rango_fechas.apply(fun_valores_eta_cajas_7500,1)
    rango_fechas.apply(fun_valores_disp_cajas_5500,1)
    rango_fechas.apply(fun_valores_disp_cajas_7500,1)  
# =============================================================================
# Computo de los contenedores, cajas, etc que se vienen en los próximos 7 días
# a la fecha de operación de referencia sin considerarla dentro del cómputo.
# =============================================================================
    maestro_if=maestro_contenedor[['Fecha disp','Contenedores','Cant cajas']].groupby(['Fecha disp']).sum()
    maestro_if.reset_index(drop=False,inplace=True)
    rango_fechas_if=pd.DataFrame(index=pd.date_range(start=minf,end=maestro_if['Fecha disp'].max()))
    rango_fechas_if.reset_index(drop=False, inplace=True)
    rango_fechas_if.rename({'index':'Fechas'},axis=1, inplace=True)
    rango_fechas_if['Fechas']=rango_fechas_if['Fechas'].apply(lambda x: x.date())
    base_if=pd.merge(left=rango_fechas_if, right=maestro_if, how='left', left_on='Fechas', right_on='Fecha disp')
    base_if.rename({'Contenedores':'Contenedores 7'},axis=1, inplace=True)
    base_if['Contenedores 7'].fillna(0, inplace=True)
    base_if.rename({'Cant cajas':'Cajas 7'},axis=1, inplace=True)
    base_if['Cajas 7'].fillna(0, inplace=True)
    base_if.drop('Fecha disp', axis=1, inplace=True)
    base_if.sort_values(by='Fechas', ascending=False, inplace=True)
    base_if['Contenedores 7']=base_if['Contenedores 7'].shift(periods=1)
    base_if['Contenedores 7'].fillna(0, inplace=True)
    base_if['Contenedores 7']=base_if['Contenedores 7'].rolling(window=7).sum()
    base_if['Contenedores 7'].fillna(0, inplace=True)
    base_if['Cajas 7']=base_if['Cajas 7'].shift(periods=1)
    base_if['Cajas 7'].fillna(0, inplace=True)
    base_if['Cajas 7']=base_if['Cajas 7'].rolling(window=7).sum()
    base_if['Cajas 7'].fillna(0, inplace=True)
    base_if.sort_values(by='Fechas', ascending=True, inplace=True)
# =============================================================================
# Cálculo final de los indicadores IF    
# =============================================================================
    rango_fechas=pd.merge(left=rango_fechas, right=base_if,how='left', left_on='Fechas operación', right_on='Fechas')
    rango_fechas.drop('Fechas', axis=1, inplace=True)
    # rango_fechas['Contenedores 7'].fillna(0, inplace=True)
    # rango_fechas['Cajas 7'].fillna(0, inplace=True)
    rango_fechas['Días_termino ETA']=rango_fechas['Valor t_0 ETA']/4
    rango_fechas['Días_termino Dead']=rango_fechas['Valor t_0 Dead']/4
    rango_fechas['Días_termino Disp']=rango_fechas['Valor t_0 Disp']/4
    rango_fechas['IF ETA']=(rango_fechas['Valor t_0 ETA']+rango_fechas['Contenedores 7'])/30
    rango_fechas['IF Dead']=(rango_fechas['Valor t_0 Dead']+rango_fechas['Contenedores 7'])/30
    rango_fechas['IF Disp']=(rango_fechas['Valor t_0 Disp']+rango_fechas['Contenedores 7'])/30
    rango_fechas['IF ETA_caja_5500']=(rango_fechas['Valor t_0 ETA_5500']+rango_fechas['Cajas 7'])/(5500*5)
    rango_fechas['IF Disp_caja_5500']=(rango_fechas['Valor t_0 Disp_5500']+rango_fechas['Cajas 7'])/(5500*5)
    rango_fechas['IF ETA_caja_7500']=(rango_fechas['Valor t_0 ETA_7500']+rango_fechas['Cajas 7'])/(7500*5)
    rango_fechas['IF Disp_caja_7500']=(rango_fechas['Valor t_0 Disp_7500']+rango_fechas['Cajas 7'])/(7500*5)


    fig,ax= plt.subplots(figsize=(14,6))
    ax.plot(rango_fechas['Fechas operación'],rango_fechas['Valor t_0 Disp'],ls='-',color='purple',label='Contenedores disponibles',alpha=1,lw=1.3, marker='o', markersize=2.9)
    ax.yaxis.set_major_locator(MaxNLocator(nbins=7,integer=True))
    plt.axhline(y=rango_fechas[rango_fechas['Fechas operación']==referencia]['Valor t_0 Disp'].unique()[0], color='purple',ls='-',alpha=0.4)
    ax1=ax.twinx()
    ax1.plot(rango_fechas['Fechas operación'],rango_fechas['IF Disp_caja_5500'],ls=':',color='black',label='IF Disp_5500',alpha=0.4,lw=2, marker='o', markersize=0.8)
    ax1.plot(rango_fechas['Fechas operación'],rango_fechas['IF Disp_caja_7500'],ls='--',color='green',label='IF Disp_7500',alpha=0.6,lw=2, marker='o', markersize=0.8)
    ax1.plot(rango_fechas['Fechas operación'],rango_fechas['IF Disp'],ls='-',color='blue',label='IF Disp',alpha=0.8,lw=2, marker='o', markersize=0.8)
    ax1.yaxis.set_major_locator(MaxNLocator(nbins=6,integer=True))
    ax.set_title('Reporte de carga de trabajo recepción desde el {}'.format(rango_fechas['Fechas operación'].iloc[0]))
    ax.set_ylabel('Contenedores disponibles')
    ax1.set_ylabel('IF')
    ax1.set_ylim([0,10])
    ax.set_xlabel('Fechas')
    h1, l1 = ax.get_legend_handles_labels()
    h2, l2 = ax1.get_legend_handles_labels()
    ax.legend(h1+h2, l1+l2, loc=2)
    plt.axvline(x=referencia, color='black',ls='-',alpha=0.4)
    plt.axvline(x=hoy, color='red',ls='-',alpha=0.5)
    plt.axvline(x=fecha_stop, color='black',ls='-',alpha=0.4)
    plt.grid(color='black',axis='y',lw=0.5,alpha=0.6)
    plt.grid(color='black',axis='x',lw=0.5,alpha=0.5,ls='--')
    plt.savefig(r'C:\Users\bpineda\Desktop\Información_llegada_disponibilidad_contenedores_IF.png', dpi=250)
    
    fig,ax= plt.subplots(figsize=(14,6))
    ax.plot(rango_fechas['Fechas operación'],rango_fechas['cont_sum_ETA'],ls='-.',color='blue',label='Contenedores ETA',alpha=1,lw=2)
    ax.plot(rango_fechas['Fechas operación'],rango_fechas['cont_sum_Dead'],ls='-',color='red',label='Contenedores Dead',alpha=1,lw=2)
    ax.plot(rango_fechas['Fechas operación'],rango_fechas['cont_sum_Disp'],ls='--',color='black',label='Contenedores Disp',alpha=1,lw=2)
    ax.plot(rango_fechas['Fechas operación'],rango_fechas['cont_sum_base'],ls=':',color='green',label='Contenedores Operación',alpha=0.7,lw=1.5)
    ax.set_title('Reporte de flujo de contenedores desde el {}'.format(rango_fechas['Fechas operación'].iloc[0]))
    ax.set_ylabel('n de contenedores acumulados')
    ax.set_xlabel('Fechas')
    ax.legend()
    plt.axvline(x=hoy, color='black',ls='dotted')
    plt.grid(color='black',axis='y',lw=0.5,alpha=0.6)
    plt.grid(color='black',axis='x',lw=0.5,alpha=0.5,ls='--')
    plt.savefig(r'C:\Users\bpineda\Desktop\CD\Programacion diaria de containers\Reporte CD\reporte IF\Información_llegada_disponibilidad_contenedores_acumulado.png', dpi=250)
    rango_fechas.to_excel(r'C:\Users\bpineda\Desktop\CD\Programacion diaria de containers\Reporte CD\reporte IF\Información_llegada_disponibilidad_contenedores_acumulado.xlsx', index=False, sheet_name='Reporte contenedores')

    fig,ax= plt.subplots(figsize=(14,6))
    ax.plot(rango_fechas['Fechas operación'],rango_fechas['Valor t_0 Disp'],ls='-.',color='blue',label='Contenedores Disponibilidad',alpha=1,lw=2,marker=8,ms=8)
    ax.plot(rango_fechas['Fechas operación'],rango_fechas['Valor t_0 Dead'],ls='-',color='red',label='Contenedores Dead',alpha=1,lw=2,marker=8,ms=8)
    # ax.plot(acum['Fechas operación'],acum['cont_sum_Disp'],ls='--',color='black',label='Contenedores Disp',alpha=1,lw=2)
    # ax.plot(acum['Fechas operación'],acum['cont_sum_base'],ls=':',color='green',label='Contenedores Operación',alpha=0.7,lw=1.5)
    ax.set_title('Reporte de flujo de contenedores desde el {}'.format(rango_fechas['Fechas operación'].iloc[0]))
    ax.set_ylabel('Flujo de contenedores')
    ax.set_xlabel('Fechas')
    ax.legend()
    plt.axvline(x=hoy, color='black',ls='dotted')
    plt.grid(color='black',axis='y',lw=0.5,alpha=0.6)
    plt.grid(color='black',axis='x',lw=0.5,alpha=0.5,ls='--')
    # plt.show()
    plt.savefig(r'C:\Users\bpineda\Desktop\CD\Programacion diaria de containers\Reporte CD\reporte IF\Información_llegada_disponibilidad_contenedores_operación.png', dpi=250)
    rango_fechas.to_excel(r'C:\Users\bpineda\Desktop\CD\Programacion diaria de containers\Reporte CD\Información_llegada_disponibilidad_contenedores.xlsx', index=False, sheet_name='Reporte contenedores')
    return 0

reporte_llegada_disponibilidad_contenedores()
# =============================================================================
# reporte demurrage en un periodo dado
# =============================================================================
# Preparación de archivos de trabajo. Maestro de contenedores programados
def reporte_demurrage_periodo():
    maestro=carga_maestro_programaciones()
    maestro=maestro[['Container','Fecha','Deadline','Dias libres','Cant productos','Cant predist','Cant cajas']]
    # maestro=maestro[maestro['Dias libres']!=0]
    tarifas= pd.read_excel(r'''C:\Users\bpineda\Desktop\CD\Programacion diaria de containers\Reporte CD\reporte demurrage\tarifas.xlsx''')
    tarifas.set_index('Contenedor', inplace=True)
    tarifa_promedio_demurrage= tarifas.loc['general']['costo dia']
    tipo_cambio_dolar= tarifas.loc['tipo_cambio_USD']['tipo_cambio']
    tipo_cambio_euro= tarifas.loc['tipo_cambio_EUR']['tipo_cambio']
    serie_descrip=pd.Series(['Item',
                             'Período',
                             'Contenedores recibidos',
                             'Cajas recibidas',
                             'Unidades recibidas',
                             'Unidades predistribuidas recibidas',
                             'Contenedores demurrage',
                             'Días demurrage acumulado',
                             'Unidades sku con demurrage',
                             'Costo demurrage acumulado planificado CLP',
                             '$ CLP por unidad sku promedio'])
    
    # fechas en formato date
    maestro['Fecha']=maestro['Fecha'].apply(fecha_eta)
    maestro['Deadline']=maestro['Deadline'].apply(fecha_eta)
    # definición de variables de trabajo
    lista_fecha=[]
    fecha_consulta_1:str=str(input('Indique la fecha inicial en formato dd/mm/aaaa:'))
    lista_fecha.append(fecha_consulta_1)
    fecha_consulta_2:str=str(input('Indique la fecha final en formato dd/mm/aaaa:'))
    lista_fecha.append(fecha_consulta_2)
    contador=0
    for i in lista_fecha:
        dia=int(i.split("/")[0])
        mes=int(i.split("/")[1])
        ano=int(i.split("/")[2])
        fecha_consulta=datetime.date(year=ano, month=mes, day=dia)
        lista_fecha[contador]=fecha_consulta
        contador+=1
        
    periodo='Periodo consultado entre {f1} y {f2}'.format(f1=lista_fecha[0],f2=lista_fecha[1])
    maestro=maestro[(maestro['Fecha']>=lista_fecha[0])&(maestro['Fecha']<=lista_fecha[1])]
    maestro['dias_demu']=maestro['Fecha']-maestro['Deadline']
    maestro['dias_demu']=maestro['dias_demu'].apply(lambda x: x.days)
    maestro['dias_demu']=maestro['dias_demu'].apply(lambda x: 0 if x<0 else x )
    maestro_planificado=maestro['Container'].nunique()
    maestro_cajas_planificado=maestro['Cant cajas'].sum()
    maestro_unids_planificado=maestro['Cant productos'].sum()
    maestro_unids_predist_planificado=maestro['Cant predist'].sum()
    maestro_demurrage=maestro[maestro['dias_demu']>0]['Container'].nunique()
    maestro_dias_demurrage=maestro['dias_demu'].sum()
    maestro_unidades_demurrage=maestro[maestro['dias_demu']>0]['Cant productos'].sum()
    costo_demurrage=math.ceil(maestro_dias_demurrage*tarifa_promedio_demurrage*tipo_cambio_dolar)
    try:
        costo_marginal_demu_unidad=round(costo_demurrage/maestro_unidades_demurrage,3)
    except:
        costo_marginal_demu_unidad=0
    serie_fecha=pd.Series(['Valor',
                           periodo,
                           maestro_planificado,
                           maestro_cajas_planificado,
                           maestro_unids_planificado,
                           maestro_unids_predist_planificado,
                           maestro_demurrage,
                           maestro_dias_demurrage,
                           maestro_unidades_demurrage,
                           costo_demurrage,
                           costo_marginal_demu_unidad])
    
    serie_descrip=pd.concat([serie_descrip,serie_fecha],axis=1)
    serie_descrip.to_excel(r'C:\Users\bpineda\Desktop\reporte_demurrage.xlsx', index=False, header=False)
    return 0
reporte_demurrage_periodo()

# =============================================================================
# Reporte de demurrage según una planificación semanal
# =============================================================================
'''Se necesita el maestro de contenedores para planificar, una tabla con tarifas/
costos diarios marginales promedio por contenedor'''

    # definición de fecha
    hoy=datetime.date.today()
    # Preparación de archivos de trabajo. Maestro de contenedores para planificar
    maestro= pd.read_csv('''C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\maestro_por_contenedores_gby.csv''',sep=';')
    try:
        plano= pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\plano\\rpt_plano_bodega_compra.csv', sep=',')
        if plano.shape[1] == 1:
            plano= pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\plano\\rpt_plano_bodega_compra.csv', sep=';')
    except:
        plano= pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\plano\\rpt_plano_bodega_compra.csv', sep='\t')
        if plano.shape[1] == 1:
            plano= pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\plano\\rpt_plano_bodega_compra.csv', sep=';')
    tarifas= pd.read_excel(r'''C:\Users\bpineda\Desktop\CD\Programacion diaria de containers\Reporte CD\reporte demurrage\tarifas.xlsx''')
    maestro=maestro[['Container','ETA','Dias libres','Fecha','Fecha disp','Cant productos','Cant predist','Cant cajas','Deadline']]
    tarifas.set_index('Contenedor', inplace=True)
    tarifa_promedio_demurrage= tarifas.loc['general']['costo dia']
    tipo_cambio_dolar= tarifas.loc['tipo_cambio_USD']['tipo_cambio']
    tipo_cambio_euro= tarifas.loc['tipo_cambio_EUR']['tipo_cambio']
    serie_descrip=pd.Series(['Item/Fecha',
                             'Contenedores planificados',
                             'Cajas planificadas',
                             'Unidades planificadas',
                             'Unidades predistribuidas planificadas',
                             'Contenedores pendientes sin programar',
                             'Contenedores disponibles dentro de 7 días',
                             'Días demurrage acumulado planificado',
                             'Costo demurrage acumulado planificado CLP',
                             'Contenedores con demurrage no planificado',
                             'Cajas acumuladas pendientes',
                             'Días demurrage acumulado no planificado',
                             'Costo demurrage acumulado no planificado CLP',
                             'Costo marginal demurrage promedio por unidad sku planificada CLP'])
    
    # fechas en formato date
    maestro['ETA']=maestro['ETA'].apply(fecha_eta)
    maestro['Fecha']=maestro['Fecha'].apply(fecha_eta)
    maestro['Deadline']=maestro['Deadline'].apply(fecha_eta)
    maestro['Fecha disp']=maestro['Fecha disp'].apply(fecha_eta)
    plano[dic_campos_plano['eta_plano']]=plano[dic_campos_plano['eta_plano']].apply(fecha_eta)
    plano= plano[plano[dic_campos_plano['container_plano']].isin(
        maestro[maestro['Fecha']<maestro['Fecha'].max()]['Container'])]
    plano_dropduplicates=plano[
        (plano[dic_campos_plano['eta_plano']]>=maestro[dic_campos_planificacion['eta_col']].min())
        &(plano[dic_campos_plano['eta_plano']]<=(datetime.date.today()+datetime.timedelta(days = 15)))][
            [dic_campos_plano["invoice_plano"],
          dic_campos_plano['departamento_plano'],
          dic_campos_plano['linea_plano'],
          dic_campos_plano['estilo_plano'],
          dic_campos_plano['producto_plano'],
          dic_campos_plano['unidades_plano']]].drop_duplicates(keep = 'first')
    plano_gby=plano_dropduplicates[[dic_campos_plano['linea_plano'], dic_campos_plano['unidades_plano']]]
    plano_gby=plano_dropduplicates[[dic_campos_plano['linea_plano'], dic_campos_plano['unidades_plano']]].groupby([dic_campos_plano['linea_plano']]).sum()
    plano_gby.reset_index(inplace= True)
    plano_gby.sort_values(by= dic_campos_plano['unidades_plano'], ascending= False, inplace= True )
    # definición de variables de trabajo
    fechas_planificado=np.sort(maestro[~maestro['Fecha'].isna()]['Fecha'].unique())
    maestro['Fecha'].fillna(datetime.date(year=2050, month=1, day=1), inplace=True)
    maestro=maestro[maestro['Dias libres']!=0]
    
    
    # g=fechas_planificado[2]
    # Consulta si es por día o por un período acotado por fechas
    for fecha_consulta in fechas_planificado:
        # dia=int(fecha_consulta.split("/")[0])
        # mes=int(fecha_consulta.split("/")[1])
        # ano=int(fecha_consulta.split("/")[2])
        # fecha_consulta=datetime.date(year=ano, month=mes, day=dia)
        maestro_planificado=maestro[maestro['Fecha']==fecha_consulta]
        
        # maestro_no_planificado=maestro[maestro['Fecha'].isna()]
        # Días acumulados de demurrage en la fecha de operación. Considera todos los contenedores
        #  de la programación.
        maestro_planificado['dias_demu']=maestro_planificado['Fecha']-maestro_planificado['Deadline']
        maestro_planificado['dias_demu']=maestro_planificado['dias_demu'].apply(lambda x: x.days)
        maestro_planificado['dias_demu']=maestro_planificado['dias_demu'].apply(lambda x: 0 if x<0 else x )
        maestro_no_planificado=maestro[(maestro['Fecha']>fecha_consulta)]
        maestro_no_planificado['dias_demu']=fecha_consulta-maestro['Deadline']
        maestro_no_planificado['dias_demu']=maestro_no_planificado['dias_demu'].apply(lambda x: x.days+1)
        maestro_no_planificado['dias_demu']=maestro_no_planificado['dias_demu'].apply(lambda x: 0 if x<0 else x )
        maestro_restante_semana=maestro[(maestro['Fecha disp']<=fecha_consulta)&(maestro['Fecha']>fecha_consulta)]
        contenedores_restante_semana=maestro_restante_semana['Container'].nunique()
        contenedores_siete_dias=maestro_no_planificado[(maestro_no_planificado['Fecha disp']<=fecha_consulta+datetime.timedelta(days=7))&(maestro_no_planificado['Fecha disp']>fecha_consulta)]['Container'].nunique()
        contenedores_fecha_planif=maestro_planificado['Container'].nunique()
        cajas_planificado=maestro_planificado['Cant cajas'].sum()
        unidades_planificado=maestro_planificado['Cant productos'].sum()
        unidades_predist_planificado=maestro_planificado['Cant predist'].sum()
        dias_demurrage_planificado=maestro_planificado['dias_demu'].sum()
        costo_demurrage_planificado=math.ceil(dias_demurrage_planificado*tarifa_promedio_demurrage*tipo_cambio_dolar)
        if fecha_consulta==max(fechas_planificado):
            contenedores_fecha_no_planif=0
            cajas_no_planificado=0
            dias_demurrage_no_planificado=0
            costo_demurrage_no_planificado=0        
        else:
            contenedores_fecha_no_planif=maestro_no_planificado[maestro_no_planificado['dias_demu']>0]['Container'].nunique()
            cajas_no_planificado=maestro_no_planificado[(maestro_no_planificado['Fecha disp']<=fecha_consulta)]['Cant cajas'].sum()
            dias_demurrage_no_planificado=maestro_no_planificado['dias_demu'].sum()
            costo_demurrage_no_planificado=math.ceil(dias_demurrage_no_planificado*tarifa_promedio_demurrage*tipo_cambio_dolar)        
        try:
            costo_marginal_demu_unidad=round(costo_demurrage_planificado/maestro_planificado['Cant productos'].sum(),3)
        except:
            costo_marginal_demu_unidad=0
        serie_fecha=pd.Series([fecha_consulta,
                               contenedores_fecha_planif,
                               cajas_planificado,
                               unidades_planificado,
                               unidades_predist_planificado,
                               contenedores_restante_semana,
                               contenedores_siete_dias,
                               dias_demurrage_planificado,
                               costo_demurrage_planificado,
                               contenedores_fecha_no_planif,
                               cajas_no_planificado,
                               dias_demurrage_no_planificado,
                               costo_demurrage_no_planificado,
                               costo_marginal_demu_unidad])
        
        serie_descrip=pd.concat([serie_descrip,serie_fecha],axis=1)

maestro_planificado.to_excel(r'C:\Users\bpineda\Desktop\reporte demurrage\reporte_demurrage_m_planif.xlsx', index=False, header=True)
serie_descrip.to_excel(r'C:\Users\bpineda\Desktop\reporte demurrage\reporte_planificación_demurrage.xlsx', index=False, header=False)
plano_gby.to_excel(r'C:\Users\bpineda\Desktop\reporte demurrage\lineas_planificacion.xlsx', index= False)

del serie_descrip, maestro_planificado    
# =============================================================================
#Preguntas 
# =============================================================================
# dia_ref='Wednesday'
dia_ref='Tuesday'
# martes más cercano o día de referencia para el reporte
count=0
while True:
    diareporte=datetime.date.today()
    diamenos=datetime.timedelta(days=count)
    diareporte= diareporte-diamenos
    if diareporte.strftime('%A') == dia_ref:
        break
    else:
        count+=1
dia1= diareporte- datetime.timedelta(days=8)
dia2= diareporte- datetime.timedelta(days=2)
# Carga de programaciones
maestro_programaciones= carga_maestro_programaciones()
# Carga maestro planificacion
maestro_contenedores= pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\maestro_por_contenedores_gby.csv',sep=',')
if maestro_contenedores.shape[1]==1:
    maestro_contenedores= pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\maestro_por_contenedores_gby.csv',sep=';')
# Carga prioridades comerciales dia lunes
prior, aux0=carga_prior()
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
fechas.remove(max(fechas))
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
prior_ant= pd.read_excel(direccion_prioridades+'\\'+nombre_prior,sheet_name='Prioridad contenedor', header=0)
prior_ant=prior_ant[['PRIORIDAD','CONTAINER']]

reporte_demu=carga_archivo_disponibilidad()

maestro_programaciones['Fecha']= maestro_programaciones['Fecha'].apply(fecha_eta)
maestro_programaciones['ETA']= maestro_programaciones['ETA'].apply(fecha_eta)
maestro_programaciones['Deadline']=maestro_programaciones['Deadline'].apply(fecha_eta)
maestro_contenedores['ETA']=maestro_contenedores['ETA'].apply(fecha_eta)
maestro_contenedores['Deadline']=maestro_contenedores['Deadline'].apply(fecha_eta)

df_eta_fecha=maestro_programaciones[['Container','ETA','Fecha','Dias libres']]
# Def del campo lead time
df_eta_fecha['Dias entrega']=df_eta_fecha['Fecha']-df_eta_fecha['ETA']
# Ordenamiento de datos
df_eta_fecha.sort_values(by='Fecha',inplace=True)
# días de lead time
df_eta_fecha['Dias entrega']=df_eta_fecha['Dias entrega'].apply(lambda x: x.days)
#  media móvil simple de los 5 días anteiores por día
# df_eta_fecha['media movil dias_entrega']=(df_eta_fecha[['Dias entrega']].rolling(5).mean()).apply(lambda x: round(x,0))
# n de contendores por día
df_eta_fecha=df_eta_fecha[df_eta_fecha['Dias libres']!=0]
df_eta_fecha['n cont']=df_eta_fecha['Fecha'].apply(cant_cont)
# preparación del df
df_eta_fecha.reset_index(inplace=True)
df_eta_fecha.drop('index',inplace=True,axis=1)
# df2=df_eta_fecha[['Fecha','n cont','Dias entrega']]
df2=df_eta_fecha[df_eta_fecha['Dias libres']!=0][['Fecha','n cont',
                                                  'Dias entrega']]
 
df2=df2.groupby(['Fecha','n cont']).mean()
df2.to_csv(r'C:\Users\bpineda\Desktop\a.csv')
df2=pd.read_csv(r'C:\Users\bpineda\Desktop\a.csv')
df2['Fecha']=df2['Fecha'].apply(fecha_eta)
df2['Dias entrega']=df2['Dias entrega'].apply(lambda x:math.ceil(x))
df2['media movil dias_entrega']=(df2[[
    'Dias entrega']].rolling(5).mean()).apply(lambda x: round(x,0))

# Promedio de tiempo de descarga

def linear_reg_descarga():
    # Cruce entre registro de tiempo de descarga y programaciones pasadas
    conj_contenedores=pd.read_excel(r'C:\Users\bpineda\Desktop\CD\Programacion diaria de containers\Control de tiempos de descarga.xlsx',sheet_name='Control')
    conj_contenedores.dropna(thresh=5,inplace=True)
    conj_contenedores=conj_contenedores[conj_contenedores['tipo']=='FCL']
    conj_contenedores=conj_contenedores[['tipo','container','diferencia']]
    conj_contenedores['Cant cajas']=0
    conj_contenedores['Cant productos']=0
    conj_contenedores['Minutos']=0
    conj_cont_prog=pd.read_csv('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Manifiestos\\programaciones\\maestro_programaciones.csv')
    conj_cont_prog=conj_cont_prog[['Container','Cant cajas','Cant productos']]
    
    
    for i in conj_contenedores['container']:
        try:
            conj_contenedores.at[conj_contenedores[conj_contenedores['container']==i].index[0],'Cant cajas']=conj_cont_prog[conj_cont_prog['Container']==i]['Cant cajas'].unique()[0]
            conj_contenedores.at[conj_contenedores[conj_contenedores['container']==i].index[0],'Cant productos']=conj_cont_prog[conj_cont_prog['Container']==i]['Cant productos'].unique()[0]
        except:
            continue
    # función para obtener los minutos que se demoró la descarga
    def minutos(i):
        hora=int(i.hour)
        hora=hora*60
        minuto=int(i.minute)
        minutos_tot=hora+minuto
        return minutos_tot
    conj_contenedores['Minutos']=conj_contenedores['diferencia'].apply(minutos)
    
    for i in ['Cant productos','Minutos']:
        z_scores = stats.zscore(conj_contenedores[[i]],nan_policy='omit')
        abs_z_scores = np.abs(z_scores)
        filtered_entries = (abs_z_scores < 3).all(axis=1)
        while False in set(filtered_entries):
            conj_contenedores = conj_contenedores[filtered_entries]
            z_scores = stats.zscore(conj_contenedores[[i]])
            abs_z_scores = np.abs(z_scores)
            filtered_entries = (abs_z_scores < 3).all(axis=1)
        conj_contenedores = conj_contenedores[filtered_entries]
    lm= LinearRegression()
    X=conj_contenedores[['Cant cajas','Cant productos']]
    Y=conj_contenedores['Minutos']
    lm.fit(X,Y)
    lm.intercept_
    lm.coef_
    # cdf= pd.DataFrame(lm.coef_,X.columns,columns=['Coeff'])
    intercepto=lm.intercept_
    coef_caja=lm.coef_[0]
    coef_unid=lm.coef_[1]
    return conj_contenedores,intercepto,coef_caja,coef_unid

#  1000 CAJAS 13000 UNIDS
#  2000 CAJAS 40199 UNIDS
#  3000 CAJAS 44331 UNIDS
def tiempo_descarga():
    conj_contenedores,intercepto,coef_caja,coef_unid=linear_reg_descarga()
    cajas_1000=conj_contenedores[(conj_contenedores['Cant cajas']<=1000)&(conj_contenedores['Cant productos']!=0)]['Cant productos'].mean()
    cajas_2000=conj_contenedores[(conj_contenedores['Cant cajas']>1000)&(conj_contenedores['Cant cajas']<=2000)&(conj_contenedores['Cant productos']!=0)]['Cant productos'].mean()
    cajas_3000=conj_contenedores[(conj_contenedores['Cant cajas']>2000)&(conj_contenedores['Cant cajas']<=3000)&(conj_contenedores['Cant productos']!=0)]['Cant productos'].mean()
    dicc_cajas={'1000':[1000,math.ceil(cajas_1000)],'1001-2000':[2000,math.ceil(cajas_2000)],'2001-3000':[3000,math.ceil(cajas_3000)]}
    dicc_tiempo={}
    for i in dicc_cajas:
        cant_cajas, cant_unid= dicc_cajas[i][0], dicc_cajas[i][1]
        tiempo_estimado=math.ceil(intercepto+coef_caja*cant_cajas+coef_unid*cant_unid)
        tiempo_estimado=datetime.time(hour=int(tiempo_estimado//60),minute=int(tiempo_estimado%60))
        dicc_tiempo[i]=tiempo_estimado
    tiempo_1000=dicc_tiempo['1000']
    tiempo_2000=dicc_tiempo['1001-2000']
    tiempo_3000=dicc_tiempo['2001-3000']
    return dicc_cajas, tiempo_1000, tiempo_2000, tiempo_3000 

dicc_cajas, tiempo_1000, tiempo_2000, tiempo_3000=tiempo_descarga()
tiempo_1000='{horas}:{minutos}'.format(horas=tiempo_1000.hour,minutos=tiempo_1000.minute)
tiempo_2000='{horas}:{minutos}'.format(horas=tiempo_2000.hour,minutos=tiempo_2000.minute)
tiempo_3000='{horas}:{minutos}'.format(horas=tiempo_3000.hour,minutos=tiempo_3000.minute)
# df2.reset_index(inplace=True)
# Contenedores programados la semana anterior

# apertura de archivo maestro de programaciones
maestro_sem_ant=maestro_programaciones[(maestro_programaciones['Fecha']>=dia1) & (maestro_programaciones['Fecha']<=dia2)]
# Contenedores de la semana anterior

n_cont_sem_ant=maestro_programaciones[(maestro_programaciones['Fecha']>=dia1) & (maestro_programaciones['Fecha']<=dia2)]['Container'].nunique()
# Contenedores programados la semana anterior que se hayan traído con 2 o menos días de demurrage

n_cont_apuro=maestro_programaciones[(maestro_programaciones['Fecha']>=dia1) & (maestro_programaciones['Fecha']<=dia2) & (maestro_programaciones['Dias libres']!=0) & (maestro_programaciones['Deadline']-maestro_programaciones['Fecha']>=datetime.timedelta(days=0))&(maestro_programaciones['Deadline']-maestro_programaciones['Fecha']<=datetime.timedelta(days=2))]['Container'].nunique()
# Cantidad de cajas ingresadas la semana anterior al envío del reporte

cant_cajas_sem_ant=maestro_sem_ant['Cant cajas'].sum()
# Cantidad de unidades ingresadas la semana anterior al envío del reporte

unid_sem_ant=maestro_sem_ant['Cant productos'].sum()
# Cantidad de contenedores que deben llegar (FCL) si o si desde el lunes anterior al informe hasta 2 semanas

n_cont_2_sem=maestro_contenedores[(maestro_contenedores['Deadline']>=diareporte)&(maestro_contenedores['Deadline']<=diareporte+datetime.timedelta(days=15))&(maestro_contenedores['Dias libres']!=0)]['Container'].nunique()
# Cantidad de unidades que deberían llegar si o si desde el miércoles siguiente al reporte hasta 2 semanas

unid_2_sem=maestro_contenedores[(maestro_contenedores['Deadline']>=diareporte+datetime.timedelta(days=1))&(maestro_contenedores['Deadline']<=diareporte+datetime.timedelta(days=14))&(maestro_contenedores['Dias libres']!=0)]['Cant productos'].sum()
# Cantidad de cajas que deberían llegar si o si desde el miércoles siguiente al reporte hasta 2 semanas

cajas_2_sem=maestro_contenedores[(maestro_contenedores['Deadline']>=diareporte+datetime.timedelta(days=1))&(maestro_contenedores['Deadline']<=diareporte+datetime.timedelta(days=14))&(maestro_contenedores['Dias libres']!=0)]['Cant cajas'].sum()
# Cuantos contenedores prioritarios de comercial hay informados al lunes anterior al informe de la semana anterior

n_cont_comercial=prior_ant['CONTAINER'].nunique()
# Cuantos contenedores prioritarios de comercial hay informados al lunes anterior al informe

n_cont_comercial_act=prior['CONTAINER'].nunique()
# Cuantos contenedores prioritarios de comercial se trajeron la semana pasada

n_cont_sem_ant_comercial=maestro_programaciones[(maestro_programaciones['Container'].isin(prior_ant['CONTAINER']))&(maestro_programaciones['Fecha']>=dia1)&(maestro_programaciones['Fecha']<=dia2)]['Container'].nunique()
# cuantos contenedores de comercial faltan

n_cont_falta_comercial=prior[~prior['CONTAINER'].isin(maestro_programaciones['Container'])]['CONTAINER'].nunique()
# Cuantos contenedores FCL o LCL se trajeron la semana anterior

n_sem_ant_fcl=maestro_programaciones[(maestro_programaciones['Fecha']>=dia1) & (maestro_programaciones['Fecha']<=dia2)&(maestro_programaciones['Dias libres']!=0)]['Container'].nunique()
n_sem_ant_lcl=maestro_programaciones[(maestro_programaciones['Fecha']>=dia1) & (maestro_programaciones['Fecha']<=dia2)&(maestro_programaciones['Dias libres']==0)]['Container'].nunique()
# Flujo acumulado de contenedores del mes al dia del reporte

primer_dia_mes=datetime.date(year=diareporte.year,month=diareporte.month,day=1)
n_flujo_acum=maestro_programaciones[(maestro_programaciones['Fecha']>=primer_dia_mes) & (maestro_programaciones['Fecha']<=diareporte)]['Container'].nunique()
# Flujo acumulado de cajas del mes

cajas_flujo_acum=maestro_programaciones[(maestro_programaciones['Fecha']>=primer_dia_mes) & (maestro_programaciones['Fecha']<=diareporte)]['Cant cajas'].sum()
# Unidades acumuladas mes actual del reporte hasta el dia del reporte

unid_flujo_acum=maestro_programaciones[(maestro_programaciones['Fecha']>=primer_dia_mes) & (maestro_programaciones['Fecha']<=diareporte)]['Cant productos'].sum()
# Contenedores mes anterior

if diareporte.month==1:
    primer_dia_mes_anterior= datetime.date(year=diareporte.year-1,month=12,day=1)
    ultimo_dia_mes_anterior=datetime.date(year=diareporte.year,month=diareporte.month,day=1)-datetime.timedelta(days=1)
else:
    primer_dia_mes_anterior= datetime.date(year=diareporte.year,month=diareporte.month-1,day=1)
    ultimo_dia_mes_anterior=datetime.date(year=diareporte.year,month=diareporte.month,day=1)-datetime.timedelta(days=1)

cont_mes_ant=maestro_programaciones[(maestro_programaciones['Fecha']>=primer_dia_mes_anterior)&(maestro_programaciones['Fecha']<=ultimo_dia_mes_anterior)]['Container'].nunique()
# Cajas mes anterior

cajas_mes_ant=maestro_programaciones[(maestro_programaciones['Fecha']>=primer_dia_mes_anterior)&(maestro_programaciones['Fecha']<=ultimo_dia_mes_anterior)]['Cant cajas'].sum()
# Unidades mes anterior

unid_mes_ant=maestro_programaciones[(maestro_programaciones['Fecha']>=primer_dia_mes_anterior)&(maestro_programaciones['Fecha']<=ultimo_dia_mes_anterior)]['Cant productos'].sum()
# Cuantos contenedores FCL y LCL, que no se han programado tienen ETA cumplida

n_cont_eta_cumplida= maestro_contenedores[maestro_contenedores['ETA']<=diareporte]['Container'].nunique()
# cuantas cajas FCL y LCL tienen ETA cumplida

n_cajas_eta_cumplida=maestro_contenedores[maestro_contenedores['ETA']<=diareporte]['Cant cajas'].sum()
# cuantas unidaddes FCL y LCL tienen ETA cumplida y no se han programado

n_unid_eta_cumplida=maestro_contenedores[maestro_contenedores['ETA']<=diareporte]['Cant productos'].sum()
# Cuantos contenedores FCL se han traído con Demurrage el mes actual hasta la fecha del reporte

n_cont_demu= maestro_programaciones[(maestro_programaciones['Dias libres']!=0)&(maestro_programaciones['Fecha']>=primer_dia_mes)&(maestro_programaciones['Fecha']<=diareporte)&(maestro_programaciones['Deadline']-maestro_programaciones['Fecha']<=datetime.timedelta(days=-1))]['Container'].nunique()
# Cuantos contenedores se trajeron con Demurrage el mes anterior

n_cont_demu_mes_ant=maestro_programaciones[(maestro_programaciones['Dias libres']!=0)&(maestro_programaciones['Fecha']>=primer_dia_mes_anterior)&(maestro_programaciones['Fecha']<=ultimo_dia_mes_anterior)&(maestro_programaciones['Deadline']-maestro_programaciones['Fecha']<=datetime.timedelta(days=-1))]['Container'].nunique()
# Contenedores que llegan a puertos chilenos a partir de la fecha del envío de reporte

cont_por_llegar= maestro_contenedores[maestro_contenedores['ETA']>diareporte]['Container'].nunique()
# Cajas que llegan a puertos chilenos desde la fecha del reporte en adelante

cajas_por_llegar= maestro_contenedores[maestro_contenedores['ETA']>diareporte]['Cant cajas'].sum()
# Unidades que llegan a puertos chilenos desde la fecha del reporte en adelante

unid_por_llegar= maestro_contenedores[maestro_contenedores['ETA']>diareporte]['Cant productos'].sum()
# Contenedores informados disponibles el lunes anterior al reporte
try:
    cont_disp=reporte_demu[(~reporte_demu[dic_campos_disponibilidad['entrega_puerto']].isnull())&(~reporte_demu[dic_campos_disponibilidad['nro_contenedor']].isin(maestro_programaciones['Container']))][dic_campos_disponibilidad['nro_contenedor']].nunique()
except:
    cont_disp=reporte_demu[(~reporte_demu[dic_campos_disponibilidad['entrega_puerto']].isnull())&(~reporte_demu[dic_campos_disponibilidad['nro_contenedor']].isin(maestro_programaciones['Container']))][dic_campos_disponibilidad['nro_contenedor']].nunique()
# Contenedores informados disponibles de comercial el lunes anterior al reporte
try:
    cont_comer_disp=reporte_demu[(~reporte_demu[dic_campos_disponibilidad['entrega_puerto']].isnull())&(reporte_demu[dic_campos_disponibilidad['nro_contenedor']].isin(prior['CONTAINER']))&(~reporte_demu[dic_campos_disponibilidad['nro_contenedor']].isin(maestro_programaciones['Container']))][dic_campos_disponibilidad['nro_contenedor']].nunique()
except:
    cont_comer_disp=reporte_demu[(~reporte_demu[dic_campos_disponibilidad['entrega_puerto']].isnull())&(reporte_demu[dic_campos_disponibilidad['nro_contenedor']].isin(prior['CONTAINER']))&(~reporte_demu[dic_campos_disponibilidad['nro_contenedor']].isin(maestro_programaciones['Container']))][dic_campos_disponibilidad['nro_contenedor']].nunique()
# Cargas LCL con ETA cumplida y que no han sido programadas, por lo tanto están a la espera.

cant_lcl_espera=maestro_contenedores[(maestro_contenedores['Dias libres']==0)&(maestro_contenedores['ETA']<diareporte)]['Container'].nunique()
# Cajas LCL en espera

cajas_lcl_espera=maestro_contenedores[(maestro_contenedores['Dias libres']==0)&(maestro_contenedores['ETA']<diareporte)]['Cant cajas'].sum()
# Unidades LCL en espera

unid_lcl_espera=maestro_contenedores[(maestro_contenedores['Dias libres']==0)&(maestro_contenedores['ETA']<diareporte)]['Cant productos'].sum()
# Dias de demora de un contenedor desde que llega a puertos chilenos hasta que es programado en CD.

len(set(prior['CONTAINER']))
dias_demora=int(df2.tail(1)['media movil dias_entrega'].iloc[0])
# Días de demora de un contenedor desde que llega a puertos chilenos hasta que es desaduanado

dias_demora_aduana=disponibilidad_eta_disp(reporte_demu, maestro_programaciones)

# Definicion de las preguntas a responder

dicc_preg={5:n_cont_sem_ant,8:n_cont_apuro,9:cant_cajas_sem_ant,10:unid_sem_ant,11:n_cont_2_sem,12:cajas_2_sem,
           13:unid_2_sem,14:n_cont_comercial,15:n_cont_sem_ant_comercial,
           6:n_sem_ant_fcl,7:n_sem_ant_lcl,18:n_flujo_acum,19:cajas_flujo_acum,
           20:unid_flujo_acum,1:cont_mes_ant,2:cajas_mes_ant,3:unid_mes_ant,
           16:n_cont_comercial_act,
           21:n_cont_eta_cumplida,22:n_cont_demu,4:n_cont_demu_mes_ant,23:cont_por_llegar,24:cont_disp,
           17:n_cont_falta_comercial,25:cont_comer_disp,26:cant_lcl_espera,27:cajas_lcl_espera,
           28:unid_lcl_espera,29:n_cajas_eta_cumplida,30:n_unid_eta_cumplida,31:cajas_por_llegar,32:unid_por_llegar,33:dias_demora,
           34:dias_demora_aduana
           }

# Elaboracion de diccionario con las preguntas y respuestas a responder
meses={1:'Enero',2:'Febrero',3:'Marzo',4:'Abril',5:'Mayo',6:'Junio',7:'Julio',8:'Agosto',9:'Septiembre',10:'Octubre',11:'Noviembre',12:'Diciembre'}

columnas=['Contenedores','Cajas','Unidades','Total','Comercial']
emp_transporte='embarcadores'
if diareporte.month==1:
    p1='Recepción mes de {mes}'.format(mes=meses[12])
    p2='Contenedores con demurrage el mes de {mes}'.format(mes=meses[12])
else:
    p1='Recepción mes de {mes}'.format(mes=meses[(diareporte.month)-1])
    p2='Contenedores con demurrage el mes de {mes}'.format(mes=meses[(diareporte).month-1])
p3='Recepción de la semana del {dia1} al {dia2}'.format(dia1=(diareporte-datetime.timedelta(days=8)).day,dia2=(diareporte-datetime.timedelta(days=2)).day,mes_actual=meses[(diareporte-datetime.timedelta(days=2)).month])
p10='Recepción acumulada de {mes_actual}'.format(mes_actual=meses[(diareporte).month])
p5='Demurrage dentro de 2 semanas'
p14='Cargas LCL en espera a desconsolidación'
p12='Contenedores en puertos Chilenos'
p11='Contenedores con demurrage acumulado de {mes_actual}'.format(mes_actual=meses[(diareporte).month])
p4='Contenedores semana anterior con 2 días o menos para el Demurrage'
p6='Contenedores de comercial informados el {dia} de {mes}'.format(dia=(diareporte-datetime.timedelta(days=8)).day,mes=meses[(diareporte-datetime.timedelta(days=8)).month])
p7='Contenedores de comercial satisfechos la semana del {dia1} al {dia2}'.format(dia1=(diareporte-datetime.timedelta(days=8)).day,dia2=(diareporte-datetime.timedelta(days=2)).day,mes_actual=meses[(diareporte-datetime.timedelta(days=2)).month])
p8='Contenedores de comercial informados el {dia}'.format(dia=(diareporte-datetime.timedelta(days=1)).day,mes=meses[(diareporte-datetime.timedelta(days=1)).month])
p9='Contenedores de comercial por recepcionar'
p13='Contenedores en tránsito a puertos Chilenos'
p15='Contenedores disponibles '+emp_transporte
p16='Dias promedio entre fecha de llegada a puertos Chilenos y llegada CD'
p17='Dias promedio entre fecha de llegada a puertos Chilenos y nacionalización'
p18='Tiempo promedio en descargar menos de 1000 cajas'
p19='Tiempo promedio en descargar entre 1001 y 2000 cajas'
p20='Tiempo promedio en descargar entre 2001 y 3000 cajas'


dp1='Es la cantidad de recepciones de contenedores que se realizaron el mes anterior al mes en que se envía este informe'
dp2='Es la cantidad de contenedores que fueron recepcionados con demurrage en alguno de los CD de Tricot el mes anterior al mes en que se envía este informe'
dp3='Es la cantidad de contenedores que fueron recepcionados en alguno de los CD de Tricot la semana anterior a la que se envía este informe'
dp4='Es la cantidad de contenedores que fueron recepcionados en alguno de los CD de Tricot la semana anterior a la que se envía este informe y que faltaban 2 días o menos para que se cumplieran los días libres de Demurrage'
dp5='Es la cantidad de contenedores que cumplirán sus días libres de demurrage en alguno de los días dentro del periodo comprendido entre el día siguiente a este informe y 14 días más adelante'
dp6='Es la cantidad de contenedores relacionados a las prioridades del Departamento comercial de Tricot que fueron informados la semana anterior al envío de este informe'
dp7='Es la cantidad de contenedores relacionados a las prioridades del Departamento comercial de Tricot que se logró satisfacer la semana anterior al envío de este informe'
dp8='Es la cantidad de contenedores relacionados a las prioridades del Departamento comercial de Tricot que fueron informados, aproximadamente el lunes de la semana en que se envío este informe'
dp9='Es la cantidad de contenedores de comercial que faltan por recepcionar'
dp10='Es la cantidad de contenedores, y su carga, que han sido recepcionados en algún CD de Tricot desde el día 1 del mes actual al día en que se envía este informe'
dp11='Es la cantidad de contenedores que se han recepcionado fuera de los días libres de demurrage el mes actual'
dp12='Es la cantidad de contenedores, y su carga, cuya nave que los transporta ya tiene una fecha de llegada/ETA menor que la fecha en que se manda este informe, por lo tanto se dice que están en puertos Chilenos'
dp13='Es la cantidad de contenedores, y su carga, cuya nave que los transporta tiene una fecha de llegada/ETA mayor que la fecha en que se manda este informe, por lo tanto, se dice que está en tránsito'
dp14='Es la cantidad de contenedores, y su carga, que cuya modalidad de transporte requiere la espera de un aviso de desconsolidación para ser programada a algún CD de Tricot'
dp15='Es la cantidad de contenedores que según la información de la empresa de transporte, están disponibles'
dp16='Es la cantidad de días en promedio de los últimos 5 días, entre la fecha de llegada de las cargas FCL a puertos Chilenos y la llegada a CD'
dp17='Es la cantidad de días en promedio de los últimos 21 días de entregas, entre la fecha de llegada de las cargas FCL a puertos Chilenos y la nacionalización de la carga'
dp18='Es la cantidad de tiempo en promedio que toma descargar un contenedor con menos de 1000 cajas dado un promedio de unidades contenidas en cargas de menos de 1000 cajas'
dp19='Es la cantidad de tiempo en promedio que toma descargar un contenedor que contiene entre 1001 y 2000 cajas dado un promedio de unidades contenidas en contenedores con este rango de carga en cajas'
dp20='Es la cantidad de tiempo en promedio que toma descargar un contenedor que contiene entre 2001 y 3000 cajas dado un promedio de unidades contenidas en contenedores con este rango de carga en cajas'

desc_preg={}
for i in range(1,20):
    tupla='dict([(p'+str(i)+',dp'+str(i)+')])'
    desc_preg.update(eval(tupla))
#df de la descripcion de las preguntas que responde este informe
df_descrip=pd.DataFrame.from_dict(data=desc_preg,orient='index',columns=['Interpretación'])

# Descripción de items
vacio='----'
pregun={p1:[dicc_preg[1],dicc_preg[2],dicc_preg[3],vacio,vacio],
        p3:[dicc_preg[5],dicc_preg[9],dicc_preg[10],vacio,vacio],
        p10:[dicc_preg[18],dicc_preg[19],dicc_preg[20],vacio,vacio],
        p5:[dicc_preg[11],dicc_preg[12],dicc_preg[13],vacio,vacio],
        p14:[dicc_preg[26],dicc_preg[27],dicc_preg[28],vacio,vacio],
        p12:[dicc_preg[21],dicc_preg[29],dicc_preg[30],vacio,vacio],
        p13:[dicc_preg[23],dicc_preg[31],dicc_preg[32],vacio,vacio],
        p2:[dicc_preg[4],vacio,vacio,vacio,vacio],
        p11:[dicc_preg[22],vacio,vacio,vacio,vacio],
        p4:[dicc_preg[8],vacio,vacio,vacio,vacio],
        p6:[dicc_preg[14],vacio,vacio,vacio,vacio],
        p7:[dicc_preg[15],vacio,vacio,vacio,vacio],
        p8:[dicc_preg[16],vacio,vacio,vacio,vacio],
        p9:[dicc_preg[17],vacio,vacio,vacio,vacio],
        p15:[vacio,vacio,vacio,dicc_preg[24],dicc_preg[25]],
        p16:[vacio,vacio,vacio,dicc_preg[33],vacio],
        p17:[vacio,vacio,vacio,dicc_preg[34],vacio],
        p18:[vacio,dicc_cajas['1000'][0],dicc_cajas['1000'][1],tiempo_1000,vacio],
        p19:[vacio,dicc_cajas['1001-2000'][0],dicc_cajas['1001-2000'][1],tiempo_2000,vacio],
        p20:[vacio,dicc_cajas['2001-3000'][0],dicc_cajas['2001-3000'][1],tiempo_3000,vacio]
        }
df=pd.DataFrame.from_dict(pregun,orient='index',columns=columnas)
# Configuración del diseño del reporte
df.to_excel(r'C:\Users\bpineda\Desktop\CD\Programacion diaria de containers\Reporte CD\temp\df.xlsx')
df_descrip.to_excel(r'C:\Users\bpineda\Desktop\CD\Programacion diaria de containers\Reporte CD\temp\df1.xlsx')
df=pd.read_excel(r'C:\Users\bpineda\Desktop\CD\Programacion diaria de containers\Reporte CD\temp\df.xlsx')
df_descrip=pd.read_excel(r'C:\Users\bpineda\Desktop\CD\Programacion diaria de containers\Reporte CD\temp\df1.xlsx')
df.rename(columns={df.columns[0]:'Item'},inplace=True)
df_descrip.rename(columns={df_descrip.columns[0]:'Item'},inplace=True)
df_eta_cumplida=maestro_contenedores[maestro_contenedores['ETA']<=diareporte][['invoice','Container','ETA','Tipo_transporte','Deadline','Cant cajas','Cant productos']]
df_prox_semana=maestro_contenedores[maestro_contenedores['ETA']<=(dia2+datetime.timedelta(days=7))-datetime.timedelta(days=dias_demora-5)][['invoice','Container','ETA','Tipo_transporte','Deadline','Temporada','Productos','Cant cajas','Cant productos']]
df_naves= maestro_contenedores.copy()
df_naves['Contenedores']=int(1)
df_naves_trans= df_naves[['Nave','ETA','Contenedores','Cant cajas','Cant productos']].groupby(['Nave','ETA']).sum()
df_naves_trans.reset_index(inplace=True)
df_naves_trans['ETA']=df_naves_trans['ETA'].apply(fecha_eta)
df_naves_trans.sort_values(by='ETA',inplace=True)
df_naves_trans.reset_index(drop=True,inplace=True)
df_naves_trans['Fecha PM']=df_naves_trans['ETA']+datetime.timedelta(days=dias_demora)
df_naves_trans=df_naves_trans[df_naves_trans['ETA']<=diareporte+datetime.timedelta(days=10)]
# df_naves_trans=df_naves[df_naves['ETA']<=diareporte+datetime.timedelta(days=12)]

# =============================================================================
# Reporte de contenedores ETA, Deadline, Disponibilidad
# =============================================================================
hoy=datetime.date.today()
rango_fechas=pd.DataFrame(index=pd.date_range(start=hoy,end=hoy+datetime.timedelta(days=21)))
rango_fechas.reset_index(inplace=True)
rango_fechas.rename({'index':'Fechas operación'},axis=1, inplace=True)
rango_fechas=rango_fechas['Fechas operación'].apply(lambda x: x.date())
maestro_contenedor= maestro_contenedores.copy()
maestro_contenedor['Contenedores']=1
maestro_contenedor=maestro_contenedor[['Dias libres','ETA','Deadline', 'Fecha disp','Contenedores' , 'Cant cajas']]
maestro_contenedor['Deadline']=maestro_contenedor['Deadline'].apply(fecha_eta)
maestro_contenedor['Fecha disp']=maestro_contenedor['Fecha disp'].apply(fecha_eta)
maestro_contenedor['ETA']=maestro_contenedor['ETA'].apply(fecha_eta)
gby_eta= maestro_contenedor[['ETA', 'Contenedores', 'Cant cajas']].groupby(['ETA']).sum()
gby_dead= maestro_contenedor[maestro_contenedor['Dias libres']!=0][['Deadline', 'Contenedores', 'Cant cajas']].groupby(['Deadline']).sum()
gby_disp= maestro_contenedor[['Fecha disp', 'Contenedores', 'Cant cajas']].groupby(['Fecha disp']).sum()
rango_fechas=pd.merge(left=rango_fechas, right=gby_eta, how='left', left_on='Fechas operación', right_on='ETA')
rango_fechas.rename({'Contenedores':'Contenedores ETA', 'Cant cajas':'Cajas ETA'},axis=1, inplace=True)
rango_fechas=pd.merge(left=rango_fechas, right=gby_dead, how='left', left_on='Fechas operación', right_on='Deadline')
rango_fechas.rename({'Contenedores':'Contenedores Deadline', 'Cant cajas':'Cajas Deadline'},axis=1, inplace=True)
rango_fechas=pd.merge(left=rango_fechas, right=gby_disp, how='left', left_on='Fechas operación', right_on='Fecha disp')
rango_fechas.rename({'Contenedores':'Contenedores Disponibilidad', 'Cant cajas':'Cajas Disponibilidad'},axis=1, inplace=True)
rango_fechas.fillna(value=0, inplace=True)
# rango_fechas.to_excel(r'C:\Users\bpineda\Desktop\jola.xlsx', index=False, sheet_name='Reporte contenedores')

# =============================================================================
# 
# =============================================================================


df2=styleframe.StyleFrame(obj=df)
estilo=styleframe.Styler(font_color='00FF0000',font_size=11,font='Calibri',wrap_text=False,shrink_to_fit=False,bold=True)
estilo_head=styleframe.Styler(font_color='00000000',font_size=12,font='Calibri',wrap_text=False,shrink_to_fit=False,bold=True)
estilo_item=styleframe.Styler(bg_color='D3D3D3',font_color='000000FF',font_size=11,font='Arial',wrap_text=False,shrink_to_fit=False,bold=False)
df2.apply_column_style(cols_to_style=['Item'], styler_obj=estilo_item)
df2.apply_column_style(cols_to_style=[i for i in df.columns if i !='Item'],styler_obj=estilo)
df2.apply_headers_style(styler_obj=estilo_head,style_index_header=True)
df2.set_column_width(columns=['Item'], width=70)
df2.set_column_width(columns=[i for i in df.columns if i !='Item'], width=15)
df2.set_row_height(rows=df2.row_indexes, height=20)

df3=styleframe.StyleFrame(obj=df_descrip)
estilo=styleframe.Styler(font_color='00000080',font_size=11,font='Arial',wrap_text=False,shrink_to_fit=False,horizontal_alignment='left')
estilo_head=styleframe.Styler(font_color='00000000',font_size=12,font='Calibri',wrap_text=False,shrink_to_fit=False,bold=True)
estilo_item=styleframe.Styler(bg_color='D3D3D3',font_color='000000FF',font_size=12,font='Calibri',wrap_text=False,shrink_to_fit=False,bold=False)
df3.apply_column_style(cols_to_style=['Item'], styler_obj=estilo_item)
df3.apply_column_style(cols_to_style=[i for i in df_descrip.columns if i !='Item'],styler_obj=estilo)
df3.apply_headers_style(styler_obj=estilo_head,style_index_header=True)
df3.set_column_width(columns=['Item'], width=70)
df3.set_column_width(columns=[i for i in df_descrip.columns if i !='Item'], width=210)
df3.set_row_height(rows=df3.row_indexes, height=20)

df4=styleframe.StyleFrame(obj=df_eta_cumplida)
estilo=styleframe.Styler(font_color='00000080',font_size=11,font='Calibri',wrap_text=False,shrink_to_fit=False,horizontal_alignment='left')
estilo_numeros=styleframe.Styler(font_color='00FF0000',font_size=11,font='Calibri',wrap_text=False,shrink_to_fit=False,horizontal_alignment='center',)
estilo_head=styleframe.Styler(font_color='00000000',font_size=12,font='Calibri',wrap_text=False,shrink_to_fit=False,bold=True)
estilo_invoice=styleframe.Styler(bg_color='D3D3D3',font_color='000000FF',font_size=12,font='Calibri',wrap_text=False,shrink_to_fit=False,bold=False)
df4.apply_column_style(cols_to_style=['invoice'], styler_obj=estilo_invoice)
df4.apply_column_style(cols_to_style=[i for i in df_eta_cumplida.columns if i not in['invoice','Cant cajas','Cant productos']],styler_obj=estilo)
df4.apply_column_style(cols_to_style=['Cant cajas','Cant productos'],styler_obj=estilo_numeros)
df4.apply_headers_style(styler_obj=estilo_head,style_index_header=True)
df4.set_column_width(columns=['invoice'], width=40)
df4.set_column_width(columns=[i for i in df_eta_cumplida.columns if i !='invoice'], width=20)
df4.set_row_height(rows=df4.row_indexes, height=20)

df5=styleframe.StyleFrame(obj=df_prox_semana)
estilo=styleframe.Styler(font_color='00000080',font_size=11,font='Calibri',wrap_text=False,shrink_to_fit=False,horizontal_alignment='left')
estilo_numeros=styleframe.Styler(font_color='00FF0000',font_size=11,font='Calibri',wrap_text=False,shrink_to_fit=False,horizontal_alignment='center',)
estilo_head=styleframe.Styler(font_color='00000000',font_size=12,font='Calibri',wrap_text=False,shrink_to_fit=False,bold=True)
estilo_invoice=styleframe.Styler(bg_color='D3D3D3',font_color='000000FF',font_size=12,font='Calibri',wrap_text=False,shrink_to_fit=False,bold=False)
df5.apply_column_style(cols_to_style=['invoice'], styler_obj=estilo_invoice)
df5.apply_column_style(cols_to_style=[i for i in df_prox_semana.columns if i not in['invoice','Cant cajas','Cant productos']],styler_obj=estilo)
df5.apply_column_style(cols_to_style=['Cant cajas','Cant productos'],styler_obj=estilo_numeros)
df5.apply_headers_style(styler_obj=estilo_head,style_index_header=True)
df5.set_column_width(columns=['invoice'], width=40)
df5.set_column_width(columns=[i for i in df_prox_semana.columns if i !='invoice'], width=20)
df5.set_row_height(rows=df5.row_indexes, height=20)

df6=styleframe.StyleFrame(obj=df_naves_trans)
estilo=styleframe.Styler(font_color='00000080',font_size=11,font='Calibri',wrap_text=False,shrink_to_fit=False,horizontal_alignment='left')
estilo_numeros=styleframe.Styler(font_color='00FF0000',font_size=11,font='Calibri',wrap_text=False,shrink_to_fit=False,horizontal_alignment='center',)
estilo_head=styleframe.Styler(font_color='00000000',font_size=12,font='Calibri',wrap_text=False,shrink_to_fit=False,bold=True)
estilo_invoice=styleframe.Styler(bg_color='D3D3D3',font_color='000000FF',font_size=12,font='Calibri',wrap_text=False,shrink_to_fit=False,bold=False)
df6.apply_column_style(cols_to_style=['Nave'], styler_obj=estilo_invoice)
df6.apply_column_style(cols_to_style=[i for i in df_naves_trans.columns if i not in['Nave','ETA','Fecha PM']],styler_obj=estilo)
df6.apply_column_style(cols_to_style=['ETA','Fecha PM'],styler_obj=estilo_numeros)
df6.apply_headers_style(styler_obj=estilo_head,style_index_header=True)
df6.set_column_width(columns=['Nave'], width=40)
df6.set_column_width(columns=[i for i in df_naves_trans.columns if i !='Nave'], width=20)
df6.set_row_height(rows=df6.row_indexes, height=20)

df7=styleframe.StyleFrame(obj=rango_fechas)
estilo=styleframe.Styler(font_color='00000080',font_size=11,font='Calibri',wrap_text=False,shrink_to_fit=False,horizontal_alignment='left')
estilo_numeros=styleframe.Styler(font_color='00FF0000',font_size=11,font='Calibri',wrap_text=False,shrink_to_fit=False,horizontal_alignment='center',)
estilo_head=styleframe.Styler(font_color='00000000',font_size=12,font='Calibri',wrap_text=False,shrink_to_fit=False,bold=True)
estilo_invoice=styleframe.Styler(bg_color='D3D3D3',font_color='000000FF',font_size=12,font='Calibri',wrap_text=False,shrink_to_fit=False,bold=False)
df7.apply_column_style(cols_to_style=['Fechas operación'], styler_obj=estilo_invoice)
df7.apply_column_style(cols_to_style=[i for i in rango_fechas.columns if i not in['Fechas operación','Contenedores ETA','Contenedores Disponibilidad', 'Contenedores Deadline']],styler_obj=estilo)
df7.apply_column_style(cols_to_style=['Contenedores ETA','Contenedores Disponibilidad','Contenedores Deadline'],styler_obj=estilo_numeros)
df7.apply_headers_style(styler_obj=estilo_head,style_index_header=True)
df7.set_column_width(columns=['Fechas operación'], width=40)
df7.set_column_width(columns=[i for i in rango_fechas.columns if i !='Fechas operación'], width=20)
df7.set_row_height(rows=df6.row_indexes, height=20)

if len(str(diareporte.month))==1 and len(str(diareporte.day))==1:
    nombre_reporte='Reporte_semanal'+'_'+str(diareporte.year)+'0'+str(diareporte.month)+'0'+str(diareporte.day)+'.xlsx'
elif len(str(diareporte.month))==1 and len(str(diareporte.day))!=1:
    nombre_reporte='Reporte_semanal'+'_'+str(diareporte.year)+'0'+str(diareporte.month)+str(diareporte.day)+'.xlsx'
elif len(str(diareporte.month))!=1 and len(str(diareporte.day))==1:
    nombre_reporte='Reporte_semanal'+'_'+str(diareporte.year)+str(diareporte.month)+'0'+str(diareporte.day)+'.xlsx'
elif len(str(diareporte.month))!=1 and len(str(diareporte.day))!=1:
    nombre_reporte='Reporte_semanal'+'_'+str(diareporte.year)+str(diareporte.month)+str(diareporte.day)+'.xlsx'
    
ew= styleframe.StyleFrame.ExcelWriter('C:\\Users\\bpineda\\Desktop\\CD\\Programacion diaria de containers\\Reporte CD\\'+nombre_reporte) 
with ew as writer:
    df2.to_excel(writer,sheet_name='Reporte semanal',index=False)
    df4.to_excel(writer,sheet_name='Contenedores en puerto Chileno',index=False)
    df5.to_excel(writer,sheet_name='Recep probable semana sgte ')
    df6.to_excel(writer,sheet_name='ETAs')
    df7.to_excel(writer,sheet_name='Reporte contenedores ')
    df3.to_excel(writer,sheet_name='Definiciones',index=True)
ew.save()

# =============================================================================
# Reporte demurrage
# =============================================================================
def items_codigo(i):
    
    if type(i) is str:
        try:
            return int(i)
        except:
            return i

items_df= pd.read_excel('C:\\Users\\bpineda\\Desktop\\items.xlsx')
items_df['Cod Barra']= items_df['Cod Barra'].apply(items_codigo)
df_desp_items= pd.merge(df, items_df[['Codigo','TipoAlmac']], how='left', on='Codigo')

items_df=items_df[items_df['Codigo'].isin(df['Codigo'])]

