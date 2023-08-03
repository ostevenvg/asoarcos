#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from utilities import *


# In[ ]:


############### Functions to get and generated paid months ################
MONTHS = ["NA", "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
          "JULIO", "AGOSTO", "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
SMONTHS = ["NA", "ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SET", "OCT", "NOV", "DIC"]

def get_months(months_list, year, get_range=False):    
    imonths = []
    for m in months_list:
        if m and not m.isspace():
            m = m.replace(' ','').replace(',','')
            if m in MONTHS:
                imonths.append(MONTHS.index(m))
            elif m in SMONTHS:
                imonths.append(SMONTHS.index(m))
            else:
                #print("Warn: month not identified:" + m + ".")
                continue
            
    if get_range and len(imonths) == 2:
        imonths = list(range(imonths[0], imonths[1]+1))
    
    out = []
    for m in imonths:
        out.append([year, m])
    return out

def gen_month_description(y, mlist):
    if len(mlist) > 1:
        return "DE {} A {} {}".format(MONTHS[mlist[0]],MONTHS[mlist[-1]],y)
    else:
        return "{} {}".format(MONTHS[mlist[0]],y)
    
def gen_next_description(lastmonth, nmonths=1):
    y = int(lastmonth[0])
    m = int(lastmonth[1])
    d = []
    
    mlist = list()
    for x in range(1, nmonths+1):
        m += 1
        if m == 13:
            if len(mlist) != 0:
                d.append(gen_month_description(y, mlist))
            y += 1
            m = 1
            mlist=[]
        mlist.append(m)
    d.append(gen_month_description(y, mlist))
    
    return "MES AL COBRO " + ' Y '.join(d)

def get_description_months(description): 
    out = list()
    d = description.upper()
    d = d.replace('.', ',')
    d = d.replace('-',', ')
    d = re.sub(r'(AAL|LA|AL|SL)','AL',d)
    d = re.sub(r'(COBRO|CORO|COBERO)','COBRO',d)
    d = re.sub(r', SE LE HIZO.*','',d)
    d = re.sub(r', ABONO A.*','',d)
    d = re.sub(r', MAS 4.*','',d)
    d = re.sub(r' \(.*','',d)
    d = re.sub(r'.*SEGURIDAD MES','MES',d)
    d = d.replace('2033', '2023')
    
    #single month match
    m = re.match(r'^\s*M?ES\s+AL\s+COBRO\s*(\w+)\s+(\d+)\s*$', d)
    if m:
        out += get_months([m.group(1)], m.group(2))

    #range matches
    m = re.match(r'^\s*MES\s+AL\s+COBRO\s*(DE\s+)?(\w+)\s+A\s+(\w+)\s+(\d+)\s*$', d)
    if m:
        out += get_months([m.group(2), m.group(3)], m.group(4), get_range=True)
    
    m = re.match(r'^\s*MES\s+AL\s+COBRO\s+((\w+,\s+)*)(\w+)\s+Y\s+(\w+)\s+(\d+)\s*$', d)
    if m:
        months = [m.group(3), m.group(4)]
        if m.group(1):
            months = m.group(1).split(', ') + months
        out += get_months(months, m.group(5))
    
    #try splitting sentence
    if not out and re.search(r'Y', d):
        d_ = d.split('Y')
        for i in range(len(d_)-1):
            out += get_description_months('Y'.join(d_[0:i+1]))
            out += get_description_months('MES AL COBRO ' + 'Y'.join(d_[i+1:]))
    
    #try replacing , with Y
    if not out and re.search(r',', d):
        out += get_description_months(d.replace(',','Y'))
    
    out.sort(reverse=False)
    
    out_ = list()
    for x in out:
        if x not in out_:
            out_.append(x)
    return out_


def process_bill_excel(inputs_dir):
    df = gendf_from_excel_table(inputs_dir + "/facturas_quickbooks.xlsx",['Memo/Description'])
    df = df[1:][['Date','Num', 'Name', 'Memo/Description','Amount']].dropna(axis=0, how="all")
    
    #get mes
    df['mes'] = df['Memo/Description'].apply(get_description_months)
    
    #create original df copy
    df_allrows=df.copy(deep=True)
    df_allrows['Date']= pd.to_datetime(df_allrows['Date'], dayfirst=True)
    
    #group by name (client), adding date,num,description as lists
    for col in ['Date', 'Num', 'Memo/Description']:
        df[col] = df[col].apply(lambda x: [x])
    df = df.groupby(['Name']).sum()
    df.insert(0, 'cliente', df.index)
    
    #add extrasummary columns
    df['ultima factura'] = df['Memo/Description'].apply(lambda x: x[-1] if x else np.nan)
    df['ultimo mes'] = df['mes'].apply(lambda x: x[-1] if x else np.nan)
    df['siguiente descripcion'] = df['ultimo mes'].apply(lambda x: gen_next_description(x,1) if isinstance(x,list) else np.nan)
    
    return [df_allrows, df] 


# In[ ]:


########### MAIN ############
work_dir=os.getcwd()
inputs_dir=work_dir+"/inputs"
if not os.path.exists(inputs_dir):
    print("Error: El directorio con los archivos de entrada no existe: " + inputs_dir)
    exit(1)
outputs_dir=work_dir+"/outputs"
if not os.path.exists(outputs_dir):
    os.makedirs(outputs_dir)

print("Procesando facturas de Quickbooks")
#Load bills
[df_bills_allrows, df_bills] = process_bill_excel(inputs_dir)


# In[ ]:


print("Generando facturas nuevas")
#Process bancos if found, to create file that will be imported to quickbooks
add_for_quickbooks = False
if os.path.exists(outputs_dir + "/bancos.xlsx"):
    #get bancos and clientes
    df_bancos = pd.read_excel(outputs_dir + "/bancos.xlsx", sheet_name='bancos')
    df_clientes = pd.read_excel(outputs_dir + "/bancos.xlsx", sheet_name='clientes')
    #validate if clientes row is correct
    df_bad_clientes = df_bancos[df_bancos['cliente'].isin(df_clientes['Customer']) == False]
    if not df_bad_clientes.empty:
        print("ERROR: los siguientes clientes no son validos, revise si se escribieron incorrectamente o sin son validos pero son mas de uno separelos en varias filas en el excel bancos.xlsx (un cliente por fila)")
        for client in list(df_bad_clientes['cliente']):
            print(' - {}'.format(client))
    else:
        #generate table to be imported in quickbooks
        df_bancos_valid =  df_bancos.loc[(df_bancos['cliente'].notna()) & (df_bancos['num casas'].notna()) & (df_bancos['num meses'].notna()) &
                  (df_bancos['num casas'] != 0) & (df_bancos['num meses'] != 0)]
        df_bancos_valid = df_bancos_valid.merge(df_bills, on=['cliente'])
        df_bancos_valid['siguiente descripcion'] = df_bancos_valid.apply(lambda x: gen_next_description(x['ultimo mes'],x['num meses']), axis=1)
        df_bancos_valid.rename(columns={"descripcion": "detalle banco"}, inplace=True)
        df_bancos_excel_colums = ['banco', 'referencia', 'detalle banco' ,'fecha', 'cliente', 'credito', 'siguiente descripcion']
        add_for_quickbooks = True
    
#Save excel
with pd.ExcelWriter("{}/facturas.xlsx".format(outputs_dir)) as writer:
    if add_for_quickbooks:
        df_bancos_valid[df_bancos_excel_colums].to_excel(writer, index=False, sheet_name='para_quickbooks')
    df_bills.to_excel(writer, index=False, sheet_name='historico')
        
print("Excel file generado: {}/facturas.xlsx".format(outputs_dir))


# In[ ]:


df_bancos[df_bancos['cliente'].isin(df_clientes['Customer']) == False]


# In[ ]:


print("Generando reporte web")
############# REPORT ###########
figs = []

#Meses cancelados
df_bills_= df_bills.dropna(axis=0, subset="ultimo mes").explode('mes')
df_bills_['mes'] = df_bills_['mes'].apply(lambda x: str(x[0])+'-'+str(x[1]))
fig = px.histogram(df_bills_, x = "mes", text_auto=True, title = "Meses cancelados")
fig.update_layout(bargap=0.2)
figs.append(fig)

#Dinero que ingresa al mes
df_bills_ = df_bills_allrows
df_bills_["mes_factura"]=df_bills_['Date'].dt.strftime('%Y-%m')
df_bills_["tipo de pago"]=df_bills_['mes'].apply(lambda x: "cuota condominal" if x else "otros")
fig = px.histogram(df_bills_, x = "mes_factura", y = "Amount", color="tipo de pago",
                   text_auto=True, title = "Monto recaudado")
fig.update_layout(bargap=0.2)
figs.append(fig)


with open("{}/reporte.html".format(outputs_dir), "w") as fh:
    for fig in figs:
        fh.write(fig.to_html())
print("Reporte web generado: {}/reporte.html".format(outputs_dir))
print("Fin del proceso.")

