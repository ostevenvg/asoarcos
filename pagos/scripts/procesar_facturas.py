#!/usr/bin/env python
# coding: utf-8

# In[1]:


import os, sys, calendar
sys.path.append(os.getcwd().replace('notebooks','scripts'))
from utilities import *


# In[2]:


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

        
def gen_next_description(lastmonth, nmonths=1, return_month=False):
    if not isinstance(lastmonth,list):
        lastmonth = [dt.datetime.today().year-1,12]

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
    
    if return_month:
        return [y, m]
    else:
        return "MES AL COBRO " + ' Y '.join(d)

def normalize_fecha(in_date):
    m = re.match(r'(\d*)/(\d*)(/(\d*))?',in_date)
    if not m:
        m = re.match(r'(\d*)-(\d*)(-(\d*))?',in_date)
    if m:
        day = m.group(1)
        month = m.group(2)
        if m.group(4):
            year = m.group(4)
            if len(year) == 2:
                year = str(dt.datetime.today().year)[:2] + year
        else:
            year = dt.datetime.today().year
            if int(month) > dt.datetime.today().month: #It has to be from previous year
                year = year -1
                print(f"Fecha '{in_date}' en el excel bancos sin year, como el mes es mayor al actual, usando '{year}'")
            year = str(year)

        norm_date = '{}/{}/{}'.format(day,month,year)
        due_date = '{}/{}/{}'.format(calendar.monthrange(int(year), int(month))[1],month,year)
        return [norm_date, due_date]
    else:
        return [None, None]

def get_account(bank):
    accounts = {
        'bac': '10205 BAC S.J.906883251 Colones',
        'bcr': '10201 BCR 160198-9',
        'bn' : '10203 BNCR 100-01-000-216002-6'}
    return accounts[bank]
    
def month2num(x):
    return (int(x[0])-2000)*12+int(x[1])
    
def num2month(x):
    return [2000 + math.floor((x-1)/12), x-12*math.floor((x-1)/12)]

def get_description_months(description):
    out = list()
    if not isinstance(description, str):
        return out

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


def group_bill_df(df, from_month=''):
    df_ = df.copy()
    df_["mes_factura"]=df_['Date'].dt.strftime('%Y,%m').str.split(',')
    df_["mes_factura_n"] = df_["mes_factura"].apply(month2num)
    if from_month:
        df_ = df_[(df_['mes_factura_n']>=month2num(from_month))]

    #group by name (client), adding date,num,description as lists
    for col in ['Date', 'Num', 'Memo/Description', 'mes_factura_n']:
        df_[col] = df_[col].apply(lambda x: [x])
    df_ = df_.groupby(['Name']).sum()
    df_.insert(0, 'cliente', df_.index)
    
    #add extrasummary columns
    df_['ultima factura'] = df_['Memo/Description'].apply(lambda x: x[-1] if x else np.nan)
    df_['ultimo mes'] = df_['mes'].apply(lambda x: x[-1] if x else np.nan)
    df_['siguiente descripcion'] = df_['ultimo mes'].apply(lambda x: gen_next_description(x,1) if isinstance(x,list) else np.nan)

    return df_

def process_bill_excel(inputs_dir):
    df = gendf_from_excel_table(os.path.join(inputs_dir, "facturas_quickbooks.xlsx"),['Memo/Description'])
    df = df[1:][['Date','Num', 'Name', 'Memo/Description','Amount']].dropna(axis=0, how="all")
    
    #get mes
    df['mes'] = df['Memo/Description'].apply(get_description_months)
    
    #create original df copy
    df_allrows=df.copy(deep=True)
    df_allrows['Date']= pd.to_datetime(df_allrows['Date'], dayfirst=True)
    
    df = group_bill_df(df_allrows)
    return [df_allrows, df] 


# In[3]:


########### MAIN ############
#work directory
if is_interactive():
    #sys.argv = ['', r'C:\Users\villalta\Documents\Personal\repos\qb_test\Junio']
    sys.argv = ['', r"/Users/oscar/Documents/work/qb_example2/"]
if len(sys.argv) < 2:
    print("Error: Se debe indicar el directorio donde estan las entradas")
    exit(1)
tasks = ['gen_qb_csv', 'gen_report']
if len(sys.argv) == 3:
    tasks = [sys.argv[2]]
    
inputs_dir = os.path.abspath(sys.argv[1])
if not os.path.exists(inputs_dir):
    print("Error: El directorio con los archivos de entrada no existe: " + inputs_dir)
    exit(1)

outputs_dir = os.path.join(inputs_dir, "salidas")
if not os.path.exists(outputs_dir):
    os.makedirs(outputs_dir)

#Load configuration
config_path = os.path.join(inputs_dir, "config.xlsx")
if os.path.exists(config_path):
    config_excel = config_path
else:
    config_excel = "config.xlsx"

print("Leyendo la configuracion desde: " + config_excel)
df_client_config = pd.read_excel(config_excel, sheet_name="excepciones clientes").dropna(axis=1, how="all")
ignore_clients = list(df_client_config[df_client_config.ignorar == "si"]['cliente'])
    
#Load bills
print("Procesando facturas de Quickbooks")
[df_bills_allrows, df_bills] = process_bill_excel(inputs_dir)

#Check if there are bills with empty description
df_bad_bills = df_bills_allrows[df_bills_allrows['Memo/Description'].isna()]
if not df_bad_bills.empty:
    print("ERROR: las siguientes facturas tienen una descripcion vacia")
    for index, row in df_bad_bills[['Num', 'Name']].iterrows():
        print(' - Factura: {}  Cliente: {}'.format(row['Num'], row['Name']))
    exit(1)


# In[29]:


#Process bancos if found, to create file that will be imported to quickbooks
add_for_quickbooks = False

banks_file = os.path.join(outputs_dir, "bancos.xlsx")
if 'gen_qb_csv' in tasks and os.path.exists(banks_file ):
    print("Generando archivo CSV para Quickbooks")
    #get bancos and clientes
    df_bancos = pd.read_excel(banks_file , sheet_name='bancos')
    df_clientes = pd.read_excel(banks_file , sheet_name='clientes')
    #validate if clientes row is correct
    df_bad_clientes = df_bancos[(df_bancos['cliente'].notna()) & (df_bancos['num meses'].notna()) & 
                                (df_bancos['cliente'].isin(df_clientes['Customer']) == False)]
    if df_bad_clientes.empty == False:
        print("\nERROR: los siguientes clientes en el archivo bancos.xlsx no son validos. Se escribieron incorrectamente o hay mas de un cliente en una misma fila:\n")
        for client in list(df_bad_clientes['cliente']):
            print(' - {}'.format(client))
        exit(1)
    else:
        df_bancos_valid =  df_bancos.loc[(df_bancos['cliente'].notna()) & (df_bancos['num meses'].notna()) &
                           (df_bancos['num meses'] != 0)]
        df_bancos_valid = df_bancos_valid.merge(df_bills, how='left', on=['cliente'])
        
        #Add month offset when more than one payment from same client in one month
        df_bancos_valid.sort_values(['cliente','fecha'],inplace=True)
        df_bancos_valid.reset_index(inplace=True)
        for i in range(1, len(df_bancos_valid)):
            if df_bancos_valid.loc[i,'cliente'] == df_bancos_valid.loc[i-1,'cliente']:
                df_bancos_valid.at[i,'ultimo mes'] = gen_next_description(df_bancos_valid.loc[i-1,'ultimo mes'],
                                 df_bancos_valid.loc[i-1,'num meses'],return_month=True)
        
        df_bancos_valid['siguiente descripcion'] = df_bancos_valid.apply(lambda x: gen_next_description(x['ultimo mes'],x['num meses']), axis=1)
        df_bancos_valid.rename(columns={"descripcion": "detalle banco"}, inplace=True)
          
        #quickbooks format
        df_bancos_qb = df_bancos_valid.rename(columns={'cliente':'Customer', 
                                                       'siguiente descripcion':'ItemDescription', 
                                                       'credito':'ItemAmount'})
        df_bancos_qb.reset_index(inplace=True)
        
        next_bill_num = df_bills_allrows['Num'].astype(int).max() + 1
        print(next_bill_num)
        df_bancos_qb.insert(0,'InvoiceNo',df_bancos_qb.index+next_bill_num)
        df_bancos_qb['InvoiceDate'], df_bancos_qb['DueDate'] = zip(*df_bancos_qb['fecha'].map(normalize_fecha))
        df_bancos_qb['Memo'] = df_bancos_qb['ItemDescription']
        df_bancos_qb['ReferenceNo'] = df_bancos_qb['banco'] + ':' + df_bancos_qb['referencia'].astype(str)
        df_bancos_qb['Item(Product/Service)'] = 40
        df_qb_csv = df_bancos_qb.copy()
        
        #add patos product if any
        for index, row in df_bancos_qb.iterrows():
            row = row.to_dict()
            if row["agua patos"] == 1:
                row['ItemDescription'] = "Aporte extraordinario agua de los lagos"
                row['Item(Product/Service)'] = "41 Cuota extraordinaria"
                row['ItemAmount'] = 0
                df_qb_csv = pd.concat([df_qb_csv, pd.DataFrame.from_dict(row, orient='index').transpose()])
        df_qb_csv.sort_values(by=['InvoiceNo','Item(Product/Service)'], inplace=True)    
        
        #save bills import csv
        csv_columns = ['ReferenceNo', 'InvoiceNo', 'InvoiceDate', 'DueDate', 'Customer', 'Item(Product/Service)', 'ItemDescription', 'Memo', 'ItemAmount']
        qb_csv = os.path.join(outputs_dir,"qb_import.csv")
        df_qb_csv[csv_columns].to_csv(qb_csv,index=False)
        df_qb_csv[csv_columns].to_excel(qb_csv.replace('csv','xlsx'),index=False)
        print("Import CSV generado (archivo para importar las facturas a Quickbooks): {}".format(qb_csv))
        
        
        #paymets import format
        df_bancos_payimport = df_bancos_qb.rename(columns={ 
            'InvoiceNo': 'Txn ID',
            'InvoiceDate': 'Payment date',
            'ReferenceNo': 'Payment Ref Number'})
        df_bancos_payimport['Amount'] = df_bancos_payimport['ItemAmount']
        df_bancos_payimport['Total amount'] = df_bancos_payimport['Amount']
        #df_bancos_payimport['Deposit to account'] = 'Cash and cash equivalents'
        df_bancos_payimport['Deposit to account'] = df_bancos_payimport['banco'].apply(get_account)
        df_bancos_payimport['Txn type'] = 'Invoice'
        df_bancos_payimport['Payment method'] = 'Transference' 
        pay_columns = ['Customer', 'Deposit to account', 'Amount', 'Total amount',
                       'Txn ID', 'Txn type', 'Payment method', 'Payment date',
                       'Memo', 'Payment Ref Number']
        pay_excel = os.path.join(outputs_dir,"qb_payment_import.xlsx")
        df_bancos_payimport[pay_columns].to_excel(pay_excel,index=False)
        print("Excel de pagos generado (Archivo que se usa para importar los pagos a Quickbooks): {}".format(pay_excel))
      
        add_for_quickbooks = True

#Save excel
facturas_file = os.path.join(outputs_dir, "facturas.xlsx")
with pd.ExcelWriter(facturas_file) as writer:
    if add_for_quickbooks:
        df_bancos_valid.to_excel(writer, index=False, sheet_name='para_quickbooks')
    df_bills.to_excel(writer, index=False, sheet_name='historico')
        
print("Excel generado: {}".format(facturas_file))


# In[ ]:


if 'gen_report' in tasks:
    print("Generando reporte web")
    ############# REPORT ###########
    figs = []

    #asociados al dia
    df_bills_ = df_bills_allrows.copy()
    df_bills_ = df_bills_.explode('mes').dropna(axis=0, subset="mes")
    df_bills_["mes_factura"]=df_bills_['Date'].dt.strftime('%Y,%m').str.split(',')
    df_bills_["mes_factura_n"] = df_bills_["mes_factura"].apply(month2num)
    df_bills_["mes_n"] =  df_bills_["mes"].apply(month2num)

    paid_months = list()
    d_map = ['mes actual', 'mes anterior', 'hace dos meses', 'hace tres meses']
    for y in range(dt.datetime.today().year-2,dt.datetime.today().year):
        for m in range(1,13):
            month = [str(y), str(m)]
            m_s = '-'.join(month)
            m_n = month2num(month)
            for d in [0, 1, 2, 3]:
                c = df_bills_[(df_bills_['mes_factura_n'] <= m_n) & (df_bills_['mes_n'] == m_n-d)].shape[0]
                paid_months.append([m_s, d_map[d], c])

    df = pd.DataFrame(paid_months, columns = ['mes', 'al dia con respecto a', 'numero de asociados'])
    fig = px.bar(df, x='mes', y='numero de asociados', color='al dia con respecto a', barmode="group",
          title = "Número de asociados al día")
    figs.append(fig)

    #Dinero que ingresa al mes
    ###############################################################
    df_bills_2 = df_bills_allrows.copy()
    df_bills_2["mes_factura"]=df_bills_2['Date'].dt.strftime('%Y-%m')
    df_bills_2["tipo de pago"]=df_bills_2['mes'].apply(lambda x: "cuota condominal" if x else "otros")
    fig = px.histogram(df_bills_2, x = "mes_factura", y = "Amount", color="tipo de pago",
                       text_auto=True, title = "Monto recaudado")
    fig.update_layout(bargap=0.2)
    figs.append(fig)

    #Meses cancelados
    ###############################################################
    df_bills_1= df_bills.dropna(axis=0, subset="ultimo mes").explode('mes')
    df_bills_1['mes'] = df_bills_1['mes'].apply(lambda x: str(x[0])+'-'+str(x[1]))
    fig = px.histogram(df_bills_1, x = "mes", text_auto=True, title = "Mensualidades pagadas")
    fig.update_layout(bargap=0.2)
    figs.append(fig)


    #Clientes perdidos
    ###############################################################
    due_alarm = [3, 2] #meses pendientes, meses pendientes diferido
    start_month = [str(dt.datetime.today().year-1), str(dt.datetime.today().month)]
    current_month = [str(dt.datetime.today().year), str(dt.datetime.today().month)]
    month_diff = month2num(current_month) - month2num(start_month) + 1

    def only_positive(x):
        if x < 0:
            return 0
        return x
        
    def elap_months (start_date):
        month_diff = month2num(current_month) - month2num([start_date.year, start_date.month]) + 1
        return month_diff

    df_bills_3 = group_bill_df(df_bills_allrows, start_month)
    df_bills_3 = df_bills_3.dropna(axis=0, subset="ultimo mes")
    df_bills_3['meses pendientes'] = df_bills_3['ultimo mes'].apply(lambda x: only_positive(month2num(current_month) - month2num(x) + 1) )
    df_bills_3['meses pendientes diferido'] = df_bills_3['mes'].apply(lambda x: only_positive(month_diff - len(x)))
    #df_bills_3['cliente perdido'] = (df_bills_3['meses pendientes'] > due_alarm[0]) |  (df_bills_3['meses pendientes diferido'] > due_alarm[1])
    df_bills_3['ultimo pago'] = df_bills_3['Date'].apply(lambda x: x[-1])
    df_bills_3['meses perdido'] = df_bills_3['Date'].apply(lambda x: elap_months(x[-1]))
    df_bills_3['cliente perdido'] = (df_bills_3['meses pendientes'] > 3) | ((df_bills_3['meses pendientes'] > 0) & (df_bills_3['meses perdido']>3))
     
    #columns = ['Date', 'Memo/Description', 'mes', 'ultimo pago', 'meses pendientes', 'meses perdido']
    columns = ['ultimo pago', 'meses pendientes', 'meses perdido']
    lost_clients = df_bills_3[df_bills_3['cliente perdido'] & ~df_bills_3.index.isin(ignore_clients)][columns].to_html()
    
    #reporte final
    ###############################################################
    html_style = """
    <style>
    table {
      font-family: consola;
      border-collapse: collapse;
      width: 100%;
    }

    td, th {
      border: 1px solid #dddddd;
      text-align: left;
      padding: 8px;
    }

    tr:nth-child(even) {
      background-color: #dddddd;
    }
    </style>
    """

    reporte_file = os.path.join(outputs_dir,"reporte.html")
    with open(reporte_file, "w") as fh:
        fh.write(html_style)
        fh.write("<center><h1> ASOARCOS: reporte de mensualidades</h1></center>")
        for fig in figs:
            fh.write(fig.to_html())
        fh.write("<center><h3> Usuarios perdidos</h3></center>")
        fh.write("<b> Meses pendientes: </b> numero de meses reales que se deben</br>")
        fh.write(lost_clients)
    print("Reporte web generado: {}".format(reporte_file))

print("Fin del proceso.")


# In[ ]:




