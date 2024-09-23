#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import os, sys
sys.path.append(os.getcwd().replace('notebooks','scripts'))
from utilities import *


# In[ ]:


def get_casas_from_description(d):
    if not isinstance(d,str):
        return []
    casas = []
    d=d.upper()
    
    d = d.replace('CSA','CASA')
    m = re.match(r'.*CASAS?\s*([\d\-Y,\s\/_]*)',d)
    if m:
        casas = m.group(1)
        if casas:
            casas = re.sub('|_|\/|\s','',casas)
            casas = casas.replace('Y',',').split(',')
            
    #cases like '1219-1220' vs '1-219'
    casas_ = []
    for c in casas:
        if re.match(r'\d\-\d.*',c):
            casas_.append(re.sub('-','',c))
        else:
            casas_ += c.split('-')
    
    #expand casa number
    casas = []
    for c in casas_:
        if len(c) == 1:
            casas.append('100' + c)
        elif len(c) == 2:
            casas.append('10' + c)
        elif len(c) == 3:
            casas.append('1' + c)
        else:
            casas.append(c[:4])
        
    return casas

def get_casa_from_customer(c):
    out = [False, None, np.nan]
    
    if isinstance(c, str):
        c = c.upper()
        #divide casa and name. Example: Casa 1209 Jaime Barrientos"
        m = re.match(r'CASA (\d+(-[A,B,C,D,E])?([-/]\d+)?)\s*-?\s*(.*)', c, re.I)
        if m:
            if m.group(4):
                name = m.group(4).upper()
            else:
                name = np.nan
            out = [True, m.group(1), name] #return: is casa, casa, name
    
    return out 

def fuzzy_match_customer(customer, description, casas):
    if customer:
        return customer

    if not isinstance(description,str):
        return []
    ratio = 0
    customers = []
    
    #try first with history
    #TO-DO
    
    #try first with casa
    customers = list(df_customers[df_customers['casa'].isin(casas)]['Customer'])
        
    #fuzzy match
    if not customers:
        d = description.upper()
        ignores = ["TRANSFERENC","BANCOBCR","DEPOSITOS", "CREDITO",
               "INTERBANCARI", "SINPE", "MOVIL", "PIN ENTRANTE",
               "TEF DE:", "-", "_", "*"]
        for ignore in ignores:
            d = d.replace(ignore,' ')

        fm = process.extract(d, dict_customers.keys(),  scorer=fuzz.token_set_ratio, limit=1)
        if fm[0][1] > 95:
            customers.append(dict_customers[fm[0][0]])
    return customers

def fuzzy_match_description_hist(description):
    if not isinstance(description,str):
        return []
    ratio = 0
    customers = []
    
    fm = process.extract(description, description_hist['descripcion'],  scorer=fuzz.token_set_ratio, limit=1)
    if fm[0][1] > 90:
        customers = list(description_hist[description_hist.descripcion == fm[0][0]].cliente)
    return customers   

def fix_credit_column(c):
    out = c
    if isinstance(c,str) or isinstance(c,int) or isinstance(c,float):
        c = str(c)
        m = re.search(r'(.*)\+', c, re.I)
        if m:
            return float(m.group(1).replace(',',''))
        m = re.search(r'(.*)\-', c, re.I)
        if m:
            return np.nan
        m = re.search(r'\-(.*)', c, re.I)
            
        if m:
            return np.nan
    return out 

#return num of mensualidades y si incluye patos
def check_payment(amount, m=None, recursive=True):
    if not m:
        m = main_config['mensualidad']

    q = int(amount / m)
    r = amount % m
    
    if q > 0 and r == 0 :
        #mensualidad exacta
        return [q, 0]
    elif q > 0 and r == q * main_config['agua lago']:
        #mensualidad con agua lagos exacta
        return [q, 1]
    else:
        #probar mensualidad anterior
        if recursive:
            return check_payment(amount, m=main_config['mensualidad anterior'], recursive=False)
        else:
            return [0, 0]
        
def process_customers_excel(excel_clientes):
    df = gendf_from_excel_table(excel_clientes,['Customer full name', 'Email address'],stop_if_empty=False)
    df.rename(columns={"Customer full name":"Customer", "Email adress": "Email"}, inplace=True)
    
    #split casa and name
    df['is_casa'], df['casa'], df['name'] = zip(*df['Customer'].map(get_casa_from_customer))
    
    return df[['Customer','is_casa','casa','name']]


def process_bank_excel(inputs_dir,bank,bank_config):
    #swaps key and values
    bank_cols = {v: k for k, v in bank_config[bank].items()}

    #get which lines have the table we want to read with pandas
    bank_excel = os.path.join(inputs_dir, "banco_" + bank + ".xls")
    if not os.path.exists(bank_excel):
         bank_excel += "x" 
    
    if not os.path.exists(bank_excel):
        print("INFO: File {}(x) not found. Bank {} not processed.".format(bank_excel, bank))
        return pd.DataFrame([])
    
    #create dataframe with normalized columns
    bankdf = gendf_from_excel_table(bank_excel,bank_cols)
    if bankdf.empty:
        print("INFO: Columns {} not found in file {}. Check configuration".format(bank_cols, bank_excel))
        return bankdf
    
    bankdf.rename(columns=bank_cols, inplace=True)
    bankdf.replace(0, np.nan, inplace=True)
    bankdf.replace('''''', np.nan, inplace=True)

    #Fix credito colums when +- symbols are used
    bankdf['credito_tmp'] = bankdf['credito'].apply(fix_credit_column)
    bankdf['credito'] = bankdf['credito_tmp']
    
    #keep only rows with credits
    bankdf.dropna(axis=0, subset="credito", inplace=True)
    
    #keep only configured columns
    bankdf=bankdf[1:][bank_cols.values()]

    #Take out non numeric credit rows and description
    bankdf=bankdf[pd.to_numeric(bankdf['credito'], errors='coerce').notnull()]
    bankdf=bankdf[bankdf['descripcion'].notnull()]

    #add bank columns
    bankdf.insert(0,'excel',bank_excel)
    bankdf.insert(0,'banco',bank)
    bankdf['casas']=bankdf['descripcion'].apply(get_casas_from_description)
    
    #Try to guess customers
    bankdf['cliente'] = bankdf.apply(lambda x: fuzzy_match_description_hist(x['descripcion']), axis=1)
    bankdf['cliente'] = bankdf.apply(lambda x: fuzzy_match_customer(x['cliente'], x['descripcion'], x['casas']), axis=1)
    bankdf['cliente'] = bankdf['cliente'].apply(lambda x: ', '.join(x) if isinstance(x,list) else '')
    
    #bankdf.credito.astype('float')
    
    #bankdf['num meses'] = bankdf['credito'].apply(check_payment)
    bankdf['num meses'], bankdf['agua patos'] = zip(*bankdf['credito'].apply(check_payment))
    
    return bankdf


# In[ ]:


########### MAIN ############

#work directory
if is_interactive():
    sys.argv = ['', r'C:\Users\villalta\Documents\Personal\repos\qb_test\Junio']
    #sys.argv = ['', r"/Users/oscar/Documents/work/qb_example/"]
if len(sys.argv) != 2:
    print("Error: Se debe indicar el directorio donde estan las entradas")
    exit(1)
    
inputs_dir = os.path.abspath(sys.argv[1])
if not os.path.exists(inputs_dir):
    print("Error: El directorio con los archivos de entrada no existe: " + inputs_dir)
    exit(1)

outputs_dir = os.path.join(inputs_dir, "salidas")
if not os.path.exists(outputs_dir):
    os.makedirs(outputs_dir)

#Double check that banks excel doesnt exists (it is dangerous to overwritte this)
bank_excel = os.path.join(outputs_dir, "bancos.xlsx")
if os.path.exists(bank_excel):
    print("Error: El archivo {} ya existe. Si desea regenerarlo borrelo manualmente antes de correr Procesar Bancos.".format(bank_excel))
    exit(1)

#Load customers
excel_clientes = os.path.join(inputs_dir, "clientes_quickbooks.xlsx")
print(f"Cargando clientes desde: {excel_clientes}")
df_customers = process_customers_excel(excel_clientes)
dict_customers = df_customers[df_customers['is_casa']][['Customer', 'name']].dropna().set_index('name').to_dict()['Customer']
casa_customers = list(df_customers[df_customers['is_casa']]['Customer'])

#Load configuration
config_path = os.path.join(inputs_dir, "config.xlsx")
if os.path.exists(config_path):
    config_excel = config_path
else:
    config_excel = "config.xlsx"
    
print("Leyendo la configuracion desde: " + config_excel)
df_main_config = pd.read_excel(config_excel, sheet_name="principal").dropna(axis=1, how="all")
main_config = df_main_config.set_index("item").to_dict()['valor']
df_bank_config = pd.read_excel(config_excel, sheet_name="columnas bancos").dropna(axis=1, how="all")
bank_config = df_bank_config.set_index("columna").to_dict()
description_hist =  pd.read_excel(config_excel, sheet_name="descripciones anteriores").dropna(axis=1, how="all")

#Load all banks data
bank_dfs = []
for bank in bank_config:
    print("Procesando banco: " + bank)
    bank_df = process_bank_excel(inputs_dir, bank, bank_config)
    if not bank_df.empty:
        bank_dfs.append(bank_df)
print("Transferencias bancarias cargadas exitosamente")

if bank_dfs:
    df = pd.concat(bank_dfs).reset_index(drop=True)
    #Save excel
    with pd.ExcelWriter(bank_excel) as writer:
        df.to_excel(writer, index=False, sheet_name='bancos') 
        df_customers.to_excel(writer, index=False, sheet_name='clientes')
    print("Excel creado: {}".format(bank_excel))
else:
    print("Error: no fue encontrado el excel de ningun banco. Nada creado.")

print("Fin del proceso.")


# In[ ]:




