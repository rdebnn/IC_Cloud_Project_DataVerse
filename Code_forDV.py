config_f = "Config.json"
log_f = 'SLA_logs.log'

import logging, json, sys, time, threading
from os import path, listdir
from random import randint
import time

try:
    import pandas as pd
    from fpdf import FPDF
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.wait import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.edge.options import Options
    import win32com.client as win32

except Exception as e:
    logging.basicConfig(filename=log_f, level=logging.INFO, format='%(asctime)s:%(levelname)s: %(message)s')
    logging.info('----- SLA processing started -------------------------')
    logging.error(f'Error during import : {e}')
    logging.info('----- SLA processing Ended ---------------------------\n')
    sys.exit()

class PDF(FPDF):  # not required now - but can be used in the future
    def header(self):
        self.set_font('helvetica','B',20)
        self.cell(80)
        self.cell(30,10,'SLA Billing', border=True, ln=1, align='C')

def check_CommonFields_structure(df):
    if len(df.axes[1]) != 2 or len(df.axes[0]) != 9:
        logging.debug('Number of Rows/Columns are changed in [CommonFields] sheet')
        return False
    if type(df.iloc[3, 1]) != type('test'):
        logging.debug('Sender Company code must be of type string [check_CommonFields_structure]')
        return False
    if len(df.iloc[4, 1]) != 3:
        logging.debug('Currency Length is not 3 -- [check_CommonFields_structure]')
        return False
    if not(type(1.1) == type(df.iloc[7,1])):
        logging.debug('Std Exchange Rate value is not float -- [check_CommonFields_structure]')
        return False
    if len(df.iloc[8, 1]) != 2:
        logging.debug('Quarter string value must be of length 2 -- [check_CommonFields_structure]')
        return False
    try:
        for i in range(2,9):
            if pd.isnull(df.iloc[i,0]) and pd.isnull(df.iloc[i,1]):
                logging.debug('Common fields sheet has empty spaces -- [check_CommonFields_structure]')
                return False
    except:
        logging.debug('Something went wrong in [check_CommonFields_structure] method')
        return False
    return True

def check_files(dict_conf):
    if len(dict_conf['Excel_File']['SLA_Columns']) != 12:  
        logging.error('All the columns are not defined in configuration file [SLA_Columns]')
        return False
    if dict_conf.get('debug_mode') is None:
        logging.error('[debug_mode] key is not defined in the configuration file')
        return False
    if not(isinstance(dict_conf['debug_mode'], bool)):
        logging.error('[debug_mode] key must be a Boolean value')
        return False
    if dict_conf.get('Create_Invoices') is None:
        logging.error('[Create_Invoices] key is not defined in the configuration file')
        return False
    if not(isinstance(dict_conf['Create_Invoices'], bool)):
        logging.error('[Create_Invoices] key must be a Boolean value')
        return False
    if dict_conf.get('Send_Invoices') is None:
        logging.error('[Send_Invoices] key is not defined in the configuration file')
        return False
    if not(isinstance(dict_conf['Send_Invoices'], bool)):
        logging.error('[Send_Invoices] key must be a Boolean value')
        return False
    if not(dict_conf.get('Excel_File')):
        logging.error('[Excel_File] key is not defined in the configuration file')
        return False
    if not(dict_conf.get('Invoices_Folder')):
        logging.error('[Invoices_Folder] key is not defined in the configuration file')
        return False
    if not(path.exists(dict_conf['Excel_File']['File_Name'])):
        logging.error('Invalid excel file path : ' + dict_conf['Excel_File']['File_Name'])
        return False
    if not(dict_conf['Excel_File'].get('File_Name')):
        logging.error('[File_Name] key is not defined in [Excel_File] configuration')
        return False
    if not(dict_conf['Excel_File'].get('Sheet2_SLA')):
        logging.error('[Sheet2_SLA] key is not defined in [Excel_File] configuration')
        return False
    if not(path.exists(dict_conf['Invoices_Folder'])):
        logging.error('Invoice Folder Path Error: '+ dict_conf['Invoices_Folder'])
        return False
    if not(dict_conf.get('Web_Driver')):
        logging.error('[Web_Driver] key is not defined in the configuration file')
        return False
    if not(path.isfile(dict_conf['Web_Driver'])):
        logging.error('Chrome Driver Path Error: '+ dict_conf['Invoices_Folder'])
        return False
    if not(dict_conf.get('SLA_Cloud')):
        logging.error('[SLA_Cloud] key is not defined in the configuration file')
        return False
    return True

def Create_Invoices(dict_conf):
    if len(listdir(dict_conf['Invoices_Folder']))>0:
        logging.error('Please clear the Invoices from the mentioned folder path before generating new')
        return False
    com_Std_Exch_rate = dict_conf["exchnage rate"]
    com_Quarter = dict_conf["Quarter"]
    # df = pd.read_excel(dict_conf['Excel_File']['File_Name'], \
            # sheet_name=dict_conf['Excel_File']['Sheet2_SLA'], engine='openpyxl')
    

    try:
        df = pd.read_excel(dict_conf['Excel_File']['File_Name'], \
            sheet_name=dict_conf['Excel_File']['Sheet2_SLA'], engine='openpyxl')
        # print(df)
    except Exception as e:
        print(e)
        logging.error('Unable to Read "'+dict_conf['Excel_File']['Sheet2_SLA']+'" sheet')
        return False
    for col in dict_conf['Excel_File']['SLA_Columns']:
        if col not in df.columns:
            logging.error('Error in SLA database column name (' + col +')')
            return False

                            #### NEW FILE EXTRACTION FOR INTEGRATION WITH DATAVERSE #####
    
    col_con = dict_conf['Excel_File']['SLA_Columns'][0]  # CountryName
    col_ccd = dict_conf['Excel_File']['SLA_Columns'][2]  # Comp Code
    col_stk = dict_conf['Excel_File']['SLA_Columns'][10]  # Stakeholder/Approver
    # col_sup = dict_conf['Excel_File']['SLA_Columns'][5]  # Support
    col_dept = dict_conf['Excel_File']['SLA_Columns'][5] # deptatment
    col_fte = dict_conf['Excel_File']['SLA_Columns'][6]  # FTEs
    col_mon = dict_conf['Excel_File']['SLA_Columns'][9]  # Months
    col_cfa = dict_conf['Excel_File']['SLA_Columns'][4]  # Cost/FTE/Annum
    col_dkk = dict_conf['Excel_File']['SLA_Columns'][7]  # Total cost (DKK)
    col_inr = dict_conf['Excel_File']['SLA_Columns'][8]  # INR
    col_text = dict_conf['Excel_File']['SLA_Columns'][1] # Text
    df['new_support']= df['new_support'].fillna('')
    df['new_comments']= df['new_comments'].fillna('N/A')
    df["Dipartment-team"] = df[col_dept].astype(str) +" "+ df["new_support"].astype(str)


    grp_fields = df.groupby([col_stk, col_ccd, col_con])
    for fields, df_grp in grp_fields:

        stakeholder = fields[0]
        ccd = fields[1]
        con = fields[2]
        print(con)
        pad_w = 20
        pad_y = 10
        cell_w = 206
        cell_h = 9
        pdf = FPDF('P', 'mm', 'Letter')
        pdf.add_page()
        pdf.set_font('helvetica', 'B', 10) 
        pdf.set_left_margin(-18)  
        pdf.cell(-15,20, ln=True)
        pdf.cell(pad_w, pad_y)
        pdf.cell(cell_w, cell_h, f'Finance GBS Bangalore FTE Costs', ln=True, border=1, align='C')
        pdf.set_font('helvetica', 'B', 6)
        cell_w = 25
        cell_h = 36
        pdf.cell(pad_w, pad_y)
        pdf.cell(cell_w+6, cell_h,f'Department', border=1, align='C')
        pdf.cell(cell_w-3, cell_h, f'Cost/FTE(DKK)', border=1, align='C')
        pdf.cell(cell_w - 10, cell_h,f'FTE', border=1, align='C')
        pdf.cell(cell_w - 10, cell_h,f'Months', border=1, align='C')
        
        pdf.cell(cell_w-2, cell_h, f'Qtr Cost(DKK)', border=1, align='C')
        pdf.cell(cell_w - 5, cell_h, f'STD Exc rate', border=1, align='C')
        pdf.cell(cell_w, cell_h, f'Qtr Cost(INR)', border=1, align='C')
        pdf.cell(cell_w+5 , cell_h, f'Std. Comments', border=1, align='L')
        pdf.cell(cell_w, cell_h, f'Comments', border=1, align='L', ln=True)
        

        pdf.set_font('helvetica', '', 5) 
        
        new_grp= df_grp.groupby(["Dipartment-team","new_months","new_ftecostdkk","new_comments"]).sum()
        print(new_grp)
        
        for field, row in new_grp.iterrows():
            month = field[1]
            dept = field[0]
            costindkk = field[2]
            additional_comments = field[3]
            effective_page_width = pdf.w - 2*pdf.l_margin
            print(effective_page_width)
            print(new_grp.columns)
            pdf.cell(pad_w, pad_y)
            pdf.cell(cell_w+6, cell_h, f'{dept}', border=1, align='C')
            pdf.cell(cell_w-3, cell_h, f'{round(costindkk,2)}', border=1, align='C')
            pdf.cell(cell_w - 10, cell_h, f'{(row["new_fte"])}', border=1, align='C')
            pdf.cell(cell_w - 10, cell_h, f'{month}', border=1, align='C')
            pdf.cell(cell_w-2, cell_h, f'{int(row["new_costindkk"])}', border=1, align='C')
            pdf.cell(cell_w - 5, cell_h, f'{com_Std_Exch_rate}', border=1, align='C')
            try:
                pdf.cell(cell_w, cell_h, f'{int(row["new_ftecostinr"])}', border=1, align='C')
                print("**********************")
            except:
                logging.error(f'Invalid value at '+round(row["new_ftecostinr"],2)+' column')
                return False
            print(dept+" "+str(row["new_fte"]) +" FTE "+" "+str(row["new_companycode"])+" Intercompany Invoice")
            text =dept+" "+str(row["new_fte"]) +" FTE "+str(int(ccd))+" IC Inv."
            pdf.set_font_size(5)
            len_text=len(text)
            len_comments=len(additional_comments)

            temp_h = cell_h / ((int)(len_text/30) + 1)
            x= pdf.get_x()
            y= pdf.get_y()
            pdf.multi_cell(w=cell_w+5, h= temp_h, txt=f'{text}', border=1)
            pdf.set_xy(x + cell_w + 5, y)
            commenet_h = cell_h / ((int)(len_comments/20) + 1)
            pdf.cell(h=commenet_h, w=cell_w, txt=f'{additional_comments}', border=1, ln=True, align='L')
                
        pdf.set_font('helvetica', 'B', 7) 
        pdf.cell(pad_w+6, pad_y)
        pdf.cell(cell_w , cell_h,f'Total', border=1, align='C')
        pdf.cell(cell_w-3 , cell_h, '', border=1, align='C')
        pdf.cell(cell_w - 10, cell_h,'', border=1, align='C')
        pdf.cell(cell_w-10, cell_h, '', border=1, align='C')
        pdf.cell(cell_w-2, cell_h, '', border=1, align='C')
        pdf.cell(cell_w-5 , cell_h, '', border=1, align='C')
        pdf.cell(cell_w, cell_h, f'{round(new_grp["new_ftecostinr"].sum(),2)}', border=1, align='C', ln=True)
        try:
            pdf.output(dict_conf['Invoices_Folder'] + '/' +con+ '_' + \
                str('0'*(4-len(str(int(ccd))))+str(int(ccd))) + '_' + stakeholder + '_' + com_Quarter + '.pdf')
        except:
            logging.error(f'Invoice for {stakeholder} is failed due to file naming issues')
            return False
        logging.debug(f'Invoice for {stakeholder} is generated successfully')
    if len(grp_fields) == 0:
        logging.warning(f'SLA database is empty')
        return False
    else:
        logging.info(f'{len(grp_fields)} invoices are generated successfully')
    return True

class cnt_Queue(object):
    def __init__(self):
        self.success_items = []
        self.failure_items = []
    def __repr__(self):
        return self.success_items, self.failure_items
    def __str__(self):
        return 'Success: ' ','.join(self.success_items) + \
        ' -- Failed' + ','.join(self.failure_items)
    def enque_s(self, add):
        self.success_items.append(add)
    def enque_f(self, add):
        self.failure_items.append(add)
    def size_s(self):
        return len(self.success_items)
    def size_f(self):
        return len(self.failure_items)

global queue
queue = cnt_Queue()

def Thread_Web(sla_path, req_lst, df_grp, inv_path, temp_timer):
    
    edge_options= Options()
    print("**********************")
    print(dict_conf['Web_Driver'])
    # driver = webdriver.Edge()
    
    # driver= webdriver.Edge(executable_path=dict_conf['Web_Driver'])
    edge_options.add_experimental_option("detach", True)
    driver= webdriver.Edge(options= edge_options)

    
    # driver= webdriver.Edge(dict_conf['Web_Driver'], options= edge_options)
    driver.implicitly_wait(30)

       
    
    driver.get(sla_path)
    time.sleep(randint(5,15))

    if driver.title != 'Intercompany Invoicing':
        time.sleep(randint(5,10))
        logging.debug('waiting 2nd time')
        if driver.title != 'Intercompany Invoicing':
            time.sleep(randint(5,15))
            logging.debug('waiting 3rd time')
            if driver.title != 'Intercompany Invoicing':
                logging.debug(f'Refreshing the page - {req_lst[3]}')
                driver.get(sla_path)                                                   #refreshing the page
                time.sleep(randint(15,25))
                if driver.title != 'Intercompany Invoicing':
                    logging.debug(f'Refreshing the page - 2nd time- {req_lst[3]}')
                    driver.get(sla_path)
                    time.sleep(randint(10,20))
                    if driver.title != 'Intercompany Invoicing':
                        logging.error(f'Error in the Website address (title != Intercompany) - {req_lst[3]}')
                        driver.close()
                        driver.quit()
                        queue.enque_f(req_lst[3])
                        return
    try:
        ele = WebDriverWait(driver, 180).until(
            EC.presence_of_element_located((By.ID, '__box4-inner')))
    except:
        driver.get(sla_path)                                                       #if the box does not appear then refresh
        logging.debug(f'Waitig for the element to Load - {req_lst[3]}')
        time.sleep(randint(20,30))
        if driver.title != 'Intercompany Invoicing':
            logging.error(f'Error in the Website address (title != Intercompany or page not loaded) - {req_lst[3]}')
            driver.close()
            driver.quit()
            queue.enque_f(req_lst[3])
            return
    try:
        # driver.find_element_by_id('__box0-inner').send_keys(req_lst[0])  # Invoice Flow
        time.sleep(1.5)
        driver.find_element(By.ID,'__box0-inner').send_keys("Charge-back")
        driver.find_element(By.ID,'__box2-inner').clear()
        driver.find_element(By.ID,'__box2-inner').send_keys(req_lst[1])  # sender cocd
        driver.find_element(By.ID,'__box3-inner').send_keys(req_lst[2])  # receiver cocd
        driver.find_element(By.ID,'__input0-inner').send_keys(req_lst[3])  # stakeholder
        driver.find_element(By.ID,'__box4-inner').send_keys(req_lst[4])  # Currency (INR)
    except Exception as e:
        logging.error(f'{req_lst[3]} - common fields update error -> {e}')
        driver.close()
        driver.quit()
        queue.enque_f(req_lst[3])
        return
    i = 0
    time.sleep(0.5)
    try:
        for _, row in df_grp.iterrows():
            driver.find_element(By.ID,'__button2-BDI-content').click()
            time.sleep(0.1)
            driver.find_element(By.ID,'__box7-__xmlview0--LineItemsTable-'+str(i)+'-inner').send_keys(req_lst[5])  # req_lst[5] - Cost Type
            # driver.find_element(By.ID,'__input4-__xmlview0--LineItemsTable-'+str(i)+'-inner').send_keys(req_lst[6])  # req_lst[6] - Cost Centre
            driver.find_element(By.ID,'__input4-__xmlview0--LineItemsTable-'+str(i)+'-inner').send_keys(row[3])
            time.sleep(0.1)
            driver.find_element(By.ID,'__box8-__xmlview0--LineItemsTable-'+str(i)+'-inner').send_keys('Mark up, Other')  # hardcoded
            driver.find_element(By.ID,'__input7-__xmlview0--LineItemsTable-'+str(i)+'-inner').clear()
            driver.find_element(By.ID,'__input7-__xmlview0--LineItemsTable-'+str(i)+'-inner').send_keys(str(round(row[req_lst[7]], 2)).replace('.',','))  # req_lst[7] - INR
            print()
            if pd.isna(df_grp.iloc[0,11]):
                team =" "
            else:
                team = df_grp.iloc[0,11]
            
            print("TEAM Selected is :  "+row[11])
            
            text = str(row[11])+" "+str(team)+" "+" FTE "+" "+f'{(row["new_fte"])}'+"  "+req_lst[2]+" Intercompany Invoice "
            driver.find_element(By.ID,'__input8-__xmlview0--LineItemsTable-'+str(i)+'-inner').send_keys(text)  # req_lst[8] - Inv Text
            i += 1
    except Exception as e:
        logging.error(f'{req_lst[3]} - table update error (in the table)-> {e}')
        driver.close()
        driver.quit()
        queue.enque_f(req_lst[3])
        return

    try:             #upload the invoice pdf                                                                                                          #Upload the Invoice PDFs
        ele = driver.find_element(By.ID,'__xmlview0--UploadAttachment-fu')
        ele.send_keys(inv_path)
        queue.enque_s(req_lst[3])
    except Exception as e:
        logging.error(f'{req_lst[3]} - Invoice upload error -> {e}')
        driver.close()
        driver.quit()
        queue.enque_f(req_lst[3])
        return
    time.sleep(temp_timer)
    # driver.find_element_by_id('__button5-BDI-content').click()  # Go Button

def Update_WebData(dict_conf):
    if len(listdir(dict_conf['Invoices_Folder']))==0:
        logging.error('No Invoices are present in the mentioned folder path')
        return False
    
    com_Invoice_Flow = dict_conf["Invoice Flow"]
    com_Sender_CoCd = dict_conf["Sender Comp code"]
    com_Currency = dict_conf["Currency"]
    com_Cost_Type = dict_conf["Cost Type"]
    com_CostCenter = dict_conf["CostCentre"]
    com_Quarter = dict_conf["Quarter"]
    df = pd.read_excel(dict_conf['Excel_File']['File_Name'], \
    sheet_name=dict_conf['Excel_File']['Sheet2_SLA'], engine='openpyxl')
    # try:
    #     df = pd.read_excel(dict_conf['Excel_File']['File_Name'], \
    #         sheet_name=dict_conf['Excel_File']['Sheet2_SLA'], engine='openpyxl')
    #     print(df)
    # except Exception as e:
    #     # logging.error('Unable to Read "'+dict_conf['Excel_File']['Sheet2_SLA']+'" sheet')
    #     logging.error(e)

    #     return False
    for col in dict_conf['Excel_File']['SLA_Columns']:
        if col not in df.columns:
            logging.error('Error in SLA database column names. (' + col +')')
            return False
    # if df.isnull().values.any() == True:
    #     logging.error('SLA Database contains empty values')
    #     return False
    # column definitions
    col_con = dict_conf['Excel_File']['SLA_Columns'][0]  # CountryName
    col_ccd = dict_conf['Excel_File']['SLA_Columns'][2]  # Comp Code
    col_stk = dict_conf['Excel_File']['SLA_Columns'][10]  # Stakeholder/Approver
    col_inr = dict_conf['Excel_File']['SLA_Columns'][8]  # INR
    # col_txt = dict_conf['Excel_File']['SLA_Columns'][]  # Inv Text

    grp_fields = df.groupby([col_stk, col_ccd, col_con])
    grp_fields.count()
    # grp_stk = df.groupby(col_stk)

    threads = []
    
    for fields, df_grp in grp_fields:
        stakeholder = fields[0]
        print(fields)
        rec_cocd = '0'*(4-len(str(int(df_grp.iloc[0,2]))))+str(int(df_grp.iloc[0,2]))
        # print(df_grp)
        if pd.isna(df_grp.iloc[0,11]):
            team =" "
        else:
            team = df_grp.iloc[0,11]

        # text = str(df_grp.iloc[0,5])+" "+str(team)+" "+" FTE "+" "+df_grp.iloc[0,6]+"  "+rec_cocd+" Intercompany Invoice "
        
        req_lst = [com_Invoice_Flow, com_Sender_CoCd, rec_cocd, 
                    stakeholder, com_Currency, com_Cost_Type, com_CostCenter,
                    col_inr] # 012,3456,78
        # pdf.output(dict_conf['Invoices_Folder'] + '/' +  row[col_cnr] + '_' + \
        #         str(row[col_ccd]) + '_' + stakeholder + '_' + com_Quarter + '.pdf')
        inv_path = dict_conf['Invoices_Folder'] + '\\' +  fields[2] + '_' + rec_cocd + '_' + stakeholder + '_' + com_Quarter + '.pdf'
        logging.info("INVOICE PATH: "+inv_path)
        if not(path.isfile(inv_path)):
            print("Path "+inv_path+" Not present")
            logging.error(f'path error for {stakeholder}: {inv_path}')
        else:
            # remove temp_timer later awrm
            if dict_conf.get('temp_timer'):
                if type(dict_conf.get('temp_timer')) == type(1):
                    t = threading.Thread(target=Thread_Web, args=(dict_conf['SLA_Cloud'], req_lst, df_grp, inv_path, dict_conf.get('temp_timer')))
                else:
                    t = threading.Thread(target=Thread_Web, args=(dict_conf['SLA_Cloud'], req_lst, df_grp, inv_path, 2))
            else:
                t = threading.Thread(target=Thread_Web, args=(dict_conf['SLA_Cloud'], req_lst, df_grp, inv_path, 2))
            threads.append(t)
            t.start()
            # Thread_Web(dict_conf['SLA_Cloud'], req_lst, df_grp, inv_path)
            if len(threads) == 5:
                for t in threads:
                    t.join()
                threads.clear()
    for t in threads:
        t.join()
    if queue.size_s()+queue.size_f() == 0:
        logging.warning(f'SLA database is empty')
    else:
        logging.info(f'Web Update Summary : Successful -> {queue.size_s()} ; Failed -> {queue.size_f()}')
        for fields, df_grp in grp_fields:  
            print("For mailing Part____________________________________________")
            stakeholder=df_grp[col_stk].iloc[0]
            companycode= str(df_grp["new_companycode"].iloc[0]).zfill(4)
            comparing_lst=list([stakeholder, companycode])
            print(comparing_lst)
            failed_list= queue.failure_items
            print(failed_list)
            if(not(comparing_lst in failed_list)):
                print(df_grp)
                send_mail((stakeholder)+"@novonordisk.com", df_grp, dict_conf)
                # send_mail("kjmu@novonordisk.com", df_grp, dict_conf)

    return True

def send_mail(initials, df,dict_conf):
    df.reset_index(drop=True, inplace=True)
    # convert dataframe to html
    html_table = df.to_html(index=False)
    # tabular_fields = df.columns
    # tabular_table = PrettyTable()
    # tabular_table.field_names = tabular_fields 
    # for i in range(len(df)):
    #     tabular_table.add_row(df.iloc[i])
    
    # html_table = tabular_table.get_html_string()
    html_body= """
    <body>Dear User,</body>
    <p>
        We have initiated Intercompany invoice via IC Cloud. Kindly review and approve the same
    </p>
    <p>
        Please find the relevant data below:
    </p>


    """
    footer_html= """
    <p>
        Link for Approval: <a href="https://flpnwc-jk5mksf7nb.dispatcher.hana.ondemand.com/sites#WorkflowTask-DisplayMyInbox">Click here</a>
        <br>If you are facing access issues: <a href="https://novonordisk.sharepoint.com/sites/CorporateAccounting/SitePages/Intercompany-invoice.aspx">Click here</a>
    </p>
    <p>
        Further queries on the IC Cloud, please write to <a href="intercompanynngroup@novonordisk.com.">intercompanynngroup@novonordisk.com.</a>
        <br>If you have queries regarding service details or invoice amount please write to your contact manager in Global IT GBS
    </p>
    <p>
        Thanks,
    </p>

    """

    o = win32.Dispatch("Outlook.Application")
    oacctouse = None
    for oacc in o.Session.Accounts:
        if oacc.SmtpAddress == dict_conf["SenderEmailID"]:
            oacctouse = oacc
            break
    Msg = o.CreateItem(0)
    if oacctouse:
        Msg._oleobj_.Invoke(*(64209, 0, 8, 0, oacctouse))  # Msg.SendUsingAccount = oacctouse

    Msg.Subject = "FTE Cost Intercompany Invoice - To be Reviewed & Approved"
    Msg.To = initials
    Msg.HTMLBody = html_body+html_table+footer_html
    Msg.Save()


def main(dict_conf):
    if dict_conf['Create_Invoices'] == True:
        if Create_Invoices(dict_conf) == True:
            if dict_conf['Send_Invoices'] == True:
                if Update_WebData(dict_conf) == True:
                    pass
                else:
                    logging.error('Error while updating the webdata (No Threads are executed)')
        else:
            logging.error('Error while creating the invoices')              
    elif dict_conf['Send_Invoices'] == True:
        if Update_WebData(dict_conf) == True:
            pass
        else:
            logging.error('Error while updating the webdata (No Threads are executed)')       
    else:
        logging.info('Please enable the actions for the script to run')

if __name__ == '__main__':
    err_msg = ''
    dict_conf= {}
    if path.exists(config_f) == True:
        with open(config_f, 'r') as f:
            try:
                dict_conf = json.loads(f.read())
            except json.JSONDecodeError:
                err_msg = 'Config File Format is corrupted'
            except:
                err_msg = 'Error while reading the Config File - This is a new exception'
    else:
        err_msg = 'Config File does not exist'

    # Debug OR Info Mode
    if not(err_msg):
        if dict_conf.get('debug_mode') is not None:
            if dict_conf['debug_mode'] == True:
                logging.basicConfig(filename=log_f, level=logging.DEBUG, format='%(asctime)s:%(levelname)s: %(message)s')
            else:
                logging.basicConfig(filename=log_f, level=logging.INFO, format='%(asctime)s:%(levelname)s: %(message)s')
            logging.info('----- SLA processing started -------------------------')
        else:
            logging.basicConfig(filename=log_f, level=logging.INFO, format='%(asctime)s:%(levelname)s: %(message)s')
            logging.info('----- SLA processing started -------------------------')
            logging.warning('[debug_mode] key is not defined')
    else:
        # since config file is corrupt debug mode is irreleavent and set to OFF here
        logging.basicConfig(filename=log_f, level=logging.INFO, format='%(asctime)s:%(levelname)s: %(message)s')
        logging.info('----- SLA processing started -------------------------')
    if not(err_msg):
        if check_files(dict_conf):
            main(dict_conf)
        else:
            logging.debug('[check_files] module returned false')
    else:
        logging.error(err_msg)
    logging.info('----- SLA processing Ended ---------------------------\n')

