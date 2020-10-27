#!/usr/bin/env python
# coding: utf-8

# In[ ]:


#updates 
  #removed negative points for multiple incentive moves 
  #stopped support for old DSS workbook


# # Definitions

# In[1]:


import pandas as pd
import numpy as np
import xlwings as xw
import sys

# Display entire Scenario string in notebook
pd.options.display.max_colwidth = 4000


# # Functions

# In[2]:


def DSS_extract_models_and_output_tabs_new_DSS(DSS_file_path):
    wb = xw.Book(DSS_file_path)
    
    output_tabs = []
    for sheet in wb.sheets:
        ws = wb.sheets[sheet]
        if ws.range("A11").value == 'Ref: BC Elasticity':
            output_tabs.append(sheet.name)
                    
    model_list = []
    for i in range(0,11):
        model = wb.sheets['Input'].range(3, 7 + i*6).value
        if model == None:
            continue
        else:
            region = str(wb.sheets['Input'].range(4, 7 + i*6).value)
            model_id = model + "_" + region
            model_list.append(model_id)
    return model_list, output_tabs


# In[3]:


def DSS_extract_models_and_output_tabs_old_DSS(DSS_file_path):
    wb = xw.Book(DSS_file_path)
    
    output_tabs = []
    for sheet in wb.sheets:
        ws = wb.sheets[sheet]
        if ws.range("A11").value == 'Ref: BC Elasticity':
            output_tabs.append(sheet.name)
                    
    model_list = []
    for i in range(0,11):
        model = str(wb.sheets['Input'].range(3, 8 + i*2).value) 
        if model == None:
            continue
        else:
            model_year = str(wb.sheets['Input'].range(4, 8 + i*2).value)
            region = str(wb.sheets['Input'].range(5, 8 + i*2).value)
            model_id = model + "_" + model_year + "_" + region
            model_list.append(model_id)
    
    return model_list, output_tabs


# In[4]:


def write_to_output_tab_new_DSS(DSS_file_path, output_tab, DSS_scenarios):
    wb = xw.Book(DSS_file_path)
    ws = wb.sheets[output_tab]
    
    # Reset scenarios to 'x'
    for i in range(36,73,2):
        ws.range(i,8).value = 'x'
    # Write Chooser scenarios    
    for index, value in enumerate(DSS_scenarios):
        ws.range(36+2*index, 8).value = value


# In[5]:


def write_to_output_tab_old_DSS(DSS_file_path, output_tab, DSS_scenarios):
    wb = xw.Book(DSS_file_path)
    ws = wb.sheets[output_tab]
        
    #  Reset scenarios to 'x'
    for i in range(36,71,2):
        ws.range(i,3).value = 'x'
    # Write Chooser scenarios    
    for index, value in enumerate(DSS_scenarios):
        ws.range(34+2*index, 3).value = value


# In[6]:


def apply_spend_filters(df, min_spend, max_spend):
    df = df[df['spend_delta'] > min_spend]
    df = df[df['spend_delta'] < max_spend]
    
    return df


# In[7]:


def convert_to_DSS_scenarios(single_lever_indices, df, max_no_of_scenarios):
    scenarios_row = list(df.index.values)
    DSS_scenarios = [x + 1 for x in scenarios_row]
    
    DSS_scenarios_combined = single_lever_indices
    [DSS_scenarios_combined.append(x) for x in DSS_scenarios if x not in single_lever_indices]
    
    DSS_scenarios_combined = DSS_scenarios_combined[0:max_no_of_scenarios]
        
    return DSS_scenarios_combined


# In[8]:


def find_single_lever_indices_new(df, baseline, delta_columns, APR_delta_columns, min_cash_bool, dbl_min_cash_bool, APR_bool, min_combo_bool, dbl_min_combo_bool, min_lease_bool, dbl_min_lease_bool):

    # Find lowest increment of enhancement
    cash_min_enh = df[13:][(df[13:]['cash_delta'] > 0) & (df[13:]['no_of_moves'] == 1)]['cash_delta'].min()
    combo_min_enh = df[13:][(df[13:]['combo_delta'] > 0) & (df[13:]['no_of_moves'] == 1)]['combo_delta'].min()
    lease_min_enh = df[13:][(df[13:]['lease_delta'] > 0) & (df[13:]['no_of_moves'] == 1)]['lease_delta'].min()
    
    # Create sum across all moves
    df['sum_of_moves'] = df[delta_columns[0:8]].sum(axis=1)
    
    # Find 'combined_APR_delta' that equals -1 move across all terms
    APR_single_lever_sum = 0
    for x in APR_delta_columns:
            if baseline[x] - 0.01 < 0:
                APR_single_lever_sum -= baseline[x]
            else:
                APR_single_lever_sum -= 0.01
    
    # ID and save single lever moves
    for index, row in df.iloc[13:].iterrows():
        # Cash single lever move
        if (row['cash_delta'] == cash_min_enh) & (row['sum_of_moves'] == cash_min_enh) & (row['no_of_moves'] == 1):
            min_cash_single_lever = index
        elif (row['cash_delta'] == 2*cash_min_enh) & (row['sum_of_moves'] == 2*cash_min_enh) & (row['no_of_moves'] == 1):
            dbl_min_cash_single_lever = index
        # APR single lever move
        elif (round(row['APR_delta_sum'],3) == APR_single_lever_sum) & (round(row['sum_of_moves'],3) == APR_single_lever_sum) & (row['no_of_moves'] == 1):
            APR_single_lever = index
        # Combo single lever move
        elif (row['combo_delta'] == combo_min_enh) & (row['sum_of_moves'] == combo_min_enh) & (row['no_of_moves'] == 1):
            min_combo_single_lever = index
        elif (row['combo_delta'] == 2*combo_min_enh) & (row['sum_of_moves'] == 2*combo_min_enh) & (row['no_of_moves'] == 1):
            dbl_min_combo_single_lever = index
        # Lease single lever move
        elif (row['lease_delta'] == lease_min_enh) & (row['sum_of_moves'] == lease_min_enh) & (row['no_of_moves'] == 1):
            min_lease_single_lever = index
        elif (row['lease_delta'] == 2*lease_min_enh) & (row['sum_of_moves'] == 2*lease_min_enh) & (row['no_of_moves'] == 1):
            dbl_min_lease_single_lever = index
            
    # Add in single lever moves
    single_lever_indices = []
    if (min_cash_bool == True) & (cash_min_enh > 0):
        try:
            single_lever_indices.append(min_cash_single_lever)
        except:
            next
    if (dbl_min_cash_bool == True) & (cash_min_enh > 0):
        try:
            single_lever_indices.append(dbl_min_cash_single_lever)
        except:
            next
    if (APR_bool == True) & (APR_single_lever_sum != 0):
        try:
            single_lever_indices.append(APR_single_lever)
        except:
            next
    if (min_combo_bool == True) & (combo_min_enh > 0):
        try:
            single_lever_indices.append(min_combo_single_lever)
        except:
            next
    if (dbl_min_combo_bool == True) & (combo_min_enh > 0):
        try:
            single_lever_indices.append(dbl_min_combo_single_lever)
        except:
            next
    if (min_lease_bool == True) & (lease_min_enh > 0):
        try:
            single_lever_indices.append(min_lease_single_lever)
        except:
            next
    if (dbl_min_lease_bool == True) & (lease_min_enh > 0):
        try:
            single_lever_indices.append(dbl_min_lease_single_lever)
        except:
            next  
        
    single_lever_indices = [x + 1 for x in single_lever_indices]
    
    return single_lever_indices


# In[9]:


def find_single_lever_indices_old(df, baseline, delta_columns, APR_delta_columns, min_cash_bool, dbl_min_cash_bool, APR_bool, min_combo_bool, dbl_min_combo_bool, min_lease_bool, dbl_min_lease_bool):

    # Find lowest increment of enhancement
    cash_min_enh = df[(df['cash_delta'] > 0) & (df['no_of_moves'] == 1)]['cash_delta'].min()
    combo_min_enh = df[(df['combo_delta'] > 0) & (df['no_of_moves'] == 1)]['combo_delta'].min()
    lease_min_enh = df[(df['lease_delta'] > 0) & (df['no_of_moves'] == 1)]['lease_delta'].min()
    
    # Create sum across all moves
    df['sum_of_moves'] = df[delta_columns[0:8]].sum(axis=1)
    
    # Find 'combined_APR_delta' that equals -1 move across all terms
    APR_single_lever_sum = 0
    for x in APR_delta_columns:
            if baseline[x] - 0.01 < 0:
                APR_single_lever_sum -= baseline[x]
            else:
                APR_single_lever_sum -= 0.01
    
    # ID and save single lever moves
    for index, row in df.iloc[8:].iterrows():
        # Cash single lever move
        if (row['cash_delta'] == cash_min_enh) & (row['sum_of_moves'] == cash_min_enh) & (row['no_of_moves'] == 1):
            min_cash_single_lever = index
        elif (row['cash_delta'] == 2*cash_min_enh) & (row['sum_of_moves'] == 2*cash_min_enh) & (row['no_of_moves'] == 1):
            dbl_min_cash_single_lever = index
        # APR single lever move
        elif (round(row['APR_delta_sum'],3) == APR_single_lever_sum) & (round(row['sum_of_moves'],3) == APR_single_lever_sum) & (row['no_of_moves'] == 1):
            APR_single_lever = index
        # Combo single lever move
        elif (row['combo_delta'] == combo_min_enh) & (row['sum_of_moves'] == combo_min_enh) & (row['no_of_moves'] == 1):
            min_combo_single_lever = index
        elif (row['combo_delta'] == 2*combo_min_enh) & (row['sum_of_moves'] == 2*combo_min_enh) & (row['no_of_moves'] == 1):
            dbl_min_combo_single_lever = index
        # Lease single lever move
        elif (row['lease_delta'] == lease_min_enh) & (row['sum_of_moves'] == lease_min_enh) & (row['no_of_moves'] == 1):
            min_lease_single_lever = index
        elif (row['lease_delta'] == 2*lease_min_enh) & (row['sum_of_moves'] == 2*lease_min_enh) & (row['no_of_moves'] == 1):
            dbl_min_lease_single_lever = index
            
    # Add in single lever moves
    single_lever_indices = []
    if min_cash_bool == True:
        try:
            single_lever_indices.append(min_cash_single_lever)
        except:
            next
    if dbl_min_cash_bool == True:
        try:
            single_lever_indices.append(dbl_min_cash_single_lever)
        except:
            next
    if APR_bool == True:
        try:
            single_lever_indices.append(APR_single_lever)
        except:
            next
    if min_combo_bool == True:
        try:
            single_lever_indices.append(min_combo_single_lever)
        except:
            next
    if dbl_min_combo_bool == True:
        try:
            single_lever_indices.append(dbl_min_combo_single_lever)
        except:
            next
    if min_lease_bool == True:
        try:
            single_lever_indices.append(min_lease_single_lever)
        except:
            next
    if dbl_min_lease_bool == True:
        try:
            single_lever_indices.append(dbl_min_lease_single_lever)
        except:
            next  
        
    single_lever_indices = [x + 1 for x in single_lever_indices]
    
    return single_lever_indices


# In[10]:


def remove_std(df):
    reg_ex = 'std'
    reg_ex_filter = df['scenarios'].str.contains(reg_ex)
    df = df[~reg_ex_filter]
    
    return df


# In[11]:


def remove_CC_APR_diff_amts(df):
    CC_APR_nonzero_filter = (df['cash_delta'] != 0) & (df['combo_delta'] != 0)
    df['cash_combo_sum'] = round(df['cash_delta']/50.0)*50 + round(df['combo_delta']/50.0)*50
    cash_combo_sum_filter = df['cash_combo_sum'] == 0.0
    df = df[~(CC_APR_nonzero_filter & cash_combo_sum_filter)]
    
    return df


# # Master function

# In[1]:


def run_chooser_new_DSS(DSS_file_path, output_tab, model_order, min_cash_bool, dbl_min_cash_bool, APR_bool, min_combo_bool, dbl_min_combo_bool, min_lease_bool, dbl_min_lease_bool, std_bool, CC_combo_bool, min_spend, max_spend, max_no_of_scenarios):
    
    # Read calc data from Excel
    column_headers = ['scenarios','cash_delta','combo_delta','APR_36_delta','APR_48_delta','APR_60_delta','APR_72_delta','APR_84_delta','lease_delta','BC_delta','DC_delta','DFC_delta','DFL_delta','spend_delta','lift_delta','elasticity']
    columns_from_excel = 'F,G,I,J,K,L,M,N,T,U,V,X,Z,KU,LE,MK'
    delta_columns = ['cash_delta', 'combo_delta', 'APR_delta_sum','lease_delta', 'DC_delta', 'DFC_delta', 'BC_delta', 'DFL_delta', 'spend_delta', 'lift_delta']    
    APR_delta_columns = ['APR_36_delta','APR_48_delta','APR_60_delta','APR_72_delta','APR_84_delta']

    df = pd.read_excel(DSS_file_path, sheet_name='Calc', names=column_headers, skiprows=500*(model_order)-1, nrows=500, usecols=columns_from_excel)

    # Remove (#) and spaces at beginning and end of Scenario
    df['scenarios'] = df['scenarios'].str.replace('\\(.\\)','', regex=True).str.lstrip().str.rstrip()

    # Create no_of_moves column
    no_of_moves = 0
    df_no_of_moves = []
    for index, row in df.iterrows():
        no_of_moves = str(row['scenarios']).count('\n') + 1
        df_no_of_moves.append(no_of_moves)
    df['no_of_moves'] = df_no_of_moves

    # Insert 'APR_delta_sum' column
    df.insert(8, 'APR_delta_sum', df[APR_delta_columns].sum(axis=1))

    # Set baseline scenario to row 12 in Excel
    baseline = df.iloc[11]

    # Remove 'Market' terms
    for x in APR_delta_columns:
        try: 
            y = float(baseline[x])
        except:
            APR_delta_columns.remove(x)

    # Calculate delta columns
    for x in delta_columns:
        df_delta = []
        if x == 'lift_delta':
            for index, row in df.iterrows():
                try:
                    delta = row[x]/baseline[x] - 1
                    df_delta.append(delta)
                except:
                    delta = row[x] - baseline[x]
                    df_delta.append(delta)
        else:
            for index, row in df.iterrows():
                try:
                    delta = row[x] - baseline[x]
                    df_delta.append(delta)
                except:
                    df_delta.append(1000000)
        df[x] = df_delta
            
    # Find single lever moves
    single_lever_indices = find_single_lever_indices_new(df, baseline, delta_columns, APR_delta_columns, min_cash_bool, dbl_min_cash_bool, APR_bool, min_combo_bool, dbl_min_combo_bool, min_lease_bool, dbl_min_lease_bool)
        
    # Filter out NAs and duplicates
    df_filtered = df.dropna()
    df_filtered = df.iloc[13:]
    df_filtered = df_filtered.drop_duplicates(subset=['no_of_moves','lift_delta','spend_delta','sum_of_moves'], keep='first')
           
    # Find efficient frontier
    df_length = df.shape[0]
    eff_front = pd.DataFrame()

    for i in range(0,df_length,df_length):
        for k in range(13,df_length):
            current_spend = df['spend_delta'][k + i]
            current_lift = df['lift_delta'][k + i]
            for j in range(13,df_length):
                new_spend = df['spend_delta'][j + i]
                new_lift = df['lift_delta'][j + i]
                if (new_spend < current_spend) & (new_lift > current_lift):
                    break
                elif (np.isnan(df['spend_delta'][j + i])) & (j == df_length-1):
                    if np.isnan(df['spend_delta'][k + i]):
                        continue
                    else:
                        eff_front = eff_front.append(df.iloc[[k+i]])
    
    # Drop N/As & duplicate scenarios
    eff_front = eff_front.dropna()
    eff_front = eff_front.drop_duplicates(subset=['scenarios'], keep='first')
    
    # Create scoring system ranks
    
    # Elasticity 
    df_filtered['elasticity_score'] = abs(df_filtered['elasticity'])

    # Find increment of lowest single lever elasticity *0100
    max_single_lever_elast =  df.iloc[[6,7,9]]['elasticity'].max()
    delta = abs(max_single_lever_elast) 
        
    # No. of moves (each additional move -X increment)
    no_of_moves_dict = {
        1: (-delta)*0,
        2: (-2*delta)*0,
        3: (-3*delta)*0,
        4: (-4*delta)*0,
        5: (-5*delta)*0,
        6: (-6*delta)*0
    }
    df_filtered['no_of_moves_score'] = df_filtered['no_of_moves'].map(no_of_moves_dict)
    
    # Calculate scores
    eff_front_list = []
    for index, row in df_filtered.iterrows():
        eff_score_adj = 0
        # If on eff frontier, plus up
        if index in list(eff_front.index):
            eff_score_adj += delta
        # If -spend +lift, bring to top of list
        if (row['spend_delta'] < 0) & (row['lift_delta'] > 0):
            eff_score_adj += -row['elasticity_score'] + abs(row['spend_delta']*row['lift_delta']) + 100*delta
        # If +spend -lift, drop to bottom of list
        elif (row['spend_delta'] > 0) & (row['lift_delta'] < 0):
            eff_score_adj += -row['elasticity_score'] + -100*delta
        # If +spend +lift, move behind -spend +lift scenarios
        elif (row['spend_delta'] > 0) & (row['lift_delta'] > 0):
            eff_score_adj += 10*delta
        # If -spend -lift, cancel elastcity, invert order (lower elast = better), and shift below other scenarios
        elif (row['spend_delta'] < 0) & (row['lift_delta'] < 0):
            eff_score_adj += -row['elasticity_score'] + delta/(row['elasticity_score']) - delta
        # Else, do nothing
        else:
            eff_score_adj = 0
        eff_front_list.append(eff_score_adj)
    
    df_filtered['score_adj'] = eff_front_list

    # Total score
    df_filtered['total_score'] = df_filtered['elasticity_score'] + df_filtered['score_adj']
    df_filtered.sort_values('total_score', ascending=False, inplace=True)
    
    # Adjust total_score to punish scenarios too close to scenarios with higher scores
    spend_list = []
    score_adj_list = []
    spend_increment = (max_spend-min_spend)/max_no_of_scenarios
    for index, row in df_filtered.iterrows():
        spend_list.append(row['spend_delta'])
        spend_score_adj = 0
        for spend in spend_list:
            if (abs(row['spend_delta'] - spend)) >= spend_increment or (abs(row['spend_delta'] - spend) == 0):
                continue
            else:
                spend_score_adj += abs(spend_increment / (row['spend_delta'] - spend)) -100*delta
        score_adj_list.append(spend_score_adj)
    df_filtered['spend_score_adj'] = score_adj_list
    df_filtered['total_score_adj'] = df_filtered['total_score'] + df_filtered['spend_score_adj']

    df_filtered.sort_values('total_score_adj', ascending=False, inplace=True)

    #Apply remove_std filter
    if std_bool == True:
        df_filtered = remove_std(df_filtered)
    
    #Apply Mitsu CC&Combo filter
    if CC_combo_bool == True:
        df_filtered = remove_CC_APR_diff_amts(df_filtered)
    
    #Apply spend filters
    df_filtered = apply_spend_filters(df_filtered, min_spend, max_spend)
    
    #Convert to scenario row numbers in Excel
    DSS_scenarios = convert_to_DSS_scenarios(single_lever_indices, df_filtered, max_no_of_scenarios)
    
    #Write to output
    write_to_output_tab_new_DSS(DSS_file_path, output_tab, DSS_scenarios)


# In[13]:


def run_chooser_old_DSS(DSS_file_path, output_tab, model_order, min_cash_bool, dbl_min_cash_bool, APR_bool, min_combo_bool, dbl_min_combo_bool, min_lease_bool, dbl_min_lease_bool, std_bool, CC_combo_bool, min_spend, max_spend, max_no_of_scenarios):
            
    # Read calc data from Excel
    column_headers = ['scenarios','cash_delta','DC_delta','DFC_delta','DFL_delta','combo_delta','APR_36_delta','APR_48_delta','APR_60_delta','APR_72_delta','APR_84_delta','lease_delta','BC_delta','spend_delta','lift_delta','elasticity']
    columns_from_excel = 'D,E,F,G,H,Q,S,T,U,V,W,Y,Z,JD,JM,KU'
    delta_columns = ['cash_delta', 'combo_delta', 'APR_delta_sum','lease_delta', 'DC_delta', 'DFC_delta', 'BC_delta', 'DFL_delta', 'spend_delta', 'lift_delta']    
    APR_delta_columns = ['APR_36_delta','APR_48_delta','APR_60_delta','APR_72_delta','APR_84_delta']
        
    df = pd.read_excel(DSS_file_path, sheet_name='DSS_Calc', names=column_headers, skiprows=100*(model_order)-1, nrows=100, usecols=columns_from_excel)
            
    # Remove (#) and spaces at beginning and end of Scenario
    df['scenarios'] = df['scenarios'].str.replace('\\(.\\)','', regex=True).str.lstrip().str.rstrip()

    # Create no_of_moves column
    no_of_moves = 0
    df_no_of_moves = []
    for index, row in df.iterrows():
        no_of_moves = str(row['scenarios']).count('\n') + 1
        df_no_of_moves.append(no_of_moves)
    df['no_of_moves'] = df_no_of_moves
    
    # Set baseline scenario to row 2 in Excel
    baseline = df.iloc[1]
    
    # Remove 'Market' terms
    for x in APR_delta_columns:
        try: 
            y = float(baseline[x])
        except:
            APR_delta_columns.remove(x)
            
    # Insert 'APR_delta_sum' column
    df.insert(8, 'APR_delta_sum', df[APR_delta_columns].sum(axis=1))
    
    # Reset baseline to include 'APR_delta_sum' column
    baseline = df.iloc[1]
    
    # Calculate delta columns
    for x in delta_columns:
        df_delta = []
        if x == 'lift_delta':
            for index, row in df.iterrows():
                try:
                    delta = row[x]/baseline[x] - 1
                    df_delta.append(delta)
                except:
                    delta = row[x] - baseline[x]
                    df_delta.append(delta)
        else:
            for index, row in df.iterrows():
                try:
                    delta = row[x] - baseline[x]
                    df_delta.append(delta)
                except:
                    df_delta.append(1000000)
        df[x] = df_delta
    
    # Find single lever moves
    single_lever_indices = find_single_lever_indices_old(df, baseline, delta_columns, APR_delta_columns, min_cash_bool, dbl_min_cash_bool, APR_bool, min_combo_bool, dbl_min_combo_bool, min_lease_bool, dbl_min_lease_bool)

    # Filter out NAs, calibration, and duplicate scenarios
    df_filtered = df.dropna()
    df_filtered = df.iloc[8:]
    df_filtered = df_filtered.drop_duplicates(subset=['no_of_moves','lift_delta','spend_delta','sum_of_moves'], keep='first')
    # Filter out negative cash/finance/BC de-escalating scenarios 
    df_filtered_copy = df_filtered
    for index, row in df_filtered_copy.iterrows():
        if (row['cash_delta'] < 0) or (row['combo_delta'] < 0) or (row['BC_delta'] < 0) or (row['DC_delta'] < 0) or (row['DFC_delta'] < 0):
            df_filtered.drop(index,inplace=True)
        else:
            continue
    
    # Find efficient frontier
    df_length = df.shape[0]
    eff_front = pd.DataFrame()

    for i in range(0,df_length,df_length):
        for k in range(8,df_length):
            current_spend = df['spend_delta'][k + i]
            current_lift = df['lift_delta'][k + i]
            for j in range(8,df_length):
                new_spend = df['spend_delta'][j + i]
                new_lift = df['lift_delta'][j + i]
                if (new_spend < current_spend) & (new_lift > current_lift):
                    break
                elif (np.isnan(df['spend_delta'][j + i])) & (j == df_length-1):
                    if np.isnan(df['spend_delta'][k + i]):
                        continue
                    else:
                        eff_front = eff_front.append(df.iloc[[k+i]])
    # Drop N/As & duplicate scenarios
    eff_front = eff_front.dropna()
    eff_front = eff_front.drop_duplicates(subset=['elasticity','lift_delta','spend_delta'], keep='first')
    
    # Create scoring system ranks
    
    # Elasticity as % of max elasticity from eff_front 
    df_filtered['elasticity_score'] = abs(df_filtered['elasticity']/eff_front['elasticity'].max())

    # Find increment of lowest single lever elasticity
    min_single_lever_elast =  df.iloc[[2,3,4,5]]['elasticity'].min()
    delta = abs(min_single_lever_elast/eff_front['elasticity'].max())  
        
    # No. of moves (each additional move = -1 increment)
    no_of_moves_dict = {
        1: (-delta),
        2: (-2*delta),
        3: (-3*delta),
        4: (-4*delta),
        5: (-5*delta),
        6: (-6*delta)
    }
    df_filtered['no_of_moves_score'] = df_filtered['no_of_moves'].map(no_of_moves_dict)

    # Eff frontier score (if on frontier = +1 increment, if +spend & -lift then -1000)
    eff_front_list = []
    for index, row in df_filtered.iterrows():
        eff_score_adj = 0
        if index in list(eff_front.index):
            eff_front_list.append(delta)
        elif (row['spend_delta'] > 0) & (row['lift_delta'] < 0):
            eff_front_list.append(-delta*10)
        else:
            eff_front_list.append(0)
    df_filtered['eff_front_score'] = eff_front_list

    # Total score
    df_filtered['total_score'] = df_filtered['elasticity_score'] + df_filtered['no_of_moves_score'] + df_filtered['eff_front_score']
    df_filtered.sort_values('total_score', ascending=False, inplace=True)
    
    # Adjust total_score to punish scenarios too close to scenarios with higher scores
    spend_list = []
    score_adj_list = []
    for index, row in df_filtered.iterrows():
        spend_list.append(row['spend_delta'])
        spend_score_adj = 0
        for spend in spend_list:
            if (abs(row['spend_delta'] - spend) >= 50) or (abs(row['spend_delta'] - spend) == 0):
                continue
            else:
                spend_score_adj += abs(50 / (row['spend_delta'] - spend))  * -delta
        score_adj_list.append(spend_score_adj)

    df_filtered['spend_score_adj'] = score_adj_list
    df_filtered['total_score_adj'] = df_filtered['total_score'] - df_filtered['spend_score_adj']

    df_filtered.sort_values('total_score_adj', ascending=False, inplace=True)

    #Apply remove_std filter
    if std_bool == True:
        df_filtered = remove_std(df_filtered)
    
    #Apply Mitsu CC&Combo filter
    if CC_combo_bool == True:
        df_filtered = remove_CC_APR_diff_amts(df_filtered)
    
    #Apply spend filters
    df_filtered = apply_spend_filters(df_filtered, min_spend, max_spend)
    
    #Convert to scenario row numbers in Excel
    DSS_scenarios = convert_to_DSS_scenarios(single_lever_indices, df_filtered, max_no_of_scenarios)
    
    #Write to output
    write_to_output_tab_old_DSS(DSS_file_path, output_tab, DSS_scenarios)


# # For later ML use

# In[14]:


def read_chosen_scenarios_new_DSS(DSS_file_path, output_tab, df):
    #work in progress
    
    chosen_scenarios = pd.read_excel(DSS_file_path, names=['target_scenarios'], sheet_name=output_tab, usecols='C', skiprows=34, nrows=46)
    chosen_scenarios = chosen_scenarios.iloc[::2]
    chosen_scenarios = chosen_scenarios.dropna()
    chosen_scenarios = chosen_scenarios['target_scenarios'].astype('int')
    df['target_scenarios'] = 0
    for index, row in chosen_scenarios.iteritems():
        df['target_scenarios'].iloc[row-1] = 1


# In[15]:


def read_chosen_scenarios_old_DSS(DSS_file_path, output_tab, df):   
   #work in progress
   
   chosen_scenarios = pd.read_excel(DSS_file_path, names=['target_scenarios'], sheet_name=output_tab, usecols='C', skiprows=32, nrows=46)
   chosen_scenarios = chosen_scenarios.iloc[::2]
   chosen_scenarios = chosen_scenarios[~chosen_scenarios.isin(['x'])]
   chosen_scenarios = chosen_scenarios.dropna()
   chosen_scenarios = chosen_scenarios['target_scenarios'].astype('int')
   df['target_scenarios'] = 0
   for index, row in chosen_scenarios.iteritems():
       df['target_scenarios'].iloc[row-1] = 1

