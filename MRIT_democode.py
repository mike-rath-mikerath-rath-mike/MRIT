'''
McKinsey RES IRR Tool  (MRIT)
Created by Mike Rath, based off an orginal excel model built by Joss Sitter
9/2024
'''

# Import packages
from matplotlib import pyplot as plt
from scipy.optimize import fmin_slsqp
from pulp import *
from math import *
import decimal
import pandas as pd
import numpy as np
import numpy_financial as npf
import statistics  
import timeit

#############################################################################################################################################################################################################################
#############################################################################################################################################################################################################################
#############################################################################################################################################################################################################################
#############################################################################################################################################################################################################################
#############################################################################################################################################################################################################################





xls = pd.ExcelFile('/home/jovyan/shared/users/mike_rath_at_mckinsey.com/MRIT/20240930 MRIT_inputs TvW.xlsx')





#############################################################################################################################################################################################################################
#############################################################################################################################################################################################################################
#############################################################################################################################################################################################################################
#############################################################################################################################################################################################################################
#############################################################################################################################################################################################################################

# Functions

'''
MRIT Levered IRR function
'''

def levered_project_summary(mrit_inputs):    

    df_depr = dfs['Depreciation Schedule']
    df_propertytax = dfs['Property Tax Schedule']
    df_propertytax = df_propertytax.drop(columns=['Unit'])
    #df_unbundled_recs = dfs['unbundled REC database']
    #df_unbundled_recs = df_unbundled_recs.replace('',np.nan) 
    #df_unbundled_recs = df_unbundled_recs.dropna(axis = "columns", how = "any") 

    ##### Inputs #####
    ##### RES ####
    # Asset Characteristics 
    mw_ac = mrit_inputs['RES Capacity'].item()
    dcac_ratio = mrit_inputs['DC:AC Ratio'].item()
    mw_dc = mw_ac*dcac_ratio
    cod_year = str(int(mrit_inputs['RES COD Year'].item()))
    operating_lifetime = int(mrit_inputs['RES Operating Life'].item()) # years
    cap_fac = mrit_inputs['RES Capacity Factor'].item() # %
    mwh_p50 = mw_ac*cap_fac*8760 # replace with mwh_system
    mwh_per_mwp = mwh_p50/mw_dc
    degradation_per = mrit_inputs['RES Degradation'].item() # p.a.

    # Grid characteristics 
    hub_price = mrit_inputs['LMP to Hub Price Basis'].item() #'base_case'
    asset_curtailment = mrit_inputs['Asset curtailment'].item() #'base_case'

    # Revenue 
    ppa_price_input = mrit_inputs['PPA Price'].item() # $/MWh
    ppa_tenor = mrit_inputs['PPA Tenor'].item() # years
    contracted_per = mrit_inputs['Percentage of energy contracted'].item() 
    ppa_escalator = mrit_inputs['PPA escalator'].item()

    merchant_energy_price_input = mrit_inputs['Merchant Prices Energy'].item() # $/MWh
    merchant_rec = mrit_inputs['REC Price'].item() # $/MWh
    merchant_escalator = mrit_inputs['Merchant escalator'].item()
    rec_escalator = mrit_inputs['REC escalator'].item()

    inflation_rate = float(mrit_inputs['Inflation rate'].item()) #'base_case'

    # Capex
    plant_construction_cost = mrit_inputs['RES Capital Cost'].item() # $/W_dc
    total_plant_construction_cost = plant_construction_cost*mw_dc*1000000

    developer_fee = mrit_inputs['Developer Fee'].item() # $/W_ac
    total_developer_fee = developer_fee*mw_ac

    res_total_capex = total_plant_construction_cost+total_developer_fee

    # ITC
    itc = mrit_inputs['RES ITC'].item()
    itc_amount = res_total_capex*itc
    depreciable_basic_reduction = mrit_inputs['Depreciable Basis Reduction'].item()
    depreciable_basic_reduction_amount = itc_amount*depreciable_basic_reduction
    res_total_depreciable_amount = res_total_capex-depreciable_basic_reduction_amount # total decpriable basis

    # PTC
    res_ptc = mrit_inputs['RES PTC'].item()
    res_ptc_inflation = mrit_inputs['RES PTC inflation'].item()
    res_ptc_length = mrit_inputs['RES PTC length'].item()
    res_ptc_value = mrit_inputs['RES PTC Sale value'].item()

    # Operating Expenditures
    royalties = mrit_inputs['Royalties'].item() # %
    insurance = mrit_inputs['Insurance'].item()
    om_annual = mrit_inputs['RES O&M Cost'].item() # $/MW
    om_escalator = mrit_inputs['RES O&M escalator'].item()
    lease_cost = mrit_inputs['Lease'].item() # $/acre
    acres_per_mwac = mrit_inputs['Acres / MWac'].item()
    lease_land = acres_per_mwac*mw_ac
    total_lease_amount = lease_land*lease_cost
    lease_escalator = mrit_inputs['Lease escalator'].item()

    # Debt
    contract_debt_coverage_ratio = mrit_inputs['Contracted Debt Coverage Ratio'].item()
    merchant_debt_coverage_ratio = mrit_inputs['Merchant Debt Coverage Ratio'].item()
    ptc_debt_coverage_ratio = mrit_inputs['PTC Debt Coverage Ratio'].item()
    debt_tenor = mrit_inputs['Debt Tenor'].item() # year
    interest = mrit_inputs['Interest'].item()

    # Tax equity
    #structure = x[31] #'yield flip'
    target_yield = mrit_inputs['Target Yield'].item()
    target_flip_yr = int(mrit_inputs['Target Flip Year'].item())
    te_preflip_tax_allocation = mrit_inputs['Pre-Flip Tax Allocation'].item()
    te_postflip_tax_allocation = mrit_inputs['Post-Flip Tax Allocation'].item()
    te_preflip_cash_allocation = mrit_inputs['Pre-Flip Cash Allocation'].item()
    te_postflip_cash_allocation = mrit_inputs['Post-Flip Cash Allocation'].item()

    # project owners
    po_preflip_tax_allocation = 1 - te_preflip_tax_allocation
    po_postflip_tax_allocation = 1 - te_postflip_tax_allocation
    po_preflip_cash_allocation = 1 - te_preflip_cash_allocation
    po_postflip_cash_allocation = 1 - te_postflip_cash_allocation

    corporate_income_tax_rate = mrit_inputs['Corporate Income Tax Rate'].item()
    nol_decution_max = mrit_inputs['NOL Deduction Max'].item()
    min_project_owner_equity = mrit_inputs['Min Project Owner Equity Precent'].item()


    #### Storage ####
    switch = int(mrit_inputs['Storage Attachment'].item())

    # Asset Characteristics 
    storage_cod_year = switch*mrit_inputs['Storage COD'].item()
    if storage_cod_year <= float(cod_year):
        storage_cod_year = float(cod_year)

    #storage_cod_year = 2027.0

    storage_lifetime = switch*mrit_inputs['Storage Lifetime'].item() # years
    storage_cap_MW = switch*mrit_inputs['Storage Power Capacity'].item() # MW
    storage_cap_MWh = switch*mrit_inputs['Storage Energy Capacity'].item() # MWh
    power_capex = switch*mrit_inputs['Storage Power Capital Cost'].item() # $/kW
    energy_capex = switch*mrit_inputs['Storage Energy Capital Cost'].item() # $/kWh
    rti = switch*mrit_inputs['Storage RTE'].item() # %
    storage_degradation = switch*mrit_inputs['Storage Degradation'].item()
    cycles = switch*mrit_inputs['Cycles per year'].item() 

    # Revenue 
    storage_arbitrage_payment = switch*mrit_inputs['Storage Arbitrage'].item() # $/kW
    storage_arbitrage_amount = switch*mrit_inputs['Storage Arbitrage Amount'].item() # % $/kW
    storage_arbitrage_escalator = switch*mrit_inputs['Storage Arbitrage Escalator'].item() # %
    storage_ancillary_payments = switch*mrit_inputs['Storage Ancillary Payments'].item() # $/kW
    storage_ancillary_amount = switch*mrit_inputs['Storage Ancillary Amount'].item() # % $/kW
    storage_ancillary_escalator = switch*mrit_inputs['Storage Ancillary Escalator'].item() # %
    storage_capacity_payments = switch*mrit_inputs['Storage Capacity Payments'].item() # $/kW
    storage_capacity_amount = switch*mrit_inputs['Storage Capacity Payment Amount'].item() # % $/kW
    storage_capacity_escalator = switch*mrit_inputs['Storage Capacity Payment Escalator'].item() # %
    tolling_amount = switch*mrit_inputs['Tolling Agreement Amount'].item() # % $/kW
    tolling_payment = switch*mrit_inputs['Tolling Agreement Payments'].item() # $/kW
    tolling_escalator = switch*mrit_inputs['Tolling Payments escalator'].item() # %
    tolling_tenor = switch*mrit_inputs['Tolling Agreement Tenor'].item() # years

    # Capex
    storage_capex_excluding_rti = switch*float((storage_cap_MW*power_capex)+(storage_cap_MWh*energy_capex))
    if rti != 0:
        storage_capex = (storage_capex_excluding_rti/rti)
    else: storage_capex = 0

    # ITC
    storage_itc = switch*mrit_inputs['Storage ITC'].item() # %
    storage_itc = switch*(storage_capex*storage_itc)
    storage_depreciable_basic_reduction_amount = switch*storage_itc*depreciable_basic_reduction
    storage_total_depreciable_amount = switch*storage_capex-storage_depreciable_basic_reduction_amount # total decpriable basis

    # PTC
    storage_ptc = mrit_inputs['Storage PTC'].item()
    storage_ptc_inflation = mrit_inputs['Storage PTC inflation'].item()
    storage_ptc_length = mrit_inputs['Storage PTC length'].item()
    storage_ptc_value = mrit_inputs['Storage PTC Sale value'].item()


    # Operating Expenditures
    storage_opex = switch*mrit_inputs['Storage O&M'].item() # $/kW
    storage_opex_escalator = switch*mrit_inputs['Storage O&M Escalator'].item() # %


    #### Code ####
    # Timeline
    start_year = int(mrit_inputs['RES COD Year'].item())-1
    year_array = [] # str
    year_array_float_list = [] # num
    operating_period_list = []
    days_in_year_list = []
    lease_factor_list = []
    om_factor_list = []
    inflation_factor_list = []
    operating_flag_list = []

    for i in range(operating_lifetime+1):
        year_array.append(str(start_year+i))
        year_array_float_list.append(start_year+i)
        operating_period_list.append(i)
        if year_array_float_list[i]%4 == 0:
            days_in_year_list.append(366)
        else: days_in_year_list.append(365)
        lease_factor_list.append(i*lease_escalator)
        om_factor_list.append(i*om_escalator)
        if i == 0 or i == 1:
            inflation_factor_list.append(1)
        else: inflation_factor_list.append(1+operating_period_list[i]*inflation_rate)
        if i == 0: 
            operating_flag_list.append(0)
        else: operating_flag_list.append(1)

    # storage Timeline
    storage_startyear = storage_cod_year - float(cod_year)
    storage_life_measure = int(storage_startyear) + int(storage_lifetime)
    if storage_life_measure <= len(operating_period_list):
        storage_operating_flag_list = [0]*(int(storage_startyear)+1) + [1]*int(storage_lifetime) +[0]*(int(len(operating_period_list)-1)-int(storage_life_measure))
    else:
        storage_operating_flag_list = [0]*(int(storage_startyear)+1) + [1]*int(storage_lifetime) +[0]*(int(len(operating_period_list)-1)-int(storage_life_measure))
        storage_operating_flag_list = storage_operating_flag_list[:int(len(operating_period_list))] 

    j = 0
    storage_om_factor_list = []
    tolling_factor_list = []
    storage_operation_period_list = [] 
    for i in range(len(storage_operating_flag_list)):
        if storage_operating_flag_list[i] == 1:
            j += 1
            storage_om_factor_list.append(-storage_opex_escalator+(1+storage_opex_escalator)**j)
            tolling_factor_list.append(-tolling_escalator+(1+tolling_escalator)**j)
            storage_operation_period_list.append(j)
        else:
            storage_om_factor_list.append(0)
            tolling_factor_list.append(0)
            storage_operation_period_list.append(0)
    j = 0
    for i in range(len(storage_operating_flag_list)):
        if storage_operating_flag_list[i] == 1:
            j += 1
            if j > tolling_tenor:
                tolling_factor_list[i] = 0

    timeline_df = pd.DataFrame(np.column_stack([np.array(year_array), 
                                        np.array(operating_period_list, dtype=float), 
                                        np.array(days_in_year_list, dtype=float), 
                                        np.array(lease_factor_list, dtype=float), 
                                        np.array(om_factor_list, dtype=float),
                                        np.array(inflation_factor_list, dtype=float),
                                        np.array(operating_flag_list, dtype=float), 
                                        np.array(storage_om_factor_list, dtype=float), 
                                        np.array(tolling_factor_list, dtype=float),
                                        np.array(storage_operation_period_list, dtype=float),
                                        np.array(storage_operating_flag_list, dtype=float)]),
                               columns = ['year_array', 
                                          'operating_period', 
                                          'days_in_year', 
                                          'lease_factor',
                                          'om_factor',
                                          'inflation_factor',
                                          'operating_flag',
                                          'storage_om_factor',
                                          'tolling_factor',
                                          'storage_operation_period',
                                          'storage_operating_flag'])
    timeLine = timeline_df.set_index('year_array').T.to_dict('list')
    (operating_period, days_in_year, lease_factor, om_factor, inflation_factor, operating_flag, storage_om_factor, tolling_factor, storage_operation_period, storage_operating_flag) = splitDict(timeLine)
    for k, v in operating_period.items(): operating_period[k] = float(v)
    for k, v in days_in_year.items(): days_in_year[k] = float(v)
    for k, v in lease_factor.items(): lease_factor[k] = float(v)
    for k, v in om_factor.items(): om_factor[k] = float(v)
    for k, v in inflation_factor.items(): inflation_factor[k] = float(v)
    for k, v in operating_flag.items(): operating_flag[k] = float(v)
    for k, v in storage_om_factor.items(): storage_om_factor[k] = float(v)
    for k, v in tolling_factor.items(): tolling_factor[k] = float(v)
    for k, v in storage_operation_period.items(): storage_operation_period[k] = float(v)
    for k, v in storage_operating_flag.items(): storage_operating_flag[k] = float(v)


    # Production
    gross_generation_list = []   # MWh annual
    module_degradation_list = [] # MWh annual
    curtailment_list = []  # MWh annual
    net_generation_list = []
    annual_storage_degradation_list = [] # % scaler
    storage_dispatchMW_list = [] # MW year
    storage_dispatchMWh_list = [] # MW year

    for i in timeLine.keys():
        loop_gen = operating_flag[i]*cap_fac*mw_ac*days_in_year[i]*24
        loop_deg = operating_flag[i]*-degradation_per*operating_period[i]*loop_gen
        loop_cur = asset_curtailment*(loop_gen+loop_deg)
        loop_stor_deg = storage_operating_flag[i]*(storage_operation_period[i]*storage_degradation-storage_degradation) # adding deg back in (by subtraction) accounts for no degridation in the first year
        gross_generation_list.append(loop_gen)
        module_degradation_list.append(loop_deg)
        curtailment_list.append(loop_cur)
        net_generation_list.append(loop_gen+loop_deg+loop_cur)
        annual_storage_degradation_list.append(1-loop_stor_deg)
        storage_dispatchMW_list.append(storage_operating_flag[i]*(rti*(storage_cap_MW*cycles)*(1-loop_stor_deg)))
        storage_dispatchMWh_list.append(storage_operating_flag[i]*(rti*(storage_cap_MWh*cycles)*(1-loop_stor_deg)))

    production_df = pd.DataFrame(np.column_stack([year_array,
                                                  gross_generation_list, 
                                                  module_degradation_list, 
                                                  curtailment_list, 
                                                  net_generation_list,
                                                  annual_storage_degradation_list,
                                                  storage_dispatchMW_list,
                                                  storage_dispatchMWh_list]), 
                                 columns = ['year_array',
                                            'gross_generation', 
                                            'module_degradation', 
                                            'curtailment', 
                                            'net_generation',
                                            'annual_storage_degradation',
                                            'storage_dispatchMW',
                                            'storage_dispatchMWh'])
    production = production_df.set_index('year_array').T.to_dict('list')
    (gross_generation, module_degradation, curtailment, net_generation,annual_storage_degradation,storage_dispatchMW,storage_dispatchMWh) = splitDict(production)
    for k, v in gross_generation.items(): gross_generation[k] = float(v)
    for k, v in module_degradation.items(): module_degradation[k] = float(v)
    for k, v in curtailment.items(): curtailment[k] = float(v)
    for k, v in net_generation.items(): net_generation[k] = float(v)
    for k, v in annual_storage_degradation.items(): annual_storage_degradation[k] = float(v)
    for k, v in storage_dispatchMW.items(): storage_dispatchMW[k] = float(v)
    for k, v in storage_dispatchMWh.items(): storage_dispatchMWh[k] = float(v)

    # Revenue, Energy & REC Prices
    ppa_price_list = []   # $/MWh annual
    merchant_energy_price_list = [] # $/MWh annual
    unbundled_rec_price_list = []  # $/MWh annual

    ppa_rev_list = []   # $/MWh annual
    merchant_energy_rev_list = [] # $/MWh annual
    unbundled_rev_price_list = []  # $/MWh annual

    storage_arb_rev_list = [] # $/MW annual
    storage_anc_rev_list = [] # $/MW annual
    storage_cap_rev_list = [] # $/MW annual
    tolling_amount_rev_list = [] # $/MW annual
    storage_rev_list = []

    contract_rev_split_list = []
    merchant_rev_split_list = []
    total_rev_list = []

    for i in timeLine.keys():
        if operating_period[i] < ppa_tenor:
            loop_ppa_price = operating_flag[i]*ppa_price_input*((ppa_escalator+1)**operating_period[i])
        else: loop_ppa_price = 0       

        loop_storage_arb_rev = storage_operating_flag[i]*(1000*storage_arbitrage_amount*storage_arbitrage_payment*storage_cap_MW*(-storage_arbitrage_escalator+(1+storage_arbitrage_escalator)**storage_operation_period[i])) # 1000 is for to turn $/kW ot $/Mw
        loop_storage_anc_rev = storage_operating_flag[i]*(1000*storage_ancillary_amount*storage_ancillary_payments*storage_cap_MW*(-storage_ancillary_escalator+(1+storage_ancillary_escalator)**storage_operation_period[i])) # 1000 is for to turn $/kW ot $/Mw
        loop_storage_cap_rev = storage_operating_flag[i]*(1000*storage_capacity_amount*storage_capacity_payments*storage_cap_MW*(-storage_capacity_escalator+(1+storage_capacity_escalator)**storage_operation_period[i])) # 1000 is for to turn $/kW ot $/Mw
        loop_tolling_amount_rev = storage_operating_flag[i]*(1000*tolling_amount*tolling_payment*storage_cap_MW*tolling_factor[i])
        loop_storage_rev = loop_storage_arb_rev+loop_storage_anc_rev+loop_storage_cap_rev+loop_tolling_amount_rev

        loop_merchant_price = operating_flag[i]*merchant_energy_price_input*((merchant_escalator+1)**operating_period[i])
        loop_rec_price = operating_flag[i]*merchant_rec*((rec_escalator+1)**operating_period[i])
        ppa_price_list.append(loop_ppa_price)
        merchant_energy_price_list.append(loop_merchant_price)
        unbundled_rec_price_list.append(loop_rec_price)
        # Revenue ($)
        loop_ppa_rev = net_generation[i]*loop_ppa_price*contracted_per
        if operating_period[i] < ppa_tenor:
            loop_mercahnt_rev = (1- contracted_per)*net_generation[i]*loop_merchant_price
        else: loop_mercahnt_rev = net_generation[i]*loop_merchant_price    
        if operating_period[i] < ppa_tenor:
            loop_rec_rev = (1- contracted_per)*net_generation[i]*loop_rec_price
        else: loop_rec_rev = net_generation[i]*loop_rec_price
        loop_total_rev = float(loop_ppa_rev+loop_mercahnt_rev+loop_rec_rev+loop_storage_rev)
        if loop_total_rev == 0:
            loop_contract_rev_split = 0.5
            loop_merchant_rev_split = 0.5
        else:
            loop_contract_rev_split = (loop_ppa_rev+loop_rec_rev+loop_storage_rev-loop_storage_arb_rev)/(loop_total_rev)
            loop_merchant_rev_split = (loop_mercahnt_rev+loop_storage_arb_rev)/(loop_total_rev)
        ppa_rev_list.append(loop_ppa_rev)
        merchant_energy_rev_list.append(loop_mercahnt_rev)
        unbundled_rev_price_list.append(loop_rec_rev)
        storage_arb_rev_list.append(loop_storage_arb_rev)
        storage_anc_rev_list.append(loop_storage_anc_rev)
        storage_cap_rev_list.append(loop_storage_cap_rev)
        tolling_amount_rev_list.append(loop_tolling_amount_rev)
        storage_rev_list.append(loop_storage_rev)
        contract_rev_split_list.append(loop_contract_rev_split)
        merchant_rev_split_list.append(loop_merchant_rev_split)
        total_rev_list.append(loop_total_rev) 

    price_rev_df = pd.DataFrame(np.column_stack([year_array,
                                                 ppa_price_list, 
                                                 merchant_energy_price_list, 
                                                 unbundled_rec_price_list, 
                                                 ppa_rev_list,
                                                 merchant_energy_rev_list,
                                                 unbundled_rev_price_list,
                                                 storage_arb_rev_list, 
                                                 storage_anc_rev_list, 
                                                 storage_cap_rev_list, 
                                                 tolling_amount_rev_list, 
                                                 storage_rev_list,
                                                 contract_rev_split_list,
                                                 merchant_rev_split_list,
                                                 total_rev_list]), 
                                 columns = ['year_array',
                                            'ppa_price', 
                                            'merchant_energy_price', 
                                            'unbundled_rec_price', 
                                            'ppa_rev',
                                            'merchant_energy_rev',
                                            'unbundled_rev_price',
                                            'storage_arb_rev', 
                                            'storage_anc_rev', 
                                            'storage_cap_rev', 
                                            'tolling_amount_rev', 
                                            'storage_rev',
                                            'contract_rev_split',
                                            'merchant_rev_split',
                                            'total_rev'])
    price_rev = price_rev_df.set_index('year_array').T.to_dict('list')
    (ppa_price, merchant_energy_price, unbundled_rec_price, ppa_rev, merchant_energy_rev, unbundled_rev_price, storage_arb_rev, storage_anc_rev, 
     storage_cap_rev, tolling_amount_rev, storage_rev, contract_rev_split, merchant_rev_split, total_rev) = splitDict(price_rev)
    for k, v in ppa_price.items(): ppa_price[k] = float(v)
    for k, v in merchant_energy_price.items(): merchant_energy_price[k] = float(v)
    for k, v in unbundled_rec_price.items(): unbundled_rec_price[k] = float(v)
    for k, v in ppa_rev.items(): ppa_rev[k] = float(v)
    for k, v in merchant_energy_rev.items(): merchant_energy_rev[k] = float(v)
    for k, v in unbundled_rev_price.items(): unbundled_rev_price[k] = float(v)
    for k, v in storage_arb_rev.items(): storage_arb_rev[k] = float(v)
    for k, v in storage_anc_rev.items(): storage_anc_rev[k] = float(v)
    for k, v in storage_cap_rev.items(): storage_cap_rev[k] = float(v)
    for k, v in tolling_amount_rev.items(): tolling_amount_rev[k] = float(v)
    for k, v in storage_rev.items(): storage_rev[k] = float(v)
    for k, v in contract_rev_split.items(): contract_rev_split[k] = float(v)
    for k, v in merchant_rev_split.items(): merchant_rev_split[k] = float(v)
    for k, v in total_rev.items(): total_rev[k] = float(v)

    # Operating Expenditures
    price_basis_list = []
    royalties_payment_list = []
    insurance_payment_list = []
    om_expenditure_list = []
    lease_expense_list = []
    property_taxes_list = []
    storage_om_expenditure_list = []
    total_operating_expenditures_list = []

    # set them to be different arrays

    depr_storage_list = list(df_depr['Storage Depreciation MACRS'].fillna(0))+[0.0]*1000
    depr_storage_timematch_list = [0]*(int(storage_cod_year)-int(cod_year)+1)+depr_storage_list

    depr_res_list = list(df_depr['RES Depreciation MACRS'].fillna(0))+[0.0]*1000

    #df_propertytax = dfs['property_tax_schedule']
    propertytax_list = list(df_propertytax['Annual Property Tax'].fillna(0))+[0.0]*1000

    depr_propertytax_temp = pd.DataFrame(np.column_stack([np.array(year_array), 
                                         np.array(depr_res_list[0:(len(year_array))], dtype=float),  
                                         np.array(depr_storage_timematch_list[0:(len(year_array))], dtype=float),                
                                         np.array(propertytax_list[0:(len(year_array))], dtype=float)]), 
                               columns = ['year_array', 
                                          'res_depreciation_MACRS',
                                          'storage_depreciation_MACRS',  
                                          'property_taxes_annual'])
    deprPropertytax = depr_propertytax_temp.set_index('year_array').T.to_dict('list')
    (res_depreciation_MACRS, storage_depreciation_MACRS, property_taxes_annual) = splitDict(deprPropertytax)
    for k, v in res_depreciation_MACRS.items(): res_depreciation_MACRS[k] = float(v)
    for k, v in storage_depreciation_MACRS.items(): storage_depreciation_MACRS[k] = float(v)
    for k, v in property_taxes_annual.items(): property_taxes_annual[k] = float(v)


    for i in timeLine.keys():
        loop_price_basis = hub_price*net_generation[i]
        loop_royalties = operating_flag[i]*royalties*total_rev[i]
        loop_insurance = operating_flag[i]*insurance*total_rev[i]
        loop_om_expenditure = operating_flag[i]*mw_ac*om_annual*inflation_factor[i]*(1+om_factor[i]) 
        loop_lease_expense = operating_flag[i]*(1+lease_factor[i])*total_lease_amount
        loop_property_taxes = operating_flag[i]*(mw_ac*acres_per_mwac)*property_taxes_annual[i]
        loop_storage_om_expenditure = storage_operating_flag[i]*storage_cap_MW*storage_opex*inflation_factor[i]*storage_om_factor[i]
        price_basis_list.append(loop_price_basis)
        royalties_payment_list.append(loop_royalties)
        insurance_payment_list.append(loop_insurance)
        om_expenditure_list.append(loop_om_expenditure)
        lease_expense_list.append(loop_lease_expense)
        property_taxes_list.append(loop_property_taxes)
        storage_om_expenditure_list.append(loop_storage_om_expenditure)
        total_operating_expenditures_list.append(loop_price_basis+
                                                 loop_royalties+
                                                 loop_insurance+
                                                 loop_om_expenditure+
                                                 loop_lease_expense+
                                                 loop_property_taxes+
                                                 loop_storage_om_expenditure)

    depr_propertytax_df = pd.DataFrame(np.column_stack([year_array,
                                                 price_basis_list, 
                                                 royalties_payment_list, 
                                                 insurance_payment_list, 
                                                 om_expenditure_list,
                                                 lease_expense_list,
                                                 property_taxes_list,
                                                 storage_om_expenditure_list,
                                                 total_operating_expenditures_list]), 
                                 columns = ['year_array',
                                            'price_basis', 
                                            'royalties_payment', 
                                            'insurance_payment', 
                                            'om_expenditure',
                                            'lease_expense',
                                            'property_taxes',
                                            'storage_om_expenditure',
                                            'total_operating_expenditures'])
    depr_propertytax = depr_propertytax_df.set_index('year_array').T.to_dict('list')
    (price_basis, royalties_payment, insurance_payment, om_expenditure, lease_expense, property_taxes, storage_om_expenditure, total_operating_expenditures) = splitDict(depr_propertytax)
    for k, v in price_basis.items(): price_basis[k] = float(v)
    for k, v in royalties_payment.items(): royalties_payment[k] = float(v)
    for k, v in insurance_payment.items(): insurance_payment[k] = float(v)
    for k, v in om_expenditure.items(): om_expenditure[k] = float(v)
    for k, v in lease_expense.items(): lease_expense[k] = float(v)
    for k, v in property_taxes.items(): property_taxes[k] = float(v)
    for k, v in storage_om_expenditure.items(): storage_om_expenditure[k] = float(v)
    for k, v in total_operating_expenditures.items(): total_operating_expenditures[k] = float(v)

    # PTC
    res_production_tax_credit_list = [] # ($/kWh)
    res_total_PTC_amount_list  = [] # ($m)
    res_PTC_value_list  = [] # ($m)
    storage_production_tax_credit_list = [] # ($/kWh)
    storage_total_PTC_amount_list  = [] # ($m)
    storage_PTC_value_list  = [] # ($m)
    totl_PTC_amount_list = [] # ($m)
    total_PTC_value_list = [] # ($m)

    for i in timeLine.keys():
        if operating_period[i] == 0:
            loop_res_production_tax_credit = 0
        elif operating_period[i] <= res_ptc_length:
            loop_res_production_tax_credit = ((res_ptc*10)*(-res_ptc_inflation+(1+res_ptc_inflation)**operating_period[i])) # /100 to convert c/Kwh ot $/kWh and *1000 to convert $/kWh to $/MWh
        else: loop_res_production_tax_credit = 0

        if storage_operation_period[i] == 0:
            loop_storage_production_tax_credit = 0
        elif storage_operation_period[i] <= storage_ptc_length:
            loop_storage_production_tax_credit = storage_operating_flag[i]*((storage_ptc*10))*(-storage_ptc_inflation+(1+storage_ptc_inflation)**storage_operation_period[i]) # /100 to convert c/Kwh ot $/kWh and *1000 to convert $/kWh to MWh
        else: loop_storage_production_tax_credit = 0

        loop_res_total_PTC_amount = loop_res_production_tax_credit*net_generation[i]
        loop_storage_total_PTC_amount = loop_storage_production_tax_credit*storage_dispatchMWh[i]
        loop_total_PTC_amount = loop_res_total_PTC_amount + loop_storage_total_PTC_amount

        loop_res_PTC_transferability_value = loop_res_total_PTC_amount*res_ptc_value
        loop_storage_PTC_transferability_value = loop_storage_total_PTC_amount*storage_ptc_value
        loop_total_PTC_transferability_value = loop_res_PTC_transferability_value + loop_storage_PTC_transferability_value

        res_production_tax_credit_list.append(loop_res_production_tax_credit)
        res_total_PTC_amount_list.append(loop_res_total_PTC_amount)
        res_PTC_value_list.append(loop_res_PTC_transferability_value)
        storage_production_tax_credit_list.append(loop_storage_production_tax_credit)
        storage_total_PTC_amount_list.append(loop_storage_total_PTC_amount)
        storage_PTC_value_list.append(loop_storage_PTC_transferability_value)
        totl_PTC_amount_list.append(loop_total_PTC_amount)
        total_PTC_value_list.append(loop_total_PTC_transferability_value)

    timeline_df = pd.DataFrame(np.column_stack([np.array(year_array), 
                                        np.array(res_production_tax_credit_list, dtype=float), 
                                        np.array(res_total_PTC_amount_list, dtype=float), 
                                        np.array(res_PTC_value_list, dtype=float), 
                                        np.array(storage_production_tax_credit_list, dtype=float),
                                        np.array(storage_total_PTC_amount_list, dtype=float),
                                        np.array(storage_PTC_value_list, dtype=float), 
                                        np.array(totl_PTC_amount_list, dtype=float),
                                        np.array(total_PTC_value_list, dtype=float)]),
                               columns = ['year_array', 
                                          'res_production_tax_credit', 
                                          'res_total_PTC_amount', 
                                          'res_PTC_value',
                                          'storage_production_tax_credit',
                                          'storage_total_PTC_amount',
                                          'storage_PTC_value',
                                          'totl_PTC_amount',
                                          'total_PTC_value'])
    timeLine = timeline_df.set_index('year_array').T.to_dict('list')
    (res_production_tax_credit, res_total_PTC_amount, res_PTC_value, storage_production_tax_credit, storage_total_PTC_amount, storage_PTC_value, totl_PTC_amount, total_PTC_value) = splitDict(timeLine)
    for k, v in res_production_tax_credit.items(): res_production_tax_credit[k] = float(v)
    for k, v in res_total_PTC_amount.items(): res_total_PTC_amount[k] = float(v)
    for k, v in res_PTC_value.items(): res_PTC_value[k] = float(v)
    for k, v in storage_production_tax_credit.items(): storage_production_tax_credit[k] = float(v)
    for k, v in storage_total_PTC_amount.items(): storage_total_PTC_amount[k] = float(v)
    for k, v in storage_PTC_value.items(): storage_PTC_value[k] = float(v)
    for k, v in totl_PTC_amount.items(): totl_PTC_amount[k] = float(v)
    for k, v in total_PTC_value.items(): total_PTC_value[k] = float(v)


    # EBITDA and cash flow calculations 
    ebitda_list = []
    macrs_depreciation_list = []
    ebit_list = []
    plus_depreciation_list = []
    asset_cashflow_list = []

    for i in timeLine.keys():
        loop_ebitda = total_rev[i]-total_operating_expenditures[i]
        loop_macrs_depreciation = -(res_total_depreciable_amount*res_depreciation_MACRS[i]+storage_total_depreciable_amount*storage_depreciation_MACRS[i])
        loop_ebit = loop_macrs_depreciation+loop_ebitda
        loop_asset_chashflow = loop_ebit-loop_macrs_depreciation
        ebitda_list.append(loop_ebitda)
        macrs_depreciation_list.append(loop_macrs_depreciation)
        ebit_list.append(loop_ebit)
        plus_depreciation_list.append(-loop_macrs_depreciation)
        asset_cashflow_list.append(loop_asset_chashflow)

    ebitda_info_df = pd.DataFrame(np.column_stack([year_array,
                                                 ebitda_list, 
                                                 macrs_depreciation_list, 
                                                 ebit_list,
                                                 plus_depreciation_list,
                                                 asset_cashflow_list]), 
                                 columns = ['year_array',
                                            'ebitda', 
                                            'macrs_depreciation', 
                                            'ebit',
                                            'plus_depreciation',
                                            'asset_cashflow'])
    ebitda_info = ebitda_info_df.set_index('year_array').T.to_dict('list')
    (ebitda, macrs_depreciation, ebit, plus_depreciation, asset_cashflow) = splitDict(ebitda_info)
    for k, v in ebitda.items(): ebitda[k] = float(v)
    for k, v in macrs_depreciation.items(): macrs_depreciation[k] = float(v)
    for k, v in ebit.items(): ebit[k] = float(v)
    for k, v in plus_depreciation.items(): plus_depreciation[k] = float(v)
    for k, v in asset_cashflow.items(): asset_cashflow[k] = float(v)

    # Projecttax benefits/ (liabilities)
    tax_credit_list = []
    tax_benefit_list = []
    total_benefit_list = []
    for i in timeLine.keys():
        if i == cod_year: 
            loop_res_tax_credit = itc_amount
        else : loop_res_tax_credit = 0
        if i == str(int(storage_cod_year)): 
            loop_storage_tax_credit = storage_itc
        else : loop_storage_tax_credit = 0
        loop_tax_credit = loop_res_tax_credit + loop_storage_tax_credit
        loop_tax_benefit = -ebit[i]*corporate_income_tax_rate
        tax_credit_list.append(loop_tax_credit)
        tax_benefit_list.append(loop_tax_benefit)
        total_benefit_list.append(loop_tax_credit+loop_tax_benefit)

    tax_benefit_df = pd.DataFrame(np.column_stack([year_array,
                                                 tax_credit_list, 
                                                 tax_benefit_list,
                                                 total_benefit_list]), 
                                 columns = ['year_array',
                                            'tax_credit', 
                                            'tax_benefit',
                                            'total_benefit'])
    tax_benefit = tax_benefit_df.set_index('year_array').T.to_dict('list')
    (tax_credit, tax_benefit, total_benefit) = splitDict(tax_benefit)
    for k, v in tax_credit.items(): tax_credit[k] = float(v)
    for k, v in tax_benefit.items(): tax_benefit[k] = float(v)
    for k, v in total_benefit.items(): total_benefit[k] = float(v)

    itc_taxbenefit_list = []
    operating_taxbenefit_list = []
    cash_allocation_list = []
    temp_total_inflow_outflow_list = []

    for i in timeLine.keys():
        if operating_period[i] <= operating_period[year_array[target_flip_yr]]:
            loop_itc_taxbenefit = tax_credit[i]*te_preflip_tax_allocation
            loop_operating_taxbenefit = tax_benefit[i]*te_preflip_tax_allocation
            loop_cash_allocation = asset_cashflow[i]*te_preflip_cash_allocation
        else:
            loop_itc_taxbenefit = tax_credit[i]*te_postflip_tax_allocation
            loop_operating_taxbenefit = tax_benefit[i]*te_postflip_tax_allocation
            loop_cash_allocation = asset_cashflow[i]*te_postflip_cash_allocation
        itc_taxbenefit_list.append(loop_itc_taxbenefit)
        operating_taxbenefit_list.append(loop_operating_taxbenefit)
        cash_allocation_list.append(loop_cash_allocation)
        temp_total_inflow_outflow_list.append(loop_itc_taxbenefit+loop_operating_taxbenefit+loop_cash_allocation)

    opt_total_inoutflow_list = []
    for i in range(target_flip_yr+1):
        opt_total_inoutflow_list.append(temp_total_inflow_outflow_list[i])

    # Desired IRR
    target_irr = 0.1
    intial_contribution_scaler = min_project_owner_equity # You need to define this value
    flip_yr = target_flip_yr 

    # Define the NPV function with the variable x
    def objective(x):
        cash_flows = []
        temp_zero = -x * intial_contribution_scaler + opt_total_inoutflow_list[0]
        cash_flows.append(temp_zero)
        temp_one = -x * (1 - intial_contribution_scaler) + opt_total_inoutflow_list[1]
        cash_flows.append(temp_one)
        for i in range(2, len(opt_total_inoutflow_list)):
            cash_flows.append(opt_total_inoutflow_list[i])
        npv = sum([cf / (1 + target_irr) ** t for cf, t in zip(cash_flows, range(flip_yr + 1))])
        return npv

    # Define the zero equation constraint
    def zero_equation(x):
        return objective(x)

    #x0 = 0

    # Use the objective function in the eqcons parameter
    sol = fmin_slsqp(objective, x0=0, eqcons=[zero_equation], acc=1.e-40, disp=False)
    tax_equity_class_a = float(sol)

    total_inflow_outflow_list = []
    caplital_contribution_list =[]
    for i in timeLine.keys():
        if operating_period[i] == 0:
            loop_caplital_contribution = -tax_equity_class_a * intial_contribution_scaler
        elif operating_period[i] == 1:
            loop_caplital_contribution = -tax_equity_class_a * (1-intial_contribution_scaler)
        else :  loop_caplital_contribution = 0
        loop_total_inflow_outflow = float(loop_caplital_contribution + temp_total_inflow_outflow_list[int(operating_period[i])])
        caplital_contribution_list.append(float(loop_caplital_contribution))
        total_inflow_outflow_list.append(loop_total_inflow_outflow)

    for_loop_total_inflow_outflow_list_list = [total_inflow_outflow_list[0],total_inflow_outflow_list[1]]
    running_tax_equity_irr_list = [0,0]
    for i in range(2,len(total_inflow_outflow_list)):
        for_loop_total_inflow_outflow_list_list.append(total_inflow_outflow_list[i])
        loop_running_tax_equity_irr = npf.irr(for_loop_total_inflow_outflow_list_list)
        running_tax_equity_irr_list.append(loop_running_tax_equity_irr)

    tax_equity_df = pd.DataFrame(np.column_stack([year_array,
                                                 itc_taxbenefit_list, 
                                                 operating_taxbenefit_list,
                                                 cash_allocation_list,
                                                 caplital_contribution_list,
                                                 total_inflow_outflow_list,
                                                 running_tax_equity_irr_list]), 
                                 columns = ['year_array',
                                            'itc_taxbenefit', 
                                            'operating_taxbenefit',
                                            'cash_allocation',
                                            'caplital_contribution',
                                            'total_inflow_outflow',
                                            'running_tax_equity_irr'])
    tax_equity = tax_equity_df.set_index('year_array').T.to_dict('list')
    (itc_taxbenefit, operating_taxbenefit, cash_allocation, caplital_contribution, total_inflow_outflow, running_tax_equity_irr) = splitDict(tax_equity)
    for k, v in itc_taxbenefit.items(): itc_taxbenefit[k] = float(v)
    for k, v in operating_taxbenefit.items(): operating_taxbenefit[k] = float(v)
    for k, v in cash_allocation.items(): cash_allocation[k] = float(v)
    for k, v in caplital_contribution.items(): caplital_contribution[k] = float(v)
    for k, v in total_inflow_outflow.items(): total_inflow_outflow[k] = float(v)
    for k, v in running_tax_equity_irr.items(): running_tax_equity_irr[k] = float(v)

    # Cash Flow Available for Debt Sizing 
    debt_sizing_flag_list = []
    remaining_cashflow_postte_list = []
    #contracted_term_cafds_list = []
    contract_term_cafds_list = []
    merchent_terms_cafds_list = []
    PTC_terms_cafds_list = []
    total_cafds_list = []

    for i in timeLine.keys():
        if operating_period[i] == 0:
            debt_sizing_flag_list.append(0)
        elif operating_period[i] <= debt_tenor:
            loop_debt_sizing_flag = 1
            debt_sizing_flag_list.append(loop_debt_sizing_flag)
        elif operating_period[i] >= debt_tenor:
            loop_debt_sizing_flag = 0
            debt_sizing_flag_list.append(loop_debt_sizing_flag)
    for i in timeLine.keys():
        if operating_period[i] == 0:
            remaining_cashflow_postte_list.append(0)        
        elif operating_period[i] <= target_flip_yr:
            loop_remaining_cashflow_postte = (1-te_preflip_cash_allocation)*asset_cashflow[i]
            remaining_cashflow_postte_list.append(loop_remaining_cashflow_postte)
        elif operating_period[i] >= target_flip_yr:
            loop_remaining_cashflow_postte = (1-te_postflip_cash_allocation)*asset_cashflow[i]
            remaining_cashflow_postte_list.append(loop_remaining_cashflow_postte)

        loop_contract_term_cafds = debt_sizing_flag_list[int(operating_period[i])]*((remaining_cashflow_postte_list[int(operating_period[i])]*contract_rev_split[i])/contract_debt_coverage_ratio)
        loop_merchent_terms_cafds = debt_sizing_flag_list[int(operating_period[i])]*((remaining_cashflow_postte_list[int(operating_period[i])]*merchant_rev_split[i])/merchant_debt_coverage_ratio)
        loop_ptc_terms_cafds = debt_sizing_flag_list[int(operating_period[i])]*((total_PTC_value[i])/ptc_debt_coverage_ratio)
        contract_term_cafds_list.append(loop_contract_term_cafds)
        merchent_terms_cafds_list.append(loop_merchent_terms_cafds)
        PTC_terms_cafds_list.append(loop_ptc_terms_cafds)
        total_cafds_list.append(loop_contract_term_cafds+loop_merchent_terms_cafds+loop_ptc_terms_cafds)

    cash_flow_for_debt_sculping_df = pd.DataFrame(np.column_stack([year_array,
                                                 debt_sizing_flag_list, 
                                                 remaining_cashflow_postte_list,
                                                 contract_term_cafds_list,
                                                 merchent_terms_cafds_list,
                                                 PTC_terms_cafds_list,
                                                 total_cafds_list]), 
                                 columns = ['year_array',
                                            'debt_sizing_flag', 
                                            'remaining_cashflow_postte',
                                            'contract_term_cafds',
                                            'merchent_terms_cafds',
                                            'PTC_terms_cafds',
                                            'total_cafds'])
    cash_flow_for_debt_sculping = cash_flow_for_debt_sculping_df.set_index('year_array').T.to_dict('list')
    (debt_sizing_flag, remaining_cashflow_postte, contract_term_cafds, merchent_terms_cafds, PTC_terms_cafds, total_cafds) = splitDict(cash_flow_for_debt_sculping)
    for k, v in debt_sizing_flag.items(): debt_sizing_flag[k] = float(v)
    for k, v in contract_term_cafds.items(): contract_term_cafds[k] = float(v)
    for k, v in merchent_terms_cafds.items(): merchent_terms_cafds[k] = float(v)
    for k, v in PTC_terms_cafds.items(): PTC_terms_cafds[k] = float(v)
    for k, v in remaining_cashflow_postte.items(): remaining_cashflow_postte[k] = float(v)
    for k, v in total_cafds.items(): total_cafds[k] = float(v)

    # Debt Sculpting

    c = interest # naming is a hold over from test code could be changed

    x_list = []  # Beginning Balance: naming is a hold over from test code could be change 
    y_list = []  # Interest: naming is a hold over from test code could be change  
    a_list = []  # Amortization: naming is a hold over from test code could be change 
    g_list = []  # Ending Balance: hold over from test code could be change 

    for i in timeLine.keys():
        b = total_cafds[i] # naming is a hold over from test code could be changed
        if operating_period[i] == 0:
            x = 0.0
            y = 0.0
            a = 0.0
            g = 0.0
        elif operating_period[i] == 1:    
            x = b*10**1
            y = b*10**0
            a = b*10**0
            g = b*10**1

            epsilon = 10**-8

            limit = 3000
            i = 0

            while True:
                i = i + 1

                xx = a + g
                yy = c * statistics.mean([x,g])
                aa = b - y
                gg = x - a       

                x = xx
                y = yy
                a = aa
                g = gg

                if i >= limit:
                    break
            x = ((2*y/c)+a)/2
            g = x - a
        elif operating_period[i] <= debt_tenor and operating_period[i] > 1:  
            x = g_list[int(operating_period[i])-1]
            g = x - a       
            y = c * statistics.mean([x,g])
            a = b - y
        elif operating_period[i] == debt_tenor:
            g = 0
        elif operating_period[i] >= debt_tenor:
            x = 0
            y = 0
            a = 0
            g = 0
        x_list.append(x)    
        y_list.append(y)
        a_list.append(a)
        g_list.append(g)

    x_list = []
    g_list = []
    for i in range(len(a_list)):
        x = sum(a_list[i:len(a_list)])
        g = x-a_list[i]
        x_list.append(x)
        g_list.append(g)

    debt = x_list[0]

    debt_sculpting_df = pd.DataFrame(np.column_stack([year_array,
                                                 x_list, 
                                                 y_list,
                                                 a_list,
                                                 g_list]), 
                                 columns = ['year_array',
                                            'beginning_balance', 
                                            'annual_interest_amount',
                                            'amortization',
                                            'ending_balance'])
    debt_sculpting = debt_sculpting_df.set_index('year_array').T.to_dict('list')
    (beginning_balance, annual_interest_amount, amortization, ending_balance) = splitDict(debt_sculpting)
    for k, v in beginning_balance.items(): beginning_balance[k] = float(v)
    for k, v in annual_interest_amount.items(): annual_interest_amount[k] = float(v)
    for k, v in amortization.items(): amortization[k] = float(v)
    for k, v in ending_balance.items(): ending_balance[k] = float(v)

    # Project Owner
    total_capex = storage_capex + res_total_capex
    project_owner = total_capex - (tax_equity_class_a + debt)

    # Operating Taxes
    profit_before_tax_list = []
    interest_payment_list = []
    nol_created_list = []
    taxable_income_prenol_list = []
    taxable_income_postnol_list = []
    beginning_balance_list = []
    nol_used_list = []
    ending_balance_list = []

    for i in timeLine.keys():
        if operating_period[i] == 0:
            loop_beginning_balance = 0
        else: loop_beginning_balance = ending_balance_list[int(operating_period[i]-1)]  
        if operating_period[i] <= target_flip_yr:
            loop_profit_before_tax = ebit[i]*po_preflip_tax_allocation
        else: loop_profit_before_tax = ebit[i]*po_postflip_tax_allocation
        profit_before_tax_list.append(loop_profit_before_tax)
        interest_payment_list.append(-annual_interest_amount[i])
        # Net Operating Loss Balance
        loop_taxable_income_prenol = max((loop_profit_before_tax+-annual_interest_amount[i]),0.0)
        loop_nol_created = -1*min((loop_profit_before_tax+-annual_interest_amount[i]),0.0)
        loop_nol_used = -1*min(loop_beginning_balance,loop_taxable_income_prenol*nol_decution_max) 
        loop_taxable_income_postnol = loop_taxable_income_prenol+loop_nol_used
        beginning_balance_list.append(loop_beginning_balance)
        nol_created_list.append(loop_nol_created)
        nol_used_list.append(loop_nol_used)
        taxable_income_prenol_list.append(loop_taxable_income_prenol)
        taxable_income_postnol_list.append(loop_taxable_income_postnol)
        ending_balance_list.append(loop_beginning_balance+loop_nol_created+loop_nol_used)

    # Owner Tax Credit Usage    
    taxable_payable_list = []
    itc_credit_list = [0]*len(timeLine.keys())
    storage_start_on_res_assest_operation_time_line = 1+(int(storage_cod_year)-int(cod_year))
    itc_credit_list[1] = (itc_amount+storage_itc)-itc_taxbenefit[cod_year]
    if cod_year != str(int(storage_cod_year)):
        itc_credit_list[1] = itc_amount-itc_taxbenefit[cod_year]
        storage_start_on_res_assest_operation_time_line = 1+(int(storage_cod_year)-int(cod_year))
        itc_credit_list[storage_start_on_res_assest_operation_time_line] = storage_itc-itc_taxbenefit[str(int(storage_cod_year))]
    else: itc_credit_list[1] = (itc_amount+storage_itc)-itc_taxbenefit[cod_year]
    tax_credit_balance_list = []
    itc_used_list = []
    tax_credit_ending_balance_list = []

    for i in timeLine.keys():
        if operating_period[i] == 0:
            loop_tax_credit_balance = 0
        else: loop_tax_credit_balance = loop_tax_credit_ending_balance
        loop_itc_used = -min(loop_tax_credit_balance,taxable_income_postnol_list[int(operating_period[i])])
        loop_tax_credit_ending_balance = loop_tax_credit_balance + loop_itc_used + itc_credit_list[int(operating_period[i])]
        tax_credit_balance_list.append(loop_tax_credit_balance)
        itc_used_list.append(loop_itc_used)
        tax_credit_ending_balance_list.append(loop_tax_credit_ending_balance)
        taxable_payable_list.append(taxable_income_postnol_list[int(operating_period[i])]*corporate_income_tax_rate+loop_itc_used)

    operating_taxes_df = pd.DataFrame(np.column_stack([year_array,
                                                 profit_before_tax_list, 
                                                 interest_payment_list]), 
                                 columns = ['year_array',
                                            'profit_before_tax', 
                                            'interest_payment'])
    operating_taxes = operating_taxes_df.set_index('year_array').T.to_dict('list')
    (profit_before_tax, interest_payment) = splitDict(operating_taxes)
    for k, v in profit_before_tax.items(): profit_before_tax[k] = float(v)
    for k, v in interest_payment.items(): interest_payment[k] = float(v)

    net_operating_loss_balance_df = pd.DataFrame(np.column_stack([year_array,
                                                 beginning_balance_list, 
                                                 nol_created_list,
                                                 nol_used_list,
                                                 ending_balance_list]), 
                                 columns = ['year_array',
                                            'beginning_balance', 
                                            'nol_created',
                                            'nol_used',
                                            'ending_balance'])
    net_operating_loss_balance = net_operating_loss_balance_df.set_index('year_array').T.to_dict('list')
    (beginning_balance, nol_created, nol_used, ending_balance) = splitDict(net_operating_loss_balance)
    for k, v in beginning_balance.items(): beginning_balance[k] = float(v)
    for k, v in nol_created.items(): nol_created[k] = float(v)
    for k, v in nol_used.items(): nol_used[k] = float(v)
    for k, v in ending_balance.items(): ending_balance[k] = float(v)

    owner_tax_credit_usage_df = pd.DataFrame(np.column_stack([year_array,
                                                 itc_credit_list, 
                                                 tax_credit_balance_list,
                                                 itc_used_list,
                                                 tax_credit_ending_balance_list]), 
                                 columns = ['year_array',
                                            'itc_credit', 
                                            'tax_credit_balance',
                                            'itc_used',
                                            'tax_credit_ending_balance'])
    owner_tax_credit_usage = owner_tax_credit_usage_df.set_index('year_array').T.to_dict('list')
    (itc_credit, tax_credit_balance, itc_used, tax_credit_ending_balance) = splitDict(owner_tax_credit_usage)
    for k, v in itc_credit.items(): itc_credit[k] = float(v)
    for k, v in tax_credit_balance.items(): tax_credit_balance[k] = float(v)
    for k, v in itc_used.items(): itc_used[k] = float(v)
    for k, v in tax_credit_ending_balance.items(): tax_credit_ending_balance[k] = float(v)


    # Capital Stack
    capital_stack_debt = debt
    capital_stack_debt_per = capital_stack_debt/total_capex

    taxable_df = pd.DataFrame(np.column_stack([year_array,
                                                 taxable_income_prenol_list, 
                                                 taxable_income_postnol_list,
                                                 taxable_payable_list]), 
                                 columns = ['year_array',
                                            'taxable_income_prenol', 
                                            'taxable_income_postnol',
                                            'taxable_payable'])
    taxable = taxable_df.set_index('year_array').T.to_dict('list')
    (taxable_income_prenol, taxable_income_postnol, taxable_payable) = splitDict(taxable)
    for k, v in taxable_income_prenol.items(): taxable_income_prenol[k] = float(v)
    for k, v in taxable_income_postnol.items(): taxable_income_postnol[k] = float(v)
    for k, v in taxable_payable.items(): taxable_payable[k] = float(v)

    capital_stack_tax_equity = tax_equity_class_a
    capital_stack_tax_equity_per = capital_stack_tax_equity/total_capex

    capital_stack_project_owner = total_capex - (debt + tax_equity_class_a)  
    capital_stack_project_owner_per = capital_stack_project_owner/total_capex

    project_owner_costofcapital = mrit_inputs['Project Owner Cost of Capital'].item() # check with Joss on this
    project_wacc = capital_stack_debt_per*interest + capital_stack_tax_equity_per*project_owner_costofcapital + capital_stack_project_owner_per*target_yield       

    # Project Owner Returns
    upfront_capital_contribution_list = [0]*len(timeLine.keys())
    upfront_capital_contribution_list[0] = -1*capital_stack_project_owner

    upfront_capital = {}
    for i in range(len(year_array)):
        upfront_capital[year_array[i]] = float(upfront_capital_contribution_list[i])

    total_cash_flow = {}
    project_cash_flow = {}
    for i in timeLine.keys():
        total_cash_flow[i] = float(upfront_capital[i]+total_PTC_value[i]+asset_cashflow[i]-cash_allocation[i]-
                                   (annual_interest_amount[i]+amortization[i])-taxable_payable[i])
        project_cash_flow[i] = float(upfront_capital[i]+asset_cashflow[i])

    total_cash_flow_list = list(total_cash_flow.values())
    project_cash_flow = list(project_cash_flow.values())
    project_owner_irr = round(npf.irr(total_cash_flow_list),4)
    project_owner_npv = round(npf.npv(project_owner_costofcapital, total_cash_flow_list),4)       
    project_npv = round(npf.npv(project_wacc,project_cash_flow),4)   
    #debt_per = debt/total_capex

    if capital_stack_project_owner_per < min_project_owner_equity:
        difference = total_capex*min_project_owner_equity - capital_stack_project_owner
        capital_stack_project_owner = capital_stack_project_owner + difference
        capital_stack_debt = capital_stack_debt - difference
        capital_stack_project_owner_per = capital_stack_project_owner/total_capex
        capital_stack_debt_per = capital_stack_debt/total_capex

    if capital_stack_debt < 0:
        difference = 0 - capital_stack_debt
        capital_stack_project_owner = capital_stack_project_owner - difference
        capital_stack_debt = total_capex-(capital_stack_tax_equity+capital_stack_project_owner)
        capital_stack_project_owner_per = capital_stack_project_owner/total_capex
        capital_stack_debt_per = capital_stack_debt/total_capex    

    npv_om = npf.npv(project_wacc,list(total_operating_expenditures.values()))
    npv_ca = npf.npv(project_wacc,list(cash_allocation.values()))
    npv_ia = npf.npv(project_wacc,list(amortization.values())+list(annual_interest_amount.values()))
    total_cost = total_capex + npv_om #+ npv_ca + npv_ia
    total_cost_taxeq = total_capex + npv_om - capital_stack_tax_equity #+ npv_ca + npv_ia

    mwh_lifetime_list = []
    for i in range(operating_lifetime):
        mwh_lifetime_list.append(mwh_p50*(1-degradation_per*i))
        mwh_lifetime = sum(mwh_lifetime_list)

    lcoe = total_cost/mwh_lifetime
    lcoe_te = total_cost_taxeq/mwh_lifetime
    lcoe-lcoe_te

    project_summary = np.array([project_owner_irr, project_owner_npv, project_npv, lcoe, lcoe_te, lcoe-lcoe_te, capital_stack_tax_equity, capital_stack_project_owner, 
                                capital_stack_debt, capital_stack_tax_equity_per, capital_stack_project_owner_per, capital_stack_debt_per, 
                                project_wacc])
    cash_flow_table_df = pd.DataFrame({'Upfront Capital': pd.Series(upfront_capital), 
                                   'Cashflow': (pd.Series(total_PTC_value)+pd.Series(asset_cashflow)),
                                   'Cash Allocation': -1*pd.Series(cash_allocation), 
                                   'Debt': -1*(pd.Series(annual_interest_amount)+pd.Series(amortization)), 
                                   'Tax Payable': -1*pd.Series(taxable_payable), 
                                   'Total Cash Flow': pd.Series(total_cash_flow), 
                                  }).T
    
    mrit_outputs = [project_summary, cash_flow_table_df]
    return mrit_outputs

'''
MRIT Unlevered IRR function
'''

def unlevered_project_summary(mrit_inputs):   
    mrit_inputs['Debt Tenor'] = 0
    return levered_project_summary(mrit_inputs)

'''
MRIT sensitivty analysis function
''' 
def mrit_value_sensitivty(xls_input,project_input,parameter_input,value_min_input,value_max_input,step_input):    
    xls = xls_input
    sheet_names = xls.sheet_names

    dfs = {}
    for sheet_name in sheet_names:
        dfs[sheet_name] = pd.read_excel(xls, sheet_name=sheet_name)

    # find general inputs to be used in the excel
    df_dcfinputs_general = dfs['General Inputs']
    df_dcfinputs_general = df_dcfinputs_general.replace('',np.nan) 
    df_dcfinputs_general = df_dcfinputs_general.dropna(axis = "columns", how = "any")
    df_dcfinputs_general = df_dcfinputs_general.drop(columns=['Unit'])

    # clean input excel to be looped/entered in the IRR function
    scenarios = list(df_dcfinputs_general.keys())
    scenarios.remove(scenarios[0])

    # find inputed advanced scenarios
    df_dcfinputs_advanced = dfs['Advanced Inputs (User)']
    df_dcfinputs_advanced = df_dcfinputs_advanced.replace('',np.nan) 
    df_dcfinputs_advanced = df_dcfinputs_advanced.dropna(axis = "columns", how = "any")
    df_dcfinputs_advanced = df_dcfinputs_advanced.drop(columns=['Unit'])
    advanced_scenarios = list(df_dcfinputs_advanced.keys())
    advanced_scenarios.remove(advanced_scenarios[0])
    # sensitivity analysis
    # IRR output
    levered_scenario_set = []
    levered_sum_set = []
    unlevered_scenario_set = []
    project_name = project_input # pick the project from the input file you want to parameterize 
    if (project_name in advanced_scenarios) == False:
        df_dcfinputs_advanced_default = dfs['Advanced Inputs (Default)']
        df_dcfinputs_advanced_default = df_dcfinputs_advanced_default.replace('',np.nan) 
        df_dcfinputs_advanced_default = df_dcfinputs_advanced_default.dropna(axis = "columns", how = "any")
        loop_df_dcfinputs_advanced = df_dcfinputs_advanced_default['Default']
    else: loop_df_dcfinputs_advanced = df_dcfinputs_advanced[project_name]
    df_mrit_inputs_general = pd.concat((df_dcfinputs_general['Input'],df_dcfinputs_general[project_name]),axis = 1)
    df_mrit_inputs_advanced = pd.concat((df_dcfinputs_advanced['Input'],loop_df_dcfinputs_advanced),axis = 1)
    df_mrit_inputs_advanced.columns = ['Input', project_name]
    mrit_inputs = pd.concat((df_mrit_inputs_general, df_mrit_inputs_advanced), ignore_index=True).T
    new_header = mrit_inputs.iloc[0] #grab the first row for the header
    new_header
    mrit_inputs = mrit_inputs[1:] #take the data less the header row
    mrit_inputs.columns = new_header #set the header row as the df header

    # value range of sensitivity analysis
    start = value_min_input
    end = value_max_input
    step = step_input

    #sensitivity_analysis_range = list(range(start, end+step, step))
    sensitivity_analysis_range = list(np.linspace(value_min, value_max, num=15))
    sensitivity_parameter = parameter_input

    project_name_col = pd.DataFrame({'Project Name' : [project_name]*len(sensitivity_analysis_range)})
    sen_name_col = pd.DataFrame({'Parameter' : [sensitivity_parameter]*len(sensitivity_analysis_range)})
    sen_value_col = pd.DataFrame({'Parameter Value' : sensitivity_analysis_range})
    levered_name_col = pd.DataFrame({'Levered' : ['Levered']*len(sensitivity_analysis_range)})
    unlevered_name_col = pd.DataFrame({'Unlevered' : ['Unlevered']*len(sensitivity_analysis_range)})
    

    levered_sensitivity_analysis_summary_set = []
    unlevered_sensitivity_analysis_summary_set = []
    sensitivity_mrit_inputs = mrit_inputs.copy()
    for i in sensitivity_analysis_range:
        sensitivity_mrit_inputs[sensitivity_parameter] = i
        sensitivity_levered_mrit_inputs = sensitivity_mrit_inputs.copy()
        sensitivity_unlevered_mrit_inputs = sensitivity_mrit_inputs.copy()

        levered_sensitivity_analysis_summary_set.append(levered_project_summary(sensitivity_levered_mrit_inputs)[0])
        unlevered_sensitivity_analysis_summary_set.append(unlevered_project_summary(sensitivity_unlevered_mrit_inputs)[0])

    levered_output_summary_table = pd.DataFrame(levered_sensitivity_analysis_summary_set)
    levered_sen_summary_table = pd.concat([levered_name_col,project_name_col,sen_name_col,sen_value_col,levered_output_summary_table], axis = 1)
    levered_sen_summary_table.columns = ['IRR Type','Run Name','Parameter','Parameter Value','IRR', 'Equity Owner NPV', 'Project NPV', 'LCOE', 'LCOE with Tax Equity','LCOE Difference', 'Tax Equity', 'Project Owner',
                                         'Debt', 'Tax Equity Precent', 'Project Owner Precent', 'Debt Precent', 'Project WACC']

    unlevered_output_summary_table = pd.DataFrame(unlevered_sensitivity_analysis_summary_set)
    unlevered_sen_summary_table = pd.concat([unlevered_name_col,project_name_col,sen_name_col,sen_value_col,unlevered_output_summary_table], axis = 1)
    unlevered_sen_summary_table.columns = ['IRR Type','Run Name','Parameter','Parameter Value','IRR', 'Equity Owner NPV', 'Project NPV', 'LCOE', 'LCOE with Tax Equity','LCOE Difference', 'Tax Equity', 'Project Owner',
                                           'Debt', 'Tax Equity Precent', 'Project Owner Precent', 'Debt Precent', 'Project WACC']

    #output_reduced_unlevered_sen_summary_table = unlevered_sen_summary_table.drop(columns=['Debt', 'Debt Precent'])
    
    sen_summary_table = pd.concat([levered_sen_summary_table, unlevered_sen_summary_table], ignore_index=True, sort=False)
    return sen_summary_table

#############################################################################################################################################################################################################################
#############################################################################################################################################################################################################################
#############################################################################################################################################################################################################################
#############################################################################################################################################################################################################################
#############################################################################################################################################################################################################################

'''
MRIT luanch code
'''

# read input excel

#xls = pd.ExcelFile('')
sheet_names = xls.sheet_names

dfs = {}
for sheet_name in sheet_names:
    dfs[sheet_name] = pd.read_excel(xls, sheet_name=sheet_name)

# find general inputs to be used in the excel
df_dcfinputs_general = dfs['General Inputs']
df_dcfinputs_general = df_dcfinputs_general.replace('',np.nan) 
df_dcfinputs_general = df_dcfinputs_general.dropna(axis = "columns", how = "any")
df_dcfinputs_general = df_dcfinputs_general.drop(columns=['Unit'])

# clean input excel to be looped/entered in the IRR function
scenarios = list(df_dcfinputs_general.keys())
scenarios.remove(scenarios[0])

# find inputed advanced scenarios
df_dcfinputs_advanced = dfs['Advanced Inputs (User)']
df_dcfinputs_advanced = df_dcfinputs_advanced.replace('',np.nan) 
df_dcfinputs_advanced = df_dcfinputs_advanced.dropna(axis = "columns", how = "any")
df_dcfinputs_advanced = df_dcfinputs_advanced.drop(columns=['Unit'])
advanced_scenarios = list(df_dcfinputs_advanced.keys())
advanced_scenarios.remove(advanced_scenarios[0])

# IRR output
levered_scenario_set = []
levered_sum_set = []
unlevered_scenario_set = []
unlevered_sum_set = []
for i in scenarios:
    # decide if you are using inputed advanced scenarios or default advanced settings
    if (i in advanced_scenarios) == False:
        df_dcfinputs_advanced_default = dfs['Advanced Inputs (Default)']
        df_dcfinputs_advanced_default = df_dcfinputs_advanced_default.replace('',np.nan) 
        df_dcfinputs_advanced_default = df_dcfinputs_advanced_default.dropna(axis = "columns", how = "any")
        loop_df_dcfinputs_advanced = df_dcfinputs_advanced_default['Default']
    else: loop_df_dcfinputs_advanced = df_dcfinputs_advanced[i]
    df_mrit_inputs_general = pd.concat((df_dcfinputs_general['Input'],df_dcfinputs_general[i]),axis = 1)
    df_mrit_inputs_advanced = pd.concat((df_dcfinputs_advanced['Input'],loop_df_dcfinputs_advanced),axis = 1)
    df_mrit_inputs_advanced.columns = ['Input', i]
    levered_scenario_set.append(i)
    unlevered_scenario_set.append(i)
    mrit_inputs = pd.concat((df_mrit_inputs_general, df_mrit_inputs_advanced), ignore_index=True).T
    new_header = mrit_inputs.iloc[0] #grab the first row for the header
    new_header
    mrit_inputs = mrit_inputs[1:] #take the data less the header row
    mrit_inputs.columns = new_header #set the header row as the df header
    levered_mrit_inputs = mrit_inputs.copy()
    unlevered_mrit_inputs = mrit_inputs.copy()
    levered_sum_set.append(levered_project_summary(levered_mrit_inputs)[0])
    unlevered_sum_set.append(unlevered_project_summary(unlevered_mrit_inputs)[0])

# output summary    
levered_summary_table = pd.DataFrame(np.column_stack([levered_scenario_set,levered_sum_set]))
levered_summary_table.columns = ['Run Name', 'IRR', 'Equity Owner NPV', 'Project NPV', 'LCOE', 'LCOE with Tax Credit','LCOE Difference', 'Tax Credit', 
                         'Project Owner', 'Debt', 'Tax Credit Precent', 'Project Owner Precent', 'Debt Precent', 'Project WACC']

unlevered_summary_table = pd.DataFrame(np.column_stack([unlevered_scenario_set,unlevered_sum_set]))
unlevered_summary_table.columns = ['Run Name', 'IRR', 'Equity Owner NPV', 'Project NPV', 'LCOE', 'LCOE with Tax Credit','LCOE Difference', 'Tax Credit', 
                         'Project Owner', 'Debt', 'Tax Credit Precent', 'Project Owner Precent', 'Debt Precent', 'Project WACC']
output_reduced_unlevered_summary_table = unlevered_summary_table.drop(columns=['Debt', 'Debt Precent'])

# ppa sensitivity 
df_dcfinputs_senana = dfs['Sensitivity Analysis Inputs']
sensitivity_analysis_summary_table = pd.DataFrame()

for index in df_dcfinputs_senana.index:
    loop_df_dcfinputs_senana = df_dcfinputs_senana.iloc[[index]]
    project = loop_df_dcfinputs_senana['Project Name'].item()
    parameter = loop_df_dcfinputs_senana['Parameter'].item()
    value_min = loop_df_dcfinputs_senana['Min Value'].item()
    value_max = loop_df_dcfinputs_senana['Max Value'].item()
    step = loop_df_dcfinputs_senana['Step'].item()
    sensitivity_analysis_summary_table = pd.concat([sensitivity_analysis_summary_table,pd.DataFrame(mrit_value_sensitivty(xls,project,parameter,value_min,value_max,step))], ignore_index=True, sort=False)

# output excel
# create a excel writer object
writer =  pd.ExcelWriter('MRIT_output_table.xlsx') # needto fix the path
# use to_excel function and specify the sheet_name and index 
# to store the dataframe in specified sheet
blank_sheet = pd.DataFrame()
blank_sheet.to_excel(writer, sheet_name="MRIT Outputs >>", index=False)
levered_summary_table.to_excel(writer, sheet_name="Levered", index=False)
output_reduced_unlevered_summary_table.to_excel(writer, sheet_name="Unlevered", index=False)
blank_sheet.to_excel(writer, sheet_name="MRIT Sensitivty Analysis >>", index=False)
sensitivity_analysis_summary_table.to_excel(writer, sheet_name='Sensitivty Analysis', index=False)
writer.close()
