import streamlit as st
import pdfplumber
import re
import pandas as pd
import os
import numpy as np
import xlsxwriter
from openpyxl.styles import PatternFill
from openpyxl.worksheet.dimensions import SheetFormatProperties
from datetime import datetime , timedelta
import datetime as dt 
import matplotlib.pyplot as plt
import time


def calculate_seasonal_consumption(df):
    seasons = {
        'Winter': [12, 1, 2],  # Winter months: Dec, Jan, Feb
        'Spring': [3, 4, 5],   # Spring months: Mar, Apr, May
        'Summer': [6, 7, 8],   # Summer months: Jun, Jul, Aug
        'Fall': [9, 10, 11]    # Fall months: Sep, Oct, Nov
    }
    
    seasonal_totals = {}
    
    for season, months in seasons.items():
        # Check if the required months are present in the DataFrame
        available_months = [month for month in months if month in df.columns]
        if available_months:
            seasonal_totals[season] = df.loc["Total Comsuption KWH", available_months].sum()
        else:
            print(f"No available months found for {season}. Skipping calculation for this season.")
    
    return seasonal_totals

# Function to extract data from a PDF file
def extract_data(pdf_file):
        # Initialize a dictionary with 0 values for all keywords
    extracted_data = {keyword: 0 for keyword in keywords}

    with pdfplumber.open(pdf_file) as pdf:
        for page_num in range(len(pdf.pages)):
            page = pdf.pages[page_num]
            text = page.extract_text()
            text = text.replace("Euel", "Fuel")

            # Replace "FEb" with "EEb"
            text = text.replace("Eeb", "Feb")
            #text = text.replace("Regulatoiy fee", "Regulatory fee")
            

            # Initialize flags for Rate and Constant
            rate_flag = False
            constant_flag = False
            service_to_flag = False

            for line in text.split('\n'):
                # Scenario 4: Find exact word "Rate"
                if "Rate:" in line:
                    rate_flag = True
                    # Use regex to extract 2 words after "Rate"
                    rate_match = re.search(r'Rate:\s*(\S+\s+\S+)', line)
                    if rate_match:
                        extracted_data["Rate"] = rate_match.group(1)

                # Scenario 5: Find exact word "Demand KW" and save the second number after that as "Const"
                if "Demand KW" in line:
                    demand_kw_match = re.findall(r'[\d,]+(\.\d+)?', line)
                    if len(demand_kw_match) >= 2:
                        if demand_kw_match and demand_kw_match[0]:
                            extracted_data["Demand KW"] = float(demand_kw_match[0].replace(',', ''))
                        else:
                            # Handle the case where demand_kw_match is empty or None
                            #print("demand_kw_match is empty or None")
                            extracted_data["Demand KW"] = None  # or some default value
                        

                        # Find the second number after "Demand KW"
                        second_number_match = re.search(r'\d+(\.\d+)?\s+(\d+(\.\d+)?)', line)
                        if second_number_match:
                            extracted_data["Const"] = float(second_number_match.group(2).replace(',', ''))
                            constant_flag = True
                        
                            # Find the third number after "Demand KW" using various patterns
                        third_number_match = re.search(r'Demand KW.*?(\d+(\.\d+)?)\s+(\d+(\.\d+)?)\s+(\d+(\.\d+)?)', line)
                        if not third_number_match:
                            third_number_match = re.search(r'Demand KW.*?(\d+(\.\d+)?)\s+(\d+(\.\d+)?)', line)
                        if not third_number_match:
                            third_number_match = re.search(r'Demand KW.*?(\d+(\.\d+)?)', line)

                        if third_number_match:
                            # Check if group(5) exists in the match object before accessing it
                            if len(third_number_match.groups()) >= 5:
                                extracted_data["Usage"] = float(third_number_match.group(5).replace(',', ''))
                            elif len(third_number_match.groups()) >= 3:
                                extracted_data["Usage"] = float(third_number_match.group(3).replace(',', ''))
                            else:
                                # Handle the case where group(3) doesn't exist (set a default value or handle it as needed)
                                extracted_data["Usage"] = 0  # You can change this default value if needed
                
                regulatory_fee_match = re.search(r'Regulatory fee (State fee)', text, re.IGNORECASE)
                if regulatory_fee_match:
                                # If found, extract the number
                                extracted_data["Regulatory fee (State fee)"] = float(regulatory_fee_match.group(1).replace(',', ''))
                                #extracted_data["Regulatory fee (State fee)"] = float(regulatory_fee_match.group(1))
                                #print("mfjkdndnsjj")
                # Scenario 5: Find exact word "Demand KW" and save the second number after that as "Const"
                if "On-peak demand" in line:
                    demand_kw_match = re.findall(r'[\d,]+(\.\d+)?', line)
                    
                    if len(demand_kw_match) >= 2:
                        extracted_data["On-peak demand"] = float(demand_kw_match[0].replace(',', ''))

                        # Find the second number after "Demand KW"
                        second_number_match = re.search(r'\d+(\.\d+)?\s+(\d+(\.\d+)?)', line)
                        if second_number_match:
                            extracted_data["Const"] = float(second_number_match.group(2).replace(',', ''))
                            constant_flag = True

                            # Find the third number after "Demand KW" using various patterns
                        third_number_match = re.search(r'On-peak demand.*?(\d+(\.\d+)?)\s+(\d+(\.\d+)?)\s+(\d+(\.\d+)?)', line)
                        if not third_number_match:
                            third_number_match = re.search(r'On-peak demand.*?(\d+(\.\d+)?)\s+(\d+(\.\d+)?)', line)
                        if not third_number_match:
                            third_number_match = re.search(r'On-peak demand.*?(\d+(\.\d+)?)', line)

                        if third_number_match:
                            # Check if group(5) exists in the match object before accessing it
                            if len(third_number_match.groups()) >= 5:
                                extracted_data["On-peak demand2"] = float(third_number_match.group(5).replace(',', ''))
                            elif len(third_number_match.groups()) >= 3:
                                extracted_data["On-peak demand2"] = float(third_number_match.group(3).replace(',', ''))
                            else:
                                # Handle the case where group(3) doesn't exist (set a default value or handle it as needed)
                                extracted_data["On-peak demand2"] = 0  # You can change this default value if needed
                
                                
                # Scenario 6: Find exact word "Service to" and save 3 words after that
                if "Service to" in line:
                    service_to_match = re.search(r'Service to\s+(\S+\s+\d{1,2},\s+\d{4})', line)
                    if service_to_match:
                        extracted_data["Service to"] = service_to_match.group(1)
                        service_to_date = extracted_data["Service to"]
                        # Parse the date string into a datetime object
                        service_date_obj = dt.datetime.strptime(service_to_date, '%b %d, %Y')
                        # Extract the month as an integer (1-12)
                        extracted_data["Service Month"] = service_date_obj.month
                        service_to_flag = True

                # Check for other keywords
                for keyword in keywords:
                    if keyword in line:
                        # Scenario 1: Find the first number next to the exact word
                        number_match = re.search(rf'{keyword}\s*([\d,]+(\.\d+)?)', line)
                        if number_match:
                            extracted_data[keyword] = float(number_match.group(1).replace(',', ''))

                        # Scenario 2: Find $ sign next to the exact word
                        dollar_match = re.search(rf'{keyword}\s*\$\s*([\d,]+(\.\d+)?)', line)
                        if dollar_match:
                            extracted_data[keyword] = float(dollar_match.group(1).replace(',', ''))

                        # Scenario 3: Find $ sign between parentheses next to the exact word
                        parentheses_match = re.search(rf'{keyword}\s*\((.*?)\)', line)
                        if parentheses_match:
                          content_inside_parentheses = parentheses_match.group(1)
                          dollar_inside_parentheses_match = re.search(r'\$\s*([\d,]+(\.\d+)?)', content_inside_parentheses)
                          if dollar_inside_parentheses_match:
                              extracted_data[keyword] = float(dollar_inside_parentheses_match.group(1).replace(',', ''))
                          else:
                              # No $ sign between parentheses, extract the number after parentheses
                              number_after_parentheses_match = re.search(r'\)\s*([\d,]+(\.\d+)?)', line)
                              if number_after_parentheses_match:
                                  extracted_data[keyword] = float(number_after_parentheses_match.group(1).replace(',', ''))
     # Inside the extract_data function, after checking for "Non-fuel energy charge:"
    found_non_fuel_energy_charge = False  # Initialize the flag
    found_on_peak = False  # Initialize the flag to track "On-peak"
    for line in text.split('\n'):
        if "Non-fuel energy charge:" in line:
            found_non_fuel_energy_charge = True
            continue  # Move to the next line to search for "On-peak"

        # Check for "On-peak" in the line, but only if we've found "Non-fuel energy charge:"
        if found_non_fuel_energy_charge and "On-peak" in line:
            on_peak_match = re.search(r'\$([\d,]+(\.\d+)?)', line)
            
            if on_peak_match:
                # Save the matched number under the key "Non-fuel energy charge: on-peak"
                extracted_data["Non-fuel energy charge: on-peak"] = float(on_peak_match.group(1).replace(',', ''))
                found_on_peak = True  # Set the flag to indicate "On-peak" was found
            continue  # Move to the next line to search for "Off-peak"

        # Check for "Off-peak" in the line, but only if "On-peak" was found in the previous line
        if found_on_peak and "Off-peak" in line:
            off_peak_match = re.search(r'\$([\d,]+(\.\d+)?)', line)
            if off_peak_match:
                # Save the matched number under the key "Non-fuel energy charge: off-peak"
                extracted_data["Non-fuel energy charge: off-peak"] = float(off_peak_match.group(1).replace(',', ''))
                found_non_fuel_energy_charge = False  # Reset the flag
                found_on_peak = False  # Reset the "On-peak" flag

    
    
    
        
        Demand_charge = False
        found_demand_charge = False
        found_on_peak1 = False
        for line in text.split('\n'):
            if "Demand charge:" in line:
                found_demand_charge = True
                # Try to extract the number from the line
                demand_charge_match = re.search(r'\$([\d,]+(\.\d+)?)', line)
                if demand_charge_match:
                    # Extract the value of "Demand charge"
                    extracted_data["Demand charge:"] = float(demand_charge_match.group(1).replace(',', ''))
                    
                    break  # Exit the loop after finding and extracting "Demand charge" value
                continue  # Move to the next line to search for "On-Peak" after "Demand charge" is found

            if found_demand_charge and "On-Peak" in line:
                on_peak_match = re.search(r'\$([\d,]+(\.\d+)?)', line)
                if on_peak_match:
                    extracted_data["Demand charge-On-peak"] = float(on_peak_match.group(1).replace(',', ''))
                    found_on_peak1 = True
                    break  # Exit the loop after finding "On-peak" since you only want to extract it once
        
        
        Demand_charge_On_peak = extracted_data.get("Demand charge-On-peak", 0)
        #print(Demand_charge_On_peak)
        
        for line in text.split('\n'):
            non_fuel_charge_match = re.search(r'Non-fuel energy charge:\s*\n\s*\$([\d,]+(\.\d+)?)', line)

            # Check if a match is found
            if non_fuel_charge_match:
                # Extract the matched number and convert it to a float
                non_fuel_charge_value = float(non_fuel_charge_match.group(1).replace(',', ''))
                extracted_data["Non-fuel energy charge:"] = non_fuel_charge_value

        lines = text.split('\n')
        for i in range(len(lines)):
            # Check if the line contains the exact phrase "Non-fuel energy charge:"
            if "Non-fuel energy charge:" in lines[i]:
                # Extract the number from the next line
                next_line = lines[i+1].strip()
                non_fuel_charge_match = re.search(r'\$([\d,]+(\.\d+)?)', next_line)

                # Check if a match is found
                if non_fuel_charge_match:
                    # Extract the matched number and convert it to a float
                    non_fuel_charge_value = float(non_fuel_charge_match.group(1).replace(',', ''))
                    extracted_data["Non-fuel energy charge:"] = non_fuel_charge_value
                    
    # Inside the extract_data function, after checking for "Fuel charge:"
        found_Fuel_charge= False  # Initialize the flag
        found_on_peak = False  # Initialize the flag to track "On-peak"
        for line in text.split('\n'):
            fuel_charge_match = re.search(r'Fuel charge: \$([\d,]+(\.\d+)?)', text)

            if fuel_charge_match:
        # Extract the matched number and convert it to a float
                fuel_charge_value = float(fuel_charge_match.group(1).replace(',', ''))
                
                # Assign the extracted value to the dictionary
                extracted_data["Fuel charge:"] = fuel_charge_value
            if "Fuel charge:" in line:
                found_Fuel_charge = True
                continue  # Move to the next line to search for "On-peak"

            # Check for "On-peak" in the line, but only if we've found "Non-fuel energy charge:"
            if found_Fuel_charge and "On-peak" in line:
                
                on_peak_match = re.search(r'\$([\d,]+(\.\d+)?)', line)
                if on_peak_match:
                    # Save the matched number under the key "Non-fuel energy charge: on-peak"
                    extracted_data["Fuel charge-On-peak"] = float(on_peak_match.group(1).replace(',', ''))
                    found_on_peak = True  # Set the flag to indicate "On-peak" was found
                continue  # Move to the next line to search for "Off-peak"
            
            
           

            


            # Check for "Off-peak" in the line, but only if "On-peak" was found in the previous line
            if found_on_peak and "Off-peak" in line:
                off_peak_match = re.search(r'\s*([+-]?\d+(\.\d+)?)', line)
                if off_peak_match:
                    # Save the matched number under the key "Non-fuel energy charge: off-peak"
                    extracted_data["Fuel charge-Off-peak"] = float(off_peak_match.group(1).replace(',', ''))
                    found_Fuel_charge = False  # Reset the flag
                    found_on_peak = False  # Reset the "On-peak" flag

            

            pattern = r'FPL SolarTogether credit\s*([−\-\d,.]+)'

            




            # Search for the pattern in the text
            match = re.search(pattern, text)
            #print(match)
            if match:
                # If a match is found, extract the number including its sign
                #print("llllllllllllllllllllllllllllllllllllllllllllllllllll")
                # Replacing non-standard minus sign with standard minus sign
                FPL_SolarTogether_credit1 = (match.group(1))
                standard_string = FPL_SolarTogether_credit1.replace('−', '-')

                # Removing comma
                FPL_SolarTogether_credit = FPL_SolarTogether_credit1.replace(',', '')
                numeric_part = re.sub(r'[^\d.-]', '', FPL_SolarTogether_credit)

                # Converting to float
                FPL_SolarTogether_credit = float(numeric_part)
                # If the original string had a negative sign, multiply the float value by -1
                if '−' in FPL_SolarTogether_credit1:
                    FPL_SolarTogether_credit *= -1
                extracted_data["FPL SolarTogether credit"] = FPL_SolarTogether_credit
                #print(FPL_SolarTogether_credit)
          # Inside the extract_data function, after extracting all other values
    rate = str(extracted_data.get("Rate", "")).strip().upper()
    valid_rates = ["GSLDT-1 GENERAL", "GSDT-1 GENERAL", "GSLD-1 GENERAL", "HLFT-2 HIGH","HLFT-2 HIGH LOAD FACTOR DEMAND TIME OF USE","HLFT-1 HIGH","OL-1 OUTDOOR","SDTR-2B SEASONAL"]

    print(rate)
    if rate in valid_rates:
        
    #if "GSLDT-1 GENERAL" in rate or  "GSLDT-1 GENERAL SVC LRG DEMAND TIME OF USE" in rate or "SDTR-2B SEASONAL DEMAND OPT B TIME OF USE" in rate or "GSDT-1 GENERAL" in rate or "GSLDT-1 GENERAL" in rate or "GSLD-1 GENERAL" in rate or "HLFT-2 HIGH LOAD FACTOR DEMAND TIME OF USE":
        print("rate hlfgfgf")
        non_fuel_off_peak = extracted_data.get("Non-fuel energy charge: off-peak", 0)
        fuel_off_peak = extracted_data.get("Fuel charge-Off-peak", 0)
        off_peak_kwh_used = extracted_data.get("Off-peak kWh used", 0)
        non_fuel_on_peak = extracted_data.get("Non-fuel energy charge: on-peak", 0)
        fuel_on_peak = extracted_data.get("Fuel charge-On-peak", 0)
        #on_peak_kwh_used = extracted_data.get("On-peak kWh used", 0)
        demand_charge = extracted_data.get("Demand charge:", 0)
        Demand_charge_On_peak = extracted_data.get("Demand charge-On-peak", 0)
        
        if demand_charge==0:
           demand_charge=Demand_charge_On_peak
           
        on_peak_demand2 = extracted_data.get("On-peak demand2", 0)
        
        
        on_peak_demand1 = extracted_data.get("On-peak demand", 0)
        
        Power_monitoring_premium_plus = extracted_data.get("Power monitoring-premium plus", 0)
        maximum_demand = extracted_data.get("Maximum demand", 0)
        maximum = extracted_data.get("Maximum", 0)
        franchise_charge = extracted_data.get("Franchise charge", 0)
        utility_tax =extracted_data.get("Utility tax", 0)
        florida_sales_tax =extracted_data.get("Florida sales tax", 0)
        gross_receipts_tax =extracted_data.get("Gross rec. tax/Regulatory fee", 0)
        gross_receipts_tax1 =extracted_data.get("Gross receipts tax", 0)
        county_sales_tax =extracted_data.get("County sales tax", 0)
        base_charge =extracted_data.get("Base charge:", 0)
        Reg_fee =extracted_data.get("Regulatory fee", 0)
        Discretionary_sales =extracted_data.get("Discretionary sales surtax", 0)
        Service_Charge =extracted_data.get("Service Charge", 0)
        franchise_fee = extracted_data.get("Franchise fee", 0)
        total_comsuption_kwh=extracted_data.get("kWh Used", 0)
        Late_payment_charge =extracted_data.get( "Late payment charge", 0)
        FPL_SolarTogether_charge =extracted_data.get( "FPL SolarTogether charge", 0)
        FPL_SolarTogether_credit =extracted_data.get( "FPL SolarTogether credit", 0)
        #print(FPL_SolarTogether_credit)
        #print("inam solar inam solar")
        Total_Comsuption_kwh= total_comsuption_kwh
        if 'Total_Comsuption_kwh' in locals() and Total_Comsuption_kwh is not None and Total_Comsuption_kwh != "":
            print("jjjjjjjjjjjjjjjjjjj")
            if off_peak_kwh_used != 0 and 'Off-peak kWh used' is not None:
               
               

   

            
                print("lolololpoloplop")
                print(off_peak_kwh_used)
                if Total_Comsuption_kwh!=0 :
                    
                    on_peak_kwh_used=total_comsuption_kwh-off_peak_kwh_used
                    print("popopopopo")
                    print(on_peak_kwh_used)
                    Total_comsuption_kwh =off_peak_kwh_used + on_peak_kwh_used
                    if gross_receipts_tax == Reg_fee :
                        Reg_fee=0
                    
                    if on_peak_demand2==0:
                        on_peak_demand=on_peak_demand1
                        
                    else: 
                        on_peak_demand=on_peak_demand2
                    
                    Total_Services_Tax=Reg_fee +base_charge + Discretionary_sales+Service_Charge+franchise_fee+county_sales_tax+gross_receipts_tax1+franchise_charge+ utility_tax  + florida_sales_tax + gross_receipts_tax+Late_payment_charge +Power_monitoring_premium_plus+FPL_SolarTogether_charge+FPL_SolarTogether_credit
                    
                    
                    Total_comsuption_kwh = on_peak_kwh_used + off_peak_kwh_used
                    
                    
                    Energy_Charge= non_fuel_on_peak*on_peak_kwh_used+non_fuel_off_peak*off_peak_kwh_used
                    Energy_Charge_On_peak=non_fuel_on_peak*on_peak_kwh_used
                    Energy_Charge_Off_peak= non_fuel_off_peak*off_peak_kwh_used
                    Fuel_Charge= fuel_off_peak* off_peak_kwh_used + fuel_on_peak* on_peak_kwh_used
                    
                    Fuel_Charge_on_peak=fuel_on_peak* on_peak_kwh_used
                    Fuel_Charge_off_peak=fuel_off_peak* off_peak_kwh_used
                    Total_Energy_Charge=Energy_Charge+Fuel_Charge
                    Total_dolar_khw=Total_Energy_Charge/Total_comsuption_kwh
                    On_Peak_demand_Charge=demand_charge * on_peak_demand
                    Maximum_demand_Charge=maximum_demand * maximum
                    Total_Demand_Charge= On_Peak_demand_Charge + Maximum_demand_Charge
                    Total_Electric_cost=Total_Energy_Charge+ Total_Demand_Charge
                    
                    Total_Charge=Total_Electric_cost+ Total_Services_Tax
                
                    Energy_Rate= (Total_Energy_Charge)/(Total_comsuption_kwh)
                    Demand_Rate=(Total_Demand_Charge)/(maximum_demand +maximum)
                    extracted_data["Power monitoring-premium plus"] = Power_monitoring_premium_plus
                    extracted_data["Energy Charge"] = Energy_Charge
                    extracted_data["Energy Charge On peak"] = Energy_Charge_On_peak
                    extracted_data["Energy Charge Off peak"] = Energy_Charge_Off_peak
                    extracted_data["Fuel Charge"] = Fuel_Charge
                    extracted_data["Fuel Charge on peak $"] = Fuel_Charge_on_peak
                    extracted_data["Fuel Charge off peak $"] = Fuel_Charge_off_peak
                    extracted_data["Total Energy Charge"] = Total_Energy_Charge
                    extracted_data["Total Electric cost"] = Total_Electric_cost
                    extracted_data["Total Services and Tax"] = Total_Services_Tax
                    extracted_data["Total Charge"] = Total_Charge
                    extracted_data["Total $/kwh cost"] = Total_dolar_khw
                    extracted_data["On Peak Demand Charge"] = On_Peak_demand_Charge
                    
                    extracted_data["Maximum Demand Charge"] = Maximum_demand_Charge
                    extracted_data["Total Demand Charge TOU ($)"] = Total_Demand_Charge
                    extracted_data["Total Comsuption KWH"] = Total_comsuption_kwh
                    extracted_data["Demand charge:"] = Demand_charge_On_peak
                    if Demand_charge_On_peak==0:
                        extracted_data["Demand charge:"] =demand_charge
                    extracted_data["Total Demand"] =on_peak_demand
                    extracted_data["Energy Rate"] = Energy_Rate
                    extracted_data["Demand Rate"] = Demand_Rate
                    extracted_data["On-Peak kWh used"] = on_peak_kwh_used
                    if on_peak_demand2==0:
                        on_peak_demand2=on_peak_demand
                        extracted_data["On-peak demand2"] = on_peak_demand2
                    
                
         
        on_peak_kwh_used=extracted_data.get('On-peak kWh used')
        print("lllllllllllll")
        print(on_peak_kwh_used)
        if off_peak_kwh_used ==0:
            
            non_fuel = extracted_data.get("Non-fuel energy charge:", 0)
            fuel = extracted_data.get("Fuel charge:", 0)
            
            if non_fuel==0:
                non_fuel=(non_fuel_off_peak+non_fuel_on_peak)/2
                
            if fuel==0:
                fuel=(fuel_on_peak+fuel_off_peak)/2 
            
            #fuel = extracted_data.get("Fuel:", 0)
            print(non_fuel)
            #demand_kw = extracted_data.get("Demand KW", 0)
            demand = extracted_data.get("Demand charge:", 0)
            kwh_used = extracted_data.get("kWh Used", 0)
            
                                       
            usage1 = extracted_data.get("Demand KW", 0)
            usage2 = extracted_data.get("Usage", 0)
            print(usage1)
            print(usage2)
            print("emeoooooooooooooooooooooooooooooooooooooooo")
            if usage1 is None:
               usage1 = 0
            if usage2 is None:
               usage2 = 0

            usage=max(usage1,usage2)
            if usage is None or usage == 0:
               usage = extracted_data.get("Usage", 0)
            
            Total_Comsuption_kwh=kwh_used
            print(Total_Comsuption_kwh)
            base_charge =extracted_data.get("Base charge:", 0)
            Customer_charge =extracted_data.get("Customer charge:", 0)
            print(base_charge)
            if base_charge==0:
                base_charge=Customer_charge
            Reg_fee =extracted_data.get("Regulatory fee", 0)
            if Reg_fee == 0:
                Reg_fee =extracted_data.get("Regulatoiy fee (State fee)", 0)
            
            Gross_reciep =extracted_data.get("Gross receipts tax", 0)
            Gross_rec =extracted_data.get("Gross rec. tax/Regulatory fee", 0)
            if Gross_reciep:
                if Reg_fee==Gross_reciep :
                    Reg_fee=0
                    extracted_data["Gross receipts tax"] = 0
            if Reg_fee:
                if Gross_rec == Reg_fee:
                    Reg_fee=0
            Energy_Charge=kwh_used*non_fuel
            Fuel_Charge= kwh_used * fuel
            print(Fuel_Charge)
            print(kwh_used)
            
            
            
            
            utility_tax =extracted_data.get("Utility tax", 0)
            franchise_fee = extracted_data.get("Franchise fee", 0)
            franchise_charge = extracted_data.get("Franchise charge", 0)
            
            
            #Total_Comsuption_kwh=extracted_data.get("kWh Used", 0)
            florida_sales_tax =extracted_data.get("Florida sales tax", 0)
            Discretionary_sales =extracted_data.get("Discretionary sales surtax", 0)
            county_sales_tax =extracted_data.get("County sales tax", 0)
            Contract_demand =extracted_data.get("Contract demand", 0)
            Late_payment_charge =extracted_data.get( "Late payment charge", 0)
            FPL_SolarTogether_charge =extracted_data.get( "FPL SolarTogether charge", 0)
            FPL_SolarTogether_credit =extracted_data.get( "FPL SolarTogether credit", 0)
            print(usage)
            print(demand)
            print("regfeee")
            print(Reg_fee)
            if on_peak_kwh_used !=0:
               kwh_used=Total_Comsuption_kwh        
            
            Total_Energy_Charge=Energy_Charge+Fuel_Charge
            if Contract_demand !=0:
                
               Total_Demand_Charge= usage * demand + Contract_demand * demand
            else: 
                Total_Demand_Charge=usage * demand 
            Total_Electric_cost=Total_Energy_Charge+ Total_Demand_Charge
            Total_Services_Tax=Gross_rec + Gross_reciep + utility_tax + franchise_fee + franchise_charge+base_charge + Reg_fee+florida_sales_tax+Discretionary_sales+county_sales_tax+Late_payment_charge+FPL_SolarTogether_credit+FPL_SolarTogether_charge+Power_monitoring_premium_plus
            Total_Charge=Total_Electric_cost+ Total_Services_Tax
            Energy_Rate= (Total_Energy_Charge)/(kwh_used)
            Demand_Rate=(Total_Demand_Charge)/(usage)
            Total_dolar_khw=Total_Energy_Charge/Total_Comsuption_kwh
            if usage==0:
                usage=1
            extracted_data["Energy Charge"] = Energy_Charge
            extracted_data["Fuel Charge"] = Fuel_Charge
            extracted_data["Total Energy Charge"] = Total_Energy_Charge
            extracted_data["Total Electric cost"] = Total_Electric_cost
            extracted_data["Total Services and Tax"] = Total_Services_Tax
            extracted_data["Total Charge"] = Total_Charge
            extracted_data["Total Comsuption KWH"] = Total_Comsuption_kwh
            extracted_data["Total Energy Charge"] = Total_Energy_Charge
            extracted_data["Total Demand Charge - Non TOU ($)"] = Total_Demand_Charge
            if Contract_demand!=0:
                Contract_demand=Contract_demand-1
            extracted_data["Total Demand"] = usage + Contract_demand
            extracted_data["Energy Rate"] = Energy_Rate
            extracted_data["Demand Rate"] = Demand_Rate
            extracted_data["Total $/kwh cost"] = Total_dolar_khw
            extracted_data["Total $/kwh cost"] = Total_dolar_khw
    # Inside the extract_data function, after extracting all other values
    else:
      print("dorostedoroste")
      if extracted_data.get("Rate", "") == "GSD-1 GENERAL":
            # Your logic for GSD-1 GENERAL goes here
            print("Rate is GSD-1 GENERAL")
      else:
            print("Rate is not GSD-1 GENERAL")
            
      non_fuel = extracted_data.get("Non-fuel:", 0)
      fuel = extracted_data.get("Fuel:", 0)
      #demand_kw = extracted_data.get("Demand KW", 0)
      demand = extracted_data.get("Demand:", 0)
      kwh_used = extracted_data.get("kWh Used", 0)
      usage = extracted_data.get("Usage", 0)
      base_charge =extracted_data.get("Base charge:", 0)
      Gross_rec =extracted_data.get("Gross rec. tax/Regulatory fee", 0)
      Gross_reciep =extracted_data.get("Gross receipts tax", 0)
      utility_tax =extracted_data.get("Utility tax", 0)
      franchise_fee = extracted_data.get("Franchise fee", 0)
      franchise_charge = extracted_data.get("Franchise charge", 0)
      Reg_fee =extracted_data.get("Regulatory fee", 0)
      Customer_charge =extracted_data.get("Customer charge:", 0)
      #Total_Comsuption_kwh=extracted_data.get("kWh Used", 0)
      florida_sales_tax =extracted_data.get("Florida sales tax", 0)
      Discretionary_sales =extracted_data.get("Discretionary sales surtax", 0)
      county_sales_tax =extracted_data.get("County sales tax", 0)
      Contract_demand =extracted_data.get("Contract demand", 0)
      Late_payment_charge =extracted_data.get( "Late payment charge", 0)
      
            
      Total_Comsuption_kwh=kwh_used
      
      if base_charge==0:
         base_charge=Customer_charge

      if Gross_reciep:
        if Reg_fee==Gross_reciep :
            Reg_fee=0
            extracted_data["Gross receipts tax"] = 0
      if Reg_fee:
        if Gross_rec == Reg_fee:
            Reg_fee=0
      Energy_Charge=kwh_used*non_fuel
      Fuel_Charge= kwh_used * fuel
      Total_Energy_Charge=Energy_Charge+Fuel_Charge
      Total_Demand_Charge= usage * demand + Contract_demand * demand
      Total_Electric_cost=Total_Energy_Charge+ Total_Demand_Charge
      Total_Services_Tax=Gross_rec + Gross_reciep + utility_tax + franchise_fee + franchise_charge+base_charge + Reg_fee+florida_sales_tax+Discretionary_sales+county_sales_tax+Late_payment_charge
      Total_Charge=Total_Electric_cost+ Total_Services_Tax
      Energy_Rate= (Total_Energy_Charge)/(kwh_used)
      Total_dolar_khw=Total_Energy_Charge/Total_Comsuption_kwh
      if usage==0:
         usage=1
      Demand_Rate=(Total_Demand_Charge)/(usage)
      extracted_data["Energy Charge"] = Energy_Charge
      extracted_data["Fuel Charge"] = Fuel_Charge
      extracted_data["Total Energy Charge"] = Total_Energy_Charge
      extracted_data["Total Electric cost"] = Total_Electric_cost
      extracted_data["Total Services and Tax"] = Total_Services_Tax
      extracted_data["Total Charge"] = Total_Charge
      extracted_data["Total Comsuption KWH"] = Total_Comsuption_kwh
      extracted_data["Total Energy Charge"] = Total_Energy_Charge
      extracted_data["Total Demand Charge - Non TOU ($)"] = Total_Demand_Charge
      if Contract_demand!=0:
          Contract_demand=Contract_demand-1
      extracted_data["Total Demand"] = usage + Contract_demand
      extracted_data["Energy Rate"] = Energy_Rate
      extracted_data["Demand Rate"] = Demand_Rate
      extracted_data["Total $/kwh cost"] = Total_dolar_khw

      

    # If "Service to" was not found in the current PDF, mark it as NaN
    if not service_to_flag:
        extracted_data["Service to"] = float('nan')

    return extracted_data


keywords=["Rate", "Service to","Service days", "Total Comsuption KWH", "Energy Charge", "Fuel:", "Fuel Charge","Fuel Charge on peak $","Fuel Charge off peak $", "Non-fuel:", "Energy Charge On peak","Energy Charge Off peak","Total Energy Charge", "Total $/kwh cost",
                 "Usage", "Total Demand Charge - Non TOU ($)","Total Demand Charge TOU ($)","Contract demand", "Total Electric cost", "Base charge:", "Gross rec. tax/Regulatory fee", "Franchise charge", "Franchise fee", "Utility tax",
                 "Florida sales tax", "Discretionary sales surtax", "Taxes and charges", "Gross receipts tax", "Regulatory fee","Regulatory fee (State fee)", "County sales tax", "Service Charge", "On-Peak kWh used", 
                 "Off-peak kWh used", "On-peak demand","FPL SolarTogether charge","FPL SolarTogether credit", "Maximum demand","Demand KW","kWh Used","Demand charge:","Maximum","Non-fuel energy charge: on-peak","Late payment charge",
                 "Non-fuel energy charge: off-peak","Regulatoiy fee (State fee)", "Fuel charge-On-peak", "Fuel charge-Off-peak", "Total Charge", "Energy Rate", "Demand Rate", "Demand:","Customer charge:","On-peak demand2","Power monitoring-premium plus"]
# Define keywords to search for

# Create a directory to store text files
if not os.path.exists("text_files"):
    os.makedirs("text_files")

# Create a list to store data for each account
data_by_account = []

# Create a list to store the PDF file names
pdf_file_names = []

# Create a set to keep track of processed PDF file names
processed_pdf_files = set()

# Create a dictionary to accumulate the values across all accounts
cumulative_data = {
    "Taxes and charges_A": 0,
    "Total charges": 0,
    "Total Comsuption KWH": 0,
    "Total Energy Charge": 0,
    "Total Demand Charge": 0,
}

def extract_and_consolidate_data(uploaded_files_by_account, num_accounts, coefficients):
    # List to hold data for all accounts
    data_by_account = []

    for account_number, uploaded_files in uploaded_files_by_account.items():
        data_for_account = []
        months_present = []  # List to store months present for this account

        # Loop through each uploaded file
        for uploaded_file in uploaded_files:
            pdf_path = uploaded_file.name  # Use the uploaded file's name as the path
            # Read the file data
            extracted_data = extract_data(uploaded_file)
            data_for_account.append(extracted_data)
            print(f"Extracted data from {pdf_path} for Account {account_number + 1}")
            # Extract the "Service Month" from the extracted data
            service_month = extracted_data.get("Service Month")
            if service_month:
                months_present.append(service_month)

        data_by_account.append(data_for_account)
            
            

    # Create an Excel file with separate sheets for each account
    excel_filename = 'all_accounts_data.xlsx'
    with pd.ExcelWriter(excel_filename, engine='xlsxwriter') as excel_writer:
        data_by_account_transposed = []  # List to store transposed data for each account
           
        missing_months_by_account = []
        for i, account_data in enumerate(data_by_account):
            accounts_usage= {}
            account_sheet_name = f'Account_{i + 1}'
            df_account = pd.DataFrame(account_data)
            # Transpose the DataFrame
            #check_values = calculate_check_values(start_month)
            missing_months = []  # Initialize the variable
            # Transpose the DataFrame
            df_account_transposed = df_account.transpose()
            print(df_account_transposed)
            # Check if "Service Month" is in the index (rows)
            if "Service Month" in df_account_transposed.index:
                # Extract the "Service Month" row
                service_month_row = df_account_transposed.loc["Service Month"]

                # Define a set of all months from 1 to 12
                all_months = set(range(1, 13))

                # Extract the actual service months as integers
                service_months = {int(month) for month in service_month_row.values if str(month).isdigit()}
                
                # Find the missing months
                missing_months_account = sorted(list(all_months - service_months))
                missing_months_account1=missing_months_account
                # Now you have the missing months for this account
                print(f"Missing months for {account_sheet_name}: {missing_months_account}")
                if not missing_months_account1:
                # Get the values in the "Service Month" row
                    service_month_values = df_account_transposed.loc["Service Month"]

                    # Rename the columns with the corresponding values
                    df_account_transposed.columns = service_month_values
                miss_NH = set()         # Initialize a set for missing months without both previous and next months               
                index_of_average=0
                
                column_values_month = []  # To store the result
                for month in missing_months_account1:
                    print("mise youuuuuu")
                    previous_month = month - 1
                    next_month = month + 1
                    column_values_month=[]
                    if previous_month== 0:
                        previous_month=12
                    if next_month==13:
                        next_month=1
                    
                    
                    if previous_month not in missing_months_account1 and next_month not in missing_months_account1:
                        miss_NH.add(month)
                        
                        column_name_p = df_account_transposed.columns[df_account_transposed.loc["Service Month"] == previous_month].item()
                        column_values_p = df_account_transposed[column_name_p]
                        
                        # Convert next_month to an integer
                        next_month = int(next_month)
                        if next_month==13: 
                            next_month=1
                        # Check if next_month exists in the "Service Month" row
                        if next_month in df_account_transposed.loc["Service Month"].values:
                            # Get the corresponding column name
                            column_name_n = df_account_transposed.columns[df_account_transposed.loc["Service Month"] == next_month].item()
                            column_values_n = df_account_transposed[column_name_n]
                            #print(column_values_n)
                        else:
                            print(f"Next month {next_month} not found in 'Service Month'.")
                        
                        
                        for n, p in zip(column_values_n, column_values_p):
                            try:
                                # Try to convert n and p to float and perform the operation
                                n_float = float(n)
                                p_float = float(p)
                                result = (n_float + p_float) / 2
                                column_values_month.append(result)
                            except ValueError:
                                # Handle non-numeric values
                                column_values_month.append(None)  # You can use None or any other value as needed
                                
                        
                        number_of_members = len(missing_months_account1)
                        index_of_average=12-number_of_members 
                        #column_name = str(index_of_average)
                        column11=index_of_average+month
                        #print(column11)
                        if column11== 12:
                            column11=112
                        df_account_transposed[column11] = column_values_month
                        df_account_transposed.at["Late payment charge", column11] = 0
                        df_account_transposed.at["Service Month",column11] = month
                        
                        
                        missing_months_account1.remove(month)
                # Find the exact column name for "12"
                column_name = None
                for col in df_account_transposed.columns:
                    if col == 18:
                        column_name = col
                        break
                # Get the values in the "Service Month" row
                service_month_values = df_account_transposed.loc["Service Month"]

                # Rename the columns with the corresponding values
                df_account_transposed.columns = service_month_values
                            
                if column_name is not None:
                    # Change the value of "Service Month" in the found column to 12
                    df_account_transposed.at["Service Month", column_name] = 12
                    #print("jujujujuju")
                    #print(df_account_transposed)
                print("Missing months without both previous and next months:", miss_NH)
                missing_months_account = list(set(missing_months_account1) - miss_NH)
                print(missing_months_account1)  
                #print("iiii") 
                if not missing_months_account1 :
                    print("empty")
                else:          
                    # Flatten the list of coefficients
                    # print("1111111111111111111111111111111111111")
                    # flat_coefficients = [coeff[0] for coeff in coefficients]
                    # Find the maximum coefficient value and its month
                    
                    # Flatten the list of coefficients
                    print("1111111111111111111111111111111111111")
                    flat_coefficients = coefficients  # Direct assignment since it's already a flat list
                    print("Flat Coefficients:", flat_coefficients)

                    # Find the maximum coefficient value and its month
                    max_coefficient = max(flat_coefficients)
                    max_coefficient_month = flat_coefficients.index(max_coefficient) + 1  # Adding 1 to make it 1-indexed

                    service_month_row = df_account_transposed.loc["Service Month"]
                    print("Service Month Row:", service_month_row)
                    print("Max Coefficient Month:", max_coefficient_month)

                    
                    
                    
                    
                    
                    
                    # max_coefficient = max(flat_coefficients)
                    # max_coefficient_month = flat_coefficients.index(max_coefficient) + 1    
                    # service_month_row = df_account_transposed.loc["Service Month"]
                    print(service_month_row)
                    print(2222222222222222222222222222222222222)
                    print(max_coefficient_month)
                    # Check if 1 is missing
                    if max_coefficient_month not in missing_months_account and missing_months_account:
                        # Find the row where "Service Month" is 1
                        print("nnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnn")
                        index_of_1 = service_month_row[service_month_row == max_coefficient_month].index[0]
                        # Extract the corresponding column
                        based_anchorm_co = df_account_transposed[index_of_1]
                        
                        
                        value1 = df_account_transposed.loc["Demand:",  index_of_1]
                        value2 = df_account_transposed.loc["Fuel:",  index_of_1]
                        value3 = df_account_transposed.loc["Fuel Charge on peak $",  index_of_1]
                        value4 = df_account_transposed.loc["Fuel Charge off peak $",  index_of_1]
                        value5 = df_account_transposed.loc["Non-fuel:",  index_of_1]
                        value6 = df_account_transposed.loc["Base charge:",  index_of_1]
                        value7 = df_account_transposed.loc["Gross rec. tax/Regulatory fee",  index_of_1]
                        value8 = df_account_transposed.loc[ "Franchise charge",  index_of_1]
                        value9 = df_account_transposed.loc["Franchise fee",  index_of_1]
                        value10 = df_account_transposed.loc["Utility tax",  index_of_1]
                        value11 = df_account_transposed.loc["Florida sales tax",  index_of_1]
                        value12 = df_account_transposed.loc["Discretionary sales surtax",  index_of_1]
                        value13 = df_account_transposed.loc["Gross receipts tax",  index_of_1]
                        value14 = df_account_transposed.loc["Regulatory fee",  index_of_1]
                        value15 = df_account_transposed.loc["County sales tax",  index_of_1]
                        value16 = df_account_transposed.loc["Service Charge",  index_of_1]
                        value17 = df_account_transposed.loc["Maximum",  index_of_1]
                        value18 = df_account_transposed.loc["Demand charge:",  index_of_1]
                        value19 = df_account_transposed.loc["Non-fuel energy charge: on-peak",  index_of_1]
                        value20 = df_account_transposed.loc["Non-fuel energy charge: off-peak", index_of_1]
                        value21 = df_account_transposed.loc["Fuel charge-On-peak",  index_of_1]
                        value22 = df_account_transposed.loc["Fuel charge-Off-peak",  index_of_1]
                        value23 = df_account_transposed.loc["Customer charge:", index_of_1]
                        value24 = df_account_transposed.loc["On-peak demand2",  index_of_1]
                        value25 = df_account_transposed.loc["Service days",  index_of_1]
                        
                    
                        # Modify the DataFrame for missing months
                        
                        for missing_month in missing_months_account:
                            
                        
                            def multiply_by_constant(value):
                                    try:
                                        missing_month1 = missing_month - 1
                                        coefficient = float(coefficients[missing_month1][0])
                                        print(f"Coefficient for missing month {missing_month1}: {coefficient}")
                                        
                                        return float(value) * coefficient
                                        
    
                                    except (ValueError, TypeError):
                                        return value

                                # Apply the function to each element in the DataFrame
                            based_anchorm_co12 = based_anchorm_co.apply(multiply_by_constant)
                            service_month_values = df_account_transposed.loc["Service Month"]
                            df_account_transposed.columns = service_month_values
                            df_account_transposed[missing_month] = based_anchorm_co12   
                            df_account_transposed.at["Late payment charge", missing_month] = 0
                            df_account_transposed.at["Demand:", missing_month] = value1
                            df_account_transposed.at["Fuel:", missing_month] = value2
                            df_account_transposed.at["Fuel Charge on peak $", missing_month] = value3
                            df_account_transposed.at["Fuel Charge off peak $", missing_month] = value4
                            df_account_transposed.at["Non-fuel:", missing_month] = value5
                            df_account_transposed.at["Base charge:", missing_month] = value6
                            df_account_transposed.at["Gross rec. tax/Regulatory fee", missing_month] = value7
                            df_account_transposed.at["Franchise charge", missing_month] = value8
                            df_account_transposed.at["Franchise fee", missing_month] = value9
                            df_account_transposed.at["Utility tax", missing_month] = value10
                            df_account_transposed.at["Florida sales tax", missing_month] = value11
                            df_account_transposed.at["Discretionary sales surtax", missing_month] = value12
                            df_account_transposed.at["Gross receipts tax", missing_month] = value13
                            df_account_transposed.at["Regulatory fee", missing_month] = value14
                            df_account_transposed.at["County sales tax", missing_month] = value15
                            df_account_transposed.at["Service Charge", missing_month] = value16
                            df_account_transposed.at["Maximum", missing_month] = value17
                            df_account_transposed.at["Demand charge:", missing_month] = value18
                            df_account_transposed.at["Non-fuel energy charge: on-peak", missing_month] = value19
                            df_account_transposed.at["Non-fuel energy charge: off-peak", missing_month] = value20
                            df_account_transposed.at["Fuel charge-On-peak", missing_month] = value21
                            df_account_transposed.at["Fuel charge-Off-peak", missing_month] = value22
                            df_account_transposed.at["Customer charge:", missing_month] = value23
                            df_account_transposed.at["On-peak demand2", missing_month] = value24
                            df_account_transposed.at["Service days", missing_month] = value25
                            df_account_transposed.at['Service Month', missing_month] = missing_month                   
                            
                            

                                            
                        
                    elif max_coefficient_month  in missing_months_account and missing_months_account:
                            alpha = next(i for i in range(1, 13) if i not in missing_months_account)
                            print(f"Selected alpha: {alpha}") 
                            # Find the row where "Service Month" is equal to alpha
                            index_of_alpha = service_month_row[service_month_row == alpha].index[0]
                            value1 = df_account_transposed.loc["Demand:", index_of_alpha]
                            value2 = df_account_transposed.loc["Fuel:", index_of_alpha]
                            value3 = df_account_transposed.loc["Fuel Charge on peak $", index_of_alpha]
                            value4 = df_account_transposed.loc["Fuel Charge off peak $",index_of_alpha]
                            value5 = df_account_transposed.loc["Non-fuel:", index_of_alpha]
                            value6 = df_account_transposed.loc["Base charge:",index_of_alpha]
                            value7 = df_account_transposed.loc["Gross rec. tax/Regulatory fee", index_of_alpha]
                            value8 = df_account_transposed.loc[ "Franchise charge", index_of_alpha]
                            value9 = df_account_transposed.loc["Franchise fee", index_of_alpha]
                            value10 = df_account_transposed.loc["Utility tax", index_of_alpha]
                            value11 = df_account_transposed.loc["Florida sales tax",index_of_alpha]
                            value12 = df_account_transposed.loc["Discretionary sales surtax", index_of_alpha]
                            value13 = df_account_transposed.loc["Gross receipts tax", index_of_alpha]
                            value14 = df_account_transposed.loc["Regulatory fee", index_of_alpha]
                            value15 = df_account_transposed.loc["County sales tax", index_of_alpha]
                            value16 = df_account_transposed.loc["Service Charge", index_of_alpha]
                            value17 = df_account_transposed.loc["Maximum", index_of_alpha]
                            value18 = df_account_transposed.loc["Demand charge:", index_of_alpha]
                            value19 = df_account_transposed.loc["Non-fuel energy charge: on-peak", index_of_alpha]
                            value20 = df_account_transposed.loc["Non-fuel energy charge: off-peak", index_of_alpha]
                            value21 = df_account_transposed.loc["Fuel charge-On-peak", index_of_alpha]
                            value22 = df_account_transposed.loc["Fuel charge-Off-peak", index_of_alpha]
                            value23 = df_account_transposed.loc["Customer charge:",index_of_alpha]
                            value24 = df_account_transposed.loc["On-peak demand2", index_of_alpha] 
                            value25 = df_account_transposed.loc["Service days",  index_of_alpha]    
                            
                            
                            # Modify the DataFrame for missing months
                            # Extract the corresponding column
                            based_anchorm_co1 = df_account_transposed[index_of_alpha]
                            
                            # Get the values in the "Service Month" row
                            service_month_values = df_account_transposed.loc["Service Month"]

                            # Rename the columns with the corresponding values
                            df_account_transposed.columns = service_month_values
                            
                            
                            for missing_month in missing_months_account:
                                    print("dididididididi")
                                
                                    def multiply_by_constant(value):
                                        try:
                                            missing_month1 = missing_month - 1
                                            coefficient = float(coefficients[missing_month1][0])
                                            print(f"Coefficient for missing month {missing_month1}: {coefficient}")
                                            
                                            return float(value) * coefficient
                                        

                                        
                                        except (ValueError, TypeError):
                                            return value
                                
                                    

                                        # Apply the function to each element in the DataFrame
                                    based_anchorm_co122 = based_anchorm_co1.apply(multiply_by_constant)
                                    df_account_transposed[missing_month] = based_anchorm_co122
                                    df_account_transposed.at["Late payment charge", missing_month] = 0 
                                    df_account_transposed.at["Demand:",  missing_month] = value1
                                    df_account_transposed.at["Fuel:",  missing_month] = value2
                                    df_account_transposed.at["Fuel Charge on peak $",  missing_month] = value3
                                    df_account_transposed.at["Fuel Charge off peak $",  missing_month] = value4
                                    df_account_transposed.at["Non-fuel:",  missing_month] = value5
                                    df_account_transposed.at["Base charge:",  missing_month] = value6
                                    df_account_transposed.at["Gross rec. tax/Regulatory fee",  missing_month] = value7
                                    df_account_transposed.at["Franchise charge",  missing_month] = value8
                                    df_account_transposed.at["Franchise fee",  missing_month] = value9
                                    df_account_transposed.at["Utility tax", missing_month] = value10
                                    df_account_transposed.at["Florida sales tax",  missing_month] = value11
                                    df_account_transposed.at["Discretionary sales surtax",  missing_month] = value12
                                    df_account_transposed.at["Gross receipts tax",  missing_month] = value13
                                    df_account_transposed.at["Regulatory fee",  missing_month] = value14
                                    df_account_transposed.at["County sales tax",  missing_month] = value15
                                    df_account_transposed.at["Service Charge",  missing_month] = value16
                                    df_account_transposed.at["Maximum",  missing_month] = value17
                                    df_account_transposed.at["Demand charge:", missing_month] = value18
                                    df_account_transposed.at["Non-fuel energy charge: on-peak",  missing_month] = value19
                                    df_account_transposed.at["Non-fuel energy charge: off-peak",  missing_month] = value20
                                    df_account_transposed.at["Fuel charge-On-peak",  missing_month] = value21
                                    df_account_transposed.at["Fuel charge-Off-peak",  missing_month] = value22
                                    df_account_transposed.at["Customer charge:",  missing_month] = value23
                                    df_account_transposed.at["On-peak demand2", missing_month] = value24
                                    df_account_transposed.at["Service days", missing_month] = value25
                                    value_at_position12 = df_account_transposed.at["Service to", missing_month]
                                    column_names = df_account_transposed.columns                          

                                    # Assuming missing_month is a column
                                    column_name = str(missing_month)
                                                                
                                    # Assuming df_account_transposed is your DataFrame
                                    # Update the "Service to" row with the new month
                                
                            month_map = {
                                            1: 'Jan', 2: 'Feb', 3: 'Mar', 4: 'Apr',
                                            5: 'May', 6: 'Jun', 7: 'Jul', 8: 'Aug',
                                            9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dec'
                                        }             
                                                
                            # Extract the column names as integers
                            column_names = df_account_transposed.columns

                            # Convert column names to month names using the provided 'month_map'
                            month_names = [month_map[column_name] for column_name in column_names]

                        # Ensure "Service to" row is a datetime series
                            df_account_transposed.loc["Service to", :] = pd.to_datetime(df_account_transposed.loc["Service to", :], errors='coerce')

                            # Extract the day and year from the existing datetime values
                            day_year_str = df_account_transposed.loc["Service to", :].apply(lambda x: x.strftime('%d, %Y') if pd.notna(x) else '')

                            # Replace NaN values (non-datetime) with the new month_names
                            df_account_transposed.loc["Service to", :] = day_year_str + ' ' + month_names

                            # Ensure "Service to" row is a datetime series again
                            df_account_transposed.loc["Service to", :] = pd.to_datetime(df_account_transposed.loc["Service to", :], errors='coerce')


                            print(f"Based on 'Service Month' = {alpha}: {based_anchorm_co1}")
               
            
            
            
            #row_label = "Service to"
            #column_names = df_account_transposed.columns
            print(df_account_transposed)
            Rate_account={}
            Rate_account1=df_account_transposed.loc['Rate']
            Rate_account[account_number]=Rate_account1.iloc[1]
            
            # Now, let's find the corresponding values in each column for the specified row
            #values_for_row = df_account_transposed.loc[row_label]               
            #value_at_position = df_account_transposed.at["Service to", 3]
            print(missing_months_account)
            #print("injaaa")
            if missing_months_account:
                                
                # Extract the values of "Service Month"
                service_month_values = df_account_transposed.loc["Service Month"].astype(int)

                # Sort the columns based on "Service Month" values
                sorted_columns = service_month_values.sort_values().index

                # Sort the DataFrame based on sorted columns
                df_account_transposed = df_account_transposed[sorted_columns]
                
                 
            else:
                
            # Extract the 'Service to' values from the transposed DataFrame
                service_to_values = pd.to_datetime(df_account_transposed.loc["Service to"], format='%b %d, %Y')

            # Sort the columns based on the "Service to" dates
                sorted_columns = service_to_values.sort_values().index
                            
                df_account_transposed = df_account_transposed[sorted_columns]
            
            # Calculate the sum of each row excluding "Service to" and "Rate"
            excluded_columns = ["Service to", "Rate"]
            sums = df_account_transposed.drop(excluded_columns, axis=0).sum(axis=1)

            # Add a new "Sum" column with the sums
            df_account_transposed["Sum"] = sums
            
            # Append the sorted and modified DataFrame to the list
            row_to_update_demand_rate = df_account_transposed.loc["Demand:"]
                        
            data_by_account_transposed.append(df_account_transposed)
            ##############################################################
            Demand_KW_Month = df_account_transposed.loc['Usage'] 
            accounts_usage[account_number] = Demand_KW_Month
            
            consolidate_transpose_df = pd.concat(data_by_account_transposed).groupby(level=0).sum()
            #print(consolidate_transpose_df)         
            # Filter the DataFrame to obtain the row with the label "Demand charge:"
            demand_charge_row = consolidate_transpose_df.loc["Demand charge:"]
            #print(demand_charge_row)
            print(consolidate_transpose_df)
            # Extract the values from the desired column (replace 'column_name' with the actual column name)
            value = demand_charge_row[11]  # Replace 'column_name' with the actual column name
            column_values = value
            consolidate_transpose_df.at["Demand charge:", 'Sum'] = column_values
            consolidate_transpose_df.loc["Demand:"] =row_to_update_demand_rate
            
            # Filter the DataFrame to obtain the row with the label "Demand charge:"
            Maximum_row = consolidate_transpose_df.loc["Maximum"]
            value = Maximum_row[11]  # Replace 'column_name' with the actual column name
            column_values = value
            # Extract the values from the desired column (replace 'column_name' with the actual column name)
            #column_values = demand_charge_row[11].tolist()
            consolidate_transpose_df.at["Maximum", 'Sum'] = column_values
            consolidate_transpose_df.loc["Service days"] = consolidate_transpose_df.loc["Service days"]/num_accounts
            
            # Calculate the new value
            new_value = (consolidate_transpose_df.at["Energy Charge", "Sum"] / consolidate_transpose_df.at["Total Comsuption KWH", "Sum"])

            # Update the value in the DataFrame
            consolidate_transpose_df.at["Non-fuel:", "Sum"] = new_value

            # Calculate the new value
            new_value = (consolidate_transpose_df.at["Fuel Charge", "Sum"] / consolidate_transpose_df.at["Total Comsuption KWH", "Sum"])

            # Update the value in the DataFrame
            consolidate_transpose_df.at["Fuel:", "Sum"] = new_value
            
            consolidate_transpose_df.loc["Total Demand Charge"]=consolidate_transpose_df.loc[ "Total Demand Charge - Non TOU ($)"]+consolidate_transpose_df.loc["Total Demand Charge TOU ($)"]
            
            #if not consolidate_transpose_df.loc["Usage"].empty and (consolidate_transpose_df.loc["Usage"] != 0).any():
            if not consolidate_transpose_df.loc["Usage"].empty and not (consolidate_transpose_df.loc["Usage"] == 0).any():
                valueeee=consolidate_transpose_df.loc["Usage"]
                consolidate_transpose_df.loc["Demand:"] = consolidate_transpose_df.loc["Total Demand Charge - Non TOU ($)"]/consolidate_transpose_df.loc["Usage"]
            # Calculate the values for the new row "Demand $/kwh"
            demand_per_kwh = consolidate_transpose_df.loc["Total Demand Charge"] / consolidate_transpose_df.loc["Total Demand"]
            # Add the new row to the DataFrame
            consolidate_transpose_df.loc["Demand $/kwh"] = demand_per_kwh
            consolidate_transpose_df.loc["Non-TOU Consumption KWH"]=consolidate_transpose_df.loc["Total Comsuption KWH"]-consolidate_transpose_df.loc["On-Peak kWh used"]-consolidate_transpose_df.loc["Off-peak kWh used"]
         
            consolidate_transpose_df.loc["Energy Charge Non-TOU ($)"]=consolidate_transpose_df.loc["Energy Charge"]-consolidate_transpose_df.loc["Energy Charge On peak"]-consolidate_transpose_df.loc["Energy Charge Off peak"]
            if not consolidate_transpose_df.loc["Non-TOU Consumption KWH"].empty and (consolidate_transpose_df.loc["Non-TOU Consumption KWH"] != 0).any():
               consolidate_transpose_df.loc["Energy $/kwh on Non-TOU"]=consolidate_transpose_df.loc["Energy Charge Non-TOU ($)"]/consolidate_transpose_df.loc["Non-TOU Consumption KWH"]

            if not consolidate_transpose_df.loc["Fuel Charge"].empty and (consolidate_transpose_df.loc["Fuel Charge"]!=0).any():
               consolidate_transpose_df.loc["Fuel Charge Non-TOU $"]=consolidate_transpose_df.loc["Fuel Charge"]-consolidate_transpose_df.loc["Fuel Charge on peak $"]-consolidate_transpose_df.loc["Fuel Charge off peak $"]
             

            if not consolidate_transpose_df.loc["Non-TOU Consumption KWH"].empty and (consolidate_transpose_df.loc["Non-TOU Consumption KWH"] != 0).any():
               consolidate_transpose_df.loc["Fuel $/KWH Non-TOU"]=consolidate_transpose_df.loc["Fuel Charge Non-TOU $"]/consolidate_transpose_df.loc["Non-TOU Consumption KWH"]

            
            consolidate_transpose_df.loc["Fuel:"] = consolidate_transpose_df.loc["Fuel Charge"] / consolidate_transpose_df.loc["Total Comsuption KWH"]
            consolidate_transpose_df.loc["Non-fuel:"] = consolidate_transpose_df.loc["Energy Charge"] / consolidate_transpose_df.loc["Total Comsuption KWH"]
            consolidate_transpose_df.loc[ "Total $/kwh cost"] = consolidate_transpose_df.loc["Total Energy Charge"] / consolidate_transpose_df.loc["Total Comsuption KWH"]
            consolidate_transpose_df.loc[ "Energy Rate"] = consolidate_transpose_df.loc["Total Energy Charge"] / consolidate_transpose_df.loc["Total Comsuption KWH"]
            consolidate_transpose_df.loc[ "Demand Rate"] = consolidate_transpose_df.loc["Total Demand Charge"] / consolidate_transpose_df.loc["Total Demand"]
            
            
               
            if consolidate_transpose_df.at["On-Peak kWh used", "Sum"]!=0:
                # Calculate the new value
                new_value = (consolidate_transpose_df.at["Energy Charge On peak", "Sum"] / consolidate_transpose_df.at["On-Peak kWh used", "Sum"])

                # Update the value in the DataFrame
                consolidate_transpose_df.at["Non-fuel energy charge: on-peak", "Sum"] = new_value 
            
            if consolidate_transpose_df.at["Off-peak kWh used", "Sum"]!=0:
                # Calculate the new value
                new_value = (consolidate_transpose_df.at["Energy Charge Off peak", "Sum"] / consolidate_transpose_df.at["Off-peak kWh used", "Sum"])

                # Update the value in the DataFrame
                consolidate_transpose_df.at["Non-fuel energy charge: off-peak", "Sum"] = new_value 
            
            
            if consolidate_transpose_df.at["On-Peak kWh used", "Sum"]!=0:
                # Calculate the new value
                new_value = (consolidate_transpose_df.at["Fuel Charge on peak $", "Sum"] / consolidate_transpose_df.at["On-Peak kWh used", "Sum"])

                # Update the value in the DataFrame
                consolidate_transpose_df.at["Fuel charge-On-peak", "Sum"] = new_value 
            
            if consolidate_transpose_df.at["Off-peak kWh used", "Sum"]!=0:
                # Calculate the new value
                new_value = (consolidate_transpose_df.at["Fuel Charge off peak $", "Sum"] / consolidate_transpose_df.at["Off-peak kWh used", "Sum"])

                # Update the value in the DataFrame
                consolidate_transpose_df.at["Fuel charge-Off-peak", "Sum"] = new_value 
                
            

            # Remove rows with "Rate" and "Service to" if present
            if "Rate" in consolidate_transpose_df.index:
                consolidate_transpose_df.drop("Rate", inplace=True)
            if "Service to" in consolidate_transpose_df.index:
                consolidate_transpose_df.drop("Service to", inplace=True)
                
            
           # Create a list to store the sum values for each quarter
            Total_amount_per_Qtr = []

            # Define the number of columns in each quarter
            columns_per_quarter = 3

            # Find the index of the "Total Charge" row
            total_charge_index = consolidate_transpose_df.index.get_loc("Total Charge")

            # Iterate through the DataFrame by selecting columns for each quarter
            
            for i in range(0, len(consolidate_transpose_df.columns), columns_per_quarter):
                quarter = consolidate_transpose_df.columns[i:i + columns_per_quarter]
                quarter_sum = consolidate_transpose_df.iloc[total_charge_index, i:i + columns_per_quarter].sum()
                Total_amount_per_Qtr.append(quarter_sum)
                
                
            # Ensure 'Quarter_sum' has 12 values, inserting zeros where needed
            while len(Total_amount_per_Qtr) < 13:
                Total_amount_per_Qtr.append(0)
            # Now, 'quarter_sum_df' contains the 'Quarter_sum' values
            
            # Convert 'Quarter_sum' to a DataFrame with the same columns as the original DataFrame
            quarter_sum_df = pd.DataFrame([Total_amount_per_Qtr], columns=consolidate_transpose_df.columns, index=["Total amount per Qtr2"])

            # Concatenate the 'quarter_sum_df' with the original DataFrame to add it as a new row
            consolidate_transpose_df = pd.concat([consolidate_transpose_df, quarter_sum_df])
            row_to_update = consolidate_transpose_df.loc["Total amount per Qtr2"]
            max_column = row_to_update.idxmax()
            consolidate_transpose_df.at["Total amount per Qtr2",max_column]=0
            
            # Create a list to hold the values for the new row
            new_row_values = []
            new22=consolidate_transpose_df.loc["Total amount per Qtr2"]
            # Convert the Series to a list
            new22_list = new22.tolist()

            # Find indices of non-zero elements
            non_zero_indices = [i for i, value in enumerate(new22_list) if value != 0]

            # Get corresponding column names from the DataFrame
            column_names = consolidate_transpose_df.columns[non_zero_indices]
            first_four_values = [new22_list[i] for i in non_zero_indices[:4]]

            
            # Convert the Series to a list
            new22_list = new22.tolist()

            # Initialize the values for "pw"
            pw_values = [0] * 13

            # Find indices of non-zero elements
            non_zero_indices = [i for i, value in enumerate(new22_list) if value != 0]

            # Assign the non-zero values to the appropriate positions in "pw"
            for i, index in enumerate([2, 5, 8, 11]):
                if i < len(non_zero_indices):
                    pw_values[index] = new22_list[non_zero_indices[i]]

            # Create a new row "pw" in the DataFrame with the values
            consolidate_transpose_df.loc["Total amount per Qtr"] = pw_values
             
            
            # Create a mapping from month numbers to month names
            month_mapping = {
                1: 'Jan',
                2: 'Feb',
                3: 'Mar',
                4: 'Apr',
                5: 'May',
                6: 'Jun',
                7: 'Jul',
                8: 'Aug',
                9: 'Sep',
                10: 'Oct',
                11: 'Nov',
                12: 'Dec',
                'Sum': 'Sum'
            }
            
            seasonal_consumption = calculate_seasonal_consumption(consolidate_transpose_df)

            # Rename the columns in the DataFrame using the mapping
            consolidate_transpose_df.columns = [month_mapping[col] for col in consolidate_transpose_df.columns]
            

            
            desired_order=["Service days",  "On-Peak kWh used", "Off-peak kWh used","Non-TOU Consumption KWH","kWh Used"," ","Energy Charge On peak","Energy Charge Off peak","Energy Charge Non-TOU ($)", "Energy Charge"," ",  
                "Non-fuel energy charge: on-peak","Non-fuel energy charge: off-peak","Energy $/kwh on Non-TOU","Non-fuel:"," ","Fuel Charge on peak $","Fuel Charge off peak $","Fuel Charge Non-TOU $", "Fuel Charge"," ","Fuel charge-On-peak", "Fuel charge-Off-peak","Fuel $/KWH Non-TOU","Fuel:"," ","Total Energy Charge", "Total $/kwh cost"," ",
                  "Usage","Contract demand", "On-peak demand2","Maximum demand","Total Demand"," ","Demand:","Demand charge:","Maximum"," ","Total Demand Charge - Non TOU ($)","Total Demand Charge TOU ($)", "Total Demand Charge", "Demand $/kwh"," ","Total Electric cost", " ","Base charge:","Service Charge", "Late payment charge"," ",
                  "Gross rec. tax/Regulatory fee", "Franchise charge", "Franchise fee", "Utility tax",
                 "Florida sales tax", "Discretionary sales surtax","FPL SolarTogether charge","FPL SolarTogether credit" , "Gross receipts tax", "Regulatory fee", "County sales tax", 
                 "Total Services and Tax"," ","Power monitoring-premium plus", "",

                   "Total Charge","Total amount per Qtr", " " ,"Energy Rate", "Demand Rate"]
            consolidate_transpose_df = consolidate_transpose_df.reindex(desired_order, axis=0)
            off_peak_kwh_used = extracted_data.get("Off-peak kWh used", 0)
            if off_peak_kwh_used !=0:
                 rename_dict = {
                 "Demand charge:" : "On-peak Demand $/kwh"}
            # Define a dictionary to map old row names to new row names
            rename_dict = {
                "Usage": "Total Demand kw - Non TOU", "Fuel :":"Fuel Charge $/kwh" , "Energy Charge" :"Total Energy Charge ($)",
                "Energy Charge On peak":"Energy Charge On peak ($)","Energy Charge Off peak": "Energy Charge Off peak ($)","Non-TOU Consumption KWH":"Non-TOU Consumption kwh",
                "Non-fuel:": "Average Energy $/kWh", "On-peak demand2":"On-peak demand kw","Contract demand" :"Contract demand kw",
                "On-Peak kWh used": "Consumption On Peak kwh","Fuel Charge" : "Total Fuel Charge ($)","Fuel Charge on peak $":"Fuel Charge on peak ($)","Fuel Charge off peak $":"Fuel Charge off peak ($)",
                "Off-peak kWh used": "Consumption off-Peak kwh","Total Energy Charge":"Total Energy & Fuel Charge ($)",
                "Demand:" : "Demand_$/kwh- Non TOU" , "Fuel:": "Average Fuel $/kWh",
                "Non-fuel energy charge: on-peak" : "Energy $/kwh on peak","Fuel Charge Non-TOU $":"Fuel Charge Non-TOU ($)",
                "Non-fuel energy charge: off-peak" :  "Energy $/kwh off peak","Total Electric cost":"Total Electric cost ($)",
                "Fuel charge-On-peak" : "Fuel Charge $/kwh on peak","Total Demand Charge":"Total Demand Charge ($)",
                 "Fuel charge-Off-peak" : "Fuel Charge $/kwh off peak","Total Demand": "Total Demand kw",
                 "Maximum demand" : "Maximum Demand kw" , "On-peak demand" : "On-peak Demand",
                 "Maximum" : "Maximum Demand $/kwh", "Base charge:" :"Base charge($)","Service Charge" : "Service Charge ($)",
                 "Late payment charge": "Late payment charge ($)",
                 "kWh Used" : "Total Comsuption kwh", "Energy Rate" : "Average $/kwh cost (Exc fees)", "Total Charge":"Total Charge ($)",
                 "Total Services and Tax": "Total Services and Tax ($)",
                 "Gross rec. tax/Regulatory fee": "Gross rec. tax/Regulatory fee ($)", "Franchise charge":"Franchise charge ($)", 
                 "Franchise fee": "Franchise fee ($)", "Utility tax":"Utility tax ($)","Florida sales tax":"Florida sales tax ($)", 
                 "Discretionary sales surtax":"Discretionary sales surtax ($)", "Taxes and charges":"Taxes and charges ($)", "Gross receipts tax": "Gross receipts tax ($)",
                 "Regulatory fee":"Regulatory fee ($)", "County sales tax":"County sales tax ($)","Total Services and Tax":"Total Services and Tax ($)",
                  "Total Charge":"Total Charge ($)","Total amount per Qtr":"Total amount per Qtr ($)"}
            
            # Use the rename method to change row names in the DataFrame
            consolidate_transpose_df = consolidate_transpose_df.rename(index=rename_dict)
             # Save the consolidated transpose data to the 'Consolidate Transpose' sheet
            #consolidate_transpose_df = consolidate_transpose_df.shift(periods=2, axis=1).shift(periods=3, axis=0)
            consolidate_transpose_df.to_excel(excel_writer, sheet_name='Consolidated')
            
            
            E_on_p1=consolidate_transpose_df.loc["Energy $/kwh on peak"]
            E_on_p=sum(E_on_p1)/12
            
            E_off_p1=consolidate_transpose_df.loc["Energy $/kwh off peak"]
            E_off_p=sum(E_off_p1)/12
            
            
            F_on_p1=consolidate_transpose_df.loc["Fuel Charge $/kwh on peak"]
            F_on_p=sum(F_on_p1)/12
            
            F_off_p1=consolidate_transpose_df.loc["Fuel Charge $/kwh off peak"]
            F_off_p=sum(F_off_p1)/12
            
            
            E_F_on_p=E_on_p +F_on_p
            E_F_off_p=E_off_p + F_off_p
            
            
            
           
            
            
            ####################################################################################################################
            total_demand_kw_row = consolidate_transpose_df.loc['Total Demand kw']
            total_demand_kw_12_months = total_demand_kw_row.iloc[:12]  
            ####################################################################################################################
            Total_Demand_Charge1= consolidate_transpose_df.loc['Total Demand Charge ($)']
            Total_Demand_Charge=Total_Demand_Charge1.iloc[:12]
            ####################################################################################################################
            months111 = consolidate_transpose_df.columns[:12] 
            ####################################################################################################################
            ##&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&////
            ####################################################################################################################
            Total_Energy_Fuel_Charge1= consolidate_transpose_df.loc['Total Energy & Fuel Charge ($)']
            Total_Energy_Fuel_Charge=Total_Energy_Fuel_Charge1.iloc[:12]
            ####################################################################################################################
            Demand_Rate1=consolidate_transpose_df.loc['Demand Rate']
            Demand_Rate=round(Demand_Rate1.loc['Sum'], 3)
            
            ####################################################################################################################
            Electricity_Rate1=consolidate_transpose_df.loc['Average $/kwh cost (Exc fees)']
            Electricity_Rate=round(Electricity_Rate1.loc['Sum'],3)
            ####################################################################################################################
            
            Late_payment_charge1=consolidate_transpose_df.loc['Late payment charge ($)']
            Late_payment_charge=round(Late_payment_charge1.loc['Sum'],3)
            ####################################################################################################################
            Consumption_on_Peak_kwh1= consolidate_transpose_df.loc['Consumption On Peak kwh']
            Consumption_on_Peak_kwh=Consumption_on_Peak_kwh1.iloc[:12]
            
            ####################################################################################################################
            Consumption_off_Peak_kwh1= consolidate_transpose_df.loc['Consumption off-Peak kwh']
            Consumption_off_Peak_kwh=Consumption_off_Peak_kwh1.iloc[:12]
            
            ####################################################################################################################
            Energy_charge_off_peak1= consolidate_transpose_df.loc['Energy $/kwh off peak'] 
            
            Energy_charge_off_peak_10=sum(Energy_charge_off_peak1.iloc[:12])/12*sum(Consumption_off_Peak_kwh)*0.1
            
            ####################################################################################################################
            Energy_charge_on_peak_1= consolidate_transpose_df.loc['Energy $/kwh on peak'] 
            print("ujujujujuj")
            Energy_charge_on_peak_10=sum(Energy_charge_on_peak_1.iloc[:12])/12*sum(Consumption_on_Peak_kwh)*0.1
            print("klklakaljsjdhfhfhfhfh")
            ####################################################################################################################
            Fuel_charge_on_peak_1= consolidate_transpose_df.loc['Fuel Charge $/kwh on peak'] 
            Fuel_charge_on_peak_10=sum(Fuel_charge_on_peak_1.iloc[:12])/12*sum(Consumption_on_Peak_kwh)*0.1

            ####################################################################################################################
            Fuel_charge_off_peak_1= consolidate_transpose_df.loc['Fuel Charge $/kwh off peak'] 
            Fuel_charge_off_peak_10=sum(Fuel_charge_off_peak_1.iloc[:12])/12*sum(Consumption_off_Peak_kwh)*0.1
            ####################################################################################################################
            Recommendation_move_on_off=(Energy_charge_on_peak_10+Fuel_charge_on_peak_10)+(Energy_charge_off_peak_10+Fuel_charge_off_peak_10)
            
            
            
            
            
            ###########AR
            ########################################################################
            implementation_cost_late=0
            total_annual_saving_late=consolidate_transpose_df.at["Late payment charge ($)", "Sum"] 
            simple_payback_late='Immediate'
            
            if total_annual_saving_late != 0:
                applicable="Applicable"
            else:
             applicable="Not Applicable"
              
            nan_value='#'
            ########################################################################
            # Find the value just before the maximum value/load factor
            
            A=consolidate_transpose_df.loc["Total Demand kw"].max()
            values_before_max = consolidate_transpose_df.loc["Total Demand kw"][consolidate_transpose_df.loc["Total Demand kw"] < A]
            # Find the maximum value from the values before the max
            value_before_max = values_before_max.max()
            B=consolidate_transpose_df.loc["Total Demand kw","Sum"]/12
            C=consolidate_transpose_df.loc[ "Demand Rate","Sum"]
            implementation_cost_load_factor=0
            AA= float(value_before_max)
            B = float(B)
            C = float(C)
            
            total_annual_saving_load_factor= (AA-B)*C
            simple_payback_load_factor='Immediate'
            ###########################################################################
            parameter=1.4
            Max_Demand_Expectation=parameter*consolidate_transpose_df.loc["Total Comsuption kwh","Sum"]/(12*consolidate_transpose_df.loc["Service days","Sum"])
            
            implementation_cost_Max_Demand_Expectation=0
            DDD=consolidate_transpose_df.loc["Total Demand kw"].min()
            if B> Max_Demand_Expectation:
               total_annual_saving_Max_Demand_Expectation=(B- Max_Demand_Expectation)*C 
            else :
                total_annual_saving_Max_Demand_Expectation=0
            simple_payback_Max_Demand_Expectation='Immediate'
            
            ###########################################################################
            
            Toatal_cost_TOU1=consolidate_transpose_df.loc["Energy Charge On peak ($)","Sum"]+consolidate_transpose_df.loc["Energy Charge Off peak ($)","Sum"]
            Toatal_cost_TOU2=consolidate_transpose_df.loc["Fuel Charge on peak ($)","Sum"]+consolidate_transpose_df.loc["Fuel Charge off peak ($)","Sum"]
            Total_demand_charge_TOU=consolidate_transpose_df.loc[ "Total Demand Charge TOU ($)","Sum"]
            
            Toatal_cost_TOU_GSDT= Toatal_cost_TOU1+Toatal_cost_TOU2+Total_demand_charge_TOU
            
            
            Total_comsuption_TOU=consolidate_transpose_df.loc["Consumption On Peak kwh","Sum"]+consolidate_transpose_df.loc["Consumption off-Peak kwh","Sum"]
            Total_rate_Fuel_Energy=consolidate_transpose_df.loc["Fuel $/KWH Non-TOU","Sum"]+consolidate_transpose_df.loc["Energy $/kwh on Non-TOU","Sum"]
            #print(consolidate_transpose_df)
            #print(Total_rate_Fuel_Energy)
            tolerance = 1e-6
            # Define the variable
            Total_rate_Fuel_Energy = np.nan  # Replace with your variable

            # Check if the variable is NaN
            if np.isnan(Total_rate_Fuel_Energy):
              Total_rate_Fuel_Energy=0.035
            Total_cost_comsuption_FE=Total_comsuption_TOU*Total_rate_Fuel_Energy
            Demand_rate_GSD=consolidate_transpose_df.loc["Demand_$/kwh- Non TOU","Sum"]
            #print(Demand_rate_GSD)
            if Demand_rate_GSD <tolerance:
               Demand_rate_GSD=11.25
            Total_cost_Demand_GSD = consolidate_transpose_df.loc["Maximum Demand kw","Sum"]*Demand_rate_GSD
            Total_cost_NON_TOU_GSD=Total_cost_comsuption_FE+Total_cost_Demand_GSD
            
            total_Annual_Saving_to_GSD=Toatal_cost_TOU_GSDT-Total_cost_NON_TOU_GSD
            implementation_cost_change_rate=0
            simple_payback_change_rate='Immediate'
            
            # Create a DataFrame with rows and calculated values for the "Summary1" sheet
            data1 = {
                "Category": ["Late payment fees were discovered upon examination of the electrical bills",
                             "Implementation Cost of Late",
                            "Total Annual Saving of Late",
                            "Simple Payback of Late",applicable
                            ],
                "$ Value": [nan_value,implementation_cost_late,
                        total_annual_saving_late,
                        simple_payback_late,nan_value
                        ]
            }
            
            
            # Create a DataFrame with rows and calculated values for the "Summary1" sheet
            data2 = {
                "Category": ["Load Factor: We can Reduce Max Demand and saving Cost",
                            "Implementation Cost of Load Factor",
                            "Total Annual Saving of Load Factor",
                            "Estimated Demand saving = (Max_Demand - Average_Demand)*Demand_Rate",applicable,
                            "Simple Payback of Load Factor"],
                "$ Value": [nan_value,implementation_cost_late,
                        total_annual_saving_load_factor,nan_value,
                        simple_payback_late,nan_value]
            }
            
            
            
            nan_value1="#"
            
            # Create a DataFrame with rows and calculated values for the "Summary1" sheet
            data3 = {
                "Category": ["Implementation Cost of Max Demand","total_annual_saving_Max_Demand_Expectation=(Average_Demand- 1.4*Max_Demand_Expectation)*Demand Rate",
                            "Total Annual Saving of Max Demand",
                            "Simple Payback of Max Demand",applicable],
                "$ Value": [
                        implementation_cost_Max_Demand_Expectation,nan_value1,
                        total_annual_saving_Max_Demand_Expectation,
                        simple_payback_Max_Demand_Expectation,nan_value]
            }
            
            
            
            
            # Create a DataFrame with rows and calculated values for the "Summary1" sheet
            data4 = {
                "Category": ["In order to switch to the GSD-1 rate","total_Annual_Saving_to_GSD=Toatal_cost_TOU_GSDT-Total_cost_NON_TOU_GSD",
                            "Implementation Cost of Change Rate",
                            "Total Annual Saving Change Rate",
                            "Simple Payback of Change Rate",applicable],
                "$ Value": [nan_value1,nan_value1,
                        implementation_cost_change_rate,
                       total_Annual_Saving_to_GSD,
                        simple_payback_change_rate,nan_value]
            }
            
            

            ad_model1=pd.DataFrame(data1)
            ad_model2=pd.DataFrame(data2)
            ad_model3=pd.DataFrame(data3)
            ad_model4=pd.DataFrame(data4)
            
            
            ad_model1.to_excel(excel_writer, sheet_name='Pay Electrical Bills On Time')
            ad_model2.to_excel(excel_writer, sheet_name='Load Factor')
            ad_model3.to_excel(excel_writer, sheet_name='Expectation of Max Demand')
            ad_model4.to_excel(excel_writer, sheet_name='Change Rate Structure to GSD')

             # Get the Excel writer's workbook and worksheet objects
            workbook = excel_writer.book
            worksheet = excel_writer.sheets['Consolidated']
            worksheet1 = excel_writer.sheets['Pay Electrical Bills On Time']
            worksheet2 = excel_writer.sheets['Load Factor']
            worksheet3 = excel_writer.sheets['Expectation of Max Demand']
            worksheet4 = excel_writer.sheets['Change Rate Structure to GSD']
         
            

             # Define the background color (e.g., green)
            green_fill = workbook.add_format({'bg_color': '00FF00'})

            # Define the row names to highlight
            rows_to_highlight = ["Average Energy $/kWh","Average Fuel $/kWh","Total Energy Charge ($)", "Total Demand Charge ($)",
                                 "Total Energy & Fuel Charge ($)","Total Fuel Charge ($)","Total $/kwh cost","Total Demand kw","Total Charge ($)","Total Services and Tax ($)", "Total Comsuption kwh","Total Electric cost ($)"]
            
            blue_fill = workbook.add_format({'bg_color': '#B4C6E7'})

                      
            consolidate_transpose_df = consolidate_transpose_df.fillna(0)


            # Iterate through rows and apply the background color
            for row_num, row_name in enumerate(consolidate_transpose_df.index):
                if row_name in rows_to_highlight:
                    for col_num in range(1, 14):  # Include column 14 for 'Sum'
                        cell_value = consolidate_transpose_df.iloc[row_num, col_num - 1]  # Subtract 1 to get the correct column index
                        worksheet.write(row_num + 1, col_num, cell_value, green_fill)  # Add +1 because row_num is zero-based
                        
                        
            rows_to_light = [ "Service days","Consumption On Peak kwh","Consumption off-Peak kwh","Non-TOU Consumption kwh","Energy Charge On peak ($)","Energy Charge Off peak ($)",
                             "Energy Charge Non-TOU ($)"," Energy $/kwh on peak"," Energy $/kwh off peak","Energy $/kwh on Non-TOU","Fuel Charge on peak ($)",
                             "Fuel Charge off peak ($)","Fuel Charge Non-TOU ($)","Fuel Charge $/kwh on peak","Fuel Charge $/kwh off peak","Fuel $/KWH Non-TOU","Total Demand kw - Non TOU",
                             "Contract demand kw","On-peak demand kw","Maximum Demand kw","Demand_$/kwh- Non TOU","On-peak Demand $/kwh","Maximum Demand $/kwh",
                             "Total Demand Charge - Non TOU ($)","Total Demand Charge TOU ($)","Demand $/kwh","Base charge($)","Service Charge ($)","Late payment charge ($)",
                             "Gross rec. tax/Regulatory fee ($)","Franchise charge ($)","Franchise fee ($)","Utility tax ($)","Florida sales tax ($)","Discretionary sales surtax ($)",
                             "Gross receipts tax ($)","Regulatory fee ($)","County sales tax ($)","Total amount per Qtr ($)","Average $/kwh cost (Exc fees)","Demand Rate"]
                
                        
        # Iterate through rows and apply the background color
            for row_num, row_name in enumerate(consolidate_transpose_df.index):
                if row_name in rows_to_light:
                    for col_num in range(1, 14):  # Include column 14 for 'Sum'
                        cell_value = consolidate_transpose_df.iloc[row_num, col_num - 1]  # Subtract 1 to get the correct column index
                        worksheet.write(row_num + 1, col_num, cell_value, blue_fill)  # Add +1 because row_num is zero-based
                        
           # Define the background color (e.g., light yellow)
            light_yellow_fill = workbook.add_format({'bg_color': '#FFFF00'})  
            rows_to_light1 = [ "Gross rec. tax/Regulatory fee ($)","Franchise charge ($)","Franchise fee ($)","Utility tax ($)","Florida sales tax ($)","Discretionary sales surtax ($)",
                             "Gross receipts tax ($)","Regulatory fee ($)","County sales tax ($)"]
           # Iterate through rows and apply the background color
            for row_num, row_name in enumerate(consolidate_transpose_df.index):
                if row_name in rows_to_light1:
                    for col_num in range(1, 14):  # Include column 14 for 'Sum'
                        cell_value = consolidate_transpose_df.iloc[row_num, col_num - 1]  # Subtract 1 to get the correct column index
                        worksheet.write(row_num + 1, col_num, cell_value, light_yellow_fill)  # Add +1 because row_num is zero-based
                        
           # Define the background color (e.g., red)
            red_fill = workbook.add_format({'bg_color': '#FF0000'})     
            rows_to_light2 = [ "Late payment charge ($)"]
            # Iterate through rows and apply the background color
            for row_num, row_name in enumerate(consolidate_transpose_df.index):
                if row_name in rows_to_light2:
                    for col_num in range(1, 14):  # Include column 14 for 'Sum'
                        cell_value = consolidate_transpose_df.iloc[row_num, col_num - 1]  # Subtract 1 to get the correct column index
                        worksheet.write(row_num + 1, col_num, cell_value, red_fill)  # Add +1 because row_num is zero-based  
                                          
           # Create a format for text justification (e.g., left alignment)
            justify_format = workbook.add_format({'align': 'left'})

            # Apply the text justification format to specific columns (e.g., columns A to Z)
            worksheet.set_column('A:Z', None, justify_format)

            # Adjust the width of specific columns (e.g., columns A to Z) to your preferred width
            worksheet.set_column('A:A', 50)  # Adjust '15' to your preferred width
            
            # Create a format for text justification (center alignment)
            center_align_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})

            # Apply the center alignment format to specific columns (e.g., columns A to Z)
            worksheet.set_column('B:Z', None, center_align_format)
           
            
            
            # Apply the text justification format to specific columns (e.g., columns A to Z)
            worksheet1.set_column('A:Z', None, justify_format)
            # Adjust the width of specific columns (e.g., columns A to Z) to your preferred width
            worksheet1.set_column('B:B', 90)  # Adjust '15' to your preferred width
            worksheet1.set_column('C:C', 50)  # Adjust '15' to your preferred width
             # Create a format for text justification (center alignment)
            center_align_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})

            # Apply the center alignment format to specific columns (e.g., columns A to Z)
            worksheet1.set_column('C:Z', None, center_align_format)
            
            

                        # Assuming total_annual_saving_late is your condition
            if total_annual_saving_late != 0:
                bold_format = workbook.add_format({'font_color': 'green', 'bold': True})
                worksheet1.write(5, 1, 'Applicable', bold_format)  # Green bold text for column B
            else:
            
                bold_format = workbook.add_format({'font_color': 'red', 'bold': True})
                worksheet1.write(5, 1, 'Not Applicable', bold_format)  # Green bold text for column B    
                # Column C for row 4
                #worksheet1.write(5, 2, 'Applicable', workbook.add_format({'font_color': 'green'}))  # Green text for column C    
               
            # Apply the text justification format to specific columns (e.g., columns A to Z)
            worksheet2.set_column('A:Z', None, justify_format)
            # Adjust the width of specific columns (e.g., columns A to Z) to your preferred width
            worksheet2.set_column('B:B', 90)  # Adjust '15' to your preferred width
            worksheet2.set_column('C:C', 50)  # Adjust '15' to your preferred width
             # Create a format for text justification (center alignment)
            center_align_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})

            # Apply the center alignment format to specific columns (e.g., columns A to Z)
            worksheet2.set_column('C:Z', None, center_align_format)
            
            
                        # Assuming total_annual_saving_late is your condition
            if total_annual_saving_load_factor > 0:
                bold_format = workbook.add_format({'font_color': 'green', 'bold': True})
                worksheet2.write(6, 1, 'Applicable', bold_format)  # Green bold text for column B
            else:
            
                bold_format = workbook.add_format({'font_color': 'red', 'bold': True})
                worksheet2.write(6, 1, 'Not Applicable', bold_format)  # Green bold text for column B   
            
            
            # Apply the text justification format to specific columns (e.g., columns A to Z)
            worksheet3.set_column('A:Z', None, justify_format)
            # Adjust the width of specific columns (e.g., columns A to Z) to your preferred width
            worksheet3.set_column('B:B', 130)  # Adjust '15' to your preferred width
            worksheet3.set_column('C:C', 50)  # Adjust '15' to your preferred width
             # Create a format for text justification (center alignment)
            center_align_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})

            # Apply the center alignment format to specific columns (e.g., columns A to Z)
            worksheet3.set_column('C:Z', None, center_align_format)
            
            
                         # Assuming total_annual_saving_late is your condition
            if total_annual_saving_Max_Demand_Expectation > 0:
                bold_format = workbook.add_format({'font_color': 'green', 'bold': True})
                worksheet3.write(5, 1, 'Applicable', bold_format)  # Green bold text for column B
            else:
            
                bold_format = workbook.add_format({'font_color': 'red', 'bold': True})
                worksheet3.write(5, 1, 'Not Applicable', bold_format)  # Green bold text for column B   
            
            # Apply the text justification format to specific columns (e.g., columns A to Z)
            worksheet4.set_column('A:Z', None, justify_format)
            # Apply the center alignment format to specific columns (e.g., columns A to Z)
            
            # Adjust the width of specific columns (e.g., columns A to Z) to your preferred width
            worksheet4.set_column('B:B', 90)  # Adjust '15' to your preferred width
            worksheet4.set_column('C:C', 50)  # Adjust '15' to your preferred width
            # Create a format for text justification (center alignment)
            center_align_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
            worksheet4.set_column('C:Z', None, center_align_format)
            
             # Assuming total_annual_saving_late is your condition
            if total_Annual_Saving_to_GSD > 0:
                bold_format = workbook.add_format({'font_color': 'green', 'bold': True})
                worksheet4.write(6, 1, 'Applicable', bold_format)  # Green bold text for column B
            else:
            
                bold_format = workbook.add_format({'font_color': 'red', 'bold': True})
                worksheet4.write(6, 1, 'Not Applicable', bold_format)  # Green bold text for column B  

            df_account.to_excel(excel_writer, sheet_name=account_sheet_name, index=False)
            
           
         
       # Save 'result_df' to the 'Consolidate' sheet
        #result_df.to_excel(excel_writer, sheet_name='Consolidate')
        for i, df_account_transposed in enumerate(data_by_account_transposed):
            account_sheet_name = f'Account_{i + 1}_Transposed'
            
            
            
            df_account_transposed.to_excel(excel_writer, sheet_name=account_sheet_name)
        

    print(f"Data saved to {excel_filename}")
    Total_Comsuption_kwh1= consolidate_transpose_df.loc['Total Comsuption kwh']
    Total_Comsuption_kwh=Total_Comsuption_kwh1.iloc[:12]
    print("emroooooooooooooooooooooooooooooooooooooooooo")
    print(Total_Comsuption_kwh1)

    # Ensure that no implicit boolean values are being evaluated or returned
    if isinstance(extracted_data, bool):
        extracted_data = None  # Set to None if a boolean is encountered unexpectedly
    return excel_filename, seasonal_consumption, accounts_usage[account_number],total_demand_kw_12_months,Total_Comsuption_kwh,months111,Demand_Rate,Electricity_Rate, Total_Demand_Charge,Total_Energy_Fuel_Charge,Late_payment_charge,Rate_account[account_number],Consumption_on_Peak_kwh,Consumption_off_Peak_kwh,Recommendation_move_on_off,E_F_on_p,E_F_off_p
    


def app():
    # Inject CSS styling for customization
    st.markdown("""
        <style>
        /* Background color */
        .stApp {
            background-color: #f7f7f7;
            font-family: 'Arial', sans-serif;
        }

        /* Customize headers */
        .stApp h1, .stApp h2, .stApp h3, .stApp h4 {
            color: #004080;
            font-weight: bold;
        }

        /* Customize input fields */
        .stNumberInput, .stRadio {
            background-color: #ffffff;
            border: 1px solid #cccccc;
            border-radius: 5px;
            padding: 5px;
        }

        /* Style radio buttons */
        .stRadio > label {
            color: #004080;
        }
        .stRadio input[type="radio"] {
            accent-color: #ff6347;
        }

        /* Customize the file uploader */
        .stFileUploader {
            background-color: #eaf2f8;
            border-radius: 5px;
            border: 1px solid #cccccc;
            padding: 10px;
            color: #004080;
        }

        /* Style the button */
        .stButton button {
            background-color: #ff6347;
            color: white;
            border: none;
            padding: 10px 20px;
            border-radius: 5px;
            font-weight: bold;
        }
        .stButton button:hover {
            background-color: #ff4500;
        }

        /* Darker and thicker table lines */
        .stDataFrame, .stTable {
            border: 10px solid #004080;
            border-radius: 10px;
        }
        
        /* Table text alignment and padding */
        .stTable th, .stTable td {
            text-align: center;
            padding: 10px;
            border: 10px solid #004080;
        }

        /* Justify table content and set font size */
        .stTable td {
            text-align: justify;
            font-size: 14px;
            color: #333333;
        }

        /* Make table headers bold and darker */
        .stTable th {
            background-color: #004080;
            color: #ffffff;
            font-weight: bold;
            text-align: center;
        }

        /* Table styling for DataFrame table */
        .stDataFrame .row_heading, .stDataFrame .blank {
            display: none;  /* Remove row numbers and blank cells */
        }
        </style>
    """, unsafe_allow_html=True)
    #st.image(r"c:\Users\mxz881\Desktop\OIP.jpeg", use_column_width=True, caption="hello")
    #st.image("C:/Users/mxz881/Desktop/logo.jpg", use_column_width=True, caption="hello")
    # image_path = r"C:\Users\mxz881\Desktop\logo1.jpg"
    # st.write("Image path exists:", os.path.exists(image_path))
    # if os.path.exists(image_path):
    #     st.image(image_path, use_column_width=True, caption="Logo")
    # else:
    #     st.write("Error: Image path does not exist. Check the path.")
    #st.image("static/logo.jpg", use_column_width=True, caption="Logo")
    import base64

    def get_base64_image(path):
        with open(path, "rb") as file:
            data = file.read()
        return base64.b64encode(data).decode()

    img_base64 = get_base64_image("Logo-University-of-Miami.jpg")
    html_code = f'<img src="data:image/jpeg;base64,{img_base64}" style="width:20%;">'
    st.markdown(html_code, unsafe_allow_html=True)

    st.title("Extractor Electricity Bills (Commercial and Industrial Building)")

    # Input for number of accounts
    num_accounts = st.number_input("Enter the number of accounts (folders):", min_value=1, step=1)
    # Radio button for window type
    window_type = st.radio(
        "Do you use regular windows or impact windows? (Select 'Yes' for regular windows, 'No' for impact windows)",
        options=["Yes", "No"]
    )

    # Convert the selection to 1 for 'Yes' and 0 for 'No' if needed
    if window_type == "Yes":
        window_value = 1
    else:
        window_value = 0
        # Define temperature coefficients for each month directly in the code
    coefficients = [0.5904, 0.6416, 0.6672, 0.6672, 0.6928, 0.7184, 0.7952, 0.8464, 1.0, 0.7696, 0.744, 0.6672]
    # Ask for working hours and working days as inputs
   # Using st.text_input for direct entry
    #working_hours = int(st.text_input("Enter the working hours per day:", value="8"))
    #working_days = int(st.text_input("Enter the working days per week:", value="5"))


    # Calculate operation hours
    #operation_hours = working_hours * working_days * 52  # Assuming 52 weeks in a year
    
    st.markdown("### Working Hours")

    # Define shifts and day types
    shifts = ["1st Shift", "2nd Shift", "3rd Shift"]
    days = ["Monday - Friday", "Saturday", "Sunday"]

    # Initialize session state for shifts and days
    if "working_hours_data" not in st.session_state:
        st.session_state["working_hours_data"] = {
            "Shift": [],
            "Day Type": [],
            "From": [],
            "To": []
        }
        # Add initial values for all shifts and days
        for shift in shifts:
            for day in days:
                st.session_state[f"{shift}_{day}_start"] = dt.time(6, 30)  # Default start time
                st.session_state[f"{shift}_{day}_end"] = dt.time(15, 0)  # Default end time

    # Ask if Saturday and Sunday are operating days
    saturday_operating = st.checkbox("Saturday an operating day", value=True)
    sunday_operating = st.checkbox("Sunday an operating day", value=True)

    # Ask if each shift is operating
    shift_checks = {}
    for shift in shifts:
        shift_checks[shift] = st.checkbox(f"{shift} an operating shift", value=True)

    # Create a dictionary to store inputs
    working_hours_data = {
        "Shift": [],
        "Day Type": [],
        "From": [],
        "To": []
    }

    # Loop through each shift and day type to create input fields
    for shift in shifts:
        if not shift_checks[shift]:
            continue  # Skip shifts that are not checked

        for day in days:
            if day == "Saturday" and not saturday_operating:
                continue  # Skip Saturday if it's not an operating day
            if day == "Sunday" and not sunday_operating:
                continue  # Skip Sunday if it's not an operating day

            col1, col2, col3 = st.columns([1, 2, 2])
            with col1:
                st.write(f"{shift} - {day}")
            with col2:
                start_time = st.time_input(
                    f"Start time ({shift} - {day})",
                    value=st.session_state[f"{shift}_{day}_start"],
                    key=f"{shift}_{day}_start_input"
                )
            with col3:
                end_time = st.time_input(
                    f"End time ({shift} - {day})",
                    value=st.session_state[f"{shift}_{day}_end"],
                    key=f"{shift}_{day}_end_input"
                )
            # Update session state with user inputs
            st.session_state[f"{shift}_{day}_start"] = start_time
            st.session_state[f"{shift}_{day}_end"] = end_time

            # Append the data (ensure day data is added to the dictionary)
            working_hours_data["Shift"].append(shift)
            working_hours_data["Day Type"].append(day)
            working_hours_data["From"].append(str(start_time))
            working_hours_data["To"].append(str(end_time))

    # Create a DataFrame from the collected inputs
    working_hours_df = pd.DataFrame(working_hours_data)

    # Function to calculate hours
    def calculate_hours(start_time_str, end_time_str): 
        """
        Calculate the duration in hours between start and end times.
        Handles overnight shifts.
        """
        # Adjust format to handle potential seconds
        if ":" in start_time_str and start_time_str.count(":") == 2:
            time_format = "%H:%M:%S"  # Format with hours, minutes, and seconds
        else:
            time_format = "%H:%M"  # Format with hours and minutes

        # Convert the string inputs to datetime.time objects
        start_time = datetime.strptime(start_time_str, time_format).time()
        end_time = datetime.strptime(end_time_str, time_format).time()
        
        # Combine with today's date
        start = datetime.combine(datetime.today(), start_time)
        end = datetime.combine(datetime.today(), end_time)
        
        # Adjust for overnight shifts
        if end <= start:
            end += timedelta(days=1)
        
        # Calculate the difference in hours
        return (end - start).total_seconds() / 3600

    # Calculate total operation hours
    def calculate_total_operation_hours(working_hours_df):
        """
        Calculate total operation hours per year based on the shifts and day types.
        """
        total_operation_hours = 0
        for _, row in working_hours_df.iterrows():
            start_time = row["From"]
            end_time = row["To"]
            day_type = row["Day Type"]
            
            # Skip rows with invalid times (e.g., 00:00 to 00:00)
            if start_time == "00:00:00" and end_time == "00:00:00":
                continue
            
            # Calculate hours for each shift
            daily_hours = calculate_hours(start_time, end_time)
            
            # Multiply hours based on day type
            if day_type == "Monday - Friday":
                total_operation_hours += daily_hours * 5  # 5 weekdays
            elif day_type == "Saturday":
                total_operation_hours += daily_hours * 1  # 1 Saturday
            elif day_type == "Sunday":
                total_operation_hours += daily_hours * 1  # 1 Sunday

        # Multiply by 52 weeks in a year
        return total_operation_hours * 52

    # Display collected working hours
    st.markdown("### Collected Working Hours")
    working_hours_df.index = working_hours_df.index + 1
    st.dataframe(working_hours_df)

    # Calculate and display total operation hours
    total_operation_hours = calculate_total_operation_hours(working_hours_df)
    st.markdown(f"### Total Operation Hours (per year): {total_operation_hours:.2f} hours")
    operation_hours=total_operation_hours
    
    
    
    # Dictionary to store the uploaded files for each account (folder)
    uploaded_files_by_account = {}

    # Loop for uploading files for each account
    for account_number in range(1, num_accounts + 1):
        uploaded_files = st.file_uploader(f"Upload PDFs for Account {account_number}", type="pdf", accept_multiple_files=True, key=f"account_{account_number}")
        if uploaded_files:
            uploaded_files_by_account[account_number] = uploaded_files

    # Placeholder for clearing previous outputs
    progress_placeholder = st.empty()

    if st.button("Extract Data"):
        if uploaded_files_by_account and num_accounts > 0:
            

            # Initialize variables for tracking seasonal consumption
            seasonal_consumption_all_accounts = {}
            consolidated_seasonal_consumption = {}
            accounts_usage = {}  # Dictionary to store usage per account
            Rate_account={}
            #for account_number, uploaded_files in uploaded_files_by_account.items():
                # Clear previous outputs from the placeholder
               # progress_placeholder.empty()

                
                # Add error handling around the data extraction
            try:
                    st.write(f"Extracting data for Account {account_number}")
                    excel_filename, seasonal_consumption, accounts_usage[account_number],total_demand_kw_12_months,Total_Comsuption_kwh,months111,Demand_Rate,Electricity_Rate, Total_Demand_Charge,Total_Energy_Fuel_Charge,Late_payment_charge,Rate_account[account_number],Consumption_on_Peak_kwh,Consumption_off_Peak_kwh,Recommendation_move_on_off,E_F_on_p,E_F_off_p = extract_and_consolidate_data(uploaded_files_by_account, num_accounts, coefficients)
                    
                   
                    
                    
                    
                    # Explicitly check if there's an issue with how data is being returned or stored
                    assert isinstance(seasonal_consumption, dict), f"Error: Unexpected type for seasonal_consumption: {type(seasonal_consumption)}"
                    
            except Exception as e:
                    st.error(f"Error in data extraction for Account {account_number}: {e}")
            #continue

                # Store seasonal consumption if it's valid
            if seasonal_consumption:
                    seasonal_consumption_all_accounts[account_number] = seasonal_consumption

                    # Consolidate seasonal consumption across accounts
                    for season, consumption in seasonal_consumption.items():
                        consolidated_seasonal_consumption[season] = consolidated_seasonal_consumption.get(season, 0) + consumption

           
                    # Provide download button for the Excel file
                    with open(excel_filename, "rb") as file:
                        st.download_button(label=f"Download Excel File for Account {account_number}",
                                           data=file,
                                           file_name=excel_filename,
                                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            
               
            
            
            #st.write(f"Demand Rate: {Demand_Rate}")
            #st.write(f"Electricity Rate: {Electricity_Rate}")
            # Assuming you've already calculated Demand_Rate and Electricity_Rate
            # Now build a table that includes rates for all accounts
            account_numbers = list(Rate_account.keys())
            account_numbers_str = [str(num) for num in account_numbers]

            rates = list(Rate_account.values())
            #data = {
             #   "Rate Type": ["Demand Rate", "Electricity Rate","Late payment charge ($)","Account Number": account_numbers],
             #   "Value": [Demand_Rate, Electricity_Rate,Late_payment_charge,"Rate (for each account)": rates]
            #}
            
           # Create a dictionary to store the rate for each account
            account_rates = {
                f"Account Number ({account_number})": rate for account_number, rate in Rate_account.items()
            }
            
            # Prepare the first table data for rates
            data_rates = {
                "Name": list(account_rates.keys()) +["Demand Rate(Consolidate) ($/kW)", "Electricity Rate(Consolidate) ($/kWh)"] ,
                "Information": list(account_rates.values())+ [Demand_Rate, Electricity_Rate] 
            }

            # Convert the data into a DataFrame for the first table
            rates_df = pd.DataFrame(data_rates)
            rates_df["Information"] = rates_df["Information"].astype(str)  # Convert values to string for display

            # Display the first table (Rates Information)
            #st.write("Rates Information:")
            
            st.markdown("**Account Information:**")
            rates_df.index=rates_df.index +1
            st.table(rates_df)

            if Late_payment_charge > 0:
                annual_saving_late = Late_payment_charge   # Assuming the charge is monthly
                recommendation_value = f" $ {annual_saving_late:.2f}"
                Late_payback="Immediate"
            else:
                recommendation_value = "Not Applicable"
                Late_payback="NA"
           
            # operation_hours=3120    
            Max_Demand = Total_Comsuption_kwh.sum() /  operation_hours
            # Logic for recommendation based on max(Total Demand) and max_Demand
            max_total_demand = max(total_demand_kw_12_months)  # Calculate the maximum total demand
             # Calculate the min and max of total_demand_kw_12_months
            min_demand = min(total_demand_kw_12_months)
            max_demand = max(total_demand_kw_12_months)
            average_demand = sum(total_demand_kw_12_months) / len(total_demand_kw_12_months)

            #if max_total_demand > Max_Demand:
            #    max_demand = max(total_demand_kw_12_months)
            #    annual_saving_Demand= (max_demand -Max_Demand)*Demand_Rate
            #    annual_saving_Demand_kw=round(max_demand -Max_Demand)
            #    demand_recommendation = f"$ {annual_saving_Demand:.2f}"
            #    Demand_payback="Immediate"
            Energy_Reduction_demand_kWh = "NA"  # Default value

            if max_demand > average_demand:
                #annual_saving_Demand= (max_demand -average_demand)*Demand_Rate
                
                annual_saving_Demand= sum(total_demand_kw_12_months*0.05)*Demand_Rate
                #annual_saving_Demand_kw=round(max_demand -average_demand)
                #st.write(total_demand_kw_12_months*0.05)
                annual_saving_Demand_kw=round(sum(total_demand_kw_12_months*0.05))
                Energy_Reduction_demand_kWh=annual_saving_Demand_kw
                annual_saving_Demand_kwh=annual_saving_Demand*Electricity_Rate
                demand_recommendation = f"$ {annual_saving_Demand+annual_saving_Demand_kwh:.2f}"
                Demand_payback="Immediate"
               
            else:
                demand_recommendation = "Not Applicable"
                Demand_payback="NA"
                annual_saving_Demand_kw="NA"
                
            if ('Consumption_on_Peak_kwh' in locals() and Consumption_on_Peak_kwh is not None and np.any(Consumption_on_Peak_kwh)) or \
                        ('Consumption_off_Peak_kwh' in locals() and Consumption_off_Peak_kwh is not None and np.any(Consumption_off_Peak_kwh)):
                annual_saving_Chage= 0.05* sum(Total_Energy_Fuel_Charge) 
                Change_Rate_Recomendation = f"$ {annual_saving_Chage:.2f}"
                Rate_paayback="Immediate"
               
                
                
                Recommendation_move_on_off_1= f" $ {Recommendation_move_on_off:.2f}"
                on_paayback="Immediate"
              
            else:
               Change_Rate_Recomendation = "Not Applicable"
               Recommendation_move_on_off_1="Not Applicable"
               Rate_paayback="NA"
               on_paayback="NA"
               
               
            if num_accounts>1:
                annual_saving_decrease_n_meter = total_demand_kw_12_months.sum()*Demand_Rate*0.1   # Assuming the charge is monthly
                meter_recommendation= f"$ {annual_saving_decrease_n_meter:.2f}"
                meter_payback="Immediate"
            else:
                meter_recommendation = "Not Applicable"
                meter_payback="NA"
                
              
            
            
            
            if window_value == 1:
                TOTAL_ENERGY_CONSUMPTION = Total_Comsuption_kwh.sum()  # Sum of 12 months total consumption
                # Calculate energy reduction due to potential tax eligibility
                energy_reduction = TOTAL_ENERGY_CONSUMPTION * 0.3*0.1  # 10% energy reduction
                annual_cost_energy_reduction = energy_reduction * Electricity_Rate  # cost saving
                
                
                if annual_cost_energy_reduction> 0:
                    
                    tax_payback="Immediate"
            else:
                recommendation_value = "Not Applicable"
                tax_payback="NA"
            
            not_available="NA"
            # Prepare the second table (Recommendation)
            data_recommendation = {
                 "No.": [1, 2, 3, 4, 5, 6],
                "Recommendation Type": [ "Late payment charge", "Max Demand Reduction","Change Rate Recommendation","Load shifting: 10% On Peak Consumption to Off peak ","Decrease number of Account","Annual Cost Savings from Energy Reduction from Tax Benefit 179D(1) or 10\% cooling reduction (by increasing the thermostat setting) "],
                "Value": [ recommendation_value, demand_recommendation,Change_Rate_Recomendation,Recommendation_move_on_off_1,meter_recommendation,f"${annual_cost_energy_reduction:.2f}"],
                "Pay Back":[Late_payback,Demand_payback,Rate_paayback,on_paayback,meter_payback,tax_payback],
                "kWh Saving":[not_available,Energy_Reduction_demand_kWh,"-" ,"-",not_available,energy_reduction],
                "kW Saving":[not_available,annual_saving_Demand_kw,"-" ,"-",not_available ,not_available]
            }

            
                            


            data_recommendation["No."] = list(range(1, len(data_recommendation["Recommendation Type"]) + 1))

            # Convert the recommendation data into a DataFrame
            recommendation_df = pd.DataFrame(data_recommendation)
            recommendation_df = recommendation_df[["Recommendation Type", "Value", "Pay Back", "kWh Saving", "kW Saving"]]

            recommendation_df["Value"] = recommendation_df["Value"].astype(str)  # Convert values to string for display
            #recommendation_df.reset_index(drop=True, inplace=True)

            # Display the second table (Recommendation)
            #st.write("Recommendation:")
            st.markdown("**Recommendation:**")

            #st.write("Recommendation number 4: It should be more deep analaysis but we calculated th")
            #st.table(recommendation_df.reset_index(drop=True))
            recommendation_df.index = recommendation_df.index + 1

            st.table(recommendation_df)
            
            #st.write("179D Commercial Buildings Energy-Efficiency Tax Deduction:")
            st.markdown("**Table Description of each Recommendation:**")
       
       
       
       
            

                #st.write("If you are eligible for the 179D Commercial Buildings Energy-Efficiency Tax Deduction and choose to upgrade your windows, you could significantly reduce your energy consumption and costs. By improving your building's windows, you can achieve a reduction of at least 10% in cooling consumption, which typically accounts for about 30% to 50% of your building's total energy consumption. This not only enhances energy efficiency but also contributes to lower utility bills, making it a financially beneficial investment. ")
            st.markdown("""
                <div style="text-align: justify;">
                    <p>1- It means that by paying on time, you can save the entire amount that would otherwise be spent on late payment fees over the year through this AR.</p>
                </div>
            """, unsafe_allow_html=True)

            st.markdown("""
                <div style="text-align: justify;">
                    <p>2- We determined the expected demand reduction by calculating a 5% reduction in demand for each month. Next, we calculated the demand cost and energy cost based on this reduction. Finally, we determined the total cost savings.</p>
                </div>
            """, unsafe_allow_html=True)

            # Add Table for 1
            st.table({"Description": ["Potential Annual Savings from 5 percent Demand Reduction"],
                    "Formula": ["sum(max deman of each month*0.05) * Demand Rate + sum(energy reduction for each month)* Energy Rate"]})

            st.markdown("""
                <div style="text-align: justify;">
                    <p>3- If your account is on a Time-of-Use plan, you have the potential for annual savings by switching to a Standard Rate.</p>
                </div>
            """, unsafe_allow_html=True)

            # Add Table for 2
            st.table({"Description": ["Savings from Change Rate"],
                    "Formula": ["(Present Rate - find another Rate) * Consumption_kWh ....."]})

            st.markdown("""
                <div style="text-align: justify;">
                    <p>4- If your account is on a Time-of-Use plan, you can reduce your bill by shifting usage from peak to off-peak hours. For example, we demonstrated an annual savings by reducing 10% of on-peak consumption and shifting it to off-peak hours.</p>
                </div>
            """, unsafe_allow_html=True)

            # Add Table for 3
            st.table({"Description": ["Savings by Shifting 10% On-Peak to Off-Peak"],
                    "Formula": ["(On_Peak_Rate - Off_Peak_Rate) * 10% On Peak Consumption_kWh"]})

            st.markdown("""
                <div style="text-align: justify;">
                    <p>5- Demand charges are typically based on the maximum demand (in kW) recorded at a single meter during peak usage times. For companies with multiple meters, each meter is assessed individually for its peak demand.</p>
                    <p>By consolidating meters, the total demand is measured across the entire facility or company, which can potentially reduce the overall peak demand. This occurs because different parts of the facility may reach their peak usage at different times, smoothing out the overall demand. As a result, the combined peak demand could be lower than the sum of the individual peaks. For example, we estimate that this consolidation could reduce the bill by approximately 10%.</p>
                </div>
            """, unsafe_allow_html=True)

            # Add Table for 4
            #st.table({"Description": ["Savings from Meter Consolidation"],
             #       "Formula": ["10% * Total_Energy_Charges"]})

            st.markdown("""
                <div style="text-align: justify;">
                    <p>6- If you are eligible for the 179D Commercial Buildings Energy-Efficiency Tax Deduction and choose to upgrade your windows, you could significantly reduce your energy consumption and costs. By improving your building's windows, you can achieve a reduction of at least 10% in cooling consumption, which typically accounts for about 30% to 50% of your building's total energy consumption. This not only enhances energy efficiency but also contributes to lower utility bills, making it a financially beneficial investment.</p>
                </div>
            """, unsafe_allow_html=True)

            # Add Table for 5
            st.table({"Description": ["Savings from 179D Tax Deduction"],
                    "Formula": ["10% * Cooling_Consumption_kWh * Electricity_Rate"]})

           

            
            # # If they use regular windows (1), calculate 10% energy reduction for tax savings
            # if window_value == 1:
            #     #st.write("179D Commercial Buildings Energy-Efficiency Tax Deduction:")
            #     st.markdown("**1- 179D Commercial Buildings Energy-Efficiency Tax Deduction:**")

            #     #st.write("If you are eligible for the 179D Commercial Buildings Energy-Efficiency Tax Deduction and choose to upgrade your windows, you could significantly reduce your energy consumption and costs. By improving your building's windows, you can achieve a reduction of at least 10% in cooling consumption, which typically accounts for about 30% to 50% of your building's total energy consumption. This not only enhances energy efficiency but also contributes to lower utility bills, making it a financially beneficial investment. ")
            #     st.markdown("""
            #     <div style="text-align: justify;">
            #     If you are eligible for the 179D Commercial Buildings Energy-Efficiency Tax Deduction and choose to upgrade your windows, you could significantly reduce your energy consumption and costs. By improving your building's windows, you can achieve a reduction of at least 10% in cooling consumption, which typically accounts for about 30% to 50% of your building's total energy consumption. This not only enhances energy efficiency but also contributes to lower utility bills, making it a financially beneficial investment.
            #     </div>
            #     """, unsafe_allow_html=True)
            #     # Calculate energy reduction due to potential tax eligibility
            #     #energy_reduction = TOTAL_ENERGY_CONSUMPTION * 0.3*0.1  # 10% energy reduction
            #     #annual_cost_energy_reduction = energy_reduction * Electricity_Rate  # cost saving


            #     # Prepare the first table data for rates
            #     data_Tax = {
            #         "Description": ["Energy Reduction from Tax Benefit (10%)", "Annual Cost Savings from Energy Reduction"],
            #         "Value": [f"{energy_reduction:.2f} kWh", f"${annual_cost_energy_reduction:.2f}"]
            #     }

            #     # Convert the data into a DataFrame for the first table
            #     Tax_df = pd.DataFrame(data_Tax)

            #     # Display the first table (Rates Information)
                
            #     st.table(Tax_df)

            

            
            
            
            
            
            
            
          
            # Define seasonal ranges based on typical months
            seasonal_mapping = {
                'Winter': Total_Comsuption_kwh[['Dec', 'Jan', 'Feb']].sum(),  # Dec (index 11), Jan (index 0), Feb (index 1)
                'Spring': Total_Comsuption_kwh[['Mar', 'Apr', 'May']].sum(),         # Mar (index 2), Apr (index 3), May (index 4)
                'Summer': Total_Comsuption_kwh[['Jun', 'Jul', 'Aug']].sum(),         # Jun (index 5), Jul (index 6), Aug (index 7)
                'Fall': Total_Comsuption_kwh[['Sep', 'Oct', 'Nov']].sum()           # Sep (index 8), Oct (index 9), Nov (index 10)
            }
            
            seasonal_mapping2 = {
                'Winter': total_demand_kw_12_months[['Dec', 'Jan', 'Feb']].sum(),  # Dec (index 11), Jan (index 0), Feb (index 1)
                'Spring':total_demand_kw_12_months[['Mar', 'Apr', 'May']].sum(),         # Mar (index 2), Apr (index 3), May (index 4)
                'Summer': total_demand_kw_12_months[['Jun', 'Jul', 'Aug']].sum(),         # Jun (index 5), Jul (index 6), Aug (index 7)
                'Fall': total_demand_kw_12_months[['Sep', 'Oct', 'Nov']].sum()           # Sep (index 8), Oct (index 9), Nov (index 10)
            }
            
            
            
            st.markdown("**Report Summerized:**")

            # Create the first figure with 2 subplots (1 row, 2 columns)
            fig1, axs1 = plt.subplots(1, 2, figsize=(12, 6))

            # Plot 1: Total Consumption KWh for 12 months
            axs1[0].bar(months111, Total_Comsuption_kwh, color='#0EAD23')
            axs1[0].set_ylabel("Total Consumption KWh")
            axs1[0].set_xlabel("Month")
            axs1[0].set_title("Total Consumption KWh ")
            axs1[0].grid(True, axis='y', linestyle='--', alpha=0.7)

            # Plot 2: Seasonal Total Consumption KWh
            # Calculate Max_Demand (kW)
            

            
            axs1[1].bar(months111, total_demand_kw_12_months, color='#CC181F')
            axs1[1].set_ylabel("Total Demand KW")
            axs1[1].set_xlabel("Month")
            #, and Expected Demand Lines
            axs1[1].set_title(f"Total Demand KW with Max, Min")
            axs1[1].grid(True, axis='y', linestyle='--', alpha=0.7)

            # Add horizontal line for Max_Demand (expected)
            #axs1[1].axhline(y=Max_Demand, color='darkblue', linestyle='--', linewidth=3, label=f'Max Demand Expectation = {Max_Demand:.2f} kW')

           
            # Add horizontal line for min demand
            axs1[1].axhline(y=min_demand, color='green', linestyle='--', linewidth=2, label=f'Min Demand = {min_demand:.2f} kW')

            # Add horizontal line for max demand
            axs1[1].axhline(y=max_demand, color='orange', linestyle='--', linewidth=2, label=f'Max Demand = {max_demand:.2f} kW')

            # Add a legend to show the constant values
            axs1[1].legend()

            # Adjust layout for better spacing in the figure
            plt.tight_layout()

            # Display the updated figure
            st.pyplot(fig1)

            # Proceed to plot only if there's valid non-zero data
            if ('Consumption_on_Peak_kwh' in locals() and Consumption_on_Peak_kwh is not None and np.any(Consumption_on_Peak_kwh)) or \
            ('Consumption_off_Peak_kwh' in locals() and Consumption_off_Peak_kwh is not None and np.any(Consumption_off_Peak_kwh)):

                # Create a figure and axis for the grouped bar chart for Peak and Off-Peak consumption
                fig, ax = plt.subplots(figsize=(12, 6))

                # Set the positions for each group (the months)
                x = np.arange(len(months111))  # the label locations

                # Define the width of the bars
                width = 0.35  # the width of the bars

                # Plot Consumption On Peak kWh
                if 'Consumption_on_Peak_kwh' in locals() and np.any(Consumption_on_Peak_kwh):
                    bars1 = ax.bar(x - width/2, Consumption_on_Peak_kwh, width, label='Consumption On Peak kWh', color='#CC181F')

                # Plot Consumption Off Peak kWh
                if 'Consumption_off_Peak_kwh' in locals() and np.any(Consumption_off_Peak_kwh):
                    bars2 = ax.bar(x + width/2, Consumption_off_Peak_kwh, width, label='Consumption Off Peak kWh', color='#0EAD23')

                # Calculate the ratio of on-peak to off-peak consumption
                ratio = np.divide(Consumption_on_Peak_kwh, Consumption_off_Peak_kwh, out=np.zeros_like(Consumption_on_Peak_kwh), where=Consumption_off_Peak_kwh != 0)

                # Normalize the ratio to match the scale of the consumption values
                #normalized_ratio = (ratio - np.min(ratio)) / (np.max(ratio) - np.min(ratio)) * (np.max(Consumption_on_Peak_kwh) - np.min(Consumption_on_Peak_kwh)) + np.min(Consumption_on_Peak_kwh)
                # Calculate the ratio of on-peak to off-peak consumption
                ratio = np.divide(Consumption_on_Peak_kwh, Consumption_off_Peak_kwh, 
                                  out=np.zeros_like(Consumption_on_Peak_kwh), where=Consumption_off_Peak_kwh != 0)
                
                # Ensure the denominator is not zero
                ratio_range = np.max(ratio) - np.min(ratio)
                
                if ratio_range == 0:
                    # If all values are the same, set a default normalized ratio to avoid division by zero
                    normalized_ratio = np.full_like(ratio, 1)  # Or another reasonable default
                else:
                    normalized_ratio = (ratio - np.min(ratio)) / ratio_range * (np.max(Consumption_on_Peak_kwh) - np.min(Consumption_on_Peak_kwh)) + np.min(Consumption_on_Peak_kwh)



                
                # Plot the normalized ratio as a line
                ax.plot(x, normalized_ratio, color='purple', marker='o', label='Normalized On-Peak/Off-Peak Ratio')
                
                # Set the labels for the axes
                ax.set_xlabel("Month")
                ax.set_ylabel("Consumption (kWh)")
                ax.set_title("Comparison of Consumption On Peak and Off Peak with Ratio")
                ax.set_xticks(x)
                ax.set_xticklabels(months111)  # Use the month labels from months111
                ax.grid(True, axis='y', linestyle='--', alpha=0.7)

                # Add the extra items to the legend for On-Peak and Off-Peak Electricity Price with values
                legend_elements = [
                    plt.Line2D([0], [0], color='#CC181F', lw=4, label='Consumption On Peak kWh'),
                    plt.Line2D([0], [0], color='#0EAD23', lw=4, label='Consumption Off Peak kWh'),
                    plt.Line2D([0], [0], color='purple', lw=2, linestyle='-', marker='o', label='Normalized On-Peak/Off-Peak Ratio')
                ]

                # Show the legend with all elements
                ax.legend(handles=legend_elements, loc='upper left')

                # Adjust layout for better spacing
                plt.tight_layout()

                # Display the plot in Streamlit
                st.pyplot(fig)


            else:
                st.write("")
                        
            
            

            #,Total_Energy_Fuel_Charge

            fig3, axs3 = plt.subplots(1, 2, figsize=(12, 6))
            seasons = list(seasonal_mapping.keys())
            total_consumption_values = list(seasonal_mapping.values())

            axs3[0].bar(seasons, total_consumption_values, color=['#ADD8E6', '#70FF18', '#CC181F', '#FFF225'])
            axs3[0].set_ylabel("Total Consumption KWh")
            axs3[0].set_title("Seasonal Total Consumption KWh")
            axs3[0].grid(True, axis='y', linestyle='--', alpha=0.7)
            
            
            seasons = list(seasonal_mapping2.keys())
            total_consumption_values2 = list(seasonal_mapping2.values())

            axs3[1].bar(seasons, total_consumption_values2, color=['#ADD8E6', '#70FF18', '#CC181F', '#FFF225'])
            axs3[1].set_ylabel("Total Demand KW")
            axs3[1].set_title("Seasonal Total Demand KW")
            axs3[1].grid(True, axis='y', linestyle='--', alpha=0.7)
            
            
            plt.tight_layout()

            # Display the second figure with two subplots
            st.pyplot(fig3)
            
            # Create a figure and axis for the grouped bar chart
            fig, ax = plt.subplots(figsize=(12, 6))

            # Set the positions for each group (the months)
            x = np.arange(len(months111))  # the label locations

            # Define the width of the bars
            width = 0.35  # the width of the bars

            # Plot the bars for Total Energy & Fuel Charge (lighter blue)
            bars1 = ax.bar(x - width/2, Total_Energy_Fuel_Charge, width, label='Total Energy & Fuel Charge', color='#1C65CF')

            # Plot the bars for Total Demand Charge (darker blue)
            bars2 = ax.bar(x + width/2, Total_Demand_Charge, width, label='Total Demand Charge', color='#CC181F')  # Darker blue

            # Add labels, title, and grid
            ax.set_xlabel("Month")
            ax.set_ylabel("Charges ($)")
            ax.set_title("Comparison of Total Energy & Fuel Charge and Total Demand Charge")
            ax.set_xticks(x)
            ax.set_xticklabels(months111)  # Use the month labels from months111
            ax.grid(True, axis='y', linestyle='--', alpha=0.7)

            # Add a legend to differentiate between the two bars
            ax.legend()

            # Adjust layout for better spacing
            plt.tight_layout()

            # Display the grouped bar chart in Streamlit
            st.pyplot(fig)
            
            
            # Calculate the sum of total consumption for 12 months
            total_consumption_12_months = sum(Total_Comsuption_kwh)

            # Calculate 1% reduction in energy consumption
            reduction_percentage = 0.01
            reduction_in_kwh = total_consumption_12_months * reduction_percentage

            # CO2 Reduction (tons)
            emission_factor = 0.29  # kg CO2 per kWh
            co2_reduction_kg = reduction_in_kwh * emission_factor
            co2_reduction_tons = co2_reduction_kg / 1000  # Convert to metric tons

            # Number of trees equivalent
            trees_equivalent = co2_reduction_kg / 22  # 22 kg CO2 absorbed per tree/year

            # Fancy Chart for CO2 Reduction and Trees Equivalent with dual-axis
            fig, ax1 = plt.subplots(figsize=(10, 6))

            # Bar for CO2 Reduction on the primary y-axis
            ax1.bar('CO2 Reduction (tons)', co2_reduction_tons, color='#ff6347')
            ax1.set_ylabel('CO2 Reduction (tons)', color='#ff6347')
            ax1.tick_params(axis='y', labelcolor='#ff6347')

            # Create a second y-axis for the tree equivalents
            ax2 = ax1.twinx()
            ax2.bar('Equivalent Trees', trees_equivalent, color='#228b22')
            ax2.set_ylabel('Equivalent Number of Trees', color='#228b22')
            ax2.tick_params(axis='y', labelcolor='#228b22')

            # Add value annotations
            ax1.text(0, co2_reduction_tons + 0.1, f'{co2_reduction_tons:.2f} tons', ha='center', va='bottom', color='#ff6347')
            ax2.text(1, trees_equivalent + 10, f'{trees_equivalent:.2f} trees', ha='center', va='bottom', color='#228b22')

            # Set title
            fig.suptitle(f'Impact of 1% Energy Reduction on CO₂ and Trees')

            # Display the dual-axis chart in Streamlit
            st.pyplot(fig)


            # Display calculation results
            #st.write(f"Total Energy Consumption: {total_consumption_12_months} kWh")
            #st.write(f"CO₂ Reduction from 1% Energy Reduction: {co2_reduction_tons:.2f} tons")
            #st.write(f"Equivalent Trees for CO₂ Reduction: {trees_equivalent:.2f} trees")
            
          
           
           
           # Constants for calculations
            CO2_EMISSION_FACTOR = 0.29  # kg CO2/kWh
            PERCENT_REDUCTION = 0.01  # 1% reduction
            TOTAL_ENERGY_CONSUMPTION = Total_Comsuption_kwh.sum()  # Sum of 12 months total consumption

            # Calculate CO2 reduction for 1% energy savings
            co2_reduction_kg = TOTAL_ENERGY_CONSUMPTION * PERCENT_REDUCTION * CO2_EMISSION_FACTOR
            co2_reduction_tons = co2_reduction_kg / 1000  # Convert to metric tons

            # Calculate equivalent number of trees needed for sequestration
            trees_needed = co2_reduction_kg / 22  # Each tree absorbs 22 kg of CO2 per year

            # Calculate life expectancy improvement from PM2.5 reduction assumption
            # Assuming 0.5 µg/m³ reduction in PM2.5 for every ton CO2 reduction (hypothetical assumption)
            # pm25_reduction = 0.5 * co2_reduction_tons  # µg/m³ reduction based on CO2 savings
            # life_expectancy_gain = 0.061 * (pm25_reduction / 10)  # Extrapolate based on Harvard study

            # # Visualization of CO2 Reduction and Life Expectancy Gains
            # fig, ax = plt.subplots(figsize=(12, 6))

            # # Bar chart showing CO2 reduction and equivalent trees
            # categories = ['CO2 Reduction (tons)', 'Equivalent Trees', 'Life Expectancy Gain (years)']
            # values = [co2_reduction_tons, trees_needed, life_expectancy_gain]

            # # Plotting bars
            # bars = ax.bar(categories, values, color=['#FF6347', '#87CEEB', '#4682B4'])

            # # Adding value annotations
            # for i, value in enumerate(values):
            #     ax.text(i, value + 0.05, f'{value:.2f}', ha='center', va='bottom')

            # # Add labels and title
            # ax.set_ylabel('Values')
            # ax.set_title(f'Impact of 1% Energy Reduction on CO2, Trees, and Life Expectancy')

            # # Show plot in Streamlit
            # st.pyplot(fig)

            # # Add textual details for better understanding
            # st.markdown(f"""
            # <div style="text-align: justify;">
            
            # ### The Impact of Clean Air on Life Expectancy
            # Research from the **Harvard School of Public Health** has shown that reductions in particulate matter, especially **PM2.5**, can have a significant impact on life expectancy. The study found that for every **10 µg/m³** reduction in **PM2.5**, life expectancy improved by **0.61 years** in the U.S. Below is a chart showing how different levels of PM2.5 reduction correlate with life expectancy improvements, based on the findings of this study.

            # Reducing energy consumption not only helps mitigate climate change by reducing CO₂ emissions but also contributes to cleaner air and longer life expectancy.
            # </div>
            # """, unsafe_allow_html=True)

          
        else:
            st.warning("Please upload at least one PDF file for each account and specify the number of accounts.")

if __name__ == "__main__":
    app()

