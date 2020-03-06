import streamlit as st
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import json
import os
import datetime

"""
# Utilization Report
"""
# TODO
# vs Planned
# Semester 1 Avg
# Semester 2 Avg
# Yearly Avg
# User input for utilization

@st.cache
def auth_gspread():
    scope = ['https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive']
    try:
        # creds for local development
        creds = ServiceAccountCredentials.from_json_keyfile_name(
            'secrets/gs_credentials.json', scope
        )
        client = gspread.authorize(creds)
    except:
        # creds for heroku deployment
        json_creds = os.environ.get("GOOGLE_SHEETS_CREDS_JSON")
        creds_dict = json.loads(json_creds)
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\\\n", "\n")
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)

    # load_hours_report:
    hours_wks = client.open("Utilization-Hours").sheet1
    data = hours_wks.get_all_values()
    headers = data.pop(0)
    df = pd.DataFrame(data, columns=headers)
    df['Entry Date'] = pd.to_datetime(df['Entry Date'])
    df['Hours Worked'] = pd.to_numeric(df['Hours Worked'])
    df['Time Off Hrs'] = pd.to_numeric(df['Time Off Hrs'])
    df['Entry Month'] = pd.DatetimeIndex(df['Entry Date']).strftime('%b')
    df['Hours Worked'] = df['Hours Worked'] + df['Time Off Hrs']
    df['Activity Name'] = df['Activity Name'] + df['Time Off Type']
    df.drop(['Time Off Hrs', 'Time Off Type'], axis=1, inplace=True)
    
    # Activity names are imported with trailing whitespace, use pd.str.strip to remove
    df['Activity Name'] = df['Activity Name'].str.strip()

    # load_activities:
    wks = client.open("Utilization-Inputs").worksheet('ACTIVITY')
    data = wks.get_all_values()
    headers = data.pop(0)
    activities = pd.DataFrame(data, columns=headers)
    
    # load_date_info:
    wks = client.open("Utilization-Inputs").worksheet('DATES')
    data = wks.get_all_values()
    headers = data.pop(0)
    dates = pd.DataFrame(data, columns=headers)
    dates['Date'] = pd.to_datetime(dates['Date'])
    dates['Remaining'] = pd.to_numeric(dates['Remaining'])
    dates['Month'] = pd.DatetimeIndex(dates['Date']).strftime('%b')
    months = dates.groupby('Month').max()
    months['FTE'] = months['Remaining'] * 8
    
    # load_employees:
    wks = client.open("Utilization-Inputs").worksheet('NAMES')
    data = wks.get_all_values()
    headers = data.pop(0)
    employees = pd.DataFrame(data, columns=headers)
    names = (['Please select your name']
             + list(employees['User Name'].unique()))
    
    # load targets
    wks = client.open("Utilization-Inputs").worksheet('TARGETS')
    data = wks.get_all_values()
    headers = data.pop(0)
    targets = pd.DataFrame(data, columns=headers)

    return df, activities, dates, months, names, targets


def build_utilization(name, hours_report, activities, dates, months, 
                      method="This Month to Date", provided_utilization=None):
    # Subset and copy (don't mutate cached data, per doc)
    df = hours_report.loc[hours_report['User Name']==name].copy()

    # Join activities
    df = df.set_index('Activity Name').join(
        activities.set_index('Activity Name')
        ).reset_index()
    
    # Set any null Classification values to 'Billable' in case not in Activity sheet
    df['Classification'].fillna('Billable', inplace=True) 
    
    # Calculate monthly total hours
    individual_hours = df.groupby(['Entry Month', 'Classification']).sum().reset_index()
    
    # Save hours by category (copy utilization and merge other columns)
    utilization = individual_hours.loc[individual_hours['Classification']=='Billable'].copy()
    r_and_d = individual_hours.loc[individual_hours['Classification']=='R&D']
    other = individual_hours.loc[individual_hours['Classification']=='Other']
    time_off = individual_hours.loc[individual_hours['Classification']=='Time Off']
    
    # Join labor categories to utilization table
    utilization = pd.merge(utilization, r_and_d, on='Entry Month', how='outer', suffixes=('','_y'))
    utilization.drop('Classification_y', inplace=True, axis=1)
    utilization.rename(columns={'Hours Worked_y': 'R&D'}, inplace=True)
    
    utilization = pd.merge(utilization, other, on='Entry Month', how='outer', suffixes=('','_y'))
    utilization.drop('Classification_y', inplace=True, axis=1)
    utilization.rename(columns={'Hours Worked_y': 'Other'}, inplace=True)
    
    utilization = pd.merge(utilization, time_off, on='Entry Month', how='outer', suffixes=('','_y'))
    utilization.drop('Classification_y', inplace=True, axis=1)
    utilization.rename(columns={'Hours Worked_y': 'Time Off'}, inplace=True)
    
    # Remove classification column and change Hours worked to Billable, update NaN to 0
    utilization.drop('Classification', inplace=True, axis=1)
    utilization.rename(columns={'Hours Worked': 'Billable'}, inplace=True)
    utilization.fillna(0, inplace=True)

    # Create list of months and month dictionary for sorting later
    global list_months, semester1, semester2
    list_months = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec',
                   'Jan', 'Feb', 'Mar']
    semester1 = list_months[:7]
    semester2 = list_months[7:]
    
    month_index = np.arange(0,12)
    month_dict = dict(zip(list_months, month_index))
    
    # Sort by month    
    utilization['id'] = utilization['Entry Month'].replace(month_dict)
    utilization.sort_values('id', inplace=True)
    utilization.set_index('Entry Month', inplace=True)
    utilization.drop('id', axis=1, inplace=True)
    
    # Save variables related to this month for prediction later on
    # this_month = utilization.last_valid_index()
    last_day_worked = df.loc[~df['Activity Name'].isin(['Holiday']), 'Entry Date'].max()
    if last_day_worked > datetime.datetime.today():
        last_day_worked = datetime.datetime.strptime(
            datetime.datetime.today().strftime('%Y-%m-%d 00:00:00'), 
            '%Y-%m-%d 00:00:00')
    global this_month
    this_month = last_day_worked.strftime('%b')
    first_day_worked = df['Entry Date'].min()
    days_remaining = dates.loc[dates['Date']==last_day_worked, 'Remaining'] - 1
    
    # Zero fill billable for remaining months
    existing_months = utilization.index
    columns = utilization.reset_index().columns
    utilization.reset_index(inplace=True)
    for m in list_months:    
        if m not in existing_months:
            new_row = pd.Series([m, 0, 0, 0, 0], columns)
            utilization = utilization.append(new_row, ignore_index=True)
    
    # Sort by month again   
    utilization['id'] = utilization['Entry Month'].replace(month_dict)
    utilization.sort_values('id', inplace=True)
    utilization.set_index('Entry Month', inplace=True)
    utilization.drop('id', axis=1, inplace=True)
    
    # Update billable with FTE per month
    utilization = utilization.join(months['FTE'])
    
    # Correct FTE for employees who start in the middle of the performance period
    first_month_worked = first_day_worked.strftime('%b')
    first_month_index = list_months.index(first_month_worked)
    first_month_FTE = dates.loc[dates['Date']==first_day_worked, 'Remaining'] * 8
    
    for m in list_months[0:first_month_index]:
        utilization.at[m, 'FTE'] = 0
    
    utilization.at[first_month_worked, 'FTE'] = first_month_FTE   
    
    # Calculate actual utilization
    utilization['Utilization'] = utilization['Billable'] / utilization['FTE']
    
    # Calculate predicted utilization for this month
    # Copy Utilization to new column, Util to Date
    utilization['Util to Date'] = utilization['Utilization']
    
    # Calculate key variables
    current_hours = utilization.loc[this_month, 'Billable']
    fte_hours = utilization.loc[this_month, 'FTE']
    fte_hours_to_date = fte_hours - days_remaining * 8
    predicted_hours = (current_hours/fte_hours_to_date) * fte_hours
    
    # Update Util to Date column at the current month with predicted
    utilization.at[this_month, 'Util to Date'] = (
        predicted_hours/fte_hours
        )
    
    # Forecast forward looking utilization    
    # Create new column for predicted hours
    utilization['Predicted Hours'] = utilization['Util to Date'] * utilization['FTE']
    
    # Predict the utilization for future months based on the method selected
    # if provided_utilization:
    #     predicted = provided_utilization/100
    if method == "Month to Date":
        predicted = utilization.loc[this_month, 'Util to Date']
    elif method == "Last Month":
        if list_months.index(this_month) > 0:
            last_month = list_months[list_months.index(this_month)-1]
            predicted = utilization.loc[last_month, 'Utilization']
        else:
            predicted = utilization.loc[this_month, 'Util to Date']
    elif method == "Year (Semester) to Date":
        current_month_index = list_months.index(this_month)
        if not by_semester:
            current_df = utilization.iloc[0:current_month_index+1]
            predicted = (
                (current_df['Predicted Hours'].sum())
                / current_df['FTE'].sum()
                )
        else:
            current_df = utilization.iloc[7:current_month_index+1]
            predicted = (
                (current_df['Predicted Hours'].sum())
                / current_df['FTE'].sum()
                )
    
    # Populate future months with predicted
    future_months = list_months[list_months.index(this_month) + 1:]
    for m in future_months:
        utilization.at[m, 'Predicted Hours'] = (predicted 
                                                * utilization.loc[m, 'FTE']
                                                )
    
    # Calculate cumulative utilization for each semester    
    utilization['Predicted Utilization'] = (utilization['Predicted Hours'].cumsum() 
                                           / utilization['FTE'].cumsum())
    # Update chart for semester to date
    if by_semester:
        utilization.loc[semester2, 'Predicted Utilization'] = (
            utilization.loc[semester2, 'Predicted Hours'].cumsum() 
            / utilization.loc[semester2, 'FTE'].cumsum()
            )
        
    # Format last day worked for printing
    last_day_f = last_day_worked.strftime('%A, %B %e, %Y')
    
    return utilization, last_day_f


def plot_hours(data, target, mode='focus'):
    plt.rcParams['font.sans-serif'] = 'Tahoma'
    plt.rcParams['font.family'] = 'sans-serif'
    plt.rcParams['font.size'] = 13

    util_color = '#006040'
    r_and_d_color = '#c89051'
    other_color = '#f0ca6c'
    time_off_color = '#A7A7A7'
    full_time_color = '#ee6642'
    
    c_util_color = '#5B9BD5'
    c_r_and_d_color = '#ED7D31'
    c_other_color = '#A5A5A5'
    c_time_off_color = '#FFC000'
    c_full_time_color = '#FF0000'

    util_target = target
    
    util_value = (data.loc[data.index=='Mar', 'Predicted Utilization']* 100)

    current_month = list_months.index(this_month)
    current_month_index = current_month

    fig, ax1 = plt.subplots(figsize=[10.75,7])    

    # Plot data by mode
    if mode == 'Predictive':
        ind = np.arange(12)
        # Hide grid lines to denote prediction portion of graph, Note zorder must be specified
        # in fill_between call
        for i in np.arange(current_month_index + 1, 13):
            ax1.axes.axvline(i, color='white', linewidth=2)
        
        if by_semester:
            ax1.plot(data.loc[semester1, 'Predicted Utilization']*100, color=util_color, linewidth=3, alpha=.85)
            ax1.plot(data.loc[semester2, 'Predicted Utilization']*100, color=util_color, linewidth=3, alpha=.85)
        else:
            ax1.plot(data.loc[:,'Predicted Utilization']*100, color=util_color, linewidth=3, alpha=.85)
        
        # Plot actuals
        ax1.plot(data['Utilization']*100, color=util_color, marker='o', lw=0)

        # Plot projected
        ax1.plot(data['Util to Date']*100, color=util_color, marker='x', lw=0, alpha=1)

        # Plot targets
        ax1.plot([util_target]*12, color=util_color, linestyle='dotted')
        
        # Label actuals
        for x, y in zip(np.arange(0,12), data['Utilization']*100):
            label = f'{y:.0f}%'
            if y > 0:
                ax1.annotate(label, 
                            (x, y), 
                            textcoords="offset points", 
                            xytext=(10,0), 
                            ha='left',
                            va='center',
                            color = 'dimgrey')
        
        # Adjust axes ranges
        ax1.set_ylim(0, 120)

        # Adjust number of labels
        ax1.yaxis.set_major_locator(plt.MaxNLocator(6))
        
        x_labels = data.index
        
        # Label
        ax1.text(11.1, util_value-3, f' Predicted \n Utilization ({int(util_value)}%)', 
                color=util_color)
        
        # Set title
        ax1.set_title('Are you on track to meet your utilization target?', 
                      loc='right', 
                      fontsize=15)

    elif mode == 'Classic':           
        width = .25
        billable_hours = data.loc[:, 'Util to Date']*100
        r_and_d_hours = (data.loc[:, 'R&D']/data.loc[:, 'FTE'])*100
        other_hours = (data.loc[:, 'Other']/data.loc[:, 'FTE'])*100
        time_off_hours = (data.loc[:, 'Time Off']/data.loc[:, 'FTE'])*100
        
        if by_semester:
            ind = np.arange(12+2)
            
            # update hours for S1
            s1_upper = list_months[min(6, current_month_index)]
            data_s1 = data.loc[:s1_upper, :]
            
            billable_hours = billable_hours.append(pd.Series(data_s1['Billable'].sum()/data_s1['FTE'].sum()*100, index=['S1']))
            r_and_d_hours = r_and_d_hours.append(pd.Series(data_s1['R&D'].sum()/data_s1['FTE'].sum()*100, index=['S1']))
            other_hours = other_hours.append(pd.Series(data_s1['Other'].sum()/data_s1['FTE'].sum()*100, index=['S1']))
            time_off_hours = time_off_hours.append(pd.Series(data_s1['Time Off'].sum()/data_s1['FTE'].sum()*100, index=['S1']))
            
            # update hours for S2
            s2_lower = list_months[6]
            s2_upper = list_months[current_month_index]
            data_s2 = data.loc[s2_lower:s2_upper, :]
            
            billable_hours = billable_hours.append(pd.Series(data_s2['Billable'].sum()/data_s2['FTE'].sum()*100, index=['S2']))
            r_and_d_hours = r_and_d_hours.append(pd.Series(data_s2['R&D'].sum()/data_s2['FTE'].sum()*100, index=['S2']))
            other_hours = other_hours.append(pd.Series(data_s2['Other'].sum()/data_s2['FTE'].sum()*100, index=['S2']))
            time_off_hours = time_off_hours.append(pd.Series(data_s2['Time Off'].sum()/data_s2['FTE'].sum()*100, index=['S2']))
            
            
            x_labels = billable_hours.index
            
        else:
            # update hours with year average
            ind = np.arange(12+1)
            billable_hours = billable_hours.append(pd.Series(data['Billable'].sum()/data.loc[:this_month, 'FTE'].sum()*100, index=['Year']))
            r_and_d_hours = r_and_d_hours.append(pd.Series(data['R&D'].sum()/data.loc[:this_month, 'FTE'].sum()*100, index=['Year']))
            other_hours = other_hours.append(pd.Series(data['Other'].sum()/data.loc[:this_month, 'FTE'].sum()*100, index=['Year']))
            time_off_hours = time_off_hours.append(pd.Series(data['Time Off'].sum()/data.loc[:this_month, 'FTE'].sum()*100, index=['Year']))
            
            x_labels = billable_hours.index
            
        billable_hours.fillna(0, inplace=True)
        r_and_d_hours.fillna(0, inplace=True)
        other_hours.fillna(0, inplace=True)
        time_off_hours.fillna(0, inplace=True)
        
        ax1.bar(ind, billable_hours, width=width, color=c_util_color, label='Utilization')
        ax1.bar(ind, r_and_d_hours, width=width, bottom=billable_hours, color=c_r_and_d_color, label='R&D')
        ax1.bar(ind, other_hours, width=width, bottom=billable_hours + r_and_d_hours, color=c_other_color, label='Other')
        ax1.bar(ind, time_off_hours, width=width, bottom=billable_hours + r_and_d_hours+other_hours, color=c_time_off_color, label='Time Off')
        
        # Plot targets
        ax1.plot([util_target]*len(ind), color=c_util_color, linestyle='dotted')
        ax1.plot([110]*len(ind), color='#70AD47', linestyle='dotted')
        ax1.plot([125]*len(ind), color=full_time_color, linestyle='dotted')        
        
        # Plot planned utilization
        target_df = targets.loc[targets['User Name'] == name, list_months]
        target_df[list_months] = target_df[list_months].apply(pd.to_numeric)
        target_df.fillna(0, inplace=True)
        if target_df.empty:
            pass
        else:
            target_util = target_df.values.tolist()[0]
            target_util = [t * 100 for t in target_util]        
        
            ax1.plot(target_util, marker='s', markerfacecolor=c_util_color, markeredgewidth=1, markeredgecolor='white', lw=0, alpha=1, label = 'Planned Utilization')
        
        # Add legend
        ax1.legend(loc='upper center', bbox_to_anchor=(0.5, -0.05), ncol=5, frameon=False, columnspacing=3)

        # Adjust axes ranges
        ax1.set_ylim(0, 140)
        
        # Adjust number of labels
        ax1.yaxis.set_major_locator(plt.MaxNLocator(7))

    # Format y labels as percent
    ax1.yaxis.set_major_formatter(plt.FuncFormatter('{:.0f}%'.format))

    # Set x labels
    ax1.set_xticks(ind)
    ax1.set_xticklabels(x_labels)

    # Add grid Lines
    ax1.yaxis.grid(False)
    ax1.xaxis.grid(True)

    # Customize grid lines
    ax1.axes.grid(axis='x', linestyle='-')

    # Set below graph objects
    ax1.set_axisbelow(True)

    # Remove Axes ticks
    ax1.tick_params(axis='both', which='both', 
                    bottom=False, top=False, left=False, right=False)

    # Recolor axis labels
    ax1.tick_params(colors='dimgrey')

    # Remove axes spines
    ax1.spines['top'].set_visible(False)
    ax1.spines['left'].set_visible(False)
    ax1.spines['right'].set_visible(False)
    ax1.spines['bottom'].set_visible(True)
    ax1.spines['bottom'].set_color('silver')

    # Indicate current month
    ax1.get_xticklabels()[current_month_index].set_fontweight('bold')
    if mode == 'Classic':
        util_color = c_util_color
    ax1.get_xticklabels()[current_month_index].set_color(util_color)
    
    return fig, util_value.item()


def message(predicted, target):
    if predicted > target and target > 0:
        message_loc.success("You're on track to meet your utilization!")
    elif predicted < target and target > 0: 
        diff = round(target - predicted, 0)
        message_loc.warning(f"You're on track to miss your target by {diff}%")

def balloons(predicted, target):
    if predicted > target and target > 0:
        st.balloons()
        

# Load data
hours_report, activities, dates, months, names, targets = auth_gspread()

# User selects name
name = st.selectbox(
    'Who are you?', 
    (names)
)

# User inputs target utilization
target_util = st.number_input("What's your target utilization?", 0, 100, 0)

chart_loc = st.empty()
message_loc = st.empty()
        
# Build utilization report for user
if name != names[0]:
    
    # User selects mode
    modes = ['Predictive', 'Classic']
    mode = st.sidebar.selectbox(
        'Select Mode', (modes)
    )
    method = "Year (Semester) to Date"
    
    # User inputs prediction method
    if mode == 'Predictive':
        methods = ["Month to Date", "Last Month", "Year (Semester) to Date"]
        method = st.sidebar.selectbox(
            'I would like to change how you predict my utilization. '
            'Use my utilization from:',
            (methods)
        )
    
    by_semester = False
    if st.sidebar.checkbox('Split the data by semester'):
        by_semester = True

# # User inputs raw value for prediction
# provided_utilization = st.number_input("Use this value to predict my utilizion "
#                                        " in the future. I plan to maintain this "
#                                        " utilization going forward.", 0, 100)

    df, valid_date = build_utilization(name, hours_report, activities, dates, months, method)

    # Plot results
    plot, predicted_utilization = plot_hours(df, target_util, mode)
    chart_loc.pyplot(plot)

    # Display a congratulatory or warning message based on prediction 
    message(predicted_utilization, target_util)
        
    # User may display data
    if st.checkbox('Show data'):
        st.subheader('Utilization Data')
        st.table(df)

    # ...And balloons, just cause
    if not st.sidebar.checkbox("That's enough balloons"):
        balloons(predicted_utilization, target_util)
    
    st.write('')
    st.write('')
    st.write('')
    st.write(f'Data valid through {valid_date}')