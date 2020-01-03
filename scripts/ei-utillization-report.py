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
# No FTE support yet
# No Practice Area support

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
    
    # return client


# @st.cache
# def load_hours_report(client):
    # df = pd.read_excel(hours_report_path, parse_dates=['Entry Date'])
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
    
    # return df


# @st.cache
# def load_activities(client):
    wks = client.open("Utilization-Inputs").worksheet('ACTIVITY')
    data = wks.get_all_values()
    headers = data.pop(0)
    activities = pd.DataFrame(data, columns=headers)
    
    # return activities


# @st.cache
# def load_date_info(client):
    wks = client.open("Utilization-Inputs").worksheet('DATES')
    data = wks.get_all_values()
    headers = data.pop(0)
    dates = pd.DataFrame(data, columns=headers)
    dates['Date'] = pd.to_datetime(dates['Date'])
    dates['Remaining'] = pd.to_numeric(dates['Remaining'])
    dates['Month'] = pd.DatetimeIndex(dates['Date']).strftime('%b')
    months = dates.groupby('Month').max()
    months['FTE'] = months['Remaining'] * 8
    
    # return dates, months


# @st.cache
# def load_employees(client):
    wks = client.open("Utilization-Inputs").worksheet('NAMES')
    data = wks.get_all_values()
    headers = data.pop(0)
    employees = pd.DataFrame(data, columns=headers)
    names = (['Please select your name']
             + list(employees['User Name'].unique()))
    
    # return names

    return df, activities, dates, months, names


def build_utilization(name, hours_report, activities, dates, months, 
                      method="This Month to Date", provided_utilization=None):
    # Subset and copy (don't mutate cached data, per doc)
    df = hours_report.loc[hours_report['User Name']==name].copy()

    # Join activities
    df = df.set_index('Activity Name').join(
        activities.set_index('Activity Name')
        ).reset_index()

    # Calculate monthly total hours
    individual_hours = df.groupby(['Entry Month', 'Classification']).sum().reset_index()
    
    # Save only billable hours
    utilization = individual_hours.loc[individual_hours['Classification']=='Billable'].copy()

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
        last_day_worked = datetime.datetime.today().date()
    global this_month
    this_month = last_day_worked.strftime('%b')
    first_day_worked = df['Entry Date'].min()
    days_remaining = dates.loc[dates['Date']==last_day_worked, 'Remaining'] - 1
    
    # Zero fill billable for remaining months (IMPROVE)
    # list_months = months.index
    existing_months = utilization.index
    columns = utilization.reset_index().columns
    utilization.reset_index(inplace=True)
    for m in list_months:    
        if m not in existing_months:
            new_row = pd.Series([m, "Billable", 0], columns)
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
    utilization['Utilization'] = utilization['Hours Worked'] / utilization['FTE']
    
    # Calculate predicted utilization for this month
    # Copy Utilization to new column, Util to Date
    utilization['Util to Date'] = utilization['Utilization']
    
    # Calculate key variables
    current_hours = utilization.loc[this_month, 'Hours Worked']
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


def plot_hours(data, target):
    plt.rcParams['font.sans-serif'] = 'Tahoma'
    plt.rcParams['font.family'] = 'sans-serif'
    plt.rcParams['font.size'] = 13

    util_color = '#006040'
    util_target = target

    current_month = list_months.index(this_month)
    current_month_index = current_month

    fig, ax1 = plt.subplots(figsize=[10.75,7])

    # Hide grid lines to denote prediction portion of graph, Note zorder must be specified
    # in fill_between call
    for i in np.arange(current_month_index + 1, 13):
        ax1.axes.axvline(i, color='white', linewidth=2)

    # Plot data
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

    # Set title
    ax1.set_title('Are you on track to meet your utilization target?', 
                  loc='right', 
                  fontsize=15)

    # Adjust axes ranges
    ax1.set_ylim(0, 120)

    # Adjust number of labels
    ax1.yaxis.set_major_locator(plt.MaxNLocator(6))

    # Format y labels as percent
    ax1.yaxis.set_major_formatter(plt.FuncFormatter('{:.0f}%'.format))

    # Set x labels
    ax1.set_xticks(np.arange(0,12))
    ax1.set_xticklabels(data.index)

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

    # Labels
    util_value = (data.loc[data.index=='Mar', 'Predicted Utilization']
                  * 100)
    ax1.text(11.1, util_value-3, f' Predicted \n Utilization ({int(util_value)}%)', 
             color=util_color)

    # Indicate current month
    ax1.get_xticklabels()[current_month_index].set_fontweight('bold')
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
hours_report, activities, dates, months, names = auth_gspread()

# User selects name
name = st.selectbox(
    'Who are you?', 
    (names)
)

# User inputs target utilization
target_util = st.number_input("What's your target utilization?", 0, 100, 0)

chart_loc = st.empty()
message_loc = st.empty()
method_loc = st.empty()
        
# Build utilization report for user
if name != names[0]:
    # User inputs prediction method
    methods = ["Month to Date", "Last Month", "Year (Semester) to Date"]
    method = method_loc.selectbox(
        'I would like to change how you predict my utilization. '
        'Use my utilization from:',
        (methods)
    )
    
    by_semester = False
    if st.checkbox('Split the data by semester'):
        by_semester = True

# # User inputs raw value for prediction
# provided_utilization = st.number_input("Use this value to predict my utilizion "
#                                        " in the future. I plan to maintain this "
#                                        " utilization going forward.", 0, 100)

    df, valid_date = build_utilization(name, hours_report, activities, dates, months, method)

    # Plot results
    plot, predicted_utilization = plot_hours(df, target_util)
    chart_loc.pyplot(plot)

    # Display a congratulatory or warning message based on prediction 
    message(predicted_utilization, target_util)
        
    # User may display data
    if st.checkbox('Show data'):
        st.subheader('Utilization Data')
        st.table(df)

    # ...And balloons, just cause
    if not st.checkbox("That's enough balloons"):
        balloons(predicted_utilization, target_util)
    
    st.write('')
    st.write('')
    st.write('')
    st.write(f'Data valid through {valid_date}')