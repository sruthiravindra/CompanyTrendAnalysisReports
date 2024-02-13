import pandas as pd 
import plotly.express as px
import streamlit as st
from PIL import Image
import matplotlib.pyplot as plt

# Define custom CSS
custom_css = """
<style>
h2{
    text-align:center
}
[data-testid="column"] {
    border: 1px solid #a28e8e;
    color: black;
    padding: 20px;
    border-radius: 6px;
}
[data-testid="column"]:nth-of-type(2) {
    border: 1px solid #94c194;
}
.stSelectbox { 
                width: 200px !important; 
            }
</style>
"""

company_options = 'ITC'
excel_file_name = 'Income and Balance sheet-ITC.xlsx'
st.set_page_config(page_title='Financials for ITC',
                   page_icon=':maple_leaf:',
                   layout='wide')


# Render custom CSS
st.markdown(custom_css, unsafe_allow_html=True)

# Define the options
company_options = ('ITC', 'Hindustan Lever', 'Britannia')

# Dropdow to select the company excel file
st.markdown('#')
st.markdown('#')
company_option = st.selectbox('Select Company:', ('ITC', 'Hindustan Lever', 'Britannia'), key='company_option')
st.markdown('#')

if company_option == 'ITC':
    st.header('Financials for ITC')
    excel_file_name = 'Income and Balance sheet-ITC.xlsx'
elif company_option == 'Hindustan Lever':
    st.header('Financials for Hindustan Lever')
    excel_file_name = 'Income and Balance sheet-HUL.xlsx'
elif company_option == 'Britannia':
    st.header('Financials for Britannia')
    excel_file_name = 'Illustration_FSA_MBA939.xlsx'

df_income_statement = pd.read_excel(
    io=excel_file_name,
    engine='openpyxl',
    sheet_name='IncomeStatement',
    skiprows=1,
    usecols='B:G',
    nrows=41
)
df_balance_sheet = pd.read_excel(
    io=excel_file_name,
    engine='openpyxl',
    sheet_name='BalanceSheet',
    skiprows=1,
    usecols='B:G',
    nrows=52
)

df_income_statement_2 = pd.read_excel(excel_file_name, sheet_name='IncomeStatement', usecols="I:N", skiprows=1, nrows=3)

df_balance_sheet_2 = pd.read_excel(excel_file_name, sheet_name='BalanceSheet', usecols="I:N", skiprows=1, nrows=6)


def display_sheet_data(df_income_statement, df_balance_sheet):
    # Dropdown to select between Income Statement and Balance Sheet
    option = st.selectbox('Select report:', ('Income Statement', 'Balance Sheet'))

    # Display DataFrame based on dropdown selection
    if option == 'Income Statement':
        st.subheader('Income Statement')
        st.dataframe(df_income_statement)
    elif option == 'Balance Sheet':
        st.subheader('Balance Sheet')
        st.dataframe(df_balance_sheet)
    st.markdown('#')


#=SERIES(,IncomeStatement!$J$2:$N$2,IncomeStatement!$J$4:$N$4,1)
def plot_revenue_trend(df):
    
    # Extract the required data
    revenues = df.iloc[8, 1:7]  # Assuming the revenue data is in the eleventh row from column C to G

    # Convert index to datetime format
    revenues.index = pd.to_datetime(revenues.index)
    
    # Group by year and sum the revenues
    yearly_revenues = revenues.groupby(revenues.index.year).sum()
        
    # Create a DataFrame with 'year' and 'revenue' columns
    revenue_df = pd.DataFrame({'Year': yearly_revenues.index, 'Total Revenue (In Crore)': yearly_revenues.values})

    revenue_df

    # Plot the line plot using Plotly Express
    fig = px.line(
        revenue_df,
        x='Year',  # Specify the column containing the x-axis values (years)
        y='Total Revenue (In Crore)',  # Specify the column containing the y-axis values (revenues)
        markers=True,
        title="Revenue Trend",
    )

    # Display the plot using Streamlit
    st.plotly_chart(fig)

#=SERIES(BalanceSheet!$I$3,BalanceSheet!$J$2:$N$2,BalanceSheet!$J$3:$N$3,1)

def plot_shareholder_trend(df):
    
    expenses = df.iloc[0, 1:]

    # Convert index to datetime format
    expenses.index = pd.to_datetime(expenses.index)
    
    # Group by year and sum the expenses
    yearly_expenses = expenses.groupby(expenses.index.year).sum()
        
    # Create a DataFrame with 'year' and 'expenses' columns
    expenses_df = pd.DataFrame({'Year': yearly_expenses.index, 'Shareholder Fund (In Crore)': yearly_expenses.values})

    expenses_df

    # Plot the line plot using Plotly Express
    fig = px.line(
        expenses_df,
        x='Year',  # Specify the column containing the x-axis values (years)
        y='Shareholder Fund (In Crore)',  # Specify the column containing the y-axis values (expenses)
        markers=True,
        title="Total Shareholder Fund",
    )

    # Display the plot using Streamlit
    st.plotly_chart(fig)


def plot_single_trend(df, x_label, y_label, title):
    
    plot_data = df

    # Convert index to datetime format
    plot_data.index = pd.to_datetime(plot_data.index)
    
    # Group by year and sum the plot_data
    yearly_plot_data = plot_data.groupby(plot_data.index.year).sum()
        
    # Create a DataFrame with 'year' and 'plot_data' columns
    plot_data_df = pd.DataFrame({x_label: yearly_plot_data.index, y_label: yearly_plot_data.values})

    plot_data_df

    # Plot the line plot using Plotly Express
    fig = px.line(
        plot_data_df,
        x=x_label,  # Specify the column containing the x-axis values (years)
        y=y_label,  # Specify the column containing the y-axis values (plot_data)
        markers=True,
        title=title,
    )

    # Display the plot using Streamlit
    st.plotly_chart(fig)


def plot_two_trend(df, df2, x_label, y_label, plot1_legend, plot2_legend, title):
    
    plot_data = df
    plot_data2 = df2

    # Convert index to datetime format
    plot_data.index = pd.to_datetime(plot_data.index)
    plot_data2.index = pd.to_datetime(plot_data2.index)
    
    # Group by year and sum the plot_data
    yearly_plot_data = plot_data.groupby(plot_data.index.year).sum()
    yearly_plot_data2 = plot_data2.groupby(plot_data2.index.year).sum()
        
    # Create a DataFrame with 'year' and 'plot_data' columns
    df_combined = pd.DataFrame({
        x_label: yearly_plot_data.index,
        plot1_legend: yearly_plot_data.values.reshape(-1),
        plot2_legend: yearly_plot_data2.values.reshape(-1)
    })

    df_combined
    
    # Melt the DataFrame to long format
    df_melted = pd.melt(df_combined, id_vars=[x_label], var_name='Dataset', value_name=y_label)
    
    # Plot the line plot using Plotly Express
    fig = px.line(
        df_melted,
        x=x_label,  # Specify the column containing the x-axis values (years)
        y=y_label,  # Specify the column containing the y-axis values (plot_data)
        color='Dataset',  # Color by dataset
        markers=True,
        title=title
    )

    # Display the plot using Streamlit
    st.plotly_chart(fig)



display_sheet_data(df_income_statement, df_balance_sheet)


col1, col2 = st.columns(2)

with col1:
   # =SERIES(,IncomeStatement!$C$2:$G$2,IncomeStatement!$C$11:$G$11,3)
   revenue = df_income_statement.iloc[8, 1:7]
   st.header("Revenue Trend")
   plot_single_trend(revenue, 'Year', 'Total Revenue (In Crores)', 'Revenue Trend')

with col2:
   #=SERIES(BalanceSheet!$I$3,BalanceSheet!$J$2:$N$2,BalanceSheet!$J$3:$N$3,1)
   shareholder = df_balance_sheet_2.iloc[0, 1:]
   st.header("Total Shareholders Fund")
   plot_single_trend(shareholder, 'Year', 'Shareholder Fund (In Crores)', 'Total Shareholders Fund')

col1, col2 = st.columns(2)
with col1:
   # =SERIES(,IncomeStatement!$J$2:$N$2,IncomeStatement!$J$4:$N$4,1)
   expenses = df_income_statement_2.iloc[1, 1:]
   st.header("Expenses Trend")
   plot_single_trend(expenses, 'Year', 'Total Expenses (In Crores)', 'Expenses Trend')

with col2:
    # =SERIES(BalanceSheet!$I$7,BalanceSheet!$J$2:$N$2,BalanceSheet!$J$7:$N$7,5)
    asset = df_balance_sheet_2.iloc[4, 1:]
    liability = df_balance_sheet_2.iloc[2, 1:] # =SERIES(BalanceSheet!$I$5,BalanceSheet!$J$2:$N$2,BalanceSheet!$J$5:$N$5,3)
    st.header("Current Assets & Current Liabilities Trend")
    plot_two_trend(asset,liability, 'Year', 'Amount (In Crores)', 'Total Current Assets', 'Total Current Liabilities', 'Current Assets & Current Liabilities Trend')

col1, col2 = st.columns(2)
with col1:
   # =SERIES(,IncomeStatement!$J$2:$N$2,IncomeStatement!$J$5:$N$5,1)
   profit = df_income_statement_2.iloc[2, 1:]
   st.header("Profit Trend")
   plot_single_trend(profit, 'Year', 'Profit for the year (In Crores)', 'Profit Trend')

with col2:
    asset = df_balance_sheet_2.iloc[3, 1:] # =SERIES(BalanceSheet!$I$6,BalanceSheet!$J$2:$N$2,BalanceSheet!$J$6:$N$6,4)
    liability = df_balance_sheet_2.iloc[1, 1:] # =SERIES(BalanceSheet!$I$4,BalanceSheet!$J$2:$N$2,BalanceSheet!$J$4:$N$4,2)
    st.header("Non Current Assets & Non Current Liabilities Trend")
    plot_two_trend(asset,liability, 'Year', 'Amount (In Crores)', 'Total Non-Current Assets', 'Total Non-Current Liabilities', 'Non Current Assets & Non Current Liabilities Trend')

col1, col2 = st.columns(2)
with col1:
   # =SERIES(BalanceSheet!$I$8,BalanceSheet!$J$2:$N$2,BalanceSheet!$J$8:$N$8,1)
   asset = df_balance_sheet_2.iloc[5, 1:]
   st.header("Total Assets")
   plot_single_trend(asset, 'Year', 'Total Assets (In Crores)', 'Total Assets')