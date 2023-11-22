import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.ticker import MaxNLocator
import numpy as np
import datetime
import matplotlib.dates as mdates
from matplotlib.ticker import FuncFormatter
import seaborn as sns
import calendar



plt.style.use('seaborn-v0_8-paper')
df = pd.read_excel('Sales01.xlsx')

df['Month'] = pd.to_datetime(df['Month']).dt.to_period('M')

pdf_path = 'shop_sales_report.pdf'
pdf_pages = PdfPages(pdf_path)


def add_trendline(ax, dates, y, order=1, color='C2', linestyle='--', linewidth=2):
    dates = np.array(dates)
    y = np.array(y)

    # Filter out any NaN values
    mask = ~np.isnan(y)
    dates = dates[mask]
    y = y[mask]

    if len(y) > order:
        x = np.array([mdates.date2num(date) for date in dates])

        coeffs = np.polyfit(x, y, order)
        trend = np.poly1d(coeffs)

        x_for_plot = np.linspace(x.min(), x.max(), len(x))

        ax.plot(mdates.num2date(x_for_plot), trend(x_for_plot), linestyle=linestyle, color=color, linewidth=linewidth)


def generate_graph(ax, filtered_data, rotation, sales=True, avg_ticket=True, deferred=True):
    sales_color = 'C0'  
    deferred_color = 'C1'
    avg_ticket_color = 'C2'

    if(sales or deferred):
        ax.set_ylabel('$', color=sales_color)
    if(sales):
        ax.plot(filtered_data['Month'], filtered_data['Sales'], label='Sales', color=sales_color, marker='o')
        add_trendline(ax, filtered_data['Month'], filtered_data['Sales'], color=sales_color)
    if(deferred):
        ax.plot(filtered_data['Month'], filtered_data['Deferred'], label='Deferred', color=deferred_color, marker='x')
        # ax.tick_params(axis='y', labelcolor=sales_color)
        add_trendline(ax, filtered_data['Month'], filtered_data['Deferred'], color=deferred_color)
    
    if(avg_ticket):
        axa = ax.twinx()
        axa.set_ylabel('AvgTicket', color=avg_ticket_color)
        line_avg_ticket, = axa.plot(filtered_data['Month'], filtered_data['AvgTicket'], label='AvgTicket', color=avg_ticket_color, marker='o')
        axa.tick_params(axis='y', labelcolor=avg_ticket_color)
        add_trendline(axa, filtered_data['Month'], filtered_data['AvgTicket'], color=avg_ticket_color)

    

    plt.setp(ax.get_xticklabels(), rotation=rotation, ha='center')
    ax.xaxis.set_major_locator(mdates.MonthLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
    ax.set_xlim(filtered_data['Month'].min(), shop_data['Month'].max())
   
    
    # # Including the line for AvgTicket
    lines, labels = ax.get_legend_handles_labels()
    if(avg_ticket):
        lines += [line_avg_ticket]
        labels += [line_avg_ticket.get_label()]
    ax.legend(lines, labels, loc='upper left', bbox_to_anchor=(1.05, 1), borderaxespad=0.)


def calculate_monthly_averages(df, shop_name):
    shop_data = df[df['ShopName'] == shop_name]
    shop_data['Month'] = shop_data['Month'].dt.to_timestamp()
    monthly_avg = shop_data.groupby(shop_data['Month'].dt.month_name()).mean().reset_index()
    return monthly_avg



# Generate a plot for each shop
for shop_name in df['ShopName'].unique():
    shop_data = df[df['ShopName'] == shop_name].sort_values('Month')    
    shop_data['Month'] = shop_data['Month'].dt.to_timestamp()
    shop_data['Year'] = shop_data['Month'].dt.year
    shop_data['MonthName'] = shop_data['Month'].dt.month_name()  

    shop_data_2022 = shop_data[shop_data['Month'] >= '2022-01-01']
    shop_data_2023 = shop_data[shop_data['Month'] >= '2023-01-01']

    # Calculate average sales for each month across all years
    monthly_avg_sales = shop_data.groupby('MonthName')['Sales'].mean()
    monthly_avg_dict = monthly_avg_sales.to_dict()

    shop_grouped = shop_data.groupby(['Year', 'MonthName']).sum().reset_index()


    fig = plt.figure(figsize=(20, 12))

    ax1 = plt.subplot2grid((3, 2), (0, 0), colspan=2, fig=fig)
    ax2 = plt.subplot2grid((3, 2), (1, 0), fig=fig)
    ax3 = plt.subplot2grid((3, 2), (1, 1), fig=fig)
    ax4 = plt.subplot2grid((3, 2), (2, 0), colspan=2, fig=fig)
    generate_graph(ax1,shop_data_2023,0,sales=True, avg_ticket=False, deferred=False)
    generate_graph(ax1,shop_data_2023,0,sales=False, avg_ticket=True, deferred=False)
    generate_graph(ax1,shop_data_2023,0,sales=True, avg_ticket=False, deferred=False)
    generate_graph(ax2,shop_data_2022,90)
    generate_graph(ax3,shop_data,90)
   
    
    months_order = list(calendar.month_name[1:])  # This will give a list of month names in order

    sns.barplot(data=shop_grouped, x='MonthName', y='Sales', hue='Year', ax=ax4, order=months_order)
    
    avg_sales_ordered = [monthly_avg_dict.get(month, 0) for month in months_order]
    avg_line, = ax4.plot(months_order, avg_sales_ordered, color='C3', marker='o', linestyle='-', linewidth=2, label='Average Sales')

    # Get handles and labels for the existing legend from the bar plot
    handles, labels = ax4.get_legend_handles_labels()

    # Check if 'Average Sales' is already in labels, if not add it
    if 'Average Sales' not in labels:
        handles.append(avg_line)
        labels.append('Average Sales')

    # Create the legend
    # ax4.legend(handles, labels)
    ax4.legend(handles, labels, loc='upper left', bbox_to_anchor=(1.05, 1), borderaxespad=0.)


    ax4.set_title(f'Monthly Sales Comparison for {shop_name}')
    ax4.set_ylabel('Total Sales')
    ax4.set_xlabel('Month')


    ax1.set_title(f'YTD {shop_name}')
    ax2.set_title(f'2022 - now for {shop_name}')
    ax3.set_title(f'2021 - now for {shop_name}')

    plt.tight_layout()
    plt.subplots_adjust(top=0.93)

    # Save the figure to the PDF
    pdf_pages.savefig(fig, bbox_inches='tight')
    plt.close(fig)

pdf_pages.close()
