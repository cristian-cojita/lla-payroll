import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.ticker import MaxNLocator
import numpy as np
import datetime
import matplotlib.dates as mdates
from matplotlib.ticker import FuncFormatter

# Load the Excel data
df = pd.read_excel('Sales01.xlsx')

# Convert 'Month' to datetime for proper sorting in plots
df['Month'] = pd.to_datetime(df['Month']).dt.to_period('M')

# Set up a PDF to save the plots
pdf_path = 'shop_sales_report.pdf'
pdf_pages = PdfPages(pdf_path)

def custom_month_formatter(x, pos=None):
    date = mdates.num2date(x)
    if date.month == 1:  # If January
        return date.strftime('%Y')  # Return the full year
    else:
        return date.strftime('%Y-%m')  # Return the abbreviated month name

def add_trendline(ax, dates, y, order=1, color='r', linestyle='--', linewidth=2):
    # Ensure that dates and y are numpy arrays
    dates = np.array(dates)
    y = np.array(y)

    # Filter out any NaN values
    mask = ~np.isnan(y)
    dates = dates[mask]
    y = y[mask]

    # Only proceed if there are enough data points
    if len(y) > order:
        # Convert dates to ordinal numbers for trendline calculation
        x = np.array([mdates.date2num(date) for date in dates])

        # Fit the polynomial
        coeffs = np.polyfit(x, y, order)
        trend = np.poly1d(coeffs)

        # Generate x values from the date range for plotting the trendline
        x_for_plot = np.linspace(x.min(), x.max(), len(x))

        # Plot the trendline
        ax.plot(mdates.num2date(x_for_plot), trend(x_for_plot), linestyle=linestyle, color=color, linewidth=linewidth)

# Generate a plot for each shop
for shop_name in df['ShopName'].unique():
    shop_data = df[df['ShopName'] == shop_name].sort_values('Month')
    shop_data['Month'] = shop_data['Month'].dt.to_timestamp()

    fig, ax1 = plt.subplots(figsize=(20, 6))
    color = 'tab:red'
    ax1.set_ylabel('Sales and Deferred', color=color)
    ax1.plot(shop_data['Month'], shop_data['Sales'], label='Sales', color=color, marker='o')
    ax1.plot(shop_data['Month'], shop_data['Deferred'], label='Deferred', color='tab:green', marker='x')



    ax1.tick_params(axis='y', labelcolor=color)
    add_trendline(ax1, shop_data['Month'], shop_data['Sales'], color=color)
    add_trendline(ax1, shop_data['Month'], shop_data['Deferred'], color='tab:green')
    
    ax2 = ax1.twinx()
    color = 'tab:blue'
    ax2.set_ylabel('AvgTicket', color=color)
    ax2.plot(shop_data['Month'], shop_data['AvgTicket'], label='AvgTicket', color=color, marker='o')
    ax2.tick_params(axis='y', labelcolor=color)
    add_trendline(ax2, shop_data['Month'], shop_data['AvgTicket'], color=color)

    plt.setp(ax1.get_xticklabels(), rotation=90, ha='center')
    ax1.xaxis.set_major_locator(mdates.MonthLocator())
    ax1.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
    ax1.set_xlim(shop_data['Month'].min(), shop_data['Month'].max())

    plt.title(f'Sales Data for {shop_name}')
    fig.tight_layout()
    fig.legend(loc='upper left', bbox_to_anchor=(0.1, 0.9))
    pdf_pages.savefig(fig)
    plt.close(fig)

pdf_pages.close()
