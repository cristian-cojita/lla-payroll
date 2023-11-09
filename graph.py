import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.ticker import MaxNLocator
import numpy as np
import datetime
import matplotlib.dates as mdates
from matplotlib.ticker import FuncFormatter

df = pd.read_excel('Sales01.xlsx')

# Convert 'Month' to datetime for proper sorting in plots
df['Month'] = pd.to_datetime(df['Month']).dt.to_period('M')

pdf_path = 'shop_sales_report.pdf'
pdf_pages = PdfPages(pdf_path)


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


def generate_graph(ax, filtered_data, rotation):
    color = 'tab:red'
    ax.set_ylabel('Sales and Deferred', color=color)
    ax.plot(filtered_data['Month'], filtered_data['Sales'], label='Sales', color=color, marker='o')
    ax.plot(filtered_data['Month'], filtered_data['Deferred'], label='Deferred', color='tab:green', marker='x')
    ax.tick_params(axis='y', labelcolor=color)
    add_trendline(ax, filtered_data['Month'], filtered_data['Sales'], color=color)
    add_trendline(ax, filtered_data['Month'], filtered_data['Deferred'], color='tab:green')
    
    axa = ax.twinx()
    color = 'tab:blue'
    axa.set_ylabel('AvgTicket', color=color)
    line_avg_ticket, = axa.plot(filtered_data['Month'], filtered_data['AvgTicket'], label='AvgTicket', color=color, marker='o')
    axa.tick_params(axis='y', labelcolor=color)
    add_trendline(axa, filtered_data['Month'], filtered_data['AvgTicket'], color=color)

    plt.setp(ax.get_xticklabels(), rotation=rotation, ha='center')
    ax.xaxis.set_major_locator(mdates.MonthLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
    ax.set_xlim(filtered_data['Month'].min(), shop_data['Month'].max())
   
    
    # Including the line for AvgTicket
    lines, labels = ax.get_legend_handles_labels()
    lines += [line_avg_ticket]
    labels += [line_avg_ticket.get_label()]
    ax.legend(lines, labels, loc='upper left', bbox_to_anchor=(1.05, 1), borderaxespad=0.)

# Generate a plot for each shop
for shop_name in df['ShopName'].unique():
    shop_data = df[df['ShopName'] == shop_name].sort_values('Month')    
    shop_data['Month'] = shop_data['Month'].dt.to_timestamp()
    shop_data_2022 = shop_data[shop_data['Month'] >= '2022-01-01']
    shop_data_2023 = shop_data[shop_data['Month'] >= '2023-01-01']

    fig = plt.figure(figsize=(20, 12))

    ax1 = plt.subplot2grid((2, 2), (0, 0), colspan=2, fig=fig)
    generate_graph(ax1,shop_data_2023,0)

    ax2 = plt.subplot2grid((2, 2), (1, 0), fig=fig)
    generate_graph(ax2,shop_data_2022,90)

    ax3 = plt.subplot2grid((2, 2), (1, 1), fig=fig)
    generate_graph(ax3,shop_data,90)
    
    ax1.set_title(f'YTD {shop_name}')
    ax2.set_title(f'2022 - now for {shop_name}')
    ax3.set_title(f'2021 - now for {shop_name}')

    plt.tight_layout()
    # ax1.legend(loc='upper left', bbox_to_anchor=(1.05, 1), borderaxespad=0.)

    # Save the figure to the PDF
    pdf_pages.savefig(fig, bbox_inches='tight')
    plt.close(fig)

pdf_pages.close()
