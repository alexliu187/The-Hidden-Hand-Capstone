##################################################################################
## The Hidden Hand: How the Dark Web is Influencing the 2024 Presidential Race  ##
## AY25 USNA Capstone Project	                                                  ##
## Authors: MIDN 1/C Alexander Liu                                              ##
## 		      MIDN 1/C Apollos Burcham                                            ##
##          MIDN 1/C Anna Cherian                                               ##
##          MIDN 1/C Izabela Gorczynski                                         ##
## Faculty Advisor: Dr. April Edwards, Cyber Science Department, USNA           ##
##################################################################################


#Import necessary modules
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from urllib.parse import urlparse
import openpyxl
from openpyxl.drawing.image import Image
import matplotlib.dates as mdates

# Load the CSV file
data = pd.read_csv("US.csv")

# Preprocess data
data['scrapeDate'] = pd.to_datetime(data['scrapeDate'])
data['date'] = data['scrapeDate'].dt.date

# Group articles by day
articles_per_day = data.groupby('date').size()

# Extract domains from URLs
data['domain'] = data['url'].apply(lambda x: urlparse(x).netloc)
top_domains = data['domain'].value_counts().head(10)

# Count selector categories
selector_counts = data['selectors'].str.split(r'\n|,', expand=True).stack().value_counts()

sns.set(style="whitegrid")

# Plot timeline
plt.figure(figsize=(10, 5))
plt.plot(articles_per_day.index, articles_per_day.values, marker='o', color='blue')
plt.gca().xaxis.set_major_locator(mdates.DayLocator(interval=1))
plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
plt.xticks(rotation=45)
plt.title('Articles Scraped per Day')
plt.tight_layout()
plt.savefig('timeline.png')
plt.close()

# Plot selector breakdown
plt.figure(figsize=(8, 5))
sns.barplot(y=selector_counts.index, x=selector_counts.values)
plt.title('Selector Frequency')
plt.tight_layout()
plt.savefig('selectors.png')
plt.close()

# Plot top domains
plt.figure(figsize=(8, 5))
sns.barplot(y=top_domains.index, x=top_domains.values)
plt.title('Top 10 Domains')
plt.tight_layout()
plt.savefig('domains.png')
plt.close()

# Create Excel file with graphs
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Visualizations"

img1 = Image('timeline.png')
img2 = Image('selectors.png')
img3 = Image('domains.png')

ws.add_image(img1, 'A1')
ws.add_image(img2, 'A20')
ws.add_image(img3, 'A40')

wb.save("US_analysis.xlsx")
