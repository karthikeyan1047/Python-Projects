import pandas as pd
import numpy as np

start_date = pd.to_datetime('2022-01-01')
end_date = pd.to_datetime('2024-12-31')

# Generate random datetimes
region = ['North', 'South', 'East', 'West']
product = ['Product-A', 'Product-B', 'Product-C', 'Product-D', 'Product-E', 'Product-F', 'Product-G', 'Product-H', 'Product-I', 'Product-J'], 
random_dates = pd.to_datetime(np.random.uniform(start_date.value, end_date.value, 7)).round('s')
# Create DataFrame
df = pd.DataFrame(
    {
        'Region' : np.random.choice(region, 12567),
        'Product' : np.random.choice(product, 12567),
        'DateTime' : random_dates
    }
)
df.to_csv(r'C:\Users\karthikeyan.s\Desktop\random_dates.csv', index=False)
print(df)
