import pandas as pd
import statsmodels.api as sm
import os

# Load the CSV file
file_path = r"C:\Users\Tom\OneDrive - Motor Depot\Reports\Sales\Regression data\Regression_Input.csv"
data = pd.read_csv(file_path)

# Clean the currency columns ('Sell', 'AT Price', 'CAP Retail Live') by removing £ and commas
currency_cols = ['Final Sell', 'AT Price', 'CAP Retail Live']
for col in currency_cols:
    data[col] = data[col].replace({'£': '', ',': ''}, regex=True).astype(float)

# Ensure 'Rating' and 'DIS' are numeric
data['AT Rating'] = pd.to_numeric(data['AT Rating'], errors='coerce')
data['DIS'] = pd.to_numeric(data['DIS'], errors='coerce')

# Separate data for each Fuel type
fuel_types = data['Fuel'].unique()

# Prepare a DataFrame to hold all the regression coefficients for each fuel type
regression_results = []

for fuel in fuel_types:
    # Filter data for the current fuel type
    fuel_data = data[data['Fuel'] == fuel]

    # Dependent variable: 'Sell'
    y = fuel_data['Final Sell']

    # Independent variables
    X = fuel_data[['AT Price', 'AT Rating', 'DIS', 'CAP Retail Live']]
    X = sm.add_constant(X)  # Adds a constant term to the predictors

    # Perform regression
    model = sm.OLS(y, X, missing='drop').fit()

    # Store the regression coefficients
    coefs = model.params
    regression_results.append({
        'Fuel': fuel,
        'const': coefs['const'],
        'AT Price': coefs['AT Price'],
        'AT Rating': coefs['AT Rating'],
        'DIS': coefs['DIS'],
        'CAP Retail Live': coefs['CAP Retail Live']
    })

# Convert the list of results to a DataFrame
regression_df = pd.DataFrame(regression_results)

# Output file path
output_file = os.path.join(os.path.dirname(file_path), 'regression_results.csv')

# Save to CSV
regression_df.to_csv(output_file, index=False)

print(f"Regression results saved to: {output_file}")
