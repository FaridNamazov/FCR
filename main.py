from flask import Flask, render_template, jsonify
import pandas as pd

app = Flask(__name__)


# Define the route for the dashboard page
@app.route('/dashboard')
def dashboard():
    return render_template('dashboard.html')


# Define the API endpoint to fetch data for a specific year
@app.route('/api/<int:year>')
def get_data(year):
    # Replace 'G:\Customer Service Center\Statistics & Dashboard\Monthly KPI\' with the actual path to your Excel files
    filepath = r'G:\Customer Service Center\Statistics & Dashboard\Monthly KPI\Monthly KPI for {}.xlsx'.format(year)

    # Read the data from the 'Call Center' sheet
    df_call_center = pd.read_excel(filepath, sheet_name='Call Center')

    # Read the data from the 'Premium' sheet
    df_premium = pd.read_excel(filepath, sheet_name='Premium')

    # Read the data from the 'Online Support' sheet
    df_online_support = pd.read_excel(filepath, sheet_name='Online Support')

    # Process the data and create a dictionary or JSON object with the required information
    # Here, we are simply converting the data to JSON format as an example
    data = {
        'call_center_data': df_call_center.to_dict(orient='records'),
        'premium_data': df_premium.to_dict(orient='records'),
        'online_support_data': df_online_support.to_dict(orient='records')
    }

    return jsonify(data)


if __name__ == '__main__':
    app.run(debug=True)