from flask import Flask, render_template, request, redirect, url_for, Response
import pandas as pd
import io

app = Flask(__name__)

# Read data from Excel file into a Pandas DataFrame
excel_file = "items.xlsx"
data = pd.read_excel(excel_file)

@app.route('/')
def index():
    # Group data by category
    grouped_data = data.groupby('Category')
    categories = []
    overall_counter = 0  # Initialize an overall counter
    for category, group in grouped_data:
        category_items = []
        for item in group.to_dict(orient='records'):
            item['overall_index'] = overall_counter
            overall_counter += 1
            category_items.append(item)
        categories.append({
            'category': category,
            'items': category_items
        })
    return render_template('index.html', categories=categories)

@app.route('/update', methods=['POST'])
def update():
    item_index = int(request.form['item_index'])
    action = request.form['action']

    if action == 'increase_in':
        data.at[item_index, 'In'] += 1
    elif action == 'decrease_in' and data.at[item_index, 'In'] > 0:
        data.at[item_index, 'In'] -= 1
    elif action == 'increase_rm':
        data.at[item_index, 'RM'] += 1
    elif action == 'decrease_rm' and data.at[item_index, 'RM'] > 0:
        data.at[item_index, 'RM'] -= 1
    elif action == 'increase_re':
        data.at[item_index, 'RE'] += 1
    elif action == 'decrease_re' and data.at[item_index, 'RE'] > 0:
        data.at[item_index, 'RE'] -= 1

    # Save the updated data to the Excel file
    # data.to_excel(excel_file, index=False)

    return redirect(url_for('index'))

@app.route('/reset-data', methods=['POST'])
def reset_data():
    global data  # Make sure 'data' is accessible globally
    data['In'] = 0
    data['RM'] = 0
    data['RE'] = 0

    # Save the updated data to the Excel file
    # data.to_excel(excel_file, index=False)

    return redirect(url_for('index'))

@app.route('/export-excel')
def export_excel():
    # Create a buffer for the Excel file
    excel_buffer = io.BytesIO()

    # Create a Pandas Excel writer using openpyxl as the engine
    excel_writer = pd.ExcelWriter(excel_buffer, engine='openpyxl')

    # Write the data to the Excel file
    data.to_excel(excel_writer, sheet_name='Sheet1', index=False)

    # Save the Pandas Excel writer to the buffer
    # excel_writer.save()
    excel_writer.close()
    excel_buffer.seek(0)

    # Return the Excel file as a response
    response = Response(excel_buffer, content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response.headers['Content-Disposition'] = 'attachment; filename=exported_data.xlsx'

    return response

if __name__ == '__main__':
    app.run(debug=True)
