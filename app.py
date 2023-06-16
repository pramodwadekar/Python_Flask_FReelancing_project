import pandas as pd
import openpyxl
from flask import Flask,render_template,request,session,redirect,url_for
data = pd.read_excel('manager_team.xlsx')
data1 = pd.read_excel('sales_team.xlsx')

app = Flask(__name__)
app.secret_key = "super secret key"

@app.route('/')
def log():
    return render_template('login.html')

@app.route('/form_login', methods =['POST','GET'] )
def login():
    name1 = request.form['username']
    pwd = request.form['password'] 

    df = pd.read_excel('sales_team.xlsx')

   
    if name1 =="Admin" and pwd =="ABCD":
        return render_template('manager.html', info = "lnvalid User")
    elif name1 =="Team" and pwd =="EFGH":
        return render_template('contacts.html', contacts=df.to_dict('records'))
    else:
            return render_template('login.html')
    
    
@app.route('/form_manager', methods =['post','GET'] )
def manager():
    
    if request.form.get('p') and request.form.get('q'): 
        return render_template('manager.html')
    elif request.form.get('p'):
        return render_template('button.html')
    elif request.form.get('q'):
        return render_template('button1.html')
    else:
        return render_template('login.html')

@app.route('/form_add', methods =['post','GET'] )
def add():
    
    if request.form.get('add'):
        return render_template('entry_form.html')
    elif request.form.get('view'):
        return render_template('home.html')
    elif request.form.get('search'):
        return render_template('search.html')
    else:
        return render_template('manager.html')
    

@app.route('/entry_form', methods =['post','GET'] )
def submit_form():
    # Get the form data
    name = request.form['Full Name']
    Phone = request.form['Phone No']
    Dmat = request.form['Dmat']
    real_estate = request.form['Real Estate']
    Insurance = request.form['Insurance']
    date = request.form['Date']
    feedback = request.form['Feedback']

    data = pd.read_excel('manager_team.xlsx')
    # Create a Pandas dataframe with the form data
    new_row = {
        'Full Name': name,
        'Phone No': Phone,
        'Dmat': Dmat,
        'Real Estate': real_estate,
        'Insurance': Insurance,
        'Date': date,
        'Feedback' : feedback
    }
    data = data.append(new_row, ignore_index=True)
    
    # Write the dataframe to an Excel file
    data.to_excel('manager_team.xlsx', index=False)
    
    return render_template('button.html')

@app.route('/home', methods=['POST'])
def filter():
    data = pd.read_excel('manager_team.xlsx')
    data.reset_index(inplace=True)
    date = request.form['date']
    start = int(request.form.get('start', 0))
    filtered_data = data[data['Date'] == date].iloc[start:start+10]
    return render_template('filtered.html', contacts=filtered_data.to_dict('records'), start=start)

@app.route('/delete', methods=['POST'])
def delete():
    # Retrieve the row index from the form data
    row_index = int(request.form['row_index'])
    # Read the Excel file into a Pandas DataFrame
    data = pd.read_excel('manager_team.xlsx')
    
    # Delete the selected row from the DataFrame
    data.drop(index = row_index -1, inplace=True)
    # Write the updated DataFrame to the Excel file
    data.to_excel('manager_team.xlsx', index=False)
    # Redirect the user back to the form
    return 'Row deleted successfully.'

@app.route('/previous', methods=['POST'])
def previous():
    data = pd.read_excel('manager_team.xlsx')
    date = request.form['date']
    start = int(request.form.get('start', 0))
    filtered_data = data[data['Date'] == date].iloc[start-10:start]
    return render_template('filtered.html', contacts=filtered_data.to_dict('records'), start=start-10)
@app.route('/homes')
def logout():
    # Implement your logout logic here, for example, clearing the user's session
    session.clear()
    return redirect(url_for('filter'))

    
# this code is Searching data
@app.route('/search' , methods =['post'])
def search():
    data = pd.read_excel('manager_team.xlsx')
    data.reset_index(inplace=True)
    search = request.form['name']
    start = int(request.form.get('start', 0))
    filtered_data = data[data['Full Name'] == search].iloc[start:start+5]
    return render_template('filtered_search.html', contacts=filtered_data.to_dict('records'), start=start)

@app.route('/search_delete', methods=['POST'])
def search_delete():
    # Retrieve the row index from the form data
    row_index = int(request.form['row_index'])
    # Read the Excel file into a Pandas DataFrame
    data = pd.read_excel('manager_team.xlsx')
    
    # Delete the selected row from the DataFrame
    data.drop(index = row_index -1, inplace=True)
    # Write the updated DataFrame to the Excel file
    data.to_excel('manager_team.xlsx', index=False)
    # Redirect the user back to the form
    return 'Row deleted successfully.'

@app.route('/nexts', methods=['POST'])
def nexts():
    data = pd.read_excel('manager_team.xlsx')
    date = request.form['name']
    start = int(request.form['start'])
    filtered_data = data[data['Full Name'] == date].iloc[start:start+5]
    return render_template('next_search.html', contacts=filtered_data.to_list(), start=start+5)

@app.route('/previouss', methods=['POST'])
def previouss():
    data = pd.read_excel('manager_team.xlsx')
    date = request.form['name']
    start = int(request.form.get('start', 0))
    filtered_data = data[data['Full Name'] == date].iloc[start-5:start]
    return render_template('filtered_search.html', contacts=filtered_data.to_dict('records'), start=start-5)
  
@app.route('/logout')
def logouts():
    # Implement your logout logic here, for example, clearing the user's session
    session.clear()
    return redirect(url_for('search'))
#sales team checkbox

@app.route('/form_add1', methods =['post','GET'] )
def adds():
    if request.form.get('add'):
        return render_template('entry_form1.html')
    elif request.form.get('view'):
        return render_template('home1.html')
    elif request.form.get('search'):
        return render_template('search1.html')
    else:
        return render_template('manager.html')
    
@app.route('/entry_form1', methods =['post','GET'] )
def submit_forms():
    # Get the form data
    name = request.form['Full Name']
    Phone = request.form['Phone No']
    Dmat = request.form['Dmat']
    real_estate = request.form['Real Estate']
    Insurance = request.form['Insurance']
    date = request.form['Date']
    feedback = request.form['Feedback']
    
    # Create a Pandas dataframe with the form data
    data1 = pd.read_excel('sales_team.xlsx')
    new_row = {
        'Full Name': name,
        'Phone No': Phone,
        'Dmat': Dmat,
        'Real Estate': real_estate,
        'Insurance': Insurance,
        'Date': date,
        'Feedback' : feedback
    }
    data1 = data1.append(new_row, ignore_index=True)
    
    # Write the dataframe to an Excel file
    data1.to_excel('sales_team.xlsx', index=False)
    
    return render_template('button1.html')
    
@app.route('/sales_home', methods=['POST'])
def filters():
    data1 = pd.read_excel('sales_team.xlsx')
    data1.reset_index(inplace=True)
    date = request.form['date']
    start = int(request.form.get('start', 0))
    filtered_data = data1[data1['Date'] == date].iloc[start:start+10]
    return render_template('sales_filtered.html', contacts=filtered_data.to_dict('records'), start=start)

@app.route('/sales_delete', methods=['POST'])
def sales_delete():
    # Retrieve the row index from the form data
    row_index = int(request.form['row_index'])
    # Read the Excel file into a Pandas DataFrame
    data1 = pd.read_excel('sales_team.xlsx')
    
    # Delete the selected row from the DataFrame
    data1.drop(index = row_index -1, inplace=True)
    # Write the updated DataFrame to the Excel file
    data1.to_excel('sales_team.xlsx', index=False)
    # Redirect the user back to the form
    return 'Row deleted successfully.'



@app.route('/sales_next', methods=['POST'])
def sales_next():
    data1 = pd.read_excel('sales_team.xlsx')
    date = request.form['date']
    start = int(request.form['start'])
    filtered_data = data1[data1['Date'] == date].iloc[start:start+5]
    return render_template('sales_next.html', contacts=filtered_data.to_dict('records'), start=start+5)

@app.route('/sales_previous', methods=['POST'])
def sales_previous():
    data1 = pd.read_excel('sales_team.xlsx')
    date = request.form['date']
    start = int(request.form.get('start', 0))
    filtered_data = data1[data1['Date'] == date].iloc[start-10:start]
    return render_template('sales_filtered.html', contacts=filtered_data.to_dict('records'), start=start-5)

@app.route('/search1' , methods =['post'])
def sales_search():
    data1 = pd.read_excel('sales_team.xlsx')
    data1.reset_index(inplace=True)
    search = request.form['name']
    start = int(request.form.get('start', 0))
    filtered_data = data1[data1['Full Name'] == search].iloc[start:start+5]
    return render_template('sales_filtered_search.html', contacts=filtered_data.to_dict('records'), start=start)   

@app.route('/sales_search_delete', methods=['POST'])
def sales_search_delete():
    # Retrieve the row index from the form data
    row_index = int(request.form['row_index'])
    # Read the Excel file into a Pandas DataFrame
    data1 = pd.read_excel('sales_team.xlsx')
    
    # Delete the selected row from the DataFrame
    data1.drop(index = row_index -1, inplace=True)
    # Write the updated DataFrame to the Excel file
    data1.to_excel('sales_team.xlsx', index=False)
    # Redirect the user back to the form
    return 'Row deleted successfully.'

@app.route('/sales_nexts', methods=['POST'])
def sales_nexts():
    data1 = pd.read_excel('sales_team.xlsx')
    date = request.form['name']
    start = int(request.form['start'])
    filtered_data = data1[data1['Full Name'] == date].iloc[start:start+5]
    return render_template('sales_next_search.html', contacts=filtered_data.to_dict('records'), start=start+5)

@app.route('/sales_previouss', methods=['POST'])
def sales_previouss():
    data1 = pd.read_excel('sales_team.xlsx')
    date = request.form['name']
    start = int(request.form.get('start', 0))
    filtered_data = data1[data1['Full Name'] == date].iloc[start-5:start]
    return render_template('sales_filtered_search.html', contacts=filtered_data.to_dict('records'), start=start-5)
  


if __name__ == "__main__":
    app.run(debug = True)