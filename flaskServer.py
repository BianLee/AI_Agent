from flask import Flask, request, render_template, jsonify
import json

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('form.html')

@app.route('/submit', methods=['POST'])
def submit():
    data = {
        'name': request.form.get('name'),
        'email': request.form.get('email'),
        'phone': request.form.get('phone'),
        'address': request.form.get('street'),
        'city': request.form.get('city'),
        'state': request.form.get('state'),
        'zip': request.form.get('zip'),
        'license': request.form.get('dl'),
        'birthday': request.form.get('birthday')
    }

    # Save the data to a JSON file
    with open('data.json', 'w') as file:
        json.dump(data, file, indent=4)
    
    return jsonify({'message': 'Data saved successfully!'})

if __name__ == '__main__':
    app.run(debug=True)

