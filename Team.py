from flask import Flask, render_template
import openpyxl
from jinja2 import Template

app = Flask(__name__)

@app.route('/')
def generate_team_page():
    # Load the Excel sheet
    workbook = openpyxl.load_workbook('Team_data.xlsx')
    worksheet = workbook.active

    # Extract the data into a list of dictionaries
    members = []
    for row in worksheet.iter_rows(min_row=2, values_only=True):
        member = {
            'name': row[0],
            'position': row[1],
            'email': row[2],
            'photo': row[3]
        }
        members.append(member)

    # Load the HTML template and render it with the data
    with open('team_page.html') as f:
        template = Template(f.read())
        output = template.render(members=members)

    return output

if __name__ == '__main__':
    app.run()
