
import smtplib
# from email.message import EmailMessage
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from dotenv import dotenv_values
from db import ConnexionOdoo
import os
from openpyxl import Workbook

cwd = os.getcwd()
config = dotenv_values(f"{cwd}/.env") 

print(config)


EMAIL_ADDRESS = config["EMAIL_ADDRESS"]
EMAIL_PASSWORD = config["EMAIL_PASSWORD"]
CC = ['Abdelhafid.BENNACI@groupe-hasnaoui.com','zahra.benali@groupe-hasnaoui.com','cheimaa.lassab@groupe-hasnaoui.com']



msg = MIMEMultipart('related')

msg['Subject'] = 'Liste des contrats'
msg['CC'] = ",".join(CC)
msg['BCC'] = 'Souheil.HADJHABIB@groupe-hasnaoui.com'
msg['From']= EMAIL_ADDRESS
msg['To'] = 'Wissem.ABDELAZIZ@groupe-hasnaoui.com'
# msg['To'] = 'odoo.test@groupe-hasnaoui.com'



employees = ConnexionOdoo.getEmployeeContracts(ConnexionOdoo,1)

print(len(employees))

wb = Workbook(write_only=True)
employees_xls = wb.create_sheet("Employees")
employees_xls.append(["ID","Employee","Date Début Contrat","Date Fin Contrat","Résponsable de l'employé","Tel du Résponsable #1","Tel du Résponsable #2"])

for employee in employees:
    employee_id = employee[0]
    employee_name = employee[1]
    employee_date_debut_contrat = employee[2]
    employee_date_fin_contrat = employee[3]
    employee_responsable = employee[4]
    tel_responsable_1 = employee[5]
    tel_responsable_2 = employee[6]
    employees_xls.append([employee_id,
                            employee_name,
                            employee_date_debut_contrat,
                            employee_date_fin_contrat,
                            employee_responsable,
                            tel_responsable_1,
                            tel_responsable_2])
wb.save("employee_contract.xlsx")

with open ("employee_contract.xlsx",'rb') as f:
    file_data = f.read()
    file_name = f.name

html_body = """
<html>
    <head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@4.0.0/dist/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous">
    <title>Email</title>
    <style type="text/css" media="screen">
    *{
    box-sizing: border-box;
    -webkit-box-sizing: border-box;
    -moz-box-sizing: border-box;
    }
    body{
    font-family: Helvetica;
    -webkit-font-smoothing: antialiased;
    background: rgba( 71, 147, 227, 1);
    }
    h2{
    text-align: center;
    font-size: 18px;
    text-transform: uppercase;
    letter-spacing: 1px;
    color: white;
    padding: 30px 0;
    }

    /* Table Styles */

    .table-wrapper{
    margin: 10px 70px 70px;
    box-shadow: 0px 35px 50px rgba( 0, 0, 0, 0.2 );
    }

    .fl-table {
    border-radius: 5px;
    font-size: 12px;
    font-weight: normal;
    border: none;
    border-collapse: collapse;
    width: 100%;
    max-width: 100%;
    white-space: nowrap;
    background-color: white;
    }

    .fl-table td, .fl-table th {
    text-align: center;
    padding: 8px;
    }

    .fl-table td {
    border-right: 1px solid #f8f8f8;
    font-size: 12px;
    }

    .fl-table thead th {
    color: #ffffff;
    background: #4FC3A1;
    }


    .fl-table thead th:nth-child(odd) {
    color: #ffffff;
    background: #324960;
    }

    .fl-table tr:nth-child(even) {
    background: #F8F8F8;
    }

    /* Responsive */

    @media (max-width: 767px) {
    .fl-table {
    display: block;
    width: 100%;
    }
    .table-wrapper:before{
    content: "Scroll horizontally >";
    display: block;
    text-align: right;
    font-size: 11px;
    color: white;
    padding: 0 0 10px;
    }
    .fl-table thead, .fl-table tbody, .fl-table thead th {
    display: block;
    }
    .fl-table thead th:last-child{
    border-bottom: none;
    }
    .fl-table thead {
    float: left;
    }
    .fl-table tbody {
    width: auto;
    position: relative;
    overflow-x: auto;
    }
    .fl-table td, .fl-table th {
    padding: 20px .625em .625em .625em;
    height: 60px;
    vertical-align: middle;
    box-sizing: border-box;
    overflow-x: hidden;
    overflow-y: auto;
    width: 120px;
    font-size: 13px;
    text-overflow: ellipsis;
    }
    .fl-table thead th {
    text-align: left;
    border-bottom: 1px solid #f7f7f9;
    }
    .fl-table tbody tr {
    display: table-cell;
    }
    .fl-table tbody tr:nth-child(odd) {
    background: none;
    }
    .fl-table tr:nth-child(even) {
    background: transparent;
    }
    .fl-table tr td:nth-child(odd) {
    background: #F8F8F8;
    border-right: 1px solid #E6E4E4;
    }
    .fl-table tr td:nth-child(even) {
    border-right: 1px solid #E6E4E4;
    }
    .fl-table tbody td {
    display: block;
    text-align: center;
    }
    }

    </style>
    </head>
    <body>
        Bonjour,
            <br /><br />

            Ci-dessous Vous trouverez la liste des employées dont leurs contrats s'achèveront le mois prochain.
            <br /><br />
                                
        <div class="table-wrapper">
            <table class="fl-table">
                <thead>
                    <tr>
                    <th style="color: #ffffff;background: #324960">ID</th>
                    <th style="border-right: 1px solid #E6E4E4;">Employee</th>
                    <th style="color: #ffffff;background: #324960">Date Début Contrat</th>
                    <th style="background: border-right: 1px solid #E6E4E4;">Date Fin Contrat</th>
                    <th style="color: #ffffff;background: #324960">Résponsable de l'employé</th>
                    <th style="background: border-right: 1px solid #E6E4E4;">Tel du Résponsable #1</th>
                    <th style="color: #ffffff;background: #324960">Tel du Résponsable #2</th>
                    </tr>
                </thead>
                <tbody>"""

msgHtml = MIMEText(html_body, 'html')

i=0
classe = ""
for employee in employees:
    if(i%2!=0):
        classe = "background: #F8F8F8;"
    else:
        classe=""
    html_body = html_body + """
    <tr>
                        <td style='"""+classe+"""'>"""+str(employee[0])+"""</td>
                        <td style='"""+classe+"""'>"""+str(employee[1])+"""</td>
                        <td style='"""+classe+"""'>"""+str(employee[2])+"""</td>
                        <td style='"""+classe+"""'>"""+str(employee[3])+"""</td>
                        <td style='"""+classe+"""'>"""+str(employee[4])+"""</td>
                        <td style='"""+classe+"""'>"""+str(employee[5])+"""</td>
                        <td style='"""+classe+"""'>"""+str(employee[6])+"""</td>
    </tr>"""
    i+=1

html_body = html_body +"""<tbody>
            </table>
        </div>
    </body>
</html>"""

msg.attach(MIMEText(html_body, 'html'))

part = MIMEBase('application', "octet-stream")
part.set_payload(open("employee_contract.xlsx", "rb").read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', 'attachment; filename="employee_contract.xlsx"')
msg.attach(part)

with smtplib.SMTP('smtp.groupe-hasnaoui.com',587) as smtp:
    smtp.ehlo()
    smtp.starttls()
    smtp.ehlo()
    print("Logging in")
    # smtp.login('cbi@groupe-hasnaoui.com', 'Cb1gsh19')
    smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
    print("Logged in")
    smtp.send_message(msg)
