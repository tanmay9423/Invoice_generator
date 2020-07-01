import gspread
from oauth2client.service_account import ServiceAccountCredentials
from PIL import Image, ImageDraw, ImageFont
from string import Template

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('invoice.json',scope)
client = gspread.authorize(creds)

sh = client.open ('invoice data')
sheet = sh.get_worksheet(0)

def read_template(filename):
    with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)

values_list = sheet.col_values(1)

run = True
while run:
    for row_count in range(len(values_list)-1):
        row_count = row_count+ 2
        status = str(sheet.cell(row_count,11).value)
        if (status != "sent"):
            date = sheet.cell(row_count, 1).value
            name = sheet.cell(row_count, 2).value
            phno = sheet.cell(row_count,3).value
            clss = sheet.cell(row_count, 4).value
            descrp = sheet.cell(row_count, 5).value
            email = sheet.cell(row_count, 6).value
            dop = sheet.cell(row_count, 7).value
            delay = sheet.cell(row_count, 8).value
            amount = str(sheet.cell(row_count, 9).value)
            quantity = sheet.cell(row_count, 10).value
            reg_num = str(row_count - 1)
    
            base = Image.open('Invoice Template.png')
            draw = ImageDraw.Draw(base)
            dscp_font = ImageFont.truetype('arial.ttf', size=28)
            reg_font = ImageFont.truetype('arial.ttf', size=25)
            name_font = ImageFont.truetype('arial.ttf', size=26)
            colour = 'rgb(0,0,0)'
    
            draw.text((680,255), name, fill=colour, font=name_font)
            draw.text((680,288), email, fill=colour, font=name_font)
            draw.text((680,319), clss, fill=colour, font=name_font)
            draw.text((680,350), phno, fill=colour, font=name_font)
            draw.text((166,498), '#'+reg_num, fill=colour, font=reg_font)
            draw.text((124,529), dop, fill=colour, font=reg_font)
            draw.text((61,680), descrp, fill=colour, font=dscp_font)
            draw.text((931,680), quantity, fill=colour, font=dscp_font)
            draw.text((1010,680), "₹"+amount, fill=colour, font=dscp_font)
            draw.text((990,990), "₹"+amount, fill=colour, font=ImageFont.truetype('arial.ttf', size=35))
            draw.text((817,680), delay, fill=colour, font=dscp_font)
    
            base.save("#"+reg_num+'.png')
            message_template = read_template('greeting_template.txt')

            s = smtplib.SMTP(host='smtp.gmail.com', port=587)
            s.starttls()
            s.login("your email" , "password")             
            msg = MIMEMultipart()
            message = message_template.substitute(name=name.title())
            msg['From']="your email"
            msg['To']=email
            msg['Subject']="Invoice"
            msg.attach(MIMEText(message, 'plain'))
            filename = reg_num+".png"
            attachment = open("#"+reg_num+".png", "rb")
            p = MIMEBase('application', 'octet-stream')
            p.set_payload((attachment).read())
            encoders.encode_base64(p)
            p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
            msg.attach(p)
            s.send_message(msg)
            del msg
            s.quit()
            sheet.update_cell(row_count,11,"sent")
        run = False
