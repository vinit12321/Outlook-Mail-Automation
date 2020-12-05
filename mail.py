import win32com.client as client
from jinja2 import Template
outlook = client.Dispatch('Outlook.Application')
message = outlook.CreateItem(0) # 0 is the code for a mail item (see the enumerations)
message.To = 'vinitdabake@outlook.com'
message.Subject = 'Redshift Details'
message.BodyFormat = 2
abc='Sumi'
username='etl_admin'
password='Abcd#2321'
html_bdy='''
<html>
    <head>

    </head>
    <body>
    Hi {{abc}} ,
<br/>
<br/>
     <table border="1">
       
     <tr>
         <th colspan="2" style="background-color: #00a2ed;color: black;">Redshift details</th>
     </tr>
        <tr>
            <td>
                Database server name :
            </td>
            <td>
                intlapadskdsndsdsndsndSKSAHDJADAKDA
            </td>
        </tr>
        <tr>
            <td>
                Port:
            </td>
            <td>
                5439
            </td>
        </tr>
        <tr>
            <td>
                Database Name:
            </td>
            <td>
                Prod
            </td>
        </tr>
        <tr>
            <td>
                User:
            </td>
            <td>
                {{username}}
            </td>
        </tr>
        <tr>
            <td>
                Password:
            </td>
            <td>
                {{password}}
            </td>
        </tr>

        
    
     </table>
<br/>
<br/>
Thanks & Regards, <br/>
Vinit
    </body>

</html>

'''
tm = Template(html_bdy)
msg=tm.render(abc=abc,username=username,password=password)
print(msg)
message.HTMLBody = msg
message.Display()
#message.Save() # save to drafts folder
#message.Send() # send to outbox
