import email
import imp
import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import excel

#credentials
senderEmail = "joseantoniomonro33@gmail.com"
password = "qfpjnxpwmepzysad"


def sendEmail(mensaje, destino):
  
    msg = MIMEMultipart('alternative')
    msg['Subject'] = "IDK Solo Prueba"
    msg['From'] = senderEmail
    msg['To'] = destino
    
    
    msg.attach(MIMEText(mensaje, 'html'))
    
    
    attachmentPath = "D:/Ergonomic/"
    try:
      with open(attachmentPath, "rb") as attachment:
        po = MIMEApplication(attachment.read(),_subtype="pdf")
        po.add_header('Content-Disposition', "attachment; filename= %s" % attachmentPath.split("\\")[-1])
        msg.attach(po)
    except Exception as e:
      print("Error Doc")
      print(str(e))
      
    context = ssl.create_default_context()
    
    
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo
    server.starttls(context=context)
    server.login(senderEmail, password)
    print("Connected to server")
    server.sendmail(senderEmail, destino, msg.as_string())
    print("Enviado")
    server.quit()
    
destino = "019101861b@uandina.edu.pe"
mensaje = '''<div role="article" aria-roledescription="email" lang="en" style="text-size-adjust:100%;-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%;background-color:#B5B2B2;">
    <table role="presentation" style="width:100%;border:none;border-spacing:0;">
      <tr>
        <td align="center" style="padding:0;">
          <!--[if mso]>
          <table role="presentation" align="center" style="width:600px;">
          <tr>
          <td>
          <![endif]-->
          <table role="presentation" style="width:94%;max-width:600px;border:none;border-spacing:0;text-align:left;font-family:Arial,sans-serif;font-size:16px;line-height:22px;color:#093469;">
            <tr>
              <td style="text-align:center;font-size:24px;font-weight:bold;">
                <a href="https://ergonomic.com.pe/index.html" style="text-decoration:none;">
                  
                  
                  <img src="https://ergonomic.com.pe/img/banner.jpg" width="165" alt="Logo" style="width:100%;max-width:100%;height:auto;border:none;text-decoration:none;color:#F5F5F5;"></a>
              </td>
            </tr>
            <tr>
              <td style="padding:30px;background-color:#F5F5F5;">
                <h1 style="margin-top:0;margin-bottom:5px;font-size:26px;line-height:10px;font-weight:bold;letter-spacing:-0.02em;">Medical Center Ergonomic </h1>
                <h2 style="margin-top:0;margin-bottom:16px;font-size:18px;line-height:32px;font-weight:bold;letter-spacing:-0.01em;">Centro de Diagnostico y Salud Ocupacional</h2>
                <p style="margin:0;">Dr/a. Patricio Gutierres , le hacemos presente los documentos de los pacientes que estan a cargo de Bambas. Espero que sea de su agrado. No dude en contactarse con los numeros en la parte inferior por alguna duda</p>
                <ul style="font-size:16px;line-height:20px;font-weight:bold;">
                  <li>41848443- GUTIERREZ TTITO NOLBERTO - SOL DEL PACIFICO - ERGONOMIC - PRUEBA ANTIGENA - 19.09.2022</li>
                </ul>

              </td>
            </tr>
            <tr>
              <td style="padding:0;font-size:24px;line-height:28px;font-weight:bold;">
                
                <a href="" style="text-decoration:none;"><img src="https://i.postimg.cc/LXtV9qC5/undraw-medicine-b1ol.png" width="600" alt="" style="width:100%;height:auto;display:block;border:none;text-decoration:none;color:#ffff;"></a>
              </td>
            </tr>
            
            <tr>
              <td style="padding:30px;text-align:center;font-size:12px;background-color:#093469;color:#cccccc;">
                <p style="margin:0;font-size:16px;line-height:20px; font-style:bold;">Raul Quispe Apaza</p>
                <p style="margin:0;font-size:14px;line-height:20px;">Encargado de Entrega de Resultados</p>
                <p style="margin:0;font-size:14px;line-height:20px;">Medicina Ocupacional</p>

                <p style="margin:0 0 8px 0;"><a href="https://www.facebook.com/Cl%C3%ADnica-Ergonomic-185440475370641" style="text-decoration:none;"><img src="https://assets.codepen.io/210284/facebook_1.png" width="40" height="40" alt="f" style="display:inline-block;color:#cccccc;"></a> <a href="http://www.twitter.com/" style="text-decoration:none;"><img src="https://cdn-icons-png.flaticon.com/512/5968/5968841.png" width="32" height="32" alt="t" style="display:inline-block;color:#cccccc;"></a></p>
                <p style="margin:0;font-size:14px;line-height:20px;">&reg; Av. Cultura Nro 1522 -  3er Piso<br><a class="unsub" href="https://ergonomic.com.pe/index.html" style="color:#cccccc;text-decoration:underline;">Visita Nuestra Web</a></p>
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
  </div>
'''

sendEmail(mensaje, destino)