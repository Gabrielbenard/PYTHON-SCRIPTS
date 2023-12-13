import win32com.client as w32


outlook = w32.Dispatch('outlook.application')

emailOutlook = outlook.CreateItem(0)
emailOutlook.To = "ana@gmail.com"
emailOutlook.Subject = "Meio com python e outlook"
emailOutlook.HTMLBody = """
<p> Boa noite <b> ana</b> </p>
<p> Email de Teste</p>

"""

emailOutlook.save() #save no draft