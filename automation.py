import win32com.client as win32

olApp = win32.Dispatch('Outlook.Application')
olNS = olApp.GetNamespace('MAPI')

mailItem = olApp.CreateItem(0)
mailItem.subject = 'Testsub'
mailItem.BodyFormat = 1
mailItem.Body = "Hello There"
mailItem.To = 'vigneshvijay15@gmail.com'

mailItem.Display()
mailItem.Save()
mailItem.Send()
