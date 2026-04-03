import win32com.client
import pythoncom
import traceback

def test():
    pythoncom.CoInitialize()
    try:
        print("Connecting to Outlook...")
        outlook = win32com.client.Dispatch("Outlook.Application")
        print("Connected.")
        mail = outlook.CreateItem(0)
        mail.To = "test@example.com"
        mail.Subject = "Test Email"
        mail.HTMLBody = "<html><body>Test</body></html>"
        print("Created item. Displaying...")
        mail.Display()
        print("Displayed successfully!")
    except Exception as exc:
        print("Error:")
        traceback.print_exc()
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    test()
