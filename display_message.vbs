Const FilePath = "C:\Users\user\git\excel_vba\display_message.xlsm"
Const PROC_NAME="display_message"
Dim app
Set app = CreateObject("Excel.Application")
With app
	.Visible=False
	Dim wb
	Set wb=.Workbooks.Open(FilePath)
	.Run wb.Name &"!"&PROC_NAME
	.DisplayAlerts = False
	wb.Save
	wb.Close
End With