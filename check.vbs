'Fecha todas as sessoes do IE e do WORD
Dim oShell : Set oShell = CreateObject("WScript.Shell")
oShell.Run "taskkill /im iexplore.exe", , True
oShell.Run "taskkill /im winword.exe", , True

WScript.Sleep 5000

'Cria objeto word
Set objWord = CreateObject("Word.Application")
objWord.Visible = False
objWord.Documents.Open "C:\workshop\vbs\checkpages\modelo\modelo.doc"
	wdStory = 6
    wdMove = 0
	
  	Set objSelection = objWord.Selection
    objSelection.EndKey wdStory, wdMove   

dim  sleepTime
sleepTime = 5000

Set objIE = CreateObject("InternetExplorer.Application")

objIE.Navigate "file:///C:/workshop/site/login.html"

objIE.Visible = 1

Do until objIE.ReadyState = 4
WScript.Sleep 50
Loop

'chama a funcao de printscreen
printscreen

objSelection.Paste
objSelection.TypeParagraph()
objSelection.TypeParagraph()
	
'move para ultima linha do doc do word
objSelection.EndKey wdStory, wdMove

'efetua login na pagina web
objIE.Document.All.Item("username").Value = "Vanderson" 
objIE.Document.All.Item("password").Value = "1234" 
printscreen 
	
objSelection.Paste
objSelection.TypeParagraph()
objSelection.TypeParagraph()
	
'move para ultima linha
objSelection.EndKey wdStory, wdMove


objIE.Document.All.Item("submit").Click

WScript.Sleep sleepTime

printscreen 
	
objSelection.Paste
objSelection.TypeParagraph()
objSelection.TypeParagraph()
	
'move para ultima linha
objSelection.EndKey wdStory, wdMove


'chama a funcao clicklink passando a url que consta no parametro href do link
clickLink "form_component.html"

printscreen 
	
objSelection.Paste
objSelection.TypeParagraph()
objSelection.TypeParagraph()
	
'move para ultima linha
objSelection.EndKey wdStory, wdMove

clickLink "form_validation.html"
printscreen 
	
objSelection.Paste
objSelection.TypeParagraph()
objSelection.TypeParagraph()
	
'move para ultima linha
objSelection.EndKey wdStory, wdMove


clickLink "general.html"
printscreen 
	
objSelection.Paste
objSelection.TypeParagraph()
objSelection.TypeParagraph()
	
'move para ultima linha
objSelection.EndKey wdStory, wdMove


clickLink "buttons.html"
printscreen 
	
objSelection.Paste
objSelection.TypeParagraph()
objSelection.TypeParagraph()
	
'move para ultima linha
objSelection.EndKey wdStory, wdMove


clickLink "grids.html"
printscreen 
	
objSelection.Paste
objSelection.TypeParagraph()
objSelection.TypeParagraph()
	
'move para ultima linha
objSelection.EndKey wdStory, wdMove


clickLink "widgets.html"
printscreen 
	
objSelection.Paste
objSelection.TypeParagraph()
objSelection.TypeParagraph()
	
'move para ultima linha
objSelection.EndKey wdStory, wdMove


clickLink "chart-chartjs.html"
printscreen
objSelection.Paste
objSelection.TypeParagraph()
objSelection.TypeParagraph()
	
'move para ultima linha
objSelection.EndKey wdStory, wdMove


clickLink "basic_table.html"
printscreen 
	
objSelection.Paste
objSelection.TypeParagraph()
objSelection.TypeParagraph()
	
'move para ultima linha
objSelection.EndKey wdStory, wdMove


objWord.ActiveDocument.SaveAs "C:\workshop\vbs\checkpages\auto_prints.doc"
objWord.ActiveDocument.Close
objword.Quit
Set objword = Nothing


msgbox "fim"

Function clickLink (ref)
For Each a In objIE.Document.GetElementsByTagName("A")
  If LCase(a.GetAttribute("href")) = LCase(ref) Then
    a.Click
    Exit For  ''# to stop after the first hit
  End If
next
WScript.Sleep 50
Do until objIE.ReadyState = 4
WScript.Sleep 1000
Loop
end function

  Function printscreen()
  'Taking Screenshot using word object
	Set oWordBasic = CreateObject("Word.Basic")
	oWordBasic.SendKeys "{prtsc}" 
	Set oWordBasic = Nothing
  end function
Set objIE = Nothing