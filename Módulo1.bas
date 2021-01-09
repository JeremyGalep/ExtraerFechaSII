Attribute VB_Name = "Módulo1"
Sub ExtraerFechaSII()
Dim ox As Integer
ox = 1
Do While Cells(ox, 1).Value <> ""
    ox = ox + 1
Loop
ox = ox - 1

For SII = 1 To ox

Dim pagina As HTMLDocument
Dim explorador As InternetExplorer
Dim direccion As String

direccion = "https://zeus.sii.cl/cvc/stc/stc.html"

usuario = Cells(SII, 1).Value
contrasena = Cells(SII, 2).Value


Set explorador = New InternetExplorer
explorador.Visible = False
explorador.GoHome
explorador.Navigate direccion

Do
DoEvents
Loop Until explorador.ReadyState = READYSTATE_COMPLETE

Set pagina = explorador.document

pagina.getElementById("RUT").innerText = usuario
pagina.getElementById("DV").innerText = contrasena
pagina.getElementById("txt_code").Value = "\"

pagina.form1.submit



Do While explorador.Busy

Loop



Cells(SII, 3) = Mid(pagina.all(63).innerText, 32)

explorador.Quit


Next SII

End Sub
