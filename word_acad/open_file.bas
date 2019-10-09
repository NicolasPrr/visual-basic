Public Graphics As AcadApplication
Public Sub OpnAcad()
    'Antes de correr se debe agregar las correspondientes referencias
    'Macro corrida en word
    
        
    Dim AngBracDwg As String
    Dim MyPath As String
    Dim MyLine As AcadLine
    
    'Crear el objeto
    Set Graphics = GetObject(, "AutoCAD.Application")
    If Err.Description > vbNullString Then
        Err.Clear
        Set Graphics = CreateObject("AutoCAD.Application")
    End If
    'Seleccionamos la localizacion del archivo
    MyPath = Application.ActiveDocument.Path
    'Concatenamos con el nobmre del archivo, en este caso Test.dwg
    MyPath = MyPath + "/test.dwg"
    AngBracDwg = MyPath
    'Abrimos el archivo
    Graphics.Documents.Open (AngBracDwg)
    
    'dibujo actual!
    Set CurrentDraw = Graphics.ActiveDocument
    
    Dim plineObj As AcadLWPolyline
    Dim points(0 To 5) As Double
    
    
      Graphics.Application.ZoomAll
         
         
      'dibujando liena
      
       
      Start = CurrentDraw.Utility.GetPoint(, "select  start point :")
      Finish = CurrentDraw.Utility.GetPoint(, "select  finish point :")
      Set MyLine = CurrentDraw.ModelSpace.AddLine(Start, Finish)
   
       Dim color As AcadAcCmColor
  '     Set color = AcadApplication.GetInterfaceObject("AutoCAD.AcCmColor.22")
   '    Call color.SetRGB(80, 100, 244)

'      MyLine.TrueColor = color
     
     
      Graphics.ActiveDocument.Utility.Prompt ("Hey everyone")
    
End Sub
