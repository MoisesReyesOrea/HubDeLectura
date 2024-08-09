Attribute VB_Name = "ModuleLoadInitialProperties"


Function PositionFrames()
    ' Ajustar posición y tamaño de fraNavbar
    FormMain.fraNavbar.Left = 30
    FormMain.fraNavbar.Top = 0
    FormMain.fraNavbar.Width = 3000
    FormMain.fraNavbar.Height = FormMain.ScaleHeight - 50
    
    ' Ajustar posición y tamaño de fraAddBooks
    FormMain.fraHome.Left = 3060
    FormMain.fraHome.Top = 0
    FormMain.fraHome.Width = FormMain.ScaleWidth - 3090 ' Ajustar el ancho según el formulario
    FormMain.fraHome.Height = FormMain.ScaleHeight - 50 ' Ajustar la altura según el formulario
    ' PanelContent.AutoScroll = True ' Habilita las barras de desplazamiento
    
    ' Ajustar posición y tamaño de frafavorites
    FormMain.fraFavorites.Left = 3060
    FormMain.fraFavorites.Top = 0
    FormMain.fraFavorites.Width = FormMain.ScaleWidth - 3090 ' Ajustar el ancho según el formulario
    FormMain.fraFavorites.Height = FormMain.ScaleHeight - 50 ' Ajustar la altura según el formulario
    ' PanelContent.AutoScroll = True ' Habilita las barras de desplazamiento
    
    ' Ajustar posición y tamaño de fraCompletedBooks
    FormMain.fraCompletedBooks.Left = 3060
    FormMain.fraCompletedBooks.Top = 0
    FormMain.fraCompletedBooks.Width = FormMain.ScaleWidth - 3090 ' Ajustar el ancho según el formulario
    FormMain.fraCompletedBooks.Height = FormMain.ScaleHeight - 50 ' Ajustar la altura según el formulario
    ' PanelContent.AutoScroll = True ' Habilita las barras de desplazamiento
    
    ' Ajustar posición y tamaño de fraHistory
    FormMain.fraHistory.Left = 3060
    FormMain.fraHistory.Top = 0
    FormMain.fraHistory.Width = FormMain.ScaleWidth - 3090 ' Ajustar el ancho según el formulario
    FormMain.fraHistory.Height = FormMain.ScaleHeight - 50 ' Ajustar la altura según el formulario
    ' PanelContent.AutoScroll = True ' Habilita las barras de desplazamiento
    
    ' Ajustar posición y tamaño de frmNoWished
    FormMain.frmNoWished.Left = 3060
    FormMain.frmNoWished.Top = 0
    FormMain.frmNoWished.Width = FormMain.ScaleWidth - 3090 ' Ajustar el ancho según el formulario
    FormMain.frmNoWished.Height = FormMain.ScaleHeight - 50 ' Ajustar la altura según el formulario
    ' PanelContent.AutoScroll = True ' Habilita las barras de desplazamiento
End Function

