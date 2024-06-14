'TODO:  Siga estos pasos para habilitar el elemento (XML) de la cinta de opciones:

'1: Copie el siguiente bloque de código en la clase ThisAddin, ThisWorkbook o ThisDocument.

'Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
'    Return New Ribbon1()
'End Function

'2. Cree métodos de devolución de llamada en el área "Devolución de llamadas de la cinta de opciones" de esta clase para controlar acciones del usuario,
'   como hacer clic en un botón. Nota: si ha exportado esta cinta desde el
'   diseñador de la cinta de opciones, deberá mover el código de los controladores de eventos a los métodos de devolución de llamada y
'   modificar el código para que funcione con el modelo de programación de extensibilidad de la cinta de opciones (RibbonX).

'3. Asigne atributos a las etiquetas de control del archivo XML de la cinta de opciones para identificar los métodos de devolución de llamada apropiados en el código.

'Para obtener más información, vea la documentación XML de la cinta de opciones en la Ayuda de Visual Studio Tools para Office.

<Runtime.InteropServices.ComVisible(True)>
Public Class Ribbon1
    Implements Office.IRibbonExtensibility

    Private ribbon As Office.IRibbonUI

    Public Sub New()
    End Sub

    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
        Return GetResourceText("Complemento_De_Econometria_Basica.Ribbon1.xml")
    End Function

#Region "Devoluciones de llamada de la cinta de opciones"
    'Cree métodos de devolución de llamada aquí. Para obtener más información sobre la adición de métodos de devolución de llamada, visite https://go.microsoft.com/fwlink/?LinkID=271226
    Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbon = ribbonUI
    End Sub

    Public Sub btn_RLUV(Control As Office.IRibbonControl)
        Dim frm_RLUV As New UF_RLUV
        frm_RLUV.Show()
    End Sub

    Public Sub btn_GHPH(Control As Office.IRibbonControl)
        On Error Resume Next
        hojaHipotesis()
    End Sub

    Public Sub btn_ScaPlot(Control As Office.IRibbonControl)
        Dim frm_GGD As New UF_GGD
        frm_GGD.Show()
    End Sub

    Public Sub btn_HipSp(Control As Office.IRibbonControl)
        Dim frm_HIPOTESIS As New UF_HIPOTESIS
        frm_HIPOTESIS.ShowDialog()
    End Sub
    Public Sub btn_IntConf(Control As Office.IRibbonControl)
        Dim frm_INTERVALO_DE_CONFIANZA As New UF_INTERVALO_DE_CONFIANZA
        frm_INTERVALO_DE_CONFIANZA.ShowDialog()
    End Sub
    Public Sub btn_CudAnov(Control As Office.IRibbonControl)
        cuadroANOVAmet()
    End Sub
    Public Sub btn_RLVV(Control As Office.IRibbonControl)
        Dim frm_RLVV As New UF_RLVV
        frm_RLVV.Show()
    End Sub

    Public Sub btn_DocsGitHub(Control As Office.IRibbonControl)
        Dim webAddress As String = "https://github.com/LASPUMSS/COMPLEMENTO-EXCEL-DE-ECONOMETRIA"
        System.Diagnostics.Process.Start(webAddress)
    End Sub

    Public Sub btn_DocsBlog(Control As Office.IRibbonControl)
        Dim webAddress As String = "https://primero-los-datos.blogspot.com/2023/01/complemento-excel-de-econometria-basica.html"
        System.Diagnostics.Process.Start(webAddress)
    End Sub

#End Region

#Region "Rutinas que permiten convertir las imágenes y cargarlas a cada botón de la cinta"
    Public Function GetCustomImage(ByVal Ctrl As Office.IRibbonControl) As stdole.IPictureDisp

        Dim pictureDisplay As stdole.IPictureDisp = Nothing

        Select Case Ctrl.Id
            Case Is = "docs_github"
                pictureDisplay = ImageConverter.Convert(My.Resources.github_logo())
            Case Is = "docs_blog"
                pictureDisplay = ImageConverter.Convert(My.Resources.blogger_logo())
        End Select

        Return pictureDisplay

    End Function

    Friend Class ImageConverter

        Inherits System.Windows.Forms.AxHost

        Sub New()
            MyBase.New(Nothing)
        End Sub

        Public Shared Function _
        Convert(ByVal image As System.Drawing.Image) _
        As stdole.IPictureDisp
            Return GetIPictureDispFromPicture(image)
        End Function
    End Class

#End Region

#Region "Asistentes"

    Private Shared Function GetResourceText(ByVal resourceName As String) As String
        Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
        Dim resourceNames() As String = asm.GetManifestResourceNames()
        For i As Integer = 0 To resourceNames.Length - 1
            If String.Compare(resourceName, resourceNames(i), StringComparison.OrdinalIgnoreCase) = 0 Then
                Using resourceReader As IO.StreamReader = New IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                    If resourceReader IsNot Nothing Then
                        Return resourceReader.ReadToEnd()
                    End If
                End Using
            End If
        Next
        Return Nothing
    End Function

#End Region

End Class
