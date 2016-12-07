Imports System.Reflection

Public Class RulesEngine

    Implements Microsoft.Vsa.IVsaSite
    Private vsaModel As New Microsoft.VisualBasic.Vsa.VsaEngine

    Public Const MONIKER As String = "net.Appagare://Rules/Engine/Class"
    Public Const NAME_SPACE As String = "Processing.Rule"
    Public Const SCRIPT As String = "Script"
    Public Const MAIN As String = "ASi_Main"
    Public Const CODE_TYPE As String = "Processing.Rule." & SCRIPT

    Private _ReturnValue As Object = Nothing

    Public ReadOnly Property ReturnValue() As Object
        Get
            Return _ReturnValue
        End Get
    End Property

    Public Function Execute(ByVal Code As String, Optional ByVal Globals As String = "", _
        Optional ByVal [Imports] As String = "") As Object

        Dim CodeItem As Vsa.VsaCodeItem
        Dim asb As [Assembly]
        Dim myType As Type
        Dim Meth As MethodInfo
        Dim ReturnValue As Object = Nothing
        Dim RefItem As Vsa.VsaReferenceItem

        Try
            vsaModel.RevokeCache()
            vsaModel.RootMoniker = MONIKER
            vsaModel.Site = Me
            vsaModel.RootNamespace = NAME_SPACE
            vsaModel.InitNew()

            'add the reference if exists
            If [Imports] <> "" AndAlso InStr([Imports], ",") > 1 Then
                Try
                    'imports is dll,object,name
                    Dim a() As String
                    a = Split([Imports], ",")
                    If UBound(a) = 2 Then
                        'need name and location
                        Dim RefName As String = a(2)
                        Dim RefLocation As String = Left(System.Reflection.Assembly.GetExecutingAssembly.Location, _
                            InStrRev(System.Reflection.Assembly.GetExecutingAssembly.Location, "\"))

                        'RefLocation = "C:\Inetpub\wwwroot\CoolCart.Net\CoolCart\bin\"
                        'If Right(RefLocation, 1) <> "\" Then
                        '    RefLocation &= "\"
                        'End If
                        RefLocation = a(0)

                        RefItem = vsaModel.Items.CreateItem(RefName, Microsoft.Vsa.VsaItemType.Reference, Microsoft.Vsa.VsaItemFlag.None)
                        RefItem.AssemblyName = RefLocation

                        [Imports] = "imports " & a(1) & vbCrLf
                    End If


                Catch ex As Exception
                    Throw New Exception("Imports error - " & ex.Message & IIf(ex.InnerException Is Nothing, "", " InnerException: " & ex.InnerException.Message))
                End Try
            End If

            'add the code item
            CodeItem = vsaModel.Items.CreateItem(SCRIPT, Microsoft.Vsa.VsaItemType.Code, Microsoft.Vsa.VsaItemFlag.Module)
            CodeItem.SourceText = "imports system" & vbCrLf & _
                [Imports] & vbCrLf & _
                "Module Script" & vbCrLf & _
                Globals & vbCrLf & _
                Code & vbCrLf & _
                "End Module" & vbCrLf

            vsaModel.Compile()
            vsaModel.Run()
            asb = vsaModel.Assembly
            myType = asb.GetType(CODE_TYPE)
            Meth = myType.GetMethod(MAIN)
            ReturnValue = Meth.Invoke(Nothing, Nothing)
            vsaModel.Close()
            Return ReturnValue
        Catch ex As Exception
            Throw New Exception(ex.Message & IIf(ex.InnerException Is Nothing, "", " InnerException: " & ex.InnerException.Message))
        End Try
    End Function
    Public Sub GetCompiledState(ByRef pe() As Byte, ByRef debugInfo() As Byte) Implements Microsoft.Vsa.IVsaSite.GetCompiledState
    End Sub
    Public Function GetEventSourceInstance(ByVal itemName As String, ByVal eventSourceName As String) As Object Implements Microsoft.Vsa.IVsaSite.GetEventSourceInstance
    End Function
    Public Function GetGlobalInstance(ByVal name As String) As Object Implements Microsoft.Vsa.IVsaSite.GetGlobalInstance
    End Function
    Public Sub Notify(ByVal notify As String, ByVal info As Object) Implements Microsoft.Vsa.IVsaSite.Notify
    End Sub
    Public Function OnCompilerError(ByVal [error] As Microsoft.Vsa.IVsaError) As Boolean Implements Microsoft.Vsa.IVsaSite.OnCompilerError
    End Function

End Class

