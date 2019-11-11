Option Explicit On

Imports Contensive
Imports Contensive.BaseClasses
Imports System.Net

Namespace Contensive.Addons.aoLatitudeandLongitude
    Public Class getLatitudeandLongitude
        Inherits AddonBaseClass
        '====================================================================================================
        ''' <summary>
        ''' addon
        ''' </summary>
        ''' <param name="CP"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overrides Function Execute(ByVal CP As BaseClasses.CPBaseClass) As Object
            '

            If Err.Number <> 0 Then
                Call CP.Site.ErrorReport("Error entering GetLatitudeandLongitude")
            End If
            If CP.Utils.EncodeBoolean(CP.Site.TrapErrors) Then
                On Error Resume Next
            End If
            '
            Dim defaultValue
            Dim stream As String = ""
            Dim address As String = ""
            Dim objHttp As WebClient = New WebClient()
            Dim serviceURL As String = ""
            Dim serviceResponse As String = ""
            Dim objXML As System.Xml.XmlDocument = New System.Xml.XmlDocument()
            Dim geometryNode
            Dim modAddress As String = ""
            Dim pointer As Integer
            Dim apiKey As String = ""
            '
            defaultValue = "0^0"
            address = CP.Doc.GetText("MSfindRideZip")
            apiKey = CP.Site.GetText("GOOGLE API KEY")

            If (String.IsNullOrEmpty(apiKey)) Then
                CP.Site.ErrorReport("The GoogleAPIKey in site properties is empty")
                Return "GoogleAPIKey is empty"
            End If

            '
            If address <> "" Then
                '
                modAddress = kmaEncodeURL(CP, address)
                '            
                serviceURL = "https://maps.googleapis.com/maps/api/geocode/xml?key=" & apiKey & "&address=" & modAddress & "&sensor=false"
                'serviceResponse = objHttp.getURL(CStr(serviceURL))

                serviceResponse = objHttp.DownloadString(serviceURL)
                objHttp = Nothing
                '
                objXML = New System.Xml.XmlDocument()

                Call objXML.LoadXml(serviceResponse)
                If objXML.HasChildNodes Then
                    Dim length = objXML.GetElementsByTagName("geometry").ItemOf(0).ChildNodes.Count - 1
                    For pointer = 0 To length
                        If objXML.GetElementsByTagName("geometry").ItemOf(0).ChildNodes(pointer).Name = "location" Then
                            stream = stream & objXML.GetElementsByTagName("geometry").ItemOf(0).ChildNodes(pointer).ChildNodes.ItemOf(0).InnerText & "^"
                            stream = stream & objXML.GetElementsByTagName("geometry").ItemOf(0).ChildNodes(pointer).ChildNodes.ItemOf(1).InnerText
                        End If
                    Next
                End If
                objXML = Nothing
            End If

            '
            If stream = "" Then
                stream = defaultValue
            End If
            '

            If Err.Number <> 0 Then
                Call CP.Site.ErrorReport("Error exiting GetLatitudeandLongitude")
            End If
            Return stream
            '

            '
        End Function
        '
        Private Function kmaEncodeURL(cp, Source) As String
            '
            If Err.Number <> 0 Then
                Call cp.Site.ErrorReport("Error entering kmaEncodeURl")
            End If
            If cp.utils.encodeBoolean(cp.Site.TrapErrors) Then
                On Error Resume Next
            End If
            '
            Dim LeftSide
            Dim RightSide
            '
            kmaEncodeURL = Source
            If Source <> "" Then
                kmaEncodeURL = Replace(kmaEncodeURL, "#", " ")
                kmaEncodeURL = Replace(kmaEncodeURL, "%", "%25")
                '
                kmaEncodeURL = Replace(kmaEncodeURL, """", "%22")
                kmaEncodeURL = Replace(kmaEncodeURL, " ", "%20")
                kmaEncodeURL = Replace(kmaEncodeURL, "$", "%24")
                kmaEncodeURL = Replace(kmaEncodeURL, "+", "%2B")
                kmaEncodeURL = Replace(kmaEncodeURL, ",", "%2C")
                kmaEncodeURL = Replace(kmaEncodeURL, ";", "%3B")
                kmaEncodeURL = Replace(kmaEncodeURL, "<", "%3C")
                kmaEncodeURL = Replace(kmaEncodeURL, "=", "%3D")
                kmaEncodeURL = Replace(kmaEncodeURL, ">", "%3E")
                kmaEncodeURL = Replace(kmaEncodeURL, "@", "%40")
                '
            End If
            '
            If Err.Number <> 0 Then
                Call cp.Site.ErrorReport("Error exiting kmaEncodeURL")
            End If
            '
            Return Source
        End Function
    End Class
End Namespace
