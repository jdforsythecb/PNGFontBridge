Imports System.Xml
Imports System.IO
Imports System.Drawing
Imports System.Drawing.Imaging


Public Class PNGFontBridge

    Private fnx As String = ""

    Private Const PNGPath As String = "g:\PNGFonts\"

    Property Folder() As String
        Get
            Return fnx
        End Get
        Set(value As String)
            fnx = value
        End Set
    End Property

    Public Function getFolderList(Optional ByVal filter As String = "") As List(Of String)
        Dim files As New List(Of String)

        If (filter = "") Then
            filter = "*.fnx"
        Else
            filter = "*" & filter & "*.fnx"
        End If

        For Each fnx As String In Directory.EnumerateFiles(PNGPath, filter, SearchOption.TopDirectoryOnly).OrderBy(Function(x) Path.GetFileName(x))
            files.Add(Path.GetFileNameWithoutExtension(fnx))
        Next

        Return files
    End Function

    Public Function getFnxPathFromFolderNumber(Optional ByVal folderNum As String = "") As String
        Try
            '' if no folder number is passed, try to use the Folder() property
            If (folderNum = "") Then folderNum = fnx
        Catch ex As Exception When folderNum = ""
            Throw New ApplicationException("PNGFontBridge:getFnxPathFromFolderNumber - No folder number specified", ex)
        End Try

        Return PNGPath & folderNum & ".fnx"
    End Function

    '' returns a dictionary of key:id, value:dictionary[image:base64imagedata, fontCode:font code, fontTitle:font title]
    Public Function getCopyListFromFolderNumber(Optional ByVal folderNum As String = "", Optional suppressErrors As Boolean = False) As Dictionary(Of String, Dictionary(Of String, String))
        Dim images As New Dictionary(Of String, Dictionary(Of String, String))
        Try
            '' if no folderNum is passed, try to use the Folder() property
            If (folderNum = "") Then folderNum = fnx
            Dim xmlReader As XmlTextReader = New XmlTextReader(getFnxPathFromFolderNumber(folderNum))

            '' we know there is a supplied folder number and that it exists, so perform the search through the font to get the copy
            Dim image As String = ""
            Dim id As String = ""
            Dim fontCode As String = ""
            Dim fontTitle As String = ""

            Do While (xmlReader.Read())
                '' we only want things returned that are elements of type "ROW" with attribute "image"
                If ((xmlReader.NodeType = XmlNodeType.Element) And
                    (xmlReader.Name.ToUpper = "ROW") And
                    (xmlReader.HasAttributes)) Then
                    '' we only want elements that have an image attribute
                    xmlReader.MoveToAttribute("Image")
                    '' if it has a value for the image attribute, get the image, id, code, and title values
                    If (xmlReader.Value.Length > 0) Then
                        image = xmlReader.Value

                        xmlReader.MoveToAttribute("Id")
                        id = xmlReader.Value

                        xmlReader.MoveToAttribute("Code")
                        fontCode = xmlReader.Value

                        xmlReader.MoveToAttribute("Title")
                        fontTitle = xmlReader.Value

                        '' create a dictionary of the values
                        Dim itemDictionary As New Dictionary(Of String, String)
                        itemDictionary.Add("fontCode", fontCode)
                        itemDictionary.Add("fontTitle", fontTitle)
                        itemDictionary.Add("image", image)

                        ' '' if the id is already in the images dictionary, remove it so it can be replaced
                        'If (images.ContainsKey(id)) Then images.Remove(id)

                        '' instead of removing existing entries in the images dictionary, we want to overwrite with
                        '' any data we have, in case the newest copy doesn't have a title, code, etc (it will always
                        '' have an image and id)
                        If (images.ContainsKey(id)) Then
                            '' we know we have an image, so overwrite the image for this id
                            images.Item(id).Item("image") = itemDictionary.Item("image")

                            '' if we have a font code, overwrite the font code for this id
                            If Not (fontCode = "") Then images.Item(id).Item("fontCode") = itemDictionary.Item("fontCode")

                            '' if we have a font title, overwrite the font title for this id
                            If Not (fontTitle = "") Then images.Item(id).Item("fontTitle") = itemDictionary.Item("fontTitle")

                        Else
                            '' if this is the first of this id, then just add what we have to the dictionary
                            images.Add(id, itemDictionary)
                        End If


                    End If
                End If
            Loop

        Catch ex As Exception When folderNum = ""
            If Not (suppressErrors) Then
                Throw New ApplicationException("PNGFontBridge:getCopyListFromFolderNumber - No folder number specified", ex)
            End If
        Catch ex As FileNotFoundException
            If Not (suppressErrors) Then
                Throw New ApplicationException("PNGFontBridge:getCopyListFromFolderNumber - Specified folder does not exist", ex)
            End If

        End Try

        Return images

    End Function


    Public Function getBitmapFromBase64(ByVal base64 As String) As Bitmap
        Try
            Dim ms As New MemoryStream(Convert.FromBase64String(base64))
            Dim bmp As New Bitmap(ms)
            ms.Close()
            Return bmp
        Catch ex As Exception
            Throw New ApplicationException("PNGFontBridge:getBitmapFromBase64 - invalid base64 image string", ex)
        End Try
    End Function

    Public Function copyBitmapToClipboardPNG(ByVal bmp As Bitmap) As Boolean
        Try
            Dim clipboardPNG As String = "c:\eps\dump\clipboard.png"
            If (File.Exists(clipboardPNG)) Then File.Delete(clipboardPNG)
            bmp.Save(clipboardPNG, ImageFormat.Png)
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

End Class
