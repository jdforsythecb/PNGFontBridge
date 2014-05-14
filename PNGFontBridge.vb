Imports System.Xml
Imports System.IO
Imports System.Drawing
Imports System.Drawing.Imaging


Public Class PNGFontBridge

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '' PRIVATE VARIABLES '''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    '' holds the folder number, or fnx filename (set by Folder() property)
    Private fnx As String = ""

    '' base path for fnx files
    Private Const PNGPath As String = "g:\PNGFonts\"



    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '' PUBLIC PROPERTIES '''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    '' publicly settable property for fnx
    Property Folder() As String
        Get
            Return fnx
        End Get
        Set(value As String)
            fnx = value
        End Set
    End Property




    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '' PUBLIC METHODS ''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    '' returns a dictionary of key:id, value:dictionary[image:base64imagedata, fontCode:font code, fontTitle:font title]
    Public Function getCopyListFromFolderNumber(Optional ByVal folderNum As String = "", Optional suppressErrors As Boolean = False) As Dictionary(Of String, Dictionary(Of String, String))
        Try
            '' if no folderNum is passed, try to use the Folder() property
            If (folderNum = "") Then folderNum = fnx

            Dim images As New Dictionary(Of String, Dictionary(Of String, String))
            Dim b64image As String = "", id As String = "", fontCode As String = "", fontTitle As String = ""

            Using fs As New FileStream(getFnxPathFromFolderNumber(folderNum), FileMode.Open, FileAccess.Read)
                Using reader As XmlReader = XmlReader.Create(fs)
                    While reader.Read

                        '' we only want elements of type row with attribute image
                        If ((reader.NodeType = XmlNodeType.Element) AndAlso
                            (reader.Name.ToUpper = "ROW") AndAlso
                            (reader.GetAttribute("Image") <> Nothing)) Then

                            id = reader.GetAttribute("Id")
                            b64image = reader.GetAttribute("Image")
                            fontCode = reader.GetAttribute("Code")
                            fontTitle = reader.GetAttribute("Title")

                            '' if the id is already in the images dictionary, we want to overwrite with the
                            '' data we have, since the further down the file, the newer the data
                            '' but sometimes we don't get a full set of data, so we'll keep the old attributes
                            '' that are already in the list
                            If (images.ContainsKey(id)) Then
                                '' we will always have an image, since we ignore anything without one
                                images.Item(id).Item("image") = b64image

                                '' if we have the other data, overwrite the old
                                If Not (fontCode = "") Then images.Item(id).Item("fontCode") = fontCode
                                If Not (fontTitle = "") Then images.Item(id).Item("fontTitle") = fontTitle

                            Else
                                '' otherwise we have no entries for this image, so create a new one
                                Dim fontDictionary As New Dictionary(Of String, String)
                                fontDictionary.Add("fontCode", fontCode)
                                fontDictionary.Add("fontTitle", fontTitle)
                                fontDictionary.Add("image", b64image)
                                images.Add(id, fontDictionary)
                            End If

                        End If
                    End While
                    'reader.Dispose()
                End Using
                fs.Dispose()
            End Using

            '' return the latest images
            Return images

        Catch e As Exception
            Throw
        End Try

    End Function

    '' give a base 64 string, get a bitmap object back
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

    '' give a bitmap object (such as a PictureBox.Image) and save it as a PNG
    '' to c:\eps\dump\clipboard.png - returns true or false, success or fail
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

    '' give a bitmap object and save it as a PNG to the supplied base path and filename
    Public Function saveBitmapToPng(ByVal bmp As Bitmap, ByVal basePath As String, ByVal filename As String, Optional ByVal overwrite As Boolean = True) As Boolean
        Try
            Dim fullPath As String = basePath & filename & ".png"
            If (File.Exists(fullPath)) Then
                If (overwrite) Then
                    File.Delete(fullPath)
                Else
                    Return False
                End If
            End If
            bmp.Save(fullPath, ImageFormat.Png)
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function


    '' returns the latest weekly image as a base64 string
    Public Function getNewestWeeklyFromFolderNumber(Optional ByVal folderNum As String = "") As Dictionary(Of String, String)
        Try
            '' if no folderNum is passed, try to use the Folder() property
            If (folderNum = "") Then folderNum = fnx

            Dim weekly As New Dictionary(Of String, String)

            Using fs As New FileStream(getFnxPathFromFolderNumber(folderNum), FileMode.Open, FileAccess.Read)
                Using reader As XmlReader = XmlReader.Create(fs)
                    While reader.Read
                        '' we only want elements of type row with attribute image and a title or font
                        '' code that identifies it as a weekly font
                        If ((reader.NodeType = XmlNodeType.Element) AndAlso
                            (reader.Name.ToUpper = "ROW") AndAlso
                            (reader.HasAttributes) AndAlso
                            Not (reader.GetAttribute("Image") = Nothing) AndAlso
                            ((Not (reader.GetAttribute("Title") = Nothing) AndAlso
                             (reader.GetAttribute("Title").ToUpper.Contains("WEEKLY"))) OrElse
                             ((Not (reader.GetAttribute("Code") = Nothing) AndAlso
                              (reader.GetAttribute("Code").Length > 1)) AndAlso
                              ((reader.GetAttribute("Code").ToUpper.Substring(0, 2) = "WK") OrElse
                              ((reader.GetAttribute("Code").ToUpper.Substring(0, 1) = "W") AndAlso
                               (IsNumeric(reader.GetAttribute("Code").Substring(1, 1)))))))) Then

                            weekly.Add("Id", reader.GetAttribute("Id"))
                            weekly.Add("Code", reader.GetAttribute("Code"))
                            weekly.Add("Title", reader.GetAttribute("Title"))
                            weekly.Add("Image", reader.GetAttribute("Image"))
                            weekly.Add("Position", reader.GetAttribute("Position"))

                        End If

                    End While
                End Using
                fs.Dispose()
            End Using

            '' return the latest image data
            Return weekly

        Catch e As Exception
            Throw
        End Try

    End Function

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '' PRIVATE METHODS '''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    '' provide a folder number, get a file path
    Private Function getFnxPathFromFolderNumber(Optional ByVal folderNum As String = "") As String
        '' if no folder number is passed, try to use the Folder() property
        If (folderNum = "") Then folderNum = fnx
        '' we don't worry about error handling here - if we fail to provide a folderNum, we'll
        '' end up generating a FileNotFoundException later on
        Return PNGPath & folderNum & ".fnx"
    End Function

End Class
