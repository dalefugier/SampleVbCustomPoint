Imports Rhino.FileIO
Imports Rhino

Namespace SampleVbCustomPoint

  ''' <summary>
  ''' SampleVbCustomPointPlugIn plug-in
  ''' </summary>
  Public Class SampleVbCustomPointPlugIn
    Inherits PlugIns.PlugIn

    Private Shared _instance As SampleVbCustomPointPlugIn
    Public Property Point0 As SampleVbPoint
    Public Property Point1 As SampleVbPoint
    Public Property Point2 As SampleVbPoint

    ''' <summary>
    ''' Constructor
    ''' </summary>
    Public Sub New()
      ' Rhino only creates one instance of a plug-in class, 
      ' so it is safe to store a reference in a static field.
      _instance = Me
      Point0 = SampleVbPoint.Unset
      Point1 = SampleVbPoint.Unset
      Point2 = SampleVbPoint.Unset
    End Sub

    ''' <summary>
    ''' Returns the only instance of the SampleVbCustomPointPlugIn plug-in
    ''' </summary>
    Public Shared ReadOnly Property Instance() As SampleVbCustomPointPlugIn
      Get
        Return _instance
      End Get
    End Property

    ''' <summary>
    ''' ShouldCallWriteDocument override
    ''' </summary>
    Protected Overrides Function ShouldCallWriteDocument(options As FileWriteOptions) As Boolean
      If (Point0.IsValid AndAlso Point1.IsValid AndAlso Point2.IsValid) Then
        Return True
      Else
        Return False
      End If
    End Function

    ''' <summary>
    ''' WriteDocument override
    ''' </summary>
    Protected Overrides Sub WriteDocument(doc As RhinoDoc, archive As BinaryArchiveWriter, options As FileWriteOptions)
      Try
        archive.Write3dmChunkVersion(1, 0)
        Dim rc As Boolean = Not archive.WriteErrorOccured
        If (rc = True) Then rc = Point0.Write(archive)
        If (rc = True) Then rc = Point1.Write(archive)
        If (rc = True) Then Point2.Write(archive)
      Catch
        ' TODO...
      End Try
    End Sub

    ''' <summary>
    ''' ReadDocument override
    ''' </summary>
    Protected Overrides Sub ReadDocument(doc As RhinoDoc, archive As BinaryArchiveReader, options As FileReadOptions)
      Try
        Dim major As Integer, minor As Integer
        archive.Read3dmChunkVersion(major, minor)
        Dim rc As Boolean = Not archive.ReadErrorOccured
        If (rc = True) Then rc = (1 = major AndAlso 0 = minor)
        If (rc = True) Then rc = Point0.Read(archive)
        If (rc = True) Then rc = Point1.Read(archive)
        If (rc = True) Then Point2.Read(archive)
      Catch
        ' TODO...
      End Try
    End Sub

  End Class
End Namespace