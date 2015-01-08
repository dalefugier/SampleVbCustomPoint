Imports System.IO
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Globalization
Imports Rhino
Imports Rhino.FileIO
Imports Rhino.Geometry

<Serializable()> Public Class SampleVbPoint

  Private Const UnsetValue As Double = -1.23432101234321E+308

  Public Property X As Double
  Public Property Y As Double
  Public Property Z As Double

  ''' <summary>
  ''' Default constructor
  ''' </summary>
  Sub New()
    X = RhinoMath.UnsetValue
    Y = RhinoMath.UnsetValue
    Z = RhinoMath.UnsetValue
  End Sub

  ''' <summary>
  ''' Constructor
  ''' </summary>
  Sub New(ByVal dx As Double, ByVal dy As Double, ByVal dz As Double)
    X = dx
    Y = dy
    Z = dz
  End Sub

  ''' <summary>
  ''' Copy constructor
  ''' </summary>
  Sub New(ByVal src As SampleVbPoint)
    X = src.X
    Y = src.Y
    Z = src.Z
  End Sub

  ''' <summary>
  ''' Construct from Rhino.Geometry.Point3d
  ''' </summary>
  Sub New(ByVal src As Point3d)
    X = src.X
    Y = src.Y
    Z = src.Z
  End Sub

  ''' <summary>
  ''' Create
  ''' </summary>
  Public Sub Create(src As SampleVbPoint)
    Me.X = src.X
    Me.Y = src.Y
    Me.Z = src.Z
  End Sub

  ''' <summary>
  ''' Returns an unset point
  ''' </summary>
  Public Shared ReadOnly Property Unset() As SampleVbPoint
    Get
      Return New SampleVbPoint(UnsetValue, UnsetValue, UnsetValue)
    End Get
  End Property

  ''' <summary>
  ''' Assignment operator
  ''' </summary>
  Public Shared Widening Operator CType(ByVal src As Point3d) As SampleVbPoint
    Return New SampleVbPoint(src)
  End Operator

  ''' <summary>
  ''' Assignment operator
  ''' </summary>
  Public Shared Widening Operator CType(ByVal src As SampleVbPoint) As Point3d
    Return New Point3d(src.X, src.Y, src.Z)
  End Operator

  ''' <summary>
  ''' Equals operator 
  ''' </summary>
  Public Shared Operator =(ByVal p0 As SampleVbPoint, ByVal p1 As SampleVbPoint) As Boolean
    If (p0.X.Equals(p1.X) AndAlso p0.Y.Equals(p1.Y) AndAlso p0.Z.Equals(p1.Z)) Then
      Return True
    Else
      Return False
    End If
  End Operator

  ''' <summary>
  ''' Not equals operator
  ''' </summary>
  Public Shared Operator <>(ByVal p0 As SampleVbPoint, ByVal p1 As SampleVbPoint) As Boolean
    If (Not p0.X.Equals(p1.X) OrElse Not p0.Y.Equals(p1.Y) OrElse Not p0.Z.Equals(p1.Z)) Then
      Return True
    Else
      Return False
    End If
  End Operator

  ''' <summary>
  ''' IsValid
  ''' </summary>
  Public ReadOnly Property IsValid() As Boolean
    Get
      If (Not X.Equals(UnsetValue) AndAlso Not Y.Equals(UnsetValue) AndAlso Not Z.Equals(UnsetValue)) Then
        Return True
      Else
        Return False
      End If
    End Get
  End Property

  ''' <summary>
  ''' Write
  ''' </summary>
  Public Function Write(ByRef archive As BinaryArchiveWriter) As Boolean
    Dim rc As Boolean = False
    Try
      archive.Write3dmChunkVersion(1, 0)
      Dim formatter As IFormatter = New BinaryFormatter()
      Dim stream As MemoryStream = New MemoryStream()
      formatter.Serialize(stream, Me)
      stream.Seek(0, 0)
      Dim bytes() As Byte = stream.ToArray()
      archive.WriteByteArray(bytes)
      stream.Close()
      rc = Not archive.WriteErrorOccured
    Catch
      ' TODO...
    End Try
    Return rc
  End Function

  ''' <summary>
  ''' Read
  ''' </summary>
  Public Function Read(ByRef archive As BinaryArchiveReader) As Boolean
    Dim rc As Boolean = False
    Dim major As Integer, minor As Integer
    Try
      archive.Read3dmChunkVersion(major, minor)
      If (1 = major AndAlso 0 = minor) Then
        Dim bytes() As Byte = archive.ReadByteArray()
        Dim stream As MemoryStream = New MemoryStream(bytes)
        Dim formatter As IFormatter = New BinaryFormatter()
        Dim data As SampleVbPoint = CType(formatter.Deserialize(stream), SampleVbPoint)
        If data IsNot Nothing Then
          Me.Create(data)
        End If
        rc = Not archive.ReadErrorOccured
      End If
    Catch
      ' TODO...
    End Try
    Return rc
  End Function

  ''' <summary>
  ''' ToString override
  ''' </summary>
  Public Overrides Function ToString() As String
    Const format As String = "F"
    Dim provider = CultureInfo.InvariantCulture
    Dim sx = X.ToString(format, provider)
    Dim sy = Y.ToString(format, provider)
    Dim sz = Z.ToString(format, provider)
    Return String.Format("{0},{1},{2}", sx, sy, sz)
  End Function
End Class
