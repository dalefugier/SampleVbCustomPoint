Imports System.Globalization
Imports Rhino
Imports Rhino.FileIO
Imports Rhino.Geometry

Public Class SampleVbVector

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
  Sub New(ByVal src As SampleVbVector)
    X = src.X
    Y = src.Y
    Z = src.Z
  End Sub

  ''' <summary>
  ''' Construct from Rhino.Geometry.Vector3d
  ''' </summary>
  Sub New(ByVal src As Vector3d)
    X = src.X
    Y = src.Y
    Z = src.Z
  End Sub

  ''' <summary>
  ''' Returns an unset vector
  ''' </summary>
  Public Shared ReadOnly Property Unset() As SampleVbVector
    Get
      Return New SampleVbVector(UnsetValue, UnsetValue, UnsetValue)
    End Get
  End Property

  ''' <summary>
  ''' Assignment operator
  ''' </summary>
  Public Shared Widening Operator CType(ByVal src As Vector3d) As SampleVbVector
    Return New SampleVbVector(src)
  End Operator

  ''' <summary>
  ''' Assignment operator
  ''' </summary>
  Public Shared Widening Operator CType(ByVal src As SampleVbVector) As Vector3d
    Return New Vector3d(src.X, src.Y, src.Z)
  End Operator

  ''' <summary>
  ''' Equals operator 
  ''' </summary>
  Public Shared Operator =(ByVal p0 As SampleVbVector, ByVal p1 As SampleVbVector) As Boolean
    If (p0.X.Equals(p1.X) AndAlso p0.Y.Equals(p1.Y) AndAlso p0.Z.Equals(p1.Z)) Then
      Return True
    Else
      Return False
    End If
  End Operator

  ''' <summary>
  ''' Not equals operator
  ''' </summary>
  Public Shared Operator <>(ByVal p0 As SampleVbVector, ByVal p1 As SampleVbVector) As Boolean
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
      archive.WriteDouble(X)
      archive.WriteDouble(Y)
      archive.WriteDouble(Z)
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
    archive.Read3dmChunkVersion(major, minor)
    If (1 = major AndAlso 0 = minor) Then
      Try
        X = archive.ReadDouble()
        Y = archive.ReadDouble()
        Z = archive.ReadDouble()
        rc = Not archive.ReadErrorOccured
      Catch
        ' TODO...
      End Try
    End If
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
