Imports Rhino
Imports Rhino.Commands
Imports Rhino.Geometry
Imports Rhino.Input
Imports Rhino.Input.Custom

Namespace SampleVbCustomPoint

  ''' <summary>
  ''' SampleVbCustomPointCommand command
  ''' </summary>
  <System.Runtime.InteropServices.Guid("58221ff9-45c8-40a1-b4ba-be9217b98235")> _
  Public Class SampleVbCustomPointCommand
    Inherits Command

    Private Shared _instance As SampleVbCustomPointCommand

    ''' <summary>
    ''' Constructor
    ''' </summary>
    Public Sub New()
      ' Rhino only creates one instance of each command class defined in a
      ' plug-in, so it is safe to store a reference in a static field.
      _instance = Me
    End Sub

    ''' <summary>
    ''' Returns the only instance of the SampleVbCustomPointCommand command
    ''' </summary>
    Public Shared ReadOnly Property Instance() As SampleVbCustomPointCommand
      Get
        Return _instance
      End Get
    End Property

    ''' <summary>
    ''' The command name as it appears on the Rhino command line
    ''' </summary>
    Public Overrides ReadOnly Property EnglishName() As String
      Get
        Return "SampleVbCustomPointCommand"
      End Get
    End Property

    ''' <summary>
    ''' Called by Rhino when the user wants to run the command
    ''' </summary>
    Protected Overrides Function RunCommand(ByVal doc As RhinoDoc, ByVal mode As RunMode) As Result

      Dim point0, point1, point2 As Point3d
      Dim rc = Rhino.Input.RhinoGet.GetPoint("First point", False, point0)
      If (rc = Result.Success) Then
        rc = Rhino.Input.RhinoGet.GetPoint("Second point", False, point1)
        If (rc = Result.Success) Then
          rc = Rhino.Input.RhinoGet.GetPoint("Third point", False, point2)
          If (rc = Result.Success) Then
            SampleVbCustomPointPlugIn.Instance.Point0 = point0
            SampleVbCustomPointPlugIn.Instance.Point1 = point1
            SampleVbCustomPointPlugIn.Instance.Point2 = point2
            RhinoApp.WriteLine("Point 0 = {0}", SampleVbCustomPointPlugIn.Instance.Point0.ToString())
            RhinoApp.WriteLine("Point 1 = {0}", SampleVbCustomPointPlugIn.Instance.Point1.ToString())
            RhinoApp.WriteLine("Point 2 = {0}", SampleVbCustomPointPlugIn.Instance.Point2.ToString())
          End If
        End If
      End If

      Return Result.Success
    End Function
  End Class

End Namespace