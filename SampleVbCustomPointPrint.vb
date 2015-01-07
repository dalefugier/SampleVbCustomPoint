Imports System
Imports Rhino
Imports Rhino.Commands

Namespace SampleVbCustomPoint

  ''' <summary>
  ''' SampleVbCustomPointPrint command
  ''' </summary>
  <System.Runtime.InteropServices.Guid("124435c4-9275-43c2-bf88-9f57c028edb3")> _
  Public Class SampleVbCustomPointPrint
    Inherits Command

    Shared _instance As SampleVbCustomPointPrint

    ''' <summary>
    ''' Constructor
    ''' </summary>
    Public Sub New()
      ' Rhino only creates one instance of each command class defined in a
      ' plug-in, so it is safe to store a reference in a static field.
      _instance = Me
    End Sub

    ''' <summary>
    ''' Returns the only instance of the SampleVbCustomPointPrint command
    ''' </summary>
    Public Shared ReadOnly Property Instance() As SampleVbCustomPointPrint
      Get
        Return _instance
      End Get
    End Property

    ''' <summary>
    ''' The command name as it appears on the Rhino command line
    ''' </summary>
    Public Overrides ReadOnly Property EnglishName() As String
      Get
        Return "SampleVbCustomPointPrint"
      End Get
    End Property

    ''' <summary>
    ''' Called by Rhino when the user wants to run the command
    ''' </summary>
    Protected Overrides Function RunCommand(ByVal doc As RhinoDoc, ByVal mode As RunMode) As Result

      If (SampleVbCustomPointPlugIn.Instance.Point0.IsValid) Then
        RhinoApp.WriteLine("Point 0 = {0}", SampleVbCustomPointPlugIn.Instance.Point0.ToString())
      Else
        RhinoApp.WriteLine("Point 0 = <invalid>")
      End If

      If (SampleVbCustomPointPlugIn.Instance.Point1.IsValid) Then
        RhinoApp.WriteLine("Point 1 = {0}", SampleVbCustomPointPlugIn.Instance.Point1.ToString())
      Else
        RhinoApp.WriteLine("Point 1 = <invalid>")
      End If

      If (SampleVbCustomPointPlugIn.Instance.Point2.IsValid) Then
        RhinoApp.WriteLine("Point 2 = {0}", SampleVbCustomPointPlugIn.Instance.Point2.ToString())
      Else
        RhinoApp.WriteLine("Point 2 = <invalid>")
      End If

      Return Result.Success

    End Function
  End Class
End Namespace