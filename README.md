# resetAllSlides

Removes all transitions, sounds, videos, animations, pictures, and most format from all slides in a PowerPoint presentation

Copyright (C) 2018 Fabio Correa correaduran.1@osu.edu

This PowerPoint macro-enabled presentation contains two procedures: (1) `resetAllSlides`, with the following effects on each and every one of the slides in your presentation:

* Remove all slide transitions
* Delete all media objects (sounds, videos)
* Delete all animations
* Set all text to Century Gothic regular
* Remove any empty lines of text
* Set all lines to 3pt solid white
* Remove borders from all shapes except for tables

(2) `removeAllPictures`, which removes all pictures from your presentation.

Make a complete backup of each and every one of your documents before using any Office macros.

Never open macro-enabled Office documents without making sure they come from a trusted source. Review them using PowerPoint Protected View and make sure you understand each and every instruction before you proceed any further.

Make sure you have two open presentations only: Presentation macros, and your presentation in which you want to perform these operations.

Your use of this file will result in loss of data for which only you will be responsible.

You have been warned!

# Copy of the source code

```vb
Sub resetAllSlides()
  Dim I As Integer
  Dim J As Integer
  Dim K As Integer
  With ActivePresentation ' In your presentation,
    For I = .Fonts.Count To 1 Step -1 ' Count backwards because .Count changes
      With .Fonts.Item(I)
        .Name = "Century Gothic" ' Change all text format to Century Gothic regular white
        .Bold = msoFalse
        .Color.RGB = vbWhite
      End With
    Next I
    For I = 1 To .Slides.Count ' For all slides,
      With .Slides(I)
        .SlideShowTransition.EntryEffect = ppEffectNone ' Remove the slide transition
        ' .FollowMasterBackground = msoTrue ' Most probably unneccessary.
        For J = .Shapes.Count To 1 Step -1
          With .Shapes(J)
            If (.HasTextFrame) Then
              If (.TextFrame2.HasText) Then
                With .TextFrame2.TextRange
                  .ParagraphFormat.IndentLevel = 1 ' Format all paragraphs as first level with no indentation
                  For K = .Lines.Count To 1 Step -1
                    If (.Lines(K, 1).Text = vbCr) Then
                      .Lines(K, 1).Delete ' Delete all empty lines
                    End If
                  Next K
                End With
              End If
            End If
            Select Case .Type
              Case msoMedia
                .Delete ' Delete all media objects altogether
              Case msoLine
                .Line.ForeColor.RGB = vbWhite ' Set all shape lines to 3pt solid white
                .Line.Weight = 3
              Case msoTable, msoPlaceholder
              Case Else
                If (.Line.Visible <> msoFalse) Then ' Only borders not previously marked as invisible
                  .Line.Visible = msoFalse ' No border
                End If
            End Select
          End With
        Next J
        For J = .TimeLine.MainSequence.Count To 1 Step -1
          .TimeLine.MainSequence(J).Delete ' Delete all animations
        Next J
      End With
    Next I
  End With
End Sub

Sub removeAllPictures()
  Dim I As Integer
  Dim J As Integer
  Dim K As Integer
  Dim ContinueIteration As Boolean
  With ActivePresentation ' In your presentation,
    For I = 1 To .Slides.Count ' For all slides,
      With .Slides(I)
        '.Select ' Debug statement
        .FollowMasterBackground = msoTrue
        Do
          ContinueIteration = False
          For J = .Shapes.Count To 1 Step -1
            With .Shapes(J)
              '.Select ' Debug statement
              'MsgBox ("Shape type " & .Type) ' Debug statement
              Select Case .Type
                Case msoPicture
                  .Delete ' Delete picture
                Case msoPlaceholder
                  If (.PlaceholderFormat.ContainedType = msoPicture) Then
                    .Delete ' Delete placeholder
                  End If
                Case msoGroup
                  .Ungroup
                  ContinueIteration = True
                  Exit For
              End Select
            End With
          Next J
        Loop While (ContinueIteration)
      End With
    Next I
  End With
End Sub

```
