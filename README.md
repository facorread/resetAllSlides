# resetAllSlides

Removes all transitions, sounds, videos, animations, and most format from all slides in a PowerPoint presentation

Copyright (C) 2018 Fabio Correa correaduran.1@osu.edu

This PowerPoint macro-enabled presentation contains one procedure, `resetAllSlides`, with the following effects on each and every one of the slides in your presentation:

* Remove all slide transitions
* Delete all media objects (sounds, videos)
* Delete all animations
* Set all text to Century Gothic regular
* Set all lines to white 3pt
* Remove borders from all shapes except for tables

Make sure you only have opened two presentations: Presentation macros, and your presentation in which you want to perform these operations. Your use of this file will result in loss of data.

Never run Office macros without exactly understanding what they do. Never open macro-enabled Office documents you got from the internet without making sure they come from a trusted source.

# Copy of the source code

```vb
Sub resetAllSlides()
  Dim I As Integer
  Dim J As Integer
  With ActivePresentation ' In your presentation,
    For I = .Fonts.Count To 1 Step -1 ' Count backwards because .Count changes
      With .Fonts.Item(I)
        .Name = "Century Gothic"
        .Bold = msoFalse
        'MsgBox .Color.RGB ' Color handling has to be manual
        'If (.Color.RGB = -2147483648#) Then
        '  .Color.RGB = vbWhite
        'End If
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
                  .ParagraphFormat.IndentLevel = 1 ' First level paragraphs with no indentation
                End With
              End If
            End If
            Select Case .Type
              Case msoMedia
                .Delete ' Delete the media object altogether
              Case msoLine
                .Line.ForeColor.RGB = vbWhite
                .Line.Weight = 3
              Case msoTable, msoPlaceholder
              Case Else
                If (Not .Line.Visible) Then ' Only borders not previously marked as invisible
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
```
