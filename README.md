# resetAllSlides

Removes all transitions, sounds, videos, animations, and most format from all slides in a PowerPoint presentation

Copyright (C) 2018 Fabio Correa correaduran.1@osu.edu

This is a PowerPoint macro-enabled presentation with one procedure, `ResetAllSlides`, with the following effects on each and every one of the slides in your presentation:

* Remove all slide transitions
* Delete all media objects (sounds, videos)
* Delete all animations

Make sure you only have opened two presentations: Presentation macros, and your presentation in which you want to perform these operations. Your use of this file will result in loss of data.

# Source code

Never run Office macros without exactly understanding what they do. Never open macro-enabled Office documents you got from the internet without making sure they come from a trusted source. Here is a copy of the contents of the ResetAllSlides procedure for you to inspect.

Sub resetAllSlides()
  Dim I As Integer
  Dim J As Integer
  With ActivePresentation ' In your presentation,
    For I = 1 To .Slides.Count ' For all slides,
      With .Slides(I)
        .SlideShowTransition.EntryEffect = ppEffectNone ' Remove the slide transition
        For J = .Shapes.Count To 1 Step -1
          With .Shapes(J)
            If (.Type = msoMedia) Then
              .Delete ' Delete the media object altogether
            End If
          End With
        Next J
        For J = .TimeLine.MainSequence.Count To 1 Step -1
          .TimeLine.MainSequence(J).Delete ' Delete the animation
        Next J
      End With
    Next I
  End With
End Sub

