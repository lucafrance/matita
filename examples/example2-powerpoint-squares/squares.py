# TODO, convert VBA to Python
# Sub squares()

#     Dim ppt As Presentation
#     Dim sld As Slide
#     Set ppt = Application.Presentations.Add()
#     Set sld = ppt.Slides.Add(1, ppLayoutBlank)
    
#     Dim i As Integer
#     Dim left, top, side As Single
#     Dim shp As Shape
#     Dim eff As Effect
#     Dim dirs(1 To 4) As MsoAnimDirection
#     dirs(1) = msoAnimDirectionLeft
#     dirs(2) = msoAnimDirectionTop
#     dirs(3) = msoAnimDirectionRight
#     dirs(4) = msoAnimDirectionBottom
    
#     i = 1000
#     While i > 0
#         side = Rnd() * ppt.PageSetup.SlideHeight / 3
#         left = -side + Rnd() * (side + ppt.PageSetup.SlideWidth)
#         top = -side + Rnd() * (side + ppt.PageSetup.SlideHeight)
#         Set shp = sld.Shapes.AddShape(msoShapeRectangle, left, top, side, side)
#         shp.Line.Visible = msoFalse
#         shp.Fill.ForeColor.RGB = CLng(256 ^ 3 * Rnd() + 1)
#         Set eff = sld.TimeLine.MainSequence.AddEffect(Shape:=shp, effectid:=msoAnimEffectFly, Trigger:=msoAnimTriggerAfterPrevious)
#         eff.EffectParameters.Direction = dirs(CInt(Int(Rnd() * 4) + 1))
#         eff.Timing.Duration = 0.2
#         i = i - 1
#     Wend
    
# End Sub
