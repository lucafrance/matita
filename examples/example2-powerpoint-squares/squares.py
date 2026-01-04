import random

from matita.office import powerpoint as pp

def squares():
    pp_app = pp.Application().new()
    pp_app.visible = True
    prs = pp_app.presentations.add()
    # Direct `Slides.add()` method unavailable
    # Will be fixed with https://github.com/MicrosoftDocs/VBA-Docs/pull/1937
    sld = pp.Slide(prs.slides.slides.Add(1, pp.ppLayoutBlank))

    for _ in range(10):
        side = random.random() * prs.pagesetup.slideheight / 3
        left = -side + random.random() * (side + prs.pagesetup.slidewidth)
        top = -side + random.random() * (side + prs.pagesetup.slideheight)
        shp = sld.shapes.addshape(pp.msoShapeRectangle, left, top, side, side)
        shp.line.visible = False
        shp.fill.forecolor.rgb = random.randint(0, 256 ** 3)
        eff = sld.timeline.mainsequence.addeffect(
            Shape=shp.shape, # TODO Fix how arguments are passed, so that .shape is not needed `TypeError: The Python instance can not be converted to a COM object`
            effectId=pp.msoAnimEffectFly,
            Level=pp.msoAnimateLevelNone, # TODO Passing None does not work in this case `TypeError: int() argument must be a string, a bytes-like object or a real number, not 'NoneType'`
            trigger=pp.msoAnimTriggerAfterPrevious,
        )
        direction = random.choice([
            pp.msoAnimDirectionLeft,
            pp.msoAnimDirectionTop,
            pp.msoAnimDirectionRight,
            pp.msoAnimDirectionBottom
        ])
        eff.effectparameters.direction = direction
        # The Timing.Duration property is no supported yet, because it can't be parsed from the documentation
        # Fill be fixed by:
        # - Addition of api_key: https://github.com/MicrosoftDocs/VBA-Docs/pull/1936
        # - Formatting adjustment: https://github.com/MicrosoftDocs/VBA-Docs/pull/1938
        eff.timing.timing.duration = 0.2

if __name__ == "__main__":
    squares()
