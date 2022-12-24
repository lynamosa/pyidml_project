# encoding: utf-8

"""
Template xml used by text and related objects
"""

from pyidml.oxml import parse_xml
from lxml import etree
from pyidml.opc.serialized import ns

font_attribs = ["AppliedFont", "FontStyle", "PointSize", "Leading", "AppliedLanguage", "KerningMethod",
                "Tracking", "Capitalization", "Position", "Ligatures", "NoBreak", "HorizontalScale",
                "VerticalScale", "BaselineShift", "Skew", "FillColor", "FillTint", "StrokeTint",
                "StrokeWeight", "OverprintStroke", "OverprintFill", "OtfFigureStyle", "OtfOrdinal",
                "OtfFraction", "OtfDiscretionaryLigature", "OtfTitling", "OtfContextualAlternate",
                "OtfSwash", "OtfSlashedZero", "OtfHistorical", "OtfStylisticSets", "StrikeThru",
                "StrikeThroughColor", "StrikeThroughGapColor", "StrikeThroughGapOverprint",
                "StrikeThroughGapTint", "StrikeThroughOffset", "StrikeThroughOverprint",
                "StrikeThroughTint", "StrikeThroughType", "StrikeThroughWeight", "StrokeColor",
                "Underline", "UnderlineColor", "UnderlineGapColor", "UnderlineGapOverprint",
                "UnderlineGapTint", "UnderlineOffset", "UnderlineOverprint", "UnderlineTint",
                "UnderlineType", "UnderlineWeight"]
para_attribs = []

class story():
    """
    Template of story
    """
    temp_str = """<idPkg:Story xmlns:idPkg="http://ns.adobe.com/AdobeInDesign/idml/1.0/packaging" DOMVersion="14.0">
	<Story Self="udd" AppliedTOCStyle="n" UserText="true" IsEndnoteStory="false" TrackChanges="false" StoryTitle="$ID/" AppliedNamedGrid="n">
		<StoryPreference OpticalMarginAlignment="false" OpticalMarginSize="12" FrameType="TextFrameType" StoryOrientation="Horizontal" StoryDirection="LeftToRightDirection" />
		<InCopyExportOption IncludeGraphicProxies="true" IncludeAllResources="false" />
		<ParagraphStyleRange AppliedParagraphStyle="ParagraphStyle/$ID/NormalParagraphStyle">
		</ParagraphStyleRange>
	</Story>
</idPkg:Story>"""
    def __init__(self, _content, _para_attribs, _font_attribs):
        self.temp_xml:etree._Element = parse_xml(self.temp_str)
        self.para:etree._Element = self.temp_xml.xpath('/idPkg:Story/Story/ParagraphStyleRange', namespaces=ns['idPkg'])[0]
        for x in _para_attribs:
            if x in para_attribs:
                para.attrib[x] = _para_attribs[x]

        content = character(_content, _font_attribs)
        self.para.append(content)

    def add_content(self, _content, _font_attribs):
        content = character(_content, _font_attribs)
        self.para.append(content)


class character():
    def __new__(cls, _content, _font_attribs):
        cls.root:etree._Element = etree.XML(f'<CharacterStyleRange AppliedCharacterStyle="CharacterStyle/$ID/[No character style]"><Content>{_content}</Content></CharacterStyleRange>')
        for x in _font_attribs:
            if x in font_attribs:
                cls.root.attrib[x] = str(_font_attribs[x])
        return cls.root
