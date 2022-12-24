from pyidml.enum.template import story
from lxml import etree


myStory = story("My heart will go on", [], {"AppliedFont": 'Arial', "FontStyle": 'Bold', "PointSize": 10})
myStory.add_content("My heart will go on 22", {"AppliedFont": 'Calibri', "FontStyle": 'Italic', "PointSize": 15})
with open('/data/Stories_demo.xml', 'wb') as pp:
    pp.write(etree.tostring(myStory.temp_xml, standalone=True, encoding='UTF-8', with_tail=False, pretty_print=True))
