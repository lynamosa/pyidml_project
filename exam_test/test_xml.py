from lxml import etree
from xml.dom.minidom import parseString
from pyidml.opc.serialized import ns

element_class_lookup = etree.ElementNamespaceClassLookup()
oxml_parser = etree.XMLParser(strip_cdata=False, resolve_entities=False)
oxml_parser.set_element_class_lookup(element_class_lookup)
fName = 'designmap.xml'
path = 'data/Coffee-Obsession_content/' + fName
path2 = 'data/addcolor/' + fName
f = open(path, 'rb').read()
d: etree._Element = etree.fromstring(f, oxml_parser)
x = d.getroottree()
if d.getroottree().xpath('/processing-instruction()'):
    d.tail = etree.tostring(d.getroottree().xpath('/processing-instruction()')[0])
stories_id = d.attrib['StoryList'].split(' ')
spreads_url = ['/'+x.attrib['src'] for x in d.xpath('//idPkg:Spread', namespaces=ns['idPkg'])]
stories_url = ['/'+x.attrib['src'] for x in d.xpath('//idPkg:Story', namespaces=ns['idPkg'])]
# print(len(stories_id), stories_id)
print(len(stories_url), len(stories_id))

with open(path2, 'wb') as pp:
    pp.write(etree.tostring(d, standalone=True, encoding='UTF-8', doctype=d.tail, with_tail=False))
# print(d.tail)
# print(d.info)
