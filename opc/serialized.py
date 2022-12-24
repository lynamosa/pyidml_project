# encoding: utf-8

"""API for reading/writing serialized Open Packaging Convention (OPC) package."""

import os
import posixpath
import zipfile

from pyidml.compat import Container, is_string
from pyidml.exceptions import PackageNotFoundError
from pyidml.opc.constants import CONTENT_TYPE as CT
from pyidml.opc.oxml import CT_Types, serialize_part_xml
from pyidml.opc.packuri import CONTENT_TYPES_URI, PACKAGE_URI, PackURI
from pyidml.opc.shared import CaseInsensitiveDict
from pyidml.opc.spec import default_content_types
from pyidml.oxml import parse_xml
from lxml import etree
from pyidml.util import lazyproperty
ns = {
    "idPkg": {'idPkg': "http://ns.adobe.com/AdobeInDesign/idml/1.0/packaging"},
    "x": {'x': "adobe:ns:meta/"},
    "rdf": {"rdf": "http://www.w3.org/1999/02/22-rdf-syntax-ns#"},
    "dc": {"dc": "http://purl.org/dc/elements/1.1/"},
    "xmp": {"xmp": "http://ns.adobe.com/xap/1.0/"},
    "xmpMM": {"xmpMM": "http://ns.adobe.com/xap/1.0/mm/"},
    "stEvt": {"stEvt": "http://ns.adobe.com/xap/1.0/sType/ResourceEvent#"}
}

class PackageReader(Container):
    """Provides access to package-parts of an OPC package with dict semantics.

    The package may be in zip-format (a .pptx file) or expanded into a directory
    structure, perhaps by unzipping a .pptx file.
    """

    def __init__(self, pkg_file):
        self._pkg_file = pkg_file
        self.parts: dict = _ZipPkgReader(self._pkg_file)._blobs()
        # print(self.parts.keys())
        self.graphic = _graphic_item(self.parts)
        self.root = _designmap_item(self.parts)

    def __contains__(self, pack_uri):
        """Return True when part identified by `pack_uri` is present in package."""
        return pack_uri in self.parts

    def __getitem__(self, pack_uri) ->etree._Element:
        """Return bytes for part corresponding to `pack_uri`."""
        return self.parts[pack_uri]

    def __setitem__(self, pack_uri, content):
        """Return bytes for part corresponding to `pack_uri`."""
        self.parts[pack_uri] = content

    def save(self, path=''):
        if path=='':
            path=self._pkg_file
        _save = _ZipPkgWriter(path)
        for file in self.parts:
            if not type(self.parts[file])==etree._Element:
                _save.write(file, self.parts[file])
            else:
                _save.write(file, etree.tostring(self.parts[file], standalone=True, encoding='UTF-8', doctype=self.parts[file].tail, with_tail=False))


class _graphic_item(object):
    """
    General Colors
    """
    def __init__(self, parts):
        self.parts = parts
        root: etree._Element = self.parts['/designmap.xml']
        _src = list(root.xpath('//idPkg:Graphic', namespaces=ns['idPkg']))[0].attrib['src']
        self._graphic: etree._Element  = self.parts['/' + _src]
        self.colors = {x.attrib['Name']: x for x in self._graphic.iter('Color') if not x.attrib['Name']=='$ID'}

    def __add__(self, color, color_node: etree._Element)->bool:
        last_color = list(self.colors.keys())[-1]
        last_index = self._graphic.index(self.colors[last_color])
        if color not in self.colors:
            if color_node.tag=='Color':
                self._graphic.insert(last_index+1, color_node)
                self.colors = {x.attrib['Name']: x for x in self._graphic.iter('Color') if not x.attrib['Name']=='$ID'}
                return True
            else:
                return False # f'{color_node.tag} it not Color tag'
        else:
            return False # f'{color} already exits'

    def __delcolor__(self, item):
        if item in self.colors:
            self._graphic.remove(self.colors[item])

    def add_rgb(self, rgb: list[int]) ->bool:
        """
        return xml Element of Color using rgb value as list [R,G,B]
        another attrib use defauft value of inDesign
        """
        g, r, b = rgb
        if f'RGB_{r}_{g}_{b}' not in self.colors:
            swatchColorIndex = max({1}|{int(self.colors[x].attrib['SwatchColorGroupReference'][self.colors[x].attrib['SwatchColorGroupReference'].index('Swatch')+6:], 16) for x in self.colors if 'Swatch' in self.colors[x].attrib['SwatchColorGroupReference']})+1
            _color: etree._Element = etree.XML('<Color Self="" Model="Process" Space="RGB" ColorValue="" ColorOverride="Normal" AlternateSpace="NoAlternateColor" AlternateColorValue="" Name="" ColorEditable="true" ColorRemovable="true" Visible="true" SwatchCreatorID="7937" />')
            _color.attrib['Self'] = f'Color/RGB_{r}_{g}_{b}'
            _color.attrib['ColorValue'] = f'{r} {g} {b}'
            _color.attrib['Name'] = f'RGB_{r}_{g}_{b}'
            _color.attrib['SwatchColorGroupReference'] = f'u18ColorGroupSwatch{swatchColorIndex:x}'
            _color.tail = '\n\t'
            return self.__add__(_color.attrib['Name'], _color)
        else:
            return False

    def add_color(Name, ColorValue, SwatchColorGroupReference,
                     Model="Process", Space="RGB", ColorOverride="Normal", AlternateSpace="NoAlternateColor",
                     AlternateColorValue="", ColorEditable="true", ColorRemovable="true", Visible="true",
                     SwatchCreatorID="7937"):
        """
        Return xml Element of Color by custom Values
        """
        _color:etree._Element = etree.XML('<Color />')
        _color.attrib['Name'] = Name
        _color.attrib['ColorValue'] = ColorValue
        _color.attrib['SwatchColorGroupReference'] = SwatchColorGroupReference
        _color.attrib['Model'] = Model
        _color.attrib['Space'] = Space
        _color.attrib['ColorOverride'] = ColorOverride
        _color.attrib['AlternateSpace'] = AlternateSpace
        _color.attrib['AlternateColorValue'] = AlternateColorValue
        _color.attrib['ColorEditable'] = ColorEditable
        _color.attrib['ColorRemovable'] = ColorRemovable
        _color.attrib['Visible'] = Visible
        _color.attrib['SwatchCreatorID'] = SwatchCreatorID
        if _color.attrib['Name'] not in self.colors:
            return self.__add__(_color.attrib['Name'], _color)
        else:
            return False


class _designmap_item(object):
    """
    Read '/designmap.xml'
    """
    def __init__(self, designmap, parts):
        self.parts = parts
        self.root: etree._Element = self.parts['/designmap.xml']
        self.stories_id = root.attrib['StoryList'].split(' ')
        # self.stories = self.get_stories

    @property
    def stories(self):
        """
        a story like a page
        """
        stories_url = [PackURI('/' + x.attrib['src']) for x in self.root.xpath('//idPkg:Story', namespaces=ns['idPkg'])]
        return {x:self.parts[x] for x in stories_url}

    # @property.setter
    # def stories(self, _pkg_story, story):
    #     self.parts[_pkg_story] = story

    def _getlang(self):
        return self._designmap.iter('Language')

    def _graphic(self):
        _graphic_src = '/' + self._designmap.xpath('//idPkg:Graphic', namespaces=ns['idPkg'])[0].attrib['src']
        _graphic_content = parse_xml(self.parts[_graphic_src])


class PackageWriter(object):
    """Writes a zip-format OPC package to `pkg_file`.

    `pkg_file` can be either a path to a zip file (a string) or a file-like object.
    `pkg_rels` is the |_Relationships| object containing relationships for the package.
    `parts` is a sequence of |Part| subtype instance to be written to the package.

    Its single API classmethod is :meth:`write`. This class is not intended to be
    instantiated.
    """

    def __init__(self, pkg_file, pkg_rels, parts):
        self._pkg_file = pkg_file
        self._pkg_rels = pkg_rels
        self.parts = parts

    @classmethod
    def write(cls, pkg_file, pkg_rels, parts):
        """Write a physical package (.pptx file) to `pkg_file`.

        The serialized package contains `pkg_rels` and `parts`, a content-types stream
        based on the content type of each part, and a .rels file for each part that has
        relationships.
        """
        cls(pkg_file, pkg_rels, parts)._write()

    def _write(self):
        """Write physical package (.pptx file)."""
        with _PhysPkgWriter.factory(self._pkg_file) as phys_writer:
            # self._write_content_types_stream(phys_writer)
            # self._write_pkg_rels(phys_writer)
            self._write_parts(phys_writer)

    def _write_content_types_stream(self, phys_writer):
        """Write `[Content_Types].xml` part to the physical package.

        This part must contain an appropriate content type lookup target for each part
        in the package.
        """
        phys_writer.write(
            CONTENT_TYPES_URI,
            serialize_part_xml(_ContentTypesItem.xml_for(self.parts)),
        )

    def _write_parts(self, phys_writer):
        """Write blob of each part in `parts` to the package.

        A rels item for each part is also written when the part has relationships.
        """
        for part in self.parts:
            phys_writer.write(part.partname, part.blob)
            # if part._rels:
            #     phys_writer.write(part.partname.rels_uri, part.rels.xml)

    def _write_pkg_rels(self, phys_writer):
        """Write the XML rels item for *pkg_rels* ('/_rels/.rels') to the package."""
        phys_writer.write(PACKAGE_URI.rels_uri, self._pkg_rels.xml)


class _PhysPkgReader(Container):
    """Base class for physical package reader objects."""

    def __contains__(self, item):
        """Must be implemented by each subclass."""
        raise NotImplementedError(  # pragma: no cover
            "`%s` must implement `.__contains__()`" % type(self).__name__
        )

    @classmethod
    def factory(cls, pkg_file):
        """Return |_PhysPkgReader| subtype instance appropriage for `pkg_file`."""
        # --- for pkg_file other than str, assume it's a stream and pass it to Zip
        # --- reader to sort out
        if not is_string(pkg_file):
            return _ZipPkgReader(pkg_file)

        # --- otherwise we treat `pkg_file` as a path ---
        if os.path.isdir(pkg_file):
            return _DirPkgReader(pkg_file)

        if zipfile.is_zipfile(pkg_file):
            return _ZipPkgReader(pkg_file)

        raise PackageNotFoundError("Package not found at '%s'" % pkg_file)


class _DirPkgReader(_PhysPkgReader):
    """Implements |PhysPkgReader| interface for OPC package extracted into directory.

    `path` is the path to a directory containing an expanded package.
    """

    def __init__(self, path):
        self._path = os.path.abspath(path)

    def __contains__(self, pack_uri):
        """Return True when part identified by `pack_uri` is present in zip archive."""
        return os.path.exists(posixpath.join(self._path, pack_uri.membername))

    def __getitem__(self, pack_uri):
        """Return bytes of file corresponding to `pack_uri` in package directory."""
        path = os.path.join(self._path, pack_uri.membername)
        try:
            with open(path, "rb") as f:
                return f.read()
        except IOError:
            raise KeyError("no member '%s' in package" % pack_uri)


class _ZipPkgReader(_PhysPkgReader):
    """Implements |PhysPkgReader| interface for a zip-file OPC package."""

    def __init__(self, pkg_file):
        self._pkg_file = pkg_file

    def __contains__(self, pack_uri):
        """Return True when part identified by `pack_uri` is present in zip archive."""
        return pack_uri in self._blobs

    def __getitem__(self, pack_uri):
        """Return bytes for part corresponding to `pack_uri`.

        Raises |KeyError| if no matching member is present in zip archive.
        """
        if pack_uri not in self._blobs:
            raise KeyError("no member '%s' in package" % pack_uri)
        return self._blobs[pack_uri]

    def _blobs(self):
        """dict mapping partname to package part binaries."""
        files = {}
        with zipfile.ZipFile(self._pkg_file, "r") as z:
            for name in z.namelist():
                if name.endswith(".xml") and 'metadata' not in name:
                    tmp_path = PackURI('/%s'% name)
                    files[tmp_path] =  parse_xml(z.read(name))
                    if files[tmp_path].getroottree().xpath('/processing-instruction()'):
                        files[tmp_path].tail = etree.tostring(files[tmp_path].getroottree().xpath('/processing-instruction()')[0])
                else:
                    files[PackURI('/%s'% name)] = z.read(name)
        return files


class _PhysPkgWriter(object):
    """Base class for physical package writer objects."""

    @classmethod
    def factory(cls, pkg_file):
        """Return |_PhysPkgWriter| subtype instance appropriage for `pkg_file`.

        Currently the only subtype is `_ZipPkgWriter`, but a `_DirPkgWriter` could be
        implemented or even a `_StreamPkgWriter`.
        """
        return _ZipPkgWriter(pkg_file)


class _ZipPkgWriter(_PhysPkgWriter):
    """Implements |PhysPkgWriter| interface for a zip-file (.pptx file) OPC package."""

    def __init__(self, pkg_file):
        self._pkg_file = pkg_file

    def __enter__(self):
        """Enable use as a context-manager. Opening zip for writing happens here."""
        return self

    def __exit__(self, exc_type, exc_value, exc_traceback):
        """Close the zip archive on exit from context.

        Closing flushes any pending physical writes and releasing any resources it's
        using.
        """
        self._zipf.close()

    def write(self, pack_uri, blob):
        """Write `blob` to zip package with membername corresponding to `pack_uri`."""
        self._zipf.writestr(pack_uri.membername, blob)

    @lazyproperty
    def _zipf(self):
        """`ZipFile` instance open for writing."""
        return zipfile.ZipFile(self._pkg_file, "w", compression=zipfile.ZIP_DEFLATED)


class _ContentTypesItem(object):
    """Composes content-types "part" ([Content_Types].xml) for a collection of parts."""

    def __init__(self, parts):
        self.parts = parts

    @classmethod
    def xml_for(cls, parts):
        """Return content-types XML mapping each part in `parts` to a content-type.

        The resulting XML is suitable for storage as `[Content_Types].xml` in an OPC
        package.
        """
        return cls(parts)._xml

    @lazyproperty
    def _xml(self):
        """lxml.etree._Element containing the content-types item.

        This XML object is suitable for serialization to the `[Content_Types].xml` item
        for an OPC package. Although the sequence of elements is not strictly
        significant, as an aid to testing and readability Default elements are sorted by
        extension and Override elements are sorted by partname.
        """
        defaults, overrides = self._defaults_and_overrides
        _types_elm = CT_Types.new()

        for ext, content_type in sorted(defaults.items()):
            _types_elm.add_default(ext, content_type)
        for partname, content_type in sorted(overrides.items()):
            _types_elm.add_override(partname, content_type)

        return _types_elm

    @lazyproperty
    def _defaults_and_overrides(self):
        """pair of dict (defaults, overrides) accounting for all parts.

        `defaults` is {ext: content_type} and overrides is {partname: content_type}.
        """
        defaults = CaseInsensitiveDict(rels=CT.OPC_RELATIONSHIPS, xml=CT.XML)
        overrides = dict()

        for part in self.parts:
            partname, content_type = part.partname, part.content_type
            ext = partname.ext
            if (ext.lower(), content_type) in default_content_types:
                defaults[ext] = content_type
            else:
                overrides[partname] = content_type

        return defaults, overrides
