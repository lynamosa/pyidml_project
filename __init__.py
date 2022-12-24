# encoding: utf-8

"""Initialization module for python-pptx package."""

__version__ = "0.6.21"


import pyidml.exc as exceptions
import sys

sys.modules["pyidml.exceptions"] = exceptions
del sys

from pyidml.api import Presentation  # noqa

from pyidml.opc.constants import CONTENT_TYPE as CT  # noqa: E402
from pyidml.opc.package import PartFactory  # noqa: E402
from pyidml.parts.chart import ChartPart  # noqa: E402
from pyidml.parts.coreprops import CorePropertiesPart  # noqa: E402
from pyidml.parts.image import ImagePart  # noqa: E402
from pyidml.parts.media import MediaPart  # noqa: E402
from pyidml.parts.presentation import PresentationPart  # noqa: E402
from pyidml.parts.slide import (  # noqa: E402
    NotesMasterPart,
    NotesSlidePart,
    SlideLayoutPart,
    SlideMasterPart,
    SlidePart,
)

content_type_to_part_class_map = {
    CT.PML_PRESENTATION_MAIN: PresentationPart,
    CT.PML_PRES_MACRO_MAIN: PresentationPart,
    CT.PML_TEMPLATE_MAIN: PresentationPart,
    CT.PML_SLIDESHOW_MAIN: PresentationPart,
    CT.OPC_CORE_PROPERTIES: CorePropertiesPart,
    CT.PML_NOTES_MASTER: NotesMasterPart,
    CT.PML_NOTES_SLIDE: NotesSlidePart,
    CT.PML_SLIDE: SlidePart,
    CT.PML_SLIDE_LAYOUT: SlideLayoutPart,
    CT.PML_SLIDE_MASTER: SlideMasterPart,
    CT.DML_CHART: ChartPart,
    CT.BMP: ImagePart,
    CT.GIF: ImagePart,
    CT.JPEG: ImagePart,
    CT.MS_PHOTO: ImagePart,
    CT.PNG: ImagePart,
    CT.TIFF: ImagePart,
    CT.X_EMF: ImagePart,
    CT.X_WMF: ImagePart,
    CT.ASF: MediaPart,
    CT.AVI: MediaPart,
    CT.MOV: MediaPart,
    CT.MP4: MediaPart,
    CT.MPG: MediaPart,
    CT.MS_VIDEO: MediaPart,
    CT.SWF: MediaPart,
    CT.VIDEO: MediaPart,
    CT.WMV: MediaPart,
    CT.X_MS_VIDEO: MediaPart,
}

PartFactory.part_type_for.update(content_type_to_part_class_map)

del (
    ChartPart,
    CorePropertiesPart,
    ImagePart,
    MediaPart,
    SlidePart,
    SlideLayoutPart,
    SlideMasterPart,
    PresentationPart,
    CT,
    PartFactory,
)
