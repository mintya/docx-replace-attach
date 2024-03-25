"""Attach a xlsx or docx to word"""
import base64
import os
import uuid
from io import BytesIO

from docx import Document
from docx.opc.part import Part
from docx.opc.constants import CONTENT_TYPE, RELATIONSHIP_TYPE
from docx.opc.oxml import parse_xml
from docx.opc.packuri import PackURI
from PIL import Image, ImageDraw, ImageFont


class _AttachmentType:
    XLSX = {
        'content_type': CONTENT_TYPE.SML_SHEET,
        'file_name': 'Microsoft_Excel____',
        'file_type': 'xlsx',
        'program_id': 'Excel.Sheet.12',
        'shape_width': '76',
        'shape_height': '48',
        'icon': (
            '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'
        )
    }

    WORD = {
        'content_type': CONTENT_TYPE.WML_DOCUMENT,
        'file_name': 'Microsoft_Word____',
        'file_type': 'docx',
        'program_id': 'Word.Document.12',
        'shape_width': '76',
        'shape_height': '48',
        'icon': (
            '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'
            'WiNuu1sAAAAAElFTkSuQmCC'
        )
    }


class _Attachment:

    def __init__(self, doc: Document, file_name: str, att_type: dict):
        self.file_name = file_name
        self.doc_part = doc.part
        self.pkg = doc.part.package
        with open(file_name, "rb") as f:
            self.blob = f.read()
        self.att_type = att_type

    def replace(self, variable: str):
        self._shape_rel_to()
        self._embedd_rel_to()
        obj_ele = self._build_obj_element()
        for p in self.doc_part.document.paragraphs:
            for r in p.runs:
                if r.text == '{' + variable + '}':
                    r.clear()
                    r.element.insert(0, obj_ele)

    def _shape_rel_to(self):
        shape_path = f'/word/media/icon_{self.att_type["file_type"]}_{str(len(self.pkg.parts))}.png'
        part_shape = None
        for part in self.pkg.parts:
            if part.partname == shape_path:
                part_shape = part
                break
        if not part_shape:
            part_shape = Part.load(
                PackURI(shape_path),
                CONTENT_TYPE.PNG, self._generate_icon(self.att_type['icon']),
                self.pkg
            )
            self.pkg.parts.append(part_shape)
        self.shape_rid = self.doc_part.relate_to(part_shape, RELATIONSHIP_TYPE.IMAGE)

    def _embedd_rel_to(self):
        file_name = self.att_type['file_name'] + str(len(self.pkg.parts)) + '.' + self.att_type['file_type']
        part_obj = Part.load(
            PackURI(f'/word/embeddings/{file_name}'),
            self.att_type['content_type'], self.blob, self.pkg
        )
        self.pkg.parts.append(part_obj)
        self.embedd_rid = self.doc_part.relate_to(part_obj, RELATIONSHIP_TYPE.PACKAGE)

    def _build_obj_element(self):
        shape_id = '_x0000_i20' + str(uuid.uuid1())
        shape_type_xml = """<v:shapetype id="_x0000_t79" coordsize="21600,21600" o:spt="75" o:preferrelative="t"
                                path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">
                                <v:stroke joinstyle="miter"/>
                                <v:formulas>
                                    <v:f eqn="if lineDrawn pixelLineWidth 0"/>
                                    <v:f eqn="sum @0 1 0"/>
                                    <v:f eqn="sum 0 0 @1"/>
                                    <v:f eqn="prod @2 1 2"/>
                                    <v:f eqn="prod @3 21600 pixelWidth"/>
                                    <v:f eqn="prod @3 21600 pixelHeight"/>
                                    <v:f eqn="sum @0 0 1"/>
                                    <v:f eqn="prod @6 1 2"/>
                                    <v:f eqn="prod @7 21600 pixelWidth"/>
                                    <v:f eqn="sum @8 21600 0"/>
                                    <v:f eqn="prod @7 21600 pixelHeight"/>
                                    <v:f eqn="sum @10 21600 0"/>
                                </v:formulas>
                                <v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>
                                <o:lock v:ext="edit" aspectratio="t"/>
                            </v:shapetype>"""
        w_object_xml = f"""<w:object xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                                    xmlns:v="urn:schemas-microsoft-com:vml"
                                    xmlns:o="urn:schemas-microsoft-com:office:office"
                                    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                                    w:dxaOrig="1520" w:dyaOrig="960">
                                    {shape_type_xml}
                                <v:shape id="{shape_id}" type="#_x0000_t79" alt="" 
                                    style="width:{self.att_type["shape_width"]}pt;height:{self.att_type["shape_height"]}pt;mso-width-percent:0;
                                    mso-height-percent:0;mso-width-percent:0;mso-height-percent:0" 
                                    o:ole="">
                                    <v:imagedata r:id="{self.shape_rid}" o:title=""/>
                                </v:shape>
                                <o:OLEObject Type="Embed" ProgID="{self.att_type["program_id"]}" ShapeID="{shape_id}" 
                                    DrawAspect="Icon" ObjectID="{shape_id}" r:id="{self.embedd_rid}">
                                    <o:FieldCodes>\\s</o:FieldCodes>\n
                                </o:OLEObject>
                            </w:object>"""
        return parse_xml(w_object_xml)

    def _generate_icon(self, base_64: str) -> bytes:
        title = os.path.basename(self.file_name)
        title2 = ''

        name, ex = os.path.splitext(title)
        if len(name) >= 14:
            title = name[:14]
            title2 = name[14:20] + ex

        base_image = Image.new('RGBA', (int(self.att_type["shape_width"]), int(self.att_type["shape_height"])))
        image = Image.open(BytesIO(base64.b64decode(base_64)))
        image.thumbnail((30, 30))

        font = ImageFont.truetype('Arial', 8)
        draw = ImageDraw.Draw(base_image, 'RGBA')
        font_length = font.getlength(title)
        draw.text((int((base_image.size[0] - font_length) / 2), 32), title, font=font, fill=(0, 0, 0, 255))
        if title2:
            font_length2 = font.getlength(title2)
            draw.text((int((base_image.size[0] - font_length2) / 2), 40), title2, font=font, fill=(0, 0, 0, 255))
        base_image.paste(image, (int((base_image.size[0] - image.size[0]) / 2), 0))
        f = BytesIO()
        base_image.save(f, 'PNG')
        return f.getvalue()


def _replace_attachment(doc: Document, variable: str, file_name: str, att_type: dict):
    _Attachment(doc, file_name, att_type).replace(variable)


def replace_xlsx(doc: Document, variable: str, file_name: str) -> None:
    """Attach a xlsx to word"""
    _replace_attachment(doc, variable, file_name, _AttachmentType.XLSX)


def replace_word(doc: Document, variable: str, file_name: str) -> None:
    """Attach a docx to word"""
    _replace_attachment(doc, variable, file_name, _AttachmentType.WORD)


def replace_xlsx_t(tpl_file: str, new_file: str, variable: str, file_name: str) -> None:
    """Attach a xlsx to word"""
    doc = Document(tpl_file)
    _replace_attachment(doc, variable, file_name, _AttachmentType.XLSX)
    doc.save(new_file)


def replace_word_t(tpl_file: str, new_file: str, variable: str, file_name: str) -> None:
    """Attach a docx to word"""
    doc = Document(tpl_file)
    _replace_attachment(doc, variable, file_name, _AttachmentType.WORD)
    doc.save(new_file)
