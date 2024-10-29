import os
import zipfile
import shutil
import json
import pprint
import copy
from lxml import etree as ET # type: ignore
from pathlib import Path

class book:
    folder_path: str = ""
    basename: str = ""
    BASE_DIR: str = ""
    templates: dict = {}
    ns: dict = {}

def init(self, path, is_unzip=True): 

    self.book = book()
    self.book.BASE_DIR = "/workdir/libs/templates"
    self.book.templates = {
        "drawing": os.path.join(self.book.BASE_DIR, "new_drawing.xml"),
        "textbox": os.path.join(self.book.BASE_DIR, "new_textbox.xml"),
        "hishigata": os.path.join(self.book.BASE_DIR, "new_hishigata.xml"),
        "clear_textbox": os.path.join(self.book.BASE_DIR, "new_clear_textbox.xml"),
        "clear_textbox_bottomtext": os.path.join(self.book.BASE_DIR, "new_clear_textbox_bottomtext.xml"),
        "connect_arrow_bottom": os.path.join(self.book.BASE_DIR, "new_connect_arrow_bottom.xml"),
        "connect_arrow_bottom_down": os.path.join(self.book.BASE_DIR, "new_connect_arrow_bottom_down.xml"),
        "connect_arrow_left_bottom": os.path.join(self.book.BASE_DIR, "new_connect_arrow_left_bottom.xml"),
        "connect_arrow_right_bottom": os.path.join(self.book.BASE_DIR, "new_connect_arrow_right_bottom.xml"),
        "connect_arrow_right_up": os.path.join(self.book.BASE_DIR, "new_connect_arrow_right_up.xml"),
        "connect_arrow_right_up_over": os.path.join(self.book.BASE_DIR, "new_connect_arrow_right_up_over.xml"),
        "drawing_rels": os.path.join(self.book.BASE_DIR, "new_drawing.xml.rels"),
    }
    self.book.ns = {
        "drawing": "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}",
        "types":"{http://schemas.openxmlformats.org/package/2006/content-types}",
        "dmain":"{http://schemas.openxmlformats.org/drawingml/2006/main}",
        "sheet": "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}",
    }
    self.book.types = {
        "drawing":"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
    }

    self.book.folder_path = "/".join(path.split("/")[0:-1])
    self.book.basename = os.path.splitext(os.path.basename(path))[0]
    self.book.zip_filename = self.book.basename+".zip"

    if is_unzip:
        shutil.copyfile(path, self.book.folder_path + "/" + self.book.zip_filename)
        os.makedirs(self.book.folder_path + "/" + self.book.basename, exist_ok=True)
        with zipfile.ZipFile(self.book.folder_path + "/" + self.book.zip_filename) as zfile:
            for info in zfile.infolist():
                _rename(info)
                zfile.extract(info, self.book.folder_path + "/" + self.book.basename)
	
def _create_newdraw(self, draw_path, draw_name):

    _dir_path = self.book.folder_path + "/" + "/".join(draw_path.split("/")[0:-1])
    if os.path.isdir(_dir_path) is False :
        os.mkdir(_dir_path)

    shutil.copy(self.book.templates["drawing"], self.book.folder_path + "/" + draw_path)

    # update Content Type
    tree = ET.parse(os.path.join(self.book.folder_path + "/" + self.book.basename, "[Content_Types].xml"))
    root = tree.getroot()
    types = ET.SubElement(
        root,
        f"{self.book.ns['types']}Override",
        attrib={
            "PartName": f"/xl/drawings/{draw_name}",
            "ContentType": "application/vnd.openxmlformats-officedocument.drawing+xml",
        },
    )
    _write_xml(self, root, os.path.join(self.book.folder_path + "/" + self.book.basename, "[Content_Types].xml"))

def _write_xml(self, root, filename, is_debug=False):
    ET.ElementTree(root).write(filename, encoding="UTF-8", xml_declaration=True)
    if is_debug:
        print(Path(filename).read_text(encoding="utf-8"))

def _update_textbox(self, draw_path, text, anchor, _type="textbox", _id=0, _connect={}, _emu={}):
        
    # 新規構造を作成
    if _type == "hishigata":
        image = ET.parse(self.book.templates["hishigata"]).getroot()
    elif _type == "connect_arrow_bottom":
        image = ET.parse(self.book.templates["connect_arrow_bottom"]).getroot()
    elif _type == "connect_arrow_left_bottom":
        image = ET.parse(self.book.templates["connect_arrow_left_bottom"]).getroot()
    elif _type == "connect_arrow_right_bottom":
        image = ET.parse(self.book.templates["connect_arrow_right_bottom"]).getroot()
    elif _type == "connect_arrow_right_up":
        image = ET.parse(self.book.templates["connect_arrow_right_up"]).getroot()
    elif _type == "connect_arrow_right_up_over":
        image = ET.parse(self.book.templates["connect_arrow_right_up_over"]).getroot()
    elif _type == "connect_arrow_bottom_down":
        image = ET.parse(self.book.templates["connect_arrow_bottom_down"]).getroot()
    elif _type == "clear_textbox":
        image = ET.parse(self.book.templates["clear_textbox"]).getroot()
    elif _type == "clear_textbox_bottomtext":
        image = ET.parse(self.book.templates["clear_textbox_bottomtext"]).getroot()
    else :
        image = ET.parse(self.book.templates["textbox"]).getroot()
    points = ["from", "to"]
    props = ["col", "colOff", "row", "rowOff"]
    for item in image.findall(f"{self.book.ns['drawing']}twoCellAnchor"):
        for p in points:
            for c in item.findall(f"{self.book.ns['drawing']}{p}"):
                for prop in props:
                    for gc in c.findall(f"{self.book.ns['drawing']}{prop}"):
                        gc.text = str(anchor[p][prop])
        for c in item.iter(f"{self.book.ns['dmain']}t"):
            c.text = text
        for c in item.iter(f"{self.book.ns['drawing']}cNvPr"):
            c.attrib['id'] = str(_id)
        if _type == "connect_arrow_bottom" or \
            _type == "connect_arrow_left_bottom" or \
            _type == "connect_arrow_right_bottom" or \
            _type == "connect_arrow_right_up" or \
            _type == "connect_arrow_bottom_down" or \
            _type == "connect_arrow_right_up_over":
            for c in item.iter(f"{self.book.ns['dmain']}stCxn"):
                c.attrib['id'] = str(_connect["s_id"])
                c.attrib['idx'] = str(_connect["s_idx"])
            for c in item.iter(f"{self.book.ns['dmain']}endCxn"):
                c.attrib['id'] = str(_connect["e_id"])
                c.attrib['idx'] = str(_connect["e_idx"])
        if _type == "connect_arrow_right_up" or _type == "connect_arrow_right_up_over":
            for c in item.iter(f"{self.book.ns['dmain']}gd"):
                if c.attrib['name'] == "adj1":
                    c.attrib['fmla'] = "val " + str(_emu["max_width_cell_EMU"])
                if c.attrib['name'] == "adj2":
                    c.attrib['fmla'] = "val " + str(_emu["max_height_cell_EMU"])

    dr = ET.parse(self.book.folder_path + "/" + draw_path).getroot()

    for child1 in image.findall(f"{self.book.ns['drawing']}twoCellAnchor"):
        child2 = copy.deepcopy(child1)
        dr.append(child2)
        # 下記を一旦、上記で賄う、ただし、あえて下記の書き方をしている点を留意する可能性がある
        #child2 = ET.Element(child1.tag, child1.attrib)
        #child2.text = child1.text
        #dr.append(child2)
        #self._copy_subtree(child1, child2)
    _write_xml(self, dr, self.book.folder_path + "/" + draw_path, False)

def write_as_xlsx(self, filename=None, delete_wokdir=True):
    
    # .DS_STORE del
    for pathname, dirnames, filenames in os.walk(self.book.folder_path):
        for _filename in filenames:
             if _filename == ".DS_Store":
                os.remove(os.path.join(pathname,_filename))

    shutil.make_archive(
        self.book.folder_path + "/" + os.path.splitext(filename)[0], 
        format='zip', 
        root_dir=self.book.folder_path + "/" + self.book.basename)
    
    if os.path.isfile(self.book.folder_path + "/" + filename):
        os.remove(self.book.folder_path + "/" + filename)

    zipfile = os.path.splitext(filename)[0]+ ".zip"
    os.rename(self.book.folder_path + "/" + zipfile, self.book.folder_path + "/" + filename)

    if delete_wokdir:
        shutil.rmtree(self.book.folder_path + "/" + self.book.basename)
        if os.path.isfile(self.book.folder_path + "/" + self.book.basename+".zip"):
            os.remove(self.book.folder_path + "/" + self.book.basename+".zip")
    
def add_textbox(self, sheet_name, text, anchor={}, _type="textbox", _id=0, _connect={}, _emu={}):
    sheet_id = sheet_name.lstrip("sheet").rstrip(".xml")
    draw_name = f"drawing{sheet_id}.xml"
    draw_path = os.path.join(self.book.basename, "xl", "drawings", draw_name)
    if not os.path.exists(self.book.folder_path + "/" + draw_path):
        _create_newdraw(self, draw_path, draw_name)
    _update_textbox(self, draw_path, text, anchor, _type, _id, _connect, _emu)

    # update drawing rels file
    # drawings _rels の方は、画像を作るとできやつい
    sheet_rels_path = os.path.join(
        self.book.basename, "xl", "worksheets", "_rels", sheet_name + ".xml.rels"
    )
    if not os.path.exists(self.book.folder_path + "/" + sheet_rels_path):
        _dir_path = self.book.folder_path + "/" + "/".join(sheet_rels_path.split("/")[0:-1])
        if os.path.isdir(_dir_path) is False :
            os.mkdir(_dir_path)
        shutil.copy(self.book.templates["drawing_rels"], self.book.folder_path + "/" + sheet_rels_path)

        tree = ET.parse(self.book.folder_path + "/" + sheet_rels_path)
        root = tree.getroot()
        types = ET.SubElement(
            root,
            f"Relationship",
            attrib={
                "Id": f"rId{sheet_id}",
                "Type": self.book.types["drawing"],
                "Target": f"../drawings/drawing{sheet_id}.xml"
            },
        )
        ET.ElementTree(root).write(self.book.folder_path + "/" + sheet_rels_path, encoding="UTF-8", xml_declaration=True)

        worksheets_path = os.path.join(
            self.book.basename, "xl", "worksheets", "sheet" + sheet_id + ".xml"
        )
        tree = ET.parse(self.book.folder_path + "/" + worksheets_path)
        ET.register_namespace('r', "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
        root = tree.getroot()
        types = ET.SubElement(
            root,
            "drawing",
            attrib={
                f"{self.book.ns['sheet']}id": f"rId{sheet_id}"
            },
        )
        ET.ElementTree(root).write(self.book.folder_path + "/" + worksheets_path, encoding="UTF-8", xml_declaration=True)

def _rename(info: zipfile.ZipInfo):
    LANG_ENC_FLAG = 0x800
    encoding = 'utf-8' if info.flag_bits & LANG_ENC_FLAG else 'cp437'
    info.filename = info.filename.encode(encoding).decode('cp932')