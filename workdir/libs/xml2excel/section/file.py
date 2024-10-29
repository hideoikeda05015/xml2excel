import copy
import os
import importlib
import math
import shutil
import uuid
import openpyxl # type: ignore
from openpyxl.utils.units import pixels_to_EMU, cm_to_EMU, pixels_to_points, EMU_to_pixels, dxa_to_inch # type: ignore
from openpyxl.drawing.spreadsheet_drawing import AbsoluteAnchor, OneCellAnchor, AnchorMarker # type: ignore
from openpyxl.drawing.xdr import XDRPositiveSize2D # type: ignore
from libs.xml2excel.utils import cell, util
from contextlib import redirect_stdout
for _m in [util, cell]:
    importlib.reload(_m)

def rendering(self, sheet, start_pos, data):
    pos = cell.setInlineTitle(self, sheet, start_pos, data["title"])
    layout = data["file_list"]["@layout_size"]["value"]
    _width_view_cell  = self._width_size - self._body_lr_padding * 2 - self._page_lr_padding * 2
    if self._image_offset_calc_type == "cell":
        _width_view_cell  = round(_width_view_cell / layout)
    else :
        _width_view_cell  = (_width_view_cell / layout) # EXCELでoffsetが効かない.....
    _bas_pos = self._start_col + self._body_lr_padding + self._page_lr_padding

    _image_start_pos = pos
    _height_row_step_stack = 0

    if "file_list" in data and "file" in data["file_list"]:
        max_pos = 0
        _height_row_step = 0
        for key, file_path in enumerate(data["file_list"]["file"]) :
            if key % layout == 0 :
                max_pos = 0
                _height_row_step = 0

            img = openpyxl.drawing.image.Image(file_path["value"])
            org_width =  img.width
            org_height = img.height
            _width_image_cell = pixels_to_points(org_width) / self._width_point_size # pt
            _height_image_cell = pixels_to_points(org_height) * _width_view_cell / _width_image_cell / self._height_point_size # pt
            img.width =  math.floor(org_width * (_width_view_cell / _width_image_cell))
            img.height =  math.floor(org_height * (_width_view_cell / _width_image_cell))

            if self._image_offset_calc_type == "cell":
                cell_address = sheet.cell(pos, _bas_pos + (_width_view_cell) * (key % layout)).coordinate
                img.anchor = cell_address
            else :
                col_offset = pixels_to_EMU((img.width) * (key % layout))
                if layout > 1:
                    row_offset = pixels_to_EMU((_height_row_step_stack))
                else :
                    row_offset = pixels_to_EMU((_height_row_step_stack))
                size_ext = XDRPositiveSize2D(pixels_to_EMU(img.width), pixels_to_EMU(img.height))
                maker = AnchorMarker(col=_bas_pos - 1, colOff=col_offset, row=_image_start_pos - 1, rowOff=row_offset)
                img.anchor = OneCellAnchor(_from=maker, ext=size_ext) # offset型

            if key % layout == layout - 1 :
                if _height_image_cell > max_pos :
                    if self._image_offset_calc_type == "cell":
                        pos = pos + round(_height_image_cell)
                    else :
                        pos = pos + (_height_image_cell) # offset型
                    _height_row_step_stack = _height_row_step_stack + img.height
                else :
                    pos = pos + max_pos
                    _height_row_step_stack = _height_row_step_stack + _height_row_step
            else :
                if _height_image_cell > max_pos :
                    if self._image_offset_calc_type == "cell":
                        max_pos = round(_height_image_cell)
                    else :
                        max_pos = (_height_image_cell) # offset型
                    _height_row_step = img.height
                if len(data["file_list"]["file"]) - 1 == key :
                    pos = pos + max_pos
                    _height_row_step_stack = _height_row_step_stack + _height_row_step
            sheet.add_image(img)
            self.diagram_number = self.diagram_number + 1
    
    if self._image_offset_calc_type != "cell":
        pos = round(pos) # offset型

    return pos

def formatting(self, data, file):
    data = copy.deepcopy(data)
    if "title" not in data:
        data["title"] = ""
    data["title"] = {
        "value" : data["title"],
    }

    if "file_list" in data and "@layout_size" in data["file_list"]:
            data["file_list"]["@layout_size"] = {
                "value" : int(data["file_list"]["@layout_size"])
            }
    else :
        data["file_list"]["@layout_size"] = {
                "value" : 1
            }

    if "file_list" in data and "file" in data["file_list"]:
        if isinstance(data["file_list"]["file"], dict):
            data["file_list"]["file"] = [data["file_list"]["file"]]
        file_path_array = []
        for tmp_file in data["file_list"]["file"]:
            file_path = ""
            if tmp_file["@src"][0] == "/": # 絶対タグ
                file_path = tmp_file["@src"][0:]
            elif tmp_file["@src"][0:2] == "./": # 同階層パス
                file_path = os.path.dirname(file) + tmp_file["@src"][1:]
            elif tmp_file["@src"][0:2] == "~/": # プロジェクトパス
                file_path = self._project_path + "/" + tmp_file["@src"][1:]
            elif tmp_file["@src"][0:3] == "../": # 同階層パス
                file_path = os.path.dirname(file) + "/" + tmp_file["@src"]
            
            if os.path.isfile(file_path):
                if ".jpeg" in file_path or ".png" in file_path :
                    file_path_array.append({
                        "value" : file_path
                    })
                if ".pu" in file_path :
                    file_path = createPlantUmlImage(self, file_path)
                    if os.path.isfile(file_path):
                        file_path_array.append({
                            "value" : file_path
                        })
                if ".mmdc" in file_path :
                    file_path = createMermaidImage(self, file_path)
                    if os.path.isfile(file_path):
                        file_path_array.append({
                            "value" : file_path
                        })
        data["file_list"]["file"] = file_path_array
    
    # print(json.dumps(data, indent=2, ensure_ascii=False))

    return data

def createMermaidImage(self, file_path):

    _uuid = str(uuid.uuid4())
    _dir_path = self._project_tmp_path + "/" + _uuid
    _file_path = _dir_path + "/" + file_path.split("/")[-1]
    os.mkdir(_dir_path)
    shutil.copy(file_path, _dir_path)

    _conf = "/workdir/libs/puppeteer-config.json"
    _tmp = "ssh -q -o StrictHostKeyChecking=no -p 22 mmdc_user@node \"mmdc -q -p %s -i %s -o %s\"" % (_conf, _file_path, _file_path.replace('.mmdc', '.png'))
    os.system(_tmp) >> 8
    
    return _file_path.replace('.mmdc', '.png')

def createPlantUmlImage(self, file_path):

    _uuid = str(uuid.uuid4())
    _dir_path = self._project_tmp_path + "/" + _uuid
    _file_path = _dir_path + "/" + file_path.split("/")[-1]
    os.mkdir(_dir_path)
    shutil.copy(file_path, _dir_path)

    test_data = open(_file_path, "r")
    _out_path = _file_path
    for line in test_data:
        if "@startuml" in line:
            _tmp = line.strip(" ").strip("　").strip("\t")
            for key, val in enumerate(_tmp.split(" ")):
                if val == "@startuml" and len(_tmp.split(" ")) >= key + 1:
                    _file = "/".join(_file_path.split("/")[0:-1])
                    _ext = _file_path.split("/")[-1].split(".")[1]
                    _out_path = _file + "/" + _tmp.split(" ")[key+1].strip(" ").strip("\n") + "." + _ext
                    break
    test_data.close()

    os.system("/usr/bin/java -jar /opt/plantuml/plantuml.jar %s" % _file_path)

    return _out_path.replace('.pu', '.png')