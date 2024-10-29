import copy
import json
from importlib import reload
from libs.xml2excel.utils import cell
reload(cell)

def rendering(self, sheet, start_pos, data):
    pos = start_pos

    cell.setBoxBorder(self, sheet, 
        pos,
        data["@height"],
        self._start_col + self._body_lr_padding + self._page_lr_padding,
        self._width_size - self._body_lr_padding * 2 - self._page_lr_padding * 2)
    
    data["title"]["font_size"] = 32
    _tmp_height = len( data["title"]["value"].split("\n"))

    cell.setCell(self, sheet, 
        pos + data["@top_padding"], 4 * _tmp_height, 
        self._start_col + self._body_lr_padding + self._page_lr_padding + 1, 
        self._width_size - self._body_lr_padding * 2 - self._page_lr_padding * 2 - 2, 
        "FFFFFF", "center", data["title"])
    
    data["sub_title"]["font_size"] = 16
    _tmp_sub_height = len( data["sub_title"]["value"].split("\n"))

    cell.setCell(self, sheet, 
        pos + data["@top_padding"] + 4 * _tmp_height, 2 * _tmp_sub_height, 
        self._start_col + self._body_lr_padding + self._page_lr_padding + 1, 
        self._width_size - self._body_lr_padding * 2 - self._page_lr_padding * 2 - 2, 
        "FFFFFF", "center", data["sub_title"])

    pos = pos + data["@height"]
     
    # print(json.dumps(data, indent=2, ensure_ascii=False))

    return pos

def formatting(self, data):
    data = copy.deepcopy(data)

    data["@height"] = int(data["@height"])
    data["@top_padding"] = int(data["@top_padding"])

    for key in ["title", "sub_title"]:
        if key not in data:
            data[key] = ""

        texts = []
        if key in data and data[key] is not None:
            texts = data[key].split("\n")
            for index in range(len(texts)):
                texts[index] = texts[index].lstrip(" ").rstrip(" ").lstrip("\t").rstrip("\t")
        data[key] = "\n".join(texts)

        data[key] = {
            "value" : data[key],
        }

    return data
