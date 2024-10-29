import copy
import importlib
from libs.xml2excel.utils import cell, util
for _m in [util, cell]:
    importlib.reload(_m)

def rendering(self, ws, data):

    if "header" not in data["sheet"]:
        return 1

    if len(data["sheet"]["header"]["head"]) > 3:
        _head_pad = 36
    elif len(data["sheet"]["header"]["head"]) > 0:
        _head_pad = 18
    else :
        _head_pad = 0

    main_header_width = self._width_size - self._header_lr_padding * 2 - self._page_lr_padding * 2 - _head_pad

    cell.setCell(self, ws, 2, 4, 
                            self._start_col + self._header_lr_padding + self._page_lr_padding, 
                            main_header_width,
                            "FFFFFF", "center", {
                                "value" : data["sheet"]["header"]["doc_title"]["value"],
                                "font_size" : 14
                            })
    cell.setBoxBorder(self, ws, 2, 4, 
                            self._start_col + self._header_lr_padding + self._page_lr_padding, 
                            self._width_size - self._header_lr_padding * 2 - self._page_lr_padding * 2)

    pos = 2
    for i in range(4):
        if len(data["sheet"]["header"]["head"]) - 1 >= i:
            cell.setCellwithBorder(self, ws, pos + i, 1, 
                                self._start_col + self._header_lr_padding + self._page_lr_padding + main_header_width, 
                                9, 
                                self._background_color, "center", data["sheet"]["header"]["head"][i]["@key"])
            cell.setCellwithBorder(self, ws, pos + i, 1, 
                                self._start_col + self._header_lr_padding + self._page_lr_padding + main_header_width + 9, 
                                9, 
                                "FFFFFF", "left", data["sheet"]["header"]["head"][i]["#text"])
    pos = 2
    for i in range(4, 8):
        if len(data["sheet"]["header"]["head"]) - 1 >= i:
            cell.setCellwithBorder(self, ws, pos + i - 4, 1, 
                                self._start_col + self._header_lr_padding + self._page_lr_padding + main_header_width + 18, 
                                9, 
                                self._background_color, "center", data["sheet"]["header"]["head"][i]["@key"])
            cell.setCellwithBorder(self, ws, pos + i - 4, 1, 
                                self._start_col + self._header_lr_padding + self._page_lr_padding + main_header_width + 27, 
                                9,
                                "FFFFFF", "left", data["sheet"]["header"]["head"][i]["#text"])

    return 6

def formatting(self, data):
    data = copy.deepcopy(data)

    data["doc_title"] = {
        "value" : data["doc_title"]["#text"],
    }

    if "head" not in data:
        data["head"] = []
    elif isinstance(data["head"], dict):
        data["head"] = [data["head"]]
    for key, val in enumerate(data["head"]) :
        data["head"][key] = {
            "@key" : {
                "value" : val["@key"]
            },
            "#text" : {
                "value" : val["#text"]
            },
        }

    return data