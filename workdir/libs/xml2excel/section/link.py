import copy
from importlib import reload
from libs.xml2excel.utils import cell
reload(cell)

def rendering(self, sheet, start_pos, data):
    pos = cell.setInlineTitle(self, sheet, start_pos, data["title"])

    if len(data["ul"]) == 0:
        pos = pos - 1
    else :
        for key, val in enumerate(data["ul"]["li"]):
            cell.setCell(self, sheet, 
                pos, 1, 
                self._start_col + self._body_lr_padding + self._page_lr_padding, 
                10, 
                "CFE2F3", "left", val["@title"])
            pos = pos + 1
            for k, v in enumerate(val["link"]):
                cell.setCell(self, sheet, 
                    pos, 1, 
                    self._start_col + self._body_lr_padding + self._page_lr_padding + 1, 
                    self._width_size - self._body_lr_padding * 2 - self._page_lr_padding * 2 - 1, 
                    "FFFFFF", "left", v["@value"])
                pos = pos + 1
            if len(data["ul"]["li"]) - 1 > key:
                pos = pos + 1

    return pos

def formatting(self, data):
    data = copy.deepcopy(data)

    if "title" not in data:
        data["title"] = ""

    if "ul" in data and isinstance(data["ul"], dict):
        if "li" in data["ul"] and isinstance(data["ul"]["li"], dict):
            data["ul"]["li"] = [data["ul"]["li"]]
        for key, val in enumerate(data["ul"]["li"]):
            if "link" in val and isinstance(val["link"], dict):
                data["ul"]["li"][key]["link"] = [data["ul"]["li"][key]["link"]]
            elif "link" not in val:
                data["ul"]["li"][key]["link"] = []
    else:
        data["ul"] = []

    data["title"] = {
        "value" : data["title"],
    }

    if "li" in data["ul"]:
        for key, val in enumerate(data["ul"]["li"]):
            data["ul"]["li"][key]["@title"] = {
                "value" : val["@title"],
            }
            for k, v in enumerate(val["link"]):
                data["ul"]["li"][key]["link"][k] = {
                    "@value" :{
                        "value" : v["@value"],
                    }
                }
    return data
