import copy
import importlib
from libs.xml2excel.utils import cell, util
from libs.xml2excel.section import table
for _m in [util, cell, table]:
    importlib.reload(_m)

def rendering(self, sheet, start_pos, data):
    pos = cell.setInlineTitle(self, sheet, start_pos, data["title"])
    pos = circular_reference_rendering(self, sheet, pos, data["ul"], [1], 0)
    return pos

def circular_reference_rendering(self, sheet, start_pos, data, hierarchy_number, indent):
    pos = start_pos
    if isinstance(data, dict):
        for key, val in enumerate(data):
            if "@title" == val:
                cell.setCell(self, sheet, 
                    pos, 1, 
                    self._start_col + self._body_lr_padding + self._page_lr_padding + len(hierarchy_number) - 2, 
                    self._width_size - self._body_lr_padding * 2 - self._page_lr_padding * 2 - len(hierarchy_number) + 2, 
                    "FFFFFF", "left", data[val])
                pos = pos + 1
                continue
            if "@bottom_space" == val:
                pos = pos + int(data[val][0]["value"])
                continue
            for k, v in enumerate(data[val]):
                if "ul" in v and isinstance(v["ul"], dict):
                    pos = circular_reference_rendering(self, sheet, pos, v["ul"], hierarchy_number + [1], indent + 1)
                elif "table" in v and isinstance(v["table"], dict):
                    if "table_title" in v and v["table_title"] != "":
                        cell.setCell(self, sheet, 
                            pos, 1, 
                            self._start_col + self._body_lr_padding + self._page_lr_padding + len(hierarchy_number) - 1, 
                            self._width_size - self._body_lr_padding * 2 - self._page_lr_padding * 2 - len(hierarchy_number) + 1, 
                            "FFFFFF", "left", {"value":v["table_title"]})
                        pos = pos + 1
                    pos = table.rendering(self, sheet, pos, v["table"], hierarchy_number)
                else : 
                    cell.setCell(self, sheet, 
                        pos, 1, 
                        self._start_col + self._body_lr_padding + self._page_lr_padding + len(hierarchy_number) - 1, 
                        self._width_size - self._body_lr_padding * 2 - self._page_lr_padding * 2 - len(hierarchy_number) + 1, 
                        "FFFFFF", "left", v)
                    pos = pos + 1
                hierarchy_number[-1] = hierarchy_number[-1] + 1

    return pos

def formatting(self, data):
    data = copy.deepcopy(data)

    if "title" not in data:
        data["title"] = ""
    data["title"] = {
        "value" : data["title"],
    }

    if "@is_hierarchy_number" in data and data["@is_hierarchy_number"] == "True":
        data["@is_hierarchy_number"] = True
    else :
        data["@is_hierarchy_number"] = False
    if "ul" in data and isinstance(data["ul"], dict):
        data["ul"] = circular_reference_formatting(self, data["ul"], [1], data["@is_hierarchy_number"], 0)
    else:
        data["ul"] = {}

    return data

def circular_reference_formatting(self, data, hierarchy_number, is_hierarchy_number, indent):
    if isinstance(data, dict):

        # 並び順にアホみたいに依存している......こえぇ.....
        if is_hierarchy_number and "@title" not in data and (indent != 0 or len(hierarchy_number) != 1):
            _tmp = {}
            _tmp["@title"] = ""
            _tmp["li"] = data["li"]
            data = _tmp
        
        if "@bottom_space" in data :
            _tmp = {}
            if "@title" in data :
                _tmp["@title"] = data["@title"]
            _tmp["li"] = data["li"]
            _tmp["@bottom_space"] = data["@bottom_space"]
            data = _tmp

        for key, val in enumerate(data):
            if "@title" == val:
                if is_hierarchy_number:
                    _tmp = '.'.join(map(str,hierarchy_number[:-1])) + ". " + data[val]
                else:
                    _tmp = data[val]
                data[val] = {
                    "value" : _tmp
                }
                continue
            if "li" in val and isinstance(data[val], list) == False:
                if data[val] is not None:
                    data[val] = [data[val]]
                else :
                    data[val] = []
            data[val] = list(filter(None, data[val]))
            for k, v in enumerate(data[val]):
                if "ul" in v and isinstance(v["ul"], dict):
                    v["ul"] = circular_reference_formatting(self, v["ul"], hierarchy_number + [1], is_hierarchy_number, indent + 1)
                elif "table" in v and isinstance(v["table"], dict):
                    data[val][k]["table"] = table.formatting(self, v["table"], hierarchy_number)
                    if is_hierarchy_number:
                        _tmp = '.'.join(map(str,hierarchy_number)) + ". " + data[val][k]["table"]["@title"]["value"]
                    else :
                        _tmp = data[val][k]["table"]["@title"]["value"]
                    data[val][k]["table_title"] = _tmp
                elif "table" in v and v["table"] is None:
                    data[val][k] = {
                        "table" : {}
                    }
                else : 
                    if is_hierarchy_number:
                        _tmp = '.'.join(map(str,hierarchy_number)) + ". " + v
                    else :
                        _tmp = v

                    data[val][k] = {
                        "value" : _tmp
                    }
                hierarchy_number[-1] = hierarchy_number[-1] + 1

    return data