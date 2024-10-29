import copy
import importlib
from libs.xml2excel.utils import cell, util
from libs.xml2excel.section import table
for _m in [util, cell, table]:
    importlib.reload(_m)

def rendering(self, sheet, start_pos, data):
    pos = cell.setInlineTitle(self, sheet, start_pos, data["title"])

    _base_col = self._start_col + self._body_lr_padding + self._page_lr_padding 
    _base_col_size = self._width_size - self._body_lr_padding * 2 - self._page_lr_padding * 2,

    return pos

def formatting(self, data, file):
    data = copy.deepcopy(data)
    if "title" not in data:
        data["title"] = ""
    data["title"] = {
        "value" : data["title"],
    }
    # print(json.dumps(data, indent=2, ensure_ascii=False))
    return data