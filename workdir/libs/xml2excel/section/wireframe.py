import copy
import json
import importlib
from libs.xml2excel.utils import cell
from libs.xml2excel.section import table
for _m in [table, cell]:
    importlib.reload(_m)

def rendering(self, sheet, start_pos, data):
    pos = cell.setInlineTitle(self, sheet, start_pos, data["title"])

    _base_width = data["@width_cell_size"]
    _padding = 1
    _start_key = ""

    if "row" in data:
        _start_key = "row"

    if "col" in data:
        _start_key = "col"

    if _start_key != "":
        pos = circular_reference_rendering(self, sheet, pos, 0, _start_key, data[_start_key], _base_width, _padding)

    return pos + 1

def circular_reference_rendering(self, sheet, pos, indent, key, val, width, padding):

    if key == "row" and isinstance(val, list):
        _max_pos = 0
        _base_start_pos = pos
        _pos_array = []
        _bottom_space_array = []
        for k, v in enumerate(val):
            _size = 1
            _bottom_space = 0
            if "@bottom_space" in v:
                _bottom_space = int(v["@bottom_space"])
            if "@size" in v:
                _size = int(v["@size"])
            _bottom_space_array.append(_bottom_space)
            _start_pos = pos
            if "col" in v:
                pos = circular_reference_rendering(self, sheet, _start_pos, indent, "col", v["col"], width, padding)
            if "table" in v:
                pos = pos + v["table"]["total_pos_height"] - 1
            pos = pos + 1
            if pos - _start_pos < _size:
                pos = _start_pos + _size
            _pos_array.append(pos - _start_pos)
            pos = pos + _bottom_space
            if pos > _max_pos:
                _max_pos = pos
        if _max_pos > 0:
            pos = _max_pos - 1
        _pos_start_step = _base_start_pos
        for k, v in enumerate(val):
            _number = ""
            _text = ""
            _is_border = not False
            _is_debug = False
            if "@number" in v and int(v["@number"]) > 0:
                _number = v["@number"]
            if "#text" in v and v["#text"] != "":
                _text = v["#text"]
            if "@is_border" in v and v["@is_border"] == "True":
                _is_border = not True
            if "@is_debug" in v and v["@is_debug"] == "True":
                _is_debug = True

            if "table" in v:
                table.rendering(self, sheet, _pos_start_step, v["table"], [0],
                                indent + self._start_col + self._body_lr_padding + self._page_lr_padding,
                                width)
            elif _text != "":

                _text_array = []
                _tmp_text = _text.split("\n")
                for tk, tv in enumerate(_tmp_text):
                    _text_array.append(
                        tv.lstrip(" ").rstrip(" ").lstrip("\t").rstrip("\t")
                    )
                _text = "\n".join(_text_array)

                cell.setCellwithBorder_Numnber(self, sheet, 
                    _pos_start_step,
                    _pos_array[k],
                    indent + self._start_col + self._body_lr_padding + self._page_lr_padding, 
                    width,
                    "FFFFFF", "center", {"value" : _text}, _is_border, "", _number, (0,0,0,255))
            elif not _is_border :
                cell.setBoxBorder(self, sheet, 
                    _pos_start_step,
                    _pos_array[k],
                    indent + self._start_col + self._body_lr_padding + self._page_lr_padding, 
                    width,
                    self._border_color, _number, (0,0,0,255))
            elif _is_debug :
                cell.setBoxBorder(self, sheet, 
                    _pos_start_step,
                    _pos_array[k],
                    indent + self._start_col + self._body_lr_padding + self._page_lr_padding, 
                    width,
                    "FF0000", _number, (0,0,0,255))
                
            _pos_start_step = _pos_start_step + _pos_array[k] + _bottom_space_array[k]
    elif key == "col" and isinstance(val, list):
        _start_pos = pos
        _max_pos = 0
        _total_size = 0
        _size_array = []
        _indent_array = []
        _tblr_padding_array = []
        _size = 0
        for k, v in enumerate(val):
            _tblr_padding = 0
            if "@padding" in v:
                _tblr_padding = int(v["@padding"])
            _tblr_padding_array.append(_tblr_padding)
        
            _indent_array.append(_total_size)
            if "@size" in v:
                _size = int(v["@size"])
            else:
                _size = 1
            _total_size = _total_size + _size
            _size_array.append(_size)
        _col_base_width = round(width / _total_size)

        for k, v in enumerate(val):
            if "row" in v:
                pos = circular_reference_rendering(self, sheet, 
                    _start_pos + _tblr_padding_array[k], 
                    indent + _col_base_width * _indent_array[k] + _tblr_padding_array[k], 
                    "row", v["row"], 
                    _col_base_width * _size_array[k] - _tblr_padding_array[k] * 2, 
                    padding) + _tblr_padding_array[k] * 1
                if pos > _max_pos:
                    _max_pos = pos
            if "table" in v:
                pos = pos + v["table"]["total_pos_height"] + _tblr_padding_array[k] * 2 - 1
                if pos > _max_pos:
                    _max_pos = pos
        if _max_pos > 0:
            pos = _max_pos

        if pos == _start_pos and len(_tblr_padding_array) > 0 and _tblr_padding_array[0] > 0:
            pos = _start_pos + _tblr_padding_array[0] * 2

        for k, v in enumerate(val):
            _number = ""
            _text = ""
            _is_border = not False
            _is_debug = False
            if "@number" in v and int(v["@number"]) > 0:
                _number = v["@number"]
            if "#text" in v and v["#text"] != "":
                _text = v["#text"]
            if "@is_border" in v and v["@is_border"] == "True":
                _is_border = not True
            if "@is_debug" in v and v["@is_debug"] == "True":
                _is_debug = True

            if "table" in v:
                table.rendering(self, sheet, _start_pos + _tblr_padding_array[k], v["table"], [0],
                    indent + _col_base_width * _indent_array[k] + self._start_col + self._body_lr_padding + self._page_lr_padding + _tblr_padding_array[k], 
                    _col_base_width * _size_array[k] - _tblr_padding_array[k] * 2)
            elif _text != "":

                _text_array = []
                _tmp_text = _text.split("\n")
                for tk, tv in enumerate(_tmp_text):
                    _text_array.append(
                        tv.lstrip(" ").rstrip(" ").lstrip("\t").rstrip("\t")
                    )
                _text = "\n".join(_text_array)

                cell.setCellwithBorder_Numnber(self, sheet, 
                    _start_pos + _tblr_padding_array[k], 
                    pos - (_start_pos + _tblr_padding_array[k] * 2) + 1,
                    indent + _col_base_width * _indent_array[k] + self._start_col + self._body_lr_padding + self._page_lr_padding + _tblr_padding_array[k], 
                    _col_base_width * _size_array[k] - _tblr_padding_array[k] * 2,
                    "FFFFFF", "center", {"value" : _text}, _is_border, "", _number, (0,0,0,255))
            elif not _is_border :
                cell.setBoxBorder(self, sheet, 
                    _start_pos + _tblr_padding_array[k], 
                    pos - (_start_pos + _tblr_padding_array[k] * 2) + 1,
                    indent + _col_base_width * _indent_array[k] + self._start_col + self._body_lr_padding + self._page_lr_padding + _tblr_padding_array[k], 
                    _col_base_width * _size_array[k] - _tblr_padding_array[k] * 2,
                    self._border_color, _number, (0,0,0,255))
            elif _is_debug :
                cell.setBoxBorder(self, sheet, 
                    _start_pos + _tblr_padding_array[k], 
                    pos - (_start_pos + _tblr_padding_array[k] * 2) + 1,
                    indent + _col_base_width * _indent_array[k] + self._start_col + self._body_lr_padding + self._page_lr_padding + _tblr_padding_array[k], 
                    _col_base_width * _size_array[k] - _tblr_padding_array[k] * 2,
                    "0000FF", _number, (0,0,0,255))

    return pos

def formatting(self, data, file_path):
    data = copy.deepcopy(data)
    if "title" not in data:
        data["title"] = ""
    data["title"] = {
        "value" : data["title"],
    }

    if "@width_cell_size" not in data:
        data["@width_cell_size"] = self._width_size - self._body_lr_padding * 2 - self._page_lr_padding * 2
    else:
        data["@width_cell_size"] = int(data["@width_cell_size"])

    _start_key = ""

    if "row" in data:
        _start_key = "row"

    if "col" in data:
        _start_key = "col"

    if _start_key != "":
        data[_start_key] = circular_reference_formatting(self, _start_key, data[_start_key])

    #print(json.dumps(data, indent=2, ensure_ascii=False))
    return data

def circular_reference_formatting(self, key, val):
    val = copy.deepcopy(val)

    if key == "row" or key == "col":
        if val is None:
            val = [{}]
        elif isinstance(val, str):
            val = [{
                "#text": val
            }]
        elif isinstance(val, list):
            for k, v, in enumerate(val):
                val[k] = circular_reference_formatting(self, k, v)
        elif isinstance(val, dict):
            _tmp = {}
            if "row" in val or "col" in val:
                for k in val:
                    if k == "row" or k == "col":
                        _tmp[k] = circular_reference_formatting(self, k, val[k])
                    else:
                        _tmp[k] = val[k]
                val = [_tmp]
            elif "table" in val:
                val["table"] = table.formatting(self, val)
                val = [val]
            else :
                val = [val]
    else:
        if val is None:
            val = {}
        elif isinstance(val, str):
            val = {"#text" : val}
        elif isinstance(val, dict):
            if "row" in val:
                val["row"] = circular_reference_formatting(self, "row", val["row"])
            if "col" in val:
                val["col"] = circular_reference_formatting(self, "col", val["col"])
            if "table" in val:
                val["table"] = table.formatting(self, val)

    return val