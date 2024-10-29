import copy
import importlib
from libs.xml2excel.utils import cell, util
for _m in [util, cell]:
    importlib.reload(_m)

def rendering(self, sheet, start_pos, data, hierarchy_number=[0], _init_start_col=-1, init_width=-1):
    pos = start_pos

    if len(hierarchy_number) == 1 and hierarchy_number[0] == 0:
        pos = cell.setInlineTitle(self, sheet, pos, data["title"])

    if isinstance(data, dict):
        _th_size = []
        if "thead" in data and isinstance(data["thead"], dict):
            if _init_start_col == -1:
                start_col = self._start_col + self._body_lr_padding + self._page_lr_padding + len(hierarchy_number) - 1
            else:
                start_col = _init_start_col
                _total_th_size = 0
                for key, val in enumerate(data["thead"]["th"]):
                    if not val["is_length_auto"]:
                        _total_th_size = _total_th_size + val["@length"]["value"]
                for key, val in enumerate(data["thead"]["th"]):
                    if val["is_length_auto"]:
                        val["@length"]["value"] = init_width - _total_th_size
                
            _max_br = 1
            for key, val in enumerate(data["thead"]["th"]):
                _tmp = val["#text"]["value"].count("\n") + 1
                if _max_br < _tmp:
                    _max_br = _tmp
                    
                _tmp = val["row_size"]
                if _max_br < _tmp:
                    _max_br = _tmp

            for key, val in enumerate(data["thead"]["th"]):
                _th_size.append(int(val["@length"]["value"]))

                if val["is_number"]:
                    cell.setCellwithBorder_Numnber(self, sheet, 
                        pos, _max_br, 
                        start_col, 
                        int(val["@length"]["value"]), 
                        self._background_color, "left", val["#text"],
                        False, "", val["number"], (0,0,0,255))
                else :
                    cell.setCellwithBorder(self, sheet, 
                        pos, _max_br, 
                        start_col, 
                        int(val["@length"]["value"]), 
                        self._background_color, "left", val["#text"])
                
                start_col = start_col + int(val["@length"]["value"])
            pos = pos + _max_br

        if "tbody" in data and isinstance(data["tbody"], dict) and "tr" in data["tbody"]:
            for key, val in enumerate(data["tbody"]["tr"]):
                if "td" in val:
                    if _init_start_col == -1:
                        start_col = self._start_col + self._body_lr_padding + self._page_lr_padding + len(hierarchy_number) - 1
                    else:
                        start_col = _init_start_col

                    _max_br = 1
                    for k, v in enumerate(val["td"]):
                        _tmp = v["value"].count("\n") + 1
                        if _max_br < _tmp:
                            _max_br = _tmp
                        _tmp = v["row_size"]
                        if _max_br < _tmp:
                            _max_br = _tmp

                    for k, v in enumerate(val["td"]):
                        if len(_th_size) > k:

                            _tr_width_size = _th_size[k]
                            _text_align = v["text_align"]
                            _bg_color = v["bg_color"]
                            _is_number = v["is_number"]
                            _number = v["number"]
                            _link = v["link"]

                            if v["is_row_merge"]:
                                _tr_width_size = sum(_th_size)

                            if _is_number:
                                cell.setCellwithBorder_Numnber(self, sheet, 
                                    pos, _max_br, 
                                    start_col, 
                                    _tr_width_size, 
                                    _bg_color, 
                                    _text_align, 
                                    v,
                                    False, _link, _number, (0,0,0,255))
                            else :
                                cell.setCellwithBorder(self, sheet, 
                                    pos, _max_br, 
                                    start_col, 
                                    _tr_width_size, 
                                    _bg_color, 
                                    _text_align, 
                                    v,
                                    False, _link)
                            start_col = start_col + int(_th_size[k])
                    pos = pos + _max_br

    return pos

def formatting(self, data, hierarchy_number=[0]):
    data = copy.deepcopy(data)

    _total_pos_height = 0

    # この変換こえぇぇ....
    if len(hierarchy_number) == 1 and hierarchy_number[0] == 0:
        _tmp = {
            "@type":"table"
        }
        if "title" not in data:
            _tmp["title"] = ""
                    
        if "title" in data:
            _tmp["title"] = {
                "value" : data["title"],
            }
        else:
            _tmp["title"] = {
                "value" : "",
            }

        if data["table"] is not None and "thead" in data["table"]:
            _tmp["thead"] = data["table"]["thead"]

        if data["table"] is not None and "tbody" in data["table"]:
            _tmp["tbody"] = data["table"]["tbody"]

        data = _tmp
        
    if isinstance(data, dict):

        if "@title" in data:
            data["@title"] = {
                "value" : data["@title"] 
            }
        else :
            data["@title"] = {
                "value" : "" 
            }

        # 合計のテーブルの高さを求めてみる
        if "thead" in data and isinstance(data["thead"], dict):
            if "th" in data["thead"] and isinstance(data["thead"]["th"], dict):
                data["thead"]["th"] = [data["thead"]["th"]]
            
            _table_size = self._width_size - self._body_lr_padding * 2 - self._page_lr_padding * 2 - len(hierarchy_number) + 1
            _total_th_size = 0
            for key, val in enumerate(data["thead"]["th"]):
                try:
                    if "@length" in val :
                        val["@length"] = int(val["@length"])
                    else :
                        val["@length"] = 0
                except ValueError:
                    val["@length"] = 0
                _total_th_size = _total_th_size + int(val["@length"])
            
            _max_height = 1

            for key, val in enumerate(data["thead"]["th"]):
                if val["@length"] > 0:
                    _tmp = int(val["@length"])
                    _is_length_auto = False
                else:
                    _tmp = _table_size - _total_th_size
                    _is_length_auto = True

                _text_array = []
                _is_number = False
                _row_size = 1
                
                if "#text" in val :
                    _tmp_text = val["#text"].split("\n")
                    for k, v in enumerate(_tmp_text):
                        _text_array.append(
                            v.lstrip(" ").rstrip(" ").lstrip("\t").rstrip("\t")
                        )

                if val is not None:
                    if isinstance(val, dict) and "@is_number" in val:
                        if val["@is_number"] == "True":
                            _is_number = True

                    if isinstance(val, dict) and "@row_size" in val:
                        if int(val["@row_size"]) > 0:
                            _row_size = int(val["@row_size"])
                        
                data["thead"]["th"][key] = {
                    "@length" : { 
                        "value" : _tmp 
                    },
                    "#text" : { 
                        "value" : "\n".join(_text_array),
                    },
                    "is_number" : _is_number,
                    "row_size" : _row_size,
                    "number" : self.view_section_number,
                    "is_length_auto" : _is_length_auto
                }

                if len(_text_array) > _max_height:
                    _max_height = len(_text_array)

                if _is_number:
                    self.view_section_number = self.view_section_number + 1
            
            _total_pos_height = _max_height

        elif "thead" in data :
            data["thead"] = []

        if "tbody" in data and isinstance(data["tbody"], dict):
            if "tr" in data["tbody"] and isinstance(data["tbody"]["tr"], dict):
                data["tbody"]["tr"] = [data["tbody"]["tr"]]
            elif "tr" not in data["tbody"] or data["tbody"]["tr"] is None:
                data["tbody"]["tr"] = []
            for key, val in enumerate(data["tbody"]["tr"]):
                if "td" in val and isinstance(val["td"], list) == False:
                    data["tbody"]["tr"][key]["td"] = [val["td"]]
                _max_height = 1
                for k, v in enumerate(data["tbody"]["tr"][key]["td"]):

                    _text_array = []
                    _tmp_is_row_merge = False
                    _text_align = "left"
                    _bg_color = "FFFFFF"
                    _link = ""
                    _sheet_link = ""
                    _is_number = False
                    _row_size = 1

                    if v is not None:
                        _tmp = v
                        if isinstance(v, dict) and "#text" in v:
                            _tmp = v["#text"]
                        
                        if isinstance(_tmp, dict):
                            _tmp = ""
                        
                        _tmp_text = _tmp.split("\n")
                        for text_k, text_v in enumerate(_tmp_text):
                            _text_array.append(
                                text_v.lstrip(" ").rstrip(" ").lstrip("\t").rstrip("\t")
                            )

                        if isinstance(v, dict) and "@is_row_merge" in v:
                            if v["@is_row_merge"] == "True":
                                _tmp_is_row_merge = True
                            else:
                                _tmp_is_row_merge = False

                        if isinstance(v, dict) and "@text_align" in v:
                            if v["@text_align"] == "center":
                                _text_align = "center"
                            else:
                                _text_align = "left"

                        if isinstance(v, dict) and "@bg_color" in v:
                            if v["@bg_color"] == "grey":
                                _bg_color = self._background_color

                        if isinstance(v, dict) and "@is_number" in v:
                            if v["@is_number"] == "True":
                                _is_number = True

                        if isinstance(v, dict) and "@row_size" in v:
                            if int(v["@row_size"]) > 0:
                                _row_size = int(v["@row_size"])

                        if isinstance(v, dict) and "@link" in v:
                            if v["@link"] != "":
                                _link = v["@link"]

                        if isinstance(v, dict) and "@sheet_link" in v:
                            if v["@sheet_link"] != "":
                                _link = "#" + v["@sheet_link"] + "!A1"

                    data["tbody"]["tr"][key]["td"][k] = {
                        "value" : "\n".join(_text_array),
                        "is_row_merge" : _tmp_is_row_merge,
                        "text_align" : _text_align,
                        "bg_color" : _bg_color,
                        "is_number" : _is_number,
                        "row_size" : _row_size,
                        "link" : _link,
                        "number" : self.view_section_number,
                    }

                    if len(_text_array) > _max_height:
                        _max_height = len(_text_array)

                    if _is_number:
                        self.view_section_number = self.view_section_number + 1

                _total_pos_height = _total_pos_height + _max_height

        elif "tbody" in data :
            data["tbody"] = []

    data["total_pos_height"] = _total_pos_height

    return data