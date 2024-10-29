import copy
import importlib
import os
import json
from libs.xml2excel.utils import cell, book
from lark import Lark # type: ignore
for _m in [cell, book]:
    importlib.reload(_m)

def preset(self):

    for _key, _val in enumerate(self.diagram_data):
        # idの再設定
        _step_index = 0
        for _k, _v in enumerate(_val):
            self.diagram_data[_key][_k]["id"] = self.diagram_number_array[_key] + _step_index
            _step_index = _step_index + 1
        self.diagram_number_array[_key] = _step_index

    ##############################################################################
    # アローdrawingを設定
    # テスト例）1シート目の2,3番目のダイアグラムのアローをテスト
    # シーケンス実装とか出てきた時にはどうするかな.....頭の片隅に

    _diagram_data = copy.deepcopy(self.diagram_data)

    for _tmp_sheet_index, __tmpval in enumerate(_diagram_data):
        for _start_diagram_key, _v in enumerate(__tmpval):

            if _v["is_last"]:
                continue

            if len(__tmpval) - 2 < _start_diagram_key:
                break
            #print(_tmp_sheet_index, _start_diagram_key, _v["id"])
            
            # デバッグする時は、左に二つインデントして、下記のコメント二つ外せばいける
            #_tmp_sheet_index = 0
            #_start_diagram_key = 5 # デバッグ中
            _end_diagram_key = _start_diagram_key + 1

            # task_id設定
            _task_ids_array = {}
            for _key, _val in enumerate(self.diagram_data[_tmp_sheet_index]):
                if "task_id" in _val and _val["task_id"] != "":
                    _task_ids_array[_val["task_id"]] = _key

            _pos = self.diagram_data[_tmp_sheet_index][_start_diagram_key]["pos"] + self.diagram_data[_tmp_sheet_index][_start_diagram_key]["pos_len"]
            _pos_len = self.diagram_data[_tmp_sheet_index][_end_diagram_key]["pos"] - _pos
            _row = self.diagram_data[_tmp_sheet_index][_start_diagram_key]["row"] + round(self.diagram_data[_tmp_sheet_index][_start_diagram_key]["row_len"] / 2)
            _row_len = self.diagram_data[_tmp_sheet_index][_end_diagram_key]["row"] + round(self.diagram_data[_tmp_sheet_index][_end_diagram_key]["row_len"] / 2)
        
            _start_id = self.diagram_data[_tmp_sheet_index][_start_diagram_key]["id"]
            _end_id = self.diagram_data[_tmp_sheet_index][_end_diagram_key]["id"]

            _block_type = self.diagram_data[_tmp_sheet_index][_start_diagram_key]["type"]

            _tmp_row_len = _row - _row_len

            # 左段落ちケース
            if _row_len < _row:
                _type = "connect_arrow_left_bottom"
                _tmp_row = _row_len
                _tmp_row_len = _row - _row_len
                _start_idx = 2
                _end_idx = 0

            # 右段落ちケース
            # このケースの場合に引き落としたケースへの線を作る必要がある
            elif _row_len > _row and _block_type == "hishigata":

                if "false_to" in self.diagram_data[_tmp_sheet_index][_start_diagram_key] and self.diagram_data[_tmp_sheet_index][_start_diagram_key]["false_to"] != "":

                    _type = "connect_arrow_bottom_down"
                    _pos = self.diagram_data[_tmp_sheet_index][_start_diagram_key]["pos"] + self.diagram_data[_tmp_sheet_index][_start_diagram_key]["pos_len"]
                    _pos_len = self.diagram_data[_tmp_sheet_index][_end_diagram_key]["pos"] - _pos
                    _row = self.diagram_data[_tmp_sheet_index][_start_diagram_key]["row"] + round(self.diagram_data[_tmp_sheet_index][_start_diagram_key]["row_len"] / 2)
                    _tmp_row = _row
                    _tmp_row_len = _row_len - _row
                    _start_idx = 2
                    _end_idx = 0

                else :

                    _type = "connect_arrow_right_bottom"
                    _pos = self.diagram_data[_tmp_sheet_index][_start_diagram_key]["pos"] + round(self.diagram_data[_tmp_sheet_index][_start_diagram_key]["pos_len"] / 2)
                    _pos_len = self.diagram_data[_tmp_sheet_index][_end_diagram_key]["pos"] - _pos
                    _row = self.diagram_data[_tmp_sheet_index][_start_diagram_key]["row"] + self.diagram_data[_tmp_sheet_index][_start_diagram_key]["row_len"]
                    _tmp_row = _row
                    _tmp_row_len = _row_len - _row
                    _start_idx = 3
                    _end_idx = 0

            elif _row_len > _row and _block_type == "textbox":
                _type = "connect_arrow_bottom_down"
                _pos = self.diagram_data[_tmp_sheet_index][_start_diagram_key]["pos"] + self.diagram_data[_tmp_sheet_index][_start_diagram_key]["pos_len"]
                _pos_len = self.diagram_data[_tmp_sheet_index][_end_diagram_key]["pos"] - _pos
                _row = self.diagram_data[_tmp_sheet_index][_start_diagram_key]["row"] + round(self.diagram_data[_tmp_sheet_index][_start_diagram_key]["row_len"] / 2)
                _tmp_row = _row
                _tmp_row_len = _row_len - _row
                _start_idx = 2
                _end_idx = 0

            # そのまま下下ろしケース
            else:
                _type = "connect_arrow_bottom"
                _tmp_row = _row
                _start_idx = 2
                _end_idx = 0
        
            self.diagram_data[_tmp_sheet_index].append({
                "pos"       : _pos,         # 縦 - 開始位置
                "pos_len"   : _pos_len,     # 縦 - 相対値
                "row"       : _tmp_row,     # 横 - 開始位置
                "row_len"   : _tmp_row_len, # 横 - 相対値
                "text"      : {"value": ""},
                "id"        : self.diagram_number_array[_tmp_sheet_index],
                "type"      : _type,
                "start_id"  : _start_id,
                "end_id"    : _end_id,
                "start_idx" : _start_idx,
                "end_idx"   : _end_idx,
            })
            self.diagram_number_array[_tmp_sheet_index] = self.diagram_number_array[_tmp_sheet_index] + 1

            # 菱形の時のテキスト
            if _block_type == "hishigata" :
                _true_text = self.diagram_data[_tmp_sheet_index][_start_diagram_key]["true_text"]
                _false_text = self.diagram_data[_tmp_sheet_index][_start_diagram_key]["false_text"]

                if _true_text != "":
                    _pos = self.diagram_data[_tmp_sheet_index][_start_diagram_key]["pos"] + round(self.diagram_data[_tmp_sheet_index][_start_diagram_key]["pos_len"] / 2) - 1
                    _pos_len = 1
                    _tmp_row = self.diagram_data[_tmp_sheet_index][_start_diagram_key]["row"] + self.diagram_data[_tmp_sheet_index][_start_diagram_key]["row_len"] - 1
                    _tmp_row_len = len(_true_text) + 4
            
                    self.diagram_data[_tmp_sheet_index].append({
                        "pos"       : _pos,         # 縦 - 開始位置
                        "pos_len"   : _pos_len,     # 縦 - 相対値
                        "row"       : _tmp_row,     # 横 - 開始位置
                        "row_len"   : _tmp_row_len, # 横 - 相対値
                        "text"      : {"value": _true_text},
                        "id"        : self.diagram_number_array[_tmp_sheet_index],
                        "type"      : "clear_textbox_bottomtext"
                    })
                    self.diagram_number_array[_tmp_sheet_index] = self.diagram_number_array[_tmp_sheet_index] + 1

                if _false_text != "":
                    _pos = self.diagram_data[_tmp_sheet_index][_start_diagram_key]["pos"] + self.diagram_data[_tmp_sheet_index][_start_diagram_key]["pos_len"] 
                    _pos_len = 1
                    _tmp_row = self.diagram_data[_tmp_sheet_index][_start_diagram_key]["row"] + round(self.diagram_data[_tmp_sheet_index][_start_diagram_key]["row_len"] / 2)
                    _tmp_row_len = len(_false_text) + 4
            
                    self.diagram_data[_tmp_sheet_index].append({
                        "pos"       : _pos,         # 縦 - 開始位置
                        "pos_len"   : _pos_len,     # 縦 - 相対値
                        "row"       : _tmp_row,     # 横 - 開始位置
                        "row_len"   : _tmp_row_len, # 横 - 相対値
                        "text"      : {"value": _false_text},
                        "id"        : self.diagram_number_array[_tmp_sheet_index],
                        "type"      : "clear_textbox_bottomtext"
                    })
                    self.diagram_number_array[_tmp_sheet_index] = self.diagram_number_array[_tmp_sheet_index] + 1

            # false_toがある場合
            if _block_type == "hishigata" and "false_to" in self.diagram_data[_tmp_sheet_index][_start_diagram_key] and self.diagram_data[_tmp_sheet_index][_start_diagram_key]["false_to"] == "":

                _row = self.diagram_data[_tmp_sheet_index][_start_diagram_key]["row"]
                _is_start_status = False
                _next_diagram_key = -1
                for _t_index, _t_value in enumerate(_diagram_data[_tmp_sheet_index]):
                    if _t_index == _start_diagram_key :
                        _is_start_status = True
                    _row_comp = self.diagram_data[_tmp_sheet_index][_t_index]["row"]
                    if _is_start_status and _row == _row_comp and _t_index > _start_diagram_key and _next_diagram_key == -1:
                        _next_diagram_key = _t_index

                # 真下に下ろす
                if _next_diagram_key != -1:
                    _end_diagram_key = _next_diagram_key

                    _pos = self.diagram_data[_tmp_sheet_index][_start_diagram_key]["pos"] + self.diagram_data[_tmp_sheet_index][_start_diagram_key]["pos_len"]
                    _pos_len = self.diagram_data[_tmp_sheet_index][_end_diagram_key]["pos"] - _pos
                    _row = self.diagram_data[_tmp_sheet_index][_start_diagram_key]["row"] + round(self.diagram_data[_tmp_sheet_index][_start_diagram_key]["row_len"] / 2)
                    _row_len = self.diagram_data[_tmp_sheet_index][_end_diagram_key]["row"] + round(self.diagram_data[_tmp_sheet_index][_end_diagram_key]["row_len"] / 2)
                    _start_id = self.diagram_data[_tmp_sheet_index][_start_diagram_key]["id"]
                    _end_id = self.diagram_data[_tmp_sheet_index][_end_diagram_key]["id"]
                    _block_type = self.diagram_data[_tmp_sheet_index][_start_diagram_key]["type"]
                    _tmp_row_len = _row - _row_len

                    _type = "connect_arrow_bottom"
                    _tmp_row = _row
                    _start_idx = 2
                    _end_idx = 0
        
                    self.diagram_data[_tmp_sheet_index].append({
                        "pos"       : _pos,         # 縦 - 開始位置
                        "pos_len"   : _pos_len,     # 縦 - 相対値
                        "row"       : _tmp_row,     # 横 - 開始位置
                        "row_len"   : _tmp_row_len, # 横 - 相対値
                        "text"      : {"value": ""},
                        "id"        : self.diagram_number_array[_tmp_sheet_index],
                        "type"      : _type,
                        "start_id"  : _start_id,
                        "end_id"    : _end_id,
                        "start_idx" : _start_idx,
                        "end_idx"   : _end_idx,
                    })
                    self.diagram_number_array[_tmp_sheet_index] = self.diagram_number_array[_tmp_sheet_index] + 1


            # false_toがある場合
            if "false_to" in self.diagram_data[_tmp_sheet_index][_start_diagram_key] and self.diagram_data[_tmp_sheet_index][_start_diagram_key]["false_to"] != "":
                _false_to = self.diagram_data[_tmp_sheet_index][_start_diagram_key]["false_to"]

                _end_diagram_key = _start_diagram_key
                _start_diagram_key = _task_ids_array[_false_to]

                # 最大横幅探索
                # [TODO] 上に戻る際に、その途中の上に戻る数分右にずらさないといけないはず
                # ここのロジック、スルッと期待値に行ってしまって自信ない
                _tmp_max_width_cell = 0
                for _tmp_index in range(_start_diagram_key, _end_diagram_key + 1):
                    _tmp_max_width_cell_value = self.diagram_data[_tmp_sheet_index][_tmp_index]["max_width_cell"]
                    if _tmp_max_width_cell_value > _tmp_max_width_cell:
                        _tmp_max_width_cell = _tmp_max_width_cell_value

                #print("=======================================")
                #print("_tmp_max_width_cell %d" % _tmp_max_width_cell)
                _tmp_stack_count = 0
                for _tmp_index in range(_start_diagram_key, _end_diagram_key + 1):
                    _inner_tmp_max_width_cell = 0
                    _tmp_false_to = self.diagram_data[_tmp_sheet_index][_tmp_index]["false_to"]
                    _tmp2_max_width_cell = self.diagram_data[_tmp_sheet_index][_tmp_index]["max_width_cell"]

                    _inner_tmp_max_width_cell = _tmp2_max_width_cell

                    # inner _tmp2_max_width_cell
                    if _tmp_false_to != "":
                        _inner_end_diagram_key = _tmp_index
                        _inner_start_diagram_key = _task_ids_array[_tmp_false_to]
                        for _inner_tmp_index in range(_inner_start_diagram_key, _inner_end_diagram_key + 1):
                            _inner_tmp_max_width_cell_value = self.diagram_data[_tmp_sheet_index][_inner_tmp_index]["max_width_cell"]
                            if _inner_tmp_max_width_cell_value > _inner_tmp_max_width_cell:
                                _inner_tmp_max_width_cell = _inner_tmp_max_width_cell_value

                    if _tmp_false_to != "" and _inner_tmp_max_width_cell == _tmp_max_width_cell:
                        _tmp_stack_count = _tmp_stack_count + 1

                    #print("false_to(%d -> %d): %s" % (_tmp2_max_width_cell, _inner_tmp_max_width_cell, _tmp_false_to))
                #print("_tmp_stack_count %d" % _tmp_stack_count)

                _tmp_max_width_cell = _tmp_max_width_cell + _tmp_stack_count - 1
                

                _pos = self.diagram_data[_tmp_sheet_index][_start_diagram_key]["pos"]
                _pos_len = self.diagram_data[_tmp_sheet_index][_end_diagram_key]["pos"] + round(self.diagram_data[_tmp_sheet_index][_end_diagram_key]["pos_len"] / 2) - _pos

                _row = self.diagram_data[_tmp_sheet_index][_start_diagram_key]["row"] + round(self.diagram_data[_tmp_sheet_index][_start_diagram_key]["row_len"] / 2)
                _row_len = self.diagram_data[_tmp_sheet_index][_end_diagram_key]["row"] + self.diagram_data[_tmp_sheet_index][_end_diagram_key]["row_len"] - _row
            
                _start_id = self.diagram_data[_tmp_sheet_index][_start_diagram_key]["id"]
                _end_id = self.diagram_data[_tmp_sheet_index][_end_diagram_key]["id"]

                # https://stackoverflow.com/questions/1829009/in-powerpoint-2007-how-can-i-position-a-callouts-tail-programatically
                # DistanceX = Coordinate.X - (Callout.X + (Callout.X_Ext/2))
                # DistanceY = Coordinate.Y - (Callout.Y + (Callout.Y_Ext/2))
                # TailX = (DistanceX/Callout.X_Ext) * 100000
                # TailY = (DistanceY/Callout.Y_Ext) * 100000

                _max_width_cell = _tmp_max_width_cell - (self.diagram_data[_tmp_sheet_index][_end_diagram_key]["row"] + self.diagram_data[_tmp_sheet_index][_end_diagram_key]["row_len"])
                _max_width_cell_EMU = -1 * round((_max_width_cell / _row_len) * 100000)
                _max_height_cell_EMU = round(((_pos_len + 1) / _pos_len) * 100000)

                # Todo: 真上か、左上への座標だと有効な感じのテンプレート
                _type = "connect_arrow_right_up"
                _start_idx = 3
                _end_idx = 0
                _tmp_row = _row
                _tmp_row_len = _row_len

                if _tmp_row_len < 0:
                    _type = "connect_arrow_right_up_over"
                    _row = self.diagram_data[_tmp_sheet_index][_end_diagram_key]["row"] + self.diagram_data[_tmp_sheet_index][_end_diagram_key]["row_len"]
                    _row_len = (self.diagram_data[_tmp_sheet_index][_start_diagram_key]["row"] + round(self.diagram_data[_tmp_sheet_index][_start_diagram_key]["row_len"] / 2)) - _row
                    _max_width_cell_EMU = round((_max_width_cell / _row_len) * 100000)

                    _tmp_row = _row
                    _tmp_row_len = _row_len
        
                self.diagram_data[_tmp_sheet_index].append({
                    "pos"       : _pos,         # 縦 - 開始位置
                    "pos_len"   : _pos_len,     # 縦 - 相対値
                    "row"       : _tmp_row,     # 横 - 開始位置
                    "row_len"   : _tmp_row_len, # 横 - 相対値
                    "text"      : {"value": ""},
                    "id"        : self.diagram_number_array[_tmp_sheet_index],
                    "type"      : _type,
                    "start_id"  : _end_id,
                    "end_id"    : _start_id,
                    "start_idx" : _start_idx,  # 上:0、左:1、下:2、右:3　・・・多分
                    "end_idx"   : _end_idx,
                    "max_width_cell_EMU"  : _max_width_cell_EMU,
                    "max_height_cell_EMU" : _max_height_cell_EMU,
                })
                self.diagram_number_array[_tmp_sheet_index] = self.diagram_number_array[_tmp_sheet_index] + 1

def rendering(self, sheet, start_pos, data):
    pos = cell.setInlineTitle(self, sheet, start_pos, data["title"])

    _start_indent = self._start_col + self._body_lr_padding + self._page_lr_padding
    _thread_lr_padding = 1
    _base_width_cell = 20
    _stack_indent_size = 0
    _thread_height = 3
    _task_height = 2
    _task_bottom_space = 2
    _max_ident = data["max_ident"]
    _indent_array = {}

    if "@base_width_cell" in data["flowchart"] and int(data["flowchart"]["@base_width_cell"]) > 0:
        _base_width_cell = int(data["flowchart"]["@base_width_cell"])

    _max_thread_br = 1
    for key, val in enumerate(data["flowchart"]["participants"]["thread"]):
        if len(val["#text"].split("\n")) > _max_thread_br:
            _max_thread_br = len(val["#text"].split("\n"))

    for key, val in enumerate(data["flowchart"]["participants"]["thread"]):
        _tmp_indent = _start_indent + _thread_lr_padding * key * 2 + _stack_indent_size * _base_width_cell
        _tmp_width = _base_width_cell * _max_ident[val["@id"]] + _thread_lr_padding * 2
        _stack_indent_size = _stack_indent_size + _max_ident[val["@id"]]

        _tmp_array = []
        for v in range(1,_max_ident[val["@id"]] + 1):
            _tmp_array.append(
                _tmp_indent + (v-1) * _base_width_cell + _thread_lr_padding
            )
        _indent_array[val["@id"]] = _tmp_array

        _text = ""
        if "#text" in val:
            _text = val["#text"]

        cell.setCellwithBorder(self, sheet, 
            pos, 
            _thread_height, 
            _tmp_indent, 
            _tmp_width,
            "FFFFFF", "center", {"value": _text})
    
    pos = pos + _thread_height + 1

    def circular_reference_rendering(self, start_pos, _data, _max_width_cell=0):
        _pos = start_pos

        for key, val in enumerate(_data):

            _text = ""
            _tmp_height = _task_height
            if "#text" in val:
                _text = val["#text"]
                if len(_text.split("\n")) > _tmp_height:
                    _tmp_height = len(_text.split("\n")) + 1
            
            if "@if" in val:
                _text = val["@if"]
                _tmp_height = _tmp_height + 2

            #cell.setCellwithBorder(self, sheet, 
            #    _pos, 
            #    _tmp_height, 
            #    _indent_array[val["@by"]][val["indent"][val["@by"]] - 1], 
            #    _base_width_cell,
            #    "FFFFFF", "center", {"value": _text})
            
            _type = "textbox"
            if "@if" in val:
                _type = "hishigata"
            
            _task_id = ""
            if "@id" in val:
                _task_id = val["@id"]

            if _task_id != "" and _pos == start_pos:
                _pos = start_pos + 1
            
            _true_text = ""
            if "@true" in val:
                _true_text = val["@true"]
            
            _false_text = ""
            if "@false" in val:
                _false_text = val["@false"]
            
            _false_to = ""
            if "@false_to" in val:
                _false_to = val["@false_to"]

            # Todo: _tmp_height 偶数倍制御しておく
            # Todo: ifの内包数分増やす必要がある？
            # Todo: false_toって上戻りで考えてたけど、下もあるのか？あ、いや下は通常分岐か？
            #if (_indent_array[val["@by"]][val["indent"][val["@by"]] - 1] - 1) + _base_width_cell + 1 > _max_width_cell:
            _max_width_cell = (_indent_array[val["@by"]][val["indent"][val["@by"]] - 1] - 1) + _base_width_cell + 2

            self.diagram_data[self.sheet_index].append({
                "pos"            : _pos - 1,
                "pos_len"        : _tmp_height,
                "row"            : _indent_array[val["@by"]][val["indent"][val["@by"]] - 1] - 1,
                "row_len"        : _base_width_cell,
                "text"           : {"value": _text},
                "id"             : 0, #self.diagram_number, # この値はシートのセクションが回り切った後に再計算させる必要がある
                "type"           : _type,
                "task_id"        : _task_id,
                "true_text"      : _true_text,
                "false_text"     : _false_text,
                "false_to"       : _false_to,
                "max_width_cell" : _max_width_cell,
                "is_last"        : False
            })
            
            if "task" in val:
                _pos = _pos + _tmp_height + _task_bottom_space
                _pos = circular_reference_rendering(self, _pos, val["task"], _max_width_cell)
            else :
                _pos = _pos + _tmp_height + _task_bottom_space

        return _pos

    pos = circular_reference_rendering(self, pos, data["flowchart"]["flow"]["task"], 0)

    # 区切りのデータにフラグ立てる
    self.diagram_data[self.sheet_index][-1]["is_last"] = True

    pos = pos - 1
    
    _stack_indent_size = 0
    for key, val in enumerate(data["flowchart"]["participants"]["thread"]):
        _tmp_indent = _start_indent + _thread_lr_padding * key * 2 + _stack_indent_size * _base_width_cell
        _tmp_width = _base_width_cell * _max_ident[val["@id"]] + _thread_lr_padding * 2
        _stack_indent_size = _stack_indent_size + _max_ident[val["@id"]]

        cell.setBoxBorder(self, sheet, 
            start_pos + 2 + _thread_height, 
            pos - start_pos - 2 - _thread_height, 
            _tmp_indent, 
            _tmp_width)
            
    return pos


def formatting(self, data, file_path):
    data = copy.deepcopy(data)
    base_indents = {}
    max_ident = {}

    if "title" not in data:
        data["title"] = ""
    data["title"] = {
        "value" : data["title"],
    }

    def circular_reference_formatting(self, data, file_path, indents):
        data = copy.deepcopy(data)
        if isinstance(data["task"], dict):
            data["task"] = [data["task"]]

        for key, val in enumerate(data["task"]):
            if "task" in val:
                _tmp_indents = copy.deepcopy(indents)
                data["task"][key]["indent"] = _tmp_indents

                _tmp_indents = copy.deepcopy(indents)
                _tmp_indents[val["@by"]] = _tmp_indents[val["@by"]] + 1
                max_ident[val["@by"]] = _tmp_indents[val["@by"]]
                data["task"][key] = circular_reference_formatting(self, data["task"][key], file_path, _tmp_indents)
            else:
                _tmp_indents = copy.deepcopy(indents)
                data["task"][key]["indent"] = _tmp_indents

            if "#text" in val:
                _text_array = []
                _tmp_text = val["#text"].split("\n")
                for tk, tv in enumerate(_tmp_text):
                    _text_array.append(
                        tv.lstrip(" ").rstrip(" ").lstrip("\t").rstrip("\t")
                    )
                _text = "\n".join(_text_array)
                data["task"][key]["#text"] = _text

        return data

    if data is not None and "flowchart" in data:
            
        if "thread" in data["flowchart"]["participants"] and isinstance(data["flowchart"]["participants"]["thread"], dict):
            data["flowchart"]["participants"]["thread"] = [data["flowchart"]["participants"]["thread"]]

        for key, val in enumerate(data["flowchart"]["participants"]["thread"]):
            base_indents[val["@id"]] = 1
            max_ident[val["@id"]] = 1

            _text_array = []
            if "#text" in val:
                _tmp_text = val["#text"].split("\n")
            else :
                _tmp_text = ""
            for tk, tv in enumerate(_tmp_text):
                _text_array.append(
                    tv.lstrip(" ").rstrip(" ").lstrip("\t").rstrip("\t")
                )
            _text = "\n".join(_text_array)
            data["flowchart"]["participants"]["thread"][key]["#text"] = _text

        if "task" in data["flowchart"]["flow"] and isinstance(data["flowchart"]["flow"]["task"], dict):
            data["flowchart"]["flow"]["task"]["indent"] = base_indents[data["flowchart"]["flow"]["task"]["by"]]
            data["flowchart"]["flow"]["task"] = [data["flowchart"]["flow"]["task"]]

        for key, val in enumerate(data["flowchart"]["flow"]["task"]):
            if "task" in val:
                _tmp_indents = copy.deepcopy(base_indents)
                data["flowchart"]["flow"]["task"][key]["indent"] = _tmp_indents

                _tmp_indents = copy.deepcopy(base_indents)
                _tmp_indents[val["@by"]] = _tmp_indents[val["@by"]] + 1
                max_ident[val["@by"]] = _tmp_indents[val["@by"]]
                data["flowchart"]["flow"]["task"][key] = circular_reference_formatting(self, data["flowchart"]["flow"]["task"][key], file_path, _tmp_indents)
            else :
                _tmp_indents = copy.deepcopy(base_indents)
                data["flowchart"]["flow"]["task"][key]["indent"] = _tmp_indents

            if "#text" in val:
                _text_array = []
                _tmp_text = val["#text"].split("\n")
                for tk, tv in enumerate(_tmp_text):
                    _text_array.append(
                        tv.lstrip(" ").rstrip(" ").lstrip("\t").rstrip("\t")
                    )
                _text = "\n".join(_text_array)
                data["flowchart"]["flow"]["task"][key]["#text"] = _text

    data["max_ident"] = max_ident

    return data
