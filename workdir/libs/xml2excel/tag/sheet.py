import copy
import importlib
import json
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment # type: ignore
from libs.xml2excel.utils import util, cell, book
from libs.xml2excel.tag import header, sheet
from libs.xml2excel.section import desc, link, table, file, hierarchy, test, wireframe, flowchart, cover
for _m in [desc, link, util, cell, sheet, header, table, file, hierarchy, test, wireframe, flowchart, book, cover]:
    importlib.reload(_m)

def compare(self, data, past_data):
    data = copy.deepcopy(data)
    past_data = copy.deepcopy(past_data)

    past_data_compare_ids = {}
    # 過去データの参照
    if "body" in past_data["sheet"] and past_data["sheet"]["body"] is not None and "section" in past_data["sheet"]["body"]:
        for key, val in enumerate(past_data["sheet"]["body"]["section"]):
            if "@id" in val:
                past_data_compare_ids[val["@id"]] = key

    # print(json.dumps(past_data_compare_ids, indent=2, ensure_ascii=False))
    
    if "body" in data["sheet"] and data["sheet"]["body"] is not None and "section" in data["sheet"]["body"]:
        for key, val in enumerate(data["sheet"]["body"]["section"]):
            if "@id" in val and val["@id"] in past_data_compare_ids:
                _past_data = past_data["sheet"]["body"]["section"][past_data_compare_ids[val["@id"]]]
                if val["@type"] == "desc": # 階層無し_説明欄
                    data["sheet"]["body"]["section"][key] = desc.compare(self, val, _past_data)

    return data

def rendering(self, sheet, data, index):
    next_pos = header.rendering(self, sheet, data)
    next_pos = next_pos + 1 # 隙間1行
    _next_pos = next_pos
    self.diagram_data.append([])

    if "body" in data["sheet"] and data["sheet"]["body"] is not None and "section" in data["sheet"]["body"]:
        for _, val in enumerate(data["sheet"]["body"]["section"]):

            if val["@type"] == "desc": # 階層無し_説明欄
                _next_pos = desc.rendering(self, sheet, next_pos, val)

            if val["@type"] == "link": # 1階層_リンク集
                _next_pos = link.rendering(self, sheet, next_pos, val)

            if val["@type"] == "hierarchy": # 階層型_説明欄
                _next_pos = hierarchy.rendering(self, sheet, next_pos, val)

            if val["@type"] == "table": # 1カテゴリ_テーブル
                _next_pos = table.rendering(self, sheet, next_pos, val)

            if val["@type"] == "file": # 1カテゴリ_テーブル
                _next_pos = file.rendering(self, sheet, next_pos, val)

            if val["@type"] == "test": # 1カテゴリ_テーブル
                _next_pos = test.rendering(self, sheet, next_pos, val)

            if val["@type"] == "wireframe": # ワイヤーフレーム
                _next_pos = wireframe.rendering(self, sheet, next_pos, val)

            if val["@type"] == "flowchart": # フローチャート
                _next_pos = flowchart.rendering(self, sheet, next_pos, val)

            if val["@type"] == "cover": # カバー
                _next_pos = cover.rendering(self, sheet, next_pos, val)

            if _next_pos != next_pos:
                next_pos = _next_pos + 1 # 隙間1行
    
    sheet.cell(row=next_pos-1, column=self._start_col + self._width_size - 1).alignment = Alignment(horizontal='center', vertical='center')

def formatting(self, data, file_path):
    data = copy.deepcopy(data)

    # 基本的には配列構造に格納する
    if "body" in data["sheet"] and data["sheet"]["body"] is not None and "section" in data["sheet"]["body"]:
        if isinstance(data["sheet"]["body"]["section"], dict):
            data["sheet"]["body"]["section"] = [data["sheet"]["body"]["section"]]

    # データフォーマット
    if "header" in data["sheet"]:
        data["sheet"]["header"] = header.formatting(self, data["sheet"]["header"])

    if "body" in data["sheet"] and data["sheet"]["body"] is not None and "section" in data["sheet"]["body"]:
        for key, val in enumerate(data["sheet"]["body"]["section"]):

            self.view_section_number = 1

            if val["@type"] == "desc": # 階層無し_説明欄
                data["sheet"]["body"]["section"][key] = desc.formatting(self, val)

            if val["@type"] == "link": # 1階層_リンク集
                data["sheet"]["body"]["section"][key] = link.formatting(self, val)

            if val["@type"] == "hierarchy": # 階層型_説明欄
                data["sheet"]["body"]["section"][key] = hierarchy.formatting(self, val)

            if val["@type"] == "table": # 1カテゴリ_テーブル
                data["sheet"]["body"]["section"][key] = table.formatting(self, val)

            if val["@type"] == "file": # ファイル一覧
                data["sheet"]["body"]["section"][key] = file.formatting(self, val, file_path)

            if val["@type"] == "test": # テスト
                data["sheet"]["body"]["section"][key] = test.formatting(self, val, file_path)

            if val["@type"] == "wireframe": # ワイヤーフレーム
                data["sheet"]["body"]["section"][key] = wireframe.formatting(self, val, file_path)

            if val["@type"] == "flowchart": # フローチャート
                data["sheet"]["body"]["section"][key] = flowchart.formatting(self, val, file_path)

            if val["@type"] == "cover": # カバー
                data["sheet"]["body"]["section"][key] = cover.formatting(self, val)

        # print(json.dumps(data, indent=2, ensure_ascii=False))

    return data

def baseRendering(self, ws, data):

    #ws.sheet_format.baseColWidth = 1 
    ws.sheet_format.defaultColWidth = 1.3 # 1.3 こっちがな....マジでよくわからん
    ws.sheet_format.defaultRowHeight = 13 # 13 # 多分、「self._height_point_size」と一緒のはず、多分
    ws.sheet_view.showGridLines = False

    self._border_color = "666666"
    self._background_color = "F0F0F0"
    self._width_size = data["sheet"]["@width_size"]
    self._start_col = data["sheet"]["@start_col"]
    self._page_lr_padding = data["sheet"]["@page_lr_padding"]
    self._header_lr_padding = data["sheet"]["@header_lr_padding"]
    self._body_lr_padding = data["sheet"]["@body_lr_padding"]
    self._inline_title_length = data["sheet"]["@inline_title_length"]
    self._image_offset_calc_type = data["sheet"]["@image_offset_calc_type"]
    self._font_size = data["sheet"]["@font_size"]
    self._width_point_size = 8 # 8 横幅
    self._height_point_size = 13 # 13 縦幅

def baseFormatting(self, data):
    data = copy.deepcopy(data)

    if "@width_size" in data and int(data["@width_size"]) >= 34:
        data["@width_size"] = int(data["@width_size"])
    else:
        data["@width_size"] = 72

    if "@start_col" in data and int(data["@start_col"]) >= 1:
        data["@start_col"] = int(data["@start_col"])
    else:
        data["@start_col"] = 1

    if "@page_lr_padding" in data and int(data["@page_lr_padding"]) >= 0:
        data["@page_lr_padding"] = int(data["@page_lr_padding"])
    else:
        data["@page_lr_padding"] = 1

    if "@header_lr_padding" in data and int(data["@header_lr_padding"]) >= 0:
        data["@header_lr_padding"] = int(data["@header_lr_padding"])
    else:
        data["@header_lr_padding"] = 0

    if "@body_lr_padding" in data and int(data["@body_lr_padding"]) >= 0:
        data["@body_lr_padding"] = int(data["@body_lr_padding"])
    else:
        data["@body_lr_padding"] = 1

    if "@inline_title_length" in data and int(data["@inline_title_length"]) >= 1:
        data["@inline_title_length"] = int(data["@inline_title_length"])
    else:
        data["@inline_title_length"] = 10

    if "@font_size" in data and int(data["@font_size"]) >= 1:
        data["@font_size"] = int(data["@font_size"])
    else:
        data["@font_size"] = 10

    if "@image_offset_calc_type" in data and data["@image_offset_calc_type"] == "cell":
        data["@image_offset_calc_type"] = "cell"
    else:
        data["@image_offset_calc_type"] = "anchor"

    return data