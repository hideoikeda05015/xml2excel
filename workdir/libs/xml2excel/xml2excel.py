import os
import openpyxl # type: ignore
import shutil
import importlib
import json
import copy
from datetime import datetime, timedelta, timezone
from openpyxl.utils.units import pixels_to_EMU, points_to_pixels # type: ignore
from tqdm.notebook import tqdm # type: ignore
from openpyxl.drawing.xdr import XDRPositiveSize2D # type: ignore
from openpyxl.drawing.spreadsheet_drawing import AbsoluteAnchor, OneCellAnchor, AnchorMarker # type: ignore
from libs.xml2excel.utils import util, book
from libs.xml2excel.tag import sheet
from libs.xml2excel.section import flowchart
for _m in [util, sheet, book, flowchart]:
    importlib.reload(_m)

class xml2excel:

    def __init__(self, project_path, xlsx_ignore_folders=[]):
        self._project_path = ""
        self._project_tmp_path = ""
        self._project_dict = {}
        self._xlsx_ignore_folders = xlsx_ignore_folders
        self._JST = timezone(timedelta(hours=+9), 'JST')
        util.checkAndSetProject(self, project_path)
    
    def createOutputs(self):
        file_paths = util.createInOutputExcelPath(self)
        
        total_step = 0
        for path in file_paths:
            for key, file in path["input_sheet_paths"].items():
                total_step = total_step + 1

        for path in tqdm(file_paths, desc='Progress (files)', leave=False):
            wb = openpyxl.Workbook()
            index = 0
            self.diagram_data = []
            self.diagram_number_array = []

            for key, file in tqdm(path["input_sheet_paths"].items(), desc='Progress (sheets)', leave=False):

                self.diagram_number = 2  # なぜか、「2」始まりっぽい、多分だけど

                self.sheet_index = index
                if index == 0:
                    wb.worksheets[index].title = key
                else:
                    wb.create_sheet(title=key)            
                xml_dict = util.getDictFromXmlPath(self, file)

                # 基本レイヤーの設定
                xml_dict["sheet"] = sheet.baseFormatting(self, xml_dict["sheet"])
                sheet.baseRendering(self, wb.worksheets[index], xml_dict)

                # 直近のファイルとの比較
                _past_data = {}
                if key in path["past_path"] and len(path["past_path"][key]) > 0:
                    _past_file_path = path["past_path"][key][-1]
                    _past_xml_dict = util.getDictFromXmlPath(self, _past_file_path)
                    _past_xml_dict["sheet"] = sheet.baseFormatting(self, _past_xml_dict["sheet"])
                    _past_data = sheet.formatting(self, _past_xml_dict, file)  # データフォーマット

                data = sheet.formatting(self, xml_dict, file)  # データフォーマット

                # 修正箇所について押さえるためには、一旦描画レベルの情報を生成してから引っ張ってくる必要がある
                # 処理イメージ
                #   1. compareフェーズで「追加、更新、削除」状態を取得する
                #   2. 描画タイミングでフラグをみて、selfスコープにpos, row、更新文章関連のデータを溜める
                #   3. 更新シート以外の描画が完了したタイミングで、更新シートについて更新をかけていく

                if  _past_data != {}:
                    data = sheet.compare(self, data, _past_data)

                # diagramの直接制御部分とのindexズレが起きるので、こっちでダミー画像登録しておく
                _tmp_image_path = util.createText2ImagePNG(self, " ", (252,255,255))
                image = openpyxl.drawing.image.Image(_tmp_image_path)
                size_ext = XDRPositiveSize2D(1, 1)
                maker = AnchorMarker(col=0, colOff=0, row=0, rowOff=0)
                image.anchor = OneCellAnchor(_from=maker, ext=size_ext)
                wb.worksheets[index].add_image(image) 
                self.diagram_number = self.diagram_number + 1
                
                sheet.rendering(self, wb.worksheets[index], data, index) # シート描画  
                index = index + 1

                self.diagram_number_array.append(self.diagram_number)

            if path["file_dir"] not in self._xlsx_ignore_folders:            
                wb.save(path["output_file_path"])
                wb.close()

                flowchart.preset(self) # フロー図 データ整理

                if "✅ XLSX解析用ファイル展開"[0] == "":
                    # 解析ファイル取り込むおおよその手順
                    #   ① できるだけ単純なファイルを作る
                    #   ② drawing1.xmlの対象構造を templates フォルダに入れる
                    #   ③ book.pyファイルをいじる
                    xlsx_path = "/workdir/projects/サンプル_プロジェクト/sample_excels/017_clear_textbox.xlsx"
                    book.init(self, xlsx_path)
                    os.remove( xlsx_path.split(".")[0] + ".zip")
                else :

                    is_data = False
                    for _v in self.diagram_data:
                        if len(_v) > 0:
                            is_data = True
                    if is_data:
                        xlsx_path = path["output_file_path"]
                        book.init(self, xlsx_path)
                        for _sheet_key, _sheet_val in enumerate(self.diagram_data):
                            for _v in _sheet_val:

                                _connect_val = {}
                                if "start_id" in _v:
                                    _connect_val = {
                                        "s_id"  : _v["start_id"],
                                        "s_idx" : _v["start_idx"],
                                        "e_id"  : _v["end_id"],
                                        "e_idx" : _v["end_idx"],
                                    }

                                _emu_val = {}
                                if "max_width_cell_EMU" in _v:
                                    _emu_val = {
                                        "max_width_cell_EMU"  : _v["max_width_cell_EMU"],
                                        "max_height_cell_EMU"  : _v["max_height_cell_EMU"],
                                    }

                                book.add_textbox(
                                    self, 
                                    sheet_name="sheet" + str(_sheet_key + 1), # 連番で行けそう
                                    text=_v["text"]["value"],
                                    anchor = {
                                        "from": {"col": _v["row"], "row": _v["pos"], "colOff": 0, "rowOff": 0},
                                        "to": {"col": (_v["row"] + _v["row_len"]), "row": (_v["pos"] + _v["pos_len"]), "colOff": 0, "rowOff": 0},
                                    },
                                    _type=_v["type"],
                                    _id=_v["id"],
                                    _connect = _connect_val,
                                    _emu = _emu_val
                                )
                        book.write_as_xlsx(self, path["output_file_path"].split("/")[-1]) # 第三引数：False・・・で生成直前のファイル構成を確認できます

        # 一次作成ファイル削除
        dirs = os.listdir(self._project_tmp_path)
        for file in sorted(dirs):
            _path = self._project_tmp_path + "/" + file
            if os.path.isfile(_path):
                os.remove(_path)
            elif os.path.isdir(_path):
                shutil.rmtree(_path)
        shutil.rmtree(self._project_tmp_path)
