import os
import re
import hashlib
import uuid
import xmltodict # type: ignore
from pathlib import Path
from datetime import datetime, timedelta, timezone
from PIL import Image, ImageDraw, ImageFont # type: ignore

def setProjectPath(self, path):
    if os.path.isdir(path) :
        self._project_path = path
        self._project_tmp_path = path + "/tmps"
    else :
        self._project_path = ""
        self._project_tmp_path = ""

def getSelfHash(self):
    file_path = os.path.dirname(__file__) + "/" + os.path.basename(__file__)
    with open(file_path, 'rb') as file:
        fileData = file.read()
        hash_md5 = hashlib.md5(fileData).hexdigest()
    return hash_md5

def getSelfUpdateAt(self):
    file_path = os.path.dirname(__file__) + "/" + os.path.basename(__file__)
    t = os.path.getmtime(file_path)
    return datetime.fromtimestamp(t, self._JST)
    
def checkAndSetProjectFolder(self):
    dirs = os.listdir(self._project_path)
    is_outputs_folder = False
    is_docs_folder = False
    is_tmps_folder = False
    res = []
    for dir in dirs:
        if dir == "docs":
            is_docs_folder = True
        elif dir == "outputs":
            is_outputs_folder = True
        elif dir == "tmps":
            is_tmps_folder = True
    
    if is_docs_folder:
        self._project_dict["docs"] = {}
    
    if is_outputs_folder == False:
        os.makedirs(self._project_path + "/outputs")
        
    self._project_dict["outputs"] = []
    
    if is_tmps_folder == False:
        os.makedirs(self._project_path + "/tmps")

    self._project_dict["tmps"] = []

    return res

def checkAndSetVersionFolder(self):
    if "docs" not in self._project_dict:
        return 

    dirs = os.listdir(self._project_path + "/docs")
    for dir in sorted(dirs):
        if re.match('^v(\d+)\.(\d+)\.(\d+)$', dir):
            self._project_dict["docs"][dir] = {}

def checkAndSetFileFolder(self):
    if "docs" not in self._project_dict:
        return 
    
    for version_dir in self._project_dict["docs"]:
        dirs = os.listdir(self._project_path + "/docs/" + version_dir)
        for dir in sorted(dirs):
            if re.match('^(\d+)_.*$', dir):
                self._project_dict["docs"][version_dir][dir] = []

def checkAndSetSheetFolder(self):
    if "docs" not in self._project_dict:
        return 
    
    for version_dir in self._project_dict["docs"]:
        for file_dir in self._project_dict["docs"][version_dir]:
            dirs = os.listdir(self._project_path + "/docs/" + version_dir + "/" + file_dir)
            for file in sorted(dirs):
                if re.match('^(\d+)_.*\.xml$', file):
                    self._project_dict["docs"][version_dir][file_dir].append(file)

def checkAndSetProject(self, path):
    setProjectPath(self, path)
    checkAndSetProjectFolder(self)
    checkAndSetVersionFolder(self)
    checkAndSetFileFolder(self)
    checkAndSetSheetFolder(self)

def getDictFromXmlPath(self, xml_path):
    if os.path.isfile(xml_path) :
        xml = Path(xml_path).read_text(encoding="utf-8")
        return xmltodict.parse(xml)
    else :
        return 

def createInOutputExcelPath(self):
    if "docs" not in self._project_dict:
        return 
    
    file_paths = []

    # 過去構造の比較が配列の並び順に依存しているので注意する
    past_version_path = {}

    last_version_dir = next(reversed(self._project_dict["docs"]), None)

    # 最新のバージョンのみを基本的に生成対象とする
    # 過去バージョンも一斉に更新するか？...悩みどころ、パフォーマスが悪ければ考慮する
    for version_dir in self._project_dict["docs"]:
        for file_dir in self._project_dict["docs"][version_dir]:
            input_sheet_paths = {}
            output_file_path = self._project_path + "/outputs/" + version_dir + "/" + file_dir + ".xlsx"

            if os.path.isdir(self._project_path + "/outputs/" + version_dir) != True:
                os.mkdir(self._project_path + "/outputs/" + version_dir)

            past_path = {}

            for file in self._project_dict["docs"][version_dir][file_dir]:
                input_sheet_paths[file.replace(".xml", "")] = self._project_path + "/docs/" + version_dir + "/" + file_dir + "/" + file

                if file_dir not in past_version_path:
                    past_version_path[file_dir] = {}
                
                if file not in past_version_path[file_dir]:
                    past_version_path[file_dir][file] = []

                if file_dir in past_version_path and file in past_version_path[file_dir] and len(past_version_path[file_dir][file]) > 0:
                    for past in past_version_path[file_dir][file]:
                        if file.replace(".xml", "") not in past_path:
                            past_path[file.replace(".xml", "")] = []
                        past_path[file.replace(".xml", "")].append(past)

                past_version_path[file_dir][file].append(
                    self._project_path + "/docs/" + version_dir + "/" + file_dir + "/" + file
                )

            # 全バージョンのファイルを生成するのであれば、この制限を解除する
            if last_version_dir == version_dir:
                file_paths.append({
                    "input_sheet_paths" : input_sheet_paths,
                    "output_file_path" : output_file_path,
                    "past_path" : past_path,
                    "file_dir" : file_dir,
                })

    return file_paths

def createText2ImagePNG(self, word, color=()):

    background_color = (255,255,255,0)
    if word != " ":
        word = str(convNumberMarusuuji(word))
    else:
        word = " "
    font_path = '/usr/share/fonts/truetype/fonts-japanese-gothic.ttf'
    text_size = 60
    text_color = color
    if len(word) == 1:
        text_size = 60
        background_size = (60,60)
        background = Image.new('RGBA', background_size, background_color)
        text_coordinate = (0,0) # -17
        text_padding = (0,0)
    elif len(word) == 2:
        text_size = 52
        background_size = (60,60)
        background = Image.new('RGBA', background_size, background_color)
        text_coordinate = (6,0) # -11
        text_padding = (0,-22)
    elif len(word) == 3:
        text_size = 38
        background_size = (60,60)
        background = Image.new('RGBA', background_size, background_color)
        text_coordinate = (16,0) # -1
        text_padding = (0,-19)

    draw_img = drawTxt(background, word, font_path, text_size, text_color, text_coordinate, text_padding)
    save_path = self._project_path + '/tmps/' + str(uuid.uuid1()) + '.png'
    draw_img.save(save_path)

    return save_path

def drawTxt(background, word, f_path, t_size , t_color, t_cood, t_pad):
    font = ImageFont.truetype(f_path, t_size)
    draw = ImageDraw.Draw(background)
    cnt = 0
    for y in range(t_cood[0], background.size[1]-t_cood[0], t_size+t_pad[0]):
        for x in range(t_cood[1], background.size[0]-t_cood[1], t_size+t_pad[1]):
            if len(word) > cnt:
                for xx in range(-5,5 + 1):
                    for yy in range(-5,5 + 1):
                        draw.text((x+xx,y+yy), word[cnt], font=font, fill=(255,255,255,255))
                for xx in range(-1,1 + 1):
                    for yy in range(-1,1 + 1):
                        draw.text((x+xx,y+yy), word[cnt], font=font, fill=t_color)
                draw.text((x,y), word[cnt], font=font, fill=t_color)
            else:
                break
            cnt += 1
    return background

def convNumberMarusuuji(number):
    _marusuuji_array = ["⓪","①","②","③","④","⑤","⑥","⑦","⑧","⑨","⑩","⑪","⑫","⑬","⑭","⑮","⑯","⑰","⑱","⑲","⑳","㉑","㉒","㉓","㉔","㉕","㉖","㉗","㉘","㉙","㉚","㉛","㉜","㉝","㉞","㉟","㊱","㊲","㊳","㊴","㊵","㊶","㊷","㊸","㊹","㊺","㊻","㊼","㊽","㊾","㊿"]
    if len(_marusuuji_array) > int(number):
        return _marusuuji_array[int(number)]
    else :
        return number