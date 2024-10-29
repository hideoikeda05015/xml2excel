import copy
import json
import Levenshtein # type: ignore
import itertools
import pprint
from importlib import reload
from libs.xml2excel.utils import cell
reload(cell)

def compare(self, data, past_data):
    data = copy.deepcopy(data)
    past_data = copy.deepcopy(past_data)
    data_array = []
    past_array = []

    for key, val in enumerate(data["text"]):
        data_array.append(val["value"])

    for key, val in enumerate(past_data["text"]):
        past_array.append(val["value"])

    p = itertools.product(range(len(data_array)), range(len(past_array)))

    # pprint.pprint(list(p)) 

    #for data_key, data_val in enumerate(data_array):
    #    for past_key, past_val in enumerate(past_array):
    #        _s1 = data_val[:6] + "..."
    #        _s2 = past_val[:6] + "..."
    #        _s3 = Levenshtein.distance(data_val, past_val)
    #        print("%s : %s => %s" % (_s1, _s2, _s3))

    #print(json.dumps(data_array, indent=2, ensure_ascii=False))
    #print(json.dumps(past_array, indent=2, ensure_ascii=False))

    return data

def rendering(self, sheet, start_pos, data):
    pos = cell.setInlineTitle(self, sheet, start_pos, data["title"])

    if len(data["text"]) == 0:
        pos = pos - 1
    else :
        for val in data["text"]:
            cell.setCell(self, 
                            sheet, 
                            pos, 1, 
                            self._start_col + self._body_lr_padding + self._page_lr_padding, 
                            self._width_size - self._body_lr_padding * 2 - self._page_lr_padding * 2, 
                            "FFFFFF", "left", val)
            pos = pos + 1
            
    # print(json.dumps(data, indent=2, ensure_ascii=False))

    return pos

def formatting(self, data):
    data = copy.deepcopy(data)

    if "title" not in data:
        data["title"] = ""

    texts = []
    if "text" in data and data["text"] is not None:
        texts = data["text"].split("\n")
        for index in range(len(texts)):
            texts[index] = texts[index].lstrip(" ").rstrip(" ").lstrip("\t").rstrip("\t")
    data["text"] = texts

    data["title"] = {
        "value" : data["title"],
    }

    tmp = []
    for val in data["text"]:
        tmp.append({
        "value" : val,
    })
    data["text"] = tmp

    return data
