# -*- coding: utf-8 -*-
"""
先建立ppt library，之後根據此資料搜尋關鍵字，可加速第二次以後的搜尋


實現:
先建立ppt library，結構為dict: key為檔名, value為list，list中塞該檔中不同的內容
"""
import os
import json
import msvcrt
from win32com.client import Dispatch
from pathlib import Path

from configs.log_config import get_logger as logger


class PowerPoint_keyword_search():
    def __init__(self):
        pass

    def filter(self, items):
        file_list = []
        for names in items:
            if (names[0] != '~') and (".ppt" in names):
                file_list.append(names)
        return file_list

    # TODO 函式跑完後，PPT沒有關閉完全，但不影響後續程式運行
    def convert_ppt_into_dict(self, file_list, path=None):
        ppt_library = dict()
        try:
            for each_file in file_list:
                filename = Path(each_file).name
                ppt_library[filename] = []
                ppt = Dispatch('PowerPoint.Application')
                pptSel = ppt.Presentations.Open(str(each_file))
                slide_count = pptSel.Slides.Count
                for i in range(1, slide_count + 1):
                    shape_count = pptSel.Slides(i).Shapes.Count
                    for j in range(1, shape_count + 1):
                        if pptSel.Slides(i).Shapes(j).HasTextFrame:
                            s = pptSel.Slides(i).Shapes(j).TextFrame.TextRange.Text
                            if s != '':
                                ppt_library[filename].append(s)
                pptSel.Close()
        except Exception as e:
            logger().error(f"extractwords_into_dict failed: {e}, file_name is: {filename}")
            raise
        return ppt_library

    def find_keyword_from_library(self, keyword):
        result = ""
        try:
            with open("ppt_library.txt", "r") as f:
                library = json.load(f)
                for key, value in library.items():
                    cunt = 0
                    for eachvalue in value:
                        if keyword in eachvalue:
                            if cunt == 0:
                                result += f"檔案名稱: {key} \n"
                                cunt += 1
                            result += f"內容: {eachvalue} \n"
                    result += "\n"
            result = result.strip()
            return result if result else "很抱歉，您輸入的關鍵字查無結果。"
        except Exception as e:
            logger().error(f"find_keyword_from_library failed: {e}, keyword is: {keyword}")


def excute():
    ppt = PowerPoint_keyword_search()
    items = os.listdir()
    search_mode = int(input('要建立json dataset，請輸入1；要透過json dataset尋找keyword，請輸入2\n'))

    if search_mode == 1:
        file_list = ppt.filter(items)
        print(file_list)
        ppt_library = ppt.convert_ppt_into_dict(file_list)
        with open("ppt_library.txt", "w") as f:
            json.dump(ppt_library, f)
            print("write file complete")
            msvcrt.getch()
    elif search_mode == 2:
        mode = int(input('若關鍵字為一組，請輸入1；若關鍵字為2組，請輸入2\n'))
        ppt.find_keyword_from_library(mode)


if __name__ == "__main__":
    excute()
