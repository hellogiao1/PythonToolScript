import os

import wpsTool

TitleDic = wpsTool.TitleDic


class readwrite_excel(wpsTool.ReadWrite_excel):
    def loop_read(self, word_list, picture_list):
        excelCount_list = []
        for self.readRow in range(1, self.ws.max_row + 1):
            row_dic = {"申请人姓名": self.ws.cell(self.readRow, 1).value,
                       "申请人电话": self.ws.cell(self.readRow, 2).value,
                       "充电设备安装地址": self.ws.cell(self.readRow, 3).value,
                       "充电桩编码": self.ws.cell(self.readRow, 4).value,
                       "充电设备停车位": ""
                       }
            apply_name = row_dic["申请人姓名"]
            picturePath = ""
            wordPath = ""
            for pl in picture_list:
                if apply_name in pl:
                    picturePath = pl
                    break
            for wl in word_list:
                if apply_name in wl:
                    wordPath = wl
                    break
            if picturePath and wordPath:
                print("图片路径：" + picturePath, "word路径：" + wordPath)
                temp = row_dic["充电设备安装地址"]
                if temp.find("【") != -1 and temp.rfind("】") != -1:
                    row_dic["充电设备停车位"] = temp[temp.find("【") + 1:temp.find("】")]
                    self.write(self.readRow, 3, temp[:temp.find("【")]
                               + temp[temp.find("【") + 1:temp.find("】")]
                               + temp[temp.find("】") + 1:])
                wpsTool.Write_Doc(wordPath, picturePath, row_dic)
            elif wordPath:
                print("word路径：" + wordPath)
                temp = row_dic["充电设备安装地址"]
                if temp.find("【") != -1 and temp.rfind("】") != -1:
                    row_dic["充电设备停车位"] = temp[temp.find("【") + 1:temp.find("】")]
                wpsTool.Write_Doc(wordPath, picturePath, row_dic)
            elif not picturePath and not wordPath:
                print("未找到\"" + apply_name + "\"的图片和Word文档")
            elif not wordPath:
                print("未找到\"" + apply_name + "\"的Word文档")
            else:
                print("未找到\"" + apply_name + "\"的图片")

            excelCount_list.append(row_dic)

        return excelCount_list


def LoopInsertData(word_list, excelPath, picture_list):
    rwe = readwrite_excel(excelPath)
    rwe.loop_read(word_list, picture_list)

    return


if __name__ == '__main__':
    currDir_list = wpsTool.get_all_doc(os.getcwd())  # r"E:\Python\汇羿自查表3.27"
    print('一共找到' + str(len(currDir_list[0])) + '个Word文档', str(len(currDir_list[2])) + '张图片')
    print("Excel表路径为：" + currDir_list[1])
    LoopInsertData(currDir_list[0], currDir_list[1], currDir_list[2])

    print("生成成功！按任意键退出")
    print("---初试版本v1.0---DaimQ")
    os.system('pause')
