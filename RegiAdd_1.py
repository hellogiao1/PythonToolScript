import wpsTool
import os

TitleDic = wpsTool.TitleDic


class readWrite_excel(wpsTool.ReadWrite_excel):
    def loop_write(self):
        self.ws.cell(self.currRow, 1, TitleDic["申请人姓名"])
        self.ws.cell(self.currRow, 2, TitleDic["申请人电话"])
        self.ws.cell(self.currRow, 3, "上海市" + TitleDic["充电设备安装区域"] + TitleDic["充电设备安装街区"] + " ")
        self.ws.cell(self.currRow, 4, TitleDic["充电桩编号"])
        self.currRow += 1
        self.dic_restore()
        self.wb.save(self.filename)


class transfer_format(wpsTool.Transfer_format):
    @staticmethod
    def ModifyWordName(fn, VIC):
        """Vehicle identification code，车辆识别代码"""
        if VIC not in fn:
            strpath = fn[:(fn.rfind("\\") + 1)] + str(VIC) + "-" + fn[fn.rfind("\\") + 1:]
            print(strpath, VIC)
            if not os.path.exists(strpath):
                os.rename(fn, strpath)
                fn = strpath
        return fn


def LoopInsertData(file_list, excelPath):
    if not excelPath:
        print("表格路径为空")
        return

    tf = transfer_format()
    we = readWrite_excel(excelPath)
    for fn in file_list:
        wpsTool.Read_DocData(fn)

        if TitleDic["车辆识别代码"] != "":
            fn = tf.ModifyWordName(fn, TitleDic["车辆识别代码"])
            # tf.word2pdf(fn)
            we.loop_write()

    return


if __name__ == '__main__':
    currDir_list = wpsTool.get_all_doc(os.getcwd())
    excel_file = currDir_list[1]
    print('一共找到' + str(len(currDir_list[0])) + '个Word文档', str(len(currDir_list[2])) + '张图片')
    print("Excel表路径为：" + excel_file)
    LoopInsertData(currDir_list[0], excel_file)

    print("生成成功！按任意键退出")
    print("---初试版本v1.0---DaimQ")
    os.system('pause')
