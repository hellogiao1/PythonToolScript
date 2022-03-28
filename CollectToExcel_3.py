import os
import time
import wpsTool


def LoopInsertData(file_list, excelPath):
    if not excelPath:
        print("表格路径为空")
        return

    # 格式化成2020/10/1 形式
    dayTime = time.strftime("%Y/%m/%d", time.localtime())
    TitleDic["投入运营时间"] = dayTime
    tf = wpsTool.Transfer_format()
    we = wpsTool.ReadWrite_excel(excelPath)
    for fn in file_list:
        wpsTool.Read_DocData(fn)
        # 客户说不需要显示读取汽车销售商，他们自己输入
        TitleDic["汽车销售商"] = ""
        if TitleDic["车辆识别代码"] != "":
            fn = tf.ModifyWordName(fn, TitleDic["车辆识别代码"])
            tf.word2pdf(fn)
            we.loop_write()

    return


def test(flist):
    for fn in flist:
        wpsTool.Read_DocData(fn)
        print(TitleDic)


TitleDic = wpsTool.TitleDic

if __name__ == '__main__':
    # 获取上一级别目录中的指定文件
    currDir_list = wpsTool.get_all_doc(os.getcwd())
    # print(currDir_list[0])
    print('一共找到' + str(len(currDir_list[0])) + '个Word文档', str(len(currDir_list[2])) + '张图片')
    excel_file = currDir_list[1]
    print("Excel表路径为：" + excel_file)
    # test(currDir_list[0])
    LoopInsertData(currDir_list[0], excel_file)

    print("生成成功！按任意键退出")
    print("---初试版本v1.0---DaimQ")
    os.system('pause')
