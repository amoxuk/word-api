# encoding=utf-8
from __future__ import print_function, unicode_literals

import math
import os
import traceback
from random import random
from time import sleep

import win32com.client


class Word:
    def __init__(self):
        pass

    def word(self):
        # word = win32com.client.Dispatch('et.Application')
        word = win32com.client.Dispatch('word.Application')
        word.Visible = 1  # 后台运行
        word.DisplayAlerts = 0  # 不显示，不警告
        # 打开一个已有的word文档
        doc = word.Documents.Open(os.path.dirname(__file__) + os.sep + u'软件测试报告.docx', Encoding="GBK")
        # doc = word.Documents.Add()  # 创建新的word文档
        # data = doc.paragraphs[0].text

        # 在文档开头添加内容
        # myRange1 = doc.Range(0, 0)
        # myRange1.InsertBefore('Hello word')

        # 在文档末尾添加内容
        # myRange2 = doc.Range()
        # myRange2.InsertAfter('Bye word')

        # 在文档i指定位置添加内容
        # myRange3= doc.Range(0, insertPos) # insertPos为数字
        # myRange3.InsertAfter('what's up, bro?')
        # word.Selection.Find.ClearFormatting()
        # word.Selection.Find.Replacement.ClearFormatting()
        # word.Selection.Find.Execute(OldStr, False, False, False, False, False, True, 1, True, NewStr, 2)
        '''
        上面涉及的 11 个参数说明
                 (OldStr--搜索的关键字,
                 True--区分大小写,
                 True--完全匹配的单词，并非单词中的部分（全字匹配）,
                 True--使用通配符,
                 True--同音,
                 True--查找单词的各种形式,
                 True--向文档尾部搜索,
                 1,
                 True--带格式的文本,
                 NewStr--替换文本,
                 2--替换个数（0表示不替换，1表示只替换匹配到的第一个，2表示全部替换）
        '''

        # self.xlApp.Selection.Find.ClearFormatting()
        # self.xlApp.Selection.Find.Replacement.ClearFormatting()

        # 循环操作，将每个匹配到的关键词进行换色
        # while word.Selection.Find.Execute("测试项目", False, False, False, False, False, True, 0, True, "", 0):
        #     word.Selection.Range.HighlightColorIndex = 11  # 替换背景颜色为绿色
        #     word.Selection.Font.Color = 255  # 替换文字颜色为红色
        # 页眉文字替换

        # word.ActiveDocument.Sections[0].Headers[0].Range.Find.ClearFormatting()
        # word.ActiveDocument.Sections[0].Headers[0].Range.Find.Replacement.ClearFormatting()
        # word.ActiveDocument.Sections[0].Headers[0].Range.Find.Execute(OldStr, False, False, False, False, False, True,
        #                                                               1, False, NewStr, 2)
        # 选中第一个表的第一行，选中之后就可以进行更改等的操作
        # doc.Tables[0].Rows[0].Range.Select()
        # ActiveDocument.Shapes(1).Anchor.Paragraphs(1).Range.Select
        # 打印
        # doc.PrintOut()
        # 根据表格第一行查找表格
        # for table in doc.Tables:
        #     if u'测试项目' in [cell.Range.Text.strip("\r\x07").strip() for cell in table.Rows[0].Cells]:
        #         # for no in range(1, table.Rows.Count):
        #         #     # 打印行数据
        #         #     # print(table.Rows[no].Range.Text)
        #         #     # 删除行列，每次操作之后行会发生改变
        #         #     table.Rows[1].Delete()
        #         # 追加一行，函数返回该行
        #         row = table.Rows.Add()
        #         # 追加到指定的行之前
        #         # row = table.Rows.Add(BeforeRow=table.Rows[1])
        #         for cell in row.Cells:
        #             # 为每个单元格添加内容
        #             cell.Range.Text = random()
        #     else:
        #         continue
        # Shape 代表一个图形层对象，例如自选图形、任意多边形、OLE 对象、ActiveX 控件、图片等。
        for shape in doc.Shapes:
            # obj Type: ActiveX
            # 12    : ActiveX
            # 1     : 形状
            print("shape")
            print(shape.Type)
            print(shape.Id)
        for inline_shape in doc.InlineShapes:
            #
            # Application属性            # 
            # Borders属性            # 
            # Creator属性            # 
            # Description属性            # 
            # Fill属性            # 
            # Height属性            # 
            # LockAspectRatio属性            # 
            # Parent属性            # 
            # PictureFormat属性            # 
            # Range属性            # 
            # ScaleHeight属性            # 
            # ScaleWidth属性            # 
            # TextEffect属性            # 
            # Type属性            返回修订类型，增加或删除
            # Width属性            #
            # Activate方法            #
            # ConvertToShape方法            # 
            # Delete方法            # 
            # Select方法
            #
            print("shape inline")
            # ActiveX 5
            # Picture 3
            # Figure  3
            if inline_shape.Type == 5:
                print("ActiveX")
                # OLEFormat 该属性由shape获取ActiveX对象
                # 函数： Activate();ActivateAs();ConvertTo();DoVerb();Edit();Open()
                # 属性：
                ac = inline_shape.OLEFormat.Activate()
                ac = inline_shape.OLEFormat.DoVerb(-4)
                # ClassType： Forms.OptionButton.1 单选按钮
                cp = inline_shape.OLEFormat.ClassType
                print(cp)
                sleep(3)

            # print(inline_shape.Delete())
            # print(inline_shape.Description)






        # 遍历表格
        # for table in doc.Tables:
        #     # 遍历行
        #     for row in table.Rows:
        #         # 遍历单元格
        #         if u'测试项目' in [cell.Range.Text.strip("\r\x07") for cell in row.Cells]:
        #             pass
        #         else:
        #             row.Delete()

            # 遍历列
            # for col in table.Columns:
                # print(col.Cells)
                # 查找表
                # if u'测试项目' in [cell.Range.Text.strip("\r\x07") for cell in col.Cells]:
                #     print(col)
                # else:
                #     col.Delete()

                # for cell in col.Cells.Range.Text:
                #     print(cell)

        # 保存
        # doc.Save()  # 保存
        # doc.SaveAs(os.path.dirname(__file__) + os.sep + 'test_save.docx')  # 另存为

        # 退出
        #
        # 　　退出操作必须得做，不然进程就会一直占据着这个文件，下次操作相同文件的时候就会报错

        # doc.Close()  # 关闭 word 文档
        # word.Documents.Close(doc.wdDoNotSaveChanges)  # 保存并关闭 word 文档
        # word.Quit()  # 关闭 office


if __name__ == '__main__':
    word = Word()
    try:
        word.word()
    except Exception as e:
        # print(e.args[1].decode('gbk'))
        # print(e.args[2][2])
        traceback.print_exc(e)
