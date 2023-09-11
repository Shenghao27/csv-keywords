#!/usr/bin/env python
# coding: utf-8

# In[ ]:


# 匯入所需的庫
import glob  # 用於查找符合特定模式的文件
import os  # 用於操作文件和文件夾路徑
import numpy as np  # 用於數值計算
import pandas as pd  # 用於處理數據表格
import re  # 用於正則表達式操作

try:
    # 獲取用戶輸入的文件夾路徑
    folder_path = input("請輸入資料夾路徑：")
    # 使用glob模塊獲取文件夾中所有以.txt結尾的文件的列表
    txt_list = glob.glob(os.path.join(folder_path, '*.txt'))

    all_list = []

    for txt in txt_list:
        with open(txt, 'r', encoding='utf-8') as f:
            text = f.read()
            # 查找文本中是否包含 'CCS通訊異常' 這個字符串
            if re.search(r'CCS通訊異常', text):
                # 使用正則表達式查找 'CCS通訊異常' 出現的位置
                CCS_list = re.finditer(r'CCS通訊異常', text)
                for s in CCS_list:
                    # span() 方法来獲取每個查找字的起始和结束索引
                    Span = s.span()
                    # 因為span()的傳回值是tuple，需轉成字串後比較好處理並接字串
                    key_1 = str(Span)
                    # 把 () 去除，並用 , 做分隔元素
                    key_2 = key_1.strip('()').split(',')
                    # 取第[0]組元素
                    key_3 = int(key_2[0])
                    # 根據位置提取時間訊息
                    timestamp = text[key_3-29:key_3-9]  # 修改字元位置
                    # 提取錯誤消息
                    Error_message = text[key_3:].split('\n')[0] # 拆分成以 \n 為分隔符的列表，然後選取列表中的第[0]部分
                    # 將時間和錯誤消息添加到列表中
                    all_list.append([timestamp, Error_message])

    # 創建包含時間戳和錯誤消息的數據框
    df = pd.DataFrame(all_list, columns=['Timestamp', 'Error_message'])

    # 獲取用戶輸入的要保存的 Excel 文件名
    output_filename = input("請輸入要輸出的 Excel 檔案名稱：")
    output_filename = output_filename + ".xlsx"
    # 將數據框保存為 Excel 文件
    df.to_excel(output_filename, encoding='UTF8')

    print("輸出成功！")

except Exception as e:
    print("發生錯誤：", str(e))

