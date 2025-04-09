#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Timestamp: "2025-03-09 12:53:43 (ywatanabe)"
# File: /home/ywatanabe/win/documents/SigmaPlot-v12.0-SigMacro/sigmapylot/docs/from_imura_sama.py

THIS_FILE = "/home/ywatanabe/win/documents/SigmaPlot-v12.0-SigMacro/sigmapylot/docs/from_imura_sama.py"
import sys
import os
import win32com.client
import subprocess



;; main("/home/ywatanabe/win/desktop/SigMacro_v1.3.JNB")


#====================================================================
#-=  メイン関数
#====================================================================
def main(sPath: str, *args) -> None:

    # SigmaPlot.Application取得
    oSigmaPlot = win32com.client.Dispatch('SigmaPlot.Application')

    # マクロ実装のJNBファイルを開く
    sCmd = f'"{sPath}"' + " /runmacro"
    subprocess.run(sCmd, shell=True)
    print(sCmd)
    # 本来は専用Openメソッドで開くが、Pythonからだと何故かアクセスできないので一旦上記CMD利用で読込
    # nbVBLib = oSigmaPlot.Notebooks.Open(sPath)

    # 指定名称からノートオブジェクト(Notebook)取得
    nbVBLib = None
    sFileName = os.path.basename(sPath)
    for i in range(0, oSigmaPlot.Notebooks.Count):
        nb = oSigmaPlot.Notebooks(i)
        if nb.Name == sFileName:
            nbVBLib = nb
            break

    # 指定名称からマクロオブジェクト(NotebookItem)取得
    nbiMacro = nbVBLib.NotebookItems("Macro1")
    if not nbiMacro is None:

        # テキスト生成
        GenerateArgumentsText(os.path.dirname(sPath),*args)

        # マクロ実行
        nbiMacro.Run()
        print("マクロ実行完了")



#====================================================================
#-=  引数受け渡し用のテキストファイルを生成
#====================================================================
def GenerateArgumentsText(sPathDir: str, *args):

    # 引数が1つ以上渡されている場合
    if args:
        # カンマ区切りで引数を結合
        arguments_text = ','.join(map(str, args))

        # フォルダが存在しない場合はエラーを出す
        if not os.path.exists(sPathDir):
            print(f"指定されたフォルダが存在しません: {sPathDir}")
            return

        # フォルダに保存するファイル名を作成
        file_path = os.path.join(sPathDir, "arguments.txt")

        # ファイルに書き込む
        with open(file_path, 'w') as file:
            file.write(arguments_text)
            print(f"カンマ区切りのテキストが {file_path} に保存されました。")
    else:
        print("引数が渡されていません。")



#====================================================================
#-=  エントリーポイント
#====================================================================
if __name__ == '__main__':

    main("C:/.../VBALib.JNB")

# EOF