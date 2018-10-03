#!/usr/bin/env python3
# coding: utf-8

import sys
import os
import pandas
import csv
import glob


def main():
    """
    メイン処理
    """
    # コマンドライン引数を格納したリストの取得
    argvs = sys.argv
    # 引数の個数取得
    argc = len(argvs)

    # 引数が足りない場合は、その旨を表示
    if (argc != 2):
        print('Usage: # python %s filename' % argvs[0])
        print(' OR')
        print('Usage: # python %s directory_name' % argvs[0])
        # プログラムの終了
        sys.exit()

    # 引数の相対パスより、フルパスを取得
    inFileName = argvs[1]
    targetPath = os.path.abspath(inFileName)
    print('read path: ', targetPath)

    # 引数で指定されたパスについて、ファイルの場合／ディレクトリの場合で処理を分岐
    if (os.path.isdir(targetPath)):
        # 指定ディレクトリ以下のExcelファイルを全量読み込む
        read_dir(targetPath)
    else:
        # パス指定されたExcelファイルを読み込む
        read_file(targetPath)

    print("!!! done !!!!")


def read_dir(dirFullPath):
    """
    指定のディレクトリ内に存在するExcelファイル全てを読み込み、指定のcsvファイルに書き出す。
    @dirFullPath 読み込み対象ディレクトリのフルパス
    @csvFileName 書込み対象のcsvファイル名。プログラムが存在するディレクトリに作成・追記する。
    """

    # 指定ディレクトリに格納されたファイルの内、Excelファイルの拡張子（xlsx）を有するファイルについて、全て読み込む
    for filePath in glob.iglob(os.path.join(dirFullPath, "*.xlsx")):
        read_file(filePath)

    print("all file readed.")


def read_file(targetPath):
    """
    指定のExcelファイルを読み込み、指定のcsvファイルに書き出す。
    @targetPath 読み込み対象Excelファイルのフルパス
    """
    print("reading... : ", targetPath)
    df = read_excel(targetPath)
    csv_file = out_csv(df, targetPath)

    last_str_del(csv_file)


def read_excel(targetPath):
    """
    指定パスのExcelファイルを読み込み、処理をした結果のpandas.DataFrame型の変数を返却する。
    @param targetPath 読み込み対象のExcelファイルのフルパス
    @return 処理をした結果のpandas.DataFrame型の変数
    """

    # df1: 1シート目の読み込み。
    df1 = pandas.read_excel(io=targetPath,
                            sheet_name=0,
                            header=0,
                            skiprows=None,
                            index_col=None,
                            usecols=None
                            )

    # 未設定セルに設定される欠損値（Nan）を、空文字（""）に変更（今後の操作のし易さのため）
    df1 = df1.fillna("")

    return df1


def out_csv(df, inFilePath):
    """
    指定のpandas.DataFrame型の変数の中身をcsv出力する。
    @param df         csv出力対象データの格納された、pandas.DataFrame型のインスタンス
    @param inFilePath 入力Excelファイルフルパス
    @return           出力したcsvファイルのフルパス
    """

    # csv出力（出力ファイル名： out_<入力ファイル名>.csv）
    outFileName = "out_" + os.path.basename(inFilePath) + ".csv"
    outTgtPath = os.path.join(os.path.dirname(inFilePath), outFileName)

    # 上書き保存。ヘッダー付き、ダブルクォーテーション囲みを指定。
    df.to_csv(outTgtPath, sep=',', na_rep='',
              header=True, index=False, index_label=True, mode='w', encoding="utf-8", compression=None,
              quotechar='"', line_terminator='\n', chunksize=None, doublequote=True,
              quoting=csv.QUOTE_ALL, date_format='%Y/%m/%d %H:%M:%S'
              )

    return outTgtPath


def last_str_del(targetFilePath):
    """
    最終文字を削除する。（空行出力されるファイルに対し、最終行、最終文字にEOFが来るようにする。）
    @param targetFilePath 書き換え対象ファイルのフルパス 
    """
    in_txt = ''
    with open(targetFilePath, 'r', encoding='utf-8', newline='\n') as in_f:
        in_txt = in_f.read()
        in_f.close()

    with open(targetFilePath, "w", encoding='utf-8', newline='\n') as out_f:
        out_f.write(in_txt[:-1])
        out_f.close()


########################################
# 直接実行された場合の処理
if __name__ == "__main__":
    # execute only if run as a script
    main()
