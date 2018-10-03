■目的
xlsxファイル、または、xlsxファイルが格納されたフォルダを指定し、
対象のxlsxファイル、または、対象ディレクトリ以下のxlsxファイル全てについて読み取り、
ファイル名＋".csv"のファイル名で、csvファイルを出力する。

■補足
出力されるcsvは、以下のような形式とする。
　・文字コード： utf-8
　・各文字はダブルクォーテーション囲み（"test"、など）

■事前準備
Python3とpandasライブラリを利用する。
下記を参照し、Windowsに動作環境を構築すること。

・Pythonのインストール
http://cheminformist.itmol.com/BEGINNER/?p=19

　※Pythonはバージョン3.5を利用する。
　　3.5より新しいバージョンでも良いが、その場合はPipfileの「python_version」の式を、該当のバージョン値に変更すること。

・pipライブラリのアップデート
下記コマンドを実行し、pipライブラリを最新化する。
　・python -m pip install --upgrade pip

・pipenvライブラリのインストール
必要なライブラリの定義を「Pipfile」に記載すると、pipenvを利用するだけで関連ライブラリを全てインストールする。
また、pipenvを呼び出したディレクトリのみにライブラリを反映させることが出来る。
Pythonのインストール、および、Pythonの実行ファイルへのパスを通した後、コマンドプロンプトを開き、下記コマンドを実行する。
　・python -m pip install pipenv
　　※"-m"オプションでpipライブラリを呼び出す。pipライブラリのinstallコマンドにより、pipenvライブラリをインストールする。

■利用法
・１．コマンドプロンプトを起動し、本ディレクトリ内に移動する。

・２．下記コマンドを実行し、Pipenvファイルを読み込み、pipenv仮想環境を起動させる。
　・python -m pipenv install
　・python -m pipenv shell

・３．pythonに、実行用ファイルを読み込ませ、引数としてファイル／ディレクトリを指定する。
　・python main.py ＜入力ファイル名／ディレクトリ名＞
