# 労働力調査 基本集計（月次）の長期時系列データ


## basic-tabulation-converter.py
[a-1 主要項目](https://www.stat.go.jp/data/roudou/longtime/03roudou.html#hyo_1)が対象。

### 使用方法
1. `basic-tabulation-converter.py`の変数を編集し、xlsxファイルのパスと読み込むシート名を指定します。（デフォルトは以下の通り）
```
#読み込むexcelファイルのパス
XLSX_PATH = r'../data/lt01-a10.xlsx'
#読み込むシート名
SHEET_NAME = '季節調整値'
```

2. [Spyder](https://www.spyder-ide.org/)などのコンソール上で`basic-tabulation-converter.py`を実行した後、`execute_pipeline()`を呼び出すことで、指定したxlsxファイル&シートのDataFrameを生成できます。
```
%run "basic-tabulation-converter.py"
df = execute_pipeline()
```

### 使用例
例えば以下のようにすることで、「2008年1月から2020年10月までの、完全失業率（男女計）」の折れ線グラフをプロットできたりします。
```
....
import matplotlib.pyplot as plt
plt.close('all')
df = execute_pipeline()
s = query_data_series(df, (2008, 1), (2020, 10), 'Unemployment rate  (percent)', 'Both sexes')
s.plot()
```
