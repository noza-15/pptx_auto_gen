# Zemi PPTX auto generator
ゼミ発表用のパワーポイントスライド(リンク付き)のジェネレータ

## Descrption
ゼミ発表用のパワーポイントスライド(リンク付き)のジェネレータです．
CSVに書いておいたゼミの日程を取り込んで生成するモードと対話的に情報を入力して生成するモードの2つがあります．
ある程度エラーの管理はしてあるので，トラブルシューティングは実行時の表示を確認してください．

## Features
ゼミ発表用のパワーポイントスライド(リンク付き)の生成．
テンプレートのpptxファイルを指定するとテンプレートに基づいたスライドを生成します．
出力ファイル名は`YYMMDD.pptx`です．

### CSVから入力
ゼミ日程をCSVに書き込んでおき一括でスライドを生成可能．

- 必要な情報は出力先ディレクトリ，ゼミ名，日付，発表者のみ．
- 発表者の名前は英字可．別途名簿を用意することで漢字を補完．
- 発表者のパワポファイル名は自動補完．
- pptxをコピーするbatファイルを生成．

### 対話型で入力
- 出力先ディレクトリ，ゼミ名，日付，発表者を直接入力．
- 発表者の名前は英字可．別途名簿を用意することで漢字を補完．
- 発表者のパワポファイル名は自動補完．任意のファイル名を指定することも可能．

## Requirement
- Python3
- [python-pptx](https://python-pptx.readthedocs.io/en/latest/index.html#)
  - インストールしていない場合は`pip install python-pptx`

## Usage
  usage: pptx_auto_gen.py [-h] [-i] [-s SCHEDULE] [-n NAMELIST] [-t TEMPLATE]
     [-o OUT_DIR] [-u UPLOAD_DIR]

各オプションは以下の通り．
通常はCSV入力モードで実行します．
`-i`を付けると対話型モードで実行します．

ゼミ日程ファイル(schedule.csv), 名前辞書ファイル(name_dic.csv), スライドのテンプレート(template.pptx)は別途指定できます．

    optional arguments:
      -h, --help            show this help message and exit
      -i, --interactive     use interactive mode
      -s SCHEDULE, --schedule SCHEDULE
                            specify zemi schedule file (csv). default:
                            schedule.csv
      -n NAMELIST, --namelist NAMELIST
                            specify namelist file (csv). default: name_dic.csv
      -t TEMPLATE, --template TEMPLATE
                            specify template file (pptx). default: template.pptx
      -o OUT_DIR, --out-dir OUT_DIR
                            specify output directory. default: current directory
      -u UPLOAD_DIR, --upload-dir UPLOAD_DIR
                            specify directory containing pptx. default:
                            X:\2019\regular_seminar

## 各種ファイルの仕様について
### schedule.csv
ゼミ日程を記入するCSVファイルです．
1行目をタイトル行とします．
タイトル行のラベルが正しくないと読み込めません．
試していませんがゼミの数に制限はありません．

- 最初の3列は出力先ディレクトリ(ラベル名:dir), ゼミの日時(ラベル:date), ゼミの名前(ラベル:name)を順不同で, 最後の列に発表者(ラベル(1人目のみ):presenter)を記述してください．
    - dir: ディレクトリ名
実際に出力されるディレクトリ名には後ろに日付が付きます．
    - date: 日付
YY/MM/DDまたはYY/M/Dの形式で入力．
    - name: ゼミの名前
特に制限なし．
    - presenter: 発表者
1人目のみラベルが必要です．必ず**最後の列**に記載してください．人数制限はありませんが多すぎるとスライドのページから溢れます．name_dic.csvで登録してある名前だと，スライドのリンクテキストを登録名にできます．この名前をもとにリンク先のファイル名を生成します．

__例__

    dir,date,name,presenter,,
    zemi01,2019/4/8,第1回合同定期ゼミ,田島 咲季,bao,hasegawa
    zemi02,2019/4/15,第2回合同定期ゼミ,oku,葉 静浩,ishikawa
    zemi03,2019/4/22,第1回通常定期ゼミ,inoue,usami,

例の場合，1つ目のゼミに関しては`zemi01_190408`フォルダの中に`190408.pptx`が生成されます．

### name_dic.csv
英語名から日本語名に変換する名簿リストです．
1行目をタイトル行とします．
タイトル行のラベルが正しくないと読み込めません．

- 英語名(ラベル:en)と日本語名(ラベル:ja)の2列を用意してください．順不同です．
    - en: 英語名
`名前.苗字`か`苗字`としてください．
    - ja: 日本語名
スライドのリンクに表示される名前です．
- 辞書に名前がない場合は表示名は入力通りになります．


__例__

    en,ja
    nozomu.togawa,戸川 望
    shigenobu.okuma,大隈 重信

### template.pptx
スライドのもととなるpptxファイルです．
フォントの種類や大きさ，テーマなどを設定しておくことでそれに沿ったスライドが生成されます．
存在しない場合は，プレーンのテーマのスライドが生成されます．