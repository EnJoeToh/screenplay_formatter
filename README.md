# screenplay_formatter
なんとなくマークアップされたテキストデータ(.txt, .docx)をシナリオっぽく整形します。

# Demo
こういうかんじのものを、
<img width=50% alt="0" src="https://user-images.githubusercontent.com/8622918/128624089-7b5274e9-78b6-4811-9978-130e27f5d501.png">

こういうかんじのものへ。
<img width=50% alt="1" src="https://user-images.githubusercontent.com/8622918/128624095-b05adec1-9fad-4e60-bf00-686b0316039e.png">

# Requirement
* python3
* python-docx

[python-docx](https://python-docx.readthedocs.io/en/latest/)

```bash
pip install python-docx
```

# Usage
- マークアップしたテキスト filename.txt を用意。
- screenplay_formatter と同じフォルダに配置。
- コマンドラインから、
```bash
python screenplay_formatter filename.txt
```
　　を実行。
- filename_formatted.html と filename_formatted.docx が同フォルダに生成される。

# マークアップ
☆タイトル  
★サブタイトル  
■柱  
＠ト書き  
人物名「セリフ」  
＃注記  
→右寄せメモ  
など



