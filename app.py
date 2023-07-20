from docx import Document
from flask import Flask, request, render_template

app = Flask(__name__)


@app.route("/")
def index():
    # HTMLフォームを表示するためのテンプレートを返す
    return render_template("input_form.html")


@app.route("/get_text", methods=["POST"])
def get_text():
    # フォームから入力されたテキストを取得
    input_text = request.form["input_text"]
    filename = request.form["filename"]

    # Word文書を生成して保存
    generate_docx(filename, input_text)

    # 応答メッセージを返す
    return f"ファイル '{filename}.docx' が保存されました。"


# 全文を改行及び文字数で区切ってリストに入れる
def slice_txt_into_list(full_txt, slice_length):
    sliced_list = []
    lines = full_txt.split("\n")  # 改行でテキストを分割

    for line in lines:
        while len(line) > slice_length:
            sliced_list.append(line[:slice_length])
            line = line[slice_length:]

        if line:
            sliced_list.append(line)

    return sliced_list


def generate_docx(new_filename, full_txt):
    # 40字ごとに区切ってリストに入れる
    draft_content = slice_txt_into_list(full_txt, 40)

    # 既存のWord文書を開く
    document = Document("TEST用起案文書様式.docx")

    # 文書内の表を取得する
    table = document.tables[0]

    # 起案文書様式の６行目から入力
    line_num = 6

    for draft in draft_content:
        # 表のline_num行目、0列目にあるセルを取得する
        cell = table.cell(line_num, 0)
        # セル内のテキストを上書きする
        cell.text = draft
        line_num += 1

    # 上書きした文書を保存
    output_path = f"{new_filename}.docx"
    document.save(output_path)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=8000, debug=True)
