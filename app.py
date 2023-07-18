from docx import Document


def slice_txt_into_list(full_txt, slice_length):
    sliced_list = []

    for i in range(0, len(full_txt), slice_length):
        sliced_list.append(full_txt[i : i + slice_length])

    return sliced_list


def main():
    new_filename = input("作成するファイル名を入力: ")
    full_txt = input("起案内容を入力:")
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
    document.save(f"{new_filename}.docx")


if __name__ == "__main__":
    main()
