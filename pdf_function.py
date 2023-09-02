# 必要な機能をインポート
import PySimpleGUI as sg
from PyPDF2 import PdfFileReader, PdfFileWriter
from os import listdir
import pdfplumber
from pdf2image import convert_from_path
import os
from os.path import isfile, join

# 以下はPDF操作関数
# PDFの存在確認
def check_file_readable(filepath):
    return os.path.exists(filepath) and os.path.isfile(filepath)

# ディレクトリの書き込み権限確認
def check_dir_writable(directory):
    return os.path.exists(directory) and os.path.isdir(directory) and os.access(directory, os.W_OK)

# PDFの分割
def split_pdf(input_filepath, output_directory):
    if not check_file_readable(input_filepath):
        return "Input file does not exist or is not readable."
    if not check_dir_writable(output_directory):
        return "Output directory does not exist or is not writable."
    pdf_reader = PdfFileReader(input_filepath)
    for page_num in range(pdf_reader.getNumPages()):
        pdf_writer = PdfFileWriter()
        pdf_writer.addPage(pdf_reader.getPage(page_num))
        output_filepath = f"{output_directory}/page_{page_num + 1}.pdf"
        with open(output_filepath, 'wb') as out_pdf_file:
            pdf_writer.write(out_pdf_file)
    return None, None

# PDFの結合
def merge_pdfs(input_filepaths, output_filepath):
    for filepath in input_filepaths:
        if not check_file_readable(filepath):
            return f"Input file {filepath} does not exist or is not readable."
    if not check_dir_writable(os.path.dirname(output_filepath)):
        return "Output directory does not exist or is not writable."
    pdf_writer = PdfFileWriter()
    for input_filepath in input_filepaths:
        pdf_reader = PdfFileReader(input_filepath)
        for page_num in range(pdf_reader.getNumPages()):
            page = pdf_reader.getPage(page_num)
            pdf_writer.addPage(page)
    with open(output_filepath, 'wb') as out_pdf_file:
        pdf_writer.write(out_pdf_file)
    return None, None

# PDFの抽出
def extract_pages(input_filepath, output_filepath, page_range):
    if not check_file_readable(input_filepath):
        return "Input file does not exist or is not readable."
    if not check_dir_writable(os.path.dirname(output_filepath)):
        return "Output directory does not exist or is not writable."
    pdf_reader = PdfFileReader(input_filepath)
    pdf_writer = PdfFileWriter()
    for page_num in page_range:
        if page_num >= pdf_reader.getNumPages():
            return "Invalid page range."
        pdf_writer.addPage(pdf_reader.getPage(page_num))
    with open(output_filepath, 'wb') as out_pdf_file:
        pdf_writer.write(out_pdf_file)
    return None, None

# PDFの抽出に関わる補助関数
def parse_page_range(page_range_str):
    try:
        ranges = page_range_str.split(',')
        pages = []
        for r in ranges:
            if '-' in r:
                start, end = map(int, r.split('-'))
                if start <= 0 or end <= 0:
                    return None, "Invalid page number. Page numbers should be greater than 0."
                pages.extend(range(start - 1, end))  # 1-based to 0-based
            else:
                page = int(r)
                if page <= 0:
                    return None, "Invalid page number. Page numbers should be greater than 0."
                pages.append(page - 1)  # 1-based to 0-based
        return pages, None
    except ValueError:
        return None, "Invalid page range format. Please enter numbers and ranges separated by commas."

# PDFの全結合（ファイル内のバッチ処理）
def batch_merge_pdfs(input_directory, output_filepath):
    if not os.path.exists(input_directory):
        return "Input directory does not exist."
    if not os.path.exists(os.path.dirname(output_filepath)):
        return "Output directory does not exist."
    pdf_writer = PdfFileWriter()
    pdf_files = [f for f in os.listdir(input_directory) if os.path.isfile(os.path.join(input_directory, f)) and f.endswith('.pdf')]
    for pdf_file in pdf_files:
        pdf_reader = PdfFileReader(os.path.join(input_directory, pdf_file))
        for page_num in range(pdf_reader.getNumPages()):
            page = pdf_reader.getPage(page_num)
            pdf_writer.addPage(page)
    with open(output_filepath, 'wb') as out_pdf_file:
        pdf_writer.write(out_pdf_file)
    return None, None

# テキストの抽出
def extract_text_from_pdf(input_filepath, output_directory):
    if not os.path.exists(input_filepath):
        return "Input PDF file does not exist."
    if not os.path.exists(output_directory):
        return "Output directory does not exist."
    with pdfplumber.open(input_filepath) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            text_file_path = os.path.join(output_directory, f"page_{i + 1}_text.txt")
            with open(text_file_path, 'w', encoding='utf-8') as f:
                f.write(text)
    return None, None

# 画像の抽出(popplerを利用する)
def extract_images_from_pdf(input_filepath, output_directory):
    images = convert_from_path(input_filepath)
    for i, image in enumerate(images):
        image_path = os.path.join(output_directory, f"page_{i + 1}.png")
        image.save(image_path, 'PNG')

# 以上が関数部分
# 以下でPySimpleGUIでUIに組み込む

# GUI設計
def main():
    while True:
        # 画面0: メインメニュー
        layout0 = [
            [sg.Text('PDF操作ツール')],
            [sg.Button('PDFの分割', key='-SPLIT-'), sg.Button('PDFの結合', key='-MERGE-'),
             sg.Button('PDFの抽出', key='-EXTRACT-'), sg.Button('PDFの全結合', key='-BATCHMERGE-'),
             sg.Button('テキストと画像の抽出', key='-TEXTIMG-')],
            [sg.Button('終了', key='-EXIT-')]
        ]
        window0 = sg.Window('PDF操作ツール', layout0)
        event, values = window0.read()
        window0.close()

        if event == '-EXIT-' or event == sg.WINDOW_CLOSED:
            break

        # 各機能に共通するレイアウト要素
        common_layout = [
            [sg.Text('入力PDFファイル'), sg.InputText(), sg.FileBrowse(file_types=(("PDF Files", "*.pdf"),))],
            [sg.Text('出力ディレクトリ'), sg.InputText(), sg.FolderBrowse()]
        ]

        # 画面1: PDFの分割
        if event == '-SPLIT-':
            layout1 = common_layout + [[sg.Button('PDF分割'), sg.Button('戻る', key='-BACK-'), sg.Button('終了')]]
            window1 = sg.Window('PDFの分割', layout1)
            event, values = window1.read()
            window1.close()
            if event == 'PDF分割':
                try:
                    input_filepath = values[0]
                    output_directory = values[1]
                    result, message = split_pdf(input_filepath, output_directory)
                    if message:
                        sg.popup_error(message)  # エラーメッセージをポップアップで表示
                except Exception as e:
                    sg.popup_error(f"Unexpected error: {str(e)}")  # 予期せぬエラーをポップアップで表示
            elif event == '-BACK-':  # 戻るボタンが押されたとき
                continue  # メインのwhileループに戻る


        # 画面2: PDFの結合
        if event == '-MERGE-':
            layout2 = [
                [sg.Text('入力PDFファイル1'), sg.InputText(key='-IN1-'), sg.FileBrowse(file_types=(("PDF Files", "*.pdf"),))],
                [sg.Text('入力PDFファイル2'), sg.InputText(key='-IN2-'), sg.FileBrowse(file_types=(("PDF Files", "*.pdf"),))],
                # さらにファイルを追加する場合は、ここ行を追加
                [sg.Text('出力PDFファイル名'), sg.InputText(key='-OUT-'), sg.SaveAs(file_types=(("PDF Files", "*.pdf"),))],
                [sg.Button('PDF結合'), sg.Button('戻る', key='-BACK-'), sg.Button('終了')]
            ]
            window2 = sg.Window('PDFの結合', layout2)
            event, values = window2.read()
            window2.close()
            if event == 'PDF結合':
                try:
                    input_filepaths = [values['-IN1-'], values['-IN2-']]
                    output_filepath = values['-OUT-']
                    result, message = merge_pdfs(input_filepaths, output_filepath)
                    if message:
                        sg.popup_error(message)
                except Exception as e:
                    sg.popup_error(f"Unexpected error: {str(e)}")  # 予期せぬエラーをポップアップで表示
            elif event == '-BACK-':  # 戻るボタンが押されたとき
                continue  # メインのwhileループに戻る

        # 画面3: PDFの抽出
        if event == '-EXTRACT-':
            layout3 = [
                [sg.Text('入力PDFファイル'), sg.InputText(), sg.FileBrowse(file_types=(("PDF Files", "*.pdf"),))],
                [sg.Text('出力PDFファイル名'), sg.InputText(), sg.SaveAs(file_types=(("PDF Files", "*.pdf"),))],
                [sg.Text('抽出するページ範囲（例: 1-20,40-50,71）'), sg.InputText()],
                [sg.Button('PDF抽出'), sg.Button('戻る', key='-BACK-'), sg.Button('終了')]
            ]
            window3 = sg.Window('PDFの抽出', layout3)
            event, values = window3.read()
            window3.close()
            if event == 'PDF抽出':
                try:
                    input_filepath = values[0]
                    output_filepath = values[1]
                    page_range_str = values[2]
                    page_range, message = parse_page_range(page_range_str)
                    if message:
                        sg.popup_error(message)
                    else:
                        result, message = extract_pages(input_filepath, output_filepath, page_range)
                        if message:
                            sg.popup_error(message)
                except Exception as e:
                    sg.popup_error(f"Unexpected error: {str(e)}")  # 予期せぬエラーをポップアップで表示
            elif event == '-BACK-':  # 戻るボタンが押されたとき
                continue  # メインのwhileループに戻る


        # 画面4: PDFの全結合（バッチ処理）
        if event == '-BATCHMERGE-':
            layout4 = [
                [sg.Text('入力PDFファイル格納ディレクトリ'), sg.InputText(), sg.FolderBrowse()],
                [sg.Text('出力PDFファイル名'), sg.InputText(), sg.SaveAs(file_types=(("PDF Files", "*.pdf"),))],
                [sg.Button('全PDF結合'), sg.Button('戻る', key='-BACK-'), sg.Button('終了')]
            ]
            window4 = sg.Window('PDFの全結合', layout4)
            event, values = window4.read()
            window4.close()
            if event == '全PDF結合':
                try:
                    input_directory = values[0]
                    output_filepath = values[1]
                    result, message = batch_merge_pdfs(input_directory, output_filepath)
                    if message:
                        sg.popup_error(message)
                except Exception as e:
                    sg.popup_error(f"Unexpected error: {str(e)}")  # 予期せぬエラーをポップアップで表示
            elif event == '-BACK-':  # 戻るボタンが押されたとき
                continue  # メインのwhileループに戻る


        # 画面5: テキストと画像の抽出
        if event == '-TEXTIMG-':
            layout5 = common_layout + [
                [sg.Radio('テキストのみ抽出', "RADIO1", default=True, key='-TEXT-'),
                sg.Radio('画像のみ抽出', "RADIO1", key='-IMG-'),
                sg.Radio('テキストと画像を抽出', "RADIO1", key='-BOTH-')],
                [sg.Button('抽出実行'), sg.Button('戻る', key='-BACK-'), sg.Button('終了')]
            ]
            window5 = sg.Window('テキストと画像の抽出', layout5)
            event, values = window5.read()
            window5.close()
            if event == '抽出実行':
                try:
                    input_filepath = values[0]
                    output_directory = values[1]
                    if values['-TEXT-']:
                        result, message = extract_text_from_pdf(input_filepath, output_directory)
                    elif values['-IMG-']:
                        result, message = extract_images_from_pdf(input_filepath, output_directory)
                    elif values['-BOTH-']:
                        result1, message1 = extract_text_from_pdf(input_filepath, output_directory)
                        result2, message2 = extract_images_from_pdf(input_filepath, output_directory)
                        message = message1 or message2
                    if message:
                        sg.popup_error(message)
                except Exception as e:
                    sg.popup_error(f"Unexpected error: {str(e)}")  # 予期せぬエラーをポップアップで表示
            elif event == '-BACK-':  # 戻るボタンが押されたとき
                continue  # メインのwhileループに戻る

if __name__ == "__main__":
    main()
