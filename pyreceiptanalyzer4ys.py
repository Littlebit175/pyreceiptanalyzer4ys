# 領収書アナライザY
# 機能
# - ヤフショ領収書PDFファイルを読み込み、抽出したデータをCSVファイルに保存する
# - CSVを"購入日"・"注文番号"の順で昇順にソートする
# - "input"フォルダに保存された複数のPDFファイルを順次読み込む
# - "output"フォルダにリネームしたPDFファイルを保存する
#   yyyy-mm-dd_領収書_仕入_<店舗名>_<決済1><決済額1>円_<決済2><決済額2>円_<決済3><決済額3>円_<注文番号>.pdf
# - ERROR発生時にtext化したPDFファイルをerror_YYYYMMDD_hhmm.txtに保存する
# CSVに保存するデータ：
# - １列目（注文番号）："注文番号XXXXX-1234567の領収書"から"XXXXX-1234567"を抽出
# - ２列目（モール名）：一律で"ヤフショ"と入力
# - ３列目（店舗名）：注文番号の次の行の"様"から改行までを店舗名として抽出
# - ４列目（購入日）："注文日: yyyy年mm月dd日"からyyyy/mm/dd形式で抽出
# - ５列目（購入月）：購入日から"yyyy/mm"形式で抽出
# - ６列目（支払金額）："\n合計金額(税込) 40,000円\n"から"40000"（例）を抽出
#   支払い内訳データ（例１）："\n支払い内訳\nPayPay（残高） 40,000円\n税率別内訳 税込金額 消費税額"
#   支払い内訳データ（例２）："\n支払い内訳\nPayPay（残高） 39,000円\n商品券 1,000円\n税率別内訳 税込金額 消費税額"
# - ７列目（決済1）：支払い内訳データから"PayPay（残高）"（例）を抽出
# - ８列目（決済額１）：決済１の金額を抽出
# - ９列目（決済２）：支払い内訳データから"商品券"（例）を抽出（データがない場合は空のデータを入力）
# - １０列目（決済額２）：決済２の金額を抽出（データがない場合は空のデータを入力）
# - １１列目（決済３）：支払い内訳データから決済方法を抽出（データがない場合は空のデータを入力）
# - １２列目（決済額３）：決済３の金額を抽出（データがない場合は空のデータを入力）
# - １３列目（注文商品）："\n注文商品 価格\n"から"\n単価(税込)"までを注文商品として抽出
# - １４列目（注文商品）：Inputフォルダから読み取ったリネーム前のPDFファイル名
# - １５列目（注文商品）：Outputフォルダに保存したリネーム後のPDFファイル名

import os
import unicodedata
import shutil
import pypdf
import pandas as pd

def extract_text_from_pdf(pdf_path):
    with open(pdf_path, 'rb') as file:
        reader = pypdf.PdfReader(file)
        text = ''
        for page in reader.pages:
            text += page.extract_text() + '\n'  # ページからテキストを抽出して追加
            
    # テキストの正規化
    # 領収書によって異なる文字コードが使用されているため
    # 例：「文」Unicodeの「U+6587」（CJK統合漢字）と「⽂」Unicodeの「U+2F8A」（CJK互換漢字）
    unicode_text = unicodedata.normalize('NFKC', text)
    
    return unicode_text

def initialize_data_structure():
    return {
        '注文番号': [],
        'モール名': [],
        '店舗名': [],
        '購入日': [],
        '購入月': [],
        '支払金額': [],
        '決済1': [],
        '決済額1': [],
        '決済2': [],
        '決済額2': [],
        '決済3': [],
        '決済額3': [],
        '注文商品': [],
        'リネーム前ファイル名': [],
        'リネーム後ファイル名': []
    }


def parse_pdf_text(text):
    isError = False
    
    # CSVに保存するデータを初期化
    data = initialize_data_structure()
    
    # "注文番号XXXXX-1234567の領収書"から"XXXXX-1234567"を抽出
    if '注文番号' in text and 'の領収書' in text:
        order_number = text.split('注文番号')[1].split('の領収書')[0]
    else:
        order_number = 'ERROR'
        isError = True

    data['注文番号'].append(order_number)

    # モール名をCSVに保存
    data['モール名'].append('ヤフショ')
    
    # "注文番号XXXXX-1234567の領収書"の次の行の"様"から改行までを店舗名として抽出
    if 'の領収書' in text and '様' in text:
        store_name = text.split('の領収書')[1].split('様')[1].split('\n')[0]
    else:
        store_name = 'ERROR'
        isError = True

    data['店舗名'].append(store_name)

    # "注文日: yyyy年mm月dd日\n"からyyyy/mm/dd形式で抽出
    if '注文日: ' in text:
        purchase_date = text.split('注文日: ')[1].split('\n')[0].replace('年', '/').replace('月', '/').replace('日', '')
        purchase_month = purchase_date.split('/')[0] + '/' + purchase_date.split('/')[1]    # 購入日から"yyyy/mm"形式で抽出
    else:
        purchase_date = 'ERROR'
        purchase_month = 'ERROR'
        isError = True
      
    data['購入日'].append(purchase_date)
    data['購入月'].append(purchase_month)

    # "\n合計金額(税込) 40,000円\n"から"40000"（例）を抽出
    if '\n合計金額(税込) ' in text and '円' in text:
        total_price = text.split('\n合計金額(税込) ')[1].split('円')[0].replace(',', '')
    elif '\n合計金額( 税込) ' in text and '円' in text:
        total_price = text.split('\n合計金額( 税込) ')[1].split('円')[0].replace(',', '')
    else:
        total_price = 'ERROR'
        isError = True

    data['支払金額'].append(total_price)

    # 決済方法の種類数を、"\n支払い内訳\n"から"税率別内訳"までの文字列の中に含まれる"\n"の数で判断
    if '\n支払い内訳\n' in text and '税率別内訳' in text:
        payment_detail_count = text.split('\n支払い内訳\n')[1].split('税率別内訳')[0].count('\n')
        # 決済方法の種類数最大値の３回ループ
        for i in range(3):
            if i < payment_detail_count:   # 決済データがある場合
                # 決済{i+1}を抽出
                payment_detail1 = text.split('\n支払い内訳\n')[1].split('税率別内訳')[0].split('\n')[i].split(' ')[0]
                # 決済方法が"商品券"の場合、"ヤフショ商品券"に書き換え
                if payment_detail1 == '商品券':
                    payment_detail1 = 'ヤフショ商品券'
                data['決済' + str(i+1)].append(payment_detail1)
                # 決済{i+1}の金額を抽出
                payment_detail1_price = text.split('\n支払い内訳\n')[1].split('税率別内訳')[0].split('\n')[i].split(' ')[1].replace('円', '').replace(',', '')
                data['決済額' + str(i+1)].append(payment_detail1_price)
            else:   # 決済データがない場合、空のデータを入力
                data['決済' + str(i+1)].append('')
                data['決済額' + str(i+1)].append('')
    else:
        data['決済1'].append('ERROR')
        data['決済額1'].append('ERROR')
        data['決済2'].append('ERROR')
        data['決済額2'].append('ERROR')
        data['決済3'].append('ERROR')
        data['決済額3'].append('ERROR')
        isError = True

    # "\n注文商品 価格\n"から次の改行までを注文商品として抽出
    if '\n注文商品 価格\n' in text and '\n単価(税込)' in text:
        order_items = text.split('\n注文商品 価格\n')[1].split('\n単価(税込)')[0]
    else:
        order_items = 'ERROR'
        isError = True

    data['注文商品'].append(order_items)
    
    return data, isError


# リネーム後ファイル名を生成（決済2、決済3が空の場合は空白にする）
def generate_new_pdf_file_name(data):
    if 'ERROR' in data['注文番号'][-1] or 'ERROR' in data['店舗名'][-1] or 'ERROR' in data['購入日'][-1] or 'ERROR' in data['支払金額'][-1]:
        return 'ERROR'
    
    new_pdf_file = data['購入日'][-1].replace('/', '-') + '_領収書_仕入_' + data['店舗名'][-1] + '_' + data['決済1'][-1] + data['決済額1'][-1] + '円'

    if data['決済2'][-1] != '':
        new_pdf_file += '_' + data['決済2'][-1] + data['決済額2'][-1] + '円'
    if data['決済3'][-1] != '':
        new_pdf_file += '_' + data['決済3'][-1] + data['決済額3'][-1] + '円'
        
    new_pdf_file += '_' + data['注文番号'][-1] + '.pdf'
    
    return new_pdf_file

def sort_data(data):
    # データを"購入日"・"注文番号"の順でソート
    df = pd.DataFrame(data)
    df.sort_values(by=['購入日', '注文番号'], inplace=True)
    return df.to_dict(orient='list')

def save_to_csv(data, output_file):
    df = pd.DataFrame(data)
    df.to_csv(output_file, index=False, encoding='utf-8')

def main():
    input_folder = 'input'
    pdf_text_with_error = ''
    
    # フォルダ内のPDFファイルを取得
    pdf_files = [file for file in os.listdir(input_folder) if file.endswith('.pdf')]

    # CSVに保存するデータを初期化
    all_data = initialize_data_structure()

    # 各PDFファイルを順次読み込んでデータを抽出
    for pdf_file in pdf_files:
        pdf_path = os.path.join(input_folder, pdf_file) # PDFファイルのパスを取得
        text = extract_text_from_pdf(pdf_path)  # PDFファイルからテキストを抽出
        data, isError = parse_pdf_text(text) # テキストから必要なデータを抽出
        if isError:
            # ERROR情報を順次追加
            pdf_text_with_error += pdf_file + '\n' + text + '\n\n'
        data['リネーム前ファイル名'].append(pdf_file)   # リネーム前ファイル名をCSVに保存
        new_pdf_file = generate_new_pdf_file_name(data) # リネーム後ファイル名を生成
        data['リネーム後ファイル名'].append(new_pdf_file)   # リネーム後ファイル名をCSVに保存
        
        # outputフォルダにリネームしたPDFファイルを保存
        output_folder = 'output'
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        # outputフォルダにリネームしたPDFファイルをコピーして保存
        if new_pdf_file != 'ERROR':
            shutil.copyfile(pdf_path, os.path.join(output_folder, new_pdf_file))
        
        for key in all_data:
            all_data[key].extend(data[key])

    # データをソートする
    sorted_data = sort_data(all_data)

    timestamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M')
    # 現在日時を語尾につけてCSVファイルに保存(list_YYYYMMDD_hhmm.csv)
    output_file = 'list_' + timestamp + '.csv'
    save_to_csv(sorted_data, output_file)
    # ERROR発生時にtext化したPDFファイルをerror_YYYYMMDD_hhmm.txtに保存する
    if pdf_text_with_error != '':
        error_file = 'error_' + timestamp + '.txt'
        with open(error_file, 'w', encoding='utf-8') as f:
            f.write(pdf_text_with_error)  
            
    # Enterを押すとコンソールを閉じる
    input('処理が完了しました。Enterを押してください。')

if __name__ == "__main__":
    main()