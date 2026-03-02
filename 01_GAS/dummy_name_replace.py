import csv
import re
from datetime import datetime

# historicalNames.js から名前リストを抽出
def extract_names(js_path):
    with open(js_path, encoding='utf-8') as f:
        content = f.read()
    # 子供用の名前リスト CHILDREN_NAMES を抽出
    match = re.search(r'const\s+CHILDREN_NAMES\s*=\s*\[([\s\S]*?)\];', content)
    if not match:
        raise ValueError("CHILDREN_NAMES 配列が見つかりません")
    names = re.findall(r'"([^"]+)"', match.group(1))
    return names

# 日付パターン（JavaScript版を参考）
# YYYY.MM.DD, YYYY/MM/DD, YYYY-MM-DD, YYYY年MM月DD日、和暦など多数のフォーマットに対応
DATE_PATTERN = r'((?:19|20)\d{2}[\.\\/\-]\d{1,2}[\.\\/\-]\d{1,2}|(?:19|20)\d{2}年\d{1,2}月\d{1,2}日?|(?:明治|大正|昭和|平成|令和)\d{1,2}年\d{1,2}月\d{1,2}日?|[MTSHRmtshr]\d{1,2}[\.\\/]\d{1,2}[\.\\/]\d{1,2}(?:生)?|(?:19|20)\d{6})'

def parse_family_info(raw_text):
    """
    コード.jsの parseFamilyInfo 関数の動作を参考に
    複数人の家族情報を解析して個別の情報に分割
    1行に複数の日付がある場合にも対応
    関係性テキスト（第1子、夫など）で分割
    """
    if not raw_text:
        return []
    
    results = []
    
    # 関係性を示すテキストで事前に分割
    # 「第1子：」「夫　：」などのパターン
    relation_pattern = r'(第\d+子|夫|妻|父|母|兄|姉|弟|妹|祖父|祖母|叔父|叔母|従兄弟|従姉妹)[\s　]*[:：]'
    
    # 関係性で分割
    sections = re.split(relation_pattern, raw_text)
    
    # 関係性テキストが見つかった場合
    if len(sections) > 1:
        # sections の構造：[前置き, 関係性1, テキスト1, 関係性2, テキスト2, ...]
        i = 1  # インデックス1から開始（0は前置き）
        while i < len(sections):
            if i + 1 < len(sections):
                relation = sections[i]  # 第1子、夫など
                text_block = sections[i + 1].strip()
            else:
                break
            
            # テキストブロック内の複数行を処理
            lines = text_block.split('\n')
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                
                # 括弧を処理
                if '（' in line and '）' in line:
                    line = line.replace('（', '(').replace('）', ')')
                
                # 日付をすべて検索
                date_matches = list(re.finditer(DATE_PATTERN, line))
                
                if date_matches:
                    # 複数の日付がある場合、各日付の直前のテキストを名前とする
                    for j, date_match in enumerate(date_matches):
                        date_str = date_match.group(0)
                        date_start = date_match.start()
                        date_end = date_match.end()
                        
                        # 前の日付の終端、または行の開始から現在の日付の開始までが名前
                        if j == 0:
                            pre_start = 0
                        else:
                            pre_start = date_matches[j-1].end()
                        
                        pre = line[pre_start:date_start].strip()
                        
                        # 括弧内の内容を削除
                        pre = re.sub(r'\([^)]*\)', '', pre).strip()
                        
                        # 不要なプレフィックスを削除
                        pre = re.sub(r'^[^：]*[:：]\s*', '', pre).strip()
                        
                        # 中点処理：最後の中点より後ろが名前
                        if '・' in pre:
                            parts = pre.split('・')
                            pre = parts[-1].strip()
                        
                        # 次の日付までの情報、または行の最後までの情報
                        if j < len(date_matches) - 1:
                            post_end = date_matches[j+1].start()
                        else:
                            post_end = len(line)
                        
                        post = line[date_end:post_end].strip()
                        
                        # 中点と括弧を削除
                        if post.startswith('・'):
                            post = post[1:].strip()
                        post = re.sub(r'\([^)]*\)', '', post).strip()
                        
                        if pre or date_str:
                            individual = {
                                'name': pre,
                                'dob': date_str,
                                'info': post
                            }
                            results.append(individual)
                else:
                    # 日付がない行でも、短いテキストは名前の可能性
                    if len(line) < 20 and not any(kw in line for kw in ["職業", "アレルギー", "園", "学校"]):
                        if line and results:
                            # 前の個人に情報として追加
                            if results[-1]['info']:
                                results[-1]['info'] += ' ' + line
                            else:
                                results[-1]['info'] = line
            
            i += 2  # 次の関係性へ
    else:
        # 関係性テキストがない場合のフォールバック処理
        # 元のテキストを改行で分割して処理
        lines = raw_text.split('\n')
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            # 括弧を処理
            if '（' in line and '）' in line:
                line = line.replace('（', '(').replace('）', ')')
            
            # 日付をすべて検索
            date_matches = list(re.finditer(DATE_PATTERN, line))
            
            if date_matches:
                # 複数の日付がある場合、各日付の直前のテキストを名前とする
                for j, date_match in enumerate(date_matches):
                    date_str = date_match.group(0)
                    date_start = date_match.start()
                    date_end = date_match.end()
                    
                    # 前の日付の終端、または行の開始から現在の日付の開始までが名前
                    if j == 0:
                        pre_start = 0
                    else:
                        pre_start = date_matches[j-1].end()
                    
                    pre = line[pre_start:date_start].strip()
                    
                    # 括弧内の内容を削除
                    pre = re.sub(r'\([^)]*\)', '', pre).strip()
                    
                    # 不要なプレフィックスを削除
                    pre = re.sub(r'^[^：]*[:：]\s*', '', pre).strip()
                    
                    # 中点処理：最後の中点より後ろが名前
                    if '・' in pre:
                        parts = pre.split('・')
                        pre = parts[-1].strip()
                    
                    # 次の日付までの情報、または行の最後までの情報
                    if j < len(date_matches) - 1:
                        post_end = date_matches[j+1].start()
                    else:
                        post_end = len(line)
                    
                    post = line[date_end:post_end].strip()
                    
                    # 中点と括弧を削除
                    if post.startswith('・'):
                        post = post[1:].strip()
                    post = re.sub(r'\([^)]*\)', '', post).strip()
                    
                    if pre or date_str:
                        individual = {
                            'name': pre,
                            'dob': date_str,
                            'info': post
                        }
                        results.append(individual)
            else:
                # 日付がない行
                # 情報キーワード
                info_keywords = [
                    "職業", "勤務", "園", "学校", "社", "アレルギー", "疾患", "病", 
                    "薬", "申請", "検討", "利用", "金額", "備考", "共有", "男児", "女児"
                ]
                is_info_keyword = any(kw in line for kw in info_keywords)
                
                # 名前として扱うかを判定（情報キーワードがなく、短い行）
                if not is_info_keyword and len(line) < 20 and line:
                    individual = {
                        'name': line,
                        'dob': '',
                        'info': ''
                    }
                    results.append(individual)
                elif results:
                    # 前の個人に情報として追加
                    if results[-1]['info']:
                        results[-1]['info'] += ' ' + line
                    else:
                        results[-1]['info'] = line
    
    return results

def replace_names_in_csv(csv_path, js_path, output_path, target_col_name):
    names = extract_names(js_path)
    name_idx = 0
    
    # エンコーディングを自動検出
    encodings = ['utf-8', 'utf-8-sig', 'utf-16', 'utf-16-le', 'utf-16-be', 'shift-jis', 'cp932']
    input_encoding = 'utf-8'
    for enc in encodings:
        try:
            with open(csv_path, encoding=enc) as f:
                f.read()
            input_encoding = enc
            break
        except (UnicodeDecodeError, UnicodeError):
            continue
    
    print(f"検出されたエンコーディング: {input_encoding}")

    with open(csv_path, encoding=input_encoding) as infile, \
         open(output_path, 'w', encoding=input_encoding, newline='') as outfile:
        reader = csv.DictReader(infile, delimiter='\t')
        fieldnames = reader.fieldnames
        
        # デバッグ: カラム名を表示
        print(f"CSVのカラム数: {len(fieldnames)}")
        
        # カラム名を検索（部分一致）
        matched_col = None
        for col in fieldnames:
            if '世帯全員の情報' in col:
                matched_col = col
                print(f"使用するカラム: {repr(matched_col)}")
                break
        
        if not matched_col:
            raise ValueError(f"カラン '{target_col_name}' が見つかりません")
        
        writer = csv.DictWriter(outfile, fieldnames=fieldnames, delimiter='\t')
        writer.writeheader()

        row_count = 0
        for row in reader:
            family_raw = row[matched_col] or ""
            
            # 家族情報をパース
            individuals = parse_family_info(family_raw)
            
            # 各人物の名前をダミー名で置換
            replaced_parts = []
            for individual in individuals:
                original_name = individual.get('name', '')
                dummy_name = names[name_idx % len(names)]
                name_idx += 1
                
                # 置換後の個人情報を再構築
                parts = [dummy_name]
                if individual.get('dob'):
                    parts.append(individual['dob'])
                if individual.get('info'):
                    parts.append(individual['info'])
                
                replaced_parts.append(' '.join(parts))
                print(f"  置換: {original_name} → {dummy_name}")
            
            # 改行で結合
            row[matched_col] = '\n'.join(replaced_parts)
            writer.writerow(row)
            row_count += 1
        
        print(f"\n処理完了: {row_count}件のレコード")

if __name__ == "__main__":
    replace_names_in_csv(
        'Kokyaku_202601191958_1.csv',
        'gas-childcare-report/historicalNames.js',
        'Kokyaku_202601191958_1_dummy.csv',
        '世帯全員の情報（名前・生年月日・職業/所属・アレルギー・その他共有事項）'
    )

