"""
現実的なダミーデータ生成スクリプト
2022年1月1日〜2025年12月31日の約40,000レコード（4年分・年間1万レコード）
"""

import csv
import random
from datetime import datetime, timedelta

# 基本設定
START_DATE = datetime(2022, 1, 1)
END_DATE = datetime(2025, 12, 31)
REGIONS = ['北海道', '東北', '関東', '中部', '近畿', '中国', '四国', '九州']
PERSONS = ['田中太郎', '佐藤花子', '鈴木一郎', '高橋美咲', '伊藤健太', '渡辺さくら', '山本大輔', '中村愛', '小林直樹', '加藤美穂']
PRODUCTS = ['製品A', '製品B', '製品C', '製品D', '製品E', 'サービスX', 'サービスY']
CATEGORIES = ['サブスク', '単発', '追加オプション', '保守', 'その他']

# 地域の特性（人口比率・経済規模）
REGION_WEIGHTS = {
    '北海道': 0.08,
    '東北': 0.10,
    '関東': 0.35,  # 最大
    '中部': 0.15,
    '近畿': 0.20,
    '中国': 0.05,
    '四国': 0.03,
    '九州': 0.04
}

# 担当者の特性（能力・経験値）
PERSON_PERFORMANCE = {
    '田中太郎': 1.3,   # エース
    '佐藤花子': 1.2,
    '鈴木一郎': 1.0,
    '高橋美咲': 1.1,
    '伊藤健太': 0.9,
    '渡辺さくら': 1.15,
    '山本大輔': 0.85,
    '中村愛': 1.05,
    '小林直樹': 1.25,
    '加藤美穂': 0.95
}

# 商品の特性（価格帯）
PRODUCT_PRICES = {
    '製品A': (30000, 50000),
    '製品B': (50000, 80000),
    '製品C': (20000, 40000),
    '製品D': (60000, 100000),
    '製品E': (15000, 30000),
    'サービスX': (100000, 200000),
    'サービスY': (80000, 150000)
}

# カテゴリと商品の関係
PRODUCT_CATEGORIES = {
    '製品A': 'サブスク',
    '製品B': '単発',
    '製品C': 'サブスク',
    '製品D': '単発',
    '製品E': '追加オプション',
    'サービスX': '保守',
    'サービスY': '保守'
}


def get_seasonal_factor(date):
    """季節性係数を取得"""
    month = date.month

    # 年末年始ブースト
    if month in [12, 1]:
        return 1.4
    # ゴールデンウィーク前の駆け込み
    elif month == 3:
        return 1.3
    # 夏季（閑散期）
    elif month in [7, 8]:
        return 0.8
    # 通常期
    else:
        return 1.0


def get_weekday_factor(date):
    """曜日係数を取得"""
    weekday = date.weekday()

    # 月曜日: 週初めで活発
    if weekday == 0:
        return 1.1
    # 火〜木: 通常
    elif weekday in [1, 2, 3]:
        return 1.0
    # 金曜日: 週末前で活発
    elif weekday == 4:
        return 1.15
    # 土日: 低調
    else:
        return 0.6


def get_trend_factor(date):
    """トレンド係数を取得（年間成長）"""
    days_from_start = (date - START_DATE).days
    total_days = (END_DATE - START_DATE).days

    # 年間で20%成長
    growth_rate = 0.20
    return 1.0 + (growth_rate * days_from_start / total_days)


def generate_sales_data():
    """売上データを生成"""
    data = []
    current_date = START_DATE

    while current_date <= END_DATE:
        # 1日あたり27〜30件の取引（4年で約40,000レコード、年間約1万）
        num_transactions = random.randint(27, 30)

        for _ in range(num_transactions):
            # 地域を選択（重み付き）
            region = random.choices(REGIONS, weights=list(REGION_WEIGHTS.values()))[0]

            # 担当者を選択
            person = random.choice(PERSONS)
            person_factor = PERSON_PERFORMANCE[person]

            # 商品を選択
            product = random.choice(PRODUCTS)
            price_range = PRODUCT_PRICES[product]
            base_price = random.randint(price_range[0], price_range[1])

            # カテゴリを取得
            category = PRODUCT_CATEGORIES[product]

            # 各種係数を適用
            seasonal = get_seasonal_factor(current_date)
            weekday = get_weekday_factor(current_date)
            trend = get_trend_factor(current_date)

            # 最終的な売上金額
            final_price = int(base_price * seasonal * weekday * trend * person_factor)

            # 数量（1〜5個）
            quantity = random.randint(1, 5)
            total_sales = final_price * quantity

            data.append({
                'Date': current_date.strftime('%Y-%m-%d'),
                'Region': region,
                'Person': person,
                'Product': product,
                'Category': category,
                'Quantity': quantity,
                'UnitPrice': final_price,
                'TotalSales': total_sales,
                'DayOfWeek': current_date.strftime('%A'),
                'Month': current_date.month,
                'Quarter': (current_date.month - 1) // 3 + 1
            })

        current_date += timedelta(days=1)

    return data


def save_to_csv(data, filename):
    """CSVファイルに保存"""
    if not data:
        return

    keys = data[0].keys()
    with open(filename, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=keys)
        writer.writeheader()
        writer.writerows(data)

    print(f"✅ {filename} に {len(data)} 件のレコードを保存しました")


def print_summary(data):
    """データのサマリーを表示"""
    total_sales = sum(row['TotalSales'] for row in data)
    total_records = len(data)

    print("\n" + "="*50)
    print("📊 データサマリー")
    print("="*50)
    print(f"総レコード数: {total_records:,} 件")
    print(f"総売上金額: ¥{total_sales:,}")
    print(f"平均単価: ¥{total_sales // total_records:,}")
    print(f"期間: {START_DATE.strftime('%Y-%m-%d')} 〜 {END_DATE.strftime('%Y-%m-%d')}")
    print("="*50 + "\n")


if __name__ == '__main__':
    import os
    from pathlib import Path

    print("🚀 ダミーデータ生成を開始します...\n")

    # データ生成
    sales_data = generate_sales_data()

    # サマリー表示
    print_summary(sales_data)

    # ダウンロードフォルダのパスを取得
    downloads_folder = str(Path.home() / "Downloads")
    output_path = os.path.join(downloads_folder, 'sales_data_2022-2025.csv')

    # CSV保存
    save_to_csv(sales_data, output_path)

    print("✅ 完了しました！")
    print(f"\n📁 保存先: {output_path}")
    print("\n次のステップ:")
    print("1. Googleスプレッドシートの「ダミーデータ作成」ダイアログを開く")
    print("2. ダウンロードフォルダから sales_data_2022-2025.csv を選択してインポート")
    print("3. 既存データは自動的に削除され、新しいデータで置き換わります")
    print("4. 集計シートとグラフが自動生成されます")
