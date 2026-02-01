import pandas as pd
import numpy as np
from datetime import datetime, timedelta

def generate_dummy_csv(filename="dummy_sales.csv"):
    # Generate data for 2 years (2024-2025)
    start_date = datetime(2024, 1, 1)
    days = 730
    dates = [start_date + timedelta(days=i) for i in range(days)]
    
    # Define Brand Strategy
    brands = {
        "Calbee": ["ポテトチップス", "じゃがりこ", "かっぱえびせん", "フルグラ"],
        "Lay's": ["クラシック", "サワークリーム", "ソルト＆ビネガー"],
        "Haitai": ["ハニーバターチップ", "ホームランボール", "エースクラッカー"],
        "Koikeya": ["カラムーチョ", "プライドポテト", "ドンタコス"],
        "Pringles": ["サワークリームオニオン", "オリジナル", "チーズ"]
    }
    
    data_rows = []
    
    for date in dates:
        # Generate 3-10 transactions per day
        num_transactions = np.random.randint(3, 10)
        for _ in range(num_transactions):
            brand = np.random.choice(list(brands.keys()))
            product = np.random.choice(brands[brand])
            
            # Seasonal trends (Summer peak for snacks)
            month = date.month
            seasonality = 1.0 + (0.3 if month in [7, 8] else 0.0)
            
            sales = int(np.random.randint(200, 1500) * seasonality)
            quantity = np.random.randint(1, 10)
            
            data_rows.append({
                "日付": date,
                "ブランド": brand,
                "商品名": product,
                "売上": sales * quantity,
                "数量": quantity,
                "カテゴリ": "スナック菓子"
            })
            
    df = pd.DataFrame(data_rows)
    df.to_csv(filename, index=False, encoding='utf-8-sig')
    print(f"Generated {filename} with {len(df)} records.")

if __name__ == "__main__":
    generate_dummy_csv()
