import streamlit as st
import pandas as pd
from io import BytesIO
from allpairspy import AllPairs

st.set_page_config(page_title="Pairwise比較法エクセル生成", layout="wide")

st.title("📊 Pairwise比較法 エクセルファイル生成")
st.markdown("**AllPairs法を使用した効率的なテストケース生成**")

# セッション状態の初期化
if 'factors' not in st.session_state:
    st.session_state.factors = [
        {'name': '主食', 'values': ['米', 'パン', 'ナン']},
        {'name': '副食', 'values': ['肉', '魚', 'たこ焼き']},
        {'name': 'デザート', 'values': ['プリン', 'ゼリー', 'ケーキ']}
    ]

def add_factor():
    """新しい因子を追加"""
    factor_num = len(st.session_state.factors) + 1
    st.session_state.factors.append({
        'name': f'因子{factor_num}',
        'values': [f'値{factor_num}-1', f'値{factor_num}-2']
    })

def remove_factor(index):
    """因子を削除"""
    if len(st.session_state.factors) > 1:
        st.session_state.factors.pop(index)

def add_value(factor_index):
    """因子に新しい値を追加"""
    value_num = len(st.session_state.factors[factor_index]['values']) + 1
    factor_num = factor_index + 1
    st.session_state.factors[factor_index]['values'].append(f'値{factor_num}-{value_num}')

def remove_value(factor_index, value_index):
    """因子の値を削除"""
    if len(st.session_state.factors[factor_index]['values']) > 2:
        st.session_state.factors[factor_index]['values'].pop(value_index)

def validate_factors():
    """因子のバリデーション"""
    if len(st.session_state.factors) < 2:
        return False, "因子を2つ以上追加してください。"
    
    for factor in st.session_state.factors:
        if len(factor['values']) < 2:
            return False, f"因子「{factor['name']}」には2つ以上の値が必要です。"
    
    return True, ""

def generate_pairwise_excel():
    """AllPairs法を使用してPairwise比較用のエクセルファイルを生成"""
    # バリデーション
    is_valid, error_msg = validate_factors()
    if not is_valid:
        raise ValueError(error_msg)
    
    output = BytesIO()
    
    # 因子の値リストを準備
    factor_values = [factor['values'] for factor in st.session_state.factors]
    factor_names = [factor['name'] for factor in st.session_state.factors]
    
    # AllPairs法でテストケースを生成
    pairs_list = []
    for pairs in AllPairs(factor_values):
        pairs_list.append(list(pairs))
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # メインシート: AllPairs法で生成されたテストケース
        df_main = pd.DataFrame(pairs_list, columns=factor_names)
        df_main.insert(0, 'テストケースNo', range(1, len(pairs_list) + 1))
        df_main['結果'] = ''
        df_main['備考'] = ''
        df_main.to_excel(writer, sheet_name='Pairwiseテストケース', index=False)
        
        # 各因子の組み合わせマトリックスシートを作成
        for i in range(len(st.session_state.factors)):
            for j in range(i + 1, len(st.session_state.factors)):
                factor1 = st.session_state.factors[i]
                factor2 = st.session_state.factors[j]
                
                # 2因子間の組み合わせを抽出
                combinations = []
                for pair in pairs_list:
                    combinations.append({
                        factor1['name']: pair[i],
                        factor2['name']: pair[j]
                    })
                
                df_combo = pd.DataFrame(combinations)
                df_combo = df_combo.drop_duplicates()
                df_combo['確認'] = ''
                
                # シート名を作成（エクセルのシート名制限に対応）
                sheet_name = f"{factor1['name'][:12]}×{factor2['name'][:12]}"
                df_combo.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # カバレッジマトリックスシートを作成
        coverage_data = []
        for i in range(len(st.session_state.factors)):
            for j in range(i + 1, len(st.session_state.factors)):
                factor1 = st.session_state.factors[i]
                factor2 = st.session_state.factors[j]
                
                # 理論上の全組み合わせ数
                total_combinations = len(factor1['values']) * len(factor2['values'])
                
                # 実際にカバーされた組み合わせ数
                covered = set()
                for pair in pairs_list:
                    covered.add((pair[i], pair[j]))
                
                coverage_data.append({
                    '因子1': factor1['name'],
                    '因子2': factor2['name'],
                    '因子1の値数': len(factor1['values']),
                    '因子2の値数': len(factor2['values']),
                    '全組み合わせ数': total_combinations,
                    'カバー数': len(covered),
                    'カバレッジ率': f"{len(covered) / total_combinations * 100:.1f}%"
                })
        
        df_coverage = pd.DataFrame(coverage_data)
        df_coverage.to_excel(writer, sheet_name='カバレッジ分析', index=False)
        
        # サマリーシートを作成
        summary_data = []
        for factor in st.session_state.factors:
            summary_data.append({
                '因子名': factor['name'],
                '値の数': len(factor['values']),
                '値': ', '.join(factor['values'])
            })
        
        summary_df = pd.DataFrame(summary_data)
        
        # 統計情報を追加
        stats_df = pd.DataFrame([
            {'項目': '総因子数', '値': len(st.session_state.factors)},
            {'項目': '総テストケース数', '値': len(pairs_list)},
            {'項目': '全組み合わせ数（総当たり）', '値': sum(len(f['values']) for f in st.session_state.factors)},
            {'項目': '削減率', '値': f"{(1 - len(pairs_list) / max(1, sum(len(f['values']) for f in st.session_state.factors))) * 100:.1f}%"}
        ])
        
        summary_df.to_excel(writer, sheet_name='サマリー', index=False, startrow=0)
        stats_df.to_excel(writer, sheet_name='サマリー', index=False, startrow=len(summary_df) + 3)
    
    output.seek(0)
    return output

# メインUI
st.markdown("### 因子と値の設定")
st.markdown("各因子に対して値を設定してください。AllPairs法により効率的なテストケースが生成されます。")

# 因子の追加ボタン
col1, col2 = st.columns([1, 5])
with col1:
    if st.button("➕ 因子を追加", key="add_factor_btn", use_container_width=True):
        add_factor()
        st.rerun()

st.markdown("---")

# 各因子の設定
for factor_idx, factor in enumerate(st.session_state.factors):
    with st.expander(f"📁 {factor['name']}", expanded=True):
        col1, col2 = st.columns([4, 1])
        
        with col1:
            # 因子名の編集
            new_name = st.text_input(
                "因子名",
                value=factor['name'],
                key=f"factor_name_{factor_idx}"
            )
            st.session_state.factors[factor_idx]['name'] = new_name
        
        with col2:
            # 因子削除ボタン
            if len(st.session_state.factors) > 1:
                if st.button("🗑️ 削除", key=f"remove_factor_{factor_idx}"):
                    remove_factor(factor_idx)
                    st.rerun()
        
        st.markdown("**値の設定:**")
        
        # 値の一覧と編集
        for value_idx, value in enumerate(factor['values']):
            col1, col2 = st.columns([5, 1])
            
            with col1:
                new_value = st.text_input(
                    f"値 {value_idx + 1}",
                    value=value,
                    key=f"value_{factor_idx}_{value_idx}",
                    label_visibility="collapsed"
                )
                st.session_state.factors[factor_idx]['values'][value_idx] = new_value
            
            with col2:
                if len(factor['values']) > 2:
                    if st.button("❌", key=f"remove_value_{factor_idx}_{value_idx}"):
                        remove_value(factor_idx, value_idx)
                        st.rerun()
        
        # 値の追加ボタン
        if st.button("➕ 値を追加", key=f"add_value_{factor_idx}"):
            add_value(factor_idx)
            st.rerun()
        
        # プレビュー
        st.info(f"📊 値の数: {len(factor['values'])}個")

st.markdown("---")

# エクセル生成とダウンロード
st.markdown("### 📥 エクセルファイルのエクスポート")

col1, col2, col3 = st.columns([2, 2, 2])

with col1:
    st.metric("総因子数", len(st.session_state.factors))

with col2:
    total_values = sum(len(f['values']) for f in st.session_state.factors)
    st.metric("総値数", total_values)

with col3:
    # AllPairs法で生成されるテストケース数を計算
    is_valid, _ = validate_factors()
    if is_valid:
        try:
            factor_values = [factor['values'] for factor in st.session_state.factors]
            pairs_count = len(list(AllPairs(factor_values)))
            st.metric("生成テストケース数", pairs_count)
        except Exception as e:
            st.metric("生成テストケース数", "-")
    else:
        st.metric("生成テストケース数", "-")

# バリデーションチェック
is_valid, error_msg = validate_factors()
if not is_valid:
    st.warning(f"⚠️ {error_msg}")

# ダウンロードボタン
if st.button("📊 エクセルファイルを生成", type="primary", use_container_width=True, disabled=not is_valid):
    try:
        excel_file = generate_pairwise_excel()
        
        st.success("✅ エクセルファイルが生成されました！")
        
        st.download_button(
            label="💾 ダウンロード",
            data=excel_file,
            file_name="pairwise_comparison.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    except Exception as e:
        st.error(f"❌ エラーが発生しました: {str(e)}")

# 使い方の説明
with st.expander("ℹ️ 使い方"):
    st.markdown("""
    ### AllPairs法（Pairwise法）とは
    すべての因子の組み合わせをテストする代わりに、任意の2つの因子のすべての組み合わせをカバーする
    最小限のテストケースを生成する手法です。テストケース数を大幅に削減しながら、
    高いカバレッジを実現できます。
    
    ### このアプリの使い方
    1. **因子の追加**: 「➕ 因子を追加」ボタンで新しい因子を追加できます
    2. **因子名の編集**: 各因子の名前を変更できます（例: OS、ブラウザ、言語など）
    3. **値の追加**: 各因子に「➕ 値を追加」ボタンで新しい値を追加できます
    4. **値の編集**: 各値の名前を変更できます（例: Windows、Mac、Linuxなど）
    5. **削除**: 不要な因子や値は削除ボタンで削除できます（最低限の数は保持されます）
    6. **エクスポート**: 設定が完了したら「エクセルファイルを生成」ボタンでファイルを作成します
    
    ### 生成されるエクセルファイル
    - **Pairwiseテストケース**: AllPairs法で生成された最適なテストケース一覧
    - **因子間組み合わせシート**: 各2因子間の組み合わせマトリックス
    - **カバレッジ分析**: 各因子ペアのカバレッジ率を表示
    - **サマリー**: 全体の統計情報と削減効果を表示
    
    ### 使用例
    **因子**: 主食 → 値: 米, パン, ナン  
    **因子**: 副食 → 値: 肉, 魚, たこ焼き
    **因子**: 言語 → 値: プリン, ゼリー, ケーキ  
    
    総当たり: 3×3×2 = 18ケース  
    AllPairs法: 約6-9ケース（削減率50-67%）
    """)

# AllPairs法の効果を表示
with st.expander("📊 AllPairs法の効果"):
    is_valid_effect, error_msg_effect = validate_factors()
    if is_valid_effect:
        try:
            factor_values = [factor['values'] for factor in st.session_state.factors]
            
            # 総当たりのケース数
            total_cases = 1
            for factor in st.session_state.factors:
                total_cases *= len(factor['values'])
            
            # AllPairs法のケース数
            allpairs_cases = len(list(AllPairs(factor_values)))
            
            # 削減率
            reduction = (1 - allpairs_cases / total_cases) * 100 if total_cases > 0 else 0
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("総当たり", f"{total_cases}ケース")
            with col2:
                st.metric("AllPairs法", f"{allpairs_cases}ケース", delta=f"-{total_cases - allpairs_cases}")
            with col3:
                st.metric("削減率", f"{reduction:.1f}%")
            
            st.success(f"✨ AllPairs法により **{total_cases - allpairs_cases}ケース** 削減できます！")
        except Exception as e:
            st.error(f"計算エラー: {str(e)}")
    else:
        st.info(f"ℹ️ {error_msg_effect}")
