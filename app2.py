#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="小红书数据批量分析平台", layout="wide")
st.title("📊 小红书数据批量分析与报告生成（在线可视化）")
st.markdown("上传一个或多个 Excel 文件，系统会逐个分析并显示结果，最后可下载汇总报告。")

# =====================================================================
# 分析函数
# =====================================================================
def analyze_and_display(df, filename):
    st.header(f"📘 分析报告：【{filename}】")

    # ---------- 列名规范 ----------
    df.columns = df.columns.astype(str).str.strip()
    rename_map = {
        "曝光量": "曝光", "阅读量": "观看量", "播放量": "观看量", "观看数": "观看量",
        "点赞数": "点赞","获赞":"点赞","获赞数":"点赞","点赞次数":"点赞",
        "收藏数": "收藏","评论数": "评论","涨粉数": "涨粉","净涨粉":"涨粉",
        "发布形式":"体裁"
    }
    df.rename(columns=rename_map, inplace=True)

    required_cols = ["笔记标题","曝光","点赞","观看量","收藏","评论","涨粉","分享",
                     "封面点击率","首次发布时间","体裁"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(f"缺少必要列：{missing}")
        return None

    # ---------- 日期处理 ----------
    df["首次发布时间"] = pd.to_datetime(df["首次发布时间"], format='%Y年%m月%d日%H时%M分%S秒',
                                   errors='coerce')
    df.dropna(subset=["首次发布时间"], inplace=True)
    df.sort_values(by="首次发布时间", ascending=True, inplace=True)
    df.reset_index(drop=True, inplace=True)
    df.insert(0, "序号", df.index + 1)
    st.markdown(f"**📅 数据时间**：{df['首次发布时间'].min().date()} ➜ {df['首次发布时间'].max().date()}")

    # ---------- 指标计算 ----------
    for c in ["曝光","封面点击率","点赞","观看量","收藏","评论","涨粉","分享"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    df["点赞率"] = df["点赞"] / df["观看量"].replace(0, pd.NA)
    df["收藏率"] = df["收藏"] / df["观看量"].replace(0, pd.NA)
    df["赞藏比"] = df["点赞"] / df["收藏"].replace(0, pd.NA)
    df["评论率"] = df["评论"] / df["观看量"].replace(0, pd.NA)
    df["互动率"] = (df["点赞"] + df["评论"] + df["收藏"]) / df["观看量"].replace(0, pd.NA)
    df["有效活跃度"] = df["评论"] / (df["点赞"] + df["收藏"]).replace(0, pd.NA)
    df["转粉率"] = df["涨粉"] / df["观看量"].replace(0, pd.NA)

    # ---------- 展示数据表 ----------
    st.subheader("📄 计算结果数据表")
    display_cols = [
        "序号","笔记标题","首次发布时间","体裁","曝光","观看量","封面点击率",
        "点赞","评论","收藏","涨粉","分享",
        "点赞率","收藏率","互动率","转粉率","赞藏比","有效活跃度"
    ]
    st.dataframe(df[display_cols].style.format({
        "首次发布时间": "{:%Y-%m-%d %H:%M}",
        "封面点击率": "{:.2%}","点赞率":"{:.2%}","收藏率":"{:.2%}",
        "互动率":"{:.2%}","转粉率":"{:.2%}","赞藏比":"{:.2f}","有效活跃度":"{:.2f}"
    }))

    # ---------- 指标平均值 ----------
    st.subheader("📈 核心指标平均值")
    avg = df[["封面点击率","点赞率","收藏率","互动率","转粉率"]].mean()
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("平均封面点击率", f"{avg['封面点击率']:.2%}")
    c2.metric("平均点赞率", f"{avg['点赞率']:.2%}")
    c3.metric("平均收藏率", f"{avg['收藏率']:.2%}")
    c4.metric("平均互动率", f"{avg['互动率']:.2%}")

    # ---------- 使用 Streamlit 原生可视化 ----------
    st.subheader("🎨 内容形式分布")
    form_count = df["体裁"].value_counts().reset_index()
    form_count.columns = ["体裁","数量"]
    st.bar_chart(data=form_count, x="体裁", y="数量")

    st.subheader("📈 核心互动指标趋势")
    chart1 = df[["序号","点赞率","收藏率","互动率"]].set_index("序号")
    st.line_chart(chart1)

    st.subheader("📈 转化与活跃度趋势")
    chart2 = df[["序号","转粉率","有效活跃度"]].set_index("序号")
    st.line_chart(chart2)

    st.subheader("📈 基础数据表现")
    chart3 = df[["序号","曝光","观看量","点赞","收藏","涨粉","分享"]].set_index("序号")
    st.line_chart(chart3)

    return df

# =====================================================================
# 主入口：多文件上传 + 汇总下载
# =====================================================================
uploaded_files = st.file_uploader("请上传 Excel 文件（可多选）",
                                  type=["xls","xlsx"],
                                  accept_multiple_files=True)

if uploaded_files:
    processed_dfs = {}
    for up in uploaded_files:
        df_raw = pd.read_excel(up, header=1)
        df_final = analyze_and_display(df_raw, up.name)
        if df_final is not None:
            sheet_name = ''.join(e for e in up.name if e.isalnum())[:31]
            processed_dfs[sheet_name] = df_final

    if processed_dfs:
        st.header("📥 下载汇总报告")
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            for sn, d in processed_dfs.items():
                d.to_excel(writer, index=False, sheet_name=sn)
        st.download_button(
            "下载完整汇总Excel",
            data=buffer.getvalue(),
            file_name="小红书分析汇总报告.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success("分析完成 ✅")
else:
    st.info("👆 上传一个或多个 Excel 文件即可开始分析。")

