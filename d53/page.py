"""Streamlit page for D53 PO GRID (called from app.py when 部門 = D53)."""
import datetime
import re

import streamlit as st

from d53 import engine as E


@st.cache_data(show_spinner=False, max_entries=2000)
def _parse_pdf(name, data, ver):  # ver: engine version, so a parser change never reuses old cache
    return E.parse_po_pdf(name, data)


@st.cache_data(show_spinner=False, max_entries=50)
def _load_pcn(name, data, ver):
    return E.load_pcn(name, data)


@st.cache_data(show_spinner=False, max_entries=50)
def _wrk_photos(name, data, ver):
    return E.extract_wrk_photos(name, data)


def _season_default(names):
    for n in names:
        m = re.search(r"(20\d\d)", n)
        if m:
            return datetime.date(int(m.group(1)), 1, 1)
    return datetime.date(datetime.date.today().year, 1, 1)


def _parse_uploads(pdf_uploads):
    files = list(E.iter_uploaded_pdfs([(f.name, f.getvalue()) for f in pdf_uploads]))
    docs, probs = [], []
    if not files:
        return docs, probs
    bar = st.progress(0.0, text="解析 PO…")
    for i, (n, b) in enumerate(files, 1):
        d, p = _parse_pdf(n, b, E.ENGINE_VERSION)
        docs += d; probs += p
        bar.progress(i / len(files), text=f"解析 PO… {i}/{len(files)}　{n}")
    bar.empty()
    return docs, probs


def _photos(wrk_uploads):
    ph = {}
    for f in wrk_uploads or []:
        try:
            ph.update(_wrk_photos(f.name, f.getvalue(), E.ENGINE_VERSION))
        except Exception as e:  # a bad WRK file should not stop the run
            st.warning(f"{f.name} 讀不到照片：{e}")
    return ph


_MODES = {"all": "整份 PCN", "include": "只納入我負責的（貼上 DPCI 或 Style#）", "exclude": "排除少數品項／工廠"}


def _scope_filter(pcns, key, preset=None):
    """品項範圍.  Returns a settings dict: filter_mode + include/exclude lists."""
    preset = preset or {}
    rows = {d: r for p in pcns for d, r in p["rows"].items()}
    facs = sorted({r["factory"] for r in rows.values() if r["factory"]})
    st.markdown("**品項範圍**")
    modes = list(_MODES)
    mode = st.radio("品項範圍", modes, index=modes.index(E.filter_mode(preset)), format_func=_MODES.get,
                    horizontal=True, key=f"{key}_mode", label_visibility="collapsed")
    out = dict(filter_mode=mode, include_dpci=[], include_factories=[], excluded_dpci=[], excluded_factories=[])
    if mode == "include":
        c1, c2 = st.columns([3, 2])
        with c1:
            pre = "\n".join(preset.get("include_dpci", []))
            txt = st.text_area("貼上 DPCI 或 Vendor Style#（從 Excel 整欄複製即可，換行、逗號、空白都可分隔）",
                               value=pre, height=150, key=f"{key}_paste",
                               placeholder="053-00-0527\n053000528\nV27PFI04")
        with c2:
            inc_f = st.multiselect("（選填）整家工廠都是我的", facs,
                                   default=[f for f in preset.get("include_factories", []) if f in facs], key=f"{key}_incf")
        matched, hits, unmatched = E.match_tokens(pcns, txt)
        out.update(include_dpci=matched, include_factories=inc_f)
        n = len(set(matched) | {d for d, r in rows.items() if r["factory"] in inc_f})
        if n:
            st.success(f"納入 {n} 個 DPCI（PCN 共 {len(rows)} 個）")
            show = sorted(set(matched) | {d for d, r in rows.items() if r["factory"] in inc_f})
            st.dataframe([{"DPCI": d, "Style": rows[d]["style"], "品名": rows[d]["desc"], "工廠": rows[d]["factory"],
                           "Program": rows[d]["program"]} for d in show], hide_index=True, use_container_width=True,
                         height=min(38 + 35 * len(show), 320))
        else:
            st.info("還沒有納入任何品項：請貼上 DPCI／Style#，或選擇工廠。")
        if unmatched:
            st.warning(f"這 {len(unmatched)} 個在 PCN 裡找不到，已略過：{'、'.join(unmatched[:30])}")
    elif mode == "exclude":
        ex_f = st.multiselect("排除這些工廠", facs, default=[f for f in preset.get("excluded_factories", []) if f in facs], key=f"{key}_f")
        opts = {f"{d}｜{r['style']}｜{r['desc'][:40]}": d for d, r in sorted(rows.items()) if r["factory"] not in ex_f}
        inv = {v: k for k, v in opts.items()}
        ex_d = st.multiselect("排除個別 DPCI", list(opts), default=[inv[d] for d in preset.get("excluded_dpci", []) if d in inv], key=f"{key}_d")
        out.update(excluded_factories=ex_f, excluded_dpci=[opts[x] for x in ex_d])
        st.caption(f"納入 {sum(1 for r in rows.values() if r['factory'] not in ex_f) - len(ex_d)} / {len(rows)} 個 DPCI")
    else:
        st.caption(f"納入 PCN 全部 {len(rows)} 個 DPCI")
    return out


def _show_result(data, stats, fname, log=None):
    c = st.columns(5)
    c[0].metric("納入 PO", f"{stats['pos']} 張")
    c[1].metric("PO 行", f"{stats['lines']:,}")
    c[2].metric("數量", f"{stats['qty']:,}")
    c[3].metric("Program 分頁", stats.get("programs", stats["factories"]), help=f"{stats['factories']} 家工廠")
    c[4].metric("成本差異行", stats["costdiff"], help=f"PO 高於 PCN {stats['cost_hi']} 行／低於 PCN {stats['cost_lo']} 行")
    chk = stats.get("checks") or []
    if chk:
        st.markdown("**核對結果**")
        st.dataframe([{"檢查項目": c, "已檢查": n, "差異": b, "結果": "—" if n == 0 else ("✔ 相符" if b == 0 else "✖ 有差異")}
                      for c, n, b in chk], hide_index=True, use_container_width=True)
    if stats["excluded"]:
        st.caption("排除：" + "；".join(f"{k} {v} 行" for k, v in stats["excluded"].items()) + "（明細在「排除清單」分頁）")
    if stats["missing_photos"]:
        st.caption(f"缺照片 {stats['missing_photos']} 個 DPCI：補上對應的 WRK 檔再更新一次即可。")
    if log:
        with st.expander(f"本次更新紀錄（{len(log)} 筆）", expanded=True):
            st.dataframe([{k: e.get(k) for k in ("action", "po", "dpci", "factory", "old", "new", "note")} for e in log],
                         use_container_width=True, hide_index=True)
    st.download_button("⬇️ 下載 PO GRID", data, file_name=fname, type="primary",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def render_d53_page():
    st.subheader("📊 D53 PO GRID")
    st.caption(f"版本 {E.ENGINE_VERSION}")
    st.caption("以 PCN 為準：只放 PCN（與品項範圍）裡的品項；BAS 補貨單、非本季、已取消的 PO 會自動排除並列在「排除清單」。"
               "每個 Program 一個分頁（頁內依 Factory 分區塊），第一頁是 Validation Summary，再來是 Program 總覽。")
    mode = st.radio("要做什麼？", ["新建 GRID", "更新現有 GRID（加新 PO／加 Program）"], horizontal=True, key="d53_mode")

    if mode == "新建 GRID":
        c1, c2 = st.columns(2)
        with c1:
            ae_pick = st.selectbox("AE", E.AE_LIST + ["其他"], key="d53_ae")
            ae = st.text_input("AE 名稱", key="d53_ae_other") if ae_pick == "其他" else ae_pick
            pcn_up = st.file_uploader("① PCN 主檔（每份＝一個 Program，可多選）", type=["xlsx"], accept_multiple_files=True, key="d53_pcn")
        with c2:
            po_up = st.file_uploader("② PO：Import 用 SPS PDF、Domestic 用 PO Visibility .csv（整包丟進來即可，可多選或 .zip）", type=["pdf", "csv", "zip"], accept_multiple_files=True, key="d53_po")
            wrk_up = st.file_uploader("③ WRK Workspace 匯出檔（照片，選填）", type=["xlsx"], accept_multiple_files=True, key="d53_wrk")
        pcns = []
        for f in pcn_up or []:
            try:
                pcns.append(_load_pcn(f.name, f.getvalue(), E.ENGINE_VERSION))
            except Exception as e:
                st.error(str(e))
        season = st.date_input("本季起始日（Ship Window 早於這天的 PO 行會排除）",
                               value=_season_default([f.name for f in pcn_up or []]), key="d53_season")
        scope = _scope_filter(pcns, "d53_new") if pcns else dict(filter_mode="all")
        empty_inc = scope["filter_mode"] == "include" and not (scope["include_dpci"] or scope["include_factories"])
        if st.button("產生 PO GRID", type="primary", disabled=not (pcns and po_up and ae) or empty_inc, key="d53_build"):
            docs, probs = _parse_uploads(po_up)
            with st.spinner("產生 GRID…"):
                state = E.new_state(ae, season, pcns, docs, probs, scope.get("excluded_factories", []),
                                    scope.get("excluded_dpci", []), scope["filter_mode"],
                                    scope.get("include_factories", []), scope.get("include_dpci", []))
                data, stats = E.render(state, _photos(wrk_up))
            st.session_state["d53_out"] = (data, stats, f"D53_PO_GRID_{ae}_{datetime.date.today():%Y%m%d}.xlsx", None, probs)
        if not (pcns and po_up and ae):
            st.info("請上傳 PCN 與 PO PDF。")
    else:
        grid_up = st.file_uploader("① 現有的 D53 PO GRID（App 產生的 .xlsx）", type=["xlsx"], key="d53_grid")
        state = photos = manual = None
        if grid_up:
            try:
                state, photos, manual = E.read_grid(grid_up.getvalue())
            except Exception as e:
                st.error(str(e))
        if state:
            S = state["settings"]
            st.success(f"AE：{S.get('ae')}｜Program：{'、'.join(state['pcn'])}｜本季起始日：{S['season_start']}｜"
                       f"PO 文件 {len(state['docs'])} 份｜上次更新：{state['changelog'][-1]['time'] if state['changelog'] else '-'}")
            c1, c2 = st.columns(2)
            with c1:
                po_up = st.file_uploader("② 新的 PO：PDF 或 Domestic .csv（新單、改單、CANCEL 都可以；可 .zip）", type=["pdf", "csv", "zip"], accept_multiple_files=True, key="d53_po_u")
                wrk_up = st.file_uploader("④ WRK 匯出檔（補照片，選填）", type=["xlsx"], accept_multiple_files=True, key="d53_wrk_u")
            with c2:
                pcn_up = st.file_uploader("③ PCN（加新的 Program，或同名 PCN 的新版本，選填）", type=["xlsx"], accept_multiple_files=True, key="d53_pcn_u")
            new_pcns = []
            for f in pcn_up or []:
                try:
                    new_pcns.append(_load_pcn(f.name, f.getvalue(), E.ENGINE_VERSION))
                except Exception as e:
                    st.error(str(e))
            season = st.date_input("本季起始日", value=datetime.date.fromisoformat(S["season_start"]), key="d53_season_u")
            all_pcns = list({**state["pcn"], **{p["program"]: p for p in new_pcns}}.values())
            scope = _scope_filter(all_pcns, "d53_upd", S)
            if st.button("更新 PO GRID", type="primary", key="d53_update"):
                docs, probs = _parse_uploads(po_up or [])
                with st.spinner("更新 GRID…"):
                    photos.update(_photos(wrk_up))
                    state, log, _ = E.merge_update(state, docs, new_pcns, probs,
                                                   dict(season_start=season.isoformat(), **scope))
                    data, stats = E.render(state, photos, manual)
                st.session_state["d53_out"] = (data, stats, f"D53_PO_GRID_{S.get('ae')}_{datetime.date.today():%Y%m%d}.xlsx", log, probs)

    out = st.session_state.get("d53_out")
    if out:
        data, stats, fname, log, probs = out
        st.markdown("---")
        for p in probs or []:
            st.warning(p)
        _show_result(data, stats, fname, log)
