import io, re, logging, time
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title='TIKTOK SHOP | HIỆU QUẢ ADS NGỌC MỸ PROCARE',page_icon='📊',layout='wide')
logging.basicConfig(level=logging.INFO)
logger=logging.getLogger('ngocmy_dashboard')
st.markdown("""<style>
:root{--ink:#12304a;--navy:#123c66;--blue:#1767a5;--teal:#087f8c;--red:#b83c47;--gold:#b7791f}
.block-container{padding-top:1.35rem;max-width:1530px;padding-bottom:3rem}
h1,h2,h3{color:#163754!important;letter-spacing:-.022em}
h1{font-size:1.85rem!important;font-weight:850!important}
div[data-testid="stMetric"]{background:linear-gradient(145deg,#f5faff,#edf5fb);border:1px solid #c9dbe9;padding:16px;border-radius:14px;box-shadow:0 3px 10px #123c6609}
div[data-testid="stMetricLabel"]{color:#31546e}
div[data-testid="stMetricValue"]{color:#0b4d79;font-weight:800}
div[role="radiogroup"]{gap:.35rem;flex-wrap:wrap}
div[role="radiogroup"] label{background:#eaf3f9;border:1px solid #c5dceb;border-radius:10px;padding:.45rem .7rem;white-space:nowrap}
div[role="radiogroup"] label:has(input:checked){background:#125c8f;color:white;border-color:#125c8f}
div[data-testid="stAlert"]{border-radius:12px}
[data-testid="stSidebar"]{background:#f0f6fa}
</style>""",unsafe_allow_html=True)
PLOT_COLORS=['#12639a','#12a59b','#e3a33d','#c9555c','#7c6cba']
def style_fig(fig,height=None):
    fig.update_layout(template='plotly_white',font=dict(family='Arial',color='#1b3952',size=13),
        paper_bgcolor='rgba(0,0,0,0)',plot_bgcolor='#ffffff',
        margin=dict(l=25,r=25,t=32,b=45),legend=dict(orientation='h',y=1.13,x=0),
        colorway=PLOT_COLORS)
    fig.update_xaxes(showgrid=True,gridcolor='#e5edf3',zeroline=False)
    fig.update_yaxes(showgrid=True,gridcolor='#e5edf3',zeroline=False)
    if height:fig.update_layout(height=height)
    return fig

def analysis(text):
    st.markdown('#### 📌 Phân tích chung')
    st.markdown(f'<div style="background:#eff7fc;border-left:5px solid #16749c;border-radius:9px;padding:15px 19px;color:#193b54;line-height:1.75">{text}</div>',unsafe_allow_html=True)


@st.cache_data(show_spinner=False, max_entries=15, ttl=3600)
def read_xlsx_bytes(data):
    return pd.read_excel(io.BytesIO(data), engine='openpyxl', dtype={'ID chiến dịch':str,'ID sản phẩm':str,'ID bài đăng':str})

@st.cache_data(show_spinner=False, max_entries=8, ttl=3600)
def read_orders_bytes(data):
    for enc in ('utf-8-sig','utf-16','cp1258'):
        try:
            return pd.read_csv(io.BytesIO(data),encoding=enc,dtype=str,sep=',',engine='c',on_bad_lines='warn')
        except (UnicodeError,ValueError):
            pass
    raise ValueError('Không đọc được file CSV')

def read_xlsx(file):
    return read_xlsx_bytes(file.getvalue())

def read_orders(file):
    return read_orders_bytes(file.getvalue())

def num(s):
    return pd.to_numeric(s.astype(str).str.replace(',','',regex=False).str.replace('₫','',regex=False).str.strip().replace({'-':None,'':None,'nan':None}),errors='coerce').fillna(0)

def clean_id(s):return s.astype(str).str.replace(r'\D','',regex=True)
def vnd(v):return f'{v:,.0f}đ'.replace(',','.')
def ratio(n,d):return n/d if d>0 else np.nan
def show_roas(v):return f'{v:.2f}' if pd.notna(v) else '—'

def find_col(df,options):
    for c in options:
        if c in df.columns:return c
    return None

st.title('TIKTOK SHOP | HIỆU QUẢ ADS NGỌC MỸ PROCARE')
st.caption('GMV Max • ROAS = Doanh thu / Chi phí Ads • Không cộng doanh thu từ nhiều cấp báo cáo')
with st.sidebar:
    st.header('1. Tải dữ liệu')
    overview_file=st.file_uploader('Campaign overview (.xlsx)',type=['xlsx'],key='overview')
    campaign_file=st.file_uploader('Product campaign (.xlsx)',type=['xlsx'],key='campaign')
    product_file=st.file_uploader('Product data (.xlsx)',type=['xlsx'],key='product')
    creative_file=st.file_uploader('Creative data (.xlsx)',type=['xlsx'],key='creative')
    order_file=st.file_uploader('Tất cả đơn hàng (.csv)',type=['csv'],key='orders')
    st.divider()
    st.header('2. Thiết lập')
    include_pending=st.toggle('Tính đơn chưa hủy nhưng chưa giao',value=True,help='Nếu tắt, chỉ tính đơn đã giao. Cả hai cách đều không phải doanh thu Ads được TikTok xác nhận.')
    st.caption('Tải từng tệp độc lập. Dashboard sẽ hiển thị những phần đủ dữ liệu.')
    st.divider()
    st.subheader('🔍 Đối soát dữ liệu')
    def open_audit():
        st.session_state['audit_page_open']=True
    st.button('Mở tab Đối soát',key='open_audit_sidebar',use_container_width=True,on_click=open_audit,
              help='Kiểm tra chênh lệch giữa các báo cáo; không cộng trùng doanh thu và chi phí.')

frames={}
for key,file,reader in [('overview',overview_file,read_xlsx),('campaign',campaign_file,read_xlsx),('product',product_file,read_xlsx),('creative',creative_file,read_xlsx),('orders',order_file,read_orders)]:
    if file is not None:
        try:frames[key]=reader(file)
        except Exception as e:
            logger.exception('Lỗi đọc %s',key)
            st.error(f'Lỗi đọc {key}: {type(e).__name__}: {e}')
if not frames:
    st.info('Hãy tải các tệp xuất từ TikTok Ads và tệp đơn hàng ở thanh bên trái. Có thể tải 1–5 tệp, nhưng cần Campaign overview + CSV đơn hàng để tính ROAS điều chỉnh tham khảo.')
    st.stop()

ov=frames.get('overview');camp=frames.get('campaign');prod=frames.get('product');creative=frames.get('creative');orders=frames.get('orders')
if ov is not None:
    ov=ov.copy();ov['Chi phí']=num(ov['Chi phí']);ov['Doanh thu gộp (Cửa hàng hiện tại)']=num(ov['Doanh thu gộp (Cửa hàng hiện tại)']);ov['Theo ngày']=pd.to_datetime(ov['Theo ngày'],errors='coerce');ov=ov[ov['Theo ngày'].notna()]
if camp is not None:
    camp=camp.copy();camp['Chi phí']=num(camp['Chi phí']);camp['Doanh thu gộp']=num(camp['Doanh thu gộp']);camp['Số lượng đơn hàng SKU']=num(camp['Số lượng đơn hàng SKU']);camp['ID chiến dịch']=clean_id(camp['ID chiến dịch']);camp['ROAS tính lại']=camp['Doanh thu gộp'].div(camp['Chi phí'].replace(0,np.nan))
if prod is not None:
    prod=prod.copy();prod['Chi phí']=num(prod['Chi phí']);prod['Doanh thu gộp']=num(prod['Doanh thu gộp']);prod['Số lượng đơn hàng SKU']=num(prod['Số lượng đơn hàng SKU']);prod['ID sản phẩm']=clean_id(prod['ID sản phẩm']);prod['ID chiến dịch']=clean_id(prod['ID chiến dịch']);prod['ROAS tính lại']=prod['Doanh thu gộp'].div(prod['Chi phí'].replace(0,np.nan))
if creative is not None:
    creative=creative.copy()
    for c in ['Chi phí','Doanh thu gộp','Số lượng đơn hàng SKU','Số lượt hiển thị quảng cáo sản phẩm','Số lượt nhấp vào quảng cáo sản phẩm']:
        if c in creative:creative[c]=num(creative[c])
if orders is not None:
    orders=orders.copy()
    needed=['Order ID','Order Status','SKU ID','Order Amount','Product Name']
    missing=[c for c in needed if c not in orders]
    if missing:
        st.error('CSV thiếu cột: '+', '.join(missing));orders=None
    else:
        orders['Order ID']=clean_id(orders['Order ID']);orders['SKU ID']=clean_id(orders['SKU ID'])
        orders['Order Amount']=num(orders['Order Amount']);orders['Quantity']=num(orders['Quantity']) if 'Quantity' in orders else 1
        orders['Đã hủy']=orders['Order Status'].astype(str).str.contains('hủy|cancel',case=False,regex=True,na=False)
        orders['Đã giao']=orders['Order Status'].astype(str).str.contains('đã giao|delivered|completed|hoàn thành',case=False,regex=True,na=False)
        orders['Hợp lệ theo lựa chọn']=(~orders['Đã hủy']) & (True if include_pending else orders['Đã giao'])
        orders['Ngày tạo']=pd.to_datetime(orders['Created Time'].astype(str).str.strip() if 'Created Time' in orders else pd.Series(dtype=str),dayfirst=True,errors='coerce')
        # Order Amount is repeated on multi-line orders; count it only once per Order ID.
        order_level=orders.groupby('Order ID',as_index=False).agg({'Order Amount':'max','Đã hủy':'max','Đã giao':'max','Hợp lệ theo lựa chọn':'max','Ngày tạo':'min'})
        # If a multi-line order contains mixed statuses, avoid treating it as fully retained.
        order_level['Hợp lệ theo lựa chọn']=(~order_level['Đã hủy']) & (True if include_pending else order_level['Đã giao'])
        order_level['Doanh thu giữ lại']=np.where(order_level['Hợp lệ theo lựa chọn'],order_level['Order Amount'],0)

st.markdown('### Điều hướng báo cáo')
page_names=['🏠 Tổng quan','📅 Theo ngày','📦 Đơn hàng','🎯 Chiến dịch','🧴 Sản phẩm','🎬 Creative','🔍 Đối soát','🎥 Phân tích Video','🛍️ Phân tích Thẻ sản phẩm','🧭 Hành động theo sản phẩm']
# Đối soát chuyển riêng sang sidebar, các trang khác giữ menu ngang.
horizontal_pages=[p for p in page_names if p!='🔍 Đối soát']
def open_horizontal_page():
    st.session_state['audit_page_open']=False
horizontal_selection=st.radio('Chọn trang phân tích',horizontal_pages,index=0,key='selected_page',horizontal=True,
                              label_visibility='collapsed',on_change=open_horizontal_page)
selected_page='🔍 Đối soát' if st.session_state.get('audit_page_open',False) else horizontal_selection
active=page_names.index(selected_page)
st.caption(f'Đang xem: {selected_page}')
with st.sidebar.expander('⚡ Hiệu năng / khắc phục tải chậm'):
    st.caption('Dữ liệu Excel/CSV được lưu cache trong 60 phút, không đọc lại khi đổi tab. Lần đầu tải có thể chậm do khởi động Streamlit Cloud.')
    if st.button('Xóa cache dữ liệu', key='clear_data_cache'):
        st.cache_data.clear()
        st.rerun()

if active == 0:
    st.subheader('01 / Tổng quan')
    spend=float(ov['Chi phí'].sum()) if ov is not None else float(camp['Chi phí'].sum()) if camp is not None else float(prod['Chi phí'].sum()) if prod is not None else np.nan
    ads_revenue=float(ov['Doanh thu gộp (Cửa hàng hiện tại)'].sum()) if ov is not None else float(camp['Doanh thu gộp'].sum()) if camp is not None else float(prod['Doanh thu gộp'].sum()) if prod is not None else np.nan
    retained=float(order_level['Doanh thu giữ lại'].sum()) if orders is not None else np.nan
    cancel_value=float(order_level.loc[order_level['Đã hủy'],'Order Amount'].sum()) if orders is not None else np.nan
    cols=st.columns(4)
    cols[0].metric('Chi phí Ads',vnd(spend) if pd.notna(spend) else '—')
    cols[1].metric('Doanh thu Ads báo cáo',vnd(ads_revenue) if pd.notna(ads_revenue) else '—')
    cols[2].metric('ROAS Ads báo cáo',show_roas(ratio(ads_revenue,spend)))
    cols[3].metric('ROAS điều chỉnh tham khảo',show_roas(ratio(retained,spend)))
    if creative is not None and len(creative):
        clicks_total=float(creative['Số lượt nhấp vào quảng cáo sản phẩm'].sum())
        sku_total=float(creative['Số lượng đơn hàng SKU'].sum())
        st.metric('CVR theo SKU / lượt nhấp (Creative)',f'{100*sku_total/clicks_total:.2f}%' if clicks_total>0 else '—')
        st.caption('CVR theo SKU = số lượng SKU quy thuộc / lượt nhấp quảng cáo sản phẩm. Đây không phải CVR theo số đơn duy nhất hay đơn đã giao.')
    if orders is not None:
        c=st.columns(4)
        c[0].metric('Số đơn (không trùng)',len(order_level))
        c[1].metric('Đơn hủy',int(order_level['Đã hủy'].sum()))
        c[2].metric('Tỷ lệ hủy',f"{order_level['Đã hủy'].mean():.1%}")
        c[3].metric('Giá trị đơn hủy',vnd(cancel_value))
        st.caption(f"Doanh thu đơn hàng theo trạng thái đã chọn: {vnd(retained)}. Tổng giá trị đơn hàng CSV: {vnd(order_level['Order Amount'].sum())}.")
        if pd.notna(ads_revenue):
            delta=ads_revenue-order_level['Order Amount'].sum()
            if abs(delta)>0.01:st.warning(f'Chênh lệch giữa doanh thu Ads và tổng giá trị đơn CSV: {vnd(delta)}. Hai nguồn có thể khác phạm vi, giá trị khuyến mãi hoặc thời điểm ghi nhận. ROAS điều chỉnh là chỉ số THAM KHẢO, không phải ROAS Ads được đối soát theo từng đơn.')
        st.info('Đơn không hủy chưa chắc đã giao/thanh toán thành công. Doanh thu CSV không đồng nghĩa toàn bộ là doanh thu được quy cho Ads. Nếu chưa có mã đơn liên kết với quảng cáo, không thể xác định ROAS Ads sau hủy chính xác.')



    if pd.notna(spend) and spend>0:
        analysis(f'Chi phí Ads {vnd(spend)} tạo doanh thu quy thuộc {vnd(ads_revenue)} (ROAS {show_roas(ratio(ads_revenue,spend))}). '+(f'Danh sách đơn hàng có {int(order_level["Đã hủy"].sum())}/{len(order_level)} đơn hủy ({order_level["Đã hủy"].mean():.1%}); doanh thu đơn không hủy theo trạng thái chọn là {vnd(retained)}. ' if orders is not None else 'Chưa tải danh sách đơn để kiểm tra hủy. ')+'Đây là tín hiệu hiệu quả ghi nhận, chưa phải lợi nhuận hay ROAS Ads đã đối soát theo mã đơn.')

if active == 1:
    if ov is not None:
        st.subheader('02 / Hiệu quả theo ngày')
        day=ov.groupby('Theo ngày',as_index=False)[['Chi phí','Doanh thu gộp (Cửa hàng hiện tại)']].sum().sort_values('Theo ngày')
        day['ROAS']=day['Doanh thu gộp (Cửa hàng hiện tại)'].div(day['Chi phí'].replace(0,np.nan))
        fig=px.bar(day.melt('Theo ngày',value_vars=['Chi phí','Doanh thu gộp (Cửa hàng hiện tại)'],var_name='Chỉ tiêu',value_name='VND'),x='Theo ngày',y='VND',color='Chỉ tiêu',barmode='group',color_discrete_sequence=['#c9555c','#12639a'])
        fig.update_layout(margin=dict(l=0,r=0,t=10,b=0),legend_title_text='')
        st.plotly_chart(style_fig(fig),use_container_width=True)
        st.dataframe(day.rename(columns={'Theo ngày':'Ngày','Doanh thu gộp (Cửa hàng hiện tại)':'Doanh thu Ads'}).style.format({'Chi phí':vnd,'Doanh thu Ads':vnd,'ROAS':show_roas}),use_container_width=True,hide_index=True)



        if len(day):
            top=day.loc[day['Doanh thu gộp (Cửa hàng hiện tại)'].idxmax()]
            empty=day[(day['Chi phí']>0)&(day['Doanh thu gộp (Cửa hàng hiện tại)']==0)]
            analysis(f'Ngày doanh thu cao nhất: {top["Theo ngày"]:%d/%m/%Y}, {vnd(top["Doanh thu gộp (Cửa hàng hiện tại)"])} trên chi phí {vnd(top["Chi phí"])}. Có {len(empty)} ngày tiêu ngân sách nhưng không ghi nhận doanh thu, tổng {vnd(empty["Chi phí"].sum())}. Biến động ngày không chứng minh nguyên nhân do thuật toán hay nội dung nếu chưa có dữ liệu thay đổi chiến dịch.')

if active == 2:
    if orders is not None:
        st.subheader('03 / Đơn hàng và hủy đơn')
        a,b=st.columns([1,2])
        with a:
            fig=px.pie(names=['Hủy','Không hủy'],values=[int(order_level['Đã hủy'].sum()),int((~order_level['Đã hủy']).sum())],hole=.62,color_discrete_sequence=['#c9555c','#10a38f']);fig.update_layout(margin=dict(l=0,r=0,t=10,b=0));st.plotly_chart(style_fig(fig),use_container_width=True)
        with b:
            view_cols=[c for c in ['Order ID','Created Time','Product Name','SKU ID','Order Status','Order Substatus','Order Amount','Cancel Reason','Cancellation Reason','Cancelled Time'] if c in orders.columns]
            st.dataframe(orders[view_cols],use_container_width=True,hide_index=True)
        st.download_button('Tải CSV đơn hàng đã chuẩn hóa',orders.to_csv(index=False,encoding='utf-8-sig').encode('utf-8-sig'),file_name='don_hang_chuan_hoa.csv',mime='text/csv')



        analysis(f'Có {len(order_level)} đơn duy nhất, {int(order_level["Đã hủy"].sum())} đơn hủy ({order_level["Đã hủy"].mean():.1%}), giá trị hủy {vnd(order_level.loc[order_level["Đã hủy"],"Order Amount"].sum())}. Doanh thu còn lại theo trạng thái lựa chọn {vnd(order_level["Doanh thu giữ lại"].sum())}. Cần phân biệt đơn chưa hủy với đơn đã giao và không tự quy toàn bộ đơn CSV cho Ads.')

if active == 3:
    if camp is not None:
        st.subheader('04 / Chiến dịch GMV Max')
        cv=camp[['ID chiến dịch','Tên chiến dịch','Chi phí','Doanh thu gộp','Số lượng đơn hàng SKU','ROAS tính lại']].copy()
        st.dataframe(cv.style.format({'Chi phí':vnd,'Doanh thu gộp':vnd,'ROAS tính lại':show_roas}),hide_index=True,use_container_width=True)
        fig=px.bar(camp.sort_values('Chi phí',ascending=True),x='Chi phí',y='Tên chiến dịch',orientation='h',color='Doanh thu gộp',color_continuous_scale=['#d5ecf5','#3188b7','#10476c']);fig.update_layout(height=max(350,len(camp)*48),margin=dict(l=0,r=0,t=15,b=0));st.plotly_chart(style_fig(fig),use_container_width=True)



        if len(camp):
            zero=camp[(camp['Chi phí']>0)&(camp['Doanh thu gộp']==0)]
            best=camp.loc[camp['Doanh thu gộp'].idxmax()]
            analysis(f'Có {len(camp)} dòng chiến dịch; {len(zero)} dòng có chi phí nhưng chưa ghi nhận doanh thu, tiêu {vnd(zero["Chi phí"].sum())}. Nhóm doanh thu cao nhất là {best.get("Tên chiến dịch","không rõ")} với {vnd(best["Doanh thu gộp"])}. ROAS cần đọc cùng quy mô chi phí và số SKU, tránh kết luận từ một đơn lẻ.')

if active == 4:
    if prod is not None:
        st.subheader('05 / Sản phẩm & đối chiếu đơn hủy theo SKU')
        p=prod.copy()
        if orders is not None:
            # SKU-level monetary allocation only safe for single-SKU orders.
            sk=orders.groupby('SKU ID',as_index=False).agg(so_dong=('Order ID','size'),so_don=('Order ID','nunique'),so_don_huy=('Đã hủy','sum'),so_dong_khong_huy=('Hợp lệ theo lựa chọn','sum'))
            p=p.merge(sk,how='left',left_on='ID sản phẩm',right_on='SKU ID')
            p['so_don_huy']=p['so_don_huy'].fillna(0).astype(int)
            p['so_don']=p['so_don'].fillna(0).astype(int)
            p['Đánh giá']=np.select([p['so_don_huy']>0,p['Số lượng đơn hàng SKU']<=0],['Có đơn hủy — cần kiểm tra','Chưa ghi nhận chuyển đổi'],default='Có đơn — kiểm tra lợi nhuận')
        else:p['Đánh giá']=np.where(p['Số lượng đơn hàng SKU']>0,'Có chuyển đổi (chưa đối chiếu hủy)','Chưa ghi nhận chuyển đổi')
        st.dataframe(p[[c for c in ['Tên sản phẩm','ID sản phẩm','Tên chiến dịch','Chi phí','Doanh thu gộp','Số lượng đơn hàng SKU','ROAS tính lại','so_don','so_don_huy','Đánh giá'] if c in p]].style.format({'Chi phí':vnd,'Doanh thu gộp':vnd,'ROAS tính lại':show_roas}),hide_index=True,use_container_width=True)
        if orders is not None:
            missing_sku=set(orders['SKU ID'])-set(prod['ID sản phẩm'])
            if missing_sku:st.warning(f'{len(missing_sku)} SKU trong đơn hàng không khớp ID sản phẩm của báo cáo Ads. Không tự gán theo tên vì có thể sai.')
        st.download_button('Tải bảng đối chiếu sản phẩm',p.to_csv(index=False).encode('utf-8-sig'),file_name='doi_chieu_san_pham.csv',mime='text/csv')



        if len(p):
            no=p[(p['Chi phí']>0)&(p['Số lượng đơn hàng SKU']==0)]
            yes=p[p['Số lượng đơn hàng SKU']>0]
            analysis(f'{len(yes)}/{len(p)} dòng sản phẩm có SKU quy thuộc; {len(no)} dòng tiêu {vnd(no["Chi phí"].sum())} mà chưa ghi nhận SKU. Ưu tiên kiểm tra nhóm chi phí cao, CTR thấp và không có SKU; đối chiếu đơn hủy theo SKU trước khi quyết định tăng ngân sách.')

if active == 5:
    st.subheader('06 / Creative – So sánh thẻ sản phẩm và video')
    if creative is None:
        st.info('Tải Creative data để xem bảng so sánh.')
    else:
        d=prep_deep(creative) if 'prep_deep' in globals() else creative.copy()
        kind=creative['Loại nội dung sáng tạo'].fillna('').astype(str)
        d=creative.copy()
        d['Hình thức']=np.select([kind.str.contains('video',case=False),kind.str.contains('thẻ',case=False)],['Video','Thẻ sản phẩm'],default='Khác')
        d=d[d['Hình thức'].isin(['Video','Thẻ sản phẩm'])]
        fields={'Chi phí':'Chi phí','Doanh thu':'Doanh thu gộp','Đơn SKU':'Số lượng đơn hàng SKU','Hiển thị SP':'Số lượt hiển thị quảng cáo sản phẩm','Nhấp SP':'Số lượt nhấp vào quảng cáo sản phẩm'}
        for col in fields.values():
            if col not in d:d[col]=0
        agg=d.groupby('Hình thức').agg(**{k:(v,'sum') for k,v in fields.items()}).reindex(['Thẻ sản phẩm','Video']).fillna(0)
        total=agg['Chi phí'].sum()
        agg['Tỷ trọng chi phí (%)']=np.where(total>0,agg['Chi phí']/total*100,np.nan)
        agg['ROAS']=agg['Doanh thu'].div(agg['Chi phí'].replace(0,np.nan))
        agg['Chi phí/đơn SKU']=agg['Chi phí'].div(agg['Đơn SKU'].replace(0,np.nan))
        agg['CTR (%)']=100*agg['Nhấp SP'].div(agg['Hiển thị SP'].replace(0,np.nan))
        agg['CVR SKU/Click (%)']=100*agg['Đơn SKU'].div(agg['Nhấp SP'].replace(0,np.nan))
        labels=[('Chi phí','Chi phí','money'),('Tỷ trọng chi phí creative','Tỷ trọng chi phí (%)','pct'),('Doanh thu gộp','Doanh thu','money'),('ROAS','ROAS','decimal'),('Đơn SKU ghi nhận','Đơn SKU','int'),('Chi phí/đơn SKU','Chi phí/đơn SKU','money'),('Hiển thị quảng cáo SP','Hiển thị SP','int'),('Nhấp quảng cáo SP','Nhấp SP','int'),('CTR tổng hợp','CTR (%)','pct'),('CVR theo SKU / lượt nhấp','CVR SKU/Click (%)','pct')]
        def fmt(value,typ):
            if pd.isna(value):return '—'
            if typ=='money':return vnd(value)
            if typ=='pct':return f'{value:.2f}%'.replace('.',',')
            if typ=='decimal':return f'{value:.2f}'.replace('.',',')
            return f'{value:,.0f}'.replace(',','.')
        rows=[{'Chỉ số':label,**{k:fmt(agg.loc[k,col],typ) for k in agg.index}} for label,col,typ in labels]
        st.dataframe(pd.DataFrame(rows),hide_index=True,use_container_width=True)
        fig=px.bar(agg.reset_index(),x='Hình thức',y='CTR (%)',color='Hình thức',text='CTR (%)',color_discrete_map={'Thẻ sản phẩm':'#12639a','Video':'#12a59b'})
        fig.update_traces(texttemplate='%{text:.2f}%',textposition='outside')
        fig.update_layout(showlegend=False,yaxis_title='CTR (%)',xaxis_title='',height=320)
        st.plotly_chart(style_fig(fig),use_container_width=True)
        card=agg.loc['Thẻ sản phẩm'];video=agg.loc['Video']
        if pd.notna(card['CTR (%)']) and pd.notna(video['CTR (%)']) and card['CTR (%)']>0:
            factor=video['CTR (%)']/card['CTR (%)']
            analysis(f'Video có CTR {video["CTR (%)"]:.2f}% so với thẻ {card["CTR (%)"]:.2f}% (gấp {factor:.2f} lần). ROAS tương ứng {show_roas(video["ROAS"])} và {show_roas(card["ROAS"])}. Lượt nhấp cao hơn không đồng nghĩa doanh thu tốt hơn; cần kiểm tra chất lượng nhấp, trang đích và đơn hủy. Chi phí Creative {vnd(total)} không cộng vào Campaign Overview.')
        st.caption('CTR = tổng nhấp / tổng hiển thị; CVR = SKU quy thuộc / lượt nhấp, không phải đơn hoàn tất. Mọi số liệu tự cập nhật theo file tải lên.')

if active == 6:
    st.subheader('07 / Kiểm tra chất lượng dữ liệu')
    checks=[]
    for name,df,sp,rev in [('Campaign overview',ov,'Chi phí','Doanh thu gộp (Cửa hàng hiện tại)'),('Product campaign',camp,'Chi phí','Doanh thu gộp'),('Product data',prod,'Chi phí','Doanh thu gộp'),('Creative data',creative,'Chi phí','Doanh thu gộp')]:
        if df is not None:checks.append({'Nguồn':name,'Số dòng':len(df),'Chi phí':df[sp].sum() if sp in df else np.nan,'Doanh thu':df[rev].sum() if rev in df else np.nan})
    if checks:st.dataframe(pd.DataFrame(checks).style.format({'Chi phí':vnd,'Doanh thu':vnd}),use_container_width=True,hide_index=True)
    st.caption('Nguồn chuẩn chi phí và doanh thu tổng quan: Campaign overview (nếu có). Các bảng chiến dịch, sản phẩm và Creative dùng để phân tích chi tiết, không cộng dồn. Dữ liệu đơn hàng chỉ tính một lần theo Order ID.')



    if checks:
        ck=pd.DataFrame(checks); base=ck.iloc[0]
        analysis(f'Có {len(ck)} nguồn Ads để đối soát. Nguồn {base["Nguồn"]} ghi nhận {vnd(base["Chi phí"])} chi phí và {vnd(base["Doanh thu"])} doanh thu. Chênh lệch chi phí giữa các cấp báo cáo có thể do phạm vi quy thuộc; không cộng các nguồn và cần kiểm tra bộ lọc ngày, múi giờ, loại quảng cáo trước khi xác nhận nguyên nhân.')

# Phân tích theo phạm vi Creative; không cộng trùng dữ liệu Campaign/Order.
def prep_deep(df):
    if df is None:return None
    d=df.copy()
    for c in ['Chi phí','Doanh thu gộp','Số lượng đơn hàng SKU','Số lượt hiển thị quảng cáo sản phẩm','Số lượt nhấp vào quảng cáo sản phẩm']:
        d[c]=num(d[c]) if c in d else 0.0
    d['Nhóm nội dung']=d.get('Loại nội dung sáng tạo',pd.Series('',index=d.index)).fillna('').astype(str)
    d['CTR (%)']=np.where(d['Số lượt hiển thị quảng cáo sản phẩm']>0,100*d['Số lượt nhấp vào quảng cáo sản phẩm']/d['Số lượt hiển thị quảng cáo sản phẩm'],np.nan)
    d['ROAS']=np.where(d['Chi phí']>0,d['Doanh thu gộp']/d['Chi phí'],np.nan)
    d['SKU/Click (%)']=np.where(d['Số lượt nhấp vào quảng cáo sản phẩm']>0,100*d['Số lượng đơn hàng SKU']/d['Số lượt nhấp vào quảng cáo sản phẩm'],np.nan)
    return d

def pct_col(df,raw):
    if raw not in df:return pd.Series(np.nan,index=df.index)
    ss=df[raw].astype(str).str.strip(); marked=ss.str.contains('%',regex=False)
    val=pd.to_numeric(ss.str.replace('%','',regex=False).str.replace(',','.',regex=False),errors='coerce')
    return pd.Series(np.where(marked,val,np.where(val.abs()<=1,val*100,val)),index=df.index)

def comment_video(r,ctr_cut,ret_cut):
    ctr=r['CTR (%)']; ret=r['Xem 6s (%)']; sku=r['Số lượng đơn hàng SKU']; cost=r['Chi phí']; roas=r['ROAS']
    good=[]; weak=[]; hypotheses=[]
    if pd.notna(ret):
        (good if ret>=ret_cut else weak).append(f'Giữ chân 6 giây {ret:.2f}% '+('từ mức tham chiếu trở lên' if ret>=ret_cut else 'dưới mức tham chiếu'))
        if ret<ret_cut:hypotheses.append(f'Giữ chân 6s {ret:.2f}% dưới mốc {ret_cut:.2f}%: xem lại 0–3s có nêu vấn đề cụ thể không; kiểm tra thời điểm xuất hiện sản phẩm và nhịp cắt 3–6s')
    if pd.notna(ctr):
        (good if ctr>=ctr_cut else weak).append(f'CTR {ctr:.2f}% '+('từ mức tham chiếu trở lên' if ctr>=ctr_cut else 'dưới mức tham chiếu'))
        if ctr<ctr_cut:hypotheses.append(f'CTR {ctr:.2f}% dưới mốc {ctr_cut:.2f}%: thử đưa công dụng có căn cứ và tình huống sử dụng trước CTA; kiểm tra nhấp theo phiên bản hook')
    if sku>0:good.append(f'Có {sku:g} SKU quy thuộc trong báo cáo Creative')
    elif cost>0:
        weak.append(f'Đã tiêu {vnd(cost)} nhưng chưa ghi nhận SKU')
        hypotheses.append(f'Chi {vnd(cost)} chưa có SKU: đối chiếu {int(r.get("Số lượt nhấp vào quảng cáo sản phẩm",0))} nhấp với trang đích, phí giao hàng, ưu đãi và đánh giá; thử thay từng yếu tố')
    if pd.notna(roas) and roas<1:weak.append(f'Doanh thu quy thuộc thấp hơn chi phí (ROAS {roas:.2f})')
    if pd.isna(ctr) or pd.isna(ret):hypotheses.append('Thiếu CTR hoặc tỷ lệ xem 6 giây: chưa thể phân loại đủ hai chiều')
    if not hypotheses:hypotheses.append(f'Đã có {sku:g} SKU, ROAS {show_roas(roas)}: giữ cấu trúc video làm đối chứng; thử một biến thể hook/CTA và đối soát đơn hủy trước khi mở rộng')
    return ' • '.join(good) or 'Chưa có tín hiệu tích cực đủ rõ', ' • '.join(weak) or 'Chưa phát hiện điểm yếu rõ từ các chỉ số có sẵn', ' • '.join(dict.fromkeys(hypotheses))

def comment_card(r,ctr_cut):
    good=[];weak=[];hyp=[]
    if pd.notna(r['CTR (%)']):
        (good if r['CTR (%)']>=ctr_cut else weak).append(f'CTR {r["CTR (%)"]:.2f}% '+('từ mức tham chiếu trở lên' if r['CTR (%)']>=ctr_cut else 'dưới mức tham chiếu'))
        if r['CTR (%)']<ctr_cut:hyp.append(f'CTR {r["CTR (%)"]:.2f}% thấp hơn mốc {ctr_cut:.2f}% trên {r["Hiển thị"]:,.0f} hiển thị: A/B ảnh chính, dòng tên đầu và thông tin giá; so CTR trên cùng thời gian')
    if r['SKU']>0:good.append(f'Ghi nhận {r["SKU"]:g} SKU bán được')
    elif r['Chi phí']>0:
        weak.append(f'Tiêu {vnd(r["Chi phí"])} nhưng chưa ghi nhận SKU')
        hyp.append(f'Đã có {r["Nhấp"]:,.0f} nhấp nhưng 0 SKU: kiểm tra giá sau voucher, phí ship, đánh giá và thông tin mô tả; thử chỉnh từng yếu tố')
    if pd.notna(r['ROAS']):
        if r['ROAS']>=1:good.append(f'Doanh thu quy thuộc lớn hơn/ bằng chi phí Ads (ROAS {r["ROAS"]:.2f})')
        else:weak.append(f'Doanh thu thấp hơn chi phí Ads (ROAS {r["ROAS"]:.2f})')
    if r['Nhấp']>0 and r['SKU']==0:hyp.append('Ưu tiên đo rơi rụng sau nhấp và số thêm giỏ; không quy nguyên nhân cho ảnh thẻ khi chưa có dữ liệu trang đích')
    if not hyp:hyp.append(f'ROAS {show_roas(r["ROAS"])} với {r["SKU"]:g} SKU: đối chiếu đơn hủy và thử biến thể ảnh/ưu đãi trong ngân sách nhỏ, giữ bản hiện tại làm đối chứng')
    return ' • '.join(good) or 'Chưa có tín hiệu tích cực đủ rõ',' • '.join(weak) or 'Chưa phát hiện điểm yếu rõ từ các chỉ số có sẵn',' • '.join(dict.fromkeys(hyp))

deep=prep_deep(creative) if active in (7,8,9) else None
if active == 7:
    st.header('08 / Phân tích chuyên sâu VIDEO')
    st.caption('Phân tích riêng từng dòng video trong Creative • Giữ chân 6 giây và CTR có mẫu số khác nhau • Giả thuyết không phải kết luận về chất lượng hình ảnh/lời thoại')
    if deep is None:st.info('Hãy tải Creative data để xem phân tích video.')
    else:
        videos=deep[deep['Nhóm nội dung'].str.contains('video',case=False,na=False)].copy()
        if videos.empty:st.info('Không có video trong dữ liệu Creative.')
        else:
            videos['Xem 2s (%)']=pct_col(videos,'Tỷ lệ xem video quảng cáo trong 2 giây')
            videos['Xem 6s (%)']=pct_col(videos,'Tỷ lệ xem video quảng cáo trong 6 giây')
            videos['Mẫu quảng cáo']=videos.get('Mẫu quảng cáo',pd.Series('',index=videos.index)).fillna('').astype(str)
            videos['Tên sản phẩm']=videos.get('Tên sản phẩm',pd.Series('',index=videos.index)).fillna('').astype(str)
            video_options=list(videos.index)
            selected=st.selectbox('Chọn video để phân tích',video_options,format_func=lambda i:f'{videos.loc[i,"Mẫu quảng cáo"][:105]} · {videos.loc[i,"Tên sản phẩm"][:45]}',key='deep_video')
            row=videos.loc[selected]
            st.subheader('Phân tích số liệu video được chọn')
            metrics=[('Chi phí',vnd(row['Chi phí'])),('Doanh thu',vnd(row['Doanh thu gộp'])),('ROAS',show_roas(row['ROAS'])),('Đơn SKU',f'{row["Số lượng đơn hàng SKU"]:g}'),('CTR',f'{row["CTR (%)"]:.2f}%' if pd.notna(row['CTR (%)']) else '—'),('Xem 6 giây',f'{row["Xem 6s (%)"]:.2f}%' if pd.notna(row['Xem 6s (%)']) else '—')]
            cols=st.columns(3)
            for i,(label,value) in enumerate(metrics):cols[i%3].metric(label,value)
            st.caption(f'Hiển thị: {row["Số lượt hiển thị quảng cáo sản phẩm"]:,.0f} · Nhấp: {row["Số lượt nhấp vào quảng cáo sản phẩm"]:,.0f} · CVR SKU/Click: {row["SKU/Click (%)"]:.2f}%' if pd.notna(row['SKU/Click (%)']) else 'Không đủ dữ liệu CVR')
            valid=videos[(videos['Xem 6s (%)'].notna())&(videos['CTR (%)'].notna())].copy()
            st.subheader('Ma trận 4 vùng: giữ chân 6 giây × CTR')
            if valid.empty:st.warning('Không đủ dữ liệu giữ chân 6 giây và CTR để tạo ma trận.')
            else:
                base_ret=float(valid['Xem 6s (%)'].median());base_ctr=float(valid['CTR (%)'].median())
                l,r=st.columns(2)
                with l:ret_cut=st.number_input('Mốc giữ chân 6s (%)',min_value=0.0,max_value=100.0,value=round(base_ret,2),step=0.1,help='Mặc định là trung vị các video có dữ liệu; không phải tiêu chuẩn TikTok')
                with r:ctr_cut=st.number_input('Mốc CTR (%)',min_value=0.0,max_value=100.0,value=round(base_ctr,2),step=0.1,help='Mặc định là trung vị các video có dữ liệu; không phải tiêu chuẩn TikTok')
                def region(row):
                    return ('Giữ chân cao' if row['Xem 6s (%)']>=ret_cut else 'Giữ chân thấp')+' – '+('CTR cao' if row['CTR (%)']>=ctr_cut else 'CTR thấp')
                valid['Vùng']=valid.apply(region,axis=1)
                maxx=max(10.0,float(valid['Xem 6s (%)'].max())*1.13,ret_cut*1.15)
                maxy=max(2.0,float(valid['CTR (%)'].max())*1.13,ctr_cut*1.15)
                fig=go.Figure()
                blocks=[(0,ret_cut,0,ctr_cut,'#f8cbd1','Giữ chân thấp / CTR thấp'),(ret_cut,maxx,0,ctr_cut,'#cbdff6','Giữ chân cao / CTR thấp'),(0,ret_cut,ctr_cut,maxy,'#f9df9b','Giữ chân thấp / CTR cao'),(ret_cut,maxx,ctr_cut,maxy,'#b8e7d6','Giữ chân cao / CTR cao')]
                for x0,x1,y0,y1,color,name in blocks:
                    fig.add_shape(type='rect',x0=x0,x1=x1,y0=y0,y1=y1,fillcolor=color,opacity=.76,line_width=0,layer='below')
                    if x1>x0 and y1>y0:fig.add_annotation(x=(x0+x1)/2,y=(y0+y1)/2,text=name.replace(' / ','<br>'),showarrow=False,font=dict(size=13,color='#1b3952'),opacity=.95)
                fig.add_vline(x=ret_cut,line_dash='dash',line_color='#334e68')
                fig.add_hline(y=ctr_cut,line_dash='dash',line_color='#334e68')
                fig.add_trace(go.Scatter(x=valid['Xem 6s (%)'],y=valid['CTR (%)'],mode='markers',text=valid['Mẫu quảng cáo'],customdata=np.stack([valid['Chi phí'],valid['Số lượng đơn hàng SKU'],valid['ROAS'].fillna(-1)],axis=-1),marker=dict(size=np.clip(9+np.sqrt(valid['Chi phí'].clip(lower=0))/35,9,34),color=valid['Số lượng đơn hàng SKU'],colorscale=[[0,'#a5d8ed'],[0.5,'#2488bf'],[1,'#073b70']],showscale=True,colorbar=dict(title='SKU'),line=dict(color='white',width=1)),hovertemplate='<b>%{text}</b><br>Xem 6s: %{x:.2f}%<br>CTR: %{y:.2f}%<br>Chi phí: %{customdata[0]:,.0f}đ<br>SKU: %{customdata[1]:.0f}<extra></extra>'))
                fig.update_layout(height=550,margin=dict(l=28,r=20,t=40,b=40),xaxis=dict(title='Giữ chân 6 giây (%)',range=[0,maxx]),yaxis=dict(title='CTR sản phẩm (%)',range=[0,maxy]),showlegend=False,plot_bgcolor='white',paper_bgcolor='white')
                st.plotly_chart(style_fig(fig),use_container_width=True)
                st.caption('Đường chia dùng trung vị của tập video hiện tại hoặc mốc tự điều chỉnh. Kích thước chấm thể hiện chi phí; màu thể hiện SKU. Vùng cao/thấp chỉ mang tính tương đối, không phải chuẩn ngành.')
                count=valid['Vùng'].value_counts()
                for name,meaning in [('Giữ chân cao – CTR cao','Nội dung có tín hiệu giữ người xem và thúc đẩy nhấp; kiểm tra thêm SKU, chi phí và đơn hủy.'),('Giữ chân cao – CTR thấp','Nội dung giữ người xem nhưng chưa thúc đẩy nhấp; thử làm rõ giá trị mua và CTA.'),('Giữ chân thấp – CTR cao','Người xem có thể nhấp sớm; không kết luận video kém chỉ vì thời gian xem thấp.'),('Giữ chân thấp – CTR thấp','Ưu tiên kiểm tra hook, cấu trúc nội dung và lời mời nhấp.')]:
                    st.markdown(f'**{name} ({count.get(name,0)} video):** {meaning}')
            st.subheader('Nhận xét video được chọn')
            cut_ret=ret_cut if not valid.empty else 8.0
            cut_ctr=ctr_cut if not valid.empty else 1.0
            good,weak,hyp=comment_video(row,cut_ctr,cut_ret)
            a,b=st.columns(2)
            with a:st.success('**Làm tốt:** '+good)
            with b:st.warning('**Chưa tốt:** '+weak)
            st.info('**Giả thuyết cần kiểm chứng:** '+hyp)
            st.markdown('**Hướng xử lý:** '+('Thử 2 cách mở đầu khác nhau và so sánh giữ chân 6 giây.' if pd.notna(row['Xem 6s (%)']) and row['Xem 6s (%)']<cut_ret else 'Giữ video làm đối chứng; thử CTA hoặc trang đích, theo dõi SKU và ROAS.'))
            if len(videos):
                zones=videos['Vùng'].value_counts()
                poor=int((videos['Số lượng đơn hàng SKU']==0).sum())
                analysis(f'{len(videos)} video được phân tích; {poor} video chưa ghi nhận SKU. Giữ chân 6s và CTR chia theo trung vị của tập dữ liệu (không phải ngưỡng chuẩn). Nhóm đông nhất: {zones.index[0]} ({zones.iloc[0]} video). Ưu tiên kiểm tra các video chi phí cao, không SKU; đánh giá lại sau khi thử hook, CTA và đối soát đơn hủy.')

if active == 8:
    st.header('09 / Phân tích chuyên sâu THẺ SẢN PHẨM')
    st.caption('Mỗi sản phẩm có nhận xét riêng theo chi phí, CTR, nhấp, SKU và ROAS • Giả thuyết ảnh/giá/trang sản phẩm cần được xác minh bằng dữ liệu bổ sung')
    if deep is None:st.info('Hãy tải Creative data để xem phân tích thẻ sản phẩm.')
    else:
        cards=deep[deep['Nhóm nội dung'].str.contains('thẻ',case=False,na=False)].copy()
        if cards.empty:st.info('Không tìm thấy thẻ sản phẩm trong Creative data.')
        else:
            cards['Sản phẩm']=cards.get('Tên sản phẩm',pd.Series('Không rõ',index=cards.index)).fillna('Không rõ').astype(str)
            imp='Số lượt hiển thị quảng cáo sản phẩm';click='Số lượt nhấp vào quảng cáo sản phẩm'
            ct=cards.groupby('Sản phẩm',as_index=False).agg(**{'Chi phí':('Chi phí','sum'),'Doanh thu':('Doanh thu gộp','sum'),'SKU':('Số lượng đơn hàng SKU','sum'),'Hiển thị':(imp,'sum'),'Nhấp':(click,'sum')})
            ct['CTR (%)']=np.where(ct['Hiển thị']>0,100*ct['Nhấp']/ct['Hiển thị'],np.nan)
            ct['SKU/Click (%)']=np.where(ct['Nhấp']>0,100*ct['SKU']/ct['Nhấp'],np.nan)
            ct['ROAS']=np.where(ct['Chi phí']>0,ct['Doanh thu']/ct['Chi phí'],np.nan)
            ct=ct.sort_values('Chi phí',ascending=False)
            selected=st.selectbox('Chọn sản phẩm để phân tích',ct['Sản phẩm'].tolist(),key='deep_card')
            row=ct.loc[ct['Sản phẩm'].eq(selected)].iloc[0]
            st.subheader('Phân tích số liệu thẻ sản phẩm')
            items=[('Chi phí',vnd(row['Chi phí'])),('Doanh thu',vnd(row['Doanh thu'])),('ROAS',show_roas(row['ROAS'])),('Đơn SKU',f'{row["SKU"]:g}'),('CTR',f'{row["CTR (%)"]:.2f}%' if pd.notna(row['CTR (%)']) else '—'),('CVR SKU/Click',f'{row["SKU/Click (%)"]:.2f}%' if pd.notna(row['SKU/Click (%)']) else '—')]
            cols=st.columns(3)
            for i,(label,value) in enumerate(items):cols[i%3].metric(label,value)
            st.caption(f'Hiển thị: {row["Hiển thị"]:,.0f} · Nhấp: {row["Nhấp"]:,.0f} · Chi phí/SKU: {vnd(row["Chi phí"]/row["SKU"]) if row["SKU"]>0 else "—"}')
            default=float(ct['CTR (%)'].median()) if ct['CTR (%)'].notna().any() else 1.0
            threshold=default
            good,weak,hyp=comment_card(row,threshold)
            st.subheader('Nhận xét thẻ sản phẩm được chọn')
            a,b=st.columns(2)
            with a:st.success('**Làm tốt:** '+good)
            with b:st.warning('**Chưa tốt:** '+weak)
            st.info('**Giả thuyết cần kiểm chứng:** '+hyp)
            st.markdown('**Hướng xử lý:** '+('Ưu tiên thử ảnh đại diện và phần đầu tên sản phẩm; đo CTR trên cùng thời gian.' if pd.notna(row['CTR (%)']) and row['CTR (%)']<threshold else 'Rà soát giá cuối cùng, voucher, phí giao hàng, đánh giá và trang chi tiết; thử từng yếu tố.'))
            st.caption('Mốc CTR tham chiếu là trung vị các thẻ trong file, không phải chuẩn ngành. Không kết luận nguyên nhân ảnh/giá nếu chưa kiểm tra thực tế.')

# 10 / Hành động theo từng sản phẩm: hai nhánh độc lập video và thẻ sản phẩm.
if active == 9:
    st.header('10 / HÀNH ĐỘNG THEO TỪNG SẢN PHẨM')
    st.caption('Kế hoạch tách riêng VIDEO và THẺ SẢN PHẨM • Đề xuất dựa trên dữ liệu Creative đã tải • Không sử dụng ROAS hòa vốn')
    if deep is None or deep.empty:
        st.info('Tải Creative data để tạo danh sách hành động theo sản phẩm.')
    else:
        src=deep.copy()
        src['Sản phẩm']=src.get('Tên sản phẩm',pd.Series('Không rõ',index=src.index)).fillna('Không rõ').astype(str).str.strip()
        src.loc[src['Sản phẩm'].eq(''),'Sản phẩm']='Không rõ'
        src['Loại']=np.where(src['Nhóm nội dung'].str.contains('video',case=False,na=False),'Video',
                     np.where(src['Nhóm nội dung'].str.contains('thẻ',case=False,na=False),'Thẻ sản phẩm','Khác'))
        src=src[src['Loại'].isin(['Video','Thẻ sản phẩm'])]
        if src.empty:
            st.warning('Không xác định được Video/Thẻ sản phẩm từ cột Loại nội dung sáng tạo.')
        else:
            src['Xem 6s (%)']=pct_col(src,'Tỷ lệ xem video quảng cáo trong 6 giây')
            imp='Số lượt hiển thị quảng cáo sản phẩm'; clicks='Số lượt nhấp vào quảng cáo sản phẩm'
            agg=src.groupby(['Sản phẩm','Loại'],as_index=False).agg(
                **{'Chi phí':('Chi phí','sum'),'Doanh thu':('Doanh thu gộp','sum'),
                   'SKU':('Số lượng đơn hàng SKU','sum'),'Hiển thị':(imp,'sum'),
                   'Nhấp':(clicks,'sum'),'Số creative':('Loại','size')})
            agg['CTR (%)']=np.where(agg['Hiển thị']>0,agg['Nhấp']/agg['Hiển thị']*100,np.nan)
            agg['SKU/Click (%)']=np.where(agg['Nhấp']>0,agg['SKU']/agg['Nhấp']*100,np.nan)
            agg['ROAS']=np.where(agg['Chi phí']>0,agg['Doanh thu']/agg['Chi phí'],np.nan)
            video_ret=src[src['Loại'].eq('Video')].groupby('Sản phẩm')['Xem 6s (%)'].median()
            agg['Giữ chân 6s trung vị (%)']=agg['Sản phẩm'].map(video_ret).where(agg['Loại'].eq('Video'))
            thresholds={kind:float(g['CTR (%)'].median()) if g['CTR (%)'].notna().any() else 1.0 for kind,g in agg.groupby('Loại')}
            ret_ref=float(src.loc[src['Loại'].eq('Video'),'Xem 6s (%)'].median())
            if not np.isfinite(ret_ref):ret_ref=8.0
            def action_for(r):
                kind=r['Loại']; cost=r['Chi phí']; sku=r['SKU']; ctr=r['CTR (%)']; roas=r['ROAS']; click=r['Nhấp']
                if kind=='Video':
                    positives=[]; issues=[]; hypotheses=[]; steps=[]
                    ret=r['Giữ chân 6s trung vị (%)']
                    if np.isfinite(ret):
                        (positives if ret>=ret_ref else issues).append(f'Giữ chân 6s {ret:.2f}% (mốc tham chiếu {ret_ref:.2f}%)')
                        if ret<ret_ref:
                            hypotheses.append('Đoạn mở đầu hoặc nhịp dựng 3–6 giây có thể chưa đủ hấp dẫn')
                            steps.append('Làm 2 hook khác nhau, kiểm tra tỷ lệ xem 2s và 6s trên từng video')
                    if np.isfinite(ctr):
                        (positives if ctr>=thresholds[kind] else issues).append(f'CTR {ctr:.2f}% (mốc {thresholds[kind]:.2f}%)')
                        if ctr<thresholds[kind]:
                            hypotheses.append('Lời mời xem sản phẩm hoặc tình huống sử dụng chưa tạo động lực nhấp')
                            steps.append('Thử CTA cụ thể và đưa tình huống sử dụng rõ hơn, giữ nguyên các yếu tố khác')
                    if sku>0:positives.append(f'{sku:g} SKU quy thuộc video')
                    if sku==0 and cost>0:
                        issues.append(f'Tiêu {vnd(cost)} nhưng 0 SKU')
                        hypotheses.append('Lượt xem/nhấp chưa chuyển thành mua; chưa thể quy lỗi cho video')
                        steps.append('Đối chiếu số nhấp với đơn và trang sản phẩm; chưa tăng ngân sách video')
                    if np.isfinite(roas) and roas<1:issues.append(f'ROAS {roas:.2f} dưới 1')
                    if not steps:steps.append('Giữ video hiện tại làm đối chứng; thử một biến thể hook hoặc CTA và theo dõi SKU, ROAS, đơn hủy')
                else:
                    positives=[];issues=[];hypotheses=[];steps=[]
                    if np.isfinite(ctr):
                        (positives if ctr>=thresholds[kind] else issues).append(f'CTR {ctr:.2f}% (mốc {thresholds[kind]:.2f}%)')
                        if ctr<thresholds[kind]:
                            hypotheses.append('Ảnh chính, tiêu đề hoặc giá hiển thị có thể chưa thúc đẩy nhấp')
                            steps.append('A/B ảnh chính hoặc 40 ký tự đầu tên sản phẩm; so CTR trên cùng khoảng thời gian')
                    if sku>0:positives.append(f'{sku:g} SKU quy thuộc thẻ')
                    if sku==0 and cost>0:
                        issues.append(f'Tiêu {vnd(cost)} nhưng 0 SKU')
                        hypotheses.append('Trang sản phẩm, giá sau ưu đãi, phí ship hoặc đánh giá có thể làm giảm mua sau nhấp')
                        steps.append('Rà soát giá cuối cùng, phí ship, đánh giá và mô tả; thử một thay đổi mỗi lần')
                    if np.isfinite(roas) and roas<1:
                        issues.append(f'ROAS {roas:.2f} dưới 1')
                        if sku>0:steps.append('Giảm thử nghiệm ngân sách thẻ, kiểm tra chi phí/SKU và đơn hủy trước khi mở rộng')
                    if not steps:steps.append('Giữ thẻ hiện tại làm đối chứng; thử một biến thể ảnh hoặc ưu đãi, theo dõi CTR và SKU/Click')
                if not hypotheses:hypotheses.append('Chưa có tín hiệu yếu nổi bật; kiểm tra tính ổn định khi tăng số lượt nhấp')
                priority='Ưu tiên cao' if cost>0 and (sku==0 or (np.isfinite(roas) and roas<1)) else 'Theo dõi / thử nghiệm'
                return pd.Series({'Làm tốt':' • '.join(positives) or 'Chưa có bằng chứng rõ',
                                  'Chưa tốt':' • '.join(issues) or 'Chưa thấy vấn đề nổi bật theo dữ liệu',
                                  'Giả thuyết độc lập':' • '.join(dict.fromkeys(hypotheses)),
                                  'Hành động đề xuất':' • '.join(dict.fromkeys(steps)),
                                  'Mức ưu tiên':priority})
            agg=pd.concat([agg,agg.apply(action_for,axis=1)],axis=1)
            products=sorted(agg['Sản phẩm'].unique())
            selected=st.selectbox('Chọn sản phẩm để xem kế hoạch VIDEO + THẺ',products,key='action_product')
            chosen=agg[agg['Sản phẩm'].eq(selected)]
            st.subheader(selected)
            cols=st.columns(4)
            cols[0].metric('Chi phí Creative',vnd(chosen['Chi phí'].sum()))
            cols[1].metric('SKU quy thuộc',f"{chosen['SKU'].sum():g}")
            cols[2].metric('ROAS Creative',show_roas(ratio(chosen['Doanh thu'].sum(),chosen['Chi phí'].sum())))
            cols[3].metric('SKU/Click',f"{100*ratio(chosen['SKU'].sum(),chosen['Nhấp'].sum()):.2f}%" if chosen['Nhấp'].sum()>0 else '—')
            left,right=st.columns(2)
            for area,kind in [(left,'Video'),(right,'Thẻ sản phẩm')]:
                with area:
                    st.markdown(f'### {"🎥" if kind=="Video" else "🛍️"} {kind}')
                    r=chosen[chosen['Loại'].eq(kind)]
                    if r.empty:
                        st.info('Không có dữ liệu Creative cho hình thức này; không suy đoán hiệu quả.')
                    else:
                        r=r.iloc[0]
                        st.caption(f"{vnd(r['Chi phí'])} chi phí · {r['Nhấp']:g} nhấp · {r['SKU']:g} SKU · ROAS {show_roas(r['ROAS'])}")
                        st.success('**Làm tốt:** '+r['Làm tốt'])
                        st.warning('**Chưa tốt:** '+r['Chưa tốt'])
                        st.info('**Giả thuyết riêng:** '+r['Giả thuyết độc lập'])
                        st.markdown('**Hành động:** '+r['Hành động đề xuất'])
                        st.caption('Ưu tiên: '+r['Mức ưu tiên'])
            st.divider()
            st.subheader('Danh sách hành động cho toàn bộ sản phẩm')
            st.dataframe(agg.sort_values(['Mức ưu tiên','Chi phí'],ascending=[True,False]),hide_index=True,use_container_width=True)
            st.download_button('Tải kế hoạch hành động từng sản phẩm (.csv)',agg.to_csv(index=False).encode('utf-8-sig'),file_name='hanh_dong_video_va_the_tung_san_pham.csv',mime='text/csv')
            analysis(f'Đã tổng hợp {len(products)} sản phẩm, {int(agg["Loại"].eq("Video").sum())} nhóm video và {int(agg["Loại"].eq("Thẻ sản phẩm").sum())} nhóm thẻ. Có {int((agg["Chi phí"].gt(0)&agg["SKU"].eq(0)).sum())} nhóm nội dung tiêu ngân sách nhưng chưa ghi nhận SKU, với tổng chi phí {vnd(agg.loc[agg["Chi phí"].gt(0)&agg["SKU"].eq(0),"Chi phí"].sum())}. Các giả thuyết của video và thẻ được xây dựng độc lập; không coi SKU/Click là tỷ lệ đơn hàng hoàn tất.')
