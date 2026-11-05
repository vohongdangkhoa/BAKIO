"""BA Studio MVP - stdlib web server + python-docx export."""
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from pathlib import Path
import json, io, re
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

ROOT = Path(__file__).resolve().parent
FIELDS = [
 ('overview','Tổng quan dự án', [('project','Tên dự án'),('type','Loại sản phẩm'),('summary','Mô tả ý tưởng'),('problem','Vấn đề cần giải quyết'),('goals','Mục tiêu'),('success','Chỉ số thành công'),('scope_in','Phạm vi thực hiện'),('scope_out','Ngoài phạm vi')]),
 ('audience','Người dùng & nghiệp vụ', [('stakeholders','Các bên liên quan'),('actors','Vai trò/tác nhân'),('personas','Chân dung người dùng'),('journeys','Hành trình người dùng'),('rules','Quy tắc nghiệp vụ')]),
 ('features','Chức năng & dữ liệu', [('features','Chức năng chính'),('priorities','Ưu tiên MVP'),('data','Dữ liệu cần thu thập/hiển thị'),('integrations','Tích hợp ngoài')]),
 ('ux','Định hướng UX/UI', [('screens','Danh sách màn hình'),('navigation','Điều hướng và luồng màn hình'),('visual','Phong cách thị giác'),('brand','Màu sắc, logo, nhận diện'),('devices','Thiết bị và kích thước'),('accessibility','Tiếp cận & tương tác'),('states','Trạng thái lỗi/trống/tải dữ liệu'),('content','Ngôn ngữ và nội dung')]),
 ('process','Khảo sát và mô hình hóa AS-IS / TO-BE', [('interview','Biên bản khảo sát và nguồn yêu cầu'),('as_is','Quy trình hiện tại (AS-IS)'),('as_is_pain','Điểm nghẽn quy trình'),('to_be','Quy trình tương lai (TO-BE)'),('exceptions','Ngoại lệ nghiệp vụ')]),
 ('usecase','User Story và Use Case', [('stories','User Stories'),('use_cases','Use Case nghiệp vụ'),('uc_alternatives','Luồng thay thế'),('acceptance_cases','Kịch bản và tiêu chí chấp nhận')]),
 ('collaboration','Phối hợp kỹ thuật, SIT, UAT và đào tạo', [('handoff_questions','Câu hỏi từ Dev/Tester/UX'),('change_log','Nhật ký yêu cầu thay đổi'),('sit','Phối hợp kiểm thử SIT'),('uat_scenarios','Kịch bản UAT'),('training','Hướng dẫn và đào tạo người dùng'),('handoff_notes','Ghi chú bàn giao')]),
 ('delivery','Ràng buộc & xác nhận', [('nfr','Phi chức năng'),('constraints','Giới hạn, ngân sách, thời gian'),('assumptions','Giả định'),('open_questions','Câu hỏi còn mở'),('acceptance','Tiêu chí chấp nhận'),('approver','Người xác nhận')])]

class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs): super().__init__(*args, directory=str(ROOT), **kwargs)
    def do_POST(self):
        if self.path != '/api/export/docx': self.send_error(404);return
        try:
            size=int(self.headers.get('Content-Length','0'))
            if size>2_000_000: raise ValueError('Dữ liệu quá lớn')
            data=json.loads(self.rfile.read(size))
            if not isinstance(data,dict): raise ValueError('Dữ liệu không hợp lệ')
            output=build_doc(data)
            name=re.sub(r'[^a-zA-Z0-9_-]+','_',str(data.get('project','Du_an')))[:48] or 'Du_an'
            self.send_response(200)
            self.send_header('Content-Type','application/vnd.openxmlformats-officedocument.wordprocessingml.document')
            self.send_header('Content-Disposition',f'attachment; filename="BA_UXUI_{name}.docx"')
            self.send_header('Content-Length',str(len(output)))
            self.end_headers(); self.wfile.write(output)
        except Exception as ex:
            msg=json.dumps({'error':str(ex)},ensure_ascii=False).encode()
            self.send_response(400);self.send_header('Content-Type','application/json; charset=utf-8');self.end_headers();self.wfile.write(msg)

def build_doc(data):
    doc=Document(); sec=doc.sections[0];sec.top_margin=Inches(.8);sec.bottom_margin=Inches(.75)
    sec.left_margin=Inches(.85);sec.right_margin=Inches(.85)
    styles=doc.styles
    styles['Normal'].font.name='Arial';styles['Normal'].font.size=Pt(10);styles['Normal'].paragraph_format.space_after=Pt(6)
    for name,size,color in [('Title',24,(18,42,78)),('Heading 1',15,(18,42,78)),('Heading 2',11,(0,112,129))]:
        s=styles[name];s.font.name='Arial';s.font.size=Pt(size);s.font.bold=True;s.font.color.rgb=RGBColor(*color)
    p=doc.add_paragraph();p.style='Title';p.add_run('HỒ SƠ YÊU CẦU BÀN GIAO UX/UI')
    doc.add_paragraph('BA REQUIREMENTS & UX/UI DESIGN HANDOFF',style='Subtitle')
    t=doc.add_table(rows=0,cols=2);t.style='Light Shading Accent 1';t.alignment=WD_TABLE_ALIGNMENT.CENTER
    for a,b in [('Dự án',data.get('project','')),('Loại sản phẩm',data.get('type','')),('Phiên bản','1.0 – Hồ sơ nghiệp vụ BA'),('Ngày xuất',datetime.now().strftime('%d/%m/%Y')),('Trạng thái','BA/PO ĐÃ XÁC NHẬN TRÊN ỨNG DỤNG' if data.get('approvalConfirmed') else 'BẢN NHÁP – CHƯA XÁC NHẬN'),('Người xác nhận',data.get('approver','Chưa xác định'))]:
        cells=t.add_row().cells;cells[0].text=a;cells[1].text=b or 'CHƯA CUNG CẤP'
    doc.add_paragraph('Lưu ý: Các mục “CHƯA CUNG CẤP” là khoảng trống yêu cầu, không được coi là quyết định thiết kế đã xác nhận.')
    doc.add_heading('Mục đích tài liệu',1)
    doc.add_paragraph('Tổng hợp đầu vào nghiệp vụ để UX Researcher, UX Designer và UI Designer xây dựng user flow, information architecture, wireframe, prototype và bộ màn hình. Cần BA/PO xác nhận trước khi dùng làm chuẩn nghiệm thu.')
    for index,(key,section,items) in enumerate(FIELDS,1):
        doc.add_heading(f'{index}. {section}',1)
        for code,label in items:
            doc.add_heading(label,2)
            value=str(data.get(code,'')).strip()
            if not value: doc.add_paragraph('CHƯA CUNG CẤP — BA cần làm rõ và xác nhận.')
            elif '\n' in value:
                for line in value.splitlines():
                    if line.strip():doc.add_paragraph(line.strip(),style='List Bullet')
            else:doc.add_paragraph(value)
    doc.add_heading(f'{len(FIELDS)+1}. Danh sách màn hình phục vụ thiết kế',1)
    screens=[x.strip(' •-\t') for x in str(data.get('screens','')).splitlines() if x.strip()]
    if screens:
        tbl=doc.add_table(rows=1,cols=4);tbl.style='Light Shading Accent 1'
        for i,v in enumerate(['Mã màn','Tên màn hình','Thông tin/Thao tác chính','Trạng thái UX cần thiết']):tbl.rows[0].cells[i].text=v
        for i,s in enumerate(screens,1):
            row=tbl.add_row().cells
            for j,v in enumerate([f'SCR-{i:03d}',s,'BA bổ sung chi tiết','Mặc định / Trống / Lỗi / Đang tải / Thành công']):row[j].text=v
    else:doc.add_paragraph('CHƯA CUNG CẤP — chưa thể lập danh mục màn hình.')
    doc.add_heading(f'{len(FIELDS)+2}. Checklist bàn giao UX/UI',1)
    for x in ['Xác nhận tác nhân và mục tiêu từng nhóm người dùng','Xác nhận phạm vi MVP và mức độ ưu tiên chức năng','Vẽ sitemap, user flow từ nghiệp vụ được duyệt','Xây dựng wireframe từng màn hình','Định nghĩa UI states: loading, empty, error, success','Kiểm tra responsive, bàn phím, tương phản và truy cập','Review prototype với BA/PO, ghi nhận phản hồi','Chốt thiết kế và lưu phiên bản bàn giao']:
        doc.add_paragraph('☐ '+x)
    doc.add_heading(f'{len(FIELDS)+3}. Các mục còn thiếu & rủi ro bàn giao',1)
    missing=[label for _,_,items in FIELDS for code,label in items if not str(data.get(code,'')).strip()]
    if missing:
        for x in missing:doc.add_paragraph(x,style='List Bullet')
    else:doc.add_paragraph('Mọi trường đã có dữ liệu; cần đánh giá tính đúng đắn, tính đầy đủ và xác nhận độc lập.')
    doc.add_heading(f'{len(FIELDS)+4}. Phê duyệt',1)
    doc.add_paragraph('Ghi chú: Xác nhận trong ứng dụng là tự khai báo và không thay thế chữ ký/kiểm soát nội bộ.')
    doc.add_paragraph('Business Analyst: __________________   Ngày: __________')
    doc.add_paragraph('Product Owner / Khách hàng: __________   Ngày: __________')
    doc.add_paragraph('UX/UI Lead: _______________________   Ngày: __________')
    footer=sec.footer.paragraphs[0];footer.alignment=WD_ALIGN_PARAGRAPH.CENTER;footer.text='BA Requirements Studio • Tài liệu nội bộ • Xem trạng thái phê duyệt tại trang bìa'
    out=io.BytesIO();doc.save(out);return out.getvalue()

if __name__=='__main__':
    print('BA Studio running at http://localhost:8000')
    ThreadingHTTPServer(('127.0.0.1',8000),Handler).serve_forever()
