# Tạo lại index.html (bản web) từ tool_dong_bo/giao_dien.html + web/tools/index_body.html.
# Chạy sau khi sửa giao diện:  python web/tools/tao_index.py
import os
os.chdir(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".."))
g = open("tool_dong_bo/giao_dien.html", encoding="utf8").read()
head = g[:g.index("</style>")]
head = head.replace("<title>Đồng bộ chấm công lương</title>",
                    "<title>Đồng bộ chấm công lương</title>\n<meta name=\"description\" content=\"Gộp bảng chấm công, bảng lương nhiều biểu mẫu về một file Excel chuẩn. Chạy ngay trong trình duyệt.\">")
head += """
.khoi-dong{display:flex;gap:10px;align-items:center;margin:0 0 20px;color:var(--muc-nhat);font-size:14px}
.khoi-dong.xong{display:none}
.rieng-tu{font-size:13px;color:var(--xanh);margin:10px 0 0}
a.lk{color:var(--muc);font-weight:600}
"""
body = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "index_body.html"), encoding="utf8").read()
# phần vẽ kết quả: lấy nguyên từ giao_dien.html, chỉ đổi link tải sang blob
s = g[g.index("const MUC = ["):g.index("veFile();\n// mở lại trang")]
a = '''`<a class="nut ${i === 0 ? "chinh" : ""}" href="/tai?f=${encodeURIComponent(f)}">'''
assert a in s
s = s.replace(a, '''`<a class="nut ${i === 0 ? "chinh" : ""}" href="${urls[i]}" download="${esc(f)}">''')
body += s + "veFile();\n</script>\n</body>\n</html>\n"
open("index.html", "w", encoding="utf8", newline="\n").write(head + body)
open(".nojekyll", "w").close()
print("ok", len(head + body))
