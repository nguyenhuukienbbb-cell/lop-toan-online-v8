LỚP TOÁN ONLINE V6.5
==================

Tài khoản giáo viên mặc định:
  Tài khoản: giaovien
  Mật khẩu: 123456

CÁC NÂNG CẤP V6
1. Tạo đề trực tiếp từ file Word .docx ngay trong trang Soạn đề.
2. Cách tính điểm mới:
   - Mỗi câu có TRỌNG SỐ (mặc định 1).
   - Mỗi bài có THANG ĐIỂM (mặc định 10).
   - Hệ thống tự quy đổi trọng số các câu về đúng thang điểm.
   - Trắc nghiệm chấm tự động; tự luận giáo viên chấm theo số điểm tối đa đã quy đổi.
3. Một bài có thể giao đồng thời cho nhiều lớp.
4. Câu hỏi có thể có ảnh minh họa.
5. Hỗ trợ công thức Toán bằng MathJax/LaTeX, ví dụ:
      $\\frac{1}{2}+\\frac{1}{3}$
      $$x^2 + 2x + 1 = 0$$
6. Có thời gian BẮT ĐẦU và KẾT THÚC bài.
7. Word import có thể lấy ảnh gần câu hỏi và phần chữ của công thức Word.
8. Vẫn giữ: trộn câu, trộn đáp án, chống làm lại, xuất Excel, Google Login, PostgreSQL/Render.

CHẠY TRÊN WINDOWS
- Cài Python 3.11+ và nhớ tích Add python.exe to PATH.
- Giải nén thư mục.
- Chạy install_local.bat một lần.
- Sau đó chạy start_local.bat hoặc run.bat.
- Mở http://127.0.0.1:5000

LƯU Ý
V6 dùng cơ sở dữ liệu mới lop_toan_online_v6.db để tránh xung đột cấu trúc với các bản cũ.


V6.5: Bộ nhập Word mới render nguyên câu để giữ đúng MathType/Equation và che đáp án; trắc nghiệm được tự chấm.

V6.5: Bộ nhập Word nhận thêm đáp án gạch chân, bảng đáp án, đúng/sai từng câu, đúng/sai nhiều mệnh đề, dữ kiện chung, điền khuyết và trả lời ngắn. Gạch chân đáp án được loại bỏ khỏi ảnh học sinh để tránh lộ đáp án.
