# PDF to Excel Converter

โปรเจค Python สำหรับแปลงไฟล์ PDF เป็น Excel (.xlsx) อย่างละเอียดและครบถ้วน

## 📋 คุณสมบัติหลัก (Features)

### 🔍 การแยกข้อมูล (Data Extraction)
- **ข้อความ (Text)**: แยกข้อความทั้งหมดจาก PDF รวมถึงการจัดรูปแบบ
- **ตาราง (Tables)**: ตรวจจับและแยกตารางอัตโนมัติ พร้อมรักษาโครงสร้าง
- **รูปภาพ (Images)**: แยกรูปภาพและบันทึกเป็นไฟล์แยก
- **เมตาดาต้า (Metadata)**: ข้อมูลเอกสารเช่น ผู้เขียน, วันที่สร้าง, หัวข้อ

### 📊 การส่งออก Excel
- **หลายชีต (Multiple Sheets)**: 
  - ชีต "Text" สำหรับข้อความทั้งหมด
  - ชีต "Tables" สำหรับตารางที่แยกได้
  - ชีต "Metadata" สำหรับข้อมูลเอกสาร
  - ชีต "Images_Info" สำหรับรายการรูปภาพ
- **การจัดรูปแบบ**: สีพื้นหลัง, ขอบตาราง, ฟอนต์
- **การปรับขนาดคอลัมน์อัตโนมัติ**

### 🔧 ตัวเลือกขั้นสูง
- **การประมวลผลหลายไฟล์**: แปลงไฟล์ PDF หลายไฟล์พร้อมกัน
- **การกรองเนื้อหา**: เลือกหน้าที่ต้องการแปลง
- **การจัดการข้อผิดพลาด**: รายงานข้อผิดพลาดอย่างละเอียด
- **แถบความคืบหน้า**: แสดงสถานะการประมวลผล

## 🛠️ เทคโนโลยีที่ใช้ (Technology Stack)

### ไลบรารีหลัก
- **PyPDF2/pdfplumber**: อ่านและแยกข้อมูลจาก PDF
- **tabula-py**: แยกตารางจาก PDF อย่างแม่นยำ
- **openpyxl**: สร้างและจัดรูปแบบไฟล์ Excel
- **Pillow (PIL)**: ประมวลผลรูปภาพ
- **pandas**: จัดการข้อมูลในรูปแบบ DataFrame

### ไลบรารีเสริม
- **tqdm**: แสดงแถบความคืบหน้า
- **python-magic**: ตรวจสอบประเภทไฟล์
- **camelot-py**: แยกตารางขั้นสูง (ทางเลือก)

## 📦 การติดตั้ง (Installation)

### 1. Clone Repository
```bash
git clone <repository-url>
cd pdf-to-excel
```

### 2. สร้าง Virtual Environment (แนะนำ)
```bash
python -m venv venv
source venv/bin/activate  # macOS/Linux
# หรือ
venv\Scripts\activate     # Windows
```

### 3. ติดตั้ง Dependencies
```bash
pip install -r requirements.txt
```

### 4. ติดตั้ง Java (สำหรับ tabula-py)
- **macOS**: `brew install openjdk`
- **Ubuntu/Debian**: `sudo apt-get install default-jdk`
- **Windows**: ดาวน์โหลดจาก Oracle JDK

## 🚀 การใช้งาน (Usage)

### การใช้งานพื้นฐาน
```python
from main import PDFToExcelConverter

# สร้าง converter
converter = PDFToExcelConverter()

# แปลงไฟล์เดียว
converter.convert_single_file("input.pdf", "output.xlsx")

# แปลงหลายไฟล์
pdf_files = ["file1.pdf", "file2.pdf", "file3.pdf"]
converter.convert_multiple_files(pdf_files, "output_directory/")
```

### การใช้งานขั้นสูง
```python
# กำหนดตัวเลือกเพิ่มเติม
converter = PDFToExcelConverter(
    extract_images=True,
    extract_tables=True,
    pages_range=(1, 5),  # แปลงเฉพาะหน้า 1-5
    output_format="detailed"
)

# แปลงพร้อมการกรอง
converter.convert_with_options(
    "input.pdf", 
    "output.xlsx",
    include_metadata=True,
    table_detection_method="camelot"
)
```

### Command Line Interface
```bash
# แปลงไฟล์เดียว
python main.py --input "document.pdf" --output "result.xlsx"

# แปลงหลายไฟล์
python main.py --batch --input-dir "pdf_files/" --output-dir "excel_files/"

# ตัวเลือกเพิ่มเติม
python main.py --input "doc.pdf" --output "result.xlsx" --extract-images --pages 1-10
```

## 📁 โครงสร้างโปรเจค (Project Structure)

```
pdf-to-excel/
├── main.py                 # ไฟล์หลัก - PDFToExcelConverter class
├── requirements.txt        # รายการ dependencies
├── README.md              # เอกสารนี้
├── utils/
│   ├── __init__.py
│   ├── pdf_reader.py      # ฟังก์ชันอ่าน PDF
│   ├── table_extractor.py # แยกตาราง
│   ├── image_extractor.py # แยกรูปภาพ
│   └── excel_writer.py    # เขียนไฟล์ Excel
├── tests/
│   ├── test_converter.py  # Unit tests
│   └── sample_files/      # ไฟล์ตัวอย่าง
├── output/                # โฟลเดอร์ผลลัพธ์
└── temp/                  # ไฟล์ชั่วคราว
```

## 🎯 ขั้นตอนการพัฒนา (Development Steps)

### Phase 1: Core Functionality ✅
- [x] สร้างโครงสร้างพื้นฐาน
- [x] อ่านข้อมูลจาก PDF
- [x] เขียนข้อมูลลง Excel
- [x] การจัดการข้อผิดพลาดพื้นฐาน

### Phase 2: Advanced Features 🔄
- [ ] แยกตารางอัตโนมัติ
- [ ] แยกรูปภาพ
- [ ] การประมวลผลหลายไฟล์
- [ ] แถบความคืบหน้า

### Phase 3: Enhancement 📋
- [ ] GUI Interface (Tkinter/PyQt)
- [ ] การกำหนดค่าผ่านไฟล์ config
- [ ] รองรับรูปแบบ PDF ที่ซับซ้อน
- [ ] การเพิ่มประสิทธิภาพ

### Phase 4: Testing & Documentation 🧪
- [ ] Unit Tests ครอบคลุม
- [ ] Integration Tests
- [ ] Performance Testing
- [ ] เอกสารการใช้งานเพิ่มเติม

## 📊 ตัวอย่างผลลัพธ์ (Sample Output)

### Excel Structure
```
Sheet 1: "Document_Text"
- Column A: Page Number
- Column B: Text Content
- Column C: Font Info (if available)

Sheet 2: "Extracted_Tables"
- ตารางที่แยกได้พร้อมโครงสร้างเดิม

Sheet 3: "Document_Metadata"
- Title, Author, Creation Date, etc.

Sheet 4: "Images_Information"
- Image filename, page number, size, format
```

## 🐛 การแก้ไขปัญหา (Troubleshooting)

### ปัญหาที่พบบ่อย
1. **Java ไม่ได้ติดตั้ง**: ติดตั้ง JDK สำหรับ tabula-py
2. **PDF ที่เข้ารหัส**: ต้องมีรหัสผ่านเพื่อเข้าถึง
3. **ตารางไม่ถูกตรวจจับ**: ลองใช้ camelot แทน tabula
4. **รูปภาพไม่สามารถแยกได้**: PDF อาจจะเป็นรูปแบบที่ไม่รองรับ

### การแก้ไข
```python
# สำหรับ PDF ที่เข้ารหัส
converter.set_password("your_password")

# สำหรับตารางที่ซับซ้อน
converter.set_table_extraction_method("camelot")

# สำหรับไฟล์ขนาดใหญ่
converter.set_memory_optimization(True)
```

## 🤝 การมีส่วนร่วม (Contributing)

1. Fork repository
2. สร้าง feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit การเปลี่ยนแปลง (`git commit -m 'Add some AmazingFeature'`)
4. Push ไปยัง branch (`git push origin feature/AmazingFeature`)
5. เปิด Pull Request

## 📄 License

MIT License - ดูรายละเอียดในไฟล์ `LICENSE`

## 📞 ติดต่อ (Contact)

- Email: your.email@example.com
- GitHub: @yourusername

---

**หมายเหตุ**: โปรเจคนี้อยู่ในขั้นตอนการพัฒนา ฟีเจอร์บางอย่างอาจยังไม่สมบูรณ์
# pdf2excel
