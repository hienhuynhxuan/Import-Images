import subprocess
import sys
import os

# Tự động cài thư viện nếu cần
def install_if_needed(package, import_name=None):
    import_name = import_name or package
    try:
        __import__(import_name)
    except ImportError:
        print(f"Installing package: {package} ...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])

# Cài đúng tên gói trên pip
install_if_needed("openpyxl")
install_if_needed("Pillow", "PIL")  # Gói trên pip là Pillow, import là PIL

# Nhập thư viện sau khi chắc chắn đã được cài
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from PIL import Image as PILImage  # Để đọc kích thước ảnh

def pixels_to_row_height(px):
    return px * 0.75  # Quy đổi pixel ➜ point (Excel)

def pixels_to_row_count(px):
    return int(pixels_to_row_height(px) / 15)  # mỗi dòng Excel ≈ 15pt

def import_images_to_excel(image_folder, output_excel):
    wb = Workbook()
    ws = wb.active
    ws.title = "Images"

    row = 1
    for filename in os.listdir(image_folder):
        if filename.lower().endswith(('.png', '.jpg', '.jpeg')):
            img_path = os.path.join(image_folder, filename)

            # Lấy kích thước ảnh
            with PILImage.open(img_path) as img_obj:
                width_px, height_px = img_obj.size

            img = ExcelImage(img_path)
            img.anchor = f'A{row}'
            ws.add_image(img)

            ws[f'B{row}'] = filename

            # Tính chiều cao dòng phù hợp với ảnh
            row_height = pixels_to_row_height(height_px)
            ws.row_dimensions[row].height = row_height

            # row += int(row_height // 15) + 2  # Nhảy dòng tránh đè ảnh
            
            # 👇 Nhảy xuống 1/2 chiều cao tương đương dòng Excel
            rows_needed = pixels_to_row_count(height_px)
            row += max(1, rows_needed // 2)  # tối thiểu nhảy 1 dòng để không bị đè
    wb.save(output_excel)
    print(f"✅ File Excel đã tạo tại: {output_excel}")

# Chạy chính
if __name__ == "__main__":
    import_images_to_excel("images", "output.xlsx")
