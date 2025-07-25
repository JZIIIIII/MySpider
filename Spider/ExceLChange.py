import sys
import os
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as ExcelImage
from io import BytesIO
from PIL import Image as PILImage

MAX_IMG_WIDTH = 80
MAX_IMG_HEIGHT = 80

class ExcelImageProcessor:
    def __init__(self, file_path, img_col_letter='L'):
        self.file_path = file_path
        self.img_col_letter = img_col_letter
        self.wb = None
        self.ws = None
        self.images = None

    def load_workbook(self):
        """加载 Excel 文件"""
        try:
            self.wb = load_workbook(self.file_path)
            self.ws = self.wb.active
            self.images = self.ws._images
            print(f"成功加载文件: {self.file_path}")
            return True
        except Exception as e:
            print(f"加载文件失败: {e}")
            return False

    def resize_image(self, img):
        """调整图片尺寸到最大宽高80*80，返回openpyxl的Image对象"""
        try:
            img_bytes = BytesIO(img._data())
            pil_img = PILImage.open(img_bytes)
            width, height = pil_img.size
            scale = min(MAX_IMG_WIDTH / width, MAX_IMG_HEIGHT / height, 1)
            new_size = (int(width * scale), int(height * scale))
            resized_img = pil_img.resize(new_size, PILImage.LANCZOS)

            output = BytesIO()
            resized_img.save(output, format='PNG')
            output.seek(0)

            return ExcelImage(output)
        except Exception as e:
            print(f"图片缩放失败: {e}")
            return img



    def process_excel(self):
        """处理 Excel 文件，按要求改写表头并保持图片，包含去重逻辑"""
        try:
            headers = [cell.value for cell in self.ws[1]]

            # 找到需要的列索引
            title_idx = headers.index("标题") if "标题" in headers else headers.index("Title")
            price_idx = headers.index("价格") if "价格" in headers else headers.index("Price")
            deal_idx = headers.index("成交量") if "成交量" in headers else headers.index("Deal")
            shop_url_idx = headers.index("店铺链接") if "店铺链接" in headers else headers.index("Title_URL")
            num_com_idx = headers.index("评论数量") if "评论数量" in headers else headers.index("Num_Com")
            shop_name_idx = headers.index("店铺名称") if "店铺名称" in headers else headers.index("Shop")
            is_post_free_idx = headers.index("包邮信息") if "包邮信息" in headers else headers.index("IsPostFree")
            tag_idx = headers.index("标签") if "标签" in headers else -1  # 没有则是空
            img_idx = headers.index("图片") if "图片" in headers else headers.index("Image")

            # 新表头
            new_headers = ["标题", "价格", "销量", "店铺链接", "店铺名称", "包邮信息", "标签", "评论数量", "图片"]

            # 新表格初始化
            new_wb = Workbook()
            new_ws = new_wb.active
            new_ws.append(new_headers)  # 写入新表头

            # 用来记录已经处理过的标题和店铺组合 (去重)
            seen = set()

            # 获取所有数据行
            data_rows = list(self.ws.iter_rows(min_row=2))
            filtered_rows = []

            for row in data_rows:
                title = row[title_idx].value
                shop = row[shop_name_idx].value
                key = (title, shop)

                # 如果标题和店铺组合已经存在，跳过
                if key in seen:
                    continue
                seen.add(key)

                # 处理这一行的数据
                new_row = [
                    row[title_idx].value,
                    row[price_idx].value,
                    row[deal_idx].value,
                    row[shop_url_idx].value,  # Shop_URL
                    row[shop_name_idx].value,
                    row[is_post_free_idx].value,
                    row[tag_idx].value if tag_idx != -1 else 'null',  # 标签，若无则是null
                    row[num_com_idx].value,
                ]

                # 保存图片并插入到 Excel
                img = self.images.pop(0) if len(self.images) > 0 else None
                if img:
                    new_img = self.resize_image(img)
                    anchor_cell = f"{get_column_letter(new_ws.max_column)}{new_ws.max_row + 1}"
                    new_img.anchor = anchor_cell
                    new_ws.add_image(new_img)

                    # 调整行高和列宽
                    new_ws.row_dimensions[new_ws.max_row].height = 60
                    new_ws.column_dimensions[get_column_letter(new_ws.max_column)].width = 14

                # 将数据添加到新行
                new_ws.append(new_row)

            output_path = os.path.splitext(self.file_path)[0] + "_去重保留图片.xlsx"
            new_wb.save(output_path)
            print(f"保存完成: {output_path}")
            return True

        except Exception as e:
            print(f"处理异常: {e}")
            return False

    def process(self):
        """处理 Excel 文件"""
        if not os.path.isfile(self.file_path):
            print("文件不存在:", self.file_path)
            return False

        ext = os.path.splitext(self.file_path)[1].lower()
        if ext not in [".xlsx", ".xlsm"]:
            print("仅支持 .xlsx 或 .xlsm 文件")
            return False

        success = self.process_excel()
        return success

def main():
    if len(sys.argv) < 2:
        print("请将 Excel 文件拖入此程序运行")
        input("按回车退出...")
        return

    file_path = sys.argv[1]
    processor = ExcelImageProcessor(file_path)

    if not processor.load_workbook():
        input("加载文件失败，按回车退出...")
        return

    success = processor.process()
    input("处理完成，按回车退出..." if success else "处理失败，按回车退出...")

if __name__ == "__main__":
    main()
