"""将竖版 logo（图标在上、文字在下）重排为横向长条（图标在左、文字在右）"""
from PIL import Image
import os

script_dir = os.path.dirname(os.path.abspath(__file__))
# 源图：竖版 logo（图标在上，QuickJudge 文字在下）
src_path = r"C:\Users\65789\.cursor\projects\d-OneDrive-09\assets\c__Users_65789_AppData_Roaming_Cursor_User_workspaceStorage_a0ca1db52a468793a047a15377281f14_images_2-d8540789-17de-4373-8ad8-7db2618066b0.png"
out_path = os.path.join(script_dir, "logo_strip.png")

if not os.path.isfile(src_path):
    print("源图不存在，请把竖版 logo 放到 logo 文件夹并改名为 source_logo.png")
    src_path = os.path.join(script_dir, "source_logo.png")
    if not os.path.isfile(src_path):
        raise SystemExit("未找到源图")

img = Image.open(src_path).convert("RGBA")
w, h = img.size
# 按比例切分：上方为图标，下方为文字（约 55% : 45%）
split = int(h * 0.52)
icon_part = img.crop((0, 0, w, split))   # 上：图标
text_part = img.crop((0, split, w, h))    # 下：QuickJudge 文字

# 去掉文字区域上下多余白边
def trim_vertical_whitespace(im, threshold=250):
    pix = im.load()
    for y in range(im.height):
        for x in range(im.width):
            if pix[x, y][3] > 0 or (im.mode == "RGB" and sum(pix[x, y][:3]) < threshold):
                break
        else:
            continue
        break
    else:
        y = 0
    top = y
    for y in range(im.height - 1, -1, -1):
        for x in range(im.width):
            if pix[x, y][3] > 0 or (im.mode == "RGB" and sum(pix[x, y][:3]) < threshold):
                break
        else:
            continue
        break
    else:
        y = im.height - 1
    bottom = y + 1
    if bottom <= top:
        return im
    return im.crop((0, top, im.width, bottom))

icon_part = trim_vertical_whitespace(icon_part)
text_part = trim_vertical_whitespace(text_part)

gap = 24
strip_w = icon_part.width + gap + text_part.width
strip_h = max(icon_part.height, text_part.height) + 32
strip = Image.new("RGBA", (strip_w, strip_h), (255, 255, 255, 0))

# 图标左对齐并垂直居中
iy = (strip_h - icon_part.height) // 2
strip.paste(icon_part, (0, iy), icon_part)

# 文字在图标右侧，垂直居中
ty = (strip_h - text_part.height) // 2
strip.paste(text_part, (icon_part.width + gap, ty), text_part)

# 导出为 PNG（白底）
result = Image.new("RGB", strip.size, (255, 255, 255))
result.paste(strip, mask=strip.split()[3])
result.save(out_path, "PNG", optimize=True)
print("已生成长条 logo:", out_path)
