import os
from paddleocr import PaddleOCR
from PIL import Image, ImageOps
import numpy as np

# 获取当前脚本所在目录
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 图片路径（相对于脚本目录，或使用绝对路径）
img_path = os.path.join(BASE_DIR, "TESTPIC.png")  # 可以修改为你的图片路径

# 检查图片是否存在
if not os.path.exists(img_path):
    print(f"错误：图片文件不存在: {img_path}")
    print(f"请将图片放在: {BASE_DIR}")
    exit(1)

# 读取图片
img = Image.open(img_path)

# 缩放图片（宽度固定为700px，高度按比例调整）
width = 700
height = int(img.height * (width / img.width))
img_resized = img.resize((width, height), Image.LANCZOS)

# 转换为灰度图 + 二值化
img_gray = ImageOps.grayscale(img_resized)
img_binary = img_gray.point(lambda x: 255 if x > 170 else 0)

# 将二值图像转换为3通道RGB格式
img = np.stack((img_binary,)*3, axis=-1)

# 模型路径配置（使用环境变量或默认路径）
PADDLE_MODELS_DIR = os.getenv('PADDLE_MODELS_DIR', 'D:/ProgramData/Paddle_models')
PP_OCR_REC_MODEL_DIR = os.path.join(PADDLE_MODELS_DIR, 'PP-OCRv4_mobile_rec')
PP_OCR_DET_MODEL_DIR = os.path.join(PADDLE_MODELS_DIR, 'PP-OCRv5_mobile_det')

# 检查模型目录是否存在，如果不存在则不指定路径（使用默认下载）
ocr_kwargs = {
    'precision': 'fp16',
    'enable_mkldnn': True,
    'cpu_threads': 2,
    'use_doc_orientation_classify': False,
    'use_doc_unwarping': False,
    'use_textline_orientation': False,
}

# 如果模型目录存在，则指定路径；否则指定模型名称让 PaddleOCR 自动下载
if os.path.exists(PP_OCR_REC_MODEL_DIR):
    ocr_kwargs['text_recognition_model_name'] = "PP-OCRv4_mobile_rec"
    ocr_kwargs['text_recognition_model_dir'] = PP_OCR_REC_MODEL_DIR
    print(f"使用本地识别模型: {PP_OCR_REC_MODEL_DIR}")
else:
    # 即使目录不存在，也指定模型名称，让 PaddleOCR 下载 mobile 版本而不是默认的 server 版本
    ocr_kwargs['text_recognition_model_name'] = "PP-OCRv4_mobile_rec"
    print(f"识别模型目录不存在，将自动下载: PP-OCRv4_mobile_rec")

if os.path.exists(PP_OCR_DET_MODEL_DIR):
    ocr_kwargs['text_detection_model_name'] = "PP-OCRv5_mobile_det"
    ocr_kwargs['text_detection_model_dir'] = PP_OCR_DET_MODEL_DIR
    print(f"使用本地检测模型: {PP_OCR_DET_MODEL_DIR}")
else:
    # 即使目录不存在，也指定模型名称，让 PaddleOCR 下载 mobile 版本
    ocr_kwargs['text_detection_model_name'] = "PP-OCRv5_mobile_det"
    print(f"检测模型目录不存在，将自动下载: PP-OCRv5_mobile_det")

# 实例化OCR
print("正在初始化 PaddleOCR...")
ocr = PaddleOCR(**ocr_kwargs)
print("PaddleOCR 初始化完成")

# 使用predict方法
print(f"开始识别图片: {img_path}")
result = ocr.predict(img)

# 解析结果
if result and len(result) > 0:
    print(f"\n识别结果类型: {type(result)}")
    print(f"结果键: {result[0].keys() if isinstance(result[0], dict) else 'N/A'}")
    
    if isinstance(result[0], dict) and 'rec_texts' in result[0]:
        print("\n识别到的文本:")
        for line in result[0]['rec_texts']:
            print(line)
    else:
        print(f"\n结果格式: {result}")
else:
    print("未识别到任何文本")
