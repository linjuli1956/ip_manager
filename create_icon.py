#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
创建IP管理器图标（可爱二次元风）
- 圆角矩形底 + 粉紫渐变
- 可爱表情（大眼、腮红、小嘴）增强亲和力
- 以 1024x1024 为基准导出多分辨率
"""

from PIL import Image, ImageDraw, ImageFont, ImageFilter, ImageChops
import os


def _rounded_rectangle_mask(size: int, radius_ratio: float = 0.22) -> Image.Image:
    """生成圆角矩形蒙版（白色区域为可见）。radius_ratio 基于边长比例。"""
    radius = int(size * radius_ratio)
    mask = Image.new("L", (size, size), 0)
    draw = ImageDraw.Draw(mask)
    draw.rounded_rectangle([0, 0, size, size], radius=radius, fill=255)
    return mask


def _linear_gradient(size: int, top_color, bottom_color) -> Image.Image:
    """创建线性竖直渐变图层。"""
    grad = Image.new("RGB", (1, size), color=0)
    draw = ImageDraw.Draw(grad)
    for y in range(size):
        ratio = y / (size - 1)
        r = int(top_color[0] * (1 - ratio) + bottom_color[0] * ratio)
        g = int(top_color[1] * (1 - ratio) + bottom_color[1] * ratio)
        b = int(top_color[2] * (1 - ratio) + bottom_color[2] * ratio)
        draw.point((0, y), fill=(r, g, b))
    return grad.resize((size, size))


def _add_inner_shadow(base: Image.Image, intensity: int = 60, blur: int = 20) -> Image.Image:
    """在圆角内添加轻微内阴影。"""
    size = base.size[0]
    shadow = Image.new("L", (size, size), 0)
    draw = ImageDraw.Draw(shadow)
    # 边缘一圈画深色边，之后高斯模糊，作为 alpha 叠加
    draw.rounded_rectangle([6, 6, size - 6, size - 6], radius=int(size * 0.22), outline=intensity, width=12)
    shadow = shadow.filter(ImageFilter.GaussianBlur(blur))
    rgba = base.convert("RGBA")
    r, g, b, a = rgba.split()
    # 将阴影作为 alpha 的反相轻微叠加
    shadow_inv = ImageChops.invert(shadow).point(lambda p: int(p * 0.12))
    new_a = ImageChops.darker(a, shadow_inv)
    rgba.putalpha(new_a)
    return rgba


def _add_top_gloss(base: Image.Image) -> Image.Image:
    """添加顶部柔和高光。"""
    size = base.size[0]
    gloss = Image.new("RGBA", (size, size), (255, 255, 255, 0))
    gdraw = ImageDraw.Draw(gloss)
    # 上半椭圆形高光
    bbox = [int(size * 0.05), int(size * 0.02), int(size * 0.95), int(size * 0.58)]
    gdraw.ellipse(bbox, fill=(255, 255, 255, 38))
    # 渐隐
    gloss = gloss.filter(ImageFilter.GaussianBlur(12))
    out = base.convert("RGBA")
    out.alpha_composite(gloss)
    return out


def create_kawaii_icon(base_size: int = 1024) -> Image.Image:
    """创建可爱二次元风图标。"""
    # 渐变颜色（粉紫）
    top = (255, 182, 193)      # LightPink
    bottom = (186, 170, 255)   # Soft Lavender

    # 渐变底 + 圆角裁切
    gradient = _linear_gradient(base_size, top, bottom)
    mask = _rounded_rectangle_mask(base_size, radius_ratio=0.23)
    rounded = Image.new("RGBA", (base_size, base_size))
    rounded.paste(gradient, (0, 0))
    rounded.putalpha(mask)

    # 轻微投影（外部阴影）
    shadow = Image.new("RGBA", (base_size, base_size), (0, 0, 0, 0))
    sdraw = ImageDraw.Draw(shadow)
    sdraw.rounded_rectangle([10, 10, base_size - 1, base_size - 1], radius=int(base_size * 0.23), fill=(0, 0, 0, 90))
    shadow = shadow.filter(ImageFilter.GaussianBlur(24))

    canvas = Image.new("RGBA", (base_size, base_size), (0, 0, 0, 0))
    canvas.alpha_composite(shadow, (0, 0))
    canvas.alpha_composite(rounded, (0, 0))

    # 顶部高光
    canvas = _add_top_gloss(canvas)

    # 绘制可爱表情
    draw = ImageDraw.Draw(canvas)
    cx, cy = base_size // 2, base_size // 2
    face_y = cy - int(base_size * 0.06)

    # 眼睛
    eye_r = int(base_size * 0.08)
    eye_offset_x = int(base_size * 0.18)
    left_eye = (cx - eye_offset_x - eye_r, face_y - eye_r,
                cx - eye_offset_x + eye_r, face_y + eye_r)
    right_eye = (cx + eye_offset_x - eye_r, face_y - eye_r,
                 cx + eye_offset_x + eye_r, face_y + eye_r)
    draw.ellipse(left_eye, fill=(255, 255, 255, 255))
    draw.ellipse(right_eye, fill=(255, 255, 255, 255))
    # 瞳孔
    pupil_r = int(eye_r * 0.55)
    def center_box(box, r):
        x0, y0, x1, y1 = box
        cxm, cym = (x0 + x1) // 2, (y0 + y1) // 2
        return (cxm - r, cym - r, cxm + r, cym + r)
    draw.ellipse(center_box(left_eye, pupil_r), fill=(60, 60, 90, 255))
    draw.ellipse(center_box(right_eye, pupil_r), fill=(60, 60, 90, 255))
    # 高光
    sparkle_r = int(pupil_r * 0.35)
    def offset_box(box, r, dx, dy):
        x0, y0, x1, y1 = box
        cxm, cym = (x0 + x1) // 2 + dx, (y0 + y1) // 2 + dy
        return (cxm - r, cym - r, cxm + r, cym + r)
    draw.ellipse(offset_box(center_box(left_eye, pupil_r), sparkle_r, -int(pupil_r*0.3), -int(pupil_r*0.3)), fill=(255,255,255,220))
    draw.ellipse(offset_box(center_box(right_eye, pupil_r), sparkle_r, -int(pupil_r*0.3), -int(pupil_r*0.3)), fill=(255,255,255,220))

    # 腮红
    blush_r = int(base_size * 0.05)
    blush_y = face_y + int(base_size * 0.10)
    draw.ellipse((cx - eye_offset_x - blush_r, blush_y - blush_r, cx - eye_offset_x + blush_r, blush_y + blush_r), fill=(255,120,140,140))
    draw.ellipse((cx + eye_offset_x - blush_r, blush_y - blush_r, cx + eye_offset_x + blush_r, blush_y + blush_r), fill=(255,120,140,140))

    # 小嘴（弧线）
    mouth_w = int(base_size * 0.16)
    mouth_h = int(base_size * 0.08)
    mx0, my0 = cx - mouth_w//2, blush_y - int(base_size*0.02)
    mx1, my1 = cx + mouth_w//2, my0 + mouth_h
    # Pillow uses start/end angles (degrees)
    draw.arc((mx0, my0, mx1, my1), start=200, end=340, fill=(90,40,60,255), width=int(base_size*0.02))

    return canvas


def create_ico_file():
    """创建ICO文件（多分辨率）。"""
    base = create_kawaii_icon(1024)
    sizes = [16, 32, 48, 64, 128, 256, 512]
    # 直接用大图按多尺寸导出，确保 ICO 内含多分辨率位图
    base.save("ip_manager.ico", format="ICO", sizes=[(s, s) for s in sizes])
    print("图标文件已创建: ip_manager.ico")


def create_png_files():
    """创建PNG文件（预览用）。"""
    base = create_kawaii_icon(1024)
    sizes = [16, 32, 48, 64, 128, 256, 512, 1024]
    for s in sizes:
        img = base.resize((s, s), Image.Resampling.LANCZOS)
        img.save(f"ip_manager_{s}x{s}.png", "PNG")
        print(f"已创建: ip_manager_{s}x{s}.png")


if __name__ == "__main__":
    print("正在创建可爱二次元风 IP 管理器图标...")
    try:
        create_ico_file()
        create_png_files()
        print("\n✅ 图标创建完成！")
        print("📁 生成的文件:")
        print("   - ip_manager.ico (Windows图标文件)")
        print("   - ip_manager_*.png (不同尺寸的PNG文件)")
        print("\n💡 使用方法:")
        print("   1. 将 ip_manager.ico 重命名为 IP管理器.ico")
        print("   2. 在构建exe时使用 --icon=IP管理器.ico 参数")
    except Exception as e:
        print(f"❌ 创建图标时出错: {e}")
        print("请确保已安装 Pillow 库: pip install Pillow") 