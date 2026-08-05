---
name: thumbnail-html-embed
description: HTML 中嵌入縮圖的最佳實踐 — 提供響應式縮圖、圖片優化、代碼範本與工具指引。當使用者說「嵌入縮圖」、「HTML 縮圖」、「圖片縮圖」、「響應式縮圖」、「優化圖片」、「圖片最佳化」、或提到 thumbnail、image embed、responsive images 時使用。
---

# HTML 中的縮圖嵌入指南

在網頁中高效展示縮略圖的完整指引，涵蓋響應式設計、性能優化與最佳實踐。

## 一、基礎 HTML 嵌入

### 簡單縮圖

```html
<img src="path/to/thumbnail.jpg" alt="縮圖描述" width="200" height="150" />
```

### 響應式縮圖（推薦）

使用 `srcset` 提供不同螢幕密度的圖片版本：

```html
<img 
  src="path/to/thumbnail-600w.jpg" 
  srcset="path/to/thumbnail-600w.jpg 600w, 
          path/to/thumbnail-900w.jpg 900w, 
          path/to/thumbnail-1200w.jpg 1200w"
  sizes="(max-width: 600px) 100vw, 
         (max-width: 900px) 50vw, 
         33vw"
  alt="縮圖描述" 
  loading="lazy" 
/>
```

## 二、現代方案 — `<picture>` 標籤

支持不同格式與螢幕尺寸的完整控制：

```html
<picture>
  <!-- WebP 格式（現代瀏覽器） -->
  <source 
    srcset="path/to/thumbnail.webp 1x, 
            path/to/thumbnail-2x.webp 2x"
    type="image/webp" 
  />
  
  <!-- JPG 後備方案 -->
  <source 
    srcset="path/to/thumbnail.jpg 1x, 
            path/to/thumbnail-2x.jpg 2x"
    type="image/jpeg" 
  />
  
  <!-- 最終後備 -->
  <img 
    src="path/to/thumbnail.jpg" 
    alt="縮圖描述" 
    loading="lazy"
  />
</picture>
```

## 三、效能最佳化

### 1. 圖片壓縮

```bash
# 使用 ImageMagick
convert original.jpg -resize 600x450 -quality 85 thumbnail-600w.jpg

# 使用 ffmpeg
ffmpeg -i original.jpg -vf scale=600:-1 thumbnail-600w.jpg
```

### 2. WebP 轉換

```bash
# 使用 cwebp
cwebp -q 85 original.jpg -o thumbnail.webp
```

### 3. 建議的圖片尺寸

| 用途 | 寬度 | 高度 | 品質 | 格式 |
|------|------|------|------|------|
| 縮圖格子 | 300-400px | 225-300px | 75-80% | JPG/WebP |
| 卡片圖 | 400-600px | 300-450px | 80-85% | JPG/WebP |
| 幻燈片 | 800-1200px | 600-900px | 85-90% | JPG/WebP |

## 四、CSS 套樣式

### 響應式容器

```css
.thumbnail-container {
  position: relative;
  width: 100%;
  max-width: 600px;
  padding-bottom: 75%; /* 4:3 比例 */
  overflow: hidden;
  border-radius: 8px;
  background: #f0f0f0;
}

.thumbnail-container img {
  position: absolute;
  top: 0;
  left: 0;
  width: 100%;
  height: 100%;
  object-fit: cover; /* 保持比例並裁切 */
  object-position: center;
}
```

### 懸停效果

```css
.thumbnail {
  transition: transform 0.3s ease, filter 0.3s ease;
}

.thumbnail:hover {
  transform: scale(1.05);
  filter: brightness(1.1);
}
```

## 五、HTML 模板範例

### 圖片網格

```html
<div class="gallery">
  <article class="thumbnail-card">
    <div class="thumbnail-container">
      <picture>
        <source 
          srcset="images/thumb-1.webp 1x, 
                  images/thumb-1-2x.webp 2x"
          type="image/webp" 
        />
        <img 
          src="images/thumb-1.jpg" 
          alt="圖片 1" 
          loading="lazy"
        />
      </picture>
    </div>
    <h3>圖片標題</h3>
    <p>圖片描述</p>
  </article>
</div>
```

### 單個縮圖 + 燈箱

```html
<a href="images/full-size.jpg" class="lightbox">
  <img 
    src="images/thumbnail.jpg" 
    srcset="images/thumbnail-2x.jpg 2x"
    alt="點擊查看大圖" 
    loading="lazy"
  />
</a>
```

## 六、常見問題

### Q: 應該用 `jpg` 還是 `webp`？
**A:** 優先用 `webp`（檔案小 25-35%），但要用 `<picture>` 提供 `jpg` 後備方案。

### Q: `lazy` 加載安全嗎？
**A:** 安全。用於螢幕外的圖片，可以加快初始加載速度。關鍵圖片（above-fold）不要加 `lazy`。

### Q: 縮圖應該有多大？
**A:** 最大寬度建議 600-800px（視使用場景而定），檔案大小控制在 50-150KB。

### Q: 如何處理不同的影像長寬比？
**A:** 使用 `object-fit: cover` 配合 `aspect-ratio` CSS 規則：
```css
.thumbnail {
  width: 100%;
  aspect-ratio: 4 / 3; /* 或其他比例 */
  object-fit: cover;
}
```

## 七、工具與資源

| 工具 | 用途 | 指令 |
|------|------|------|
| ImageMagick | 圖片縮放壓縮 | `convert` / `mogrify` |
| FFmpeg | 視訊截圖為縮圖 | `ffmpeg -i video.mp4 -ss 0 -vf scale=600:-1 thumb.jpg` |
| cwebp | WebP 轉換 | `cwebp -q 85 input.jpg -o output.webp` |
| TinyPNG | 線上壓縮 | https://tinypng.com |
| Squoosh | 線上最佳化 | https://squoosh.app |

## 八、檢查清單

部署前確認：

- [ ] 所有縮圖都有 `alt` 屬性（SEO 與無障礙）
- [ ] 圖片尺寸適配不同螢幕（使用 srcset 或 picture）
- [ ] 檔案大小優化（單個縮圖 < 150KB）
- [ ] 使用 WebP 並提供後備格式
- [ ] 關鍵圖片移除 `loading="lazy"`
- [ ] 響應式容器使用 `aspect-ratio` 或 padding-bottom 技巧
- [ ] 色彩空間正確（sRGB 用於 Web）
- [ ] 在各瀏覽器與設備測試（Chrome, Safari, Firefox, 行動設備）

## 九、效能指標

目標指標（Lighthouse）：

- **LCP（最大內容繪製）**：圖片應在 2.5 秒內載入
- **CLS（累積排版轉移）**：設置 `width` / `height` 或 `aspect-ratio` 防止抖動
- **圖片最佳化評分**：應達 90+ 分

## 十、進階 — 動態生成縮圖

### Python 範例

```python
from PIL import Image
import os

def create_thumbnails(source_dir, output_dir, size=(600, 450)):
    os.makedirs(output_dir, exist_ok=True)
    
    for filename in os.listdir(source_dir):
        if filename.endswith(('.jpg', '.png')):
            img = Image.open(os.path.join(source_dir, filename))
            img.thumbnail(size, Image.Resampling.LANCZOS)
            
            # 儲存為 JPG 與 WebP
            base_name = os.path.splitext(filename)[0]
            img.save(f"{output_dir}/{base_name}.jpg", "JPEG", quality=85)
            img.save(f"{output_dir}/{base_name}.webp", "WEBP", quality=85)

create_thumbnails("./original", "./thumbnails")
```

### Node.js 範例（Sharp）

```javascript
const sharp = require('sharp');
const fs = require('fs').promises;
const path = require('path');

async function createThumbnails(sourceDir, outputDir, width = 600, height = 450) {
  await fs.mkdir(outputDir, { recursive: true });
  
  const files = await fs.readdir(sourceDir);
  
  for (const file of files) {
    if (!/\.(jpg|png)$/i.test(file)) continue;
    
    const input = path.join(sourceDir, file);
    const baseName = path.parse(file).name;
    
    // JPG
    await sharp(input)
      .resize(width, height, { fit: 'cover' })
      .jpeg({ quality: 85 })
      .toFile(path.join(outputDir, `${baseName}.jpg`));
    
    // WebP
    await sharp(input)
      .resize(width, height, { fit: 'cover' })
      .webp({ quality: 85 })
      .toFile(path.join(outputDir, `${baseName}.webp`));
  }
}

createThumbnails('./original', './thumbnails');
```

---

需要快速實作縮圖嵌入？告訴我：
- 使用場景（格子、卡片、單張等）
- 圖片來源格式
- 目標平台（Web、行動、電子報）
