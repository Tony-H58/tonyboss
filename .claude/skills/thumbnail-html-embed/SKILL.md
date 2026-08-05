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

## 十一、常見問題排除

| 問題 | 原因 | 解決方案 |
|------|------|---------|
| 縮圖不顯示 | 路徑錯誤 | 檢查 `src` 是否相對於 HTML 檔位置 |
| 縮圖模糊 | 原始圖過小 | 使用 ≥ 600px 寬的圖片 |
| 縮圖被裁切 | `object-fit: cover` 導致 | 改用 `contain` 或調整容器比例 |
| 載入緩慢 | 檔案過大 | 壓縮至 50-150KB |
| 手機版排版亂 | CSS 不響應式 | 添加 `@media` 查詢 |
| WebP 不支援 | 舊瀏覽器 | 使用 `<picture>` 提供後備 |
| 色彩不符 | 色彩空間不對 | 確保 sRGB，避免 CMYK |
| 比例不一致 | 原始圖長寬比不同 | 統一至 4:3 或 16:9 |

## 十二、快速故障排除流程

### 縮圖不顯示？

```html
<!-- 開發者工具診斷 (F12) -->

1. 檢查 Network 分頁
   → 右鍵檢查元素
   → 查看 src 屬性的完整路徑

2. 檢查 Console 分頁
   → 是否有 404 或跨域錯誤？

3. 驗證檔案存在
   → 在檔案管理器中導航到該路徑
   → 確認檔案確實存在

4. 測試相對路徑
   → HTML: E:\project\index.html
   → 圖片: E:\project\images\thumb.jpg
   → src="images/thumb.jpg" ✅
   → src="/images/thumb.jpg" ❌ (絕對路徑)
```

### 載入緩慢？

```
1. 打開瀏覽器開發者工具 (F12)
   → Performance 分頁
   → 點「錄製」→ 重載頁面 → 停止

2. 查看圖片載入時間
   → 若某張圖超過 2 秒
   → 該圖需壓縮

3. 使用線上工具檢查
   → https://tinypng.com
   → https://squoosh.app
   → 目標：將 300KB 壓至 80KB
```

## 十三、實際應用範例

### 範例 1：電商產品格子

```html
<div class="product-grid">
  <article class="product-card">
    <div class="image-container">
      <picture>
        <source srcset="product-thumb.webp" type="image/webp" />
        <img src="product-thumb.jpg" alt="產品名稱" loading="lazy" />
      </picture>
    </div>
    <h3>產品名稱</h3>
    <p class="price">$99.99</p>
    <button>加入購物車</button>
  </article>
</div>

<style>
  .product-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    gap: 20px;
  }
  
  .product-card {
    border: 1px solid #ddd;
    border-radius: 8px;
    overflow: hidden;
    transition: transform 0.3s;
  }
  
  .product-card:hover {
    transform: scale(1.05);
    box-shadow: 0 4px 12px rgba(0,0,0,0.2);
  }
  
  .image-container {
    width: 100%;
    aspect-ratio: 1;
    overflow: hidden;
    background: #f5f5f5;
  }
  
  .image-container img {
    width: 100%;
    height: 100%;
    object-fit: cover;
  }
</style>
```

### 範例 2：部落格文章卡片

```html
<article class="blog-card">
  <div class="thumbnail">
    <img src="blog-thumb-400w.jpg" 
         srcset="blog-thumb-600w.jpg 600w, 
                 blog-thumb-900w.jpg 900w"
         sizes="(max-width: 600px) 100vw, 50vw"
         alt="文章預覽圖" />
    <span class="date-badge">2026/08/05</span>
  </div>
  
  <div class="content">
    <h2>文章標題</h2>
    <p>文章摘要...</p>
    <a href="/article" class="read-more">閱讀更多 →</a>
  </div>
</article>

<style>
  .blog-card {
    display: flex;
    gap: 20px;
    border: 1px solid #eee;
    border-radius: 8px;
    overflow: hidden;
    transition: all 0.3s;
  }
  
  .thumbnail {
    position: relative;
    width: 300px;
    height: 200px;
    flex-shrink: 0;
    overflow: hidden;
  }
  
  .thumbnail img {
    width: 100%;
    height: 100%;
    object-fit: cover;
  }
  
  .date-badge {
    position: absolute;
    bottom: 10px;
    right: 10px;
    background: rgba(0,0,0,0.7);
    color: white;
    padding: 5px 10px;
    border-radius: 4px;
    font-size: 12px;
  }
  
  .blog-card:hover {
    box-shadow: 0 8px 16px rgba(0,0,0,0.1);
  }
</style>
```

## 十四、效能檢查清單

部署前檢查以下項目：

- [ ] **圖片大小** — 每個縮圖 < 150KB
- [ ] **解析度** — 寬度至少 600px（2x 高 DPI）
- [ ] **格式** — 優先 WebP，備用 JPG / PNG
- [ ] **Alt 文字** — 所有圖片都有描述性 `alt` 屬性
- [ ] **響應式測試** — 在手機、平板、電腦上測試
- [ ] **載入速度** — LCP（最大內容繪製）< 2.5 秒
- [ ] **跨瀏覽器** — Chrome、Firefox、Safari、Edge 都支援
- [ ] **無障礙** — 螢幕閱讀器能正確讀取圖片描述
- [ ] **SEO** — 圖片命名清晰（不用 `image1.jpg`）
- [ ] **色彩** — 在淺色 / 深色主題中都清晰可見

## 十五、線上工具速查表

| 工作 | 推薦工具 | URL |
|------|---------|-----|
| 圖片壓縮 | TinyPNG | https://tinypng.com |
| 進階優化 | Squoosh | https://squoosh.app |
| Base64 轉換 | Base64 Image | https://www.base64-image.de/ |
| Lighthouse 測試 | Chrome DevTools | F12 → Lighthouse |
| 色彩檢查 | WebAIM | https://webaim.org/resources/contrastchecker/ |

---

需要快速實作縮圖嵌入？告訴我：
- 使用場景（格子、卡片、單張等）
- 圖片來源格式
- 目標平台（Web、行動、電子報）
- 特殊需求（動畫、濾鏡等）
