# 蝦皮商品 PR 品質分數儀表板

蝦皮商品品質分數分析工具，依據 CVR / CTR / 跳出率 / 購物車率 / 搜尋點擊自動計算 PR 分數，並提供主圖、商品說明、價格三方向改善建議。

## 使用方式

上傳蝦皮賣家中心匯出的 `parentskudetail_*.xlsx`，系統自動解析並更新排行。

## 本機開發

```bash
npm install
npm run dev
```

## 部署（GitHub + Vercel）

1. 將此資料夾上傳到 GitHub Repository
2. 登入 Vercel → Import Repository → Deploy
3. Framework Preset 選 **Vite**，其餘預設即可
