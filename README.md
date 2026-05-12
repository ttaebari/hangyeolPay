# hangyeolPay

한결님의 일을 하나라도 줄이기 위한 급여명세서 작성 도움 웹입니다.

급여대장 XLSX를 업로드하면 직원별 급여명세서 XLSX를 생성합니다.

## 실행

```bash
npm install
npm run dev
```

## 빌드

```bash
npm run build
```

## Vercel 배포

Vercel에서 GitHub 저장소를 연결하면 `vercel.json` 설정에 따라 빌드됩니다.

- Framework Preset: Vite
- Build Command: `npm run build`
- Output Directory: `dist`
