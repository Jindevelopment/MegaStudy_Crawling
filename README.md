# 미맥스터디 Q&A 크롤러

미맥스터디 사이트의 Q&A 게시글을 날짜 범위와 선생님별로 수집해 Excel로 다운로드하는 웹 앱입니다.

---

## 로컬 실행

### 1. 패키지 설치

```bash
pip install -r requirements.txt
```

### 2. 서버 실행

```bash
python run.py
```

### 3. 브라우저에서 접속

```
http://localhost:5000
```

---

## Vercel 배포

### 1. GitHub 저장소 생성 후 push

```bash
git init
git add .
git commit -m "init"
git remote add origin https://github.com/<계정>/<저장소명>.git
git push -u origin main
```

### 2. Vercel 배포

**방법 A — Vercel 대시보드 (권장)**

1. [vercel.com](https://vercel.com) 로그인
2. **Add New Project** → GitHub 저장소 선택
3. 설정 변경 없이 **Deploy** 클릭

**방법 B — Vercel CLI**

```bash
npm install -g vercel
vercel
```

배포 완료 후 발급된 URL로 접속하면 바로 사용 가능합니다.

---

## 주의사항

- Vercel **Hobby** 플랜은 함수 실행 시간이 최대 **60초**입니다. 선생님이 많거나 날짜 범위가 길면 **Pro** 플랜(300초)이 필요할 수 있습니다.
- 선생님 목록은 구글 시트에서 자동으로 불러옵니다. 시트 공유 설정이 **링크가 있는 모든 사용자 → 뷰어**로 되어 있어야 합니다.
