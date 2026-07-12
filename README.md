# Math Platform — joylashtirish (deploy) qo'llanmasi

Bu papkada saytingiz uchun to'liq loyiha bor (Vite + React). Quyidagi qadamlarni bosqichma-bosqich bajaring.

---

## 1-BOSQICH: Kompyuteringizga kerakli dasturlarni o'rnatish

1. **Node.js**ni o'rnating (agar hali yo'q bo'lsa): https://nodejs.org — "LTS" versiyani yuklab, o'rnating.
2. **Git**ni o'rnating: https://git-scm.com/downloads
3. GitHub'da bepul akkaunt oching (agar yo'q bo'lsa): https://github.com/signup

---

## 2-BOSQICH: Loyihani mahalliy (o'z kompyuteringizda) sinab ko'rish (ixtiyoriy, lekin tavsiya etiladi)

Bu papkani kompyuteringizga tushirgandan so'ng, terminal (Windows'da "Command Prompt" yoki "PowerShell", Mac'da "Terminal") oching va shu papka ichiga o'ting:

```bash
cd math-platform-deploy
npm install
npm run dev
```

Terminal sizga bir manzil beradi (masalan `http://localhost:5173`) — uni brauzerda oching, saytingiz shu yerda ishlashi kerak. Xatolik bo'lsa, screenshot olib menga yuboring.

---

## 3-BOSQICH: GitHub'ga yuklash

1. https://github.com ga kiring → yuqori o'ngdagi **"+"** → **"New repository"**.
2. Nom bering (masalan `math-platform`) → **"Create repository"**.
3. Terminalda shu papka ichida:

```bash
git init
git add .
git commit -m "Birinchi versiya"
git branch -M main
git remote add origin https://github.com/FOYDALANUVCHI_NOMINGIZ/math-platform.git
git push -u origin main
```

(`FOYDALANUVCHI_NOMINGIZ` o'rniga o'z GitHub login'ingizni yozing — GitHub bu buyruqni repository yaratgandan keyin sizga aynan shu holda ko'rsatadi, shunchaki nusxalab qo'ysangiz bo'ldi.)

---

## 4-BOSQICH: Vercel'ga joylashtirish (bepul, HTTPS manzil bilan)

1. https://vercel.com ga kiring → **"Sign up"** → **"Continue with GitHub"** (shu bilan ikkalasi avtomatik bog'lanadi).
2. Kirgandan so'ng **"Add New..." → "Project"**.
3. Ro'yxatdan `math-platform` repositoriyangizni tanlang → **"Import"**.
4. Sozlamalar avtomatik to'g'ri aniqlanadi (Framework: Vite) — hech narsani o'zgartirmasdan **"Deploy"** tugmasini bosing.
5. 1-2 daqiqada tayyor bo'ladi va sizga shunday manzil beradi:
   `https://math-platform-sizning-nomingiz.vercel.app`

**Shu manzil — sizning HTTPS saytingiz!** Uni Telegram botga ulash uchun ishlatasiz.

---

## 5-BOSQICH: Firebase sozlamalarini kiritish

Eslatma: kodda `FIREBASE_CONFIG` bo'sh holatda turibdi (fayl boshida, `src/App.jsx` ichida). Firebase loyihangizni yaratgan bo'lsangiz, o'sha config qiymatlarini shu yerga qo'ying, so'ng qayta:

```bash
git add .
git commit -m "Firebase config qo'shildi"
git push
```

Vercel buni ko'rib, saytni **avtomatik qayta joylashtiradi** (har safar GitHub'ga push qilganingizda shunday bo'ladi — qo'lda hech narsa qilish shart emas).

---

## 6-BOSQICH: Telegram botga ulash

1. Telegram'da **@BotFather** bilan suhbatga kiring.
2. `/mybots` → botingizni tanlang → **"Bot Settings"** → **"Menu Button"** (yoki **"Configure Mini App"**).
3. Vercel bergan HTTPS manzilni kiriting.
4. Tayyor — endi botingizda tugma bosilganda sayt Telegram ichida ochiladi.

---

## Muammo bo'lsa

Har qanday bosqichda xatolik chiqsa, terminal/brauzer ekranining screenshotini olib yuboring — birga hal qilamiz.
