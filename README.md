# MF-V-01 — Masaüstü (Offline) Sürüm

KB Group markalı Maliye-Finans uygulamasının **çevrimdışı, tek cihaz** sürümü. Electron ile paketlenir, veriler yalnızca bu bilgisayarda (`localStorage`) tutulur — sunucu, internet bağlantısı veya paylaşımlı veritabanı gerektirmez.

Bu proje, aynı ürünün **çevrimiçi/çok kullanıcılı** sürümünden (Postgres + Node, ayrı bir repo) bilinçli olarak ayrı tutulur. İkisi aynı arayüz/özellik kümesinden başlamış olsa da mimarileri artık farklıdır; bkz. [`MF-V-01-ANA-PROMPT-1.0.8.md`](./MF-V-01-ANA-PROMPT-1.0.8.md) → "Durum güncellemesi" bölümü.

## Bu sürüm ne yapar

Kasa, Gelir, Gider, Rapor Hazırla, Mali Özel Notlar, Arşiv, Kullanıcılar ve Ayarlar bölümlerini tek bir masaüstü penceresinde yönetir. Tüm veriler tarayıcı depolama alanında (`localStorage`) tutulur; "Ayarlar → Veri (Database) Ayarları" üzerinden JSON olarak dışa/içe aktarılabilir.

- **Kullanıcı Yönetimi**: ekle/düzenle/sil, izin bazlı sayfa erişimi, başarısız girişte otomatik hesap kilitleme (admin açar).
- **Şifreler** düz metin değil, PBKDF2-SHA256 (tuzlu) olarak saklanır.
- **Rapor Hazırla**: Kasa/Gelir/Gider kayıtlarından seçerek rapor oluşturma, Excel'e aktarma.

## Gereksinimler

- Node.js `>=22.13.0`
- Windows üzerinde paketleme için `electron-builder` (bağımlılıklarla birlikte gelir)

## Komutlar

```bash
npm install          # bağımlılıkları kur
npm run dev           # geliştirme sunucusu (Vite + Cloudflare-plugin sandbox; hızlı önizleme için)
npm run build          # üretim derlemesi (dist/)
npm test               # derleme + render/asset testleri
npm run desktop:dist    # Windows NSIS kurulum dosyasını üret (release/MF-V-01-Setup-<sürüm>.exe)
```

`npm run dev`, Cloudflare Workers (Miniflare) sandbox'ı içinde çalışır — bu nedenle bazı Windows/Git-Bash ortamlarında `WRANGLER_LOG_PATH=...` ön eki script hatası verebilir; bu durumda doğrudan `npx vite` çalıştırılabilir veya üretim derlemesi (`npm run build && PORT=<port> npx vinext start`) üzerinden test edilebilir. Gerçek çalışma zamanı her hâlükârda Electron'un kendi sunucusudur (`desktop/main.cjs`), ham HTTP sunucusu değil.

## Kaynak yapısı

- Ana arayüz: `app/page.tsx` (tek dosya, modüllere bölünmemiş — web sürümünden farklı olarak)
- Stil: `app/globals.css`
- Masaüstü ana süreci: `desktop/main.cjs`
- Windows kurucu ayarları: `package.json` → `"build"` alanı (electron-builder), NSIS betiği `build/mf-v-01-installer.nsi`
- Testler: `tests/*.test.mjs`

## Teslim / dağıtım

`npm run desktop:dist` sonrası `release/MF-V-01-Setup-<sürüm>.exe` üretilir. Bu proje şu an bir Git deposu olarak hazırlanmıştır ama herhangi bir uzak sunucuya (GitHub dâhil) bağlı değildir — istendiğinde `git remote add origin <url> && git push -u origin main` ile bağlanabilir.
