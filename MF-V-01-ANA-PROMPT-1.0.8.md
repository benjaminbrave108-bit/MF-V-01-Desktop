# MF-V-01 — Ana Proje Promptu

## Durum güncellemesi (2026-08-18 — Kullanıcı yönetimi, güvenlik ve rapor sistemi güncellemesi)

Bu belgenin geri kalanı masaüstü sürümünün genel kurallarını anlatır ve hâlâ
geçerlidir. Bu bölüm, aynı üründen ayrı bir çevrimiçi/çok kullanıcılı sürüm
(Postgres + Node, ayrı bir repo) türetildikten *sonra* bu masaüstü sürümüne
yapılan ve o web sürümüyle **senkron olmayan**, bağımsız iyileştirmeleri
özetler:

- **Kullanıcı Yönetimi** artık gerçek: Kullanıcılar sayfasından ekleme,
  düzenleme, silme ve sayfa bazlı izin ataması yapılabiliyor (bkz.
  `UserAccount` tipi, `Users`/`UserModal` bileşenleri). Admin olmayan bir
  kullanıcı yalnız kendisine tanımlanan bölümleri görür (`canAccess()`).
- **Şifreler artık düz metin değil**: PBKDF2-SHA256 (tuzlu, 100.000
  iterasyon, Web Crypto `crypto.subtle`) ile hash'leniyor. 5 başarısız giriş
  denemesinden sonra hesap otomatik kilitleniyor; yalnız bir admin
  Kullanıcılar sayfasından kilidi açabiliyor. Bu, sunucu tarafı olmayan
  tek-cihaz bir uygulama için ulaşılabilir en iyi seviyedir — gerçek bir
  sunucu güvenlik sınırının yerini tutmaz.
- **"Birden Fazla Şirket Veritabanı" özelliği kaldırıldı** (çevrimiçi
  sürümle bilinçli olarak eşitlendi). Yerine web sürümündeki gibi tek bir
  "Veri (Database) Ayarları" bölümü geldi: şifre onaylı Dışa Aktar / İçe
  Aktar / Tüm Mali Verileri Sil.
- **"Rapor Hazırla" tamamen web sürümündeki sistemle değiştirildi.** Eski
  "Aylık Rapor + Birim Bütçe Dağıtımı" sistemi (Excel şablon içe/dışa
  aktarma, birim bazlı gelir dağıtımı) tamamen kaldırıldı. Bu, bilinçli bir
  özellik kaybıdır — proje sahibinin açık onayıyla yapılmıştır.
- Bu proje artık bir **Git deposu** (bkz. kök `.git/`) ama şu an herhangi bir
  uzak sunucuya (GitHub dâhil) bağlı değil; istendiğinde
  `git remote add origin <url> && git push -u origin main` ile bağlanabilir.

Çevrimiçi sürümdeki sunucuya özgü maddeler (IP engelleme, oturum çerezi,
Zod doğrulama, rate limiting) bu mimariye taşınmadı — bu mimaride sunucu
olmadığı için karşılığı yok.

## Proje kimliği

- Uygulama adı: **MF-V-01**
- Sürüm: **1.0.8**
- Geliştirici markası: **KB Group**
- Hedef sistem: 64-bit Windows
- Çalışma biçimi: Electron tabanlı bağımsız masaüstü uygulaması; varsayılan tarayıcıyı açmaz.
- Kurulum biçimi: Standart Windows NSIS kurulum sihirbazı (`MF-V-01-Setup-1.0.8.exe`).
- Kullanıcıya ayrıca Node.js kurdurulmaz; gerekli çalışma motoru pakete gömülüdür.

## Temel işlevler

MF-V-01; kasa, gelir, gider, rapor hazırlama, mali not, arşiv, kullanıcı ve uygulama/veritabanı ayarlarını tek bir masaüstü arayüzünde yönetir. Veriler cihazdaki uygulama profilinde yerel olarak saklanır. Tüm mali veriler (kayıt/arşiv/not/rapor) tek bir JSON dosyası olarak dışa aktarılabilir ve yeniden içe aktarılabilir (bkz. Durum güncellemesi).

## Dil kuralları

- Desteklenen diller: **English, Türkçe, Kurdî (Kurmancî / Latin alfabesi)**.
- İlk açılışta üç dilli dil seçim ekranı gösterilir.
- Seçilen dil giriş ekranına ve uygulamanın tüm bölümlerine uygulanır.
- Dil daha sonra üst çubuktan değiştirilebilir.
- Yeni bir kullanıcı metni eklenirken Türkçe, İngilizce ve Kürtçe karşılıkları birlikte eklenmelidir.
- Kürtçe bir terim belirsizse tahmin edilmemeli; proje sahibine Türkçe olarak sorulmalıdır.

## Logo kuralları

- Varsayılan logo: `public/kb-logo.png` içindeki KB Group logosudur.
- Giriş/ilk açılış ekranında KB Group logosu gösterilir.
- Sol menünün üstündeki şirket logosu kullanıcı tarafından değiştirilebilir.
- Kullanıcı şirket logosunu kaldırırsa üst alan varsayılan KB Group logosuna döner.
- Sol menünün altındaki “KB Group Tarafından Geliştirilmiştir” alanında KB Group logosu sabittir ve şirket logosundan etkilenmez.

## Veritabanı güvenliği

- Ayarlar > Veri (Database) Ayarları bölümünde veritabanı yükleme, dışa aktarma ve silme seçenekleri bulunur.
- “Veritabanını Sil” işlemi iki ayrı onay sorusu göstermelidir.
- İkinci onay verilmeden hiçbir mali veri silinmemelidir.
- Onay tamamlanınca yerel mali kayıtlar, arşiv, notlar, raporlar ve yüklenmiş veritabanı paketleri temizlenir; boş veritabanıyla başlanır.

## Ayarlar sayfası düzeni

- Uygulama Ayarları sekmesinde “Genel Ayarlar” masaüstünde sol sütunda bulunur.
- “Yazı ve Renk Ayarları” sağ sütunda, Genel Ayarlar ile yan yana bulunur.
- Küçük ekranlarda bölümler okunabilir biçimde alt alta geçer.
- Yazı boyutu, yazı tipi ve renk ayarlarının mevcut işlevleri korunur.

## Masaüstü başlatma ve hata güvenliği

- Electron, paket içindeki Sites/Vinext sunucu çıktısını yerel `127.0.0.1` adresinde rastgele bir portta çalıştırır ve kendi penceresinde gösterir.
- Windows dosya yolları `pathToFileURL` ile güvenli modül adresine çevrilir.
- Uygulama verileri ve Chromium önbelleği kullanıcının yazılabilir AppData klasöründe tutulur.
- Pencere birincil ekranın çalışma alanında ortalanır; görünmez kalırsa zorla gösterilir.
- Aynı anda yalnız bir uygulama örneği çalışır; ikinci açılış mevcut pencereyi öne getirir.
- Başlangıç hataları `startup.log` dosyasına yazılır ve kullanıcıya hata penceresi gösterilir.
- `dist/client` içindeki JavaScript, CSS, logo ve diğer statik dosyalar yerel sunucudan doğru MIME türleriyle sunulur; ana HTML'nin `200 OK` vermesi tek başına başarı sayılmaz.
- Electron renderer konsol hataları, yükleme hataları ve renderer kapanmaları `startup.log` dosyasına yazılır.
- Arayüz yükleme sonrası otomatik sağlık kontrolünden geçmezse boş/beyaz ekran bırakılmamalı; kullanıcıya görünür hata ve günlük dosyasının konumu gösterilmelidir.
- Kurulum ve uygulama kodunda bilgisayarı kapatma veya yeniden başlatma komutu bulunmamalıdır.

## Kaynak yapısı

- Ana arayüz: `app/page.tsx`
- Stil: `app/globals.css`
- Masaüstü ana süreci: `desktop/main.cjs`
- Sabit logo: `public/kb-logo.png`
- Windows kurucu tanımı: `build/standalone-installer.nsi`
- Paketleme ayarları: `package.json`
- Sites kimliği: `.openai/hosting.json` (kimlik korunmalı ve değiştirilmemelidir)

## Derleme ve doğrulama

1. Bağımlılıkları kilit dosyasına göre kurun.
2. `npm test` ile üretim derlemesini, HTML testini ve HTML'nin çağırdığı her masaüstü JavaScript/CSS varlığının paket içinde bulunduğunu doğrulayan testi çalıştırın.
3. 64-bit Windows NSIS kurucusunu üretin.
4. Kurucunun PE/NSIS yapısını, dosya boyutunu ve SHA-256 özetini doğrulayın; kurulan gerçek `app.asar` içinden sürüm, `desktop/main.cjs`, sunucu ve tüm arayüz varlıklarını ayrıca kontrol edin.
5. Windows üzerinde kurulum, masaüstü kısayolu, ilk dil seçimi, oturum açma, logo görünümü, iki aşamalı veritabanı silme ve kaldırma işlemlerini gerçek cihazda test edin.

## Teslim kuralları

Her yeni sürümde birlikte teslim edilmelidir:

1. `MF-V-01-Setup-<sürüm>.exe`
2. `MF-V-01-Source-Code-<sürüm>.zip`
3. Güncellenmiş ana prompt dosyası

Kurulum dosyası kod imzalama sertifikasıyla imzalanmamışsa SmartScreen uyarısı çıkabileceği açıkça belirtilmelidir. Gerçek Windows açılış testi yapılmadıysa yalnız Linux paketleme veya içerik testi “Windows’ta kesin çalışıyor” şeklinde sunulmamalıdır.


## 1.0.8 Responsive Arayüz
- Küçük ekranlarda yan menü otomatik daralır ve elle açılıp kapatılabilir.
- Arayüz zoom seçenekleri: %80, %90, %100, %110, %125. Ctrl + / Ctrl - / Ctrl 0 kısayolları desteklenir.
- Form/modaller ekran yüksekliğine göre kaydırılır; Kaydet/İptal alanı erişilebilir kalır.
- Küçük Windows ekranlarında pencere minimum 680x500 boyuta kadar küçültülebilir.
- Tablolar ve yoğun paneller gerektiğinde yatay kaydırılır.

### V6 - Veri Kaliciligi Duzeltmesi
- Electron masaustu sunucusu sabit 127.0.0.1:47831 adresini kullanir.
- Kayitli veriler okunmadan bos baslangic verisi yazilmaz.
- Uygulama kapatilip yeniden acildiginda Kasa, Gelir, Gider, rapor ve diger yerel kayitlar korunur.
