# Exchange Message Tracking Pro

Bu sürüm, incelenen 2019 tarihli PowerShell Studio betiğinin işlevini daha güvenli ve anlaşılır bir arayüzle yeniden kurar. Orijinal masaüstü dosyasını değiştirmez.

## Kısa özet

Exchange Server üzerindeki ileti izleme günlüklerinde e-postanın kimden geldiğini, kime gittiğini, hangi sunuculardan geçtiğini ve teslim edilip edilmediğini araştırır. Sonuçları Türkçe açıklamalarla gösterir; mevcut sonuçlarda arama/filtreleme yapar ve raporu CSV veya HTML olarak dışa aktarır.

Başlıca yetenekleri: tarih ve ileti ölçütleriyle arama, tek veya çoklu sunucu sorgusu, teslimat/hata olaylarını anlamlandırma, sonuç içi metin-durum-olay filtreleri, ayrıntılı tanı bilgisi, iptal/ilerleme takibi ve uzak Exchange PowerShell bağlantısıdır.

## Öne çıkan geliştirmeler

- Komut metni üretip `Invoke-Expression` çalıştırmak yerine doğrulanmış parametreler ve splatting kullanır.
- Eski `Get-PSSnapin` yaklaşımı yerine hazır Exchange Management Shell oturumunu algılar; gerekirse uzak Exchange PowerShell oturumu açabilir.
- Sonuçları ayrı `Out-GridView` penceresinde değil, uygulama içinde sıralanabilir bir tabloda gösterir.
- Mevcut sonuçlarda gönderen, alıcı, konu, sunucu, durum, Message-ID ve tanı metinleri dahil önemli alanlarda anlık arama yapar; durum ve olay türü ayrıca filtrelenebilir.
- Olayları Microsoft'un resmî tanımlarına dayanan Türkçe açıklamalarla sunar. `Teslim edildi`, `Aktarım`, `Alındı`, `Gecikme`, `Teslim edilmedi` ve `Başarısız` gibi sonuçları birbirinden ayırır; SMTP `4xx` ve `5xx` durumlarını ayrıca değerlendirir.
- Tüm taşıma sunucularını veya tek sunucuyu sorgulayabilir; sunucu bazındaki hataları günlüğe kaydeder ve diğer sunucularla devam eder.
- Varsayılan sorguyu son 24 saat ve sunucu başına 1000 sonuçla sınırlar. Geniş/sınırsız sorgularda kullanıcıyı uyarır.
- HealthMailbox trafiğini daha dar ve anlamlı bir kuralla filtreler.
- Ayrıntılı kayıt paneli, sorgu önizlemesi, durum çubuğu, ilerleme, iptal, `Ctrl+F`, F5 ve Escape kısayolları içerir.
- UTF-8 CSV ve biçimlendirilmiş HTML raporu dışa aktarabilir.
- Uzak oturumu uygulama kapanırken temizler.

## Gereksinimler

- Windows PowerShell 5.1 (`powershell.exe`); PowerShell 7 değil.
- Şirket içi Exchange Server 2016, 2019 veya Subscription Edition.
- `Get-MessageTrackingLog` çalıştırmaya ve hedef sunucuları görmeye yeterli Exchange RBAC yetkisi.
- Uzak bağlantı kullanılacaksa WinRM/Kerberos erişimi ve Exchange Remote PowerShell izni.

> `Get-MessageTrackingLog` yalnızca on-premises Exchange içindir; Exchange Online Message Trace için farklı cmdlet ve API gerekir.

## Çalıştırma

Exchange Management Shell içinde:

```powershell
.\Exchange-Message-Tracking-Pro.ps1
```

Standart Windows PowerShell 5.1 içinde otomatik bağlantı ile:

```powershell
.\Exchange-Message-Tracking-Pro.ps1 -ExchangeServerFqdn ex01.contoso.local
```

Farklı bir yönetici hesabı sorulması için:

```powershell
.\Exchange-Message-Tracking-Pro.ps1 -ExchangeServerFqdn ex01.contoso.local -UseCredential
```

Betik yürütme politikası tarafından engellenirse kurum politikasını atlamayın. Kurumsal ortamda tercih edilen yöntem betiği kod imzalama sertifikasıyla imzalamak ve BT ekibinizin onayladığı yürütme politikasını korumaktır.

## Kullanım notları

- `Alıcılar` alanında birden çok adresi virgül, noktalı virgül veya yeni satırla ayırabilirsiniz.
- `[Tümü]` seçimi her taşıma sunucusunda ayrı sorgu yapar; `Sunucu başı` sınırı her sunucuya uygulanır.
- `Message-ID` değerini tam girin. Açılı/kapanış parantezleri (`< >`) korunabilir.
- `Network ID` filtresi eski Exchange sürümlerinde olmayabilir; uygulama desteklenmeyen parametreyi sorgudan önce bildirir.
- Kırmızı satır başarısızlık, sarı satır gecikme/inceleme, yeşil satır doğrulanmış yerel teslim olayıdır. `SEND`, `RECEIVE` ve `SUBMIT` yalnızca ara taşıma aşamalarıdır ve son teslim olarak işaretlenmez. `DELIVER`, yerel posta kutusuna teslimin temel göstergesidir.
- `Sonuçlarda ara` alanına birden çok sözcük yazıldığında her sözcük önemli sonuç sütunlarından en az birinde bulunmalıdır. Örneğin `kullanici@contoso.com FAIL` birlikte aranabilir.
- `Ctrl+F` sonuç arama alanına gider. Escape etkin sonuç filtresini temizler; arama devam ederken Escape sorguya iptal isteği gönderir.
- CSV ve HTML dışa aktarma, etkin filtre varsa yalnızca tabloda gösterilen kayıtları aktarır. Üstteki sayaç `Gösterilen / Toplam` değerini belirtir.
- CSV dosyası Windows bölgesel liste ayırıcısını kullanır (`-UseCulture`) ve Excel ile daha rahat açılır.

## İncelenen eski betikteki başlıca sorunlar

1. Kullanıcı girdileri komut metnine ekleniyor ve `Invoke-Expression` ile çalıştırılıyordu. Bu hem kaçış/yerel tarih sorunlarına hem de komut enjeksiyonu riskine yol açıyordu.
2. `$viewcommand` ve `$userviewcommand` her aramada temizlenmediği için aynı GUI oturumundaki sonraki aramalarda komutlar üst üste eklenebiliyordu.
3. Exchange snap-in yokken yalnızca hata etiketi gösteriliyor, fakat akış durdurulmadığı için yine sorgu deneniyordu.
4. `Get-PSSnapin`/`Add-PSSnapin *Exchange*` modern Exchange Management Shell ve uzak PowerShell kullanım modeli için kırılgandı.
5. Arama GUI iş parçacığında tek parça çalışıyor; kullanıcıya gerçek ilerleme, iptal veya sunucu bazında hata bilgisi vermiyordu.
6. `Get-TransportService | Get-MessageTrackingLog -WarningAction 0` yaklaşımı sunucu hatalarını gizliyor ve hangi sunucunun başarısız olduğunu göstermiyordu.
7. Health filtresi alıcıda geçen genel `Health` sözcüğünü dışladığı için gerçek kullanıcı adreslerini yanlışlıkla gizleyebilirdi.
8. Tarihler kültüre göre komut metnine çevrildiğinden farklı bölgesel ayarlarda kırılma riski vardı. Yeni sürüm gerçek `DateTime` nesnesi geçirir.
9. CSV her seferinde sabit `TrackingLogSearch.csv` adına `-Force` ile yazılıyor, önceki dosyayı sessizce ezebiliyordu.
10. CSV’de hesaplanmış sütunların adı anlamlı değildi; hata yakalama ve dışa aktarma sonucu doğrulaması yoktu.
11. Sonuçlar yalnızca ham alanlarla `Out-GridView` içinde açılıyor; olay anlamı, genel durum ve benzersiz ileti özeti sunulmuyordu.
12. Çok sayıda boş olay yordamı, oluşturucu kurtarma verisi, eski .NET 2.0 assembly referansları ve kullanılmayan yardımcı kod bakım yükünü ciddi biçimde artırıyordu.

## Resmî başvuru kaynakları

- [Get-MessageTrackingLog — Microsoft Learn](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/get-messagetrackinglog?view=exchange-ps)
- [Message tracking — Microsoft Learn](https://learn.microsoft.com/en-us/exchange/mail-flow/transport-logs/message-tracking)
- [Exchange sunucularına Remote PowerShell ile bağlanma — Microsoft Learn](https://learn.microsoft.com/en-us/powershell/exchange/connect-to-exchange-servers-using-remote-powershell?view=exchange-ps)
