# Excel VBA Tabanlı Satın Alma ve Bütçe Yönetim Sistemi
1.	GİRİŞ
Geliştirilen ERP modeli talep oluşturma, bütçe kontrolü, onay mekanizması ve satın alma süreçlerini tek bir sistemde toplayan bir yapı sunmaktadır. Modülün temel işlevi birimlerin yıllık bütçe limitlerini kontrol etmek aşmasını engellemektir ve limit üstü harcamaların üst kademeye onaya sunulmasıdır. Kurum içerisinde yer alan IT, İnsan Kaynakları ve diğer birimlerin ofis ihtiyaçlarını karşılamak amacıyla oluşturduğu satın alma taleplerinin dijital olarak alınması ve toplanması, bütçe kontrolünden geçirilmesi ve gerektiğinde de müdür onayına sunulması ve satın alma birimi tarafından işleme alınmasını amaçlar.
Kullanılacak Hesaplamalar:
Toplam Tutar = Adet x Birim Fiyat
Toplam Tutar > Birim Kalan Bütçe  İse Reddedilir
Toplam Tutar <= Birim Kalan Bütçe ise Süreç Devam eder
Sistemde tanımlı bir eşik değeri olucak Esik = 5000 TL
Eğer Toplam Tutar > Esik ise MüdürOnayDurumu=Gerekli  olur ve talep Müdür onay ekranına düşer, Aksi halde direk satın alma ekranına düşer .
Satın alma tamamlandıktan sonra Birim Kalan Bütçe formülü güncellenir: Birim_Kalan_Bütce_Yeni = Birim_Kalan_Bütce_Eski – ToplamTutar 
Bir Talep Onay Bekliyor veya Tamamlandı Durumundayken aynı kullanıcı tarafından talep verilmesine sınırlama getirilebilir.
Her kullanıcının yapmış olduğu işlemler kaydedilecektir, harcama geçmişi görülebilecektir. Bu yapı sayesinde satın alma ve bütçe yönetimi süreçleri kontrollü, şeffaf ve izlenebilir bir sistem üzerinden yürütülmektedir.

 
2.	AMAÇ
Bu modülün amacı, kurum içinde yer alan birimlerin satın alma ihtiyaçlarını düzenli ve kontrollü şekilde yönetmektir. Birimlerin yıllık bütçelerini takip edebilmesi ona göre harcama yapabilmesi. Harcama yaparken şirketin belirlediği bedel üzerinde bir harcama yapmalarına kısıt getirmek ve yıllık harcama bütçelerini aşmalarını engellemek. Üst yönetimin birimlerin bütçe ve taleplerini kontrol altında tutabilmesi ve gözlemleyebilmesi, Böylece hem mali disiplin sağlamakta hem de satın alma sürecinde kurumsal verimlilik sağlamak.
3.	SİSTEME GENEL BAKIŞ
Sistem kullanıcı giriş ekranı ile kullanıcının giriş yapması ile başlamaktadır. Kullanıcı giriş yaptığı rol ile belirli sayfalara erişebilir. Sayfalar ilk olarak birimlerin ofis ihtiyaçları için talep oluşturması ile başlamaktadır. Ürün adı, adet ve fiyat bilgisi girildiğinde sistem otomatik olarak tutarı hesaplar. Ardından talep edilen tutar, birimin kalan bütçesi ile karşılaştırılır, bütçe yeterliyse sistem devam eder. Sonrasında eğer talep edilen ürünün tutarı belirlenen eşik değerinden yüksek ise üst yönetime iletilir, üst yönetim talebi onayladığında ürün alımı onaylanır ve kalan bütçe güncellenir. Onaylanan talepler satın alma birimine düşer ve onay bekler sistem satın alındığına dair uyarılar verir ve bütçe otomatik güncellenir. Geçmiş satın alımlarda üst yönetim tarafından görüntülenebilmektedir. Ayrıca yönetici rolündeki kullanıcılar programdan bütçe yönetimi kısmı ile departmanların bütçelerini belirleyebiliyor.
4.	SİSTEM TANITIMI 
Sistem, Excel VBA kullanılarak geliştirilen yedi adet kullanıcı formu ve bu formlar aracılığıyla yönetilen veri sayfalarından oluşmaktadır. Sistem modüler bir yapıya sahip olup, her kullanıcı formu belirli bir işlevi yerine getirmektedir. Her kullanıcı belli formlara erişebilmektedir.
4.1	Ekran Tasarımları

Kullanıcı Giriş İşlemi: Kullanıcı giriş işlemi başlangıç ekranıdır ve tüm kullanıcıların kimlik doğrulaması bu ekran üzerinden yapılmaktadır. Kullanıcı adı ve şifresini girdikten sonra sisteme giriş yapabilir. Giriş işlemi sırasında boş alan kontrolü ve hatalı giriş uyarı mesajı kontrolü yapılmaktadır. Giriş yaptıktan sonra sistem kullanıcının rolünü kontrol ederek tek bir ekranda yer alan bölümlerin görünürlüğünü ayarlamaktadır ve kullanıcı ana menü ekranına yönlendirilmektedir. Kullanıcının rolüne göre erişim yetkisi olan alanlar gösterilmektedir.
<img width="765" height="373" alt="image" src="https://github.com/user-attachments/assets/21e5ae32-3a57-44ab-8b5e-e2583024075c" />









YENİ KULLANICI KAYDI:
Yeni kullanıcı kaydı oluşturmak için oluşturulmuş bir sayfadır, Burada kullanıcı adı benzerlik kontrolü yapılıyor ve kullanıcı şifresini oluşturuyor ve departmanını seçiyor. Yeni eklenen kullanıcılar sisteme güvenlik amacıyla otomatik olarak personel olarak ekleniyor.
<img width="567" height="325" alt="image" src="https://github.com/user-attachments/assets/fffb6524-e861-4ec5-b095-d022eb7f8608" />









ANA MENÜ:
Ana menü ekranında kullanıcıların görüp işlem yapabileceği alanlar rol bazlı olarak sınırlandırılmıştır. Kullanıcı giriş yaptıktan sonra yalnızca yetkili olduğu modüllere ait butonları görüntüleyebilir. Bu yapı sayesinde yetkisiz erişimlere karşı güvenlik önlemi alınmıştır. Sistem rollerinin erişebileceği modüller;
Personel: Talep Oluşturma
Müdür: Talep Onaylama ve Raporlama
Satın Alma: Satın alma işlemleri ve Raporlama
Finans/Yönetim: Tüm modüller
<img width="485" height="428" alt="image" src="https://github.com/user-attachments/assets/79b0c541-dd24-4e5b-8c4c-17c7fccf733b" />











TALEP OLUŞTURMA İŞLEMİ:
Personel olarak giriş yapan kullanıcılar için aktif bir bölümdür. Personelin isim soy isim, bağlı bulunduğu birim bilgisi otomatik olarak gözükmekte. Kullanıcı ürün adı, adet ve birim fiyat bilgilerini girdikten sonra sistem tarafından toplam tutar otomatik olarak hesaplanmaktadır. Personel “Siparişi Gönder” butonu ile talebi taslak olarak kaydedebilir. Boş alan kontrolü vardır. Göndere basınca önce birimin bütçesi ile karşılaştırılır. Bütçe yetersiz ise talep reddedilir ve uyarı mesajı görülür. Bütçe yeterli ise belirlenen eşik değerini aşmıyorsa kullanıcı doğrudan satın alma sürecine gönderilir. Eşik değeri aşıyorsa talep otomatik olarak müdür onayına gönderilir.
<img width="605" height="447" alt="image" src="https://github.com/user-attachments/assets/b694b03e-b922-4308-8ace-f1a85ac02fde" />













MÜDÜR ONAY İŞLEMİ:
Müdür onay işlemi yalnızca müdür rolü ile giriş yapan kullanıcılara görünür. Müdürün sorumlu olduğu personel tarafından gönderilen talepler ve “Müdür Onayı Bekliyor” durumundaki talepler listelenir. Reddet butonuna basması durumunda talep iptal edilir ve talep eden personelin ekranında “reddedildi” olarak gözükür. Müdür onayla butonuna basması durumunda talep tutarı birimin bütçesinden düşer ve müdür ekranında talebin onaylanıp satın alma birimine yönlendirildiği mesajı çıkar. Onaylanan talepler satın alma biriminin işlem ekranına yönlendirilir.
<img width="391" height="450" alt="image" src="https://github.com/user-attachments/assets/f9c86dd0-f5a2-460f-920d-4b41d112b9f8" />




SATIN ALMA İŞLEMİ:
Satın alma işlemi yalnızca satın alma rolü ile giriş yapan kullanıcılar tarafından kullanılabilmektedir. Bu bölümde müdür tarafından onaylanmış veya eşik değer altında olduğu için müdür onayı gerektirmeyen talepler listelenmektedir. Satın alma personeli gerekli kontrolü sağladıktan sonra “satın al” butonuna tıklayarak işlemi tamamlar. İşlem tamamlandığında sistem otomatik olarak birimin kalan bütçesini günceller ve ekranda “Tebrikler satın alma işlemi tamamlandı, … bütçesinden … düşüldü” mesajını verir. Satın alma işlemi tamamlanınca durum otomatik olarak “tamamlandı” olarak güncellenir.
<img width="681" height="308" alt="image" src="https://github.com/user-attachments/assets/867b7c1a-0ca5-49df-b90b-59b5d6ad20eb" />

















BÜTÇE YÖNETİMİ:
Bu ekran finans veya yönetici rolüne sahip kullanıcılar için geliştirilmiştir. Departman bazlı bütçe görüntüleme, güncelleme ve kontrol işlemleri bu ekran üzerinden yapılmaktadır. Yeni girilen bütçe tutarının, daha önce harcanan tutardan düşük olmaması sistem tarafından kontrol edilerek bütçe tutarlılığı sağlanmaktadır.
<img width="547" height="335" alt="image" src="https://github.com/user-attachments/assets/5e6464e8-4ad3-49a3-9ea1-3597324a00da" />



 DASHBOARD VE RAPORLAMA EKRANI:
Sistem kapsamında yöneticilerin ve yetkili kullanıcıların satın alma ve bütçe durumlarını anlık olarak takip edebilmesi için tasarlanmıştır. Bu dashboard pivot tablo ,grafik ve dilimleyici özellikleri ile makrolarla birlikte oluşturulmuştur. Dashboard ekranında sistemde yer alan data sayfasından alınan veriler özetlenerek görsel hale getirilmiştir. Dashboard da ki yenile tuşu ile bilgiler yenilenebilmektedir, Ana menüye dön tuşu ile ana menüye geri dönülür.

Dashboard üzerinde yer alan başlıca bileşenler:

	TOPLAM HARCAMA DAĞILIMI GRAFİĞİ:
Departman bazlı toplam harcama tutarları pasta grafik ile gösterilmektedir. Bu grafik sayesinde hangi departmanın toplam harcama içerisindeki payı görsel olarak analiz ediliyor.
	TOPLAM KALAN BÜTÇE GRAFİĞİ:
Departmanların kalan bütçeleri sütun grafiği ile gösterilerek bütçe durumunun hızlı bir şekilde analiz edilmesi sağlanmıştır.
	BÜTÇE VE HARCAMA TABLOSU:
Departmanlara ait toplam bütçe, harcanan tutar ve kalan bütçe bilgileri tablo halinde sunulmaktadır.
	Onay Durumu ve Departman Filtreleri:
Dashboard üzerinde bulunan dilimleyiciler aracılığı ile veriler departman, toplam tutar ve talep onay durumuna göre filtrelenebilmektedir. 
<img width="269" height="255" alt="image" src="https://github.com/user-attachments/assets/f67c4fc3-65cd-45da-bfb5-59805fd5177c" />

<img width="1003" height="470" alt="image" src="https://github.com/user-attachments/assets/e88126e2-3177-4af9-8b6a-166910c4d4f3" />












5.	SONUÇ

Sonuç olarak birimlerin ofis ihtiyaçlarına ilişkin talepler sistem üzerinden alınarak, bütçe kontrolü, müdür onayı ve satın alma birimi tarafından izlenebilir bir süreç haline getirilmiştir. Bu yapı sayesinde hem talep oluşturma hem de onay işlemlerini hızlıca gerçekleştirilebilmekte. Sistem üzerinden yapılan işlemler kayıt altına alındığından, bütçe kullanımı ve satın alma süreçleri şeffaf bir şekilde yönetilebilmekte ve olası hataların önüne geçilebilmektedir. Bu modül ile kurum içi taleplerin düzenli, kontrollü ve izlenebilir yapıda yürütülmesi sağlanmıştır.
