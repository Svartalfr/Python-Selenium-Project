from selenium import webdriver
import xlsxwriter

#excel de bir dosya oluşturu
kitap=xlsxwriter.Workbook("yaseminkitapyurdu.xlsx")

#oluşan excel dosyama bir tane BOŞ sayfa ekler
sayfa=kitap.add_worksheet()

#bu fonkisyon oluşan boş excel sayfamın 1. satırana BAŞLIK ekler
def excelbaslik():
    sayfa.write(0,0,"SN")
    sayfa.write(0,1,"BARKOD")
    sayfa.write(0,2,"KİTAP ADI")
    sayfa.write(0,3,"YAZAR ADI")
    sayfa.write(0,4,"YAYIN EVİ")
    sayfa.write(0,5,"BASIM TARİHİ")
    sayfa.write(0,6,"DİL")
    sayfa.write(0,7,"SAYFA SAYISI")
    sayfa.write(0,8,"KAPAK TURU")
    sayfa.write(0,9,"KAĞIT TÜRÜ")
    sayfa.write(0,10,"İNDİRİMSİZ FİYAT")
    sayfa.write(0,11,"İNDİRİMLİ FİYAT")
    
    
    
#bu fonksiyon, SELENIUM ile kitapyurdu.com dan okunan tüm değerleri
#hem excel sayfama aktarır hem de ekranda gösterir
def excelyazdir(satirno,barkod,kitapadi,yazaradi,yayinevi,basimtarihi,dil,sayfasayisi,kapakturu,kagitturu,eskifiyat,yenifiyat):
    sayfa.write(satirno,0,str(satirno))
    sayfa.write(satirno,1,barkod)
    sayfa.write(satirno,2,kitapadi)
    sayfa.write(satirno,3,yazaradi)
    sayfa.write(satirno,4,yayinevi)
    sayfa.write(satirno,5,basimtarihi)
    sayfa.write(satirno,6,dil)
    sayfa.write(satirno,7,sayfasayisi)
    sayfa.write(satirno,8,kapakturu)
    sayfa.write(satirno,9,kagitturu)
    sayfa.write(satirno,10,eskifiyat)
    sayfa.write(satirno,11,yenifiyat)
    print("Satir no:",satirno)
    print("Barkod:",barkod)
    print("Kitap adi:",kitapadi)
    print("Yazar adı:",yazaradi)
    print("Basım Evi:",yayinevi)
    print("Basım Tarihi:",basimtarihi)
    print("Dili:",dil)
    print("Sayfa Sayısı:",sayfasayisi)
    print("Kapak Turu:",kapakturu)
    print("Kağıt Turu:",kagitturu)
    print("İndirimsiz Fiyat:",eskifiyat)
    print("İndirimli Fiyat:",yenifiyat)
    print("="*50)


#sayfaya dair başlıkları olusturma fonksiyonu
excelbaslik()

#firefox açılmasını sağlar
driver=webdriver.Firefox()

#satir no için başlangıç değeri verir
satirno=0

#kitapyurdu.com da her sayfa 100 kitap olmak üzere
#ortalama olarak 1000 sayfanın kitap verileri çekilebilir.
#kabaca 100.000 adet kitap verisi çekilebilir

maxsayfa=2    #çalışma uzun sürmesin diye 4 sayfa ile sınırlandırıldı
for sayfano in range(1,maxsayfa):

    #ulaşmak istediğimiz adresimizi
    adres="https://www.kitapyurdu.com/index.php?route=product/category&filter_category_all=true&path=1&filter_in_stock=1&sort=p.sort_order&order=ASC&limit=100&page="+str(sayfano)
    driver.get(adres)
    
    #a linkinin sınıcı pr-img-link olanlarını bul
    #onların içinde yer alan tüm img (resim) nesnlerini al
    #resmin de hem src hem de alt özerlliğini yazdır
    #kitapresmi, kitapadi
    resimler=driver.find_elements_by_xpath("//a[@class='pr-img-link']//img")
    #yayinevi alıancak
    yayinevleri=driver.find_elements_by_xpath("//div[@class='publisher']//a//span")

    #yazarlar alıancak
    yazarlar=driver.find_elements_by_xpath("//div[@class='author']//a//span")

    #ozelikler (sayfa sayısı, dili cildi vb)

    ozellikler=driver.find_elements_by_xpath("//div[@class='product-info']")

    eskifiyatlar=driver.find_elements_by_xpath("//div[@class='price-old price-passive']//span[@class='value']")
    yenifiyatlar=driver.find_elements_by_xpath("//div[@class='price-new ']//span[@class='value']")
    try:
        #her sayfa da 100 adet kitap listesi ele alnır
        for i in range(0,101):
            
            kitapadi=resimler[i].get_attribute("alt")
            yayinevi=yayinevleri[i].get_attribute('innerHTML')
            yazaradi=yazarlar[i].get_attribute('innerHTML')
            fiyaticin=eskifiyatlar[i].get_attribute('innerHTML')
            eskifiyat=fiyaticin[-5:]
            fiyaticin=yenifiyatlar[i].get_attribute('innerHTML')
            yenifiyat=fiyaticin[-5:]
            metin=ozellikler[i].get_attribute('innerHTML')
            parca1=metin.split("<br>")[0]
            basimtarihi=metin.split("<br>")[1]
            
            kacparca=len(parca1.split(" | "))
            if kacparca==5:
                satirno=satirno+1
                barkod=parca1.split(" | ")[0]
                dil=parca1.split(" | ")[1]
                sayfasayisi=parca1.split(" | ")[2]
                kapakturu=parca1.split(" | ")[3]
                kagitturu=parca1.split(" | ")[4]

                #aldığımız veriler, aşağıdaki fonksiyon ile EXCEL ve EKRANA yazdırılır
                excelyazdir(satirno,barkod,kitapadi,yazaradi,yayinevi,basimtarihi,dil,sayfasayisi,kapakturu,kagitturu,eskifiyat,yenifiyat)  
                
                
                
    except:
        print("Yeni Sayfaya Geçiliyor")
                
            


#veri çekme işlemi tamamlanınca kitap kapatılır    
kitap.close()
    
    

    

