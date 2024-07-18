import os
import pandas as pd
import xlwings as xw

# Konsoldan kullanıcıdan arama metni ve yeni metni al
arama_metni = input("Değiştirilecek metni girin: ")
yeni_metni = input("Yeni metni girin: ")

# Konsoldan dosya yolunu al
dosya_yolu = input("Dosya yolunu girin (boş bırakılırsa scriptin bulunduğu klasör taranacak): ").strip()
if not dosya_yolu:
    # Dosya yolunu python scriptinin bulunduğu klasör olarak ayarla
    dosya_yolu = os.path.dirname(os.path.abspath(__file__))

# Klasördeki tüm Excel dosyalarını tara
for dosya in os.listdir(dosya_yolu):
    if dosya.endswith('.xlsx') or dosya.endswith('.xls'):
        dosya_tam_yolu = os.path.join(dosya_yolu, dosya)
        # Excel dosyasını oku
        if dosya.endswith('.xlsx'):
            df = pd.read_excel(dosya_tam_yolu, engine='openpyxl')
        elif dosya.endswith('.xls'):
            with xw.App(visible=False) as app:
                wb = app.books.open(dosya_tam_yolu)
                df = wb.sheets[0].used_range.options(pd.DataFrame, index=False, header=True).value
                wb.close()
        
        # Tüm DataFrame'de arama_metni'ni yeni_metni ile değiştir
        if arama_metni in df.values:
            df = df.replace(arama_metni, yeni_metni)
            # Değişiklikleri aynı dosyaya yaz
            if dosya.endswith('.xlsx'):
                df.to_excel(dosya_tam_yolu, index=False, engine='openpyxl')
            elif dosya.endswith('.xls'):
                with xw.App(visible=False) as app:
                    wb = app.books.open(dosya_tam_yolu)
                    wb.sheets[0].range('A1').options(index=False, header=True).value = df
                    wb.save()
                    wb.close()
            # Konsola değişikliği yaz
            print(f"{dosya_tam_yolu} dosyasında '{arama_metni}' metni '{yeni_metni}' ile değiştirildi.")
        else:
            print(f"{dosya_tam_yolu} dosyasında '{arama_metni}' metni bulunamadı.")
        
print("Tüm dosyalarda değişiklik yapıldı.")
