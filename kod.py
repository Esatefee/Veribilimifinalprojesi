import sys
import subprocess
import os
from collections import Counter
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import itertools

# 1. Kütüphane Yüklemesi
def install_and_import(package):
    try:
        __import__(package)
    except ImportError:
        print(f"-> '{package}' yukleniyor (Lutfen bekleyin)...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
install_and_import("matplotlib")
install_and_import("pandas")
install_and_import("openpyxl")
install_and_import("seaborn")
install_and_import("spacy")

# Türkçe dil modelini kontrol et ve indir
try:
    import tr_core_news_lg
except ImportError:
    print("-> Türkçe dil modeli indiriliyor (tr_core_news_lg)...")
    subprocess.check_call([sys.executable, "-m", "spacy", "download", "tr_core_news_lg"])
    import tr_core_news_lg

import spacy

sys.stdout.reconfigure(encoding='utf-8')

#2. AYARLAR: FILTRELER VE HARİTALAMA

BLACKLIST = {
    "Ah", "Oh", "Aman", "Tanrim", "Sen", "Ben", "O", "Biz", "Siz", "Onlar",
    "Bunu", "Sey", "Ama", "Fakat", "Eger", "Yok", "Var", "Evet", "Hayir",
    "Bey", "Bay", "Hanim", "Efendi", "Oyle", "Simdi", "Sonra", "Burada", 
    "Orada", "Neden", "Nicin", "Kim", "Nasil", "Cunku", "Vay", "Yaa", "Hah",
    "Bu", "Su", "Bir", "Hic", "Belki", "Lutfen", "Tam", "Daha", "En", "Cok", 
    "Az", "Baska", "Hangi", "Kendi", "Kendisi", "Biri", "Diye", "Şey",
    "Son", "Dun", "Bugun", "Yarin", "Lyon", "Misir", "Seyler", "Dün"
}

MAPPING = {
    "Rodya": "Raskolnikov", "Rodion": "Raskolnikov", "Rodion Romanovic": "Raskolnikov",
    "Raskolnikovun": "Raskolnikov", "Raskolnikova": "Raskolnikov Ailesi",
    "Sonya": "Sofya Semyonovna", "Sonyanin": "Sofya Semyonovna", "Sofya": "Sofya Semyonovna",
    "Dunya": "Avdotya Romanovna", "Dunyanin": "Avdotya Romanovna", "Avdotya": "Avdotya Romanovna",
    "Razumihin": "Dmitri Prokofyic", "Dmitri": "Dmitri Prokofyic",
    "Lujin": "Pyotr Petrovic", "Petrovic": "Pyotr Petrovic", "Pyotr": "Pyotr Petrovic",
    "Porfiriy": "Porfiriy Petrovic", "Katerina": "Katerina Ivanovna",
    "Andrey": "Lebezyatnikov", "Lebezyatnikov": "Lebezyatnikov",
    "Alyona": "Alyona Ivanovna", "Lizaveta": "Lizaveta Ivanovna",
    "Zosimov": "Doktor Zosimov", "Zametov": "Aleksandr Grigoryevic",
    "Nastas": "Nastasya (Hizmetci)", "Nastasya": "Nastasya (Hizmetci)",
    "Amal": "Amalya Ivanovna", "Amalya": "Amalya Ivanovna",
    "Polina": "Polina Mihaylovna",
    "Petersburg": "St. Petersburg", "Petersburgu": "St. Petersburg", "Petersburga": "St. Petersburg",
    "Sibirya": "Sibirya", "Sibiryaya": "Sibirya",
    "Hayman": "Samanpazari", "Neva": "Neva Nehri", 
    "Karakol": "Polis Karakolu", "Karakola": "Polis Karakolu"
}

FORCE_PERSON = {
    "Nastasya (Hizmetci)", "Amalya Ivanovna", "Raskolnikov", "Sofya Semyonovna", 
    "Avdotya Romanovna", "Dmitri Prokofyic", "Katerina Ivanovna"
}

FORCE_LOC = {
    "St. Petersburg", "Sibirya", "Samanpazari", "Neva Nehri", "Polis Karakolu", "Meyhane"
}

# 3. YARDIMCI FONKSIYONLAR 
def clean_entity(text):
    clean_text = text.replace("\u2014", "").replace("-", "").replace("'", "").replace("\u2019", "").strip()
    
    # Ek temizliği
    suffixes = ["da", "de", "ta", "te", "ya", "ye", "dan", "den", "nin", "nun", "yi", "yu", "un", "in", "a", "e"]
    if len(clean_text) > 4:
        for suf in suffixes:
            if clean_text.endswith(suf):
                clean_text = clean_text[:-len(suf)]
                break 

    # Mapping kontrolü
    if clean_text in MAPPING: return MAPPING[clean_text]
    for key, val in MAPPING.items():
        if key in clean_text.split(): return val
            
    # Blacklist ve uzunluk kontrolü
    if len(clean_text) < 3 or clean_text in BLACKLIST: return None
    return clean_text

def create_plot(data, title, filename, output_type="bar", series_data=None, color='#4a90e2'):
    plt.figure(figsize=(12, 7))
    if output_type == "bar":
        names = [x[0] for x in data]
        values = [x[1] for x in data]
        plt.bar(names, values, color=color, edgecolor='black')
        plt.xticks(rotation=45, ha='right')
        plt.ylabel("Gecme Sayisi")
        
    elif output_type == "line":
        top_items = [x[0] for x in data[:5]]
        x_axis = range(1, len(series_data) + 1)
        markers = ['o', 's', '^', 'D', 'v']
        for i, item in enumerate(top_items):
            y_axis = [chunk[item] for chunk in series_data]
            plt.plot(x_axis, y_axis, marker=markers[i%len(markers)], linewidth=2.5, label=item)
        plt.xlabel("Romanin Bolumleri (Bastan Sona)")
        plt.ylabel("Gecme Sayisi")
        plt.legend(bbox_to_anchor=(1.02, 1), loc='upper left')

    plt.title(title, fontsize=14)
    plt.grid(True, linestyle='--', alpha=0.5)
    plt.tight_layout()
    plt.savefig(filename, dpi=300)
    print(f"   -> Grafik Kaydedildi: {filename}")
    plt.close()

def create_heatmap(df, title, filename, color_map="YlGnBu"):
    plt.figure(figsize=(14, 9))
    top_rows = df.sum(axis=1).nlargest(15).index
    top_cols = df.sum(axis=0).nlargest(15).index
    df_filtered = df.loc[top_rows, top_cols]

    sns.heatmap(df_filtered, annot=True, fmt="d", cmap=color_map, linewidths=.5)
    plt.title(title, fontsize=16)
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.savefig(filename, dpi=300)
    print(f"   -> Isı Haritası Kaydedildi: {filename}")
    plt.close()

# 4. ANA PROGRAM
def main():
    try:
        print("1. Model yukleniyor...")
        nlp = tr_core_news_lg.load()
        nlp.max_length = 3000000 
        
        dosya_adi = "suc_ve_ceza.txt"
        print(f"2. {dosya_adi} okunuyor...")
        
        with open(dosya_adi, "r", encoding="utf-8", errors="ignore") as f:
            text = f.read()
        
        print("3. Detaylı Analiz Başlıyor (Cümle Bazlı)...")
        # Analiz için metni parçalara ayır
        chunk_size = 100000
        chunks = [text[i:i+chunk_size] for i in range(0, len(text), chunk_size)]
        
        all_p, all_l, all_t = [], [], []
        series_p, series_l, series_t = [], [], []
        
  
        char_char_rel = []  
        char_loc_rel = []   
        char_time_rel = []  
        pair_loc_rel = []   

        for i, chunk in enumerate(chunks):
            doc = nlp(chunk)
            chunk_p, chunk_l, chunk_t = [], [], []
            
            for sent in doc.sents:
                s_p, s_l, s_t = set(), set(), set()
                
                for ent in sent.ents:
                    cleaned = clean_entity(ent.text)
                    if not cleaned: continue
                    
                    # Etiket Zorlama
                    lbl = "PERSON" if cleaned in FORCE_PERSON else ("LOC" if cleaned in FORCE_LOC else ent.label_)
                    
                    if lbl == "PERSON":
                        s_p.add(cleaned); chunk_p.append(cleaned); all_p.append(cleaned)
                    elif lbl in ["GPE", "LOC"]:
                        s_l.add(cleaned); chunk_l.append(cleaned); all_l.append(cleaned)
                    elif lbl in ["DATE", "TIME"]:
                        s_t.add(cleaned); chunk_t.append(cleaned); all_t.append(cleaned)
                # 1. Karakter - Karakter
                if len(s_p) > 1:
                    pairs = list(itertools.combinations(sorted(list(s_p)), 2))
                    char_char_rel.extend(pairs)
                    
                    # 2. Karakter Çifti - Mekan
                    if s_l:
                        for pair in pairs:
                            pair_str = f"{pair[0]} & {pair[1]}"
                            for loc in s_l:
                                pair_loc_rel.append((pair_str, loc))
                
                # 3. Karakter - Mekan
                for c in s_p:
                    for l in s_l: char_loc_rel.append((c, l))
                
                # 4. Karakter - Zaman
                for c in s_p:
                    for t in s_t: char_time_rel.append((c, t))
            
            # Zaman serisi için kaydet
            series_p.append(Counter(chunk_p))
            series_l.append(Counter(chunk_l))
            series_t.append(Counter(chunk_t))
            
            print(f"   -> Parça {i+1}/{len(chunks)} işlendi...")

        print("4. İstatistikler Hesaplanıyor...")
        top_p = Counter(all_p).most_common(10)
        top_l = Counter(all_l).most_common(10)
        top_t = Counter(all_t).most_common(10)

        # Matrisler
        df_cc = pd.DataFrame(char_char_rel, columns=["K1", "K2"])
        # Simetrik matris yapma
        df_cc_full = pd.concat([df_cc, df_cc.rename(columns={"K1":"K2", "K2":"K1"})])
        mat_cc = pd.crosstab(df_cc_full.K1, df_cc_full.K2)
        
        df_cl = pd.DataFrame(char_loc_rel, columns=["K", "M"])
        mat_cl = pd.crosstab(df_cl.K, df_cl.M)
        
        df_ct = pd.DataFrame(char_time_rel, columns=["K", "Z"])
        mat_ct = pd.crosstab(df_ct.K, df_ct.Z)
        
        df_pair_loc = pd.DataFrame(pair_loc_rel, columns=["Cift", "Mekan"])
        mat_pair_loc = pd.crosstab(df_pair_loc.Cift, df_pair_loc.Mekan)

        # 5. ÇIKTILAR (GRAFİKLER VE EXCEL)
        print("5. Görseller Oluşturuluyor...")
        
        
        if top_p:
            create_plot(top_p, "En Çok Geçen Karakterler", "1_karakter_bar.png", "bar", color="#3498db")
            create_plot(top_p, "Karakterlerin Hikaye Akışı", "2_karakter_line.png", "line", series_data=series_p)
        
        if top_l:
            create_plot(top_l, "En Sık Geçen Mekanlar", "3_yer_bar.png", "bar", color="#e67e22")
            create_plot(top_l, "Mekanların Kullanım Sıklığı", "4_yer_line.png", "line", series_data=series_l)
            
        if top_t:
            create_plot(top_t, "En Sık Geçen Zaman İfadeleri", "5_zaman_bar.png", "bar", color="#9b59b6")
            create_plot(top_t, "Zaman Kavramının Değişimi", "6_zaman_line.png", "line", series_data=series_t)

        # Isı Haritaları 
        if not mat_cl.empty:
            create_heatmap(mat_cl, "Karakter - Mekan İlişkisi (Kim Nerede?)", "7_mekan_heatmap.png", "YlOrBr")
        
        if not mat_ct.empty:
            create_heatmap(mat_ct, "Karakter - Zaman İlişkisi (Kim Ne Zaman?)", "8_zaman_heatmap.png", "Purples")
            
        if not mat_cc.empty:
            create_heatmap(mat_cc, "Sosyal Ağ Analizi (Kim Kiminle?)", "9_karakter_etkilesim_heatmap.png", "Reds")
            
        if not mat_pair_loc.empty:
            create_heatmap(mat_pair_loc, "Karakter Çiftleri ve Mekan (Birlikte Neredeler?)", "12_karakter_ciftleri_mekan_heatmap.png", "YlGnBu")

        # Excel Kaydı
        print("6. Excel Raporu Kaydediliyor...")
        base_filename = "VERI_BILIMI_FINAL_PROJESI_ANALIZ"
        
        for attempt in range(10):
            try:
                fname = f"{base_filename}.xlsx" if attempt == 0 else f"{base_filename}_{attempt}.xlsx"
                with pd.ExcelWriter(fname, engine='openpyxl') as writer:
                    pd.DataFrame(top_p, columns=["Karakter", "Sayi"]).to_excel(writer, sheet_name="Karakterler", index=False)
                    pd.DataFrame(top_l, columns=["Yer", "Sayi"]).to_excel(writer, sheet_name="Yerler", index=False)
                    pd.DataFrame(top_t, columns=["Zaman", "Sayi"]).to_excel(writer, sheet_name="Zamanlar", index=False)
                    
                    if not mat_cc.empty: mat_cc.to_excel(writer, sheet_name="Matris-Karakterler")
                    if not mat_cl.empty: mat_cl.to_excel(writer, sheet_name="Matris-Mekan")
                    if not mat_pair_loc.empty: mat_pair_loc.to_excel(writer, sheet_name="Matris-Cift-Mekan")
                
                print(f"   -> Excel Başarıyla Kaydedildi: {fname}")
                break
            except PermissionError:
                print(f"   ! {fname} dosyası açık, farklı isim deneniyor...")

        print("\n" + "="*50)
        print(" TEBRİKLER! TÜM ANALİZLER TAMAMLANDI.")
        print(" Rapor için gerekli tüm görseller ve Excel dosyası hazır.")
        print("="*50)

    except Exception as e:
        print(f"\nKRİTİK HATA OLUŞTU: {e}")

if __name__ == "__main__":
    main()
