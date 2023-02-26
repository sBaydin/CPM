def CPM():
    import sys
    import pandas as pd
    import numpy as np
    import tkinter as tk
    from tkinter import filedialog
    from tkinter import messagebox
    
    # Uyarı penceresini aç
    answer = messagebox.askquestion("Dosya Seçimi", "Kritik yol hesabı için dosyayı seçmek istiyor musunuz?")
    # Kullanıcının seçimine göre işlem yap
    if answer == "yes":
        # Dosya seçme işlemini gerçekleştir
        root = tk.Tk()
        root.withdraw()  # Tk penceresini gizle
        file_path = filedialog.askopenfilename(filetypes=[("Excel", ".xlsx"),("All files", "*.*")])  # Dosya seçme penceresini aç
        if not file_path:  # Dosya seçilmedi ise programı sonlandır
            raise SystemExit
    else:
        # Programdan çıkılacak
        sys.exit()

    # CSV dosyasından verileri yükleme
    df = pd.read_excel(file_path)
    df["Öncül Faaliyeti"].fillna(value='0', inplace=True)
    df["Öncül Faaliyeti"] = df["Öncül Faaliyeti"].apply(lambda x: [int(i) for i in x.split(",")] if type(x) == str else [int(x)])
    """ Ardılları excelde belirtirsem işe yarar ama gerek yok adj_matrix1 ile ardılları buldurdum...
    df["Ardıl Faaliyeti"].fillna(value='0', inplace=True)
    df["Ardıl Faaliyeti"] = df["Ardıl Faaliyeti"].apply(lambda x: [int(i) for i in x.split(",")] if type(x) == str else [int(x)])
    """

    # Faaliyet öncüllük ve ardıllık matrisini oluşturma
    num_activities = len(df)
    adj_matrix = np.zeros((num_activities, num_activities))     #*Öncüller Matrisi
    adj_matrix1 = np.zeros((num_activities, num_activities))    #*Ardıllar Matrisi

    for i, row in df.iterrows():
        for j in row["Öncül Faaliyeti"]:
            if j != 0:
                if type(j) == int:
                    adj_matrix[i, j-1] = 1
                elif type(j) == list:
                    for k in j:
                        adj_matrix[i, k-1] = 1
                        
    for i, row in df.iterrows():
        for j in row["Öncül Faaliyeti"]:
            if j != 0:
                if type(j) == int:
                    adj_matrix1[j-1, i] = 1
                elif type(j) == list:
                    for k in j:
                        adj_matrix1[k-1, i] = 1

    # Aktivitelerin ES ve EF değerlerini hesapla
    es = np.zeros(num_activities)
    ef = np.zeros(num_activities)

    # Öncülü olmayan faaliyetlerin ES değerlerini 0 olarak ata ve EF değerlerini bul
    for i in range(num_activities):
        if not np.any(adj_matrix[i]):
            es[i] = 0
            ef[i] = df.at[i, "Süresi"]

    # Diğer aktivitelerin ES ve EF değerlerini hesapla
    for i in range(num_activities):
        if np.any(adj_matrix[i, :]):
            max_predecessor_ef = np.max(ef[adj_matrix[i, :] == 1])
            es[i] = max_predecessor_ef
            ef[i] = es[i] + df.at[i, "Süresi"]

    # Geri Gidiş (Backward Pass Yöntemi)
    # Geriye doğru hesaplama yaparak LS ve LF değerlerini bulma
    ls = np.zeros(num_activities)
    lf = np.zeros(num_activities)

    # En son bitirilmesi gereken faaliyetin indeksi
    last_activity = np.argmax(ef)

    # Ardılı olmayan faaliyetlerin LF ve LS değerleri EF ve ES değerlerine eşitlenir kukla faaliyetler ve kritik yol sonu
    for i in range(num_activities):
        if not np.any(adj_matrix1[i,:]):
            lf[i] = ef[last_activity]
            ls[i] = ef[last_activity] - df.at[last_activity, "Süresi"]

    # LF ve LS değerlerinin bulunması
    for i in range(num_activities -1, -1, -1):
        successors = np.where(adj_matrix1[i, :])[0]
        if len(successors) == 0:
            lf[i] = ef[last_activity]
        else:
            lf[i] = min([ls[j] for j in successors])
            
        predecessors = np.where(adj_matrix[:, i])[0]
        if len(predecessors) == 0:
            ls[i] = lf[i] - df.at[i, "Süresi"]
        else:
            ls[i] = lf[i] - df.at[i, "Süresi"]

    # Toplam bolluk (TB)
    tb = lf - ef

    # Serbest bolluk (SB) hesapla #! Hatalı hesaplıyor gibi kontorl edilmesi gerek
    sb = np.zeros(num_activities)
    for i in range(num_activities):
        successors = np.where(adj_matrix1[i, :])[0]
        if len(successors) == 0:
            sb[i] = np.inf
        else:
            sb[i] = min([ls[j] - ef[i] for j in successors])

    # Kritik yolun süresini hesapla
    cp_duration = int(max(ef))

    # Kritik yoldaki faaliyetleri bul
    cp_activities = []
    for i in range(num_activities):
        if tb[i] == 0:
            activity_no = df.at[i, "Faaliyet No"]
            cp_activities.append(activity_no)

    # Verileri DataFrame'e ekle
    df["ES"] = es
    df["EF"] = ef
    df["LS"] = ls
    df["LF"] = lf
    df["TB"] = tb
    df["SB"] = sb

    # DataFrame bilgileri
    print(df)
    print("-" * 150)
    print(f"Projenin Toplam Bitiş Süresi         : {cp_duration}")
    print(f"Kritik Yoldaki Aktivitelerin Listesi : {cp_activities}")

    # Dosya seçme işlemi için Tkinter dosya seçme penceresini açma
    root = tk.Tk()
    root.withdraw() # Tkinter penceresini gizle
    file_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[("Excel", ".xlsx"),("All files", "*.*")])

    # Dosya yolu alındıktan sonra, verileri seçilen dosyaya kaydetme
    if file_path:
        df.to_excel(file_path, index=False)
        
CPM()