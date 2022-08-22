# coding:utf-8 -*-

from PyQt5 import QtCore, QtGui, QtWidgets
from eclipse_main import Ui_MainWindow
import sys, time, os
from shutil import copy2
import xlwings as xw

class mywindow(QtWidgets.QMainWindow):    
    def __init__(self):
        super(mywindow, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        
        #Sinyal ve slotları bağlayan fonksiyyon çağırıyoruz
        self.triggerfinger()
        
        #Zaman ve çalışma yolu belirleme
        t = time.strftime("%d %m %Y")
        self.tarih = t.replace(" ", ".")
        self.ana_dizin = os.getcwd()
        self.desktop_path = os.path.expanduser("~") + "\\Desktop"
    
    """Buradaki kodları gui üzerindeki elamanlara yaptırtabilmek için
    fonksiyonuu hazırlayıp init içinde çağırarak gui ile bağlayacağız
    """
    def triggerfinger(self):
        #Anasayfa buton bağlantıları
        self.ui.petrobutton.clicked.connect(self.goto_petro)
        self.ui.lintecbutton.clicked.connect(self.goto_lintec)
        self.ui.hapagbutton.clicked.connect(self.goto_hapag)
        self.ui.lighthousebutton.clicked.connect(self.goto_lighthouse)
        self.ui.bqsbutton.clicked.connect(self.goto_bqs)
        self.ui.cleanlinesbutton.clicked.connect(self.goto_cleanlines)
        self.ui.lpgbutton.clicked.connect(self.goto_lpg)
        self.ui.bargeloadingbutton.clicked.connect(self.goto_bargeloading)
        self.ui.undeliveredbutton.clicked.connect(self.goto_undelivered)
        self.ui.quantitycontrolbutton.clicked.connect(self.goto_quantitycontrol)
        self.ui.onhirebutton.clicked.connect(self.goto_onhire)
        self.ui.supplierbargebutton.clicked.connect(self.goto_supplier)
        self.ui.ittbutton.clicked.connect(self.goto_itt)
        self.ui.chemicalbutton.clicked.connect(self.goto_chemical)
        self.ui.joblistbutton.clicked.connect(self.goto_joblist)
        
        #iç sayfadan anasayfaya dönmek için geri tuş bağlantıları
        self.ui.exitpetropage.clicked.connect(self.goto_homepage)
        self.ui.exitlintecpage.clicked.connect(self.goto_homepage)
        self.ui.exithapagpage.clicked.connect(self.goto_homepage)
        self.ui.exitlighthousepage.clicked.connect(self.goto_homepage)
        self.ui.exitbqspage.clicked.connect(self.goto_homepage)
        self.ui.exitcleanlinespage.clicked.connect(self.goto_homepage)
        self.ui.exitlpgpage.clicked.connect(self.goto_homepage)
        self.ui.exitbargeloadingpage.clicked.connect(self.goto_homepage)
        self.ui.exitundeliveredpage.clicked.connect(self.goto_homepage)
        self.ui.exitquantitcontrolpage.clicked.connect(self.goto_homepage)
        self.ui.exitonhirepage.clicked.connect(self.goto_homepage)
        self.ui.exitsupplierbargepage.clicked.connect(self.goto_homepage)
        self.ui.exitittpage.clicked.connect(self.goto_homepage)
        self.ui.exitchemicalpage.clicked.connect(self.goto_homepage)
        self.ui.exitjoblistpage.clicked.connect(self.goto_homepage)
        
        #fonksyionları slotlara bağlama
        
        self.ui.petroolusturbutton.clicked.connect(self.petro_apply)
    
    def goto_petro(self):
        self.ui.stackedWidget.setCurrentIndex(1)
    
    def goto_lintec(self):
        self.ui.stackedWidget.setCurrentIndex(2)
    
    def goto_hapag(self):
        self.ui.stackedWidget.setCurrentIndex(3)
    
    def goto_lighthouse(self):
        self.ui.stackedWidget.setCurrentIndex(4)
    
    def goto_bqs(self):
        self.ui.stackedWidget.setCurrentIndex(5)
    
    def goto_cleanlines(self):
        self.ui.stackedWidget.setCurrentIndex(6)
    
    def goto_lpg(self):
        self.ui.stackedWidget.setCurrentIndex(7)
    
    def goto_bargeloading(self):
        self.ui.stackedWidget.setCurrentIndex(8)
    
    def goto_undelivered(self):
        self.ui.stackedWidget.setCurrentIndex(9)
    
    def goto_quantitycontrol(self):
        self.ui.stackedWidget.setCurrentIndex(10)
    
    def goto_onhire(self):
        self.ui.stackedWidget.setCurrentIndex(11)
    
    def goto_supplier(self):
        self.ui.stackedWidget.setCurrentIndex(12)
    
    def goto_itt(self):
        self.ui.stackedWidget.setCurrentIndex(13)
    
    def goto_chemical(self):
        self.ui.stackedWidget.setCurrentIndex(14)
    
    def goto_joblist(self):
        self.ui.stackedWidget.setCurrentIndex(15)           
    
    def goto_homepage(self):
        self.ui.stackedWidget.setCurrentIndex(0)
    
    def petro_apply(self):
        # gemi adını alıp büyük harf yaptık
        gemi_ad = self.ui.petrogemiline.text().upper()
        
        #kaynak dosyasını değişkene atadık
        dosya_yolu = "\\libs\\files\\petro\\BunkerManager v1.12.xlsm"
        
        #masaüstü klasör adını belirledik
        dizin_ad = gemi_ad +" "+ self.tarih
        
        #masaüstü yolu ve klasör adını birleştirdik
        path = os.path.join(self.desktop_path, dizin_ad)
        
        #masaüstüne klasörü oluşturduk
        os.mkdir(path)
        
        #klasör içine kopyalanacak dosyalara için hedef ve kaynak tanımlamalarını yaptık
        kaynak = self.ana_dizin + dosya_yolu
        hedef = path + "\\BunkerManager v1.12 {} {}.xlsm".format(gemi_ad, self.tarih)
        copy2(kaynak, hedef)
        
        if (self.ui.checkBox_petro_lop.isChecked()):
            lop_source = "\\libs\\files\\petro\\lop.xlsx"
            lop_kaynak = self.ana_dizin + lop_source
            lop_hedef = path + "\\lop {} {}.xlsx".format(gemi_ad, self.tarih)
            copy2(lop_kaynak, lop_hedef)
        else:
            pass
        
        if (self.ui.checkBox_petro_prosedur.isChecked()):
            pro_source = "\\libs\\files\\petro\\petro prosedür.pdf"
            pro_kaynak = self.ana_dizin + pro_source
            pro_hedef = path + "\\Petro Prosedür.pdf"
            copy2(pro_kaynak, pro_hedef)
        else:
            pass
        
        if (self.ui.checkBox_petro_gt.isChecked()):
            gt_source = "\\libs\\files\\petro\\Gauging Tickets.pdf"
            gt_kaynak = self.ana_dizin + gt_source
            gt_hedef = path + "\\Petro Prosedür.pdf"
            copy2(gt_kaynak, gt_hedef)
        else:
            pass
    
    def yaz(self):
        print("mehmet")
        
    
    
        

app = QtWidgets.QApplication([])
application = mywindow()
application.show()
sys.exit(app.exec())