# Data-Parser
 -------------------------- For Turkish ----------------------------------------------------------------------------------------------------------------------
Bu Python kodu, bir metin dosyasından veri okuyarak, bu verileri filtreleyerek ve farklı kategorilere ayırarak bir Excel dosyasına dönüştüren bir betiği temsil eder. Ayrıca bu kod, Excel dosyasını oluştururken belirli ayarları ve işlevleri kullanır. İşte bu kodun yaptığı temel işlevler:

Kullanıcıya bir dosya seçmesi için bir dosya açma penceresi gösterir.
Seçilen dosyanın yolunu alır ve dosya içeriğini okur.
Belirli anahtar kelimelere göre verileri sınıflandırır: firstData, secondData, ve otherData.
Excel dosyasını oluşturur ve sayfaları ekler: "first_data", "second_data", ve "other_data".
Verileri ilgili sayfalara ekler ve sütun genişliklerini ayarlar.
"Freeze Panes" özelliği ile Excel sayfalarında başlıkları sabitler.
"AutoFilter" özelliği ile filtreler ekler.
Oluşturulan Excel dosyasını belirtilen adla kaydeder.
Bu kod, özellikle büyük metin verilerini filtrelemek ve düzenlemek ve bunları Excel formatına dönüştürmek isteyenler için kullanışlıdır. Örneğin, bir veri raporu oluşturmak veya analiz etmek için metin verilerini düzenlemek için kullanabilirsiniz.

Kodu kullanmadan önce dikkat etmeniz gereken bazı noktalar:

firstDataString ve secondDataString değişkenlerine, verileri sınıflandırmak için kullanılacak belirli anahtar kelimeleri eklemeniz gerekmektedir.
Verilerin metin dosyasındaki formatına ve verilerin nasıl ayrıldığına dikkat etmelisiniz (örneğin, verilerin virgülle ayrıldığı varsayılmıştır).
Verilerin nasıl sınıflandırılacağını ve Excel sayfalarına nasıl ekleneceğini özelleştirebilirsiniz.
Bu kodu kullanarak, metin verilerini daha düzenli ve analiz edilebilir bir Excel formatına dönüştürebilirsiniz.




 ---------------------------- For English ------------------------------------------------------------------------------------------------------------
 
 This Python code represents a script that reads data from a text file, filters the data based on specific keywords, and categorizes it into different categories, then converts it into an Excel file. Additionally, this code uses specific settings and functionalities when creating the Excel file. Here are the basic functions of this code:

Displays a file open dialog to allow the user to choose a file.
Retrieves the path of the selected file and reads its content.
Categorizes the data based on specific keywords: firstData, secondData, and otherData.
Creates an Excel file and adds sheets: "first_data," "second_data," and "other_data."
Adds the data to the respective sheets and sets column widths.
Freezes headers in Excel sheets using the "Freeze Panes" feature.
Adds filters using the "AutoFilter" feature.
Saves the created Excel file with the specified name.
This code is useful, especially for those who want to filter and organize large text data and convert it into an Excel format. For example, you can use it to create a data report or to prepare text data for analysis.

Before using the code, here are some points to consider:

You should add specific keywords to the firstDataString and secondDataString variables to categorize the data.
Pay attention to the format of data in the text file and how the data is separated (it assumes data is separated by commas).
You can customize how the data is categorized and added to Excel sheets.
By using this code, you can convert text data into a more organized and analyzable Excel format.
