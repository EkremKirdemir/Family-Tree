PROGRAMLAMA LABORATUVARI

PROJE 3

EKREM KIRDEM˙IR

Kocaeli Univer¨ sitesi Muhendislik¨ Fakultesi¨ Bilgisayar Muhendisli¨ gi˘ Bol¨ um¨ u¨ 2. sınıf kirdemir.ekrem@gmail.com

210202017

Abstract—Bu program, Excel dosyasından verileri oku- yarak aile uy¨ eleri arasındaki ilis¸kileri ve bu ilis¸kileri goster¨ en grafikleri olus¸turur. Program, bir Form1 sınıfı ve bu sınıfın ic¸inde CreateFamilyTrees ve CreateButtonsList adlı iki yontem¨ ic¸erir. CreateFamilyTrees yontemi,¨ Excel dosyasını ac¸arak sayfalarındaki verileri okur ve PersonUI sınıfını kullanarak grafikselaile agac˘ ¸ları olus¸turur. Create- ButtonsList yontemi¨ ise GUI ic¸in butonları olus¸turur. Pro- gramın sonucu, aile uy¨ eleri arasındaki ilis¸kileri goster¨ en soy agacı˘ grafikleridir.

1. GIRIS¸

Soy Agacı˘ programı, kullanıcıların aile uyeleri¨ hakkında grafiksel agac˘ ¸ bic¸iminde bilgi olus¸turma ve gor¨ unt¨ uleme¨ imkanı sunan bir arac¸tır. Program C# diliyle yazılmıs¸tır ve Microsoft Visual Basic ve Microsoft Office Interop Excel kut¨ uphanelerini¨ kullanarak Microsoft Excel hesap tablosundan veri okur ve soy agacını˘ olus¸turur.

Program kullanımı kolay ve kullanıcı dostu tasarlanmıs¸tır. Basit ve kolayca kullanılabilen bir arayuze¨ sahiptir ve kullanıcıların aile uyeleri¨ hakkında hızlı ve kolayca bilgi gor¨ unt¨ ulemesine¨ imkan verir. Program ayrıca, kullanıcının farklı soy agacı˘ panelleri arasında gec¸is¸ yapmasına olanak saglayan˘ bir node kumesi¨ de ic¸erir.

2. YO¨ NTEM

Program birkac¸ liste ve degis˘ ¸ken bildirerek bas¸lar. Roots listesi soy agac˘ ¸larının root du¨g˘umlerini,¨ satırlar listesi aile uyeleri¨ arasındaki baglantıları˘ ve du¨gmeler˘ listesi GUI’de kullanılan du¨gmeleri˘ sak- layacaktır. tablePath degis˘ ¸keni, programın verileri okuyacagı˘ Excel tablosunun yolunu saklar.

YUNUS ERDEM AKPINAR

Kocaeli Univer¨ sitesi Muhendislik¨ Fakultesi¨ Bilgisayar Muhendisli¨ gi˘ Bol¨ um¨ u¨ 2. sınıf akpinaryunuserdem@gmail.com        210202012

Program GUI’nin ana formunu temsil eden bir Form1 sınıfına sahiptir. Form1 sınıfı, form yuklendi¨ ginde˘ c¸alıs¸tırılan bir Form1 Load yontemine¨ sahiptir. Bu fonksiyon diger˘ iki yontemi¨ c¸agırır:˘ CreateFamilyTrees ve CreateButtonsList.

CreateFamilyTrees yontemi¨ Excel tablosundaki verileri okur ve PersonUI sınıfını kullanarak aile agac˘ ¸larının grafiksel gosterimlerini¨ olus¸turur. Fonksiyon, ExcelApp.Application sınıfının bir orne¨ gini˘ olus¸turarak ve tablePath degis˘ ¸keni tarafından belirtilen Excel tablosunu ac¸mak ic¸in bunu kullanarak bas¸lar. Fonksiyon daha sonra tablodaki sayfalar uzerinde¨ yineleme yaparak her sayfa ic¸in yeni bir soy agacı˘ olus¸turur.

Fonksiyon, her sayfa ic¸in sayfadaki hucrelerden¨ verileri okur ve bunları soy agacındaki˘ birey- lerin grafiksel temsilleri olan PersonUI nesnelerini olus¸turmak ic¸in kullanır. Fonksiyon daha sonra bu nesneleri GUI’deki uygun panele ekler. Fonksiyon ayrıca c¸izgi nesneleri olus¸turarak ve bunları lines listesine ekleyerek aile uyeleri¨ arasında baglantılar˘ olus¸turur.

CreateButtonsList fonksiyonu, GUI’de kullanılan butonların bir listesini olus¸turur. butonlar, Button sınıfı kullanılarak olus¸turulur ve buttons listesine eklenir.

Son olarak, program, formun ve biles¸enlerinin bas¸latılmasından sorumlu olan bir Form1 yapıcıya sahiptir. InitializeComponent fonksiyonu, formu ve biles¸enlerini ayarlamak ic¸in c¸agrılır˘ .

Kodun nasıl c¸alıs¸tıgını˘ anlamak ic¸in, programın verileri okudugu˘ Excel tablosunun yapısını anlamak onemlidir¨ . Tablonun, her biri farklı bir soy agacını˘

temsil eden dort¨ sayfaya sahip oldugu˘ varsayılır. Her sayfa, soy agacındaki˘ bireyler hakkında satırlar ve sutunlar¨ halinde duzenlenmis¨ ¸ veriler ic¸erir. Sutunlar¨ her bir birey hakkında isim, dogum˘ tarihi ve kan grubu gibi farklı bilgi parc¸alarını temsil etmektedir. Satırlar ise soy agacındaki˘ farklı bireyleri temsil eder.

CreateFamilyTrees fonksiyonu, her sayfanın satırları ve sutunları¨ uzerinde¨ yineleme yaparak Ex- cel tablosundaki verileri okur. Fonksiyon, her satır ic¸in hucrelerdeki¨ verileri okur ve bunları bir Per- sonUI nesnesi olus¸turmak ic¸in kullanır. Fonksiyon daha sonra PersonUI nesnesini GUI’deki uygun pan- ele ekler. Yontem¨ ayrıca, c¸izgi nesneleri olus¸turup bunları lines listesine ekleyerek aile uyeleri¨ arasında baglantılar˘ olus¸turur.

PersonUI sınıfı, soy agacındaki˘ bir bireyin grafik- sel temsilini temsil eden ozel¨ bir sınıftır. Sınıf, bireyin adını, soyadını ve dogum˘ tarihini temsil eden name, surname ve dateOfBirth gibi c¸es¸itli ozelliklere¨ sahiptir. Sınıf ayrıca PersonUI nes- nesinin GUI’de c¸izilmesinden sorumlu olan bir Draw fonksiyonuna sahiptir.

CreateButtonsList fonksiyonu GUI’de kullanılan butonların bir listesini olus¸turur. Bu Butonlar Button sınıfı kullanılarak olus¸turulur ve du¨gmeler˘ listesine eklenir.

\1) Person Classı:

Bu sınıf, C# programlama dilinde FamilyTree ad alanında bir Person sınıfını temsil eder. Bu sınıf, soy agacında˘ bir kis¸iyi temsil eder ve kis¸inin kimligi,˘ ilis¸kileri ve kis¸isel ayrıntıları hakkında bilgi ic¸erir.

Person sınıfı, kis¸i hakkında bilgi saklayan birc¸ok ozel¨ alan ic¸erir, bunlar arasında kis¸inin adı, soyadı, dogum˘ tarihi, annesinin ve babasının adları, kan grubu, meslegi,˘ evlendikten onceki¨ soyadı ve cin- siyeti yer alır. Sınıf ayrıca, bu bilgilere eris¸imi saglayan˘ birc¸ok yayın ozelli¨ gi˘ de ic¸erir.

Person sınıfı, birc¸ok arguman¨ alan ve bu argumanlara¨ gore¨ alanların degerlerini˘ ayarlayan bir yapıcı metodu vardır. Ayrıca, bir kis¸inin bilgi- lerinin guncellenmesine¨ izin veren UpdateInfo adlı bir metodu da vardır.

Person sınıfı, bir kis¸inin bas¸ka bir kis¸iyle evlen- mesine izin veren AddSpouse adlı bir metodu vardır. Eger˘ kis¸i zaten evliyse, metod once¨ mevcut es¸in yeni es¸le aynı olup olmadıgını˘ kontrol eder. Aynı degillerse,˘ metod mevcut es¸i yeni es¸le degis˘ ¸tirir.

Person sınıfı, bir kis¸inin c¸ocukları olmasına izin veren AddChild adlı bir metodu da vardır. Bu metod, Person nesnesini bir arguman¨ olarak alır ve mevcut kis¸inin c¸ocukları listesine ekler.

Person sınıfı ayrıca, aile agacında˘ bir kis¸i ara- mayı, aile agacından˘ bir kis¸i c¸ıkarmayı ve kul- lanıcı arayuz¨ unde¨ gor¨ unt¨ ulenen¨ bilgileri guncelleme¨ gibi c¸es¸itli amac¸lar ic¸in Search, SearchInFami- lyTree, Remove ve RemoveFromFamilyTree gibi diger˘ metodlar da ic¸erir. Bu metodlar, aile agacında˘ bir kis¸i aramayı, aile agacından˘ bir kis¸i c¸ıkarmayı ve diger˘ c¸es¸itli gore¨ vleri yapar.

3. SONUC¸

Soy Agacı˘ programı, kullanıcıların aile uyeleri¨ hakkında grafiksel agac˘ ¸ bic¸iminde bilgi olus¸turma ve gor¨ unt¨ uleme¨ imkanı sunan bir arac¸tır. Program C# diliyle yazılmıs¸tır ve Microsoft Visual Basic ve Microsoft Office Interop Excel kut¨ uphanelerini¨ kullanarak Microsoft Excel hesap tablosundan veri okur ve soy agacını˘ olus¸turur.

4. KAYNAKC¸A
1. Juan, Angel (2006). ”Ch20 –Data Structures; ID06 - PRO- GRAMMING with JAVA (slide part of the book ’Big Java’, by CayS. Horstmann)” (PDF). p. 3. Archived from the original (PDF) on 2012-01-06. Retrieved 2011-07-10.
1. Black, Paul E. (2004-08-16). Pieterse, Vreda; Black, Paul E. (eds.). ”linked list”. Dictionary of Algorithms and Data Struc- tures. National Institute of Standards and Technology. Retrieved 2004-12-14.
1. Antonakos, James L.; Mansfield, Kenneth C. Jr. (1999). Practi- cal Data Structures Using C/C++. Prentice-Hall. pp. 165–190. ISBN 0-13-280843-9.
1. https://medium.com/kodcular/adan-z-ye-c-oop-2d766cf2d144
1. https://www.c-sharpcorner.com/UploadFile/84c85b/object- oriented-programming-using-C-Sharp-net/
1. Collins, William J. (2005) [2002]. Data Structures and the Java Collections Framework. New York: McGraw Hill. pp. 239–303. ISBN 0-07-282379-8.
1. https://learn.microsoft.com/en-us/dotnet/csharp/programming- guide/classes-and-structs/access-modifiers
1. https://www.pluralsight.com/courses/c-sharp-code-more-object -oriented?aid=7010a000002BWqGAAW&promo=&utm source non¯ branded&utm medium=digital paid search google&utm campaign=EMEA![](Aspose.Words.bf976d4e-5d9c-434a-a8a7-4587f2d1c872.001.png)![](Aspose.Words.bf976d4e-5d9c-434a-a8a7-4587f2d1c872.002.png) Dynamic&utm content=&gclid=Cj0KCQiA veebBhD ARIsAFaAvrHyp395fdtFdwBUKknQuSz0xxfOI0ILd xihO33PeQN3-n5pHLZzNRcaArXnEALw wcB
1. Green, Bert F. Jr. (1961). ”Computer Languages for Symbol Manipulation”. IRE Transactions on Human Factors in Elec- tronics (2): 3–8. doi:10.1109/THFE2.1961.4503292.
1. McCarthy, John (1960). ”Recursive Functions of Symbolic Ex- pressions and Their Computation by Machine, Part I”. Commu- nications of the ACM. 3 (4): 184. doi:10.1145/367177.367199. S2CID 1489409.
1. Parlante, Nick (2001). ”Linked list basics” (PDF). Stanford University. Retrieved 2009-09-21
1. Shanmugasundaram, Kulesh (2005-04-04). ”Linux Kernel Linked List Explained”. Retrieved 2009-09-21.
1. https://circuitstream.com/blog/learn-c-for-unity-lesson-6- inheritance-and-interfaces/
1. https://www.youtube.com/watch?v=2LA3BLqOw9g
1. https://en.wikipedia.org/wiki/List of terms relating to algorit hms and data structures
1. https://www.gencayyildiz.com/blog/cta- inheritancekalitimmiras/
1. Microsoft documentation on dynamic interfaces: https://docs.microsoft.com/en-us/dotnet/csharp/programming- guide/interfaces/dynamic-interfaces
1. ”C# 8.0 and .NET Core 3.0 - Modern Cross-Platform De- velopment - Fourth Edition” by Mark J. Price: This book includes a chapter on dynamic interfaces that provides a detailed overview of the topic, including examples of how to use dynamic interfaces in C#.
1. C# Corner tutorial on dynamic interfaces: https://www.c- sharpcorner.com/article/dynamic-interface-in-c-sharp/
1. C# Station tutorial on dynamic interfaces: https://www.csharp- station.com/Tutorial/CSharp/Lesson24
1. https://www.javatpoint.com/c-sharp-abstract
1. The Microsoft documentation on the Label control in C#: https://docs.microsoft.com/en- us/dotnet/api/system.windows.forms.label?view=netframework-

4.8

23. A tutorial on creating dynamic labels in C#: https://www.c- sharpcorner.com/article/creating-dynamic-labels-in-c-sharp/
23. A forum discussion on dynamically updating the text of a label in C#: https://www.dreamincode.net/forums/topic/246873- dynamically-update-label-text-c%23/

![](Aspose.Words.bf976d4e-5d9c-434a-a8a7-4587f2d1c872.003.png)

Fig. 2. psuedo-2![](Aspose.Words.bf976d4e-5d9c-434a-a8a7-4587f2d1c872.004.png)

Fig. 1. psuedo-1![](Aspose.Words.bf976d4e-5d9c-434a-a8a7-4587f2d1c872.005.png)

Fig. 3. psuedo-3
