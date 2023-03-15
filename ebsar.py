import random
import pandas as pd
import xlwings as xw

#keywords = pd.read_excel('C:/Users/bassa/OneDrive/سطح المكتب/KeyWords.xltx') #فتح ملفات البيانات
#تعريف كل صفحه باسم متغير
Questions = xw.Book('C:/Users/bassa/OneDrive/سطح المكتب/KeyWords.xltx').sheets('Questions')
Q16_25 = xw.Book('C:/Users/bassa/OneDrive/سطح المكتب/KeyWords.xltx').sheets('Q16-25')
Q26_70 = xw.Book('C:/Users/bassa/OneDrive/سطح المكتب/KeyWords.xltx').sheets('Q26-70')
#متغير لتحدد عدد الدورة
rand = 0
#متغير لحساب نسبة الاكتأب
count = 0

# رسائل ترحيب
GREETINGS = ["مرحبا", "السلام عليكم", "اهلا", "السلام"]

# ردود لرسائل الترحيب
GREETING_RESPONSES = ["اهلا!", "مرحبا بك!", "يا مرحبا!", "اهلا,كيف حالك اليوم؟"]
print("مرحبا!")
name = input("كيف تحب نناديك؟\n")
age = int(input("اخبرني بعمرك:\n"))

# رسائل وداعية
FAREWELLS = ["وداعا", "مع السلامه", "الى اللقاء", "اراك مجددا"]

# ردود رسائل الودعية
FAREWELL_RESPONSES = ["الى اللقاء اراك مجددا!", "يوم سعيد لك الى اللقاء!", "اراك قريبا!"]

#تعريف بكيان
print("اهلا", name, "\nمعاك كيان انا AI وجد من أجل مساعدتك.\nهيا لنبدا مع مجموعة من الاسألة")

#بداية الاسألة التشخيصيه
smoker = input("هل انت مدخن؟\n"
               "1- نعم\n"
               "2- لا  \n")
if smoker == 'نعم':
    count =+ 10
# اسأله عشوائة
while True:
    print(random.choice(Questions.range("A1:A19").value))
    rand = rand + 1
    # اخذ رد العميل
    user_input = input("أنت: ")

    # انهاء المحادثة
    if user_input in FAREWELLS:
        print(random.choice(FAREWELL_RESPONSES))
        break

    # ارسال الرسالة ترحيبية بشكل عشوائي
    if user_input in GREETINGS:
        print(random.choice(GREETING_RESPONSES))

    if user_input == 'نعم':
        count = count + 10

    if rand == 3:
        break

while True:
    rand = rand + 1
    # اسالة محدده لاعمار 16 الى 25
    if age >= 16 or age <= 25:
        print(random.choice(Q16_25.range("A1:A16").value))
        user_input = input("أنت: ")
        if user_input == 'نعم':
            count = count + 10
    # اسلة محدد لاعمار 26 واكثر
    if age >= 26 or age <= 90:
        print(random.choice(Q26_70.range("A1:A13").value))
        user_input = input("أنت: ")
        if user_input == 'نعم':
            count = count + 10

    # انهاء المحادثة
    elif user_input in FAREWELLS:
        print(random.choice(FAREWELL_RESPONSES))
        break
    #الخروج بعد سؤال 6 أسالة
    if rand == 6:
        break

print("عبر عن ما تشعر به حاليا:")

contains_all = 0
key_count = 0
# مقارنه الكلمات المدخلة من المستخدم مع الكلمات المفتاحيه بملف الاكسل
while True:
    text_input = input("أنت: ").split()
    print(text_input)

    Keys = xw.Book('C:/Users/bassa/OneDrive/سطح المكتب/KeyWords.xltx').sheets["Keys"]
    print(Keys.range("A1:A69").value)

    for user_input in text_input:
        for keyword in Keys.range("A1:A69").value:
            if keyword == user_input:

                key_count += 10
    break
# طباعة نسبة الاكتأب
print("انت تعاني من بالاكتأب بنسبة", ((count + key_count) / 200) * 100 , "%")
print("اعمل حاليا على تجهيز خطتك العلاجية.\nيمكنك التحدث معي في أي وقت انا هنا من أجلك.")
