# वोटर सुविधा (Voter Suvidha) - Voters Slip Generator 🗳️

[![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows&logoColor=white)](#)
[![Python](https://img.shields.io/badge/Python-3.x-3776AB?logo=python&logoColor=white)](#)
[![Offline](https://img.shields.io/badge/Internet-100%25%20Offline-success)](#)

वार्ड-वार मतदाता सूची विश्लेषण एवं प्रिंट-रेडी मतदाता पर्ची (Voter Suvidha Slip) निर्माण प्रणाली।

यह सॉफ्टवेयर **100% ऑफलाइन (बिना इंटरनेट)** किसी भी विंडोज कंप्यूटर (Windows 10 / 11 / 7) पर बिना किसी इंस्टॉलेशन के पोर्टेबल रूप से काम करता है।

---

## 🌟 मुख्य विशेषताएं (Key Features)

- **📄 ऑटोमैटिक PDF एक्सट्रैक्शन**:
  - निर्वाचन नामावली (Electoral Roll PDF) से वार्ड, भाग और मतदान केंद्र का नाम व पता स्वतः निष्कर्षण।
  - कृती देव (Krutidev) फॉन्ट का शुद्ध देवनागरी (Unicode) में रूपांतरण।
  - विलोपन सूची (Deleted / Shifted Voters) की स्वतः छंटनी।

- **🎨 कस्टमाइज़्ड वोटर स्लिप (Voter Suvidha Slip)**:
  - **प्रत्याशी की फोटो (Candidate Photo)** और **पार्टी चुनाव चिन्ह (Party Symbol)** जोड़ने का विकल्प।
  - दाईं ओर (~36% क्षेत्र) सुंदर फ्रेम बॉक्स में प्रत्याशी की फोटो, नाम, पार्टी चुनाव चिन्ह, पार्टी नाम व अपील संदेश।
  - बाईं ओर (~63% क्षेत्र) में शीर्षक **"वोटर सुविधा स्लिप"**, वार्ड नं के ठीक नीचे स्पष्ट **क्रम संख्या (Serial No)** बैज, मतदाता का नाम, पिता/पति का नाम, मकान नं, उम्र, लिंग, EPIC एवं मतदान केंद्र।
  - न्यूनतम खाली जगह (Zero Waste Space) के साथ सघन और स्पष्ट लेआउट।

- **🖨️ लचीला A4 प्रिंटिंग ग्रिड**:
  - 4 स्लिप्स प्रति पेज (2 × 2)
  - 6 स्लिप्स प्रति पेज (2 × 3)
  - 8 स्लिप्स प्रति पेज (2 × 4) — मानक व अनुशंसित
  - 10 स्लिप्स प्रति पेज (2 × 5)
  - 12 स्लिप्स प्रति पेज (2 × 6)

- **⚡ तेज़ एवं स्थिर सर्वर**:
  - मल्टी-थ्रेडेड HTTP सर्वर (`ThreadedHTTPServer`)।
  - HTTP Range Requests (`206 Partial Content`) सपोर्ट, जिससे Chrome और Edge में 200+ पेजों की PDF बिना हैंग हुए तुरंत लोड व प्रिंट होती है।
  - 11-कॉलम मास्टर वोटर लिस्ट Excel (`.xlsx`) एक्सपोर्ट।

---

## 🚀 शुरुआत कैसे करें (Quick Start)

### 1. पोर्टेबल ऐप चलाएं (Windows Portable)
1. **`VoterSuvidha.exe`** या **`Voter_Suvidha.bat`** पर डबल-क्लिक करें।
2. आपका डिफॉल्ट ब्राउज़र स्वतः खुल जाएगा:
   ```
   http://127.0.0.1:5000
   ```
3. वार्ड की PDF अपलोड करें या मौजूदा मास्टर डेटा से सीधे स्लिप बनाएं।

### 2. पायथन के माध्यम से चलाएं (Developers)
```powershell
python voter_suvidha/server.py
start http://127.0.0.1:5000
```

---

## 📁 प्रोजेक्ट संरचना (Repository Structure)

```
Voters-Slip/
├── VoterSuvidha.exe             # 1-क्लिक पोर्टेबल स्टार्टर
├── Voter_Suvidha.bat            # वैकल्पिक बैच स्टार्टर
├── Voter_Suvidha_Portable.zip   # रेडी-टू-शेयर पोर्टेबल पैकेज
├── README.md                    # प्रोजेक्ट जानकारी
├── README_INSTRUCTIONS.txt      # हिंदी में विस्तृत निर्देश
├── beawar_ward_001_part_001.xlsx# 1,200 मतदाताओं का सत्यापित मास्टर डेटा
├── sample_voter_slips_page1.png # जनरेटेड वोटर स्लिप का नमूना
├── voter_slips_ward_001.pdf     # 200 पेजों की प्रिंट-रेडी वोटर स्लिप्स PDF
├── voter_list_ward_001.xlsx     # मास्टर वोटर लिस्ट एक्सेल एक्सपोर्ट
│
└── voter_suvidha/
    ├── server.py                # मल्टी-थ्रेडेड HTTP बैकएंड सर्वर
    ├── server.ps1               # PowerShell आधारित बैकएंड सर्वर
    ├── extractor.ps1            # PDF से वोटर डेटा एक्सट्रैक्टर
    ├── pdf_builder.ps1          # हाई-स्पीड HTML से PDF कनवर्टर
    ├── excel_builder.ps1        # Excel जनरेटर
    ├── template.xlsx            # 11-कॉलम एक्सेल टेम्पलेट
    ├── words_dict.json          # देवनागरी वर्तनी सुधार शब्दकोश
    │
    └── web/                     # मॉडर्न वेब इंटरफ़ेस
        ├── index.html           # मुख्य डैशबोर्ड व कंट्रोल पैनल
        ├── app.js               # इंटरएक्टिव लॉजिक, प्रीव्यू व अपलोड
        ├── style.css            # आधुनिक रिस्पॉन्सिव स्टाइलिंग
        └── vote_list.html       # लाइव प्रिंट प्रीव्यू टेम्पलेट
```

---

## 💻 सिस्टम आवश्यकताएं (Requirements)

- **OS**: Windows 7, 8, 10, 11 (32-bit या 64-bit)
- **Browser**: Google Chrome या Microsoft Edge
- **Internet**: शून्य (0 MB - पूर्णतः ऑफलाइन)
