# ContextPack Desktop

ContextPack Desktop გარდაქმნის PDF-ს, სკანს, Excel-სა და Office დოკუმენტებს AI-სთვის მოსახერხებელ პაკეტებად. Markdown გამოიყენება სწრაფი წაკითხვისთვის, ხოლო ორიგინალი, ფორმულები, სტრუქტურის მონაცემები და გვერდების სურათები — ზუსტი გადამოწმებისთვის.

> Markdown — კითხვა. ორიგინალი — სიზუსტე. სურათები — ვიზუალური მტკიცებულება.

[English documentation](README.md)

## შესაძლებლობები

- PDF, DOCX, PPTX, XLSX, HTML, CSV, JSON და სხვა მხარდაჭერილი ფორმატების Markdown-ად გარდაქმნა;
- ქართული და ინგლისური სკანების OCR;
- PDF-ის სრული პაკეტი: Markdown, საწყისი PDF, გვერდების PNG-ები და ორენოვანი AI ინსტრუქციები;
- Excel-ის სრული პაკეტი: მნიშვნელობები, ფორმულები, cached შედეგები, workbook-ის სტრუქტურა, ორიგინალი და ვიზუალური გვერდები;
- ყველა სამუშაო ფაილის ლოკალურად დამუშავება;
- ნებისმიერ Windows კომპიუტერზე მუშაობა პირადი ან მყარად ჩაწერილი გზების გარეშე.

## მოთხოვნები

- Windows 10 ან 11;
- PowerShell 5.1 ან ახალი;
- Python 3.10 ან ახალი;
- ინტერნეტი პირველი ინსტალაციისას;
- Microsoft Excel მხოლოდ Excel-ის PDF/PNG რენდერისთვის.

OCR-ს სჭირდება Tesseract 5. `setup.ps1` იპოვის უკვე დაყენებულ ვერსიას ან `winget`-ით დააყენებს Windows-ის UB Mannheim პაკეტს.

## ახალ კომპიუტერზე ინსტალაცია

```powershell
git clone https://github.com/akaki21/ContextPack-Desktop.git
cd ContextPack-Desktop
Set-ExecutionPolicy -Scope Process Bypass
.\setup.ps1
.\check-environment.ps1
```

Setup:

1. შექმნის `.venv` გარემოს;
2. დააყენებს Python პაკეტებს;
3. იპოვის ან დააყენებს Tesseract-ს;
4. მოამზადებს ლოკალურ `tessdata`-ს;
5. ჩამოტვირთავს ინგლისურ, ქართულ და orientation OCR მოდელებს;
6. შექმნის `input` და `output` საქაღალდეებს.

`.venv`-ის გააქტიურება აუცილებელი არაა — სკრიპტები საკუთარ გარემოს პირდაპირ იყენებს.

## სწრაფი გამოყენება

ჩვეულებრივი ფაილი:

```powershell
.\convert-to-markdown.ps1 ".\input\document.docx"
```

დასკანერებული PDF:

```powershell
.\ocr-and-convert.ps1 ".\input\scan.pdf"
```

ერთი სურათი:

```powershell
.\ocr-image.ps1 ".\input\page.png"
```

PDF-ის სრული პაკეტი:

```powershell
.\pdf-package.ps1 ".\input\document.pdf"
```

OCR-იანი სრული პაკეტი:

```powershell
.\pdf-package.ps1 ".\input\scan.pdf" -Ocr
```

ტექნიკური ნახაზისთვის:

```powershell
.\pdf-package.ps1 ".\input\drawing.pdf" -Dpi 240
```

Excel-ის სრული პაკეტი:

```powershell
.\excel-package.ps1 ".\input\workbook.xlsx"
```

Excel-ის პაკეტში შედის `workbook-info.md`, `formulas.md`, `values.md`, ორიგინალი workbook, ვიზუალური გვერდები და AI-სთვის ორენოვანი ინსტრუქციები. cached ფორმულის შედეგები რომ აქტუალური იყოს, workbook წინასწარ Excel-ში გადაათვლევინე და შეინახე.

## AI-სთვის ეფექტური მიწოდება

1. ჯერ მიაწოდე Markdown და ზუსტად აღწერე მიზანი.
2. მიუთითე საჭირო გვერდები, ფურცლები, პერიოდი, სვეტები ან უჯრედების დიაპაზონი.
3. ორიგინალი დაურთე ზუსტი შემოწმების ან ცვლილებისას.
4. PNG გვერდები დაურთე მხოლოდ მაშინ, როცა ვიზუალური განლაგება, ნახაზი, ხელმოწერა, ბეჭედი ან რთული ცხრილია მნიშვნელოვანი.

სრული პაკეტი ავტომატურად ქმნის `AI-HANDOFF.md` და `AI-HANDOFF.ka.md` ფაილებს. დამატებითი მაგალითებია [AI-HANDOFF.ka.md](AI-HANDOFF.ka.md)-ში.

## კონფიდენციალურობა

ContextPack ფაილებს ლოკალურად ამუშავებს და თვითონ არაფერს ტვირთავს AI სერვისში. შემდგომ შენ ირჩევ, რომელი შედეგი ან ორიგინალი გააზიარო.

## შემოწმება

```powershell
.\convert-to-markdown.ps1 ".\examples\sample.csv"
```

შედეგი უნდა შეიქმნას `output\sample.md`-ში.

## პრობლემები

- PowerShell ბლოკავს სკრიპტს: `Set-ExecutionPolicy -Scope Process Bypass`.
- Tesseract სხვა ადგილასაა: მიუთითე `$env:CONTEXTPACK_TESSERACT`.
- OCR-ში შეცდომებია: კრიტიკული სახელები და რიცხვები გადაამოწმე ორიგინალთან.
- Excel-ის რენდერი არ მუშაობს: გადაამოწმე, რომ Microsoft Excel დაყენებული და გააქტიურებულია.
- `ffmpeg` გაფრთხილება PDF/Excel/Word დამუშავებას ხელს არ უშლის.

## ლიცენზია

[MIT](LICENSE)
