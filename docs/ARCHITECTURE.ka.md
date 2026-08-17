# ContextPack Desktop — არქიტექტურის რუკა

ეს დოკუმენტი აღწერს პროგრამის მიმდინარე მუშაობას მარტივი ენით. რეფაქტორის დროს მისი დანიშნულებაა დაგვეხმაროს, რომ კოდის სტრუქტურა შევცვალოთ, მაგრამ პროგრამის ქცევა შემთხვევით არ დავაზიანოთ.

შემდეგი სამუშაო სესიის ზუსტი გეგმა იხილე `docs/REFACTOR_HANDOFF.ka.md`-ში, ხოლო დამწყები მფლობელის უსაფრთხო სამუშაო წესები — `docs/BEGINNER_SAFETY.ka.md`-ში.

## პროგრამის მთავარი გზა

1. მომხმარებელი ფაილსა და დამუშავების რეჟიმს ირჩევს Desktop GUI-ში.
2. GUI ცალკე PowerShell runner-ს გადასცემს უსაფრთხოდ აგებულ არგუმენტებს.
3. runner სტრუქტურირებულ JSON შეტყობინებებს აბრუნებს, რომ GUI-მ აჩვენოს პროგრესი, შეცდომა ან დასრულებული შედეგი.
4. მთავარი dispatcher ფაილის გაფართოებით ირჩევს PDF, Excel, image OCR ან ზოგადი Markdown conversion-ის გზას.
5. დამუშავების შედეგი ჯერ დროებით build საქაღალდეში იქმნება.
6. მხოლოდ წარმატებული დასრულების შემდეგ ცვლის build საბოლოო output პაკეტს.

## შესასვლელი წერტილები

- `Start-ContextPack-GUI.cmd` — მომხმარებლისთვის Desktop launcher.
- `contextpack_gui.py` — GUI-ის მიმდინარე Python შესასვლელი წერტილი.
- `contextpack-gui-runner.ps1` — GUI-სა და დამუშავების სკრიპტებს შორის საზღვარი.
- `contextpack.ps1` — command-line dispatcher, რომელიც processor-ს ირჩევს.
- `contextpack-routing.ps1` — input-ის ტიპის ამოცნობა და PDF/Excel/Image/Document processing route-ები.
- `common.ps1` — ძველი processing scripts-ის სტაბილური compatibility loader.
- `ContextPack.Environment.ps1` — Python/Tesseract-ის პოვნა და OCR გარემოს მომზადება.
- `ContextPack.Build.ps1` — atomic build, უსაფრთხო replacement/rollback და cleanup საზღვრები.
- `ContextPack.Manifest.ps1` — source hash-ის, outputs-ის, settings-ისა და warnings-ის manifest-ში ჩაწერა.
- `ContextPack.ExcelCom.ps1` — Excel COM retry, უსაფრთხოების პარამეტრები, workbook-ის read-only გახსნა და COM lifecycle cleanup.

## Python-ის გამოყოფილი საფუძველი

- `contextpack/job_options.py` — ერთი დავალების პარამეტრების უცვლელი მონაცემთა მოდელი.
- `contextpack/file_types.py` — input-ისთვის შესაბამისი processing route-ის არჩევა.
- `contextpack/validation.py` — GUI-სგან და ენისგან დამოუკიდებელი validation წესები.
- `contextpack/paths.py` — პროექტის, runner-ისა და ნაგულისხმევი output-ის სტაბილური მისამართები.
- `contextpack/runner_command.py` — PowerShell runner-ისთვის shell parsing-ის გარეშე უსაფრთხო არგუმენტების აგება.
- `contextpack/runner_events.py` — runner-ის სტრუქტურირებული JSON event-ების ამოცნობა.
- `contextpack/environment.py` — engine, OCR და Excel შესაძლებლობების ხელმისაწვდომობის შემოწმება.
- `contextpack/localization.py` — ქართული/ინგლისური ტექსტები, რეჟიმების წარწერები და runner-ის სტატუსების თარგმნა.
- `contextpack/gui/theme.py` — Desktop ფანჯრის ფერები და ttk ვიზუალური სტილები.
- `contextpack/gui/layout.py` — ფანჯრის ხუთი ვიზუალური სექცია: header, file selection, options, action/progress და details/results.
- `contextpack/job_controller.py` — hidden PowerShell subprocess-ის გაშვება, output-reading thread, GUI event queue და cooperative Cancel token.
- `contextpack/excel/cells.py` — მხოლოდ რეალურად შენახული/შევსებული Excel უჯრედების უსაფრთხო წაკითხვა.
- `contextpack/excel/markdown.py` — rectangular და sparse Markdown table-ების შექმნა და მნიშვნელობების escaping.
- `contextpack/excel/workbook.py` — sheet folder-ის უსაფრთხო სახელი და workbook calculation mode.
- `contextpack/excel/analysis.py` — worksheet bounds, formulas/cached errors, sparse mode და layout-risk metrics ერთ immutable `SheetAnalysis` მოდელში.
- `contextpack/excel/reporting.py` — sheet Markdown ფაილები, workbook indexes, quality report და renderer-ის `excel-metrics.json`.
- `contextpack/excel/extractor.py` — workbook-ების გახსნა, sheet analysis/report orchestration და failure-ის დროს handle-ების გარანტირებული დახურვა.

`layout.py` მხოლოდ widget-ებს ქმნის და application-ის callback-ებს უკავშირებს. ფაილის დამუშავება, subprocess და Cancel ლოგიკა მასში განზრახ არ არის, რათა ვიზუალური განლაგება processing behavior-ს არ შეერიოს.

`job_controller.py` Tkinter widget-ებს არ ეხება. ის ტექნიკურ background execution-ს მართავს, ხოლო `ContextPackGui` მიღებული event-ების მიხედვით მხოლოდ status/progress/dialog მდგომარეობას აახლებს.

`contextpack_gui.py` ამ ფუნქციებს ისევ ძველი სახელებით import-ავს. ეს დროებითი compatibility layer ძველ ტესტებს, launcher-სა და installer-ს მუშაობას უნარჩუნებს.

## დამუშავების გზები

- PDF: საჭიროების დადგენა → სურვილისამებრ OCR → Markdown/text extraction → გვერდების render → quality report და manifest.
- Excel: workbook-ის უსაფრთხოდ გახსნა → values/formulas extraction → layout diagnostics → PDF/PNG render → quality report და manifest.
- Image: ქართული და ინგლისური OCR → ტექსტური შედეგი.
- სხვა დოკუმენტი: MarkItDown-ის საშუალებით Markdown conversion.

Root-level `contextpack.ps1` სტაბილური command-line entry point-ია: ის პარამეტრებს იღებს, input-ს ამოწმებს და routing მოდულს იძახებს. ფორმატების კონკრეტული გადაწყვეტილებები `contextpack-routing.ps1`-შია თავმოყრილი, რათა მათი ცალ-ცალკე წაკითხვა და ტესტირება შეიძლებოდეს.

## უსაფრთხოების მნიშვნელოვანი წესები

- წყარო read-only რეჟიმში მუშავდება, სადაც გამოყენებული ბიბლიოთეკა ან Excel ამის საშუალებას იძლევა.
- Excel macros და events არ უნდა გაეშვას.
- ნაწილობრივ შექმნილმა output-მა წინა კარგი პაკეტი არ უნდა ჩაანაცვლოს.
- დროებითი build მხოლოდ მისთვის განკუთვნილი output root-ის შიგნით იშლება.
- GUI cancellation უსაფრთხო ოპერაციის საზღვარზე სრულდება.

`ContextPack.Build.ps1`-ის ცვლილებები მოწმდება ცალკე დროებით საქაღალდეში: ტესტი ადასტურებს source hash-ს, ერთსახელიანი სხვადასხვა წყაროების გაყოფას, იგივე source-ის package replacement-ს და unrelated output ფაილების შენარჩუნებას.

## რეფაქტორის მიმართულება

რეფაქტორისას GUI, validation, translations, process orchestration და format-specific processing ცალკე პასუხისმგებლობებად გაიყოფა. ძველი launcher და საჯარო entry point-ები თავსებადობისთვის შენარჩუნდება, სანამ installer ახალ სტრუქტურაზე სრულად არ შემოწმდება.
