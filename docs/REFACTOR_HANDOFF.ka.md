# ContextPack Desktop — რეფაქტორის handoff

ეს ფაილი განკუთვნილია Codex-ის შემდეგი სესიისთვის. სამუშაოს დაწყებამდე სრულად წაიკითხე ეს დოკუმენტი, `docs/ARCHITECTURE.ka.md` და `docs/BEGINNER_SAFETY.ka.md`.

## პროექტის მფლობელთან მუშაობის კონტექსტი

მფლობელი დამწყები პროგრამისტია, სწავლობს მუშაობის პროცესში და ძირითადად Python-თან ჰქონია შეხება. PowerShell, Excel COM, installer და release engineering მისთვის ჯერ უცნობი ან რთული სფეროებია.

ამიტომ შემდეგი სესია ვალდებულია:

- ყოველი milestone წინასწარ ახსნას მარტივი ქართულით;
- მიუთითოს რისკი, შეცვლილი ფაილები და rollback გზა;
- ცვლილების შემდეგ ახსნას diff-ის არსი და ტესტების მტკიცებულება;
- არ მოითხოვოს მომხმარებლისგან უცნობი ტექნიკური გადაწყვეტილების ბრმად დამტკიცება;
- მაღალი რისკის ცვლილებაზე შესთავაზოს დამოუკიდებელი review;
- თანდათან ასწავლოს პროექტის არქიტექტურა და Git/test workflow.

სრული წესები და checklist ინახება `docs/BEGINNER_SAFETY.ka.md`-ში და რეფაქტორისას სავალდებულოდ უნდა დაიცვა.

## მომხმარებლის საბოლოო მიზანი

ContextPack Desktop უნდა იყოს ლოკალური, მაქსიმალურად ავტომატური AI document-preparation pipeline. მომხმარებელი აძლევს ჩვეულებრივ ფაილს, პროგრამა კი თვითონ:

1. ამოიცნობს ფაილის ტიპსა და სტრუქტურას;
2. გადაწყვეტს, რომელი extraction/OCR/rendering არის საჭირო;
3. უცვლელად ინახავს ორიგინალს;
4. ქმნის AI-სთვის მსუბუქ Markdown/text მასალას;
5. ქმნის მხოლოდ საჭირო ვიზუალურ მტკიცებულებებს;
6. აფიქსირებს formulas, pages, sheets, warnings და quality risks;
7. ქმნის manifest-ს, quality report-სა და მკაფიო reading order-ს;
8. საბოლოოდ იძლევა პაკეტს, რომელსაც Codex პირდაპირ და სწრაფად დაამუშავებს.

პროგრამა საბოლოოდ მხოლოდ PDF/Excel-ზე არ უნდა იყოს შეზღუდული, თუმცა ახალი ფორმატების დამატებამდე არსებული pipeline-ები უნდა გახდეს სუფთა, სტაბილური და ტესტირებადი.

## მიმდინარე Git მდგომარეობა

- repository: `https://github.com/akaki21/ContextPack-Desktop.git`
- სამუშაო ბრენჩი: `refactor/readable-foundation`
- ამ handoff-ის შექმნამდე ბოლო commit: `490f4cf — Move Excel extraction orchestration into package`
- `main` ჯერ არ შეცვლილა;
- ყველა დასრულებული milestone ატვირთულია სამუშაო ბრენჩზე;
- ახალი სამუშაო ყოველთვის მცირე commit-ებად დაყავი და წარმატებული სრული შემოწმების შემდეგ push გააკეთე.

სესიის დაწყებისას აუცილებლად გაუშვი:

```powershell
git status --short --branch
git fetch --all --prune
git status --short --branch
```

არ გადახვიდე `main`-ზე და არ გახსნა/merge გააკეთო PR მომხმარებლის მკაფიო მითითების გარეშე.

## უკვე დასრულებული რეფაქტორი

### Python GUI foundation

- `contextpack_gui.py` შემცირდა და დარჩა თავსებადი entry point;
- `contextpack/localization.py` — ქართული/ინგლისური ტექსტები;
- `contextpack/gui/layout.py` — GUI-ის ვიზუალური სექციები;
- `contextpack/gui/theme.py` — ფერები და ttk styles;
- `contextpack/job_controller.py` — subprocess/thread/cancel token;
- `contextpack/job_options.py` — job data model;
- `contextpack/validation.py` — GUI-სგან დამოუკიდებელი validation;
- `contextpack/file_types.py` — input classification;
- `contextpack/environment.py` — engine/OCR/Excel readiness;
- `contextpack/runner_command.py` და `runner_events.py` — runner protocol.

ძველი `contextpack_gui.py` imports განზრახ ინარჩუნებს compatibility-ს installer-თან და არსებულ ტესტებთან.

### Processing foundation

- `contextpack.ps1` არის მოკლე, სტაბილური entry point;
- `contextpack-routing.ps1` შეიცავს PDF/Excel/Image/Document route-ებს;
- `common.ps1` არის compatibility loader;
- `ContextPack.Environment.ps1` — Python/Tesseract/OCR environment;
- `ContextPack.Build.ps1` — atomic build/replacement/rollback/cleanup;
- `ContextPack.Manifest.ps1` — manifest writer;
- `ContextPack.ExcelCom.ps1` — COM retry, უსაფრთხო Excel application configuration, read-only workbook open და lifecycle cleanup;
- `ContextPack.ExcelDiagnostics.ps1` — manual page break და shapes diagnostics უსაფრთხო fallback-ებით;
- `ContextPack.ExcelAutoFit.ps1` — AutoFit eligibility condition order და skip reasons Excel COM mutation-ის გარეშე;
- `ContextPack.ExcelWorkbookLayout.ps1` — authoritative workbook-layout PDF/PNG paths, renderer metrics და warnings orchestration;
- portable ZIP build უკვე აკოპირებს `contextpack/` Python package-ს.

### Excel Python extractor

- root `extract-excel-package.py` მხოლოდ 20-ხაზიანი CLI wrapper-ია;
- `contextpack/excel/cells.py` — stored/populated cells;
- `contextpack/excel/markdown.py` — rectangular/sparse Markdown;
- `contextpack/excel/workbook.py` — safe sheet name და calculation mode;
- `contextpack/excel/analysis.py` — immutable `SheetAnalysis`, formulas/errors/layout risks/metrics;
- `contextpack/excel/reporting.py` — Markdown/quality/metrics output;
- `contextpack/excel/extractor.py` — orchestration და workbook handle-ების `try/finally` cleanup.

## მიმდინარე შემოწმებული baseline

მიმდინარე სამუშაო branch-ის ამ მდგომარეობაში წარმატებით გადის:

- 28 Python unit/integration test;
- `tests/test-static.ps1`;
- `tests/test-excel-com.ps1` COM retry/configuration/read-only/cleanup tests;
- `tests/test-excel-diagnostics.ps1` manual break/shapes counting, fallback და release tests;
- `tests/test-excel-autofit.ps1` hidden/empty/wide/drawing/manual-break precedence და eligible-sheet tests;
- `tests/test-excel-workbook-layout.ps1` output paths, renderer arguments, warnings და failure propagation tests;
- `tests/test-common.ps1` atomic build/manifest tests;
- `tests/test-e2e.ps1 -RequireExcel` სრული E2E:
  - Markdown conversion;
  - PDF package და PNG render;
  - Microsoft Excel COM Workbook + AutoFit render;
  - Excel formulas/cached values;
  - image OCR;
  - scanned PDF OCR;
  - searchable OCR PDF;
  - manifest/quality/metrics.

ყველა მნიშვნელოვანი ცვლილების შემდეგ გამოიყენე:

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -p "test_*.py"
.\tests\test-common.ps1
.\tests\test-static.ps1
.\tests\test-e2e.ps1 -RequireExcel
git diff --check
```

E2E script თვითონ ქმნის უნიკალურ Windows Temp საქაღალდეს და უსაფრთხოდ ასუფთავებს. `-KeepArtifacts` არ გამოიყენო, თუ კონკრეტულად არ გჭირდება შედეგის ხელით დათვალიერება.

## შემდეგი ზუსტი ეტაპი

Excel COM lifecycle, layout-risk diagnostics, Workbook orchestration და AutoFit eligibility milestones დასრულებულია. AutoFit-ის skip-condition order და reasons `ContextPack.ExcelAutoFit.ps1`-შია გამოყოფილი, ხოლო რეალური PageSetup mutation ჯერ `excel-package.ps1`-ში რჩება.

შემდეგ მცირე milestone-ში მხოლოდ horizontal pagination და wide-sheet safeguards გამოყავი `excel-package.ps1`-დან. არ შეცვალო და არ შეეხო:

- Workbook vs AutoFit ქცევა;
- horizontal pagination;
- manual page break logic;
- chart/image/merged-cell decisions;
- `print-layout-report.json` schema;
- output package schema.

სასურველი root-level ფაილები უნდა იყოს installer/portable packaging-თან თავსებადი. ახალი nested PowerShell საქაღალდის დამატებამდე გადაამოწმე `installer/build-release.ps1` და `installer/ContextPack.iss`. Root-level `.ps1` ფაილები portable build-ში ავტომატურად ხვდება.

პირველი COM milestone-ის შემდეგ აუცილებლად გაუშვი სრული `-RequireExcel` E2E. მხოლოდ static test საკმარისი არ არის.

## Excel COM-ის შემდგომი ეტაპები

AutoFit eligibility-ის შემდეგ ცალკე მცირე milestone-ებად:

1. horizontal pagination და wide-sheet safeguards;
2. `print-layout-report.json` writing;
3. root `excel-package.ps1`-ის საბოლოო orchestration script-ად შემცირება.

ყოველი milestone-ის შემდეგ tests → commit → push.

## Excel-ის შემდეგ

1. `pdf-package.ps1` დაყავი inspection/OCR/extraction/rendering/quality/orchestration პასუხისმგებლობებად;
2. დაამატე ერთიანი `DocumentProfile`;
3. დაამატე ავტომატური `ProcessingPlan`;
4. ყველა package-ში შექმენი მოკლე `package-summary.md`;
5. warnings-ს მიეცი severity/location/type;
6. დაამატე selective rendering დიდი ფაილებისთვის;
7. მხოლოდ ამის შემდეგ შექმენი Word/PowerPoint/CSV და სხვა ფორმატების ცალკე processors.

სასურველი საერთო processor contract:

```text
inspect → plan → process → verify → summarize
```

## კოდის სტილის პრინციპები

- თითოეულ მოდულს ერთი მკაფიო პასუხისმგებლობა ჰქონდეს;
- root entry point-ები პატარა და უკან თავსებადი დარჩეს;
- ფუნქციის სახელი უნდა ხსნიდეს მოქმედებას;
- კომენტარი ხსნიდეს „რატომ“-ს, განსაკუთრებით COM/Windows/უსაფრთხოების უცნაურობებზე;
- მარტივი ხაზები ზედმეტი კომენტარებით არ გადატვირთო;
- რთულ ფუნქციას ჰქონდეს მოკლე docstring/comment-based help;
- ქცევის რეფაქტორი და ახალი feature ერთ commit-ში არ აურიო;
- output schema და public CLI arguments განზრახ არ შეცვალო შესაბამისი migration გეგმის გარეშე.

## უსაფრთხოების უცვლელი წესები

- source ფაილი არ შეცვალო და არ შეინახო;
- Excel workbook read-only რეჟიმში გახსენი;
- `AutomationSecurity = 3` შეინარჩუნე;
- `EnableEvents = $false` შეინარჩუნე;
- `AskToUpdateLinks = $false` შეინარჩუნე;
- COM objects ყოველთვის release/cleanup უნდა გაიაროს;
- atomic build მხოლოდ სრული წარმატების შემდეგ finalize-დება;
- cleanup მხოლოდ verified output/temp path-ზე შესრულდეს;
- არსებული კარგი package შეცდომისას უნდა გადარჩეს;
- installer და portable ZIP ყოველი ახალი code location-ისას გადაამოწმე.

## მომხმარებელთან კომუნიკაცია

მომხმარებელი დამწყები პროგრამისტია. ახსენი მარტივი ქართულით:

- რა ნაწილი რაზეა პასუხისმგებელი;
- რატომ კეთდება ცვლილება;
- რა რისკია;
- რომელი ტესტი ამტკიცებს გამართულ მუშაობას;
- რა დარჩა შემდეგ ეტაპზე.

არ განაცხადო „სრულად მუშაობს“ მხოლოდ unit tests-ის საფუძველზე. PDF/Excel/OCR-ის მნიშვნელოვანი ცვლილების შემდეგ მიუთითე სრული E2E შედეგიც.

## სავარაუდო პროგრესი

სრული დაგეგმილი რეფაქტორის დაახლოებით 55–60% დასრულებულია. დარჩენილია:

- Excel PowerShell COM/layout pipeline;
- PDF/OCR pipeline;
- intelligent profiling/planning layer;
- package summary/severity/selective rendering;
- ფორმატების გაფართოება;
- installer/portable final QA;
- PR, `main` merge და ახალი release.
