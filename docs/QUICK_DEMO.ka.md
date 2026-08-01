# ContextPack Desktop — 60-წამიანი დემო

დემო იყენებს რეპოზიტორის პატარა, არაკონფიდენციალურ `examples\sample.csv` ფაილს. დაწყებამდე ერთხელ დაასრულე `setup.ps1`.

## GUI

1. ორჯერ დააწკაპუნე `Start-ContextPack-GUI.cmd` ფაილზე.
2. აირჩიე `examples\sample.csv`.
3. დატოვე **ავტომატური — რეკომენდებული**.
4. დააჭირე **დამუშავების დაწყებას** და აირჩიე დანიშნულების საქაღალდე, მაგალითად Desktop.
5. როდესაც პროგრესი 100%-ს მიაღწევს, გახსენი შედეგი.

არჩეულ ადგილას უნდა შეიქმნას `sample.md`. მისი ცხრილი უნდა ემთხვეოდეს [მოსალოდნელ შედეგს](../examples/demo-output/sample.md).

## ბრძანების სტრიქონი

```powershell
$demoOutput = Join-Path $env:TEMP 'ContextPack-Demo'
.\contextpack.ps1 ".\examples\sample.csv" -OutputDirectory $demoOutput
```

მოსალოდნელი ბოლო შეტყობინება:

```text
Ready: ...\ContextPack-Demo\sample.md
```

ეს მოკლე დემო ამოწმებს launcher-ს, გარემოს, router-ს, MarkItDown conversion-ს, არჩეულ output-ში ჩაწერასა და საბოლოო შედეგის დამუშავებას. სრული ვიზუალური პაკეტის შესამოწმებლად შემდეგ გამოიყენე PDF ან Excel workbook.

[English version](QUICK_DEMO.md)
