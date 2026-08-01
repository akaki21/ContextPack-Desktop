# ContextPack Desktop-ის ინსტალაცია

## რეკომენდებული გზა: Windows installer

1. GitHub-ის უახლესი Release-დან ჩამოტვირთე `ContextPack-Setup-2.2.0.exe`.
2. სურვილის შემთხვევაში ფაილი გადაამოწმე `SHA256SUMS-2.2.0.txt`-თან.
3. installer-ს ორჯერ დააწკაპუნე. თუ Desktop shortcut გინდა, შესაბამისი არჩევანი დატოვე ჩართული.
4. პირველი გამართვისას ინტერნეტი ჩართული უნდა იყოს. ContextPack ავტომატურად ამოწმებს Python-ს, Tesseract OCR-ს, ვირტუალურ გარემოს, Python-ის პაკეტებს და ქართულ/ინგლისურ OCR მოდელებს; საჭირო კომპონენტებს თავად აყენებს.
5. Start მენიუდან ან Desktop-იდან გახსენი **ContextPack Desktop**.

Microsoft Excel სავალდებულო არაა. ის საჭიროა მხოლოდ Excel-ის PDF/PNG ვიზუალური ხედის ზუსტი გენერირებისთვის; მნიშვნელობებისა და ფორმულების ამოღება მის გარეშეც შესაძლებელია.

## თუ ინსტალაციისას შეცდომა გამოჩნდა

Start მენიუდან გახსენი **ContextPack Setup Repair**. ის კომპონენტების მომზადებას თავიდან ცდის და დეტალებს აქ ინახავს:

```text
%LOCALAPPDATA%\Programs\ContextPack Desktop\logs\install.log
%LOCALAPPDATA%\Programs\ContextPack Desktop\logs\last-install-error.txt
```

GUI-ში ასევე არის **გარემოს შემოწმება** და **გარემოს აღდგენა**.

## Portable ZIP

1. `ContextPack-Desktop-2.2.0-portable.zip` ამოიღე ჩვეულებრივ, ჩაწერად საქაღალდეში.
2. ამ საქაღალდეში გახსენი PowerShell.
3. გაუშვი:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\setup.ps1
.\Start-ContextPack-GUI.cmd
```

პროგრამა პირდაპირ ZIP-ის შიგნიდან არ გაუშვა. `.venv`-ის გააქტიურება საჭირო არაა.

## წაშლა

გახსენი **Settings → Apps → Installed apps**, მონახე **ContextPack Desktop** და აირჩიე **Uninstall**. იგივე შეგიძლია Start მენიუში **Uninstall ContextPack Desktop**-ით. იშლება პროგრამა, ლოკალური Python გარემო, OCR მოდელები, ლოგები და shortcut-ები. მომხმარებლის დოკუმენტების დასაცავად შევსებული `input` და `output` საქაღალდეები არ იშლება.

## SHA-256-ით გადამოწმება

```powershell
Get-FileHash .\ContextPack-Setup-2.2.0.exe -Algorithm SHA256
```

მიღებული მნიშვნელობა შეადარე `SHA256SUMS-2.2.0.txt`-ის შესაბამის ხაზს.
