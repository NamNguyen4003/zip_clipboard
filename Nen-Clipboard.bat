<# : phần cmd: chạy PowerShell đọc lại chính file này (cmd bỏ qua dòng nhãn, PowerShell coi đây là comment)
@echo off
start "" powershell.exe -STA -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -Command "& ([scriptblock]::Create((Get-Content -LiteralPath '%~f0' -Raw -Encoding UTF8))) %*"
exit /b
#>
# Nén ảnh trong clipboard (copy từ OneNote) để paste sang New Outlook không bị treo.
# Dùng: copy trang OneNote (desktop hoặc web) -> chạy script (phím tắt) -> Ctrl+V vào Outlook.
# Tự kiểm tra (xem kết quả): powershell -STA -NoProfile -ExecutionPolicy Bypass -Command "& ([scriptblock]::Create((Get-Content -LiteralPath 'Nen-Clipboard.bat' -Raw -Encoding UTF8))) -SelfTest"
# Chẩn đoán: thêm -Diag (ghi báo cáo cấu trúc clipboard ra Desktop) hoặc -TestPaste (đặt 2 ảnh test lên clipboard).
param([switch]$SelfTest, [switch]$Diag, [switch]$TestPaste)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

# Clipboard qua Win32 để đọc/ghi "HTML Format" dạng byte thô (UTF-8), tránh .NET làm hỏng tiếng Việt.
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Threading;
public static class ClipNative {
    [DllImport("user32.dll")] static extern bool OpenClipboard(IntPtr h);
    [DllImport("user32.dll")] static extern bool CloseClipboard();
    [DllImport("user32.dll")] static extern bool EmptyClipboard();
    [DllImport("user32.dll")] static extern IntPtr GetClipboardData(uint f);
    [DllImport("user32.dll")] static extern IntPtr SetClipboardData(uint f, IntPtr h);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern uint RegisterClipboardFormat(string s);
    [DllImport("kernel32.dll")] static extern IntPtr GlobalLock(IntPtr h);
    [DllImport("kernel32.dll")] static extern bool GlobalUnlock(IntPtr h);
    [DllImport("kernel32.dll")] static extern UIntPtr GlobalSize(IntPtr h);
    [DllImport("kernel32.dll")] static extern IntPtr GlobalAlloc(uint flags, UIntPtr size);
    [DllImport("kernel32.dll")] static extern IntPtr GlobalFree(IntPtr h);
    const uint CF_UNICODETEXT = 13;
    static uint HtmlFormat { get { return RegisterClipboardFormat("HTML Format"); } }
    [DllImport("user32.dll")] static extern uint EnumClipboardFormats(uint f);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClipboardFormatName(uint f, System.Text.StringBuilder sb, int n);
    public static string[] ListFormats() {
        Open();
        try {
            var l = new System.Collections.Generic.List<string>(); uint f = 0;
            while ((f = EnumClipboardFormats(f)) != 0) {
                var sb = new System.Text.StringBuilder(256);
                l.Add(f + " " + (GetClipboardFormatName(f, sb, 256) > 0 ? sb.ToString() : "(chuẩn)"));
            }
            return l.ToArray();
        } finally { CloseClipboard(); }
    }

    // Pixel BGRA 32bpp: true nếu mọi byte alpha đều 255 (ảnh không trong suốt).
    public static bool IsOpaque(byte[] bgra) {
        for (int i = 3; i < bgra.Length; i += 4) if (bgra[i] != 255) return false;
        return true;
    }

    static void Open() {
        for (int i = 0; i < 40; i++) { if (OpenClipboard(IntPtr.Zero)) return; Thread.Sleep(50); }
        throw new Exception("Clipboard đang bị chương trình khác giữ, thử lại.");
    }
    static byte[] Read(uint f) {
        IntPtr h = GetClipboardData(f);
        if (h == IntPtr.Zero) return null;
        IntPtr p = GlobalLock(h);
        try { int n = (int)GlobalSize(h); byte[] b = new byte[n]; Marshal.Copy(p, b, 0, n); return b; }
        finally { GlobalUnlock(h); }
    }
    static void Put(uint f, byte[] b) {
        IntPtr h = GlobalAlloc(0x0002, (UIntPtr)b.Length);
        if (h == IntPtr.Zero) throw new OutOfMemoryException("GlobalAlloc");
        IntPtr p = GlobalLock(h); Marshal.Copy(b, 0, p, b.Length); GlobalUnlock(h);
        if (SetClipboardData(f, h) == IntPtr.Zero) { GlobalFree(h); throw new Exception("SetClipboardData lỗi"); }
    }
    public static byte[] GetHtml() { Open(); try { return Read(HtmlFormat); } finally { CloseClipboard(); } }
    public static string GetText() {
        Open();
        try {
            byte[] b = Read(CF_UNICODETEXT); if (b == null) return "";
            string s = System.Text.Encoding.Unicode.GetString(b); int z = s.IndexOf('\0');
            return z >= 0 ? s.Substring(0, z) : s;
        } finally { CloseClipboard(); }
    }
    public static void SetHtmlAndText(byte[] html, string text) {
        byte[] h = new byte[html.Length + 1]; Buffer.BlockCopy(html, 0, h, 0, html.Length);
        byte[] t = System.Text.Encoding.Unicode.GetBytes(text + "\0");
        Open();
        try { EmptyClipboard(); Put(HtmlFormat, h); Put(CF_UNICODETEXT, t); } finally { CloseClipboard(); }
    }
}
'@

# CF_HTML: header chứa byte offset (không phải offset ký tự) vào phần UTF-8 phía sau.
function Read-CfHtml([byte[]]$raw) {
    $head = [Text.Encoding]::ASCII.GetString($raw, 0, [Math]::Min(400, $raw.Length))
    $get = { param($k) if ($head -match "${k}:(-?\d+)") { [int]$matches[1] } else { -1 } }
    $s = & $get 'StartHTML'; $e = & $get 'EndHTML'
    if ($s -lt 0) { $s = & $get 'StartFragment'; $e = & $get 'EndFragment' }
    if ($s -lt 0) { throw 'Header "HTML Format" không hợp lệ.' }
    if ($e -lt 0 -or $e -gt $raw.Length) { $e = $raw.Length }
    while ($e -gt $s -and $raw[$e - 1] -eq 0) { $e-- }
    [Text.Encoding]::UTF8.GetString($raw, $s, $e - $s)
}

function Build-CfHtml([string]$html) {
    $u = [Text.Encoding]::UTF8
    $tpl = "Version:0.9`r`nStartHTML:{0:D10}`r`nEndHTML:{1:D10}`r`nStartFragment:{2:D10}`r`nEndFragment:{3:D10}`r`n"
    $hl = ($tpl -f 0, 0, 0, 0).Length
    $body = $u.GetBytes($html)
    $sm = '<!--StartFragment-->'; $em = '<!--EndFragment-->'
    $i = $html.IndexOf($sm, [StringComparison]::Ordinal); $j = $html.IndexOf($em, [StringComparison]::Ordinal)
    if ($i -ge 0 -and $j -gt $i) {
        $fs = $hl + $u.GetByteCount($html.Substring(0, $i + $sm.Length))
        $fe = $hl + $u.GetByteCount($html.Substring(0, $j))
    } else { $fs = $hl; $fe = $hl + $body.Length }
    $head = $u.GetBytes(($tpl -f $hl, ($hl + $body.Length), $fs, $fe))
    $out = New-Object byte[] ($head.Length + $body.Length)
    [Array]::Copy($head, $out, $head.Length)
    [Array]::Copy($body, 0, $out, $head.Length, $body.Length)
    , $out
}

function Get-Mime([byte[]]$b) {
    if ($b.Length -lt 4) { return $null }
    if ($b[0] -eq 0x89 -and $b[1] -eq 0x50) { return 'image/png' }
    if ($b[0] -eq 0xFF -and $b[1] -eq 0xD8) { return 'image/jpeg' }
    if ($b[0] -eq 0x47 -and $b[1] -eq 0x49 -and $b[2] -eq 0x46) { return 'image/gif' }
    $null
}

# Kích thước hiển thị (CSS px) từ thuộc tính width/height hoặc style.
function Get-Dim([string]$tag, [string]$name) {
    if ($tag -match "(?i)\s$name\s*=\s*[""']?(\d+(?:\.\d+)?)") { return [double]$matches[1] }
    if ($tag -match "(?i)[;""'\s]$name\s*:\s*(\d+(?:\.\d+)?)\s*(px|pt|in)") {
        $v = [double]$matches[1]
        if ($matches[2] -eq 'pt') { return $v * 96 / 72 }
        if ($matches[2] -eq 'in') { return $v * 96 }
        return $v
    }
    0
}

# Vẽ về w x h rồi chọn PNG hoặc JPEG: ảnh không trong suốt mà JPEG (chất lượng 85) nhỏ hơn 1/3 PNG thì lấy JPEG.
function Resize-Encode($bmp, [int]$w, [int]$h) {
    $dst = New-Object Drawing.Bitmap $w, $h
    $g = [Drawing.Graphics]::FromImage($dst)
    $attr = New-Object Drawing.Imaging.ImageAttributes
    $ms = New-Object IO.MemoryStream
    try {
        $g.InterpolationMode = 'HighQualityBicubic'
        $g.PixelOffsetMode = 'HighQuality'
        $attr.SetWrapMode('TileFlipXY')  # tránh viền mờ ở mép ảnh khi thu nhỏ
        $g.DrawImage($bmp, (New-Object Drawing.Rectangle 0, 0, $w, $h), 0, 0, $bmp.Width, $bmp.Height, 'Pixel', $attr)
        $dst.Save($ms, [Drawing.Imaging.ImageFormat]::Png)
        $png = $ms.ToArray()
        $d = $dst.LockBits((New-Object Drawing.Rectangle 0, 0, $w, $h), 'ReadOnly', $dst.PixelFormat)
        try { $px = New-Object byte[] ($d.Stride * $h); [Runtime.InteropServices.Marshal]::Copy($d.Scan0, $px, 0, $px.Length) } finally { $dst.UnlockBits($d) }
        if (-not [ClipNative]::IsOpaque($px)) { return , $png }
        $ms.SetLength(0)
        $codec = [Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' }
        $ep = New-Object Drawing.Imaging.EncoderParameters 1
        $ep.Param[0] = New-Object Drawing.Imaging.EncoderParameter ([Drawing.Imaging.Encoder]::Quality), 85L
        try { $dst.Save($ms, $codec, $ep) } finally { $ep.Param[0].Dispose(); $ep.Dispose() }
        if ($ms.Length * 3 -lt $png.Length) { , $ms.ToArray() } else { , $png }
    } finally { $ms.Dispose(); $attr.Dispose(); $g.Dispose(); $dst.Dispose() }
}

# Thu ảnh về kích thước hiển thị; giữ nguyên byte gốc nếu không nhỏ đi được.
# Ảnh không lớn hơn kích thước hiển thị: JPEG/GIF giữ nguyên (chạy lại không đổi thêm), PNG chỉ đổi sang JPEG nếu lợi.
function Shrink-Image([byte[]]$orig, [double]$w, [double]$h) {
    $ms = New-Object IO.MemoryStream (, $orig)
    try { $bmp = [Drawing.Image]::FromStream($ms) } catch { $ms.Dispose(); return , $orig }
    try {
        $srcW = $bmp.Width; $srcH = $bmp.Height  # PowerShell không phân biệt hoa thường: không đặt tên $W/$H
        if ($w -le 0 -and $h -le 0) { $w = $srcW; $h = $srcH }
        elseif ($h -le 0) { $h = $srcH * $w / $srcW }
        elseif ($w -le 0) { $w = $srcW * $h / $srcH }
        $tw = [Math]::Max(1, [int][Math]::Round($w)); $th = [Math]::Max(1, [int][Math]::Round($h))
        $mime = Get-Mime $orig
        $fits = -not ($tw -lt $srcW -and $th -lt $srcH)
        if ($fits) {
            if ($mime -eq 'image/jpeg' -or $mime -eq 'image/gif') { return , $orig }
            $tw = $srcW; $th = $srcH  # PNG: mã hóa lại ở cỡ gốc; BMP... (trình duyệt không hiển thị): đổi sang PNG/JPEG
        }
        $new = Resize-Encode $bmp $tw $th
        if ($mime -and $new.Length -ge $orig.Length) { return , $orig }
        if ($fits -and $mime -eq 'image/png' -and (Get-Mime $new) -eq $mime) { return , $orig }  # PNG mã hóa lại chỉ lệch vài byte: không coi là đổi
        , $new
    } finally { $bmp.Dispose(); $ms.Dispose() }
}

# Thẻ <img>: thuộc tính trong ngoặc kép/đơn có thể chứa '>' (vd alt="log > 12"); ngoặc không đóng thì cắt ở '>' đầu tiên.
$ImgRe = '(?is)<img\b(?:(?>[^>"'']+|"[^"]*"|''[^'']*'')*>|[^>]*>)'

# Ảnh file: (OneNote desktop) hoặc data: base64 (OneNote web). Mỗi ảnh bị bỏ qua ghi lý do vào $st.Skip để báo người dùng.
function Convert-ImgTag([string]$tag, $st) {
    $m = [regex]::Match($tag, '(?i)\bsrc\s*=\s*(?:"([^"]*)"|''([^'']*)''|([^\s>]+))')
    if (-not $m.Success) { $st.Skip['không có src']++; return $tag }
    $src = $m.Groups[1].Value + $m.Groups[2].Value + $m.Groups[3].Value
    $isData = $src -match '^data:image/[^;,]*;base64,'
    if ($isData) {
        try { $orig = [Convert]::FromBase64String($src.Substring($matches[0].Length)) } catch { $st.Skip['định dạng không hỗ trợ']++; return $tag }
    } elseif ($src -match '^file:') {
        $path = ([Uri]$src).LocalPath
        if (-not (Test-Path -LiteralPath $path)) { $st.Skip['file ảnh không tồn tại']++; return $tag }
        $orig = [IO.File]::ReadAllBytes($path)
    } else {
        $why = if ($src -match '^https?:') { 'link web' } else { 'kiểu khác' }
        $st.Skip[$why]++; return $tag
    }
    $bytes = Shrink-Image $orig (Get-Dim $tag 'width') (Get-Dim $tag 'height')
    $mime = Get-Mime $bytes
    if (-not $mime) { $st.Skip['định dạng không hỗ trợ']++; return $tag }
    $st.Count++; $st.Before += $orig.Length; $st.After += $bytes.Length
    if ($isData -and [object]::ReferenceEquals($bytes, $orig)) { return $tag }  # data: không đổi: giữ nguyên thẻ
    $st.Changed++  # file: luôn tính (phải nhúng vì Outlook không đọc được file:)
    $tag.Substring(0, $m.Index) + 'src="data:' + $mime + ';base64,' + [Convert]::ToBase64String($bytes) + '"' + $tag.Substring($m.Index + $m.Length)
}

function Compress-Html([string]$html) {
    $st = @{ Count = 0; Changed = 0; Before = 0L; After = 0L; Skip = @{} }
    # Office ghi mỗi ảnh hai lần (VML + <img>); bỏ bản VML.
    $html = [regex]::Replace($html, '(?s)<!--\[if gte vml 1\]>.*?<!\[endif\]-->', '')
    $html = $html.Replace('<![if !vml]>', '').Replace('<![endif]>', '')
    $sb = New-Object Text.StringBuilder
    $last = 0
    foreach ($m in [regex]::Matches($html, $ImgRe)) {
        [void]$sb.Append($html, $last, $m.Index - $last)
        [void]$sb.Append((Convert-ImgTag $m.Value $st))
        $last = $m.Index + $m.Length
    }
    [void]$sb.Append($html, $last, $html.Length - $last)
    [pscustomobject]@{ Html = $sb.ToString(); Count = $st.Count; Changed = $st.Changed; Before = $st.Before; After = $st.After; Skip = $st.Skip }
}

# Chỉ ghi clipboard ở bước cuối: lỗi ở giữa chừng thì clipboard gốc vẫn còn.
function Invoke-Compress {
    $raw = [ClipNative]::GetHtml()
    if (-not $raw) { throw 'Clipboard không có nội dung HTML. Hãy copy trang OneNote trước.' }
    $html = Read-CfHtml $raw
    $text = [ClipNative]::GetText()
    $r = Compress-Html $html
    if ($r.Count -eq 0) {
        if ($r.Skip.Count -eq 0) { throw ('Phần đã copy không có ảnh (HTML {0:N0} KB). Trong OneNote bấm Ctrl+A vài lần tới khi chọn cả trang, Ctrl+C, rồi chạy lại.' -f ($html.Length / 1KB)) }
        $why = ($r.Skip.GetEnumerator() | ForEach-Object { "$($_.Key): $($_.Value)" }) -join ', '
        throw "Không nén được ảnh nào ($why)."
    }
    if ($r.Changed -eq 0) { return $r }  # mọi ảnh đã đủ nhỏ: không ghi lại clipboard
    [ClipNative]::SetHtmlAndText((Build-CfHtml $r.Html), $text)
    $r
}

function Show-Note([string]$text, [bool]$isError = $false) {
    $n = New-Object Windows.Forms.NotifyIcon
    $n.Icon = if ($isError) { [Drawing.SystemIcons]::Error } else { [Drawing.SystemIcons]::Information }
    $n.Visible = $true
    $n.ShowBalloonTip(5000, 'OneNote -> Outlook', $text, $(if ($isError) { 'Error' } else { 'Info' }))
    Start-Sleep -Seconds 5
    $n.Dispose()
}

# ---------------- Self-test: report giả ~30MB, chạy qua clipboard thật ----------------
function New-FakeShot([string]$path, [int]$w, [int]$h, $rnd) {
    $bmp = New-Object Drawing.Bitmap $w, $h, ([Drawing.Imaging.PixelFormat]::Format24bppRgb)
    $g = [Drawing.Graphics]::FromImage($bmp)
    $g.Clear([Drawing.Color]::White)
    for ($k = 0; $k -lt 15; $k++) {
        $br = New-Object Drawing.SolidBrush ([Drawing.Color]::FromArgb($rnd.Next(256), $rnd.Next(256), $rnd.Next(256)))
        $g.FillRectangle($br, $rnd.Next($w), $rnd.Next($h), $rnd.Next(50, [int]($w / 3)), $rnd.Next(20, [int]($h / 4)))
        $br.Dispose()
    }
    $font = New-Object Drawing.Font 'Segoe UI', 14
    for ($y = 0; $y -lt $h; $y += 40) { $g.DrawString("Dòng $y - log 12:34:56 giá trị=$($rnd.Next())", $font, [Drawing.Brushes]::Black, 10, $y) }
    $font.Dispose(); $g.Dispose()
    # Dải nhiễu ~1MB để PNG không nén quá nhỏ (giống screenshot có ảnh/gradient).
    $band = [Math]::Min($h, [int](1MB / ($w * 3)))
    $data = $bmp.LockBits((New-Object Drawing.Rectangle 0, 0, $w, $band), 'WriteOnly', $bmp.PixelFormat)
    $buf = New-Object byte[] ($data.Stride * $band); $rnd.NextBytes($buf)
    [Runtime.InteropServices.Marshal]::Copy($buf, 0, $data.Scan0, $buf.Length)
    $bmp.UnlockBits($data)
    $bmp.Save($path, [Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
}

function Test-CfHtml([byte[]]$raw, [string]$label) {
    $head = [Text.Encoding]::ASCII.GetString($raw, 0, 200)
    $n = @{}
    foreach ($k in 'StartHTML', 'EndHTML', 'StartFragment', 'EndFragment') { if ($head -match "${k}:(\d+)") { $n[$k] = [int]$matches[1] } }
    $u = [Text.Encoding]::UTF8
    $sm = '<!--StartFragment-->'; $em = '<!--EndFragment-->'
    Check ($u.GetString($raw, $n.StartHTML, 5) -eq '<html') "${label}: StartHTML trỏ đúng <html"
    Check ($u.GetString($raw, $n.EndHTML - 7, 7) -eq '</html>') "${label}: EndHTML trỏ đúng cuối </html>"
    Check ($u.GetString($raw, $n.StartFragment - $sm.Length, $sm.Length) -eq $sm) "${label}: StartFragment đúng"
    Check ($u.GetString($raw, $n.EndFragment, $em.Length) -eq $em) "${label}: EndFragment đúng"
}

function Check($ok, [string]$msg) {
    if ($ok) { Write-Host "  ok   $msg" } else { Write-Host "  FAIL $msg" -ForegroundColor Red; $script:fail++ }
}

# Ca 0 ảnh: đếm đúng lý do bỏ qua, và Invoke-Compress phải throw mà không đụng clipboard.
function Test-ZeroImages([string]$dir) {
    $png = Join-Path $dir 'tiny.png'
    $bmp = New-Object Drawing.Bitmap 4, 4; $bmp.Save($png, [Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
    $u = ([Uri]$png).AbsoluteUri
    $miss = ([Uri](Join-Path $dir 'khong-co.png')).AbsoluteUri
    $vml = '<!--[if gte vml 1]><v:shape><v:imagedata src="' + $miss + '"/></v:shape><![endif]-->'

    $a = Compress-Html ('<p><img alt="log > 12" src="' + $u + '"></p>')
    Check ($a.Count -eq 1) 'alt chứa > đứng trước src: vẫn xử lý 1 ảnh'
    $q = Compress-Html ('<img alt=5" src=' + $u + '>')
    Check ($q.Count -eq 1) 'thẻ có ngoặc không đóng: vẫn xử lý 1 ảnh'
    $b = Compress-Html '<img src="data:image/png;base64,AAAA"><img src="https://x/y.png"><img alt="x">'
    Check ($b.Count -eq 0 -and $b.Skip['định dạng không hỗ trợ'] -eq 1 -and $b.Skip['link web'] -eq 1 -and $b.Skip['không có src'] -eq 1) 'data: không giải mã được/https:/không src: Count=0 và đúng lý do'
    $b2 = Compress-Html '<img src="data:text/plain;base64,AAAA"><img src="data:image/svg+xml,%3Csvg%3E">'
    Check ($b2.Count -eq 0 -and $b2.Skip['kiểu khác'] -eq 2) 'data: không phải ảnh base64: Count=0, lý do kiểu khác'
    $c = Compress-Html ('<img src="' + $miss + '">')
    Check ($c.Count -eq 0 -and $c.Skip['file ảnh không tồn tại'] -eq 1) 'file thiếu: Count=0 và đúng lý do'
    $d = Compress-Html $vml
    Check ($d.Count -eq 0 -and $d.Skip.Count -eq 0) 'chỉ VML: Count=0, không có thẻ <img>'

    foreach ($case in @(@('file thiếu', ('<p>Chữ gốc</p><img src="' + $miss + '">'), 'Không nén được'), @('chỉ VML', ('<p>Chữ gốc</p>' + $vml), 'không có ảnh'))) {
        $orig = Build-CfHtml ('<html><body><!--StartFragment-->' + $case[1] + '<!--EndFragment--></body></html>')
        [ClipNative]::SetHtmlAndText($orig, 'Chữ gốc')
        $orig = [ClipNative]::GetHtml()
        $err = $null
        try { [void](Invoke-Compress) } catch { $err = $_.Exception.Message }
        Check ($null -ne $err) "$($case[0]): Invoke-Compress throw khi 0 ảnh"
        Check ($err -and $err.Contains($case[2])) "$($case[0]): thông báo nêu đúng lý do"
        Check ($err -and ($err.Length + 'Lỗi:  Clipboard giữ nguyên.'.Length) -lt 255) "$($case[0]): thông báo < 255 ký tự"
        Check ([Convert]::ToBase64String([ClipNative]::GetHtml()) -eq [Convert]::ToBase64String($orig)) "$($case[0]): clipboard giữ nguyên byte"
    }
}

# Ảnh test cho data: URI. 'photo' = gradient + nhiễu nhẹ (như ảnh chụp), 'alpha' = như photo nhưng chỉ vẽ hình elip trên nền trong suốt, 'text' = chữ trên nền trắng (như screenshot).
function New-TestBitmap([string]$kind, [int]$w, [int]$h, $rnd) {
    $bmp = New-Object Drawing.Bitmap $w, $h, ([Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [Drawing.Graphics]::FromImage($bmp)
    $rect = New-Object Drawing.Rectangle 0, 0, $w, $h
    try {
        if ($kind -eq 'text') {
            $g.Clear([Drawing.Color]::White)
            $font = New-Object Drawing.Font 'Segoe UI', 14
            for ($y = 0; $y -lt $h; $y += 40) { $g.DrawString("Dòng $y - log 12:34:56", $font, [Drawing.Brushes]::Black, 10, $y) }
            $font.Dispose()
            $g.FillRectangle([Drawing.Brushes]::SteelBlue, 600, 100, 300, 120)
            return $bmp
        }
        $c1 = [Drawing.Color]::FromArgb($rnd.Next(256), $rnd.Next(256), $rnd.Next(256)); $c2 = [Drawing.Color]::FromArgb($rnd.Next(256), $rnd.Next(256), $rnd.Next(256))
        $br = New-Object Drawing.Drawing2D.LinearGradientBrush $rect, $c1, $c2, 45.0
        if ($kind -eq 'alpha') { $g.FillEllipse($br, $rect) } else { $g.FillRectangle($br, $rect) }
        $br.Dispose()
        $nz = New-Object Drawing.Bitmap $w, $h, ([Drawing.Imaging.PixelFormat]::Format24bppRgb)
        $d = $nz.LockBits($rect, 'WriteOnly', $nz.PixelFormat)
        $buf = New-Object byte[] ($d.Stride * $h); $rnd.NextBytes($buf)
        [Runtime.InteropServices.Marshal]::Copy($buf, 0, $d.Scan0, $buf.Length)
        $nz.UnlockBits($d)
        $cm = New-Object Drawing.Imaging.ColorMatrix; $cm.Matrix33 = 0.06
        $attr = New-Object Drawing.Imaging.ImageAttributes; $attr.SetColorMatrix($cm)
        $g.DrawImage($nz, $rect, 0, 0, $w, $h, 'Pixel', $attr)
        $attr.Dispose(); $nz.Dispose()
        $bmp
    } catch { $bmp.Dispose(); throw } finally { $g.Dispose() }
}

function New-TestPng([string]$kind, [int]$w, [int]$h, $rnd) {
    $bmp = New-TestBitmap $kind $w $h $rnd
    $ms = New-Object IO.MemoryStream
    $bmp.Save($ms, [Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
    , $ms.ToArray()
}

# Thẻ <img> giống OneNote web: src data:, style dài, kích thước hiển thị ở cuối style.
function New-WebImg([byte[]]$b, [int]$dw, [int]$dh) {
    '<img class="Image" src="data:' + (Get-Mime $b) + ';base64,' + [Convert]::ToBase64String($b) + '" style="font-family: ''Segoe UI''; border-width: medium; border-style: none; width: ' + $dw + 'px; height: ' + $dh + 'px;">'
}

# Giải mã mọi ảnh data: trong HTML thành mime, kích thước, alpha pixel (0,0).
function Get-WebImgs([string]$html) {
    foreach ($m in [regex]::Matches($html, 'src="data:image/\w+;base64,([^"]+)"')) {
        $b = [Convert]::FromBase64String($m.Groups[1].Value)
        $ms = New-Object IO.MemoryStream (, $b); $img = New-Object Drawing.Bitmap $ms
        [pscustomobject]@{ Mime = Get-Mime $b; W = $img.Width; H = $img.Height; A = $img.GetPixel(0, 0).A; Len = $b.Length }
        $img.Dispose(); $ms.Dispose()
    }
}

# Ảnh data: (OneNote web): thu về kích thước hiển thị, chọn PNG hoặc JPEG.
function Test-DataImages {
    $rnd = New-Object Random 7
    $a = Compress-Html ('<p>' + (New-WebImg (New-TestPng 'photo' 1024 690 $rnd) 480 323) + '</p>')
    $ia = @(Get-WebImgs $a.Html)
    Check ($a.Count -eq 1 -and $ia.Count -eq 1 -and $ia[0].W -eq 480 -and $ia[0].H -eq 323) 'data: PNG 1024x690 hiển thị 480x323: Count=1, ra 480x323'

    $photo = Get-WebImgs (Compress-Html (New-WebImg (New-TestPng 'photo' 1024 690 $rnd) 480 323)).Html
    Check ($photo.Mime -eq 'image/jpeg') "ảnh chụp đục: ra image/jpeg (thực tế $($photo.Mime))"
    $text = Get-WebImgs (Compress-Html (New-WebImg (New-TestPng 'text' 1024 690 $rnd) 480 323)).Html
    Check ($text.Mime -eq 'image/png') "screenshot chữ: ra image/png (thực tế $($text.Mime))"
    $alpha = Get-WebImgs (Compress-Html (New-WebImg (New-TestPng 'alpha' 1024 690 $rnd) 480 323)).Html
    Check ($alpha.Mime -eq 'image/png' -and $alpha.A -lt 255) "ảnh trong suốt: ra image/png, còn pixel trong suốt (mime $($alpha.Mime), alpha $($alpha.A))"

    # Ảnh đã nhỏ sẵn: PNG phẳng và JPEG phải giữ nguyên thẻ, không tính là đổi.
    $flat = New-Object Drawing.Bitmap 100, 50; $ms = New-Object IO.MemoryStream
    $flat.Save($ms, [Drawing.Imaging.ImageFormat]::Jpeg); $jpg = $ms.ToArray(); $flat.Dispose(); $ms.Dispose()
    $same = (New-WebImg (New-TestPng 'text' 100 50 $rnd) 100 50) + (New-WebImg $jpg 100 50)
    $s = Compress-Html $same
    Check ($s.Count -eq 2 -and $s.Changed -eq 0 -and $s.Html -ceq $same) 'ảnh nhỏ sẵn (PNG chữ, JPEG): Count=2, Changed=0, thẻ giữ nguyên'

    $bad = $null; $e = $null
    try { $bad = Compress-Html '<img src="data:image/png;base64,@@@không-phải-base64">' } catch { $e = $_.Exception.Message }
    Check ($null -eq $e -and $bad.Count -eq 0 -and $bad.Skip['định dạng không hỗ trợ'] -eq 1) "base64 hỏng: Count=0, bỏ qua 'định dạng không hỗ trợ', không throw ($e)"
}

# Report OneNote web qua clipboard thật: 17 ảnh data: PNG kiểu ảnh chụp; lần nén thứ hai không được đổi gì.
function Test-WebReport {
    $rnd = New-Object Random 11
    $viet = 'Báo cáo web: Tiếng Việt có dấu – ắ ằ ẳ ẵ ặ đ Ư ơ ữ'
    $sb = New-Object Text.StringBuilder
    [void]$sb.Append('<html><body><!--StartFragment-->')
    $before = 0L
    for ($i = 1; $i -le 17; $i++) {
        $png = New-TestPng 'photo' 1024 690 $rnd; $before += $png.Length
        [void]$sb.Append('<div class="OutlineElement"><p class="Paragraph"><span lang="VI">' + $viet + " $i</span></p><div>" + (New-WebImg $png 480 323) + '</div></div>')
    }
    [void]$sb.Append('<!--EndFragment--></body></html>')
    Check ($before -ge 15MB) ('report web đủ nặng: {0:N1} MB' -f ($before / 1MB))
    [ClipNative]::SetHtmlAndText((Build-CfHtml $sb.ToString()), $viet)

    $sw = [Diagnostics.Stopwatch]::StartNew(); $r = Invoke-Compress; $sw.Stop()
    $raw1 = [ClipNative]::GetHtml(); $out = Read-CfHtml $raw1
    $imgs = @(Get-WebImgs $out)
    Check ($r.Count -eq 17 -and $r.Changed -eq 17 -and $imgs.Count -eq 17) "report web: 17 ảnh, đổi 17 (Count $($r.Count), Changed $($r.Changed), ảnh $($imgs.Count))"
    Check (@($imgs | Where-Object { $_.W -ne 480 -or $_.H -ne 323 }).Count -eq 0) 'report web: mọi ảnh ra 480x323'
    Check (($r.Before / [double]$r.After) -ge 5) ('report web: giảm {0:N1} lần (cần >= 5)' -f ($r.Before / [double]$r.After))
    Check ($sw.Elapsed.TotalSeconds -lt 15) ('report web: {0:N1}s (cần < 15s)' -f $sw.Elapsed.TotalSeconds)
    Check ($out.Contains($viet + ' 17') -and [ClipNative]::GetText() -eq $viet) 'report web: tiếng Việt còn nguyên'

    $r2 = Invoke-Compress
    $raw2 = [ClipNative]::GetHtml()
    Check ($r2.Count -eq 17 -and $r2.Changed -eq 0) "nén lần hai: Count 17, Changed 0 (Count $($r2.Count), Changed $($r2.Changed))"
    Check ([Convert]::ToBase64String($raw2) -eq [Convert]::ToBase64String($raw1)) 'nén lần hai: clipboard giữ nguyên byte'
    Write-Host ('  report web: {0:N1} MB -> {1:N1} MB, {2:N1}s' -f ($r.Before / 1MB), ($r.After / 1MB), $sw.Elapsed.TotalSeconds)
}

function Invoke-SelfTest {
    $script:fail = 0
    $root = Join-Path $env:TEMP 'msohtmlclip_selftest'
    $dir = Join-Path $root '01'
    New-Item -ItemType Directory -Force $dir | Out-Null
    Test-ZeroImages $dir
    Test-DataImages
    Test-WebReport
    $rnd =New-Object Random 42
    $viet = 'Báo cáo tiến độ: Tiếng Việt có dấu – ắ ằ ẳ ẵ ặ đ Ư ơ ữ'
    $sb = New-Object Text.StringBuilder
    [void]$sb.Append('<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><head><meta http-equiv=Content-Type content="text/html; charset=utf-8"></head><body lang=VI>' + "`r`n<!--StartFragment--><p>$viet</p>")
    $files = @()
    Write-Host 'Tạo 30 ảnh 2560x1440 + 1 ảnh nhỏ...'
    for ($i = 1; $i -le 31; $i++) {
        $f = Join-Path $dir ('clip_image{0:D3}.png' -f $i)
        if ($i -le 30) { New-FakeShot $f 2560 1440 $rnd; $dw = 640; $dh = 360 } else { New-FakeShot $f 400 200 $rnd; $dw = 400; $dh = 200 }
        $u = ([Uri]$f).AbsoluteUri
        [void]$sb.Append("<p>Hình $i</p><!--[if gte vml 1]><v:shape id=""Picture_$i"" style='width:${dw}pt;height:${dh}pt'><v:imagedata src=""$u"" o:title=""""/></v:shape><![endif]--><![if !vml]><img width=$dw height=$dh src=""$u"" v:shapes=""Picture_$i""><![endif]>")
        $files += $f
    }
    [void]$sb.Append("<!--EndFragment-->`r`n</body></html>")
    $before = ($files | ForEach-Object { (Get-Item $_).Length } | Measure-Object -Sum).Sum

    [ClipNative]::SetHtmlAndText((Build-CfHtml $sb.ToString()), $viet)
    Test-CfHtml ([ClipNative]::GetHtml()) 'input'

    $sw = [Diagnostics.Stopwatch]::StartNew()
    $r = Invoke-Compress
    $sw.Stop()

    $raw = [ClipNative]::GetHtml()
    Test-CfHtml $raw 'output'
    $out = Read-CfHtml $raw
    Check ($before -ge 25MB) ('dữ liệu test đủ nặng: {0:N1} MB' -f ($before / 1MB))
    Check ($r.Count -eq 31) "xử lý 31 ảnh (thực tế $($r.Count))"
    Check ($out.IndexOf('file:///', [StringComparison]::Ordinal) -lt 0) 'không còn tham chiếu file:///'
    Check ($out.IndexOf('v:imagedata', [StringComparison]::Ordinal) -lt 0) 'đã bỏ khối VML'
    Check ($out.Contains($viet)) 'tiếng Việt trong HTML còn nguyên'
    Check ([ClipNative]::GetText() -eq $viet) 'text thuần còn nguyên'
    $uris = [regex]::Matches($out, 'src="data:image/\w+;base64,([^"]+)"')
    Check ($uris.Count -eq 31) "31 ảnh nhúng data URI (thực tế $($uris.Count))"
    $okDim = $true
    for ($i = 0; $i -lt 30; $i++) {
        $ms = New-Object IO.MemoryStream (, [Convert]::FromBase64String($uris[$i].Groups[1].Value))
        $img = [Drawing.Image]::FromStream($ms)
        if ($img.Width -ne 640 -or $img.Height -ne 360) { $okDim = $false }
        $img.Dispose(); $ms.Dispose()
    }
    Check $okDim '30 ảnh lớn về đúng 640x360'
    $small = @(Get-WebImgs $out)[30]
    Check ($small.W -eq 400 -and $small.H -eq 200 -and $small.Len -lt (Get-Item $files[30]).Length) 'ảnh nhỏ (nhiễu): vẫn 400x200 và nhỏ hơn bản gốc'
    Check ($r.Changed -eq 31) "Changed = 31 (file: luôn nhúng), thực tế $($r.Changed)"
    $ratio = $r.Before / [double]$r.After
    Check ($ratio -ge 5) ('giảm {0:N1} lần (cần >= 5)' -f $ratio)
    Check ($sw.Elapsed.TotalSeconds -lt 15) ('thời gian nén {0:N1}s (cần < 15s)' -f $sw.Elapsed.TotalSeconds)
    $peak = [Diagnostics.Process]::GetCurrentProcess().PeakWorkingSet64
    Check ($peak -lt 1GB) ('bộ nhớ tối đa {0:N0} MB (cần < 1024)' -f ($peak / 1MB))
    Write-Host ('{0} ảnh: {1:N1} MB -> {2:N1} MB, clipboard HTML {3:N1} MB, {4:N1}s' -f $r.Count, ($r.Before / 1MB), ($r.After / 1MB), ($raw.Length / 1MB), $sw.Elapsed.TotalSeconds)
    $r2 = Invoke-Compress
    Check ($r2.Changed -eq 0 -and [Convert]::ToBase64String([ClipNative]::GetHtml()) -eq [Convert]::ToBase64String($raw)) "report desktop nén lần hai: clipboard giữ nguyên byte (Changed $($r2.Changed))"
    Remove-Item -Recurse -Force $root
    if ($script:fail) { Write-Host "FAIL ($script:fail)" -ForegroundColor Red } else { Write-Host 'PASS (clipboard đang chứa report test đã nén, có thể Ctrl+V vào Outlook để xem)' -ForegroundColor Green }
}

# ---------------- Chẩn đoán (chỉ ghi cấu trúc, không ghi chữ trong report) ----------------
function Invoke-Diag {
    $o = New-Object Collections.Generic.List[string]
    $red = { param($s) $s = [regex]::Replace($s, '(?i)(Users[\\/])[^\\/]+', '$1<user>'); $s = [regex]::Replace($s, '(?i)\b(alt|title)\s*=\s*(?:"[^"]*"|''[^'']*''|[^\s>]+)', '$1="<đã xóa>"'); if ($s.Length -gt 400) { $s.Substring(0, 400) + '...' } else { $s } }
    $o.Add('== Định dạng clipboard ==')
    foreach ($f in [ClipNative]::ListFormats()) { $o.Add($f) }
    $raw = [ClipNative]::GetHtml()
    if (-not $raw) { $o.Add('KHÔNG có HTML Format') } else {
        $o.Add("== Header (HTML Format {0:N0} bytes) ==" -f $raw.Length)
        $o.Add([Text.Encoding]::ASCII.GetString($raw, 0, [Math]::Min(300, $raw.Length)).Split('<')[0].Trim())
        $html = Read-CfHtml $raw
        $o.Add('== Meta ==')
        foreach ($m in [regex]::Matches($html, '(?i)<meta[^>]*>')) { $o.Add((& $red $m.Value)) }
        $o.Add('== Đếm ==')
        foreach ($p in '<img', '<v:imagedata', '<v:shape', 'file:', 'data:image', 'http', '<!--StartFragment', '<![if !vml]>', '<!--[if gte vml 1]>') {
            $o.Add("$p : " + [regex]::Matches($html, [regex]::Escape($p), 'IgnoreCase').Count)
        }
        foreach ($t in @(@('<img', $ImgRe), @('<v:imagedata', '(?is)<v:imagedata\b[^>]*>'))) {
            $o.Add("== 5 thẻ $($t[0]) đầu ==")
            foreach ($m in @([regex]::Matches($html, $t[1])) | Select-Object -First 5) {
                $sm = [regex]::Match($m.Value, '(?i)\bsrc\s*=\s*(?:"([^"]*)"|''([^'']*)''|([^\s>]+))')
                $src = $sm.Groups[1].Value + $sm.Groups[2].Value + $sm.Groups[3].Value
                $ok = ''
                if ($src -match '^file:') { try { $ok = ' [file tồn tại: ' + (Test-Path -LiteralPath ([Uri]$src).LocalPath) + ']' } catch { $ok = ' [src lỗi]' } }
                $o.Add((& $red $m.Value) + $ok)
            }
        }
        $r = Compress-Html $html
        $o.Add('== Thử nén (không ghi clipboard) ==')
        $o.Add(('ảnh xử lý: {0}, {1:N0} -> {2:N0} bytes; sau nén còn <img>: {3}, data:image: {4}' -f $r.Count, $r.Before, $r.After, [regex]::Matches($r.Html, '(?i)<img\b').Count, [regex]::Matches($r.Html, 'data:image').Count))
        $o.Add('bỏ qua: ' + (($r.Skip.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ', ') + '; độ dài chữ thuần: ' + [ClipNative]::GetText().Length)
    }
    $file = Join-Path ([Environment]::GetFolderPath('Desktop')) 'nen-clipboard-diag.txt'
    [IO.File]::WriteAllLines($file, $o, (New-Object Text.UTF8Encoding $true))
    Start-Process notepad.exe $file
}

# Hai ảnh test: A nhúng data URI, B trỏ file:/// — dán vào Outlook để biết kiểu nào hiện.
function Invoke-TestPaste {
    $mk = {
        param($color, $label)
        $bmp = New-Object Drawing.Bitmap 240, 120
        $g = [Drawing.Graphics]::FromImage($bmp); $g.Clear($color)
        $font = New-Object Drawing.Font 'Segoe UI', 28
        $g.DrawString($label, $font, [Drawing.Brushes]::White, 20, 30); $font.Dispose(); $g.Dispose()
        $ms = New-Object IO.MemoryStream; $bmp.Save($ms, [Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
        , $ms.ToArray()
    }
    $a = & $mk ([Drawing.Color]::Red) 'A: data'
    $fb = Join-Path $env:TEMP 'nen-clipboard-testB.png'
    [IO.File]::WriteAllBytes($fb, (& $mk ([Drawing.Color]::Blue) 'B: file'))
    $html = '<html><body><!--StartFragment--><p>Ảnh A (data URI):</p><img width=240 height=120 src="data:image/png;base64,' + [Convert]::ToBase64String($a) + '"><p>Ảnh B (file):</p><img width=240 height=120 src="' + ([Uri]$fb).AbsoluteUri + '"><p>Hết test.</p><!--EndFragment--></body></html>'
    [ClipNative]::SetHtmlAndText((Build-CfHtml $html), 'Test A B')
    Show-Note 'Đã đặt ảnh test A (đỏ) và B (xanh) vào clipboard. Ctrl+V vào Outlook, xem ảnh nào hiện.'
}

if ($SelfTest) { Invoke-SelfTest; exit $script:fail }
if ($Diag) { Invoke-Diag; exit }
if ($TestPaste) { Invoke-TestPaste; exit }
try {
    $r = Invoke-Compress
    $msg = if ($r.Changed -eq 0) { '{0} ảnh đã đủ nhỏ, cứ Ctrl+V vào Outlook.' -f $r.Count }
    else { '{0} ảnh: {1:N1} MB -> {2:N1} MB. Giờ Ctrl+V vào Outlook.' -f $r.Count, ($r.Before / 1MB), ($r.After / 1MB) }
    $k = [int]($r.Skip.Values | Measure-Object -Sum).Sum
    if ($k) { $msg += " Bỏ qua $k ảnh." }
    Show-Note $msg
} catch {
    Show-Note ('Lỗi: ' + $_.Exception.Message + ' Clipboard giữ nguyên.') $true
}
